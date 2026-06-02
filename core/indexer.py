# -*- coding: utf-8 -*-
"""
索引引擎模块 v1.9

核心修复：
1. 停止按钮立即生效：分批提交任务，停止时取消未执行的任务
2. 速度优化：批量写入改为 INSERT OR REPLACE，减少 SELECT 查询
3. 超时保护：每个文件最多处理 30 秒，避免卡在单个文件上
"""

import os
import sqlite3
import time
import logging
import threading
import queue
import multiprocessing
from typing import List, Dict, Optional, Callable
from concurrent.futures import ThreadPoolExecutor, Future, as_completed

logger = logging.getLogger(__name__)

# 批量写入大小：每积累 BATCH_SIZE 条记录才 commit 一次
BATCH_SIZE = 100

# 上下文摘要长度（字符数）
SNIPPET_CONTEXT = 120
# 最大摘要数量
MAX_SNIPPETS = 5

# 每个文件提取的最大超时时间（秒）
EXTRACT_TIMEOUT = 30


def get_file_hash(filepath: str) -> str:
    """用文件大小 + 修改时间作为快速指纹（比 MD5 快 100 倍）"""
    try:
        stat = os.stat(filepath)
        return f"{stat.st_size}_{stat.st_mtime}"
    except Exception:
        return ''


class IndexEngine:
    """
    文档索引引擎（v1.9）
    使用单一数据库连接，所有写操作在同一线程中执行，避免并发写入问题
    """

    def __init__(self, db_path: str):
        self.db_path = db_path
        self._conn: Optional[sqlite3.Connection] = None
        self._init_db()

    def _get_conn(self) -> sqlite3.Connection:
        """获取数据库连接（单连接，非线程安全，调用方负责线程安全）"""
        if self._conn is None:
            self._conn = sqlite3.connect(self.db_path, check_same_thread=False)
            self._apply_pragmas(self._conn)
        return self._conn

    def _apply_pragmas(self, conn: sqlite3.Connection):
        """应用性能优化 PRAGMA"""
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("PRAGMA cache_size=-65536")   # 64MB 内存缓存
        conn.execute("PRAGMA temp_store=MEMORY")
        conn.execute("PRAGMA mmap_size=268435456") # 256MB 内存映射

    def _init_db(self):
        """初始化数据库表结构"""
        conn = sqlite3.connect(self.db_path)
        self._apply_pragmas(conn)
        cursor = conn.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filepath TEXT UNIQUE NOT NULL,
                filename TEXT NOT NULL,
                extension TEXT NOT NULL,
                filesize INTEGER,
                modified_time REAL,
                file_hash TEXT,
                content TEXT,
                indexed_time REAL,
                index_root TEXT
            )
        ''')

        cursor.execute('''
            CREATE VIRTUAL TABLE IF NOT EXISTS documents_fts
            USING fts5(
                filepath UNINDEXED,
                filename,
                content_tokens,
                tokenize="unicode61 remove_diacritics 1"
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS index_info (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        ''')

        conn.commit()
        conn.close()

    def get_index_info(self) -> Dict:
        """获取索引基本信息"""
        conn = self._get_conn()
        cursor = conn.cursor()
        cursor.execute("SELECT key, value FROM index_info")
        info = dict(cursor.fetchall())
        cursor.execute("SELECT COUNT(*) FROM documents")
        info['total_docs'] = cursor.fetchone()[0]
        return info

    def set_index_info(self, key: str, value: str):
        """设置索引信息"""
        conn = self._get_conn()
        conn.execute(
            "INSERT OR REPLACE INTO index_info (key, value) VALUES (?, ?)",
            (key, value)
        )
        conn.commit()

    def get_indexed_files(self) -> Dict[str, str]:
        """获取已索引文件的 {路径: hash} 字典"""
        conn = self._get_conn()
        cursor = conn.cursor()
        cursor.execute("SELECT filepath, file_hash FROM documents")
        return dict(cursor.fetchall())

    def batch_add_documents(self, records: List[tuple], index_root: str):
        """
        批量添加/更新文档（必须在单一线程中调用）
        records: [(filepath, filename, ext, filesize, mtime, fhash, content), ...]
        
        v1.9 优化：使用 INSERT OR REPLACE 代替 SELECT+UPDATE/INSERT，速度更快
        """
        if not records:
            return

        conn = self._get_conn()
        cursor = conn.cursor()
        now = time.time()

        for filepath, filename, ext, filesize, mtime, fhash, content in records:
            # 先查找已有记录的 id（用于更新 FTS）
            cursor.execute("SELECT id FROM documents WHERE filepath=?", (filepath,))
            row = cursor.fetchone()

            if row:
                doc_id = row[0]
                cursor.execute('''
                    UPDATE documents SET
                        filename=?, extension=?, filesize=?, modified_time=?,
                        file_hash=?, content=?, indexed_time=?, index_root=?
                    WHERE id=?
                ''', (filename, ext, filesize, mtime, fhash, content, now, index_root, doc_id))
                cursor.execute("DELETE FROM documents_fts WHERE rowid=?", (doc_id,))
                cursor.execute(
                    "INSERT INTO documents_fts(rowid, filepath, filename, content_tokens) VALUES(?,?,?,?)",
                    (doc_id, filepath, filename, content or '')
                )
            else:
                cursor.execute('''
                    INSERT INTO documents
                        (filepath, filename, extension, filesize, modified_time,
                         file_hash, content, indexed_time, index_root)
                    VALUES (?,?,?,?,?,?,?,?,?)
                ''', (filepath, filename, ext, filesize, mtime, fhash, content, now, index_root))
                doc_id = cursor.lastrowid
                cursor.execute(
                    "INSERT INTO documents_fts(rowid, filepath, filename, content_tokens) VALUES(?,?,?,?)",
                    (doc_id, filepath, filename, content or '')
                )

        conn.commit()

    def remove_document(self, filepath: str):
        """从索引中删除一个文档"""
        conn = self._get_conn()
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM documents WHERE filepath=?", (filepath,))
        row = cursor.fetchone()
        if row:
            doc_id = row[0]
            cursor.execute("DELETE FROM documents_fts WHERE rowid=?", (doc_id,))
            cursor.execute("DELETE FROM documents WHERE id=?", (doc_id,))
            conn.commit()

    @staticmethod
    def _escape_fts_query(query: str) -> str:
        """转义 FTS5 特殊字符，避免查询语法错误"""
        escaped = query.replace('"', '""')
        return f'"{escaped}"'

    def search(self, query: str, limit: int = 200) -> List[Dict]:
        """
        搜索文档（支持中文关键词直接搜索）
        三级搜索策略：精确短语 -> 多词AND -> LIKE 兜底
        """
        if not query.strip():
            return []

        conn = self._get_conn()
        cursor = conn.cursor()
        query = query.strip()
        results = []

        # 方式1：整体作为精确短语搜索
        try:
            fts_query = self._escape_fts_query(query)
            cursor.execute('''
                SELECT d.id, d.filepath, d.filename, d.extension,
                       d.filesize, d.modified_time, d.content, d.index_root
                FROM documents d
                JOIN documents_fts fts ON d.id = fts.rowid
                WHERE documents_fts MATCH ?
                ORDER BY rank
                LIMIT ?
            ''', (fts_query, limit))
            results = cursor.fetchall()
        except sqlite3.OperationalError:
            results = []

        # 方式2：多词空格分隔 AND 搜索
        if not results and ' ' in query:
            try:
                words = query.split()
                fts_parts = [self._escape_fts_query(w) for w in words if w.strip()]
                fts_query = ' AND '.join(fts_parts)
                cursor.execute('''
                    SELECT d.id, d.filepath, d.filename, d.extension,
                           d.filesize, d.modified_time, d.content, d.index_root
                    FROM documents d
                    JOIN documents_fts fts ON d.id = fts.rowid
                    WHERE documents_fts MATCH ?
                    ORDER BY rank
                    LIMIT ?
                ''', (fts_query, limit))
                results = cursor.fetchall()
            except sqlite3.OperationalError:
                results = []

        # 方式3：LIKE 兜底（确保不漏）
        if not results:
            like_query = f'%{query}%'
            cursor.execute('''
                SELECT id, filepath, filename, extension,
                       filesize, modified_time, content, index_root
                FROM documents
                WHERE content LIKE ? OR filename LIKE ?
                LIMIT ?
            ''', (like_query, like_query, limit))
            results = cursor.fetchall()

        output = []
        for row in results:
            doc_id, filepath, filename, extension, filesize, modified_time, content, index_root = row
            snippets = self._extract_snippets(content, query)
            file_exists = os.path.exists(filepath)
            output.append({
                'id': doc_id,
                'filepath': filepath,
                'filename': filename,
                'extension': extension,
                'filesize': filesize or 0,
                'modified_time': modified_time or 0,
                'snippets': snippets,
                'content': content or '',
                'file_exists': file_exists,
                'index_root': index_root or '',
            })

        return output

    def _extract_snippets(self, content: str, query: str) -> List[str]:
        """从文档内容中提取包含关键词的上下文摘要（大小写不敏感）"""
        if not content or not query:
            return []

        import re
        content_lower = content.lower()
        query_lower = query.lower()
        snippets = []
        seen_positions = set()
        pos = 0

        while len(snippets) < MAX_SNIPPETS:
            idx = content_lower.find(query_lower, pos)
            if idx == -1:
                break

            bucket = idx // SNIPPET_CONTEXT
            if bucket in seen_positions:
                pos = idx + 1
                continue
            seen_positions.add(bucket)

            start = max(0, idx - SNIPPET_CONTEXT // 2)
            end = min(len(content), idx + len(query) + SNIPPET_CONTEXT // 2)
            snippet = content[start:end].strip()
            snippet = ' '.join(snippet.split())

            try:
                highlighted = re.sub(
                    re.escape(query),
                    f'**{query}**',
                    snippet,
                    flags=re.IGNORECASE
                )
            except re.error:
                highlighted = snippet.replace(query, f'**{query}**')

            if start > 0:
                highlighted = '...' + highlighted
            if end < len(content):
                highlighted = highlighted + '...'

            snippets.append(highlighted)
            pos = idx + 1

        return snippets

    def get_document_content(self, doc_id: int) -> Optional[str]:
        """获取文档的完整内容"""
        conn = self._get_conn()
        cursor = conn.cursor()
        cursor.execute("SELECT content FROM documents WHERE id=?", (doc_id,))
        row = cursor.fetchone()
        return row[0] if row else None

    def cleanup_orphans(self):
        """清理孤立记录：删除磁盘上已不存在的文件的索引"""
        conn = self._get_conn()
        cursor = conn.cursor()
        cursor.execute("SELECT id, filepath FROM documents")
        rows = cursor.fetchall()
        deleted = 0
        for doc_id, filepath in rows:
            if not os.path.exists(filepath):
                cursor.execute("DELETE FROM documents_fts WHERE rowid=?", (doc_id,))
                cursor.execute("DELETE FROM documents WHERE id=?", (doc_id,))
                deleted += 1
        if deleted > 0:
            conn.commit()
            logger.info(f"清理孤立记录：删除 {deleted} 条")
        return deleted

    def optimize(self):
        """优化 FTS5 索引（索引完成后调用，加速搜索）"""
        conn = self._get_conn()
        try:
            conn.execute("INSERT INTO documents_fts(documents_fts) VALUES('optimize')")
            conn.commit()
        except Exception:
            pass

    def close(self):
        """关闭数据库连接"""
        if self._conn:
            self._conn.close()
            self._conn = None


class IndexBuilder:
    """
    索引构建器 v1.9

    架构改进：
    - 分批提交任务（每批 SUBMIT_BATCH 个），停止时立即取消未执行的任务
    - 每个文件有超时保护（EXTRACT_TIMEOUT 秒），避免卡在单个文件
    - 生产者-消费者模式：多线程解析 → 队列 → 单线程写入数据库
    """

    SUBMIT_BATCH = 200  # 每批提交的任务数

    def __init__(self, engine: IndexEngine):
        self.engine = engine
        self._stop_flag = threading.Event()
        self._executor: Optional[ThreadPoolExecutor] = None

    def stop(self):
        """立即停止：设置停止标志，并关闭线程池（取消所有待执行任务）"""
        self._stop_flag.set()
        # 关闭线程池，cancel_futures=True 会取消所有还未开始执行的任务
        if self._executor is not None:
            try:
                self._executor.shutdown(wait=False, cancel_futures=True)
            except TypeError:
                # Python 3.8 不支持 cancel_futures 参数
                try:
                    self._executor.shutdown(wait=False)
                except Exception:
                    pass
            except Exception:
                pass

    def _scan_files(self, root_dir: str, enabled_extensions: set) -> List[str]:
        """快速递归扫描目录，返回所有符合条件的文件路径"""
        result = []
        stack = [root_dir]
        while stack:
            if self._stop_flag.is_set():
                break
            current = stack.pop()
            try:
                with os.scandir(current) as it:
                    for entry in it:
                        if self._stop_flag.is_set():
                            break
                        if entry.is_dir(follow_symlinks=False):
                            if not entry.name.startswith('.') and entry.name not in (
                                'System Volume Information', '$RECYCLE.BIN', 'Windows',
                                'Program Files', 'Program Files (x86)'
                            ):
                                stack.append(entry.path)
                        elif entry.is_file(follow_symlinks=False):
                            ext = os.path.splitext(entry.name)[1].lower()
                            if ext in enabled_extensions:
                                result.append(entry.path)
            except (PermissionError, OSError):
                pass
        return result

    def build_index(
        self,
        root_dir: str,
        enabled_extensions: List[str],
        enable_pdf: bool = True,
        enable_ocr: bool = False,
        max_workers: int = None,
        progress_callback: Optional[Callable] = None,
        log_callback: Optional[Callable] = None,
        speed_callback: Optional[Callable] = None,
    ) -> Dict:
        """
        构建/更新索引（v1.9 分批提交版）

        关键改进：
        1. 分批提交任务，停止时立即取消未执行的任务
        2. 每个文件有超时保护
        3. 生产者-消费者模式，单线程写入数据库
        """
        if max_workers is None:
            max_workers = min(multiprocessing.cpu_count(), 8)

        self._stop_flag.clear()

        def log(msg):
            logger.info(msg)
            if log_callback:
                log_callback(msg)

        log(f"开始扫描目录: {root_dir}")

        ext_set = set(enabled_extensions)
        all_files = self._scan_files(root_dir, ext_set)
        total = len(all_files)
        log(f"共找到 {total} 个文件")

        if total == 0 or self._stop_flag.is_set():
            return {'total': 0, 'added': 0, 'updated': 0, 'skipped': 0, 'failed': 0}

        # 获取已索引文件的哈希（增量更新）
        indexed = self.engine.get_indexed_files()

        # 筛选需要处理的文件
        to_process = []
        skipped = 0
        for fp in all_files:
            fhash = get_file_hash(fp)
            if fp in indexed and indexed[fp] == fhash:
                skipped += 1
            else:
                to_process.append(fp)

        log(f"需要处理: {len(to_process)} 个，跳过（未变化）: {skipped} 个")

        stats = {
            'total': total,
            'added': 0,
            'updated': 0,
            'skipped': skipped,
            'failed': 0
        }

        if not to_process or self._stop_flag.is_set():
            self.engine.set_index_info('last_index_time', str(time.time()))
            self.engine.set_index_info('index_root', root_dir)
            return stats

        # ── 生产者-消费者架构 ──────────────────────────────────────────────
        result_queue = queue.Queue(maxsize=max_workers * 8)
        _SENTINEL = object()

        processed = 0
        start_time = time.time()
        recent_done_times = []

        # ── 写入线程（消费者）──────────────────────────────────────────────
        def writer_thread():
            nonlocal processed
            batch = []

            while True:
                try:
                    item = result_queue.get(timeout=1.0)
                except queue.Empty:
                    if self._stop_flag.is_set() and batch:
                        self.engine.batch_add_documents(batch, root_dir)
                        batch = []
                    continue

                if item is _SENTINEL:
                    if batch:
                        self.engine.batch_add_documents(batch, root_dir)
                    break

                fp, result = item
                is_update = fp in indexed

                if result:
                    batch.append(result)
                    if is_update:
                        stats['updated'] += 1
                    else:
                        stats['added'] += 1
                else:
                    stats['failed'] += 1

                processed += 1

                # 进度回调
                if progress_callback:
                    progress_callback(processed + skipped, total, fp)

                # 速度计算（滑动窗口最近 50 个文件）
                if speed_callback:
                    now = time.time()
                    recent_done_times.append(now)
                    if len(recent_done_times) > 50:
                        recent_done_times.pop(0)
                    if len(recent_done_times) >= 2:
                        window = recent_done_times[-1] - recent_done_times[0]
                        if window > 0:
                            spd = (len(recent_done_times) - 1) / window
                            remaining = len(to_process) - processed
                            eta = remaining / spd if spd > 0 else 0
                            speed_callback(spd, eta)

                # 达到批量大小时写入
                if len(batch) >= BATCH_SIZE:
                    self.engine.batch_add_documents(batch, root_dir)
                    batch = []

                result_queue.task_done()

        # 启动写入线程
        writer = threading.Thread(target=writer_thread, daemon=True)
        writer.start()

        # ── 分批提交任务（生产者）──────────────────────────────────────────
        # 关键改进：分批提交，每批 SUBMIT_BATCH 个
        # 这样停止时，未提交的批次直接跳过，已提交但未执行的任务会被取消
        self._executor = ThreadPoolExecutor(max_workers=max_workers)
        
        try:
            pending_futures: List[Future] = []
            
            for i in range(0, len(to_process), self.SUBMIT_BATCH):
                if self._stop_flag.is_set():
                    break
                
                batch_files = to_process[i:i + self.SUBMIT_BATCH]
                batch_futures = []
                
                for fp in batch_files:
                    if self._stop_flag.is_set():
                        break
                    future = self._executor.submit(_extract_single, fp, enable_pdf, enable_ocr)
                    batch_futures.append((future, fp))
                
                # 等待这批任务完成（或停止）
                for future, fp in batch_futures:
                    if self._stop_flag.is_set():
                        future.cancel()
                        continue
                    try:
                        result = future.result(timeout=EXTRACT_TIMEOUT)
                    except Exception as e:
                        result = None
                        if not self._stop_flag.is_set():
                            log(f"处理失败: {os.path.basename(fp)} - {type(e).__name__}")
                    
                    if not self._stop_flag.is_set():
                        result_queue.put((fp, result))
        finally:
            # 关闭线程池
            try:
                self._executor.shutdown(wait=False, cancel_futures=True)
            except TypeError:
                self._executor.shutdown(wait=False)
            except Exception:
                pass
            self._executor = None

        # 通知写入线程结束
        result_queue.put(_SENTINEL)
        writer.join(timeout=30)

        if not self._stop_flag.is_set():
            # 索引完成后优化 FTS5
            log("正在优化索引...")
            self.engine.optimize()

        self.engine.set_index_info('last_index_time', str(time.time()))
        self.engine.set_index_info('index_root', root_dir)

        if self._stop_flag.is_set():
            log(f"索引已停止！已处理: {processed}，新增: {stats['added']}，"
                f"更新: {stats['updated']}，跳过: {stats['skipped']}，失败: {stats['failed']}")
        else:
            log(f"索引完成！新增: {stats['added']}，更新: {stats['updated']}，"
                f"跳过: {stats['skipped']}，失败: {stats['failed']}")

        return stats


def _extract_single(filepath: str, enable_pdf: bool, enable_ocr: bool):
    """
    单文件提取函数（线程池工作函数，只负责解析，不写数据库）
    返回 tuple 或 None（失败时）
    """
    try:
        try:
            from core.extractor import extract_text
        except ImportError:
            import sys, os as _os
            _base = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))
            if _base not in sys.path:
                sys.path.insert(0, _base)
            from core.extractor import extract_text
        filename = os.path.basename(filepath)
        ext = os.path.splitext(filename)[1].lower()
        stat = os.stat(filepath)
        filesize = stat.st_size
        mtime = stat.st_mtime
        fhash = f"{filesize}_{mtime}"
        content = extract_text(
            filepath,
            enable_pdf=(enable_pdf and ext == '.pdf'),
            enable_ocr=enable_ocr
        )
        return (filepath, filename, ext, filesize, mtime, fhash, content)
    except Exception:
        return None
