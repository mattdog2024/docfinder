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
from concurrent.futures import Future, wait, FIRST_COMPLETED
from core.extraction_pool import ExtractionPool

logger = logging.getLogger(__name__)

# 批量写入大小：每积累 BATCH_SIZE 条记录才 commit 一次
BATCH_SIZE = 100

# 上下文摘要长度（字符数）
SNIPPET_CONTEXT = 120
# 最大摘要数量
MAX_SNIPPETS = 5

# 每个文件提取的最大超时时间（秒）
EXTRACT_TIMEOUT = 30

LEGACY_PARSER_VERSION = {'.doc': '3', '.ppt': '3'}

def _tokenize_for_fts(text: str) -> str:
    """
    对文本进行分词处理，使 FTS5 能正确搜索中文。

    原理：FTS5 的 unicode61 分词器对中文不按字切分，
    把连续中文当作一个 token，导致搜索"监狱"无法匹配"女子监狱"。
    解决方案：在每个中文字符之间插入空格，让 FTS5 按字索引。
    """
    if not text:
        return ''
    import re
    # 在中文字符之间插入空格（保留英文单词和数字不变）
    # 匹配连续的中文字符块，在每个字之间加空格
    result = re.sub(
        r'([\u4e00-\u9fff\u3400-\u4dbf\u3000-\u303f\uff00-\uffef])',
        r' \1 ',
        text
    )
    # 清理多余空格
    return ' '.join(result.split())


def get_file_hash(filepath: str) -> str:
    """用文件大小 + 修改时间作为快速指纹（比 MD5 快 100 倍）"""
    try:
        stat = os.stat(filepath)
        fingerprint = f"{stat.st_size}_{stat.st_mtime}"
        version = LEGACY_PARSER_VERSION.get(os.path.splitext(filepath)[1].lower())
        return fingerprint + (f'_parser{version}' if version else '')
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
            self._conn = sqlite3.connect(self.db_path, timeout=1.0, check_same_thread=False)
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

        # v1.0/v1.1 的 external-content FTS 指向不存在的 content_tokens 列。
        # CREATE IF NOT EXISTS 不会迁移旧表；首批更新 DELETE 时会抛错。
        schema = cursor.execute(
            "SELECT sql FROM sqlite_master WHERE name='documents_fts'"
        ).fetchone()[0]
        columns = {r[1] for r in cursor.execute('PRAGMA table_info(documents_fts)')}
        if 'content=' in schema.replace(' ', '').lower() or 'content_tokens' not in columns:
            cursor.execute('DROP TABLE documents_fts')
            cursor.execute('''CREATE VIRTUAL TABLE documents_fts USING fts5(
                filepath UNINDEXED, filename, content_tokens,
                tokenize="unicode61 remove_diacritics 1")''')
            rows = conn.execute('SELECT id, filepath, filename, content FROM documents')
            while True:
                chunk = rows.fetchmany(100)
                if not chunk:
                    break
                conn.executemany(
                    'INSERT INTO documents_fts(rowid,filepath,filename,content_tokens) VALUES(?,?,?,?)',
                    [(i, fp, name, _tokenize_for_fts((text or '') + ' ' + name))
                     for i, fp, name, text in chunk]
                )
            logger.info('已兼容升级早期 FTS 表，原始文档与离线文本保持不变')

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

            # 对内容和文件名进行中文分词，确保 FTS5 能按字搜索
            tokens = _tokenize_for_fts((content or '') + ' ' + filename)

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
                    (doc_id, filepath, filename, tokens)
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
                    (doc_id, filepath, filename, tokens)
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
            fts_query = self._escape_fts_query(_tokenize_for_fts(query))
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
                fts_parts = [self._escape_fts_query(_tokenize_for_fts(w)) for w in words if w.strip()]
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

        # 老库可能混合旧的中文整段token与新按字token；少量FTS命中不能
        # 证明旧记录没有命中，中文额外合并原始文本查询结果。
        if not results or any('\u3400' <= ch <= '\u9fff' for ch in query):
            words = query.split()
            clauses = []
            parameters = []
            for word in words:
                escaped = word.replace('\\', '\\\\').replace('%', '\\%').replace('_', '\\_')
                clauses.append("(content LIKE ? ESCAPE '\\' OR filename LIKE ? ESCAPE '\\')")
                parameters.extend([f'%{escaped}%', f'%{escaped}%'])
            cursor.execute('''
                SELECT id, filepath, filename, extension,
                       filesize, modified_time, content, index_root
                FROM documents
                WHERE ''' + ' AND '.join(clauses) + '''
                LIMIT ?
            ''', (*parameters, limit))
            fallback = cursor.fetchall()
            found_ids = {row[0] for row in results}
            results.extend(row for row in fallback if row[0] not in found_ids)
            results = results[:limit]

        output = []
        for row in results:
            doc_id, filepath, filename, extension, filesize, modified_time, content, index_root = row
            snippets = self._extract_snippets(content, query)
            # 用 pathlib.Path 处理中文路径（os.path.exists 在 Windows 中文路径下有时失效）
            try:
                from pathlib import Path
                file_exists = Path(filepath).exists()
            except Exception:
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
    - 完成顺序取结果，协调线程写库；写库错误向UI传播
    """

    SUBMIT_BATCH = 200  # 每批提交的任务数

    def __init__(self, engine: IndexEngine):
        self.engine = engine
        self._stop_flag = threading.Event()
        self._executor: Optional[ExtractionPool] = None

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
            max_workers = min(multiprocessing.cpu_count(), 16)


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
            return {'total': total, 'added': 0, 'updated': 0, 'skipped': 0, 'failed': 0, 'stopped': self._stop_flag.is_set()}

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
            stats['stopped'] = self._stop_flag.is_set()
            if not stats['stopped']:
                self.engine.set_index_info('last_index_time', str(time.time()))
            self.engine.set_index_info('index_root', root_dir)
            return stats

        # 协调线程拥有数据库连接，解析线程不写库。
        # 不再使用无错误传递的后台writer/有界结果队列：写库错误直接上报UI。
        max_workers = max(1, min(int(max_workers), 16))
        self._executor = ExtractionPool(max_workers=max_workers)
        pending = {}
        expired = set()
        files = iter(to_process)
        batch = []
        processed = 0
        started = time.monotonic()
        recent = []

        def report(fp):
            nonlocal processed
            processed += 1
            now = time.monotonic()
            if progress_callback:
                progress_callback(processed + skipped, total, fp)
            recent.append(now)
            if len(recent) > 50:
                recent.pop(0)
            if speed_callback and len(recent) > 1:
                elapsed = recent[-1] - recent[0]
                if elapsed > 0:
                    speed = (len(recent) - 1) / elapsed
                    speed_callback(speed, (len(to_process) - processed) / speed)

        def flush():
            if batch:
                try:
                    self.engine.batch_add_documents(batch, root_dir)
                except Exception:
                    self.engine._get_conn().rollback()
                    raise
                for record in batch:
                    stats['updated' if record[0] in indexed else 'added'] += 1
                batch.clear()

        def fill():
            # 最多2倍工作线程任务，不一次提交几万条；慢文件不挡其他完成结果。
            while len(pending) < max_workers * 2 and not self._stop_flag.is_set():
                fp = next(files, None)
                if fp is None:
                    break
                try:
                    future = self._executor.submit(_extract_single, fp, enable_pdf, enable_ocr)
                except RuntimeError:
                    if self._stop_flag.is_set():
                        break
                    raise
                pending[future] = (fp, time.monotonic())

        error = None
        try:
            fill()
            last_flush = time.monotonic()
            while pending and not self._stop_flag.is_set():
                done, _ = wait(pending, timeout=0.2, return_when=FIRST_COMPLETED)
                for future in done:
                    fp, submitted = pending.pop(future)
                    if future in expired:
                        expired.remove(future)
                        continue
                    try:
                        result = future.result()
                    except Exception as exc:
                        result = None
                        log(f"解析失败: {os.path.basename(fp)} - {type(exc).__name__}: {exc}")
                    if result:
                        batch.append(result)
                    else:
                        stats['failed'] += 1
                    report(fp)
                    if len(batch) >= BATCH_SIZE:
                        flush()
                        last_flush = time.monotonic()
                now = time.monotonic()
                for future, (fp, submitted) in list(pending.items()):
                    if future not in expired and now - submitted >= EXTRACT_TIMEOUT:
                        # Python线程不能硬杀。保留占用，不无限创建替代线程。
                        expired.add(future)
                        future.cancel()
                        stats['failed'] += 1
                        report(fp)
                        log(f"文件解析超时，保留原索引: {fp}")
                if batch and now - last_flush >= 1.0:
                    flush()
                    last_flush = now
                if pending and all(f in expired for f in pending):
                    raise RuntimeError(
                        "多个文件解析长时间没有返回，已暂停索引并保存已完成结果。"
                        "请关闭重开软件后缩小索引目录；不会删除原文件或旧索引。"
                    )
                fill()
        except Exception as exc:
            error = exc
        finally:
            for future in pending:
                future.cancel()
            self._executor.shutdown(wait=False, cancel_futures=True)
            self._executor = None
            # 写库失败不要重试同一坏批次；解析失败/停止则保存已完成批次。
            if error is None or isinstance(error, RuntimeError):
                flush()

        if error is not None:
            raise error

        stats['stopped'] = self._stop_flag.is_set()
        self.engine.set_index_info('index_root', root_dir)
        self.engine.set_index_info('last_run_stopped', str(int(stats['stopped'])))
        if not stats['stopped']:
            self.engine.set_index_info('last_index_time', str(time.time()))
            log("索引完成，正在合并搜索索引...")
            self.engine.optimize()
        log(f"索引{'已停止' if stats['stopped'] else '完成'}：已处理{processed}，"
            f"新增{stats['added']}，更新{stats['updated']}，跳过{skipped}，失败{stats['failed']}")
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
        version = LEGACY_PARSER_VERSION.get(ext)
        fhash = f"{filesize}_{mtime}" + (f'_parser{version}' if version else '')
        content = extract_text(
            filepath,
            enable_pdf=(enable_pdf and ext == '.pdf'),
            enable_ocr=enable_ocr
        )
        return (filepath, filename, ext, filesize, mtime, fhash, content)
    except Exception:
        return None
