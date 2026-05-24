# -*- coding: utf-8 -*-
"""
索引引擎模块
使用 SQLite FTS5 + jieba 中文分词
"""

import os
import sqlite3
import hashlib
import time
import logging
import threading
from typing import List, Dict, Optional, Tuple, Callable
from concurrent.futures import ThreadPoolExecutor, as_completed

import jieba

from .extractor import extract_text, SUPPORTED_EXTENSIONS

logger = logging.getLogger(__name__)

# jieba 静默模式
jieba.setLogLevel(logging.WARNING)

# 上下文摘要长度（字符数）
SNIPPET_CONTEXT = 120
# 最大摘要数量
MAX_SNIPPETS = 5


def tokenize_chinese(text: str) -> str:
    """
    对文本进行 jieba 分词，返回空格分隔的词语
    用于写入 FTS5 索引
    """
    if not text:
        return ''
    words = jieba.cut(text, cut_all=False)
    return ' '.join(w for w in words if w.strip())


def get_file_hash(filepath: str) -> str:
    """计算文件的 MD5 哈希，用于判断文件是否变化"""
    try:
        stat = os.stat(filepath)
        # 用文件大小 + 修改时间作为快速哈希（比读取全文件快得多）
        return f"{stat.st_size}_{stat.st_mtime}"
    except Exception:
        return ''


class IndexEngine:
    """
    文档索引引擎
    数据库结构：
      - documents: 存储文件元数据和原始文本
      - documents_fts: FTS5 虚拟表，存储分词后的文本
    """

    def __init__(self, db_path: str):
        self.db_path = db_path
        self._local = threading.local()
        self._init_db()

    def _get_conn(self) -> sqlite3.Connection:
        """获取线程本地的数据库连接"""
        if not hasattr(self._local, 'conn') or self._local.conn is None:
            conn = sqlite3.connect(self.db_path, check_same_thread=False)
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA synchronous=NORMAL")
            conn.execute("PRAGMA cache_size=10000")
            self._local.conn = conn
        return self._local.conn

    def _init_db(self):
        """初始化数据库表结构"""
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA journal_mode=WAL")
        cursor = conn.cursor()

        # 文档元数据表
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

        # FTS5 全文搜索表（存储分词后的文本）
        cursor.execute('''
            CREATE VIRTUAL TABLE IF NOT EXISTS documents_fts
            USING fts5(
                filepath UNINDEXED,
                filename,
                content_tokens,
                content="documents",
                content_rowid="id",
                tokenize="unicode61"
            )
        ''')

        # 索引信息表
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

    def add_document(self, filepath: str, filename: str, extension: str,
                     filesize: int, modified_time: float, file_hash: str,
                     content: str, index_root: str):
        """添加或更新一个文档到索引"""
        conn = self._get_conn()
        cursor = conn.cursor()

        # 对内容进行分词
        tokens = tokenize_chinese(filename + ' ' + content)

        # 检查是否已存在
        cursor.execute("SELECT id FROM documents WHERE filepath = ?", (filepath,))
        row = cursor.fetchone()

        if row:
            doc_id = row[0]
            cursor.execute('''
                UPDATE documents SET
                    filename=?, extension=?, filesize=?, modified_time=?,
                    file_hash=?, content=?, indexed_time=?, index_root=?
                WHERE id=?
            ''', (filename, extension, filesize, modified_time,
                  file_hash, content, time.time(), index_root, doc_id))
            # 更新 FTS 表
            cursor.execute("DELETE FROM documents_fts WHERE rowid=?", (doc_id,))
            cursor.execute(
                "INSERT INTO documents_fts(rowid, filepath, filename, content_tokens) VALUES (?,?,?,?)",
                (doc_id, filepath, filename, tokens)
            )
        else:
            cursor.execute('''
                INSERT INTO documents
                    (filepath, filename, extension, filesize, modified_time,
                     file_hash, content, indexed_time, index_root)
                VALUES (?,?,?,?,?,?,?,?,?)
            ''', (filepath, filename, extension, filesize, modified_time,
                  file_hash, content, time.time(), index_root))
            doc_id = cursor.lastrowid
            cursor.execute(
                "INSERT INTO documents_fts(rowid, filepath, filename, content_tokens) VALUES (?,?,?,?)",
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

    def search(self, query: str, limit: int = 200) -> List[Dict]:
        """
        搜索文档
        返回匹配结果列表，每项包含文件信息和命中摘要
        """
        if not query.strip():
            return []

        conn = self._get_conn()
        cursor = conn.cursor()

        # 对查询词进行分词
        query_words = list(jieba.cut(query.strip(), cut_all=False))
        query_words = [w for w in query_words if w.strip()]

        if not query_words:
            return []

        # 构建 FTS5 查询（支持多词 AND 搜索）
        fts_query = ' '.join(f'"{w}"' for w in query_words)

        try:
            cursor.execute('''
                SELECT d.id, d.filepath, d.filename, d.extension,
                       d.filesize, d.modified_time, d.content, d.index_root
                FROM documents d
                JOIN documents_fts fts ON d.id = fts.rowid
                WHERE documents_fts MATCH ?
                ORDER BY rank
                LIMIT ?
            ''', (fts_query, limit))
        except sqlite3.OperationalError:
            # FTS 查询失败时降级为 LIKE 搜索
            like_query = f'%{query}%'
            cursor.execute('''
                SELECT id, filepath, filename, extension,
                       filesize, modified_time, content, index_root
                FROM documents
                WHERE content LIKE ? OR filename LIKE ?
                LIMIT ?
            ''', (like_query, like_query, limit))

        rows = cursor.fetchall()
        results = []

        for row in rows:
            doc_id, filepath, filename, extension, filesize, modified_time, content, index_root = row

            # 生成关键词命中摘要
            snippets = self._extract_snippets(content, query_words)

            # 检查文件是否实际存在
            file_exists = os.path.exists(filepath)

            results.append({
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

        return results

    def _extract_snippets(self, content: str, query_words: List[str]) -> List[str]:
        """从文档内容中提取包含关键词的上下文摘要"""
        if not content or not query_words:
            return []

        content_lower = content.lower()
        snippets = []
        seen_positions = set()

        for word in query_words:
            word_lower = word.lower()
            pos = 0
            while len(snippets) < MAX_SNIPPETS:
                idx = content_lower.find(word_lower, pos)
                if idx == -1:
                    break

                # 避免重叠摘要
                bucket = idx // SNIPPET_CONTEXT
                if bucket in seen_positions:
                    pos = idx + 1
                    continue
                seen_positions.add(bucket)

                # 提取上下文
                start = max(0, idx - SNIPPET_CONTEXT // 2)
                end = min(len(content), idx + len(word) + SNIPPET_CONTEXT // 2)
                snippet = content[start:end].strip()

                # 高亮关键词（用 ** 标记）
                for qw in query_words:
                    snippet = snippet.replace(qw, f'**{qw}**')
                    snippet = snippet.replace(qw.lower(), f'**{qw.lower()}**')

                if start > 0:
                    snippet = '...' + snippet
                if end < len(content):
                    snippet = snippet + '...'

                snippets.append(snippet)
                pos = idx + 1

        return snippets

    def get_document_content(self, doc_id: int) -> Optional[str]:
        """获取文档的完整内容"""
        conn = self._get_conn()
        cursor = conn.cursor()
        cursor.execute("SELECT content FROM documents WHERE id=?", (doc_id,))
        row = cursor.fetchone()
        return row[0] if row else None

    def close(self):
        """关闭数据库连接"""
        if hasattr(self._local, 'conn') and self._local.conn:
            self._local.conn.close()
            self._local.conn = None


class IndexBuilder:
    """
    索引构建器：扫描目录，多线程提取文档内容，写入索引
    """

    def __init__(self, engine: IndexEngine):
        self.engine = engine
        self._stop_flag = threading.Event()

    def stop(self):
        """停止正在进行的索引任务"""
        self._stop_flag.set()

    def build_index(
        self,
        root_dir: str,
        enabled_extensions: List[str],
        enable_pdf: bool = True,
        enable_ocr: bool = False,
        max_workers: int = 4,
        progress_callback: Optional[Callable] = None,
        log_callback: Optional[Callable] = None,
    ) -> Dict:
        """
        构建/更新索引
        
        Args:
            root_dir: 要索引的根目录
            enabled_extensions: 启用的文件扩展名列表，如 ['.docx', '.pdf']
            enable_pdf: 是否处理 PDF
            enable_ocr: 是否对扫描版 PDF 启用 OCR
            max_workers: 并行线程数
            progress_callback: 进度回调 (current, total, filepath)
            log_callback: 日志回调 (message)
        
        Returns:
            统计信息字典
        """
        self._stop_flag.clear()

        def log(msg):
            logger.info(msg)
            if log_callback:
                log_callback(msg)

        log(f"开始扫描目录: {root_dir}")

        # 扫描所有符合条件的文件
        all_files = []
        for dirpath, dirnames, filenames in os.walk(root_dir):
            if self._stop_flag.is_set():
                break
            # 跳过隐藏目录
            dirnames[:] = [d for d in dirnames if not d.startswith('.')]
            for filename in filenames:
                ext = os.path.splitext(filename)[1].lower()
                if ext in enabled_extensions:
                    all_files.append(os.path.join(dirpath, filename))

        total = len(all_files)
        log(f"共找到 {total} 个文件")

        if total == 0:
            return {'total': 0, 'added': 0, 'updated': 0, 'skipped': 0, 'failed': 0}

        # 获取已索引文件的哈希
        indexed = self.engine.get_indexed_files()

        # 筛选需要处理的文件（新文件或已修改的文件）
        to_process = []
        skipped = 0
        for fp in all_files:
            fhash = get_file_hash(fp)
            if fp in indexed and indexed[fp] == fhash:
                skipped += 1
            else:
                to_process.append(fp)

        log(f"需要处理: {len(to_process)} 个，跳过（未变化）: {skipped} 个")

        stats = {'total': total, 'added': 0, 'updated': 0, 'skipped': skipped, 'failed': 0}
        processed = 0
        lock = threading.Lock()

        def process_file(filepath):
            nonlocal processed
            if self._stop_flag.is_set():
                return

            try:
                filename = os.path.basename(filepath)
                ext = os.path.splitext(filename)[1].lower()
                stat = os.stat(filepath)
                filesize = stat.st_size
                modified_time = stat.st_mtime
                file_hash = get_file_hash(filepath)

                # 提取文本内容
                content = extract_text(
                    filepath,
                    enable_pdf=(enable_pdf and ext == '.pdf'),
                    enable_ocr=enable_ocr
                )

                # 写入索引
                is_update = filepath in indexed
                self.engine.add_document(
                    filepath=filepath,
                    filename=filename,
                    extension=ext,
                    filesize=filesize,
                    modified_time=modified_time,
                    file_hash=file_hash,
                    content=content,
                    index_root=root_dir
                )

                with lock:
                    processed += 1
                    if is_update:
                        stats['updated'] += 1
                    else:
                        stats['added'] += 1
                    if progress_callback:
                        progress_callback(processed + skipped, total, filepath)

            except Exception as e:
                with lock:
                    stats['failed'] += 1
                    processed += 1
                    if log_callback:
                        log_callback(f"处理失败: {os.path.basename(filepath)} - {e}")
                    if progress_callback:
                        progress_callback(processed + skipped, total, filepath)

        # 多线程处理
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(process_file, fp) for fp in to_process]
            for future in as_completed(futures):
                if self._stop_flag.is_set():
                    executor.shutdown(wait=False, cancel_futures=True)
                    break
                try:
                    future.result()
                except Exception:
                    pass

        # 更新索引信息
        self.engine.set_index_info('last_index_time', str(time.time()))
        self.engine.set_index_info('index_root', root_dir)

        log(f"索引完成！新增: {stats['added']}，更新: {stats['updated']}，"
            f"跳过: {stats['skipped']}，失败: {stats['failed']}")

        return stats
