# -*- coding: utf-8 -*-
import os
import sqlite3
import subprocess
import sys
import tempfile
import threading
import time
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from core.indexer import IndexEngine, IndexBuilder
import core.indexer as module


def make_docx(path, text='文明单位创建 女子监狱 志愿服务'):
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>' + text + '</w:t></w:r></w:p></w:body></w:document>')


class IndexTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.engine = IndexEngine(str(self.root / '索引.db'))
        self.builder = IndexBuilder(self.engine)

    def tearDown(self):
        self.engine.close()
        self.temp.cleanup()

    def corpus(self, count):
        for n in range(count):
            make_docx(self.root / f'文明资料{n:04d}.docx')

    def test_over_100_files_and_incremental(self):
        self.corpus(357)
        progress = []
        stats = self.builder.build_index(str(self.root), ['.docx'], max_workers=8, progress_callback=lambda c,t,p: progress.append(c))
        self.assertEqual(stats['added'], 357)
        self.assertEqual(progress[-1], 357)
        self.assertEqual(self.engine.get_index_info()['total_docs'], 357)
        self.assertEqual(len(self.engine.search('监狱', 500)), 357)
        self.assertEqual(len(self.engine.search('文明 志愿', 500)), 357)
        stats = self.builder.build_index(str(self.root), ['.docx'], max_workers=4)
        self.assertEqual(stats['skipped'], 357)
        self.assertEqual(stats['added'], 0)

    def test_write_failure_is_reported_not_deadlocked(self):
        self.corpus(400)
        start = time.monotonic()
        with patch.object(self.engine, 'batch_add_documents', side_effect=sqlite3.OperationalError('模拟磁盘写入故障')):
            with self.assertRaisesRegex(sqlite3.OperationalError, '模拟磁盘'):
                self.builder.build_index(str(self.root), ['.docx'], max_workers=8)
        self.assertLess(time.monotonic() - start, 3)
        self.assertEqual(self.engine.get_index_info()['total_docs'], 0)

    def test_sqlite_failure_preserves_old_committed_content(self):
        self.corpus(120)
        self.builder.build_index(str(self.root), ['.docx'], max_workers=4)
        for fp in self.root.glob('*.docx'):
            make_docx(fp, '新的文明创建台账内容')
        conn = self.engine._get_conn()
        conn.execute("CREATE TRIGGER deny_update BEFORE UPDATE ON documents BEGIN SELECT RAISE(ABORT, '模拟更新写入受阻'); END")
        conn.commit()
        with self.assertRaisesRegex(sqlite3.IntegrityError, '模拟更新'):
            self.builder.build_index(str(self.root), ['.docx'], max_workers=4)
        self.assertEqual(self.engine.get_index_info()['total_docs'], 120)
        self.assertEqual(len(self.engine.search('志愿', 500)), 120)

    def test_stop_during_hung_parser(self):
        self.corpus(200)
        entered, release = threading.Event(), threading.Event()
        def blocked(*args):
            entered.set()
            release.wait(10)
            return None
        stats = {}
        def build():
            stats.update(self.builder.build_index(str(self.root), ['.docx'], max_workers=4))
            self.engine.close()
        try:
            with patch.object(module, '_extract_single', blocked):
                thread = threading.Thread(target=build)
                thread.start()
                self.assertTrue(entered.wait(2))
                start = time.monotonic()
                self.builder.stop()
                thread.join(2)
                self.assertFalse(thread.is_alive())
                self.assertLess(time.monotonic()-start, 1.0)
                self.assertTrue(stats['stopped'])
        finally:
            release.set()

    def test_slow_first_file_does_not_block_completed_files(self):
        self.corpus(150)
        real = module._extract_single
        seen = []
        def slow(fp, *args):
            if fp.endswith('0000.docx'):
                time.sleep(0.5)
            return real(fp, *args)
        with patch.object(module, '_extract_single', slow):
            stats = self.builder.build_index(str(self.root), ['.docx'], max_workers=8, progress_callback=lambda c,t,p: seen.append(p))
        self.assertEqual(stats['added'], 150)
        self.assertFalse(seen[0].endswith('0000.docx'))

    def test_all_parsers_timeout_has_clear_error(self):
        self.corpus(20)
        release = threading.Event()
        def blocked(*args):
            release.wait(5)
            return None
        try:
            with patch.object(module, '_extract_single', blocked), patch.object(module, 'EXTRACT_TIMEOUT', 0.25):
                with self.assertRaisesRegex(RuntimeError, '已暂停索引'):
                    self.builder.build_index(str(self.root), ['.docx'], max_workers=2)
        finally:
            release.set()

    def test_old_external_content_db_migrates_without_losing_text(self):
        conn = self.engine._get_conn()
        conn.execute('DROP TABLE documents_fts')
        conn.execute('CREATE VIRTUAL TABLE documents_fts USING fts5(filepath UNINDEXED,filename,content_tokens,content="documents",content_rowid="id")')
        conn.execute('INSERT INTO documents(filepath,filename,extension,content) VALUES(?,?,?,?)', ('离线.doc','旧资料.doc','.doc','女子监狱文明创建档案'))
        conn.commit()
        self.engine.close()
        self.engine = IndexEngine(str(self.root / '索引.db'))
        result = self.engine.search('监狱')
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]['content'], '女子监狱文明创建档案')
        self.assertFalse(result[0]['file_exists'])
        self.assertTrue(result[0]['snippets'])
        self.engine.batch_add_documents([('离线.doc','旧资料.doc','.doc',1,1,'new','文明创建志愿活动')], str(self.root))
        self.assertTrue(self.engine.search('志愿'))

    def test_search_while_indexing_with_separate_connection(self):
        self.corpus(230)
        reader = IndexEngine(self.engine.db_path)
        found = []
        def progress(c,t,p):
            if c == 110:
                found.extend(reader.search('文明', 500))
        try:
            self.builder.build_index(str(self.root), ['.docx'], max_workers=4, progress_callback=progress)
            self.assertGreater(len(found), 0)
        finally:
            reader.close()

    def test_legacy_raw_tokens_multiword_fallback(self):
        self.engine.batch_add_documents([('offline.doc','旧资料.doc','.doc',1,1,'h','女子监狱文明单位创建')], str(self.root))
        conn = self.engine._get_conn()
        conn.execute('UPDATE documents_fts SET content_tokens=?', ('女子监狱文明单位创建',))
        conn.commit()
        self.assertEqual(len(self.engine.search('监狱 文明')), 1)

    def test_pool_does_not_prevent_process_exit(self):
        code = '''import time
from core.extraction_pool import ExtractionPool
p=ExtractionPool(2)
p.submit(time.sleep, 60)
time.sleep(.1)
p.shutdown(wait=False, cancel_futures=True)
'''
        result = subprocess.run([sys.executable, '-c', code], timeout=3, capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr.decode())


if __name__ == '__main__':
    unittest.main()
