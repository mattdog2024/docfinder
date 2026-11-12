"""DOC regressions: one real LibreOffice fixture plus explicitly synthetic streams.

Run: python -m pytest -q tests/test_legacy_extract.py
Optional frozen integration: DOCFINDER_TEST_PYINSTALLER=1 python -m pytest -q
    tests/test_legacy_extract.py
No Word, network downloads, or user documents are used by these tests.
"""
import importlib.util
import os
from pathlib import Path
import struct
import subprocess
import sys
import zipfile

import pytest

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
from core import extractor
from third_party import THIRD_PARTY_NOTICES, olefile

REAL_DOC = PROJECT / 'third_party' / 'test_fixtures' / 'mixed_table.doc'


@pytest.fixture(autouse=True)
def no_word(monkeypatch):
    monkeypatch.setattr(extractor, '_WIN32COM_AVAILABLE', False)


def synthetic_streams(pieces, *, body_chars=None, table_one=True, prc=False):
    """Synthetic FIB + CLX, NOT a real Office-produced DOC/OLE container.

    Pieces are (text, compressed), stored in reverse physical order so CP order
    differs from the order a naive whole-stream scan would produce.
    """
    wd = bytearray(1024 + len(pieces) * 1024)
    struct.pack_into('<HH', wd, 0, 0xA5EC, 0x00C1)
    struct.pack_into('<H', wd, 10, 0x0200 if table_one else 0)
    struct.pack_into('<H', wd, 32, 14)  # csw
    struct.pack_into('<H', wd, 62, 22)  # cslw
    lengths = [len(text) if compressed else len(text.encode('utf-16le')) // 2
               for text, compressed in pieces]
    struct.pack_into('<I', wd, 76, sum(lengths) if body_chars is None else body_chars)
    struct.pack_into('<H', wd, 152, 93)  # cbRgFcLcb: count of pairs
    # A convincing non-text metadata string: must never be returned.
    noise = 'MetadataOnly 不应进入正文'.encode('utf-16le')
    wd[920:920 + len(noise)] = noise
    cps = [0]
    pcds = []
    for i, ((text, compressed), count) in enumerate(zip(pieces, lengths)):
        fc = 1024 + (len(pieces) - i - 1) * 1024
        raw = text.encode('cp1252') if compressed else text.encode('utf-16le')
        wd[fc:fc + len(raw)] = raw
        encoded_fc = (fc * 2 | 0x40000000) if compressed else fc
        pcds.append(struct.pack('<HIH', 0, encoded_fc, 0))
        cps.append(cps[-1] + count)
    plc = struct.pack('<' + 'I' * len(cps), *cps) + b''.join(pcds)
    clx = (b'\x01\x02\x00\x00\x00' if prc else b'')
    clx += b'\x02' + struct.pack('<I', len(plc)) + plc
    table = bytearray(b'NOT-TEXT' + clx)
    struct.pack_into('<II', wd, 154 + 33 * 8, 8, len(clx))
    return wd, table


def test_vendor_is_static_local_and_license_is_complete():
    assert olefile.__version__ == '0.47'
    assert Path(olefile.__file__).resolve().is_relative_to(PROJECT / 'third_party')
    assert THIRD_PARTY_NOTICES == (PROJECT / 'third_party/olefile/LICENSE.txt').read_text()
    assert 'Philippe Lagadec' in THIRD_PARTY_NOTICES
    assert 'Secret Labs AB' in THIRD_PARTY_NOTICES


def test_real_libreoffice_mixed_text_and_table(monkeypatch):
    # Fail if DOC ever attempts to check/use COM; this is an actual CFB DOC.
    def forbid_com():
        raise AssertionError('DOC extraction must not invoke Word')
    monkeypatch.setattr(extractor, '_check_win32com', forbid_com)
    assert olefile.isOleFile(str(REAL_DOC))
    text = extractor.extract_doc(str(REAL_DOC))
    assert text.startswith('中文 English 混合文档，测试检索 Alpha Beta 2026。')
    assert text.count('重复段落 Repeat paragraph.') == 2
    for token in ('项目 Item', '数量 Qty', '苹果 Apple', '12', '香蕉 Banana', '34',
                  '结束 End — café “quoted”.'):
        assert token in text
    assert text.index('苹果 Apple') < text.index('12') < text.index('香蕉 Banana')
    assert 'DocFinder legacy regression' not in text  # document title metadata
    assert extractor.extract_text(str(REAL_DOC)) == text


@pytest.mark.parametrize('table_one', [False, True])
def test_synthetic_noncontiguous_mixed_pieces_and_prc(table_one):
    wd, table = synthetic_streams([
        ('English café “quoted” —\r', True),
        ('中文\r单\x0712\r', False),
        ('Tail\r', True),
    ], table_one=table_one, prc=True)
    assert extractor._doc_fib_info(wd)[0] == ('1Table' if table_one else '0Table')
    assert extractor._extract_doc_piece_table(wd, table) == (
        'English café “quoted” —\n中文\n单\t12\nTail')


def test_synthetic_only_main_story_not_header_or_stale_text():
    wd, table = synthetic_streams([('正文 Main\r', False), ('Header NotMain\r', True)],
                                  body_chars=8)
    assert extractor._extract_doc_piece_table(wd, table) == '正文 Main'


def test_synthetic_short_repeated_and_numeric_cells():
    wd, table = synthetic_streams([('单\r单\r12\x0734\r', False)])
    assert extractor._extract_doc_piece_table(wd, table) == '单\n单\n12\t34'


def test_synthetic_unicode_surrogate_across_pieces():
    wd, table = synthetic_streams([('😀', False), ('x', True)])
    # Reuse storage of the second piece for the UTF-16 low surrogate.
    wd[1024:1026] = b'\x00\xde'
    # Replace PLC with three pieces: high surrogate, low surrogate, ASCII x.
    wd[1100:1101] = b'x'
    plc = struct.pack('<4I', 0, 1, 2, 3)
    plc += struct.pack('<HIH', 0, 2048, 0)
    plc += struct.pack('<HIH', 0, 1024, 0)
    plc += struct.pack('<HIH', 0, 1100 * 2 | 0x40000000, 0)
    table = bytearray(b'NOT-TEXT\x02' + struct.pack('<I', len(plc)) + plc)
    struct.pack_into('<II', wd, 418, 8, len(table) - 8)
    assert extractor._extract_doc_piece_table(wd, table) == '😀x'


def test_synthetic_content_limit(monkeypatch):
    wd, table = synthetic_streams([('abcdef', True), ('中文后续', False)])
    monkeypatch.setattr(extractor, 'MAX_CONTENT_CHARS', 7)
    assert extractor._extract_doc_piece_table(wd, table) == 'abcdef中'


@pytest.mark.parametrize('damage', [
    'bad_magic', 'pre97', 'encrypted', 'short_fib', 'missing_fc_pairs',
    'clx_out_of_bounds', 'bad_marker', 'bad_plc_size', 'truncated_clx',
    'bad_prc', 'negative_prc', 'bad_cp_start', 'nonascending_cp',
    'text_out_of_bounds', 'odd_compressed_fc', 'short_main_story',
])
def test_synthetic_malformed_rejected_without_scanning(damage):
    wd, table = synthetic_streams([('Valid English\r', True)])
    if damage == 'bad_magic':
        wd[0:2] = b'xx'
    elif damage == 'pre97':
        struct.pack_into('<H', wd, 2, 0x0065)
    elif damage == 'encrypted':
        struct.pack_into('<H', wd, 10, 0x0300)
    elif damage == 'short_fib':
        wd = wd[:40]
    elif damage == 'missing_fc_pairs':
        struct.pack_into('<H', wd, 152, 33)
    elif damage == 'clx_out_of_bounds':
        struct.pack_into('<I', wd, 418, len(table) + 1)
    elif damage == 'bad_marker':
        table[8] = 3
    elif damage == 'bad_plc_size':
        struct.pack_into('<I', table, 9, 11)
    elif damage == 'truncated_clx':
        table = table[:-1]
    elif damage == 'bad_prc':
        table[8:11] = b'\x01\xff\x7f'
    elif damage == 'negative_prc':
        table[8:11] = b'\x01\xff\xff'
    elif damage == 'bad_cp_start':
        struct.pack_into('<I', table, 13, 1)
    elif damage == 'nonascending_cp':
        struct.pack_into('<I', table, 17, 0)
    elif damage == 'text_out_of_bounds':
        struct.pack_into('<I', table, 23, 0x40000000 | (len(wd) * 2))
    elif damage == 'odd_compressed_fc':
        struct.pack_into('<I', table, 23, 0x40000000 | 2049)
    elif damage == 'short_main_story':
        struct.pack_into('<I', wd, 76, 999)
    with pytest.raises((ValueError, struct.error)):
        extractor._extract_doc_piece_table(wd, table)


def test_non_ole_missing_and_truncated_inputs(tmp_path):
    bad = tmp_path / 'bad.doc'
    bad.write_bytes('Not OLE 中文'.encode())
    assert extractor.extract_doc(str(bad)) == ''
    assert extractor.extract_doc(str(tmp_path / 'missing.doc')) == ''
    bad.write_bytes(REAL_DOC.read_bytes()[:512])
    assert extractor.extract_doc(str(bad)) == ''


@pytest.mark.parametrize('damage', ['encrypted', 'missing_table', 'out_of_bounds'])
def test_real_ole_fail_closed_and_resources_closed(tmp_path, monkeypatch, damage):
    path = tmp_path / 'bad.doc'
    path.write_bytes(REAL_DOC.read_bytes())
    if damage != 'missing_table':
        with olefile.OleFileIO(str(path), write_mode=True) as doc:
            wd = bytearray(doc.openstream('WordDocument').read())
            if damage == 'encrypted':
                flags = struct.unpack_from('<H', wd, 10)[0]
                struct.pack_into('<H', wd, 10, flags | 0x0100)
            else:
                struct.pack_into('<I', wd, 418, 0xFFFFFFFF)
            doc.write_stream('WordDocument', bytes(wd))
    else:
        original = olefile.OleFileIO.exists
        monkeypatch.setattr(olefile.OleFileIO, 'exists',
                            lambda self, name: False if name == '1Table' else original(self, name))
    assert extractor.extract_doc(str(path)) == ''
    path.unlink()  # handle must be closed, especially on Windows


def test_other_xml_formats_unchanged(tmp_path):
    # XML-only smoke tests: no third-party Office dependency required.
    samples = [
        ('docx', 'word/document.xml', '<w:p><w:r><w:t>中文 Alpha</w:t></w:r></w:p>',
         extractor.extract_docx),
        ('pptx', 'ppt/slides/slide1.xml', '<a:t>中文 Alpha</a:t>', extractor.extract_pptx),
    ]
    for extension, member, xml, extract in samples:
        path = tmp_path / ('sample.' + extension)
        with zipfile.ZipFile(path, 'w') as z:
            z.writestr(member, xml)
        assert extract(str(path)) == '中文 Alpha'


def test_ppt_uses_same_vendored_olefile(tmp_path, monkeypatch):
    # Synthetic fake OLE stream; this checks dependency resolution, NOT PPT fidelity.
    path = tmp_path / 'fake.ppt'
    path.write_bytes(b'dummy')
    class FakeOle:
        def listdir(self):
            return [['PowerPoint Document']]
        def openstream(self, name):
            import io
            return io.BytesIO('English fallback 中文测试'.encode('utf-16le'))
        def close(self):
            pass
    monkeypatch.setattr(olefile, 'isOleFile', lambda p: True)
    monkeypatch.setattr(olefile, 'OleFileIO', lambda p: FakeOle())
    assert 'English fallback 中文测试' in extractor.extract_ppt(str(path))


@pytest.mark.skipif(os.environ.get('DOCFINDER_TEST_PYINSTALLER') != '1',
                    reason='set DOCFINDER_TEST_PYINSTALLER=1 for frozen integration')
def test_pyinstaller_collects_vendor_without_hidden_import(tmp_path):
    if importlib.util.find_spec('PyInstaller') is None:
        pytest.skip('PyInstaller not installed')
    script = tmp_path / 'legacy_smoke.py'
    script.write_text(
        'import sys\n'
        'from core.extractor import extract_doc\n'
        'from third_party import THIRD_PARTY_NOTICES\n'
        'assert "Secret Labs AB" in THIRD_PARTY_NOTICES\n'
        'text = extract_doc(sys.argv[1])\n'
        'assert "中文 English" in text and "苹果 Apple" in text and "34" in text\n'
        'print("FROZEN_DOC_OK", len(text))\n', encoding='utf-8')
    built = subprocess.run([
        sys.executable, '-m', 'PyInstaller', '--clean', '--noconfirm', '--onefile',
        '--name', 'legacy_smoke', '--paths', str(PROJECT), '--noupx',
        '--distpath', str(tmp_path / 'dist'), '--workpath', str(tmp_path / 'build'),
        '--specpath', str(tmp_path), str(script),
    ], cwd=tmp_path, text=True, capture_output=True, timeout=180)
    assert built.returncode == 0, built.stdout + built.stderr
    executable = tmp_path / 'dist' / ('legacy_smoke.exe' if os.name == 'nt' else 'legacy_smoke')
    result = subprocess.run([str(executable), str(REAL_DOC)], cwd=tmp_path,
                            text=True, capture_output=True, timeout=30)
    assert result.returncode == 0, result.stdout + result.stderr
    assert 'FROZEN_DOC_OK' in result.stdout
    toc = (tmp_path / 'build/legacy_smoke/Analysis-00.toc').read_text()
    assert 'third_party.olefile.olefile' in toc
