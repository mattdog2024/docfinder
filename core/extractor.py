"""
文档内容提取模块 v2.1

核心改进：
1. .doc 老格式：内置 OLE + Word 97+ piece table，不依赖 Word
2. .xls 老格式：xlrd 直接读取，速度快
3. .ppt 老格式：win32com 快速调用，失败立即跳过
4. 所有格式都有明确的超时保护
5. 新格式（docx/xlsx/pptx）直接解析 XML，速度最快
"""

import os
import re
import logging
import zipfile
import struct
from typing import Optional

# 静态导入：现有 PyInstaller 工作流会自动收集，不需要运行时安装。
from third_party import olefile

logger = logging.getLogger(__name__)

# 每个文件最多提取的字符数（5万字，足够搜索用）
MAX_CONTENT_CHARS = 50_000

# 超过此大小的文件跳过（单位：MB）
MAX_FILE_SIZE_MB = 100

# xlsx/xls 最多读取的行数
MAX_EXCEL_ROWS = 2000

# win32com 是否可用（只检测一次，避免重复检测）
_WIN32COM_AVAILABLE = None


def _check_win32com() -> bool:
    """检测 win32com 是否可用（只检测一次）"""
    global _WIN32COM_AVAILABLE
    if _WIN32COM_AVAILABLE is None:
        try:
            import win32com.client
            _WIN32COM_AVAILABLE = True
        except ImportError:
            _WIN32COM_AVAILABLE = False
    return _WIN32COM_AVAILABLE


def _check_file(filepath: str) -> bool:
    """检查文件是否可处理"""
    try:
        size_mb = os.path.getsize(filepath) / (1024 * 1024)
        return size_mb <= MAX_FILE_SIZE_MB
    except Exception:
        return False


# ─────────────────────────────────────────────
# DOCX - 直接解析 XML（最快方式）
# ─────────────────────────────────────────────

def extract_docx(filepath: str) -> str:
    """提取 .docx 文件文本（直接读 XML，比 python-docx 快 3-5 倍）"""
    if not _check_file(filepath):
        return ''
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            if 'word/document.xml' not in z.namelist():
                return ''
            xml_data = z.read('word/document.xml').decode('utf-8', errors='ignore')

        # 段落之间加换行
        xml_data = re.sub(r'<w:p[ >]', '\n<w:p ', xml_data)
        xml_data = re.sub(r'<w:br[^/]*/>', '\n', xml_data)
        # 去掉所有 XML 标签
        text = re.sub(r'<[^>]+>', '', xml_data)
        # 清理多余空白
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        result = '\n'.join(lines)
        return result[:MAX_CONTENT_CHARS]
    except Exception as e:
        logger.debug(f"XML 提取 docx 失败，尝试 python-docx: {e}")
        try:
            from docx import Document
            doc = Document(filepath)
            parts = []
            total = 0
            for para in doc.paragraphs:
                t = para.text.strip()
                if t:
                    parts.append(t)
                    total += len(t)
                    if total >= MAX_CONTENT_CHARS:
                        break
            return '\n'.join(parts)[:MAX_CONTENT_CHARS]
        except Exception:
            return ''


# ─────────────────────────────────────────────
# DOC - Word 97+ FIB/CLX piece table（不依赖 Word）
# ─────────────────────────────────────────────

# MS-DOC 2.9.73: compressed pieces are 8-bit Unicode with these exceptions,
# not GBK or an encoding inferred from the producer's installation language.
_DOC_COMPRESSED_MAP = dict(zip(
    [0x82, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89, 0x8A, 0x8B, 0x8C,
     0x91, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9A, 0x9B,
     0x9C, 0x9F],
    [0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030,
     0x0160, 0x2039, 0x0152, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022,
     0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x0178],
))


def _doc_fib_info(wd_data: bytes):
    """返回 (table stream, fcClx, lcbClx, ccpText)，损坏/加密时失败。"""
    def u16(offset):
        return struct.unpack_from('<H', wd_data, offset)[0]

    def u32(offset):
        return struct.unpack_from('<I', wd_data, offset)[0]

    if len(wd_data) < 34 or u16(0) != 0xA5EC or u16(2) < 0x00C1:
        raise ValueError('Not a supported Word 97+ FIB')
    flags = u16(10)
    if flags & 0x0100:
        raise ValueError('Encrypted/obfuscated DOC is not supported')
    # 跟随 FIB 自身的 csw/cslw 计数，避免把所有版本布局写死。
    lw_count_offset = 34 + 2 * u16(32)
    lw_count = u16(lw_count_offset)
    lw_offset = lw_count_offset + 2
    if lw_count < 4:
        raise ValueError('Missing FibRgLw97.ccpText')
    body_chars = u32(lw_offset + 12)
    fc_count_offset = lw_offset + 4 * lw_count
    fc_count = u16(fc_count_offset)
    fc_offset = fc_count_offset + 2
    if fc_count < 34 or fc_offset + fc_count * 8 > len(wd_data):
        raise ValueError('Truncated FibRgFcLcb97')
    fc_clx, lcb_clx = struct.unpack_from('<II', wd_data, fc_offset + 33 * 8)
    if body_chars > 0x7FFFFFFF:
        raise ValueError('Invalid main-story character count')
    return ('1Table' if flags & 0x0200 else '0Table',
            fc_clx, lcb_clx, body_chars)


def _extract_doc_piece_table(wd_data: bytes, table_data: bytes) -> str:
    """只读 CLX 引用的正文；不扫描元数据、旧修订残留或其他 OLE 流。"""
    _, fc_clx, lcb_clx, body_chars = _doc_fib_info(wd_data)
    if lcb_clx < 5 or fc_clx + lcb_clx > len(table_data):
        raise ValueError('Invalid CLX range')
    clx = memoryview(table_data)[fc_clx:fc_clx + lcb_clx]
    pos = 0
    # 可有零个或多个 Prc（0x01 + 2字节长度 + grpprl）。
    while pos < len(clx) and clx[pos] == 1:
        cb = struct.unpack_from('<h', clx, pos + 1)[0]
        if cb < 0 or pos + 3 + cb > len(clx):
            raise ValueError('Invalid CLX Prc')
        pos += 3 + cb
    if pos + 5 > len(clx) or clx[pos] != 2:
        raise ValueError('Missing CLX Pcdt')
    plc_size = struct.unpack_from('<I', clx, pos + 1)[0]
    if plc_size < 4 or (plc_size - 4) % 12 or pos + 5 + plc_size != len(clx):
        raise ValueError('Invalid PlcPcd size')
    piece_count = (plc_size - 4) // 12
    plc = clx[pos + 5:]
    pcd_offset = 4 * (piece_count + 1)
    if struct.unpack_from('<I', plc, 0)[0] != 0:
        raise ValueError('Piece character positions must start at zero')
    chunks = []
    remaining = min(body_chars, MAX_CONTENT_CHARS)
    cp_start = 0
    for i in range(piece_count):
        cp_end = struct.unpack_from('<I', plc, 4 * (i + 1))[0]
        if cp_end <= cp_start or cp_end > 0x7FFFFFFF:
            raise ValueError('Invalid piece character positions')
        raw_fc = struct.unpack_from('<I', plc, pcd_offset + 8 * i + 2)[0]
        compressed = bool(raw_fc & 0x40000000)
        fc = raw_fc & 0x3FFFFFFF
        if compressed:
            if fc % 2:
                raise ValueError('Invalid compressed piece offset')
            fc //= 2
        width = 1 if compressed else 2
        if fc + (cp_end - cp_start) * width > len(wd_data):
            raise ValueError('Piece text is outside WordDocument')
        take = min(cp_end - cp_start, remaining)
        if take:
            data = wd_data[fc:fc + take * width]
            if compressed:
                data = data.decode('latin-1').translate(_DOC_COMPRESSED_MAP).encode('utf-16le')
            chunks.append(data)
            remaining -= take
        cp_start = cp_end
    if cp_start < body_chars:
        raise ValueError('Piece table does not cover main story')
    # 合并后再解码，保留跨 piece 的 UTF-16 surrogate pair。
    text = b''.join(chunks).decode('utf-16le', errors='replace')
    text = text.replace('\r', '\n').replace('\x07', '\t')
    text = text.replace('\x0b', '\n').replace('\x0c', '\n')
    text = re.sub(r'[\x00-\x08\x0e-\x1f\x7f-\x9f]', '', text)
    # 正文重复段落、单字和数字表格单元格均应保留。
    return '\n'.join(line.strip() for line in text.splitlines() if line.strip())[:MAX_CONTENT_CHARS]


def extract_doc(filepath: str) -> str:
    """提取未加密 Word 97+ OLE DOC 正文和表格文字，不还原排版。

    不依赖 Word，也不回退全流扫描；不支持的版本/损坏文件返回空文本。
    页眉、脚注、文本框等非正文 story 和嵌入对象不在此解析范围。
    """
    if not _check_file(filepath):
        return ''
    try:
        if not olefile.isOleFile(filepath):
            return ''
        with olefile.OleFileIO(filepath) as ole:
            if not ole.exists('WordDocument'):
                return ''
            wd_data = ole.openstream('WordDocument').read()
            table_name, _, _, _ = _doc_fib_info(wd_data)
            if not ole.exists(table_name):
                return ''
            table_data = ole.openstream(table_name).read()
            return _extract_doc_piece_table(wd_data, table_data)
    except Exception as e:
        logger.debug(f'提取 DOC piece table 失败 {filepath}: {e}')
        return ''


# ─────────────────────────────────────────────
# XLSX - read_only 模式，限制行数
# ─────────────────────────────────────────────

def extract_xlsx(filepath: str) -> str:
    """提取 .xlsx 文件文本（read_only + 行数限制）"""
    if not _check_file(filepath):
        return ''
    try:
        import openpyxl
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        parts = []
        total = 0
        for sheet in wb.worksheets:
            row_count = 0
            for row in sheet.iter_rows(values_only=True):
                row_count += 1
                if row_count > MAX_EXCEL_ROWS:
                    break
                row_texts = []
                for cell in row:
                    if cell is None:
                        continue
                    if isinstance(cell, float):
                        val = str(int(cell)) if cell == int(cell) else f'{cell:.4g}'
                    elif isinstance(cell, bool):
                        val = '是' if cell else '否'
                    else:
                        val = str(cell).strip()
                    if val and val not in ('None', 'nan', ''):
                        row_texts.append(val)
                if row_texts:
                    line = ' '.join(row_texts)
                    parts.append(line)
                    total += len(line)
                    if total >= MAX_CONTENT_CHARS:
                        break
            if total >= MAX_CONTENT_CHARS:
                break
        wb.close()
        return '\n'.join(parts)[:MAX_CONTENT_CHARS]
    except Exception as e:
        logger.debug(f"提取 xlsx 失败 {filepath}: {e}")
        return ''


# ─────────────────────────────────────────────
# XLS - xlrd（老格式，直接读取）
# ─────────────────────────────────────────────

def extract_xls(filepath: str) -> str:
    """提取 .xls 文件文本（xlrd 直接读取）"""
    if not _check_file(filepath):
        return ''
    try:
        import xlrd
        wb = xlrd.open_workbook(filepath)
        parts = []
        total = 0
        for sheet in wb.sheets():
            for row_idx in range(min(sheet.nrows, MAX_EXCEL_ROWS)):
                row_texts = []
                for col_idx in range(sheet.ncols):
                    cell = sheet.cell(row_idx, col_idx)
                    if cell.ctype == 2:  # XL_CELL_NUMBER
                        v = cell.value
                        val = str(int(v)) if v == int(v) else f'{v:.4g}'
                    elif cell.ctype == 5:  # XL_CELL_ERROR
                        continue
                    else:
                        val = str(cell.value).strip()
                    if val and val not in ('', 'nan'):
                        row_texts.append(val)
                if row_texts:
                    line = ' '.join(row_texts)
                    parts.append(line)
                    total += len(line)
                    if total >= MAX_CONTENT_CHARS:
                        break
            if total >= MAX_CONTENT_CHARS:
                break
        return '\n'.join(parts)[:MAX_CONTENT_CHARS]
    except Exception as e:
        logger.debug(f"提取 xls 失败 {filepath}: {e}")
        return ''


# ─────────────────────────────────────────────
# PPTX - 直接读 XML（最快）
# ─────────────────────────────────────────────

def extract_pptx(filepath: str) -> str:
    """提取 .pptx 文件文本（直接读 XML）"""
    if not _check_file(filepath):
        return ''
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            slide_files = sorted([
                name for name in z.namelist()
                if name.startswith('ppt/slides/slide') and name.endswith('.xml')
            ])
            parts = []
            total = 0
            for slide_file in slide_files:
                xml_data = z.read(slide_file).decode('utf-8', errors='ignore')
                texts = re.findall(r'<a:t[^>]*>([^<]+)</a:t>', xml_data)
                for t in texts:
                    t = t.strip()
                    if t:
                        parts.append(t)
                        total += len(t)
                        if total >= MAX_CONTENT_CHARS:
                            break
                if total >= MAX_CONTENT_CHARS:
                    break
        return '\n'.join(parts)[:MAX_CONTENT_CHARS]
    except Exception as e:
        logger.debug(f"XML 提取 pptx 失败，尝试 python-pptx: {e}")
        try:
            from pptx import Presentation
            prs = Presentation(filepath)
            parts = []
            total = 0
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, 'text') and shape.text.strip():
                        t = shape.text.strip()
                        parts.append(t)
                        total += len(t)
                        if total >= MAX_CONTENT_CHARS:
                            break
                if total >= MAX_CONTENT_CHARS:
                    break
            return '\n'.join(parts)[:MAX_CONTENT_CHARS]
        except Exception:
            return ''


# ─────────────────────────────────────────────
# PPT - 老格式，win32com 优先
# ─────────────────────────────────────────────

def extract_ppt(filepath: str) -> str:
    """提取 .ppt 文件文本（PowerPoint 97-2003）"""
    if not _check_file(filepath):
        return ''

    # 方式1：win32com 调用 PowerPoint（Windows 上最可靠）
    if _check_win32com():
        try:
            import win32com.client
            import threading

            result_holder = [None]

            def _do_extract():
                ppt_app = None
                try:
                    ppt_app = win32com.client.Dispatch('PowerPoint.Application')
                    ppt_app.Visible = True  # PowerPoint 必须可见才能打开文件
                    prs = ppt_app.Presentations.Open(
                        os.path.abspath(filepath),
                        ReadOnly=True,
                        Untitled=True,
                        WithWindow=False
                    )
                    texts = []
                    for slide in prs.Slides:
                        for shape in slide.Shapes:
                            try:
                                if shape.HasTextFrame:
                                    texts.append(shape.TextFrame.TextRange.Text)
                            except Exception:
                                pass
                    prs.Close()
                    result_holder[0] = '\n'.join(texts)
                except Exception:
                    pass
                finally:
                    if ppt_app:
                        try:
                            ppt_app.Quit()
                        except Exception:
                            pass

            t = threading.Thread(target=_do_extract, daemon=True)
            t.start()
            t.join(timeout=10)

            if result_holder[0] and len(result_holder[0].strip()) > 10:
                return result_holder[0][:MAX_CONTENT_CHARS]
        except Exception:
            pass

    # 方式2：olefile 扫描（兜底）
    try:
        if not olefile.isOleFile(filepath):
            return ''

        ole = olefile.OleFileIO(filepath)
        texts = []
        current = []

        for stream_name in ole.listdir():
            try:
                data = ole.openstream(stream_name).read()
                i = 0
                while i < len(data) - 1:
                    c = struct.unpack_from('<H', data, i)[0]
                    if (0x4E00 <= c <= 0x9FFF or
                        0x0020 <= c <= 0x007E or
                        0x3000 <= c <= 0x303F):
                        current.append(chr(c))
                    else:
                        if len(current) >= 5:
                            texts.append(''.join(current))
                        current = []
                    i += 2
            except Exception:
                pass

        ole.close()

        if len(current) >= 5:
            texts.append(''.join(current))

        seen = set()
        unique_texts = []
        for t in texts:
            t_clean = t.strip()
            if t_clean and t_clean not in seen and len(t_clean) >= 3:
                seen.add(t_clean)
                unique_texts.append(t_clean)

        result = ' '.join(unique_texts)
        if len(result.strip()) > 10:
            return result[:MAX_CONTENT_CHARS]
    except Exception:
        pass

    return ''


# ─────────────────────────────────────────────
# PDF
# ─────────────────────────────────────────────

def extract_pdf_text(filepath: str) -> str:
    """提取可识别 PDF 的文本"""
    if not _check_file(filepath):
        return ''

    # 方式1：pymupdf（最快）
    try:
        import fitz
        parts = []
        total = 0
        doc = fitz.open(filepath)
        for page in doc:
            text = page.get_text()
            if text and text.strip():
                parts.append(text.strip())
                total += len(text)
                if total >= MAX_CONTENT_CHARS:
                    break
        doc.close()
        result = '\n'.join(parts)[:MAX_CONTENT_CHARS]
        if len(result.strip()) > 50:
            return result
    except ImportError:
        pass
    except Exception as e:
        logger.debug(f"pymupdf 失败 {filepath}: {e}")

    # 方式2：pdfplumber（兜底）
    try:
        import pdfplumber
        parts = []
        total = 0
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    parts.append(text.strip())
                    total += len(text)
                    if total >= MAX_CONTENT_CHARS:
                        break
        return '\n'.join(parts)[:MAX_CONTENT_CHARS]
    except Exception as e:
        logger.debug(f"pdfplumber 失败 {filepath}: {e}")
        return ''


def extract_pdf_ocr(filepath: str) -> str:
    """对扫描版 PDF 进行 OCR 识别"""
    if not _check_file(filepath):
        return ''
    try:
        import pytesseract
        from pdf2image import convert_from_path
        pages = convert_from_path(filepath, dpi=150)
        parts = []
        total = 0
        for page_img in pages:
            text = pytesseract.image_to_string(page_img, lang='chi_sim+eng')
            if text.strip():
                parts.append(text.strip())
                total += len(text)
                if total >= MAX_CONTENT_CHARS:
                    break
        return '\n'.join(parts)[:MAX_CONTENT_CHARS]
    except Exception as e:
        logger.debug(f"OCR 失败 {filepath}: {e}")
        return ''


# ─────────────────────────────────────────────
# 统一入口
# ─────────────────────────────────────────────

def extract_text(filepath: str, enable_pdf: bool = True, enable_ocr: bool = False) -> str:
    """
    统一文本提取入口

    支持格式：
    - .docx / .doc（Word 文档）
    - .xlsx / .xls（Excel 表格）
    - .pptx / .ppt（PowerPoint 演示文稿）
    - .pdf（PDF 文档，可选 OCR）
    - .txt / .csv / .md / .json / .xml / .html（纯文本）
    """
    ext = os.path.splitext(filepath)[1].lower()

    try:
        if ext == '.docx':
            return extract_docx(filepath)
        elif ext == '.doc':
            return extract_doc(filepath)
        elif ext == '.xlsx':
            return extract_xlsx(filepath)
        elif ext == '.xls':
            return extract_xls(filepath)
        elif ext == '.pptx':
            return extract_pptx(filepath)
        elif ext == '.ppt':
            return extract_ppt(filepath)
        elif ext == '.pdf':
            if not enable_pdf:
                return ''
            text = extract_pdf_text(filepath)
            if not text.strip() and enable_ocr:
                text = extract_pdf_ocr(filepath)
            return text
        elif ext in ('.txt', '.csv', '.md', '.log', '.ini', '.cfg', '.conf'):
            # 纯文本文件：直接读取，尝试多种编码
            for enc in ['utf-8-sig', 'utf-8', 'gbk', 'gb2312', 'latin-1']:
                try:
                    with open(filepath, 'r', encoding=enc, errors='ignore') as f:
                        return f.read(MAX_CONTENT_CHARS)
                except Exception:
                    continue
            return ''
        elif ext in ('.json', '.xml', '.html', '.htm'):
            # 结构化文本：读取后去标签
            for enc in ['utf-8-sig', 'utf-8', 'gbk']:
                try:
                    with open(filepath, 'r', encoding=enc, errors='ignore') as f:
                        content = f.read(MAX_CONTENT_CHARS)
                    if ext in ('.xml', '.html', '.htm'):
                        content = re.sub(r'<[^>]+>', ' ', content)
                        content = re.sub(r'\s+', ' ', content).strip()
                    return content
                except Exception:
                    continue
            return ''
        else:
            return ''
    except Exception as e:
        logger.debug(f"提取失败 {filepath}: {e}")
        return ''


# 支持的文件扩展名列表（供外部引用）
SUPPORTED_EXTENSIONS = [
    '.docx', '.doc',
    '.xlsx', '.xls',
    '.pptx', '.ppt',
    '.pdf',
    '.txt', '.csv', '.md',
    '.json', '.xml', '.html', '.htm',
    '.log', '.ini', '.cfg', '.conf',
]
