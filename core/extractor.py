# -*- coding: utf-8 -*-
"""
文档内容提取模块 v1.5
支持: docx, doc, xlsx, xls, pptx, ppt, pdf

速度优化原则：
1. 每种格式只提取"够用"的内容（前 50000 字符），不全量读取
2. docx: 用 zipfile 直接读 XML，比 python-docx 快 3-5 倍
3. xlsx: openpyxl read_only 模式，只读前 500 行
4. pptx: 只提取文本框，跳过图片/图表
5. 所有格式都有超时保护（signal 在 Windows 无效，改用内容长度限制）
6. doc/ppt 老格式：docx2txt 兜底（无需 LibreOffice）
"""

import os
import re
import logging
import zipfile
from typing import Optional

logger = logging.getLogger(__name__)

# 每个文件最多提取的字符数（5万字，足够搜索用，减少一半提升速度）
MAX_CONTENT_CHARS = 50_000

# 超过此大小的文件跳过（单位：MB）
MAX_FILE_SIZE_MB = 100

# xlsx/xls 最多读取的行数（超大表格只读前面部分）
MAX_EXCEL_ROWS = 2000


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

        # 去掉 XML 标签，提取纯文本
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
        # 降级：用 python-docx
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
# DOC - 老格式，多种方式兜底
# ─────────────────────────────────────────────

def extract_doc(filepath: str) -> str:
    """提取 .doc 文件文本"""
    if not _check_file(filepath):
        return ''

    # 方式1：docx2txt（纯 Python，速度快）
    try:
        import docx2txt
        text = docx2txt.process(filepath)
        if text and len(text.strip()) > 20:
            return text[:MAX_CONTENT_CHARS]
    except Exception:
        pass

    # 方式2：尝试用 python-docx 直接读（部分 .doc 可以）
    try:
        from docx import Document
        doc = Document(filepath)
        parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        if parts:
            return '\n'.join(parts)[:MAX_CONTENT_CHARS]
    except Exception:
        pass

    # 方式3：antiword（需要系统安装）
    try:
        import subprocess
        result = subprocess.run(
            ['antiword', filepath],
            capture_output=True, timeout=15
        )
        if result.returncode == 0:
            text = result.stdout.decode('utf-8', errors='ignore')
            if len(text.strip()) > 20:
                return text[:MAX_CONTENT_CHARS]
    except Exception:
        pass

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
# XLS - xlrd
# ─────────────────────────────────────────────

def extract_xls(filepath: str) -> str:
    """提取 .xls 文件文本"""
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
                # 提取 <a:t> 标签内容（文本框文字）
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
# PPT - 老格式
# ─────────────────────────────────────────────

def extract_ppt(filepath: str) -> str:
    """提取 .ppt 文件文本"""
    if not _check_file(filepath):
        return ''

    # 方式1：docx2txt
    try:
        import docx2txt
        text = docx2txt.process(filepath)
        if text and len(text.strip()) > 20:
            return text[:MAX_CONTENT_CHARS]
    except Exception:
        pass

    # 方式2：LibreOffice（慢，但兼容性好）
    try:
        import subprocess
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            result = subprocess.run(
                ['libreoffice', '--headless', '--convert-to', 'txt:Text',
                 '--outdir', tmpdir, filepath],
                capture_output=True, timeout=30
            )
            base = os.path.splitext(os.path.basename(filepath))[0]
            txt_path = os.path.join(tmpdir, base + '.txt')
            if os.path.exists(txt_path):
                with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read(MAX_CONTENT_CHARS)
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
        pages = convert_from_path(filepath, dpi=150)  # 降低 DPI 提升速度
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

def extract_text(filepath: str, enable_pdf: bool = True,
                 enable_ocr: bool = False) -> str:
    """
    统一入口：根据文件扩展名自动选择提取方式
    """
    ext = os.path.splitext(filepath)[1].lower()

    extractors = {
        '.docx': extract_docx,
        '.doc': extract_doc,
        '.xlsx': extract_xlsx,
        '.xls': extract_xls,
        '.pptx': extract_pptx,
        '.ppt': extract_ppt,
    }

    if ext in extractors:
        try:
            return extractors[ext](filepath) or ''
        except Exception as e:
            logger.warning(f"提取失败 {filepath}: {e}")
            return ''
    elif ext == '.pdf':
        if not enable_pdf:
            return ''
        text = extract_pdf_text(filepath)
        if len(text.strip()) < 50 and enable_ocr:
            ocr_text = extract_pdf_ocr(filepath)
            if len(ocr_text) > len(text):
                return ocr_text
        return text
    return ''


# 支持的文件格式
SUPPORTED_EXTENSIONS = {
    '.docx': 'Word 文档 (docx)',
    '.doc': 'Word 文档 (doc)',
    '.xlsx': 'Excel 表格 (xlsx)',
    '.xls': 'Excel 表格 (xls)',
    '.pptx': 'PowerPoint 演示 (pptx)',
    '.ppt': 'PowerPoint 演示 (ppt)',
    '.pdf': 'PDF 文档',
}
