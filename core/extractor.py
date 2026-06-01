# -*- coding: utf-8 -*-
"""
文档内容提取模块 v1.4
支持: docx, doc, xlsx, xls, pptx, ppt, pdf

修复点（v1.4）：
1. doc/ppt 老格式：新增 docx2txt 兜底，无需 LibreOffice 也能提取
2. xls 数字格式：修复浮点数显示（1.0 -> 1，避免噪声）
3. pptx 表格重复提取：shape.text 已包含表格内容，去掉重复的 table 遍历
4. 超大文件保护：文件超过 MAX_FILE_SIZE_MB 直接跳过，避免内存溢出
5. 编码容错：所有文本提取都加 errors='ignore'
"""

import os
import logging
from typing import Optional

logger = logging.getLogger(__name__)

# 每个文件最多提取的字符数（约 10 万字，足够搜索用）
MAX_CONTENT_CHARS = 100_000

# 超过此大小的文件跳过（单位：MB），防止内存溢出
MAX_FILE_SIZE_MB = 200


def _check_file_size(filepath: str) -> bool:
    """检查文件大小是否在允许范围内"""
    try:
        size_mb = os.path.getsize(filepath) / (1024 * 1024)
        if size_mb > MAX_FILE_SIZE_MB:
            logger.warning(f"文件过大（{size_mb:.1f}MB），跳过：{filepath}")
            return False
        return True
    except Exception:
        return True


def extract_docx(filepath: str) -> str:
    """提取 .docx 文件文本"""
    if not _check_file_size(filepath):
        return ''
    try:
        from docx import Document
        doc = Document(filepath)
        parts = []
        total = 0

        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                parts.append(text)
                total += len(text)
                if total >= MAX_CONTENT_CHARS:
                    break

        if total < MAX_CONTENT_CHARS:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text = cell.text.strip()
                        if text:
                            parts.append(text)
                            total += len(text)
                            if total >= MAX_CONTENT_CHARS:
                                break
                    if total >= MAX_CONTENT_CHARS:
                        break
                if total >= MAX_CONTENT_CHARS:
                    break

        return '\n'.join(parts)[:MAX_CONTENT_CHARS]
    except Exception as e:
        logger.warning(f"提取 docx 失败 {filepath}: {e}")
        return ''


def extract_doc(filepath: str) -> str:
    """提取 .doc 文件文本
    优先级：antiword > docx2txt > LibreOffice
    """
    if not _check_file_size(filepath):
        return ''

    # 方式1：antiword（最快，需要系统安装）
    try:
        import subprocess
        result = subprocess.run(
            ['antiword', '-t', filepath],
            capture_output=True, timeout=30
        )
        if result.returncode == 0 and result.stdout.strip():
            text = result.stdout.decode('utf-8', errors='ignore')
            if len(text.strip()) > 20:
                return text[:MAX_CONTENT_CHARS]
    except (FileNotFoundError, Exception):
        pass

    # 方式2：docx2txt（纯 Python，无需额外安装）
    try:
        import docx2txt
        text = docx2txt.process(filepath)
        if text and len(text.strip()) > 20:
            return text[:MAX_CONTENT_CHARS]
    except Exception:
        pass

    # 方式3：尝试用 python-docx 直接读（部分 .doc 文件可以）
    try:
        from docx import Document
        doc = Document(filepath)
        parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        if parts:
            return '\n'.join(parts)[:MAX_CONTENT_CHARS]
    except Exception:
        pass

    # 方式4：LibreOffice（最慢，兼容性最好）
    try:
        import subprocess
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            result = subprocess.run(
                ['libreoffice', '--headless', '--convert-to', 'txt:Text',
                 '--outdir', tmpdir, filepath],
                capture_output=True, timeout=60
            )
            base = os.path.splitext(os.path.basename(filepath))[0]
            txt_path = os.path.join(tmpdir, base + '.txt')
            if os.path.exists(txt_path):
                with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read(MAX_CONTENT_CHARS)
    except Exception as e:
        logger.warning(f"提取 doc 失败 {filepath}: {e}")

    return ''


def extract_xlsx(filepath: str) -> str:
    """提取 .xlsx 文件文本（read_only 模式，速度快）"""
    if not _check_file_size(filepath):
        return ''
    try:
        import openpyxl
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        parts = []
        total = 0
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                row_texts = []
                for cell in row:
                    if cell is not None:
                        # 修复：整数不显示小数点（1.0 -> 1）
                        if isinstance(cell, float) and cell == int(cell):
                            val = str(int(cell))
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
        logger.warning(f"提取 xlsx 失败 {filepath}: {e}")
        return ''


def extract_xls(filepath: str) -> str:
    """提取 .xls 文件文本"""
    if not _check_file_size(filepath):
        return ''
    try:
        import xlrd
        wb = xlrd.open_workbook(filepath)
        parts = []
        total = 0
        for sheet in wb.sheets():
            for row_idx in range(sheet.nrows):
                row_texts = []
                for col_idx in range(sheet.ncols):
                    cell = sheet.cell(row_idx, col_idx)
                    # 修复：xlrd 数字类型（type=2）去掉多余小数点
                    if cell.ctype == 2:  # XL_CELL_NUMBER
                        val_f = cell.value
                        val = str(int(val_f)) if val_f == int(val_f) else str(val_f)
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
        logger.warning(f"提取 xls 失败 {filepath}: {e}")
        return ''


def extract_pptx(filepath: str) -> str:
    """提取 .pptx 文件文本（修复：去掉重复的表格遍历）"""
    if not _check_file_size(filepath):
        return ''
    try:
        from pptx import Presentation
        prs = Presentation(filepath)
        parts = []
        total = 0
        for slide in prs.slides:
            for shape in slide.shapes:
                # shape.text 已经包含了表格内容，不需要单独遍历 table
                if hasattr(shape, 'text') and shape.text.strip():
                    text = shape.text.strip()
                    parts.append(text)
                    total += len(text)
                    if total >= MAX_CONTENT_CHARS:
                        break
            if total >= MAX_CONTENT_CHARS:
                break
        return '\n'.join(parts)[:MAX_CONTENT_CHARS]
    except Exception as e:
        logger.warning(f"提取 pptx 失败 {filepath}: {e}")
        return ''


def extract_ppt(filepath: str) -> str:
    """提取 .ppt 文件文本
    优先级：docx2txt > LibreOffice
    """
    if not _check_file_size(filepath):
        return ''

    # 方式1：docx2txt（部分 ppt 可以直接读）
    try:
        import docx2txt
        text = docx2txt.process(filepath)
        if text and len(text.strip()) > 20:
            return text[:MAX_CONTENT_CHARS]
    except Exception:
        pass

    # 方式2：LibreOffice
    try:
        import subprocess
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            result = subprocess.run(
                ['libreoffice', '--headless', '--convert-to', 'txt:Text',
                 '--outdir', tmpdir, filepath],
                capture_output=True, timeout=60
            )
            base = os.path.splitext(os.path.basename(filepath))[0]
            txt_path = os.path.join(tmpdir, base + '.txt')
            if os.path.exists(txt_path):
                with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read(MAX_CONTENT_CHARS)
    except Exception as e:
        logger.warning(f"提取 ppt 失败 {filepath}: {e}")
    return ''


def extract_pdf_text(filepath: str) -> str:
    """提取可识别 PDF 的文本（优先 pymupdf，速度更快；兜底 pdfplumber）"""
    if not _check_file_size(filepath):
        return ''

    # 方式1：pymupdf（fitz），速度比 pdfplumber 快 3-5 倍
    try:
        import fitz  # pymupdf
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
        logger.warning(f"pymupdf 提取 pdf 失败 {filepath}: {e}")

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
        logger.warning(f"pdfplumber 提取 pdf 失败 {filepath}: {e}")
        return ''


def extract_pdf_ocr(filepath: str, progress_callback=None) -> str:
    """对扫描版 PDF 进行 OCR 识别（需要 pytesseract + pdf2image）"""
    if not _check_file_size(filepath):
        return ''
    try:
        import pytesseract
        from pdf2image import convert_from_path
        pages = convert_from_path(filepath, dpi=200)
        parts = []
        total = 0
        for i, page_img in enumerate(pages):
            if progress_callback:
                progress_callback(i + 1, len(pages))
            text = pytesseract.image_to_string(page_img, lang='chi_sim+eng')
            if text.strip():
                parts.append(text.strip())
                total += len(text)
                if total >= MAX_CONTENT_CHARS:
                    break
        return '\n'.join(parts)[:MAX_CONTENT_CHARS]
    except Exception as e:
        logger.warning(f"OCR 识别失败 {filepath}: {e}")
        return ''


def extract_text(filepath: str, enable_pdf: bool = True,
                 enable_ocr: bool = False,
                 progress_callback=None) -> str:
    """
    统一入口：根据文件扩展名自动选择提取方式
    返回提取到的文本内容（最多 MAX_CONTENT_CHARS 字符）
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
        return extractors[ext](filepath)
    elif ext == '.pdf':
        if not enable_pdf:
            return ''
        text = extract_pdf_text(filepath)
        if len(text.strip()) < 50 and enable_ocr:
            logger.info(f"PDF 文字内容较少，尝试 OCR: {filepath}")
            ocr_text = extract_pdf_ocr(filepath, progress_callback)
            if len(ocr_text) > len(text):
                return ocr_text
        return text
    else:
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
