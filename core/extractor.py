# -*- coding: utf-8 -*-
"""
文档内容提取模块 v1.3 - 性能优化版
支持: docx, doc, xlsx, xls, pptx, ppt, pdf

优化点：
1. 内容截断：每个文件最多提取 MAX_CONTENT_CHARS 字符，避免超大文件拖慢速度
2. xlsx 使用 read_only 模式，内存占用更低
3. 所有提取函数加入超时保护（通过调用方控制）
"""

import os
import logging
from typing import Optional

logger = logging.getLogger(__name__)

# 每个文件最多提取的字符数（约 10 万字，足够搜索用）
# 超过此长度的内容直接截断，避免超大文件拖慢索引
MAX_CONTENT_CHARS = 100_000


def extract_docx(filepath: str) -> str:
    """提取 .docx 文件文本（优化版）"""
    try:
        from docx import Document
        doc = Document(filepath)
        parts = []
        total = 0

        # 提取段落
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                parts.append(text)
                total += len(text)
                if total >= MAX_CONTENT_CHARS:
                    break

        # 提取表格（如果还没超限）
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
    """提取 .doc 文件文本（使用 antiword 或 LibreOffice）"""
    # 先尝试 antiword（速度快）
    try:
        import subprocess
        result = subprocess.run(
            ['antiword', filepath],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout[:MAX_CONTENT_CHARS]
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    # 再尝试 LibreOffice（速度慢，但兼容性好）
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
    try:
        import openpyxl
        # read_only=True 大幅减少内存占用和解析时间
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        parts = []
        total = 0
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                row_texts = []
                for cell in row:
                    if cell is not None:
                        val = str(cell).strip()
                        if val and val != 'None':
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
    try:
        import xlrd
        wb = xlrd.open_workbook(filepath)
        parts = []
        total = 0
        for sheet in wb.sheets():
            for row_idx in range(sheet.nrows):
                row_texts = []
                for col_idx in range(sheet.ncols):
                    val = str(sheet.cell(row_idx, col_idx).value).strip()
                    if val and val != '':
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
    """提取 .pptx 文件文本"""
    try:
        from pptx import Presentation
        prs = Presentation(filepath)
        parts = []
        total = 0
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'text') and shape.text.strip():
                    text = shape.text.strip()
                    parts.append(text)
                    total += len(text)
                    if total >= MAX_CONTENT_CHARS:
                        break
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            text = cell.text.strip()
                            if text:
                                parts.append(text)
                                total += len(text)
            if total >= MAX_CONTENT_CHARS:
                break
        return '\n'.join(parts)[:MAX_CONTENT_CHARS]
    except Exception as e:
        logger.warning(f"提取 pptx 失败 {filepath}: {e}")
        return ''


def extract_ppt(filepath: str) -> str:
    """提取 .ppt 文件文本（使用 LibreOffice 转换）"""
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
    """提取可识别 PDF 的文本"""
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
        logger.warning(f"提取 pdf 文本失败 {filepath}: {e}")
        return ''


def extract_pdf_ocr(filepath: str, progress_callback=None) -> str:
    """对扫描版 PDF 进行 OCR 识别（需要 pytesseract + pdf2image）"""
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
