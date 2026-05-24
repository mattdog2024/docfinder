# -*- coding: utf-8 -*-
"""
文档内容提取模块
支持: docx, doc, xlsx, xls, pptx, ppt, pdf
"""

import os
import logging
from typing import Optional

logger = logging.getLogger(__name__)


def extract_docx(filepath: str) -> str:
    """提取 .docx 文件文本"""
    try:
        from docx import Document
        doc = Document(filepath)
        paragraphs = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                paragraphs.append(text)
        # 也提取表格中的文字
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text.strip()
                    if text:
                        paragraphs.append(text)
        return '\n'.join(paragraphs)
    except Exception as e:
        logger.warning(f"提取 docx 失败 {filepath}: {e}")
        return ''


def extract_doc(filepath: str) -> str:
    """提取 .doc 文件文本（使用 antiword 或 LibreOffice）"""
    try:
        import subprocess
        # 尝试 antiword
        result = subprocess.run(
            ['antiword', filepath],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    try:
        import subprocess
        import tempfile
        # 尝试用 LibreOffice 转换为 txt
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
                    return f.read()
    except Exception as e:
        logger.warning(f"提取 doc 失败 {filepath}: {e}")

    return ''


def extract_xlsx(filepath: str) -> str:
    """提取 .xlsx 文件文本"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        texts = []
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                for cell in row:
                    if cell is not None:
                        val = str(cell).strip()
                        if val and val != 'None':
                            texts.append(val)
        wb.close()
        return ' '.join(texts)
    except Exception as e:
        logger.warning(f"提取 xlsx 失败 {filepath}: {e}")
        return ''


def extract_xls(filepath: str) -> str:
    """提取 .xls 文件文本"""
    try:
        import xlrd
        wb = xlrd.open_workbook(filepath)
        texts = []
        for sheet in wb.sheets():
            for row_idx in range(sheet.nrows):
                for col_idx in range(sheet.ncols):
                    cell = sheet.cell(row_idx, col_idx)
                    val = str(cell.value).strip()
                    if val and val != '':
                        texts.append(val)
        return ' '.join(texts)
    except Exception as e:
        logger.warning(f"提取 xls 失败 {filepath}: {e}")
        return ''


def extract_pptx(filepath: str) -> str:
    """提取 .pptx 文件文本"""
    try:
        from pptx import Presentation
        prs = Presentation(filepath)
        texts = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'text'):
                    text = shape.text.strip()
                    if text:
                        texts.append(text)
                # 提取表格
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            text = cell.text.strip()
                            if text:
                                texts.append(text)
        return '\n'.join(texts)
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
                    return f.read()
    except Exception as e:
        logger.warning(f"提取 ppt 失败 {filepath}: {e}")
    return ''


def extract_pdf_text(filepath: str) -> str:
    """提取可识别 PDF 的文本"""
    try:
        import pdfplumber
        texts = []
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    texts.append(text.strip())
        return '\n'.join(texts)
    except Exception as e:
        logger.warning(f"提取 pdf 文本失败 {filepath}: {e}")
        return ''


def extract_pdf_ocr(filepath: str, progress_callback=None) -> str:
    """对扫描版 PDF 进行 OCR 识别（需要 pytesseract + pdf2image）"""
    try:
        import pytesseract
        from pdf2image import convert_from_path
        pages = convert_from_path(filepath, dpi=200)
        texts = []
        for i, page_img in enumerate(pages):
            if progress_callback:
                progress_callback(i + 1, len(pages))
            text = pytesseract.image_to_string(page_img, lang='chi_sim+eng')
            if text.strip():
                texts.append(text.strip())
        return '\n'.join(texts)
    except Exception as e:
        logger.warning(f"OCR 识别失败 {filepath}: {e}")
        return ''


def extract_text(filepath: str, enable_pdf: bool = True,
                 enable_ocr: bool = False,
                 progress_callback=None) -> str:
    """
    统一入口：根据文件扩展名自动选择提取方式
    返回提取到的文本内容
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
        # 如果提取到的文字很少（可能是扫描版），且开启了 OCR
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
