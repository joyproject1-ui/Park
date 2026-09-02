# -*- coding: utf-8 -*-
"""옛 워드(.doc) 결재본을 .docx 로 바꾼다.

담당자 PC(Windows) 에는 Word 가 있으므로 Word 자동화(COM)로 바꾼다 — 글꼴·조판이
가장 정확하다. Word 가 없으면 LibreOffice(soffice) 를 찾아 쓴다. 둘 다 없으면 담당자가
Word 에서 '다른 이름으로 저장 → .docx' 로 만들어 두도록 안내한다.
"""
import os
import shutil
import subprocess
import sys


class ConvertError(Exception):
    pass


def _with_word(src, dst):
    try:
        import win32com.client  # pywin32
    except ImportError:
        return False
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(src), ReadOnly=True)
        doc.SaveAs2(os.path.abspath(dst), FileFormat=16)   # wdFormatXMLDocument
        doc.Close(False)
    finally:
        word.Quit()
    return os.path.isfile(dst)


def _soffice():
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    for cand in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.isfile(cand):
            return cand
    return None


def _with_soffice(src, dst):
    exe = _soffice()
    if not exe:
        return False
    outdir = os.path.dirname(os.path.abspath(dst))
    subprocess.run([exe, "--headless", "--convert-to", "docx", "--outdir", outdir, os.path.abspath(src)],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    made = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".docx")
    if os.path.isfile(made) and os.path.abspath(made) != os.path.abspath(dst):
        shutil.move(made, dst)
    return os.path.isfile(dst)


def to_docx(src, dst):
    """src(.doc/.docx) → dst(.docx). 이미 .docx 면 복사만 한다."""
    if src.lower().endswith(".docx"):
        shutil.copyfile(src, dst)
        return "copy"
    if sys.platform == "win32" and _with_word(src, dst):
        return "word"
    if _with_soffice(src, dst):
        return "soffice"
    raise ConvertError(
        "결재본이 옛 워드 형식(.doc)인데 바꿀 도구가 없습니다. Word 에서 결재본을 열어 "
        "'다른 이름으로 저장 → Word 문서(*.docx)' 로 저장해 같은 폴더에 두세요.")


def _xls_with_excel(src, dst):
    try:
        import win32com.client
    except ImportError:
        return False
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(src), ReadOnly=True)
        wb.SaveAs(os.path.abspath(dst), FileFormat=51)     # xlOpenXMLWorkbook
        wb.Close(False)
    finally:
        excel.Quit()
    return os.path.isfile(dst)


def _xls_with_soffice(src, dst):
    exe = _soffice()
    if not exe:
        return False
    outdir = os.path.dirname(os.path.abspath(dst))
    subprocess.run([exe, "--headless", "--convert-to", "xlsx", "--outdir", outdir, os.path.abspath(src)],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    made = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".xlsx")
    if os.path.isfile(made) and os.path.abspath(made) != os.path.abspath(dst):
        shutil.move(made, dst)
    return os.path.isfile(dst)


def to_xlsx(src, dst):
    """옛 엑셀(.xls) → .xlsx (수식 보존). 이미 .xlsx 면 복사."""
    if src.lower().endswith((".xlsx", ".xlsm")):
        shutil.copyfile(src, dst)
        return "copy"
    if sys.platform == "win32" and _xls_with_excel(src, dst):
        return "excel"
    if _xls_with_soffice(src, dst):
        return "soffice"
    raise ConvertError("Cpk 계산 파일(.xls)을 바꿀 도구가 없습니다 (Excel 또는 LibreOffice).")


def _pdf_with_word(src, dst):
    try:
        import win32com.client
    except ImportError:
        return False
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(src), ReadOnly=True)
        doc.SaveAs2(os.path.abspath(dst), FileFormat=17)      # wdFormatPDF
        doc.Close(False)
    finally:
        word.Quit()
    return os.path.isfile(dst)


def _pdf_with_soffice(src, dst):
    exe = _soffice()
    if not exe:
        return False
    outdir = os.path.dirname(os.path.abspath(dst))
    subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", outdir, os.path.abspath(src)],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
    made = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    if os.path.isfile(made) and os.path.abspath(made) != os.path.abspath(dst):
        shutil.move(made, dst)
    return os.path.isfile(dst)


def to_pdf(src, dst):
    """보고서(.docx/.doc)를 화면에서 볼 수 있게 PDF 로 바꾼다.

    담당자 PC 에는 Word 가 있으므로 Word 로 만든다(쪽 나눔·글꼴이 실제 인쇄본과 같다).
    없으면 LibreOffice 를 쓴다. 이미 PDF 면 복사만 한다.
    """
    if src.lower().endswith(".pdf"):
        shutil.copyfile(src, dst)
        return "copy"
    if sys.platform == "win32" and _pdf_with_word(src, dst):
        return "word"
    if _pdf_with_soffice(src, dst):
        return "soffice"
    raise ConvertError("보고서를 화면에 띄우려면 Word 또는 LibreOffice 가 필요합니다.")
