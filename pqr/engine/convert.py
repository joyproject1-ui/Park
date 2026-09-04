# -*- coding: utf-8 -*-
"""옛 워드(.doc) 결재본을 .docx 로 바꾼다.

담당자 PC(Windows) 에는 Word 가 있으므로 Word 자동화(COM)로 바꾼다 — 글꼴·조판이
가장 정확하다. Word 가 없으면 LibreOffice(soffice) 를 찾아 쓴다. 둘 다 없으면 담당자가
Word 에서 '다른 이름으로 저장 → .docx' 로 만들어 두도록 안내한다.
"""
import io
import os
import shutil
import subprocess
import tempfile
import sys


class ConvertError(Exception):
    pass


last_error = []          # 변환에 실패했을 때 마지막 오류 — 안내 문구에 붙인다


def _run(cmd, why):
    """명령을 돌리고 성공 여부를 돌려준다. 실패하면 까닭을 last_error 에 남긴다."""
    try:
        run = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=300)
    except (OSError, subprocess.TimeoutExpired) as error:
        last_error.append("%s: %s" % (why, error))
        return False
    if run.returncode != 0:
        text = (run.stdout or b"").decode("utf-8", "replace")
        if not text.strip():
            text = (run.stdout or b"").decode("cp949", "replace")
        last_error.append("%s: %s" % (why, " ".join(text.split())[:300] or "코드 %d" % run.returncode))
    return run.returncode == 0


def _script_run(text, suffix, cmd_for):
    """스크립트를 임시 파일에 써서 돌린다.

    명령줄에 통째로 넘기면(-Command) 따옴표·한글 경로에서 잘리는 PC 가 있어 파일로 넘긴다.
    """
    work = tempfile.mkdtemp(prefix="pqr-conv-")
    path = os.path.join(work, "convert" + suffix)
    try:
        with io.open(path, "w", encoding="utf-8-sig" if suffix == ".ps1" else "utf-8") as handle:
            handle.write(text)
        return _run(cmd_for(path), suffix.lstrip("."))
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _powershell(script):
    """PowerShell 로 COM 자동화 — pywin32 가 없는 PC 에서도 Word·Excel 을 쓸 수 있다."""
    exe = shutil.which("powershell") or shutil.which("pwsh")
    if not exe:
        last_error.append("PowerShell 을 찾지 못했습니다")
        return False
    return _script_run(script, ".ps1", lambda path: [
        exe, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", path])


def _vbscript(script):
    """VBScript(cscript) 로 COM 자동화 — PowerShell 이 막힌 회사 PC 에서도 대개 돈다."""
    exe = shutil.which("cscript")
    if not exe:
        last_error.append("cscript 를 찾지 못했습니다")
        return False
    return _script_run(script, ".vbs", lambda path: [exe, "//nologo", "//B", path])


def _vbs_path(path):
    return '"' + os.path.abspath(path).replace('"', '""') + '"'


def _word_via_vbscript(src, dst, fmt=16):
    """Word 를 VBScript 로 불러 문서를 다른 형식으로 저장한다."""
    return _vbscript(
        'On Error Resume Next\n'
        'Set w = CreateObject("Word.Application")\n'
        'If Err.Number <> 0 Then WScript.StdErr.WriteLine "Word 를 열지 못함: " & Err.Description : WScript.Quit 1\n'
        'w.Visible = False\n'
        'w.DisplayAlerts = 0\n'
        'Set d = w.Documents.Open(%s, False, True)\n'
        'If Err.Number <> 0 Then WScript.StdErr.WriteLine "문서를 열지 못함: " & Err.Description : w.Quit : WScript.Quit 1\n'
        'd.SaveAs %s, %d\n'
        'If Err.Number <> 0 Then WScript.StdErr.WriteLine "저장하지 못함: " & Err.Description : d.Close False : w.Quit : WScript.Quit 1\n'
        'd.Close False\n'
        'w.Quit\n'
        'WScript.Quit 0\n' % (_vbs_path(src), _vbs_path(dst), fmt))


def _ps_path(path):
    return "'" + os.path.abspath(path).replace("'", "''") + "'"


def _word_via_powershell(src, dst, fmt=16):
    """Word 를 PowerShell COM 으로 불러 문서를 다른 형식으로 저장한다."""
    script = (
        "$ErrorActionPreference='Stop'\n"
        "$w = New-Object -ComObject Word.Application\n"
        "$w.Visible = $false\n"
        "$w.DisplayAlerts = 0\n"
        "try {\n"
        "  $d = $w.Documents.Open(%s, $false, $true)\n"
        "  $d.SaveAs(%s, %d)\n"
        "  $d.Close($false)\n"
        "} finally { $w.Quit() }\n" % (_ps_path(src), _ps_path(dst), fmt))
    return _powershell(script) and os.path.isfile(dst)


def _word_convert(src, dst, fmt=16):
    """Word 로 변환 — PowerShell 을 먼저, 막혀 있으면 VBScript 로. 하나라도 되면 True."""
    for how in (_word_via_powershell, _word_via_vbscript):
        try:
            if how(src, dst, fmt) and os.path.isfile(dst):
                return True
        except Exception as error:                 # 한 방법이 터져도 다음 방법을 본다
            last_error.append("%s: %s" % (how.__name__, error))
    return False


def _with_word(src, dst):
    try:
        import win32com.client  # pywin32
    except ImportError:
        return _word_convert(src, dst, 16)
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
    detail = (" (마지막 오류: %s)" % last_error[-1]) if last_error else ""
    raise ConvertError(
        "전년도 결재본이 옛 워드 형식(.doc)인데 .docx 로 바꾸지 못했습니다%s. "
        "Word 에서 그 파일을 열어 '다른 이름으로 저장 → Word 문서(*.docx)' 로 저장한 뒤 "
        "제품 폴더에 두고 '보고서 작성' 을 다시 누르세요." % detail)


def _excel_convert(src, dst):
    """엑셀도 같은 두 경로로 — .xls → .xlsx (51 = xlOpenXMLWorkbook)."""
    ps = ("$ErrorActionPreference='Stop'\n"
          "$x = New-Object -ComObject Excel.Application\n"
          "$x.Visible = $false\n$x.DisplayAlerts = $false\n"
          "try { $b = $x.Workbooks.Open(%s, 0, $true); $b.SaveAs(%s, 51); $b.Close($false) } "
          "finally { $x.Quit() }\n" % (_ps_path(src), _ps_path(dst)))
    if _powershell(ps) and os.path.isfile(dst):
        return True
    vbs = ('On Error Resume Next\n'
           'Set x = CreateObject("Excel.Application")\n'
           'If Err.Number <> 0 Then WScript.Quit 1\n'
           'x.Visible = False\nx.DisplayAlerts = False\n'
           'Set b = x.Workbooks.Open(%s, 0, True)\n'
           'If Err.Number <> 0 Then x.Quit : WScript.Quit 1\n'
           'b.SaveAs %s, 51\n'
           'If Err.Number <> 0 Then b.Close False : x.Quit : WScript.Quit 1\n'
           'b.Close False\nx.Quit\nWScript.Quit 0\n' % (_vbs_path(src), _vbs_path(dst)))
    return _vbscript(vbs) and os.path.isfile(dst)


def _xls_with_excel(src, dst):
    try:
        import win32com.client
    except ImportError:
        return _excel_convert(src, dst)
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
