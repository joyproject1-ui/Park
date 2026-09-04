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
    """Word 로 .doc → .docx. pywin32 가 있으면 그 길을 먼저, 안 되면 스크립트 두 길로 넘어간다.

    pywin32 가 터져도 여기서 멈추면 안 된다 — 담당자 PC 에서 pywin32·PowerShell·VBScript 가
    모두 되는데도 보고서가 안 나온 일이 있었다(2026-09, 디겐타안연고). pywin32 길만 쓰고 그
    오류를 그대로 밖으로 던지는 바람에 나머지 두 길을 아예 시도하지 않았다.

    ConfirmConversions=False 를 꼭 넘긴다 — 옛 워드(.doc)를 열 때 뜨는 '파일 변환' 확인
    대화상자를 끈다. 화면 없이 도는 자동화에서 대화상자가 뜨면 그대로 멈춘다. 스크립트 두 길은
    Open(경로, False, True) 로 이미 끄고 있었는데 이 길만 빠져 있었다.
    """
    try:
        import win32com.client  # pywin32
    except ImportError:
        return _word_convert(src, dst, 16)
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0                             # wdAlertsNone — 대화상자를 띄우지 않는다
        try:
            # (파일, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
            # 자리로 넘긴다 — 이름으로 넘기는 것보다 확실하고, 되는 것이 확인된 VBScript 길과 같다.
            doc = word.Documents.Open(os.path.abspath(src), False, True, False)
            try:
                doc.SaveAs2(os.path.abspath(dst), 16)      # 16 = wdFormatXMLDocument (.docx)
            except Exception:                              # 옛 Word 에는 SaveAs2 가 없다
                doc.SaveAs(os.path.abspath(dst), 16)
            doc.Close(False)
        finally:
            try:
                word.Quit()
            except Exception:                              # Quit 이 터져도 원래 오류를 가리지 않는다
                pass
        if os.path.isfile(dst):
            return True
        last_error.append("pywin32: Word 가 저장한 파일이 없습니다")
    except Exception as error:
        last_error.append("pywin32: %s" % error)
    return _word_convert(src, dst, 16)                     # PowerShell → VBScript 로 다시 해 본다


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


def _guard(how, src, dst, name):
    """한 길이 터져도 다음 길을 본다 — 까닭은 last_error 에 남겨 안내 문구에 붙인다.

    변환 실패는 언제나 ConvertError 로만 밖에 나가야 한다. COM 오류가 그대로 새어 나가면
    부르는 쪽이 '변환 실패' 로 알아보지 못해 EDMS 빈 서식으로 물러나는 길까지 막힌다.
    """
    try:
        return bool(how(src, dst))
    except Exception as error:
        last_error.append("%s: %s" % (name, error))
        return False


def to_docx(src, dst):
    """src(.doc/.docx) → dst(.docx). 이미 .docx 면 복사만 한다."""
    if src.lower().endswith(".docx"):
        shutil.copyfile(src, dst)
        return "copy"
    if sys.platform == "win32" and _guard(_with_word, src, dst, "Word"):
        return "word"
    if _guard(_with_soffice, src, dst, "LibreOffice"):
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
    """엑셀도 같다 — pywin32 가 터지면 스크립트 두 길로 넘어간다."""
    try:
        import win32com.client
    except ImportError:
        return _excel_convert(src, dst)
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(os.path.abspath(src), UpdateLinks=0, ReadOnly=True)
            wb.SaveAs(os.path.abspath(dst), FileFormat=51)     # xlOpenXMLWorkbook
            wb.Close(False)
        finally:
            try:
                excel.Quit()
            except Exception:
                pass
        if os.path.isfile(dst):
            return True
        last_error.append("pywin32(Excel): 저장한 파일이 없습니다")
    except Exception as error:
        last_error.append("pywin32(Excel): %s" % error)
    return _excel_convert(src, dst)


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
    if sys.platform == "win32" and _guard(_xls_with_excel, src, dst, "Excel"):
        return "excel"
    if _guard(_xls_with_soffice, src, dst, "LibreOffice"):
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


# ---------- 목차 쪽수 등 필드 다시 계산 ----------
def _fields_via_vbscript(path):
    return _vbscript(
        'On Error Resume Next\n'
        'Set w = CreateObject("Word.Application")\n'
        'If Err.Number <> 0 Then WScript.Quit 1\n'
        'w.Visible = False\n'
        'w.DisplayAlerts = 0\n'
        'Set d = w.Documents.Open(%s, False, False, False)\n'
        'If Err.Number <> 0 Then w.Quit : WScript.Quit 1\n'
        'For pass = 1 To 2\n'
        '  d.Repaginate\n'
        '  For Each s In d.StoryRanges\n'
        '    s.Fields.Update\n'
        '    Do While Not (s.NextStoryRange Is Nothing)\n'
        '      Set s = s.NextStoryRange\n'
        '      s.Fields.Update\n'
        '    Loop\n'
        '  Next\n'
        'Next\n'
        'd.Save\n'
        'If Err.Number <> 0 Then d.Close 0 : w.Quit : WScript.Quit 1\n'
        'd.Close 0\n'
        'w.Quit\n'
        'WScript.Quit 0\n' % _vbs_path(path))


def _fields_via_powershell(path):
    return _powershell(
        "$ErrorActionPreference='Stop'\n"
        "$w = New-Object -ComObject Word.Application\n"
        "$w.Visible = $false\n"
        "$w.DisplayAlerts = 0\n"
        "try {\n"
        "  $d = $w.Documents.Open(%s, $false, $false, $false)\n"
        "  for ($i = 0; $i -lt 2; $i++) {\n"
        "    $d.Repaginate()\n"
        "    foreach ($s in $d.StoryRanges) {\n"
        "      $r = $s\n"
        "      while ($r -ne $null) { $r.Fields.Update() | Out-Null; $r = $r.NextStoryRange }\n"
        "    }\n"
        "  }\n"
        "  $d.Save()\n"
        "  $d.Close(0)\n"
        "} finally { $w.Quit() }\n" % _ps_path(path))


def _fields_via_pywin32(path):
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(os.path.abspath(path), False, False, False)
        try:
            for _ in range(2):
                doc.Repaginate()
                for story in doc.StoryRanges:
                    here = story
                    while here is not None:
                        here.Fields.Update()
                        here = here.NextStoryRange
            doc.Save()
        finally:
            doc.Close(0)
    finally:
        word.Quit()
    return True


def refresh_fields(path):
    """Word 로 열어 쪽 나눔을 다시 하고 모든 필드(목차 쪽수·머리글 Page)를 계산해 저장한다.

    담당자 PC 에서 목차 쪽수가 전부 '1' 로 나왔다(2026-09, 디겐타안연고). 필드에
    dirty 표시를 달아 두면 Word 가 열면서 갱신하지만, 쪽 나눔이 끝나기 전에 갱신해
    PAGEREF 가 1 로 잡히는 일이 있다. 만든 자리에서 Word 로 한 번 계산해 두면 파일을
    열자마자 맞는 쪽수가 보인다. Word 가 없으면(리눅스 등) 아무 일도 하지 않고 False.
    """
    del last_error[:]
    for how in (_fields_via_pywin32, _fields_via_powershell, _fields_via_vbscript):
        try:
            if how(path):
                return True
        except Exception as error:
            last_error.append("%s: %s" % (how.__name__, error))
    return False
