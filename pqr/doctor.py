# -*- coding: utf-8 -*-
"""이 PC 에서 무엇이 되고 무엇이 막혔는지 알아봅니다.

옛 워드(.doc) 결재본을 .docx 로 바꾸지 못하면 '보고서 작성' 이 결재본 양식 대신 요약본만
만듭니다. 그 변환은 Word 를 COM 으로 불러 하는데, 회사 PC 마다 막히는 지점이 다릅니다
(PowerShell 정책, COM 차단, cscript 없음 …). 무엇이 막혔는지 알아야 고칠 수 있으므로,
담당자가 두 번 눌러 결과를 알려 줄 수 있게 이 검사를 둡니다.

실제로 임시 .doc 를 만들어 바꿔 보지는 않습니다 — Word 를 띄웠다 닫는 것까지만 해 봅니다.
"""
import os
import shutil
import subprocess
import sys
import tempfile


def _line(ok, label, detail=""):
    mark = "O" if ok else ("-" if ok is None else "X")
    return "  [%s] %-28s %s" % (mark, label, detail)


def _try(cmd, timeout=120):
    """명령을 돌려 (성공 여부, 화면에 나온 글) 을 돌려준다."""
    try:
        run = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=timeout)
    except (OSError, subprocess.TimeoutExpired) as error:
        return False, str(error)
    text = (run.stdout or b"")
    for enc in ("utf-8", "cp949"):
        try:
            text = text.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if isinstance(text, bytes):
        text = text.decode("utf-8", "replace")
    return run.returncode == 0, " ".join(text.split())[:200]


def _com_check(kind):
    """Word/Excel 을 COM 으로 띄웠다 닫아 본다. [(방법, 성공, 설명)]"""
    out = []
    ps = shutil.which("powershell") or shutil.which("pwsh")
    if ps:
        work = tempfile.mkdtemp(prefix="pqr-doctor-")
        path = os.path.join(work, "check.ps1")
        with open(path, "w", encoding="utf-8-sig") as handle:
            handle.write("$ErrorActionPreference='Stop'\n"
                         "$a = New-Object -ComObject %s.Application\n"
                         "$a.Quit()\n" % kind)
        ok, why = _try([ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                        "-File", path])
        shutil.rmtree(work, ignore_errors=True)
        out.append(("PowerShell", ok, why if not ok else "됩니다"))
    else:
        out.append(("PowerShell", False, "PowerShell 을 찾지 못했습니다"))
    cs = shutil.which("cscript")
    if cs:
        work = tempfile.mkdtemp(prefix="pqr-doctor-")
        path = os.path.join(work, "check.vbs")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write('On Error Resume Next\n'
                         'Set a = CreateObject("%s.Application")\n'
                         'If Err.Number <> 0 Then WScript.StdErr.WriteLine Err.Description : WScript.Quit 1\n'
                         'a.Quit\nWScript.Quit 0\n' % kind)
        ok, why = _try([cs, "//nologo", "//B", path])
        shutil.rmtree(work, ignore_errors=True)
        out.append(("VBScript(cscript)", ok, why if not ok else "됩니다"))
    else:
        out.append(("VBScript(cscript)", False, "cscript 를 찾지 못했습니다"))
    return out


def report():
    """검사 결과를 줄 목록으로 돌려줍니다."""
    lines = ["PQR 프로그램 진단", "=" * 58, "",
             "이 PC:", _line(True, "파이썬", sys.version.split()[0]),
             _line(True, "운영체제", sys.platform), ""]

    lines.append("옛 워드(.doc) 결재본을 바꾸는 길:")
    try:
        import win32com.client  # noqa: F401
        lines.append(_line(True, "pywin32", "설치되어 있습니다"))
    except ImportError:
        lines.append(_line(None, "pywin32", "없습니다 — 아래 두 길로 대신합니다"))

    if sys.platform == "win32":
        for how, ok, why in _com_check("Word"):
            lines.append(_line(ok, "Word (%s)" % how, why))
    else:
        lines.append(_line(None, "Word COM", "Windows 가 아니라 건너뜁니다"))

    from .engine import convert
    soffice = convert._soffice()
    lines.append(_line(bool(soffice), "LibreOffice", soffice or "없습니다 (없어도 됩니다)"))
    lines.append("")

    lines.append("첨부 Cpk 엑셀(그래프까지) 채우는 길:")
    if sys.platform == "win32":
        for how, ok, why in _com_check("Excel"):
            lines.append(_line(ok, "Excel (%s)" % how, why))
    else:
        lines.append(_line(None, "Excel COM", "Windows 가 아니라 건너뜁니다"))
    lines.append("")

    word_ok = any(ok for _, ok, _ in _com_check("Word")) if sys.platform == "win32" else False
    lines.append("판정:")
    if word_ok or soffice:
        lines.append("  옛 워드(.doc) 결재본을 바꿀 수 있습니다 — '보고서 작성' 이 결재본 양식으로 만듭니다.")
    else:
        lines.append("  옛 워드(.doc) 를 바꿀 길이 없습니다. 둘 중 하나를 하세요:")
        lines.append("   1) 전년도 PQR(.doc)을 Word 로 열어 '다른 이름으로 저장 → Word 문서(*.docx)' 로")
        lines.append("      저장한 뒤 제품 폴더에 두세요. 그러면 바꿀 일이 없어집니다.")
        lines.append("   2) 위 [X] 줄에 적힌 까닭을 그대로 알려 주시면 그 길을 뚫겠습니다.")
    lines.append("")
    lines.append("이 화면을 그대로 알려 주시면 됩니다.")
    return lines


def main(argv=None):
    for line in report():
        try:
            print(line)
        except UnicodeEncodeError:
            print(line.encode("utf-8", "replace").decode("utf-8", "replace"))
    return 0
