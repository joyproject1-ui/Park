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


def _make_doc(path):
    """Word 로 시험용 옛 워드(.doc) 하나를 만든다 (0 = wdFormatDocument). 되면 True."""
    from .engine import convert
    ps = ("$ErrorActionPreference='Stop'\n"
          "$w = New-Object -ComObject Word.Application\n"
          "$w.Visible = $false\n$w.DisplayAlerts = 0\n"
          "try { $d = $w.Documents.Add(); $d.Content.Text = 'PQR 진단'; "
          "$d.SaveAs(%s, 0); $d.Close($false) } finally { $w.Quit() }\n" % convert._ps_path(path))
    if convert._powershell(ps) and os.path.isfile(path):
        return True
    vbs = ('On Error Resume Next\n'
           'Set w = CreateObject("Word.Application")\n'
           'If Err.Number <> 0 Then WScript.Quit 1\n'
           'w.Visible = False\nw.DisplayAlerts = 0\n'
           'Set d = w.Documents.Add()\n'
           'd.Content.Text = "PQR 진단"\n'
           'd.SaveAs %s, 0\n'
           'If Err.Number <> 0 Then w.Quit : WScript.Quit 1\n'
           'd.Close False\nw.Quit\nWScript.Quit 0\n' % convert._vbs_path(path))
    return convert._vbscript(vbs) and os.path.isfile(path)


def convert_check():
    """진짜로 .doc 를 만들어 .docx 로 바꿔 본다. (되는지, 설명)

    Word 를 띄웠다 닫는 것만으로는 모자랐다 — 담당자 PC 에서 그 검사는 모두 '됩니다' 였는데
    보고서는 안 나왔다(2026-09, 디겐타안연고). 실제로 열고 저장하는 데서 막히면
    (옛 워드의 '파일 변환' 확인 대화상자 같은 것) 여기서만 드러난다.
    """
    from .engine import convert
    if sys.platform != "win32":
        return None, "Windows 가 아니라 건너뜁니다"
    work = tempfile.mkdtemp(prefix="pqr-doctor-")
    src = os.path.join(work, "pqr-check.doc")
    try:
        del convert.last_error[:]
        if not _make_doc(src):
            why = convert.last_error[-1] if convert.last_error else "까닭 모름"
            return False, "시험용 .doc 를 만들지 못했습니다 — %s" % why
        try:
            how = convert.to_docx(src, os.path.join(work, "pqr-check.docx"))
            return True, "됩니다 (%s)" % how
        except convert.ConvertError as error:
            return False, " ".join(str(error).split())[:220]
    finally:
        shutil.rmtree(work, ignore_errors=True)


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
    # 띄웠다 닫는 것만으로는 모자란다 — 실제로 .doc 를 만들어 .docx 로 바꿔 본다.
    really, why = convert_check()
    lines.append(_line(really, "실제로 바꿔 보기", why))
    lines.append("")

    lines.append("첨부 Cpk 엑셀(그래프까지) 채우는 길:")
    if sys.platform == "win32":
        for how, ok, why in _com_check("Excel"):
            lines.append(_line(ok, "Excel (%s)" % how, why))
    else:
        lines.append(_line(None, "Excel COM", "Windows 가 아니라 건너뜁니다"))
    lines.append("")

    lines.append("손글씨 안정성시험일지 판독(13항):")
    key_on = bool(os.environ.get("ANTHROPIC_API_KEY"))
    try:
        from .engine import handwriting
        offline = handwriting.available()
    except Exception:
        offline = False
    lines.append(_line(key_on, "ANTHROPIC_API_KEY (Claude 비전)",
                       "켜짐 — 시험일지 PDF 를 정확히 읽습니다" if key_on else "없음 — 오프라인 판독기로 대신합니다"))
    try:
        from .engine import claude_cli
        cli = claude_cli._exe()
    except Exception:
        cli = ""
    lines.append(_line(bool(cli), "이 PC 의 Claude Code (claude)",
                       cli if cli else "찾지 못했습니다 — 아래 '찾아본 자리' 참고"))
    if not cli:
        try:
            from .engine import claude_cli as _cc
            lines.append("      찾아본 자리: PATH 의 claude/claude.cmd/claude.exe 와")
            for guess in _cc.places()[:12]:
                lines.append("        %s" % guess)
            lines.append("      · 브라우저나 Claude 앱에서 쓰는 화면은 이 PC 의 명령(claude)과 다릅니다.")
            lines.append("      · 이미 깔려 있다면 명령 프롬프트에서  where claude  로 경로를 확인해,")
            lines.append("        그 경로 한 줄을 프로그램 폴더의 '%s' 에 적어 두면 그것을 씁니다." % _cc.HINT_FILE)
            lines.append("      · 새로 깔았다면 창을 닫았다 다시 여세요(PATH 는 새 창부터 반영됩니다).")
            lines.append("      · 없다면: Node.js 18 이상을 깔고  npm install -g @anthropic-ai/claude-code")
            lines.append("        → claude 를 한 번 실행해 로그인하면 '안정성 판독' 이 이 PC 안에서 돕니다.")
        except Exception:
            pass
    lines.append(_line(offline, "오프라인 판독기 (RapidOCR)",
                       "설치됨 — 깨끗이 읽힌 값만 쓰고 애매한 칸은 노랑/주황으로 표시합니다" if offline
                       else "없음 — %s / PQR-업데이트.bat 을 다시 실행하고 'install-log.txt' 를 보내 주세요"
                            % (handwriting.why() if offline is False else "")))
    try:
        from .engine import handreq
        pc_on = handreq.pc_reading_on(None)
    except Exception:
        pc_on = False
    try:
        from .engine import privacy
        lines.append(_line(True, "외부 전송", privacy.note()))
    except Exception:
        pass
    lines.append(_line(pc_on, "'PC 판독 사용.txt'",
                       "있음 — 이 PC 로 직접 읽습니다(한 장에 몇 분)" if pc_on
                       else "없음 — 위 두 길이 모두 없으면 '13. Claude 판독 요청 …zip' 을 만들어 둡니다"))
    if not key_on and not cli:
        lines.append("   → 13항을 자동으로 채우려면: 이 PC 에 Claude Code 를 깔거나(위 줄),")
        lines.append("     제품 폴더에 'PC 판독 사용.txt' 빈 파일을 두거나,")
        lines.append("     만들어진 '13. Claude 판독 요청 …zip' 을 Claude 대화에 올려 판독 json 을 받으세요.")
    lines.append("")

    lines.append("판정:")
    if really or (really is None and soffice):
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
