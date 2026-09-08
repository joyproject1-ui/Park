# -*- coding: utf-8 -*-
"""만든 보고서를 PC 의 Claude Code 가 첨부 자료와 대조해 문의 목록에 보탠다 — '보고서 작성' 의 마지막 단계.

담당자 2026-09-08: "대시보드에서 보고서 만들기를 누르면 네가 위에 말한 Cowork 절차를 실행해 주면
안 돼?" — 절차의 ⑤ '결과를 눈으로 검토' 를 단추에 붙인 것이다. ⑥ '코드 고치기' 는 붙이지
않는다: 단추를 누를 때마다 프로그램이 스스로 규칙을 바꾸면 어제 되던 것이 오늘 안 되고 왜
그런지 아무도 모르게 된다. 검토가 "코드가 원인" 이라고 하면 그 줄을 제작자에게 보낸다.

Claude Code 는 Read·Glob 만 쓴다(자료를 바꾸지 않는다). 보고서는 .docx 를 못 읽으므로 항·표를
글로 풀어 `PQR 검토용 본문 - <코드>.txt` 로 놓고, 엑셀도 첫 시트를 글로 풀어 옆에 둔다.
"""
import os
import re
import shutil
import tempfile

from docx import Document
from docx.oxml.ns import qn

REVIEW_TEXT = "PQR 검토용 본문 - %s.txt"          # 'PQR ' 로 시작해야 어느 항의 자료로도 셈해지지 않는다
MAX_FINDINGS = 20
XLSX_ROWS = 120

PROMPT = """다음은 자동으로 만든 제품품질평가(PQR) 보고서의 본문(표를 글로 푼 것)과, 그 근거가 된
첨부 자료 목록입니다. 보고서의 항 번호와 첨부 자료 번호는 같습니다 — 7항 표는 첨부 7(수율현황표),
9.2.4 표는 첨부 9.2.4(포장 완료 후 성적서), 13항은 첨부 13(안정성 시험일지)로 만듭니다.

보고서 본문:
%(report)s

첨부 자료 (항 번호 → 파일):
%(sources)s

이미 문의 목록에 적힌 것 (다시 적지 마세요):
%(issues)s

평가 기간: %(period)s — **이 기간 안에 생산·시험·완료된 것만** 보고서에 실립니다. 기간 뒤(작성 연도)에
끝난 안정성 시점이나 뒤에 생산된 Lot 이 보고서에 없는 것은 어긋남이 아닙니다. 13항 '시험 기간' 은
평가 기간 안에 완료된 시점만 적고, 13.3 경향은 평가 연도까지의 값만 씁니다.

할 일: 보고서의 채워진 칸을 같은 번호의 첨부 자료와 **대조**하여 어긋난 곳만 찾습니다.
· 첨부에 값이 있는데 보고서 칸이 비었거나 사선·'확인 필요' 인 것
· 보고서 값이 첨부의 값과 다른 것 (Lot 이 뒤바뀐 것, 단위·자릿수가 아니라 숫자 자체가 다른 것)
· 첨부에 있는 Lot·문서·설비가 보고서에서 빠진 것, 또는 그 반대
스캔본(글자 없는 PDF)은 직접 열어 보고 읽습니다. 확실하지 않으면 적지 않습니다 — 지어내지 않습니다.
서식·글씨·줄 간격은 보지 않습니다. 최대 %(max)d건.

JSON 만 출력하세요:
{"findings": [{"item": "9.2.4", "lot": "LWY201", "why": "제제균일성 칸이 비었는데 성적서에는 2.7%% 가 있음",
               "cause": "자료|코드"}]}
· item: 항 번호. lot: 제조번호나 문서번호(없으면 빈 값). why: 한 문장, 첨부의 값을 함께.
· cause: 첨부 자료가 없거나 담당자가 정할 일이면 "자료", 자료가 있는데 프로그램이 못 읽은 것이면 "코드".
"""


def dump_report(path):
    """보고서 .docx 를 항 차례대로 글로 — 문단은 그대로, 표는 줄마다 '칸 | 칸 | …'."""
    document = Document(path)
    lines = []
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            text = "".join(t.text or "" for t in child.iter(qn("w:t"))).strip()
            if text:
                lines.append(text)
        elif child.tag == qn("w:tbl"):
            for tr in child.findall(qn("w:tr")):
                cells = []
                for tc in tr.findall(qn("w:tc")):
                    text = " ".join("".join(t.text or "" for t in p.iter(qn("w:t"))).strip()
                                    for p in tc.findall(qn("w:p")))
                    cells.append(re.sub(r"\s+", " ", text).strip() or "-")
                lines.append("  | " + " | ".join(cells) + " |")
            lines.append("")
    return "\n".join(lines)


def _dump_xlsx(path, out_dir):
    """엑셀 첫 시트들을 글로 — Claude Code 의 Read 는 .xlsx 를 못 읽는다."""
    try:
        from openpyxl import load_workbook
        book = load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return None
    lines = []
    for ws in book.worksheets[:3]:
        lines.append("[시트] %s" % ws.title)
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= XLSX_ROWS:
                lines.append("  … (이하 생략)")
                break
            if any(v is not None for v in row):
                lines.append("  | " + " | ".join("" if v is None else str(v) for v in row) + " |")
    dst = os.path.join(out_dir, os.path.basename(path) + ".txt")
    try:
        with open(dst, "w", encoding="utf-8") as handle:
            handle.write("\n".join(lines))
    except OSError:
        return None
    return dst


def collect_sources(folder, workdir, log=None):
    """[(항, 경로)] — 압축은 풀고, 엑셀은 글로 풀어 그 경로를 준다."""
    from . import collect
    say = log or (lambda *a: None)
    got = collect.discover(folder, workdir)
    out = []
    for item in sorted(got, key=lambda k: [int(x) if x.isdigit() else 0 for x in k.split(".")]):
        for path in got[item]:
            low = path.lower()
            if low.endswith((".xlsx", ".xlsm", ".xls")):
                text = _dump_xlsx(path, workdir) if not low.endswith(".xls") else None
                if text:
                    out.append((item, text))
                    continue
            if low.endswith((".pdf", ".txt", ".docx", ".json")):
                out.append((item, path))
    say("  검토 자료 %d개" % len(out))
    return out


def _findings(text):
    from .claude_cli import _json_object
    got = _json_object(text)
    out = []
    for one in (got.get("findings") or [])[:MAX_FINDINGS]:
        item = str((one or {}).get("item") or "").strip()
        why = str(one.get("why") or "").strip()
        if not item or not why:
            continue
        cause = str(one.get("cause") or "").strip()
        tail = " — 프로그램이 못 읽은 것으로 보임: 이 줄을 제작자에게 보내 주세요" if cause == "코드" else ""
        out.append((item, str(one.get("lot") or "").strip(), "검토: %s%s" % (why, tail)))
    return out


def review(folder, product, report_path, issues, log=None, period=None):
    """보고서를 첨부와 대조한 검토 결과 [(항, Lot, 설명)]. Claude Code 가 없거나 실패하면 []."""
    say = log or (lambda *a: None)
    if os.environ.get("PQR_REVIEW", "").strip() == "0":     # 시험이나 급한 작성에서 끄는 스위치
        say("검토: PQR_REVIEW=0 이라 건너뜁니다")
        return []
    try:
        from . import claude_cli
        if not claude_cli.available():
            say("검토: 이 PC 에 Claude Code 가 없어 건너뜁니다 (PQR-Claude설치.bat)")
            return []
    except Exception:
        return []
    work = tempfile.mkdtemp(prefix="pqr-review-")
    try:
        code = (product or {}).get("code", "")
        made = os.path.dirname(os.path.abspath(report_path))
        text_path = os.path.join(made, REVIEW_TEXT % code)
        with open(text_path, "w", encoding="utf-8") as handle:
            handle.write(dump_report(report_path))
        sources = collect_sources(folder, work, say)
        period = period or {}
        span = "%s ~ %s" % (period.get("from") or "?", period.get("to") or "?")
        prompt = PROMPT % {
            "period": span,
            "report": text_path,
            "sources": "\n".join("  %s → %s" % (item, path) for item, path in sources) or "  (없음)",
            "issues": "\n".join("  [%s] %s — %s" % (i, f, w) for i, f, w in (issues or [])[:60]) or "  (없음)",
            "max": MAX_FINDINGS,
        }
        say("검토: Claude Code 가 보고서를 첨부 %d개와 대조합니다 (몇 분 걸립니다)" % len(sources))
        answer = claude_cli._ask(claude_cli._exe(), prompt, folder, say,
                                 [text_path] + [p for _, p in sources])
        found = _findings(answer)
        say("검토: %d건을 문의 목록에 보탭니다" % len(found))
        return found
    except Exception as error:                      # 검토가 넘어져도 보고서는 이미 있다
        say("검토를 하지 못했습니다 — %s" % error)
        return []
    finally:
        shutil.rmtree(work, ignore_errors=True)
