# -*- coding: utf-8 -*-
"""변경요청서(CC) PDF — 12항 변경관리 표."""
import re

from ..pdftext import read_layout, squash
from .blocks import form_rows, pick


def _grab(pattern, text, flags=0):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else None


TEAM = re.compile(r"^\s*((?:QA|QC|RA)\d*팀|위수탁팀|생산\d*부|기술지원부|품질보증\s*\d*팀|품질관리\s*\d*팀|"
                  r"[가-힣A-Z0-9]{2,8}팀|[가-힣]{2,6}부)\s*")
NOISE = re.compile(r"부서/팀|조치사항|완료예정일|Departments|Teams|Follow-?up|Target Date|"
                   r"첨부문서|한림제약|Rev\.|변경 실행 계획|Change Execution|문서번호|변\s*경\s*명|"
                   r"완료\s*목표일|실행계획서|변경동의서|승인 후|N/?A")
DATE_TAIL = re.compile(r"\s*\d{4}[.\-]\d{2}[.\-]\d{2}\s*$")


def _actions(pages, doc_no=None, title=None):
    """'변경 실행 계획' 표의 부서별 조치사항 제목 — [(팀, 조치사항), ...]

    12항 조치사항은 이 표를 간추려 적는다(담당자 지시 2026-09). 딸린 줄('- QC-142 …')과 서식
    글자(문서번호·변경명·완료 목표일 …)는 빼고 제목 줄만 모은다. 위탁사·위수탁 관련 줄은
    부르는 쪽에서 걸러 낸다. 말은 지어내지 않고 원본 줄을 그대로 옮긴다.
    """
    whole = "\n".join(pages)
    start = whole.find("변경 실행 계획")
    if start < 0:
        return []
    block = whole[start:]
    end = block.find("첨부문서")
    if end > 0:
        block = block[:end]
    skip = {re.sub(r"\s+", "", x) for x in (doc_no or "", title or "") if x}
    out, team, seen = [], "", set()
    for raw in block.split("\n"):
        line = re.sub(r"\s+", " ", raw).strip()
        if not line or NOISE.search(line):
            continue
        m = TEAM.match(line)
        if m:
            team = re.sub(r"\s+", "", m.group(1))
            line = line[m.end():].strip()
        line = DATE_TAIL.sub("", line).strip()
        if not line or line.startswith(("-", "·", ":")):      # 딸린 줄은 제목이 아니다
            continue
        if re.match(r"^[\d.,\-/() ]+$", line) or line.startswith("("):
            continue
        key = re.sub(r"\s+", "", line)
        if key in skip or key in seen or len(key) < 3:
            continue
        seen.add(key)
        out.append((team, line))
    return out


def read_change(path):
    pages = read_layout(path)
    text = squash("\n".join(pages))
    out = {}
    out["doc_no"] = _grab(r"(CC-\d{6}-\d{2})", text)
    out["title"] = _grab(r"변\s*경\s*명\s*\n\s*(.+?)\n", text)
    out["target_date"] = _grab(r"완료 목표일.*?(\d{4}-\d{2}-\d{2})", text, re.S)
    out["products"] = _grab(r"관련 제품\s+(.+?)\n", text)
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        rows = form_rows(pdf.pages[0])
    out["reason"] = "\n".join(pick(rows, ("변경사유",)))
    out["description"] = "\n".join(pick(rows, ("변경내용",)))
    out["attachments"] = "\n".join(pick(rows, ("첨부문서",)))
    out["actions"] = _actions(pages, out.get("doc_no"), out.get("title"))
    dates = re.findall(r"(\d{4}-\d{2}-\d{2})", text)
    out["approved"] = None
    m = re.search(r"승인.*?(\d{4}-\d{2}-\d{2})", text, re.S)
    if m:
        out["approved"] = m.group(1).replace("-", ".")
    out["all_dates"] = sorted(set(dates))
    return out
