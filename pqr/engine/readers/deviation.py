# -*- coding: utf-8 -*-
"""일탈발생보고서 / 일탈완료보고서 PDF — 11항 표와 7항 각주에 쓸 값."""
import re

from ..pdftext import read_layout, squash
from .blocks import form_rows, pick


def _grab(pattern, text, flags=0):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else None


def _block(text, start, end_labels):
    """start 라벨 뒤부터 end 라벨 중 먼저 나오는 곳 앞까지의 글 — 라벨 열(영문 설명)을 걷어낸다."""
    i = text.find(start)
    if i < 0:
        return ""
    rest = text[i + len(start):]
    cut = len(rest)
    for label in end_labels:
        j = rest.find(label)
        if 0 <= j < cut:
            cut = j
    lines = []
    for line in rest[:cut].split("\n"):
        line = line.strip()
        if not line or re.match(r"^[A-Za-z()/ ,&.'-]+$", line):   # 영문 설명 줄
            continue
        line = re.sub(r"^(계획된 일탈 내용|또는 일탈 발생|현황|필요한 시정 사항|또는 즉시 취해진|시정 사항|일탈의 \(예상\)|원인)\s*", "", line)
        if line:
            lines.append(line)
    return "\n".join(lines)


def read_deviation(path):
    pages = read_layout(path)
    text = squash("\n".join(pages))
    out = {}
    out["doc_no"] = _grab(r"(DR-\d{6}-\d{2})", text)
    out["title"] = _grab(r"일탈명\s*\n?\s*(?:Title\s*)?\n?\s*(.+?)\n", text)
    if out["title"] and "Title" in out["title"]:
        out["title"] = _grab(r"Title\s*\n\s*(.+?)\n", text)
    out["lot"] = _grab(r"\(([A-Z]{2}[A-Z0-9]{4})\)", text)
    out["occurred"] = _grab(r"Occurred \(Expected\)\s+(\d{4}-\d{2}-\d{2})", text) \
        or _grab(r"발생\(예정\)일시.*?(\d{4}-\d{2}-\d{2})", text, re.S)
    out["planned"] = "✓ 계획된 일탈" in text or "☑ 계획된 일탈" in text
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        rows1 = form_rows(pdf.pages[0])
        rows2 = form_rows(pdf.pages[1]) if len(pdf.pages) > 1 else []
    out["description"] = "\n".join(pick(rows1, ("계획된일탈내용",)))
    out["correction"] = "\n".join(pick(rows1, ("필요한시정사항",)))
    out["cause"] = "\n".join(pick(rows1, ("일탈의(예상)",)))
    out["reviewer_comment"] = "\n".join(pick(rows1, ("검토자의견",)))
    out["qa_review"] = "\n".join(pick(rows2, ("검토의견",)))
    out["level"] = _grab(r"[☑✓]\s*(Critical|Major|Minor)", text)
    out["quality_impact"] = None
    m = re.search(r"제품\s*품질\s*영향.*?([☑✓☐])\s*있음\s*Yes\s*([☑✓☐])\s*없음\s*No\s*([☑✓☐])\s*있음\s*Yes\s*([☑✓☐])\s*없음\s*No", text, re.S)
    if m:                      # 유사 일탈(앞 둘) · 제품 품질 영향(뒤 둘)
        out["similar"] = m.group(1) in "☑✓"
        out["quality_impact"] = m.group(3) in "☑✓"
    # 완료일: 일탈완료보고서의 '조치사항 완료일' > 없으면 품질보증팀장 서명일
    out["completed"] = _grab(r"Completion Date of\s+(\d{4}[.\-]\d{2}[.\-]\d{2})", text) \
        or _grab(r"조치사항 완료일.*?(\d{4}[.\-]\d{2}[.\-]\d{2})", text, re.S)
    qa_dates = re.findall(r"QA1T\s+\S+\s+(\d{4}-\d{2}-\d{2})", text)
    out["qa_signed"] = qa_dates[0] if qa_dates else None
    if not out["completed"] and out["qa_signed"]:
        out["completed"] = out["qa_signed"]
    if out["completed"]:
        out["completed"] = out["completed"].replace("-", ".")
    out["has_completion_report"] = "일탈완료보고서" in text and "Completion Date" in text
    return out
