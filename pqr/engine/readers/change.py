# -*- coding: utf-8 -*-
"""변경요청서(CC) PDF — 12항 변경관리 표."""
import re

from ..pdftext import read_layout, squash
from .blocks import form_rows, pick


def _grab(pattern, text, flags=0):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else None


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
    dates = re.findall(r"(\d{4}-\d{2}-\d{2})", text)
    out["approved"] = None
    m = re.search(r"승인.*?(\d{4}-\d{2}-\d{2})", text, re.S)
    if m:
        out["approved"] = m.group(1).replace("-", ".")
    out["all_dates"] = sorted(set(dates))
    return out
