# -*- coding: utf-8 -*-
"""품목허가(신고)증 PDF — 3항 대상 제품 표에 들어갈 값."""
import re

from ..pdftext import read_layout, squash


def read_license(path):
    pages = read_layout(path, pages={1, 2})
    text = squash("\n".join(pages))
    out = {}
    m = re.search(r"제\s*(\d+)\s*호", text)
    out["license_no"] = ("제 %s 호" % m.group(1)) if m else None
    m = re.search(r"제품명\S*\s+(\S+)", text)
    out["product_name"] = m.group(1) if m else None
    m = re.search(r"성상\s+(.+?)\n", text)
    out["appearance"] = m.group(1).strip() if m else None
    m = re.search(r"저장방법 및 사용\(유효\)기간\s+(.+?)\n", text)
    if m:
        out["storage_and_shelf"] = m.group(1).strip()
        s = m.group(1)
        mm = re.search(r"제조일로부터\s*(\d+)\s*개월", s)
        out["shelf_life"] = ("제조일로부터 %s개월" % mm.group(1)) if mm else None
        out["storage"] = s.split("제조일")[0].strip(" ,") if "제조일" in s else s.strip()
    m = re.search(r"(\d{4}\.\d{2}\.\d{2})\s+최초허가\(신고\)일자", text)
    out["first_approved"] = m.group(1) if m else None
    m = re.search(r"품목기준코드\s+(\d+)", text)
    out["item_code"] = m.group(1) if m else None
    m = re.search(r"\[\s*[V∨v]\s*\]\s*(전문|일반)", text)
    out["category"] = m.group(1) if m else None
    out["domestic"] = "내수용" in text
    return out
