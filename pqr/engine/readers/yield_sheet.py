# -*- coding: utf-8 -*-
"""7항 수율현황표 — 담당자가 개인 작성하는 엑셀: 첫 열 Lot, 이어서 공정별 수율(%)."""
import re
from openpyxl import load_workbook


def read_yields(path):
    """[(Lot, {공정명: 값문자열}), ...] — 값은 소수점 둘째 자리 문자열로 맞춘다."""
    ws = load_workbook(path, data_only=True, read_only=True).worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not rows:
        return []
    header = [str(v or "").strip() for v in rows[0]]
    out = []
    for row in rows[1:]:
        lot = str(row[0] or "").strip()
        if not re.match(r"^[A-Z0-9]{5,7}$", lot):
            continue
        vals = {}
        for name, v in zip(header[1:], row[1:]):
            if not name or v is None or str(v).strip() == "":
                continue
            try:
                vals[name] = "%.2f" % float(str(v).replace("%", ""))
            except ValueError:
                vals[name] = str(v).strip()
        out.append((lot, vals))
    return out
