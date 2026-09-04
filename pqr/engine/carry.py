# -*- coding: utf-8 -*-
"""전년도 결재본에서 제품 고유 정보를 이어받는다.

담당자 말 그대로다(2026-09): “전년도 PQR 결재본은 장비 및 원료 등 정보를 참고하는거고,
올해 자료를 참고해서 새롭게 작성해야 하는거야.”

그래서 이 단계는 **빈 칸에만** 값을 넣는다. 올해 자료로 채운 칸은 건드리지 않는다.
옮겨 오는 것은 해마다 바뀌지 않는 값뿐이다 — 원료 규격(KP·USP·자사규격), 제조단위·포장단위,
포장 형태·보관 조건. 시험값·일자·수량처럼 해마다 다른 것은 옮기지 않는다.
"""
import re

from docx.oxml.ns import qn

from . import docedit as E
from .locate import outline

# 해마다 바뀌지 않는 열만 옮긴다
SAME_EVERY_YEAR = ("규격", "제조단위", "포장단위", "포장형태", "보관조건", "제형", "제품분류",
                   "제조원", "문서번호", "제조업체")
KEY_WORDS = ("관리번호", "코드", "점검항목")


def squeeze(text):
    return re.sub(r"[\s ]+", "", text or "")


def _section(text):
    m = re.match(r"^(\d{1,2}(?:\.\d+)*)[.\s]", text or "")
    return m.group(1) if m else None


def _tables_by_section(document):
    """{항 번호: [표 …]} — 표 바로 앞의 번호 붙은 제목으로 묶는다."""
    out, here = {}, None
    for kind, value, _ in outline(document):
        if kind == "h":
            got = _section(value)
            if got:
                here = got
        elif here:
            out.setdefault(here, []).append(document.tables[value])
    return out


def _headers(table):
    """열 이름 (머리행 여러 줄이면 이어 붙여). 그리드 열 수만큼."""
    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    width = len(grid.findall(qn("w:gridCol"))) if grid is not None else 0
    trs = tbl.findall(qn("w:tr"))
    if not width or not trs:
        return []
    out, col = [""] * width, 0
    for tc in trs[0].findall(qn("w:tc")):
        pr = tc.find(qn("w:tcPr"))
        span_el = pr.find(qn("w:gridSpan")) if pr is not None else None
        span = int(span_el.get(qn("w:val"))) if span_el is not None else 1
        text = squeeze("".join(t.text or "" for t in tc.iter(qn("w:t"))))
        for k in range(col, min(col + span, width)):
            out[k] = text
        col += span
    return out


def _rows(table):
    """[(그리드 열 번호 → 셀)] — 자료 행만 (머리행 하나 뺀 나머지)."""
    grid = table._tbl.find(qn("w:tblGrid"))
    width = len(grid.findall(qn("w:gridCol"))) if grid is not None else 0
    out = []
    for row in table.rows[1:]:
        cells, col = {}, 0
        for cell in E.raw_cells(row):
            pr = cell._tc.find(qn("w:tcPr"))
            span_el = pr.find(qn("w:gridSpan")) if pr is not None else None
            span = int(span_el.get(qn("w:val"))) if span_el is not None else 1
            if col < width:
                cells[col] = cell
            col += span
        out.append(cells)
    return out


def _key_col(headers):
    for k, h in enumerate(headers):
        if any(w in h for w in KEY_WORDS):
            return k
    return None


def _wanted(headers):
    return {k: h for k, h in enumerate(headers) if any(w in h for w in SAME_EVERY_YEAR)}


def carry(document, old_document, log=None):
    """빈 칸에 전년도 값을 넣는다. 넣은 칸 수를 돌려준다."""
    new_by, old_by = _tables_by_section(document), _tables_by_section(old_document)
    done = 0
    for section, tables in sorted(new_by.items()):
        olds = old_by.get(section)
        if not olds:
            continue
        for i, table in enumerate(tables):
            if i >= len(olds):
                break
            old = olds[i]
            head, old_head = _headers(table), _headers(old)
            want = _wanted(head)
            if not want:
                continue
            if [h for h in head if h] != [h for h in old_head if h]:
                continue                       # 열 구성이 다르면 옮기지 않는다
            key, old_key = _key_col(head), _key_col(old_head)
            old_rows = _rows(old)
            by_key = {}
            for cells in old_rows:
                if old_key is not None and old_key in cells:
                    by_key.setdefault(squeeze(E.cell_text(cells[old_key])), cells)
            for cells in _rows(table):
                code = squeeze(E.cell_text(cells[key])) if key is not None and key in cells else ""
                source = by_key.get(code) if code else None
                for col, name in want.items():
                    cell = cells.get(col)
                    if cell is None or E.cell_text(cell).strip():
                        continue
                    text = ""
                    if source is not None and col in source:
                        text = E.cell_text(source[col]).strip()
                    elif key is None:
                        # 줄 열쇠가 해마다 다른 표(6항 Lot No.) — 전년도 값이 한 가지면 그 값을 쓴다.
                        # 관리번호로 짝지을 수 있는 표에서는 짝이 없으면 가져오지 않는다 — 다른
                        # 원료의 규격을 끌어오면 안 된다.
                        seen = {E.cell_text(c[col]).strip() for c in old_rows if col in c}
                        seen = {s for s in seen if s}
                        text = seen.pop() if len(seen) == 1 else ""
                    if text:
                        E.set_cell(cell, *text.split("\n"))
                        done += 1
    if log:
        log("전년도에서 이어받은 칸: %d" % done)
    return done
