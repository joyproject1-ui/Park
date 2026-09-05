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


# 왼쪽에 항목 이름이 오고 오른쪽이 그 값인 표 — 3항 '대상 제품' 이 그렇다.
LABEL_KEY = "점검항목"
ROW_NUMBER = ("No.", "연번", "번호")


def _label_columns(headers):
    """줄 이름이 곧 항목인 표의 값 열 — 항목 열과 연번 열만 뺀다.

    3항 대상 제품(제형·제품분류·제품명·허가번호·허가일자·사용기한·보관조건 …)은
    해마다 바뀌지 않고, 허가증 PDF 에 글자 정보가 없으면 읽을 데가 전년도 결재본뿐이다
    (담당자 2026-09: "3항 대상 제품 정보는 전년도 pqr 에서 정보를 가져와").
    """
    if not any(LABEL_KEY in h for h in headers):
        return {}
    return {k: h for k, h in enumerate(headers)
            if h and LABEL_KEY not in h and not any(w in h for w in ROW_NUMBER)}


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
            want = _wanted(head) or _label_columns(head)
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


# 제조번호는 여섯 자(OGY301). 여러 개가 줄바꿈 없이 붙어 있어도 갈라내려고 낱말 경계를 쓰지 않는다.
LOT = re.compile(r"[A-Z]{2}[A-Z0-9]{4}")


def pv_reasons(document):
    """10.1 공정밸리데이션 표에서 {제조번호: 실시 사유}.

    안정성 시험의 '실시 사유' 는 그 Lot 을 다시 시험하게 만든 변경, 곧 공정밸리데이션의
    사유다(담당자 2026-09: "실시 사유는 작년 내용 것을 그대로 가져오면 돼"). 전년도 결재본의
    10.1 에는 그해 PV 대상 Lot 과 사유가 적혀 있어, 지난해 Lot 의 사유를 거기서 물려받는다.
    """
    out = {}
    for section, tables in _tables_by_section(document).items():
        if not section.startswith("10.1"):
            continue
        for table in tables:
            head = _headers(table)
            lot_col = next((k for k, h in enumerate(head) if "Lot" in h or "제조번호" in h), None)
            why_col = next((k for k, h in enumerate(head) if "비고" in h), None)
            if lot_col is None or why_col is None:
                continue
            for cells in _rows(table):
                if lot_col not in cells or why_col not in cells:
                    continue
                why = " ".join(E.cell_text(cells[why_col]).split())
                if not why or why.upper() in ("N/A", "-"):
                    continue
                for lot in LOT.findall(E.cell_text(cells[lot_col]).replace("\n", " ")):
                    out.setdefault(lot, why)
    return out


def _grid_text(table):
    """[[그리드 열마다의 글자]] — 가로 병합은 덮는 열 모두에 같은 글자를 넣는다."""
    grid = table._tbl.find(qn("w:tblGrid"))
    width = len(grid.findall(qn("w:gridCol"))) if grid is not None else 0
    out = []
    for row in table.rows:
        line, col = [""] * width, 0
        for cell in E.raw_cells(row):
            pr = cell._tc.find(qn("w:tcPr"))
            span_el = pr.find(qn("w:gridSpan")) if pr is not None else None
            span = int(span_el.get(qn("w:val"))) if span_el is not None else 1
            text = E.cell_text(cell)
            for k in range(col, min(col + span, width)):
                line[k] = text
            col += span
        out.append(line)
    return out


def section_grids(old_document, *sections):
    """전년도 결재본의 항 표를 글자 그대로 읽는다 — {항: [[칸 글자]]}.

    서식이 개정되며 항 번호가 밀리므로(13.2 → 13.3), 앞에 적은 항부터 찾아 하나만 담는다.
    """
    by = _tables_by_section(old_document)
    out = {}
    for section in sections:
        for name in ([section] if isinstance(section, str) else section):
            got = by.get(name)
            if got:
                out[section if isinstance(section, str) else section[0]] = _grid_text(got[0])
                break
    return out


def stability_tables(old_document):
    """전년도 결재본의 13항 표를 글자 그대로 읽는다 — {"13.1": [[…]], "13.3": [[…]]}.

    담당자 2026-09: "안정성도 공란인데 전년도 PQR 결재본 참고해서 작성한 다음에
    13항 최신 안정성 시험 파일로 업로드해서 작성하면 돼."
    올해 시험일지를 읽지 못했을 때, 13항을 빈칸으로 두는 대신 여기서 읽은 전년도
    내용을 옮겨 놓고 '갱신 필요' 로 알린다. 서식은 해가 바뀌며 13.2 가 13.3 이 되었다.
    """
    return section_grids(old_document, "13.1", ("13.3", "13.2"))
