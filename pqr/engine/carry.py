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
from lxml import etree

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


def section_grids_all(old_document, section):
    """전년도 결재본에서 section(예: '10.1') 아래의 표를 모두 글자 그대로 — [[[칸 글자]], …].
    내수용·수출용처럼 하위 항(10.1.1·10.1.2)이 있으면 그 차례대로 담긴다."""
    out, here = [], False
    for kind, value, _ in outline(old_document):
        if kind == "h":
            got = _section(value)
            if got:
                here = got == section or got.startswith(section + ".")
        elif here:
            out.append(_grid_text(old_document.tables[value]))
    return out


def stability_packs(old_document):
    """전년도 결재본 13항 표의 '포장 형태' — {"by_lot": {제조번호: 포장}, "by_market": {"내수"|"수출": 포장}}.
    올해 시험일지에서 포장을 못 읽은 Lot(퀴노비드 '1Tube/Gab')은 여기서 채운다."""
    by_lot, counts, market = {}, {}, None
    market_by_lot, prefix_markets = {}, {}
    for kind, value, _ in outline(old_document):
        if kind == "h":
            got = _section(value)
            if not got:
                continue
            if not got.startswith("13"):
                market = None
            elif "수출" in value:
                market = "수출"
            elif "내수" in value or re.match(r"^13(\.\d+)?[.\s]*$|^13(\.\d+)?[.\s]", value):
                market = "내수"
            continue
        if market is None:
            continue
        table = old_document.tables[value]
        heads = _headers(table)
        lot_col = next((k for k, h in enumerate(heads) if "제조번호" in h.replace(" ", "")), None)
        pack_col = next((k for k, h in enumerate(heads) if "포장" in h), None)
        if lot_col is None or pack_col is None:
            continue
        for cells in _rows(table):
            lot = squeeze(E.cell_text(cells[lot_col])) if lot_col in cells else ""
            pack = E.cell_text(cells[pack_col]).strip() if pack_col in cells else ""
            for code in LOT.findall(lot):
                market_by_lot.setdefault(code, market)
                prefix_markets.setdefault(code[:2], set()).add(market)
            if pack:
                for code in LOT.findall(lot):
                    by_lot.setdefault(code, pack)
                counts.setdefault(market, {}).setdefault(pack, 0)
                counts[market][pack] += 1
    by_market = {m: max(c.items(), key=lambda kv: kv[1])[0] for m, c in counts.items()}
    # 제조번호 앞 두 글자가 한 시장에만 쓰였으면(내수 OE…, 수출 OA…·OZ…) 새 Lot 의 시장도 그것으로 본다
    market_by_prefix = {pf: next(iter(ms)) for pf, ms in prefix_markets.items() if len(ms) == 1}
    return {"by_lot": by_lot, "by_market": by_market,
            "market_by_lot": market_by_lot, "market_by_prefix": market_by_prefix}


EQUIP = re.compile(r"^[A-Z]{3}\d{4}")


def equipment_rows(old_document):
    """전년도 결재본 10.2~10.5 설비 표 — {"10.3": [{"mid", "name", "docs": {"IQ": (문서, 완료일), …}}, …], …}.
    빈 공양식의 10.3~10.5 는 관리번호·설비명이 없어, 여기서 줄을 세우고 마스터파일로 문서를 채운다."""
    out = {}
    for section, tables in _tables_by_section(old_document).items():
        if section not in ("10.2", "10.3", "10.4", "10.5"):
            continue
        for table in tables:
            grid = _grid_text(table)
            if not grid or len(grid[0]) < 5:
                continue
            kind_col = {}
            for row in grid[:4]:
                for ci, t in enumerate(row):
                    if squeeze(t).upper() in ("IQ", "OQ", "PQ") and squeeze(t).upper() not in kind_col:
                        kind_col[squeeze(t).upper()] = ci
            rows = []
            for i, row in enumerate(grid):
                mid = squeeze(row[1]) if len(row) > 1 else ""
                if not EQUIP.match(mid) or i + 1 >= len(grid):
                    continue
                docs = {}
                for kind, ci in kind_col.items():
                    doc = (row[ci] if ci < len(row) else "").strip()
                    day = (grid[i + 1][ci] if ci < len(grid[i + 1]) else "").strip()
                    if doc:
                        docs[kind] = (doc, day)
                rows.append({"mid": mid, "name": (row[2] if len(row) > 2 else "").strip(), "docs": docs})
            if rows:
                out.setdefault(section, []).extend(rows)
    return out


_PERIOD = re.compile(r"^(\d{1,3})\s*M$", re.I)


def _months(period):
    m = _PERIOD.match(squeeze(period or ""))
    return int(m.group(1)) if m else None


def stability_entries(old_document):
    """전년도 결재본 13항 실시 내역 — Lot 마다 하나.

    [{"kind": "장기"|"시판후", "market": "내수"|"수출", "lot", "lot_text", "year", "pack", "store",
      "periods": [...], "dones": [...], "last": 실시 사유|비고, "ongoing": bool, "range": {열 이름: "최소 ~ 최대"}}]
    'ongoing' 은 올해도 이어지는 시험 — 장기는 마지막 시점이 36M 미만, 시판 후는 비고가 '완료' 가 아닐 때.
    'range' 는 같은 시장 13.3 경향 표의 그 Lot 줄(해마다 Lot 차례가 같다)의 값이다.

    담당자 2026-09: "안정성 자료가 모두 입력이 안됐네 — 잘 모르겠으면 작년 PQR 에서 정보를 가져오고
    첨부된 안정성 자료로 값을 입력". 올해 시험일지를 읽지 못한 Lot 은 여기서 옮기고 확인을 남긴다.
    """
    entries, trends = [], {}
    kind, market = None, "내수"
    for what, value, _ in outline(old_document):
        if what == "h":
            got = _section(value)
            if not got:
                continue
            if not got.startswith("13"):
                kind = None
                continue
            t = squeeze(value)
            if got.count(".") <= 1:
                kind = "시판후" if "시판" in t else ("경향" if "경향" in t else ("장기" if "장기" in t else None))
                market = "내수"
            if "수출" in t:
                market = "수출"
            elif "내수" in t:
                market = "내수"
            continue
        if not kind:
            continue
        table = old_document.tables[value]
        if kind == "경향":
            trends.setdefault(market, []).extend(_trend_rows(table))
            continue
        heads = _headers(table)
        col = lambda *words: next((k for k, h in enumerate(heads) if any(w in h for w in words)), None)
        c_lot, c_year, c_period = col("제조번호"), col("해당연도", "연도"), col("시험기간")
        c_pack, c_store, c_done = col("포장"), col("보관"), col("완료")
        c_last = col("실시사유", "비고")
        if c_lot is None:
            continue
        for row in _grid_vmerged(table)[1:]:
            text = lambda c: (row[c] if c is not None and c < len(row) else "").strip()
            lot_text = text(c_lot)
            codes = LOT.findall(squeeze(lot_text))
            if not codes:
                continue
            periods = [p for p in text(c_period).split("\n") if p.strip()]
            dones = [d for d in text(c_done).split("\n") if d.strip()]
            prev = entries[-1] if entries else None
            if prev and prev["kind"] == kind and prev["market"] == market and prev["lot"] == codes[0]:
                prev["periods"] += periods                      # 시점마다 한 줄인 표(2025 결재본)
                prev["dones"] += dones
                for k in ("year", "pack", "store", "last"):
                    prev[k] = prev[k] or text({"year": c_year, "pack": c_pack, "store": c_store, "last": c_last}[k])
                continue
            entries.append({"kind": kind, "market": market, "lot": codes[0], "lot_text": lot_text,
                            "year": text(c_year), "pack": text(c_pack), "store": text(c_store),
                            "periods": periods, "dones": dones, "last": text(c_last), "range": {}})
    for e in entries:
        months = [m for m in (_months(p) for p in e["periods"]) if m is not None]
        if e["kind"] == "시판후":
            e["ongoing"] = "완료" not in squeeze(e["last"])
        else:
            e["ongoing"] = not months or max(months) < 36
    # 13.3 경향 줄 ↔ Lot: 같은 시장·같은 구분·같은 해 안에서 차례가 같다
    for market, rows in trends.items():
        for label, year, values in rows:
            kind_ = "시판후" if "시판" in label else "장기"
            same = [e for e in entries if e["market"] == market and e["kind"] == kind_ and e["year"] == year]
            taken = [e for e in same if e["range"]]
            if len(taken) < len(same):
                same[len(taken)]["range"] = values
    return entries


def _grid_vmerged(table):
    """_grid_text 와 같되 세로 병합으로 이어진 칸(vMerge 계속)에는 위 칸의 글을 넣는다 —
    시점마다 한 줄인 13항 표에서 제조번호·연도·실시 사유가 첫 줄에만 있는 것을 줄마다 되살린다."""
    grid = table._tbl.find(qn("w:tblGrid"))
    width = len(grid.findall(qn("w:gridCol"))) if grid is not None else 0
    out = []
    for row in table.rows:
        line, col = [""] * width, 0
        for cell in E.raw_cells(row):
            pr = cell._tc.find(qn("w:tcPr"))
            span_el = pr.find(qn("w:gridSpan")) if pr is not None else None
            span = int(span_el.get(qn("w:val"))) if span_el is not None else 1
            vm = pr.find(qn("w:vMerge")) if pr is not None else None
            text = E.cell_text(cell)
            if vm is not None and vm.get(qn("w:val")) != "restart" and out and not text.strip():
                text = out[-1][col] if col < width else ""
            for k in range(col, min(col + span, width)):
                line[k] = text
            col += span
        out.append(line)
    return out


def _trend_rows(table):
    """13.3 경향 표의 연도 줄 — [(줄 이름 '장기'|'시판후', 연도, {열 이름: 값 글})]."""
    grid = _grid_text(table)
    first = next((i for i, row in enumerate(grid) if squeeze(row[0]) in ("장기", "시판후")), None)
    if first is None:
        return []
    names = grid[first - 1] if first >= 1 else []
    if not any(squeeze(n) and "시험항목" not in n for n in names[2:]) and first >= 2:
        names = grid[first - 2]
    out, label = [], ""
    for row in grid[first:]:
        head = squeeze(row[0])
        if head in ("장기", "시판후"):
            label = head
        elif head:
            break
        year = re.match(r"(\d{4})", squeeze(row[1]) if len(row) > 1 else "")
        if not year:
            continue
        values = {}
        for k in range(2, len(row)):
            name = squeeze(names[k]) if k < len(names) else ""
            if "시험항목" in name or not (row[k] or "").strip():
                continue
            if name in values:
                continue
            values[name or "함량"] = row[k].strip()
        out.append((label, year.group(1), values))
    return out


# ---------- 공양식에 전년도 표 뼈대 옮겨 심기 ----------
# 담당자 PC 2026-09-07 (한림포비돈점안액): (v) 0항에 올린 빈 공양식이 바탕이 되면서 9.1 시험항목·허용기준,
# 9.2 열 머리글, 10.2 설비 목록, 13.3 경향표 머리글, 4항 문안이 통째로 없어져 전부 사선이 됐다
# ("도대체 왜 작성이 안 되는 거야?"). 새 양식(표지·항 구성·문안)은 그대로 쓰되, 그 항의 표가 빈 뼈대뿐이면
# 전년도 결재본의 그 항 내용(표·소제목·문단)을 옮겨 심는다 — 그 뒤 채우기는 전년도 결재본을 바탕으로
# 할 때와 똑같이 돌아간다.
SKELETON = ("4", "9.1", "9.2.1", "9.2.2", "9.2.3", "9.2.4", "10.2", "13.3")
EMPTYISH = ("", "N/A", "-", "—")


def _spans(document):
    """{항 번호: (제목 el, [본문 el …])} — 제목부터, 그 항 아래가 아닌 다음 제목 앞까지."""
    body = document.element.body
    items = [el for el in body if el.tag != qn("w:sectPr")]
    heads = []
    for i, el in enumerate(items):
        if el.tag != qn("w:p"):
            continue
        text = " ".join(_text_of(el).split())
        sec = _section(text)
        if sec and len(text) < 80:
            heads.append((i, sec))
    out = {}
    for n, (i, sec) in enumerate(heads):
        stop = len(items)
        for j, other in heads[n + 1:]:
            if not (other == sec or other.startswith(sec + ".")):
                stop = j
                break
        out.setdefault(sec, (items[i], items[i + 1:stop]))
    return out


def _text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _cells_text(tbl):
    """[[칸 글 …] …] — 표의 줄마다 원 칸(raw) 글."""
    rows = []
    for tr in tbl.findall(qn("w:tr")):
        rows.append([" ".join(_text_of(tc).split()) for tc in tr.findall(qn("w:tc"))])
    return rows


def _blank_skeleton(section, elements):
    """그 항이 빈 뼈대뿐인가 — 항마다 '값이 있어야 할 자리' 가 다르다. '※' 안내 블록(Cpk 판정표)은 안 본다."""
    before = []
    for el in elements:
        if el.tag == qn("w:p") and _text_of(el).strip().startswith("※"):
            break
        before.append(el)
    elements = before
    tables = [el for el in elements if el.tag == qn("w:tbl")]
    paras = [el for el in elements if el.tag == qn("w:p")]
    if section == "4":
        return not tables and not any(_text_of(p).strip() for p in paras)
    if not tables:
        return False
    grids = [_cells_text(t) for t in tables]
    if section == "9.1":                      # 시험항목 칸이 하나도 안 적혀 있다
        return not any(len(r) > 1 and r[1] for g in grids for r in g[1:])
    if section.startswith("9.2"):             # 열 머리글(시험항목 이름)이 비어 있다
        return not any(c for g in grids for c in g[0][2:])
    if section == "10.2":                     # 관리번호 적힌 줄이 없다
        return not any(r[0].isdigit() and len(r) > 1 and r[1] not in EMPTYISH for g in grids for r in g)
    if section == "13.3":                     # 시험항목 머리글이 비어 있다
        return not any(c for g in grids for c in g[0][1:])
    return False


_DROP = ("w:drawing", "w:pict", "w:object", "w:bookmarkStart", "w:bookmarkEnd", "w:proofErr",
         "w:commentRangeStart", "w:commentRangeEnd", "w:commentReference", "w:footnoteReference",
         "w:endnoteReference", "w:lastRenderedPageBreak", "w:numPr", "w:sectPr")


def _clean(el, style_ids):
    """다른 문서에서 온 요소를 이 문서에 심을 수 있게 다듬는다 — 없는 스타일·번호·그림·필드·책갈피를 뗀다."""
    import copy
    el = copy.deepcopy(el)
    for tag in _DROP:
        for bad in list(el.iter(qn(tag))):
            parent = bad.getparent()
            if parent is not None:
                parent.remove(bad)
    for tag in ("w:pStyle", "w:rStyle", "w:tblStyle"):
        for st in list(el.iter(qn(tag))):
            if st.get(qn("w:val")) not in style_ids:
                st.getparent().remove(st)
    for link in list(el.iter(qn("w:hyperlink"))):          # 링크는 풀고 글만 남긴다
        parent = link.getparent()
        at = parent.index(link)
        for child in list(link):
            parent.insert(at, child)
            at += 1
        parent.remove(link)
    for fld in list(el.iter(qn("w:fldSimple"))):
        parent = fld.getparent()
        at = parent.index(fld)
        for child in list(fld):
            parent.insert(at, child)
            at += 1
        parent.remove(fld)
    for run in list(el.iter(qn("w:r"))):                    # 필드 코드(목차·쪽 번호 참조) 런은 뗀다
        if run.find(qn("w:fldChar")) is not None or run.find(qn("w:instrText")) is not None:
            run.getparent().remove(run)
    for br in list(el.iter(qn("w:br"))):
        if br.get(qn("w:type")) == "page":
            br.getparent().remove(br)
    for fonts in el.iter(qn("w:rFonts")):                  # LibreOffice 가 남긴 '굴림;Gulim' 꼴
        for key in list(fonts.attrib):
            val = fonts.get(key)
            if val and ";" in val:
                fonts.set(key, val.split(";")[0])
    return el


def _style_ids(document):
    try:
        return {s.style_id for s in document.styles}
    except Exception:
        return set()


def _heading_template(document, spans):
    """소제목(10.2.1 …)을 세울 때 본뜰 공양식의 셋째 단계 제목 문단."""
    for sec, (head, _) in spans.items():
        if sec.count(".") == 2:
            return head
    return None


def _as_heading(text, template, style_ids):
    """전년도 소제목 글을 공양식 제목 문단 꼴로 — 글자만 바꾼다."""
    import copy
    if template is None:
        return None
    para = _clean(template, style_ids)
    runs = para.findall(qn("w:r"))
    for run in runs[1:]:
        para.remove(run)
    if runs:
        run = runs[0]
        for t in run.findall(qn("w:t")):
            run.remove(t)
        t = etree.SubElement(run, qn("w:t"))
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    else:
        run = etree.SubElement(para, qn("w:r"))
        t = etree.SubElement(run, qn("w:t"))
        t.text = text
    return para


def _take(section, elements):
    """전년도 항에서 옮길 것 — '※' 안내 블록 앞까지, 빈 문단 뭉치는 하나로, 끝의 빈 문단은 뗀다."""
    got = []
    for el in elements:
        text = _text_of(el).strip() if el.tag == qn("w:p") else None
        if text is not None and text.startswith("※"):
            break
        if text == "" and got and got[-1].tag == qn("w:p") and not _text_of(got[-1]).strip():
            continue
        if text == "" and not got:
            continue
        got.append(el)
    while got and got[-1].tag == qn("w:p") and not _text_of(got[-1]).strip():
        got.pop()
    return got


def _tc_clear(tc):
    """칸의 글을 지운다 — 첫 문단(서식)만 남긴다."""
    paras = tc.findall(qn("w:p"))
    for para in paras[1:]:
        tc.remove(para)
    if paras:
        for child in list(paras[0]):
            if child.tag != qn("w:pPr"):
                paras[0].remove(child)
    else:
        etree.SubElement(tc, qn("w:p"))


def _blank_values(section, tbl):
    """옮겨 심은 표에서 전년도 값을 비운다 — 뼈대(시험항목·허용기준·열 머리글·요약 줄 이름)만 남긴다.

    작년 숫자가 올해 칸에 그대로 보이면 안 된다. 9.1 은 결과 열, 9.2 는 Lot 줄(하나만 남김)과
    최댓값·최솟값·평균 줄의 값 칸.
    """
    trs = tbl.findall(qn("w:tr"))

    def first(tr):
        tcs = tr.findall(qn("w:tc"))
        return " ".join(_text_of(tcs[0]).split()) if tcs else ""

    if section == "9.1":
        for tr in trs[1:]:
            tcs = tr.findall(qn("w:tc"))
            if tcs:
                _tc_clear(tcs[-1])
    elif section.startswith("9.2"):
        data = [tr for tr in trs if first(tr).isdigit()]
        summary = [tr for tr in trs if first(tr).startswith(("최댓", "최솟", "평균"))]
        for tr in data[1:]:
            tbl.remove(tr)
        for tr in data[:1] + summary:
            for tc in tr.findall(qn("w:tc"))[1:]:
                _tc_clear(tc)


def _shift_years(text, period):
    """4항 문안의 연도를 올해 평가 기간에 맞춘다 — '2024년 1월 ~ 12월 … 2025년도 1분기' → 한 해씩."""
    if not period or not period.get("from"):
        return text
    years = [int(y) for y in re.findall(r"(?<!\d)(20\d{2})(?!\d)", text)]
    if not years:
        return text
    delta = int(str(period["from"])[:4]) - min(years)
    if not delta:
        return text
    return re.sub(r"(?<!\d)(20\d{2})(?!\d)", lambda m: str(int(m.group(1)) + delta), text)


def _set_para_text(para, text):
    runs = para.findall(qn("w:r"))
    first = None
    for run in runs:
        ts = run.findall(qn("w:t"))
        if ts and first is None:
            first = ts[0]
            for extra in ts[1:]:
                run.remove(extra)
        elif ts:
            for t in ts:
                run.remove(t)
    if first is None:
        run = etree.SubElement(para, qn("w:r"))
        first = etree.SubElement(run, qn("w:t"))
    first.text = text
    first.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def adopt_skeleton(document, old_document, period=None, log=None):
    """공양식의 빈 항에 전년도 결재본의 표 뼈대를 옮겨 심는다. 옮긴 항 목록을 돌려준다.

    옮기는 항: 4(문안, 연도만 올해로), 9.1(시험항목·허용기준), 9.2.1~9.2.4(열 머리글이 든 표들),
    10.2(라인별 소제목과 설비 목록), 13.3(시험항목 머리글). 그 항의 표에 이미 값이 있으면 건드리지 않는다.
    공양식에만 있는 '※ 안내' 블록은 남긴다.
    """
    new_spans, old_spans = _spans(document), _spans(old_document)
    style_ids = _style_ids(document)
    template = _heading_template(document, new_spans)
    moved = []
    for section in SKELETON:
        here, there = new_spans.get(section), old_spans.get(section)
        if not here or not there:
            continue
        head, elements = here
        if not _blank_skeleton(section, elements):
            continue
        take = _take(section, there[1])
        if not any(el.tag == qn("w:tbl") for el in take) and section != "4":
            continue
        # 공양식 쪽: '※' 안내 블록부터는 남기고 그 앞의 빈 뼈대를 걷어 낸다
        keep_from = None
        for i, el in enumerate(elements):
            if el.tag == qn("w:p") and _text_of(el).strip().startswith("※"):
                keep_from = i
                break
        for el in (elements if keep_from is None else elements[:keep_from]):
            el.getparent().remove(el)
        # 전년도 쪽을 다듬어 제목 뒤에 차례로 심는다
        anchor = head
        for el in take:
            if el.tag == qn("w:p"):
                text = " ".join(_text_of(el).split())
                sub = _section(text)
                if sub and len(text) < 80 and template is not None:       # 소제목(10.2.1 …)
                    new = _as_heading(text, template, style_ids)
                elif section == "4" and text:
                    new = _clean(el, style_ids)
                    _set_para_text(new, _shift_years(text, period))
                else:
                    new = _clean(el, style_ids)
            else:
                new = _clean(el, style_ids)
                if section == "9.1" or section.startswith("9.2"):
                    _blank_values(section, new)
            anchor.addnext(new)
            anchor = new
        if keep_from is None or not (anchor.tag == qn("w:p") and not _text_of(anchor).strip()):
            blank = etree.SubElement(document.element.body, qn("w:p"))   # 표 뒤에는 빈 문단 하나
            document.element.body.remove(blank)
            anchor.addnext(blank)
        moved.append(section)
    if "9.1" in moved:
        # 쪽을 나누려고 둘로 갈라 둔 9.1 표(조제·충전 / 포장)는 하나로 — 둘째 표를 수출용으로 오해하지 않게
        tables = _tables_by_section(document).get("9.1") or []
        if len(tables) > 1:
            E.join_continuations(tables)
    if log and moved:
        log("공양식이 빈 뼈대뿐인 항은 전년도 결재본의 표를 옮겨 심음: %s" % ", ".join(moved))
    return moved
