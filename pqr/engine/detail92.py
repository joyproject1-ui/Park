# -*- coding: utf-8 -*-
"""9.2 세부 시험결과 표 채우기 — 열을 **머리글로** 짚는다.

제품마다, 해마다 열 구성이 다르다. 디겐타안연고 2026 양식은 함량이 성분별로 두 열이고
포장 표에 입자도·금속성이물이 함께 있는데, 퀴노비드안연고 결재본은 함량이 한 열이고 표가
내수용·수출용으로 나뉜다. 자리를 정해 두면 입자도 값이 함량 칸에 들어간다 — 실제로 그렇게
나왔다(2026-09). 그래서 머리행(여러 줄일 수 있다)을 이어 붙인 이름으로 열을 짚는다.

표 하나가 어느 공정 것인지도 머리글과 항 제목으로 정한다 — 9.2.x 번호는 제품마다 뜻이
다르다(디겐타는 공정별, 퀴노비드는 내수용·수출용).
"""
import re

from docx.oxml.ns import qn

from . import docedit as E

SUMMARY = ("최댓값", "최솟값", "평균", "공정능력지수", "Cpk")
HEAD_FIRST = ("연번", "no.", "번호")


def squeeze(text):
    return re.sub(r"[\s ]+", "", text or "")


def _grid_labels(tr, width):
    """그 행의 칸 글자를 그리드 열 수만큼 펼친다 (가로 병합은 같은 글자를 여러 열에)."""
    out, col = [""] * width, 0
    for tc in tr.findall(qn("w:tc")):
        pr = tc.find(qn("w:tcPr"))
        span_el = pr.find(qn("w:gridSpan")) if pr is not None else None
        span = int(span_el.get(qn("w:val"))) if span_el is not None else 1
        text = "".join(t.text or "" for t in tc.iter(qn("w:t")))
        for k in range(col, min(col + span, width)):
            out[k] = text
        col += span
    return out


def head_rows(table):
    """머리행 개수 — 첫 칸이 '연번'·'No.' 이거나 비어 있는 줄이 이어지는 만큼."""
    trs = table._tbl.findall(qn("w:tr"))
    n = 0
    for tr in trs:
        first = squeeze("".join(t.text or "" for t in tr.findall(qn("w:tc"))[0].iter(qn("w:t"))))
        if n == 0 or first.lower() in [squeeze(x) for x in HEAD_FIRST] or not first:
            if n and first and first.lower() not in [squeeze(x) for x in HEAD_FIRST]:
                break
            n += 1
            continue
        break
    return max(1, n)


def labels(table):
    """열마다 머리글을 이어 붙인 이름 (빈칸 없이). 그리드 열 수만큼."""
    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    width = len(grid.findall(qn("w:gridCol"))) if grid is not None else 0
    trs = tbl.findall(qn("w:tr"))
    if not width or not trs:
        return []
    rows = [_grid_labels(tr, width) for tr in trs[:head_rows(table)]]
    out = []
    for k in range(width):
        parts, seen = [], set()
        for row in rows:
            text = squeeze(row[k])
            if text and text not in seen:
                seen.add(text)
                parts.append(text)
        out.append("".join(parts))
    return out


def data_range(table):
    """(첫 자료 행, 마지막 자료 행, [요약 행 …]) — 요약은 최댓값·최솟값·평균·Cpk 줄."""
    trs = table._tbl.findall(qn("w:tr"))
    first = head_rows(table)
    summary = []
    for i in range(len(trs) - 1, first - 1, -1):
        text = squeeze("".join(t.text or "" for t in trs[i].findall(qn("w:tc"))[0].iter(qn("w:t"))))
        if any(squeeze(w) in text for w in SUMMARY):
            summary.insert(0, i)
        else:
            break
    last = (summary[0] - 1) if summary else (len(trs) - 1)
    return first, last, summary


# ---------- 허용기준 → 결과 문구 ----------
PREFIX = re.compile(r"^\s*(?:\[(?P<part>[^\]]*)\]\s*)?(?:자가|허가)\s*\)\s*", re.S)
PART = re.compile(r"\[([^\]]*)\]")


def criterion_text(text):
    """허용기준 글에서 결과 칸에 적을 문구를 만든다.

    한림 결재본이 그렇게 쓴다 — “자가) 튜브개봉 시 튜브 막힘이 없음” 은 결과도 같은 글이고,
    “… 식별 가능해야 한다.” 는 “… 식별 가능함” 으로 바꿔 적는다.
    """
    out = PREFIX.sub("", (text or "").strip())
    out = re.sub(r"[\s\u00a0]+", " ", out)          # 칸 안의 줄바꿈은 글자를 나눈 것이 아니라 줄 넘김이다
    out = re.sub(r"\s*해야\s*한다\.?\s*$", "함", out)
    return out.strip()


def criteria(table):
    """9.1 표에서 [{공정, 항목, 성분, 구분, 기준}] 을 뽑는다. 공정·항목은 병합된 위 칸을 잇는다."""
    out, process, item = [], "", ""
    for row in table.rows[1:]:
        cells = E.raw_cells(row)
        texts = [E.cell_text(c) for c in cells]
        ci = next((i for i, t in enumerate(texts) if PREFIX.search(t or "")), None)
        if ci is None:
            continue
        if texts[0].strip():
            process = squeeze(texts[0])
        if len(texts) > 1 and ci > 1 and texts[1].strip():
            item = squeeze(texts[1])
        sub = "".join(squeeze(t) for i, t in enumerate(texts[2:ci], 2)) + \
              "".join(squeeze(t) for t in texts[ci + 1:-1])
        m = PART.search(texts[ci] or "")
        out.append({"process": process, "item": item, "part": squeeze(m.group(1)) if m else "",
                    "sub": sub, "text": criterion_text(texts[ci])})
    return out


def criterion_for(rules, process, item, part="", sub=""):
    """공정·항목(·성분·구분)에 맞는 기준 문구. 공정이 다르면 다른 공정 것이라도 쓴다."""
    def score(r):
        s = 0
        if r["process"] and process and r["process"] == process:
            s += 8
        if item and item in r["item"]:
            s += 4
        if part:
            s += 2 if r["part"] and (r["part"] in part or part in r["part"]) else -2
        elif r["part"]:
            s -= 1
        if sub:
            s += 1 if sub in r["sub"] else -1
        return s
    best = max(rules, key=score) if rules else None
    if best is None or not item or item not in best["item"]:
        return None
    return best["text"]


# ---------- 값 채우기 ----------
NUM = re.compile(r"-?\d+(?:\.\d+)?")


def _decimals(texts):
    d = 0
    for t in texts:
        for m in NUM.finditer(t):
            if "." in m.group(0):
                d = max(d, len(m.group(0).split(".")[1]))
    return d


def _stats(label, texts):
    """열의 자료 글에서 (최댓값, 최솟값, 평균) 글을 만든다. 쓰지 않을 자리는 None.

    “a ~ b” 는 범위(개개)라 평균이 없고, “… 이상”·“… 이하” 는 뜻이 있는 쪽만 적는다 —
    한림 2026 결재본이 그렇게 쓴다.
    """
    lows, highs, plain = [], [], []
    for t in texts:
        got = [float(x) for x in NUM.findall(t)]
        if not got:
            return None, None, None
        if "~" in t and len(got) >= 2:
            lows.append(min(got)); highs.append(max(got))
        else:
            plain.append(got[0]); lows.append(got[0]); highs.append(got[0])
    fmt = "%%.%df" % _decimals(texts)
    ranged = len(plain) != len(texts)
    only_low = ("이상" in label) or any("이상" in t for t in texts)
    only_high = ("이하" in label) or any("이하" in t for t in texts)
    top = fmt % max(highs) if not only_low else None
    bottom = fmt % min(lows) if not only_high else None
    mean = None
    if plain and not ranged and not only_low and not only_high:
        mean = fmt % (sum(plain) / len(plain))
    return top, bottom, mean


def _grid_cells(row, width):
    """그리드 열 번호 → 그 행의 셀. 가로 병합은 첫 열에만 담는다."""
    out, col = {}, 0
    for cell in E.raw_cells(row):
        pr = cell._tc.find(qn("w:tcPr"))
        span_el = pr.find(qn("w:gridSpan")) if pr is not None else None
        span = int(span_el.get(qn("w:val"))) if span_el is not None else 1
        if col < width:
            out[col] = cell
        col += span
    return out


def fill(table, lots, value, cpk=None):
    """자료 행을 채우고 요약 행(최댓값·최솟값·평균·Cpk)을 계산해 넣는다.

    value(label, lot, i) — label 은 머리글을 이어 붙여 빈칸을 지운 열 이름. None 이면 그 칸은 둔다.
    cpk(label, values) — (지수, 판정) 또는 None.
    돌려주는 값: {열 이름: [자료 글 …]}
    """
    labs = labels(table)
    if not labs:
        return {}
    first, last, summary = data_range(table)
    f, l = E.fit_rows(table, first, last, max(1, len(lots)))
    summary = [ri + (l - last) for ri in summary]    # 자료 줄이 늘거나 줄면 요약 줄도 밀린다
    got = {}
    for i, lot in enumerate(lots):
        cells = _grid_cells(table.rows[f + i], len(labs))
        for k, lab in enumerate(labs):
            cell = cells.get(k)
            if cell is None:
                continue
            if "연번" in lab:
                text = str(i + 1)
            elif lab.lower().startswith("lotno") or lab.lower() == "no.":
                text = lot
            else:
                text = value(lab, lot, i)
            if text is None:
                continue
            E.set_cell(cell, *str(text).split("\n"))
            got.setdefault(k, []).append(str(text))
    if not summary:
        return {labs[k]: v for k, v in got.items()}
    stats = {}
    for k, texts in got.items():
        lab = labs[k]
        if "연번" in lab or lab.lower().startswith("lotno"):
            continue
        if "금속성이물" in lab and len(set(texts)) == 1:
            continue                       # 모두 같은 값이면 적지 않는다 — 표 아래 각주가 그렇게 밝힌다
        stats[k] = _stats(lab, texts)
    for ri in summary:
        cells = _grid_cells(table.rows[ri], len(labs))
        head = squeeze(E.cell_text(cells[0])) if cells.get(0) is not None else ""
        which = 0 if "최댓값" in head else 1 if "최솟값" in head else 2 if "평균" in head else None
        for k, three in stats.items():
            cell = cells.get(k)
            if cell is None:
                continue
            if which is not None:
                if three[which] is not None:
                    E.set_cell(cell, three[which])
            elif cpk is not None:
                pair = cpk(labs[k], got[k])
                if pair:
                    # 디겐타 2026 양식은 함량 열의 'Cpk' 칸과 'Cpk 판정 결과' 칸이 세로로 병합돼
                    # 있어 판정('충분')이 숨었다(2026-09 점검). 값을 쓰는 칸은 병합을 푼다.
                    if _has_vmerge(cell):
                        E.set_vmerge(cell, False)
                    E.set_cell(cell, pair[0] if "Cpk" in head and "판정" not in head else pair[1])
    return {labs[k]: v for k, v in got.items()}


def _has_vmerge(cell):
    pr = cell._tc.find(qn("w:tcPr"))
    return pr is not None and pr.find(qn("w:vMerge")) is not None


# ---------- 9.2 아래 표 모으기 ----------
SECTION = re.compile(r"^(\d{1,2}(?:\.\d+)*)[.\s]")
STAGES = ("조제", "충전", "포장")


def tables_92(document):
    """[(표, 공정)] — 9.2 아래 표들과 그 공정(조제·충전·포장).

    표 사이의 각주 줄('1) 모든 Lot 의 …')이 제목처럼 보여 항이 끊기므로, 번호가 제대로 붙은
    제목만 항의 끝으로 본다. 양식에 '9.2.1.3 포장 완료 후' 처럼 번호가 어긋난 곳이 있어
    공정은 번호가 아니라 제목 낱말로 정한다.
    """
    from .locate import outline
    out, inside, stage = [], False, ""
    for kind, value, _ in outline(document):
        if kind == "h":
            m = SECTION.match(value)
            if m and m.group(1).startswith("9.2"):
                inside = True
                for s in STAGES:
                    if s in value:
                        stage = s
            elif m and m.group(1) != "9":
                inside = False
            continue
        if inside:
            table = document.tables[value]
            if any(l.lower().startswith("lotno") for l in labels(table)):
                out.append((table, stage))          # Cpk 판정 기준 같은 안내표는 뺀다
    return out


def simple(table):
    """자료 블록이 하나뿐인 표인가 — 확인·수치 두 블록이 한 표에 든 결재본(퀴노비드)은 아니다."""
    trs = table._tbl.findall(qn("w:tr"))
    marks = [i for i, tr in enumerate(trs)
             if squeeze("".join(t.text or "" for t in tr.findall(qn("w:tc"))[0].iter(qn("w:t")))).startswith("최댓값")]
    return len(marks) <= 1
