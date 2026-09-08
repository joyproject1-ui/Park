# -*- coding: utf-8 -*-
"""7항 수율현황표를 읽는다.

회사 수율현황표는 제품마다 머리 부분이 다르다 — 어떤 것은 첫 줄이 바로 공정 이름이고, 어떤
것은 '연번 / 중요공정 별 수율 현황 (%)' 아래에 공정·기준·Lot No. 줄이 더 있고 Lot 은 둘째
칸이다(디겐타안연고). 자리를 정해 두면 다른 제품에서 값을 통째로 놓친다 — 실제로 디겐타
2026 자료에서 아홉 칸이 모두 '확인 필요' 로 나왔다. 그래서 자리를 정하지 않고 표를 보고 찾는다.
"""
import re

from openpyxl import load_workbook

# 제조번호: 영문과 숫자가 섞인 5~7 글자 (OGY301, OEY101 …). '조제'·'기준' 같은 말은 걸리지 않는다.
LOT = re.compile(r"^(?=.*[A-Z])(?=.*\d)[A-Z0-9]{5,7}$")


def _number(value):
    if value is None:
        return None
    try:
        return float(str(value).replace("%", "").replace(",", "").strip())
    except ValueError:
        return None


def _grid(path):
    """표를 [[칸 값]] 로. 가로·세로로 병합된 칸은 덮는 자리 모두에 같은 값을 넣는다.

    담당자 2026-09-08: 수율현황표의 '포장' 머리 칸이 내수·베트남 두 열에 걸쳐 병합돼 있어
    베트남 열의 이름이 빈 값이 되고, 그 열의 수율(99.75)이 통째로 버려졌다.
    """
    book = load_workbook(path, data_only=True)          # 병합 정보를 보려면 read_only 를 쓸 수 없다
    ws = book.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not rows:
        return [], 0
    width = max(len(r) for r in rows)
    rows = [r + [None] * (width - len(r)) for r in rows]
    for spot in getattr(ws, "merged_cells", None).ranges if getattr(ws, "merged_cells", None) else []:
        r0, c0, r1, c1 = spot.min_row - 1, spot.min_col - 1, spot.max_row - 1, spot.max_col - 1
        if not (0 <= r0 < len(rows) and 0 <= c0 < width):
            continue
        value = rows[r0][c0]
        if value is None:
            continue
        for r in range(r0, min(r1 + 1, len(rows))):
            for c in range(c0, min(c1 + 1, width)):
                if rows[r][c] is None:
                    rows[r][c] = value
    return rows, width


def _lot_column(rows, width):
    """Lot 처럼 생긴 값이 가장 많은 열. 없으면 None."""
    hits = {}
    for row in rows:
        for c in range(width):
            if LOT.match(str(row[c] or "").strip()):
                hits[c] = hits.get(c, 0) + 1
    return max(sorted(hits), key=lambda c: hits[c]) if hits else None



SPEC = re.compile(r"이상|이하|±|~")


def _layout(path):
    """(표, Lot 열, 자료 줄 번호들, 값 열들, {값 열: 공정 이름}). 못 읽으면 (None, ...)."""
    rows, width = _grid(path)
    lot_col = _lot_column(rows, width)
    if lot_col is None:
        return None, None, [], [], {}
    data_rows = [i for i, row in enumerate(rows) if LOT.match(str(row[lot_col] or "").strip())]
    value_cols = [c for c in range(lot_col + 1, width)
                  if any(_number(rows[i][c]) is not None for i in data_rows)]
    if not value_cols:
        return None, None, [], [], {}
    names = _column_names(rows, value_cols, data_rows[0])
    return rows, lot_col, data_rows, value_cols, names


def _column_names(rows, value_cols, first_data):
    """열마다 머리 이름 — 머리가 여러 줄이면 위에서 아래로 이어 '포장(내수)' 로 만든다.

    담당자 2026-09-08: 아이퓨어 수율현황표는 머리가 두 줄이다 —
    첫 줄 '조제 | 충전 | 포장 | 포장', 둘째 줄 '조제 | 충전 | 내수 | 온누리에이치엔씨'.
    가장 아래 줄만 보면 이름이 '내수'·'온누리에이치엔씨' 가 되어 '포장' 이라는 공정을 잃는다.
    기준 줄('98.0 ± 2.0%')처럼 숫자가 든 칸은 이름이 아니다.
    """
    # 표 제목('중요공정 별 수율 현황(%)')은 모든 값 열에 같은 글로 퍼져 있다 — 이름이 아니다
    title_rows = set()
    for i in range(0, first_data):
        texts = {str(rows[i][c] or "").strip() for c in value_cols}
        if len(texts) == 1 and texts != {""}:
            title_rows.add(i)
    names = {}
    for c in value_cols:
        parts = []
        for i in range(0, first_data):
            if i in title_rows:
                continue
            text = str(rows[i][c] or "").strip()
            if not text or SPEC.search(text) or re.search(r"\d", text) or "%" in text:
                continue                      # 기준 칸이거나 표 제목('…수율 현황 (%)')
            text = re.sub(r"[(（][^)）]*[)）]", "", text).strip() or text
            if text and (not parts or text != parts[-1]):
                parts.append(text)
        if parts:
            names[c] = parts[0] if len(parts) == 1 else "%s(%s)" % (parts[-2], parts[-1])
    return names





def read_specs(path):
    """{공정명: 기준 글자} — 그해 수율현황표에 적힌 기준 줄. 없으면 {}.

    기준은 해가 바뀌며 개정된다(디겐타안연고 충전: 96.0 ± 3.5% → 86.5 ± 6.5%). 전년도 결재본에
    적힌 기준으로 견주면 멀쩡한 Lot 이 모두 '기준 벗어남' 으로 잡힌다 — 기준도 그해 자료에서 읽는다.
    """
    rows, lot_col, data_rows, value_cols, names = _layout(path)
    if rows is None:
        return {}
    # 기준이 한 줄에 다 있지 않다 — 조제·충전은 위 줄에, 포장은 그 아래 줄에 적힌 표가 있다
    # (담당자 2026-09-08 한림포비돈 수율현황표). 자료 줄 위쪽을 모두 훑어 열마다 하나씩 모은다.
    got = {}
    for i in range(data_rows[0] - 1, -1, -1):
        texts = {c: str(rows[i][c] or "").strip() for c in value_cols}
        if not any(SPEC.search(t) for t in texts.values()):
            continue
        for c, t in texts.items():
            if names.get(c) and t and SPEC.search(t) and names[c] not in got:
                got[names[c]] = t
    if not got:
        return {}
    # 시장별로 갈린 공정('포장(내수)'·'포장(베트남)')은 기준이 하나뿐일 때 나눠 쓴다
    def _base(text):
        return re.sub(r"\s", "", re.sub(r"[(（][^)）]*[)）]", "", text or ""))

    for name in names.values():
        if not name or name in got:
            continue
        같은것 = {v for k, v in got.items() if _base(k) == _base(name)}
        if len(같은것) == 1:
            got[name] = 같은것.pop()
    return got


def read_yields(path):
    """[(Lot, {공정명: 값문자열}), ...] — 값은 소수점 둘째 자리 문자열로 맞춘다."""
    rows, lot_col, data_rows, value_cols, names = _layout(path)
    if rows is None:
        return []
    out = []
    for i in data_rows:
        vals = {}
        for c in value_cols:
            name = names.get(c)
            if not name:
                continue
            number = _number(rows[i][c])
            if number is None:
                text = str(rows[i][c] or "").strip()
                if text:
                    vals[name] = text
            else:
                vals[name] = "%.2f" % number
        out.append((str(rows[i][lot_col]).strip(), vals))
    return out
