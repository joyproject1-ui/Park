# -*- coding: utf-8 -*-
"""공급업체 목록(8.1.1 원료 · 8.1.3 자재) 과 주성분 공급망 마스터파일(8.1.2)."""
import re
from openpyxl import load_workbook


def _cell(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


HEADER_WORDS = ("공급업체명", "원료코드", "자재코드", "문서번호", "평가승인일", "공급되는품목", "평가등급",
                "주성분", "제조소", "납품처")                      # 뒤 셋은 주성분 공급망 마스터파일
HEADER_ROWS = 15                                     # 제목·개정번호 줄 아래 이만큼 안에 머리행이 있다


def _latest_sheet(wb):
    """읽을 시트 — **'실시간' 시트가 있으면 그것**, 없으면 이름의 Rev. 번호가 가장 높은 개정 시트.

    담당자 2026-09-28: "실시간 시트 또는 가장 최종 작성본으로 읽어야돼". 목록 파일에는
    '202503(Rev.21)' … '202603(Rev.25)' 개정별 시트와 담당자가 계속 고치는 '실시간' 시트,
    '202210~개정이력' 이 함께 있다.
    """
    names = wb.sheetnames
    live = [n for n in names if n.strip().startswith("실시간")]   # 담당자가 계속 고치는 현재 시트
    if live:
        return live[0]
    revs = []
    for n in names:
        m = re.search(r"Rev\.?\s*(\d+)", n)
        revs.append((int(m.group(1)) if m else -1, n))
    return max(revs)[1]


def _is_header(row):
    """머리행인가 — 첫 칸이 'No.'/'NO'/'연번' 이거나, 목록 열 이름이 둘 이상 든 줄."""
    first = re.sub(r"[^a-z가-힣]", "", (row[0] if row else "").lower())
    if first in ("no", "연번", "번호"):
        return True
    flat = [re.sub(r"\s", "", c) for c in row]
    return sum(1 for c in flat if any(w in c for w in HEADER_WORDS)) >= 2


def _table(path):
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[_latest_sheet(wb)]
    rows = [[_cell(v) for v in r] for r in ws.iter_rows(values_only=True)]
    for i, row in enumerate(rows[:HEADER_ROWS]):
        if row and _is_header(row):
            # '원료코드 Material Code' → '원료코드' (영문 설명은 뗀다). 머리 칸의 줄바꿈은 _cell 이 빈칸으로 바꿨다.
            header = [re.sub(r"\s*[A-Za-z(].*$", "", c).strip() for c in row]
            return header, rows[i + 1:], ws.title
    raise ValueError("공급업체 목록 '%s' 시트의 위 %d줄에서 머리행(No. · 공급업체명 · 원료코드 …)을 찾지 못했습니다: %s"
                     % (ws.title, HEADER_ROWS, path))


def sheet_of(records):
    """읽은 줄들이 어느 시트에서 왔는지 — 기록에 적는다."""
    for rec in records or []:
        if rec.get("_sheet"):
            return rec["_sheet"]
    return ""


def read_supplier_list(path):
    """[{공급업체명, 제조소 국가, 공급되는 품목, 원료코드?, 평가방법, 평가등급, 문서번호, 개정번호, 평가승인일}, ...]"""
    header, rows, sheet = _table(path)
    out = []
    for row in rows:
        if not row or not row[0] or not any(row[1:]):
            continue
        rec = {h: (row[i] if i < len(row) else "") for i, h in enumerate(header) if h}
        rec["_sheet"] = sheet
        out.append(rec)
    return out


def find_supplier(records, code=None, item=None, name=None):
    """원료코드나 품목명(부분 일치)으로 한 줄 찾기."""
    for rec in records:
        if code and code in (rec.get("원료코드") or ""):
            return rec
    for rec in records:
        if item and item in (rec.get("공급되는 품목") or ""):
            if not name or name in (rec.get("공급업체명") or ""):
                return rec
    return None


def _pick(rec, *words):
    """열 이름에 낱말이 모두 든 칸의 값 — '주성분 명'·'주성분명', '제조소/국가'·'제조소 (국가)' 처럼 띄어쓰기가 달라도."""
    for key, value in rec.items():
        flat = re.sub(r"\s", "", str(key or ""))
        if all(w in flat for w in words):
            return value or ""
    return ""


def read_api_chain(path):
    """주성분 공급망 마스터파일 — {원료코드: {"api": 주성분명, "manufacturer": 제조소/국가, "chain": [1차, 2차, 3차], "_sheet": 시트}}"""
    header, rows, sheet = _table(path)
    out = {}
    for row in rows:
        if not row or not row[0]:
            continue
        rec = {h: (row[i] if i < len(row) else "") for i, h in enumerate(header) if h}
        code = (_pick(rec, "원료코드") or _pick(rec, "코드") or "").strip()
        if code:
            chain = [rec.get(k, "") for k in header if "납품처" in re.sub(r"\s", "", k)]
            out[code] = {"api": _pick(rec, "주성분"), "manufacturer": _pick(rec, "제조소"),
                         "chain": [c for c in chain if c and c != "N/A"], "_sheet": sheet}
    return out
