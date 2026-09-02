# -*- coding: utf-8 -*-
"""적격성 마스터파일 — 10.2 제조설비, 10.3~10.5 제조지원 설비의 IQ/OQ/PQ 문서번호·완료일."""
import re
from openpyxl import load_workbook

DOC = re.compile(r"^(DQ|IQ|OQ|PQ|IOQ|OPQ|QM)\d*")
DATE = re.compile(r"\d{4}[.\-/]\s?\d{1,2}[.\-/]\s?\d{1,2}")


def _cell(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def _kind(doc):
    m = re.match(r"^([A-Z]+)", doc or "")
    return m.group(1) if m else ""


def equipment_docs(path, sheet="제조시설적격성평가현황"):
    """{관리번호: {"name": 장비명, "line": 라인, "docs": [(문서번호, 완료일), ...]}}

    마스터파일은 장비 한 대가 여러 행(연번 행 + 이어지는 행)이라, 관리번호가 나온 뒤
    관리번호 칸이 빈 행은 같은 장비의 문서로 본다.
    """
    ws = load_workbook(path, data_only=True, read_only=True)[sheet]
    rows = list(ws.iter_rows(values_only=True))
    header = None
    for i, row in enumerate(rows[:10]):
        cells = [_cell(v) for v in row]
        if "관리번호" in cells and "보고서 문서 번호" in " ".join(cells):
            header = i
            col = {name: cells.index(name) for name in cells if name}
            break
    if header is None:
        raise ValueError("설비 마스터파일의 머리행(관리번호·보고서 문서 번호)을 찾지 못했습니다.")
    c_id = col["관리번호"]; c_doc = next(k for k in col if "문서 번호" in k or "문서번호" in k); c_doc = col[c_doc]
    c_date = col.get("완료일"); c_name = col.get("장비명"); c_line = col.get("라인")
    out, current = {}, None
    for row in rows[header + 1:]:
        cells = list(row) + [None] * 25
        mid = _cell(cells[c_id])
        if mid:
            current = mid
            out.setdefault(current, {"name": _cell(cells[c_name]) if c_name is not None else "",
                                     "line": _cell(cells[c_line]) if c_line is not None else "", "docs": []})
        if current is None:
            continue
        doc = _cell(cells[c_doc])
        if doc and DOC.match(doc):
            date = _cell(cells[c_date]) if c_date is not None else ""
            out[current]["docs"].append((doc, date))
    return out


def latest_by_kind(docs):
    """[(doc, date)] → {"IQ": (doc, date), "OQ": …, "PQ": …} 가장 최근 완료일 기준 (같은 종류 여러 건이면 전부)."""
    by = {}
    for doc, date in docs:
        by.setdefault(_kind(doc), []).append((doc, date))
    return by


def support_docs(path, sheet="제조지원 설비 & IT 시스템"):
    """{관리번호(앞 토큰): {"name": 설비명, "system": 시스템, "IQ": [(doc, date)], "OQ": …, "PQ": …, "DQ": …}}"""
    ws = load_workbook(path, data_only=True, read_only=True)[sheet]
    rows = list(ws.iter_rows(values_only=True))
    header = None
    for i, row in enumerate(rows[:12]):
        cells = [_cell(v) for v in row]
        if "관리번호" in cells and any("IQ 문서번호" in c for c in cells):
            header = i
            col = {name: j for j, name in enumerate(cells) if name}
            break
    if header is None:
        raise ValueError("제조지원 설비 마스터파일의 머리행을 찾지 못했습니다.")
    out = {}
    system = ""
    for row in rows[header + 1:]:
        cells = list(row) + [None] * 25
        sysname = _cell(cells[col["시스템"]]) if "시스템" in col else ""
        if sysname:
            system = sysname
        raw_id = _cell(cells[col["관리번호"]])
        if not raw_id:
            continue
        key = raw_id.split()[0]
        entry = {"name": _cell(cells[col.get("설비명", 0)]), "system": system, "raw_id": raw_id}
        for kind in ("DQ", "IQ", "OQ", "PQ"):
            dcol = next((col[k] for k in col if k.startswith(kind + " 문서")), None)
            acol = next((col[k] for k in col if k.startswith(kind + " 승인")), None)
            docs = [d for d in re.split(r"\s+", _cell(cells[dcol])) if d] if dcol is not None else []
            dates = DATE.findall(_cell(cells[acol])) if acol is not None else []
            entry[kind] = [(docs[i] if i < len(docs) else "", dates[i] if i < len(dates) else "")
                           for i in range(max(len(docs), len(dates)))]
        out[key] = entry
    return out


def pv_by_code(path, code, sheet=None):
    """PV 마스터에서 문서 코드(QUIO3 · QUIO2 …)가 든 계획/보고서 묶음을 제품명과 상관없이 찾는다.

    [{plan, reason, kind, report, report_date, lots:[(seq, lot, mfg)], revalidation}] — 계획 행 뒤에
    Lot 행(1st·2nd·3rd)이 이어진다.
    """
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = [[_cell(v) for v in r] for r in ws.iter_rows(values_only=True)]
    header = None
    for i, r in enumerate(rows[:15]):
        if any("Lot" in c or "제조번호" in c for c in r) and any("계획" in c or "PV" in c or "Plan" in c for c in r):
            header = i; break
    out, current = [], None
    for r in rows[(header + 1) if header is not None else 0:]:
        r = r + [""] * 14
        plan_col = next((i for i, c in enumerate(r) if re.match(r"^PV\d{2}-.*-P", c)), None)
        if plan_col is not None:
            plan = r[plan_col]
            if code not in plan.upper().replace("0", "O"):
                current = None; continue
            rep_col = next((i for i in range(plan_col + 1, len(r)) if re.match(r"^PV\d{2}-.*-R", r[i])), None)
            rep = r[rep_col] if rep_col is not None else ""
            m = re.search(r"(\d{4}\.\d{2}\.\d{2})", rep)
            current = {"plan": plan.split("(")[0].strip(), "reason": r[plan_col + 1], "kind": r[plan_col + 2],
                       "report": re.sub(r"\s*\(\d{4}\.\d{2}\.\d{2}\)?\s*$", "", rep).strip(),
                       "report_date": m.group(1) if m else "", "lots": [],
                       "revalidation": r[rep_col + 1] if rep_col is not None else ""}
            out.append(current)
        if current is None:
            continue
        seq_col = next((i for i, c in enumerate(r) if re.match(r"^\d(st|nd|rd|th)$", c)), None)
        if seq_col is not None and r[seq_col + 1]:
            current["lots"].append((r[seq_col], r[seq_col + 1], r[seq_col + 2]))
    return out
