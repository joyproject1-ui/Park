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


def _plain(text):
    return re.sub(r"\s+", "", text or "")


def _sheets(path, preferred):
    """정해 둔 시트를 먼저, 그다음 나머지 시트 — 회사가 서식을 개정하면 시트 이름이 바뀐다.

    2026-09 디겐타 자료의 마스터파일은 시트가 '생산장비' 였다(예전 '제조시설적격성평가현황').
    이름을 정해 두면 개정될 때마다 그 항이 통째로 비므로, 머리행을 보고 고른다.
    """
    wb = load_workbook(path, data_only=True, read_only=True)
    order = ([preferred] if preferred in wb.sheetnames else []) + \
            [n for n in wb.sheetnames if n != preferred]
    for name in order:
        yield name, list(wb[name].iter_rows(values_only=True))


def _header(rows, test, limit=12):
    """머리행 번호와 그 칸 글자. 없으면 (None, None)."""
    for i, row in enumerate(rows[:limit]):
        cells = [_cell(v) for v in row]
        if test(cells):
            return i, cells
    return None, None


def _is_equipment_header(cells):
    joined = _plain(" ".join(cells))
    return "관리번호" in cells and "보고서문서번호" in joined


def _is_support_header(cells):
    return "관리번호" in cells and any("IQ 문서번호" in c for c in cells)


def equipment_docs(path, sheet="제조시설적격성평가현황"):
    """{관리번호: {"name": 장비명, "line": 라인, "docs": [(문서번호, 완료일), ...]}}

    마스터파일은 장비 한 대가 여러 행(연번 행 + 이어지는 행)이라, 관리번호가 나온 뒤
    관리번호 칸이 빈 행은 같은 장비의 문서로 본다.
    """
    rows = header = col = None
    for _name, sheet_rows in _sheets(path, sheet):
        header, cells = _header(sheet_rows, _is_equipment_header)
        if header is not None:
            rows = sheet_rows
            col = {name: cells.index(name) for name in cells if name}
            break
    if header is None:
        raise ValueError("설비 마스터파일의 머리행(관리번호·보고서 문서번호)을 찾지 못했습니다.")
    c_id = col["관리번호"]
    c_doc = col[next(k for k in col if "문서 번호" in k or "문서번호" in k)]
    # 완료일은 개정되며 '승인일자' 로 이름이 바뀌었다 — 둘 다 받는다.
    c_date = next((col[k] for k in ("완료일", "승인일자", "승인일") if k in col), None)
    c_name = next((col[k] for k in ("장비명", "설비명") if k in col), None)
    c_line = col.get("라인")
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
    rows = header = col = None
    for _name, sheet_rows in _sheets(path, sheet):
        header, cells = _header(sheet_rows, _is_support_header)
        if header is not None:
            rows = sheet_rows
            col = {name: j for j, name in enumerate(cells) if name}
            break
    if header is None:
        # 개정된 마스터파일은 제조지원설비도 제조설비와 같은 꼴이다(관리번호·보고서 문서번호·승인일자).
        # 읽지 못했다고 물러나면 10.3~10.5 가 통째로 빈다 — 같은 값을 그 꼴에서 읽어 온다.
        return _support_from_equipment(path)
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


def _support_from_equipment(path):
    """제조설비 꼴(관리번호·보고서 문서번호·승인일자)로 적힌 제조지원설비 마스터파일을 읽는다."""
    out = {}
    for key, entry in equipment_docs(path, sheet=None).items():
        by = latest_by_kind(entry["docs"])
        row = {"name": entry.get("name", ""), "system": entry.get("line", ""), "raw_id": key}
        for kind in ("DQ", "IQ", "OQ", "PQ"):
            row[kind] = by.get(kind, [])
        out[key.split()[0] if key.split() else key] = row
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
