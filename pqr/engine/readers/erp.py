# -*- coding: utf-8 -*-
"""ERP 에서 내려받은 표 — 주원료·자재 시험번호(.xls), 제조내역(PDF)."""
import os
import re

from ..pdftext import read_text, squash


def _sheet_rows(path):
    if path.lower().endswith(".xls"):
        import xlrd
        book = xlrd.open_workbook(path)
        sheet = book.sheet_by_index(0)
        return [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    from openpyxl import load_workbook
    ws = load_workbook(path, data_only=True, read_only=True).worksheets[0]
    return [list(row) for row in ws.iter_rows(values_only=True)]


def test_no(raw):
    """ERP 표기 'R202405130070' → 보고서 표기 'R-2024-05-13-0070'."""
    s = str(raw or "").strip()
    m = re.match(r"^([A-Z])(\d{4})(\d{2})(\d{2})(\d{4})$", s)
    return "%s-%s-%s-%s-%s" % m.groups() if m else s


def read_material_tests(path):
    """원료/자재 코드 · 시험번호 · Lot 목록.  [(코드, 시험번호, Lot), ...] — Lot 없는 행은 뺀다."""
    rows = _sheet_rows(path)
    out = []
    for row in rows[1:]:
        cells = [str(c or "").strip() for c in row[:3]]
        if len(cells) < 3 or not cells[2]:
            continue
        out.append((cells[0], test_no(cells[1]), cells[2]))
    return out


def group_by_test(records):
    """같은 시험번호를 쓴 Lot 을 묶는다 (보고서 8.2 표의 세로 병합 단위). 순서는 첫 등장 순."""
    order, groups = [], {}
    for code, test, lot in records:
        key = (code, test)
        if key not in groups:
            groups[key] = []
            order.append(key)
        if lot not in groups[key]:
            groups[key].append(lot)
    return [(code, test, groups[(code, test)]) for code, test in order]


LOT_LINE = re.compile(r"^\s*([A-Z]{2}[A-Z0-9]{4})\s+(\S.*?)\s+(\d{4}\.\d{2}\.\d{2})\s+(\d{4}\.\d{2}\.\d{2})\s*$", re.M)


def read_manufacturing(path):
    """제조내역 PDF — [(Lot, 품명, 제조일자, 사용기한), ...]"""
    text = squash(read_text(path))
    return [(m.group(1), m.group(2).strip(), m.group(3), m.group(4)) for m in LOT_LINE.finditer(text)]
