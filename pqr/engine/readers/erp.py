# -*- coding: utf-8 -*-
"""ERP 에서 내려받은 표 — 주원료·자재 시험번호(.xls), 제조내역(PDF)."""
import os
import re

from ..pdftext import read_text, squash


def _sheet_rows(path):
    if path.lower().endswith(".xls"):
        import xlrd
        book = xlrd.open_workbook(path)
        if not book.codepage:
            # ERP 가 내려 준 .xls 에는 코드페이지 기록이 없어 xlrd 가 iso-8859-1 로 읽는다 —
            # 제품명이 'µð°ÕÅ¸¾È¿¬°í' 처럼 깨져 제품을 가려낼 수 없다.
            book = xlrd.open_workbook(path, encoding_override="cp949")
        sheet = book.sheet_by_index(0)
        return [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    from openpyxl import load_workbook
    ws = load_workbook(path, data_only=True, read_only=True).worksheets[0]
    return [list(row) for row in ws.iter_rows(values_only=True)]


def test_no(raw):
    """시험 성적 번호는 ERP 에 적힌 그대로 쓴다 — 'R202305090013'.

    담당자 2026-09-08: "시험성적번호는 오른쪽처럼 하이픈 없이 작성해줘." 전에는 날짜로 끊어
    'R-2023-05-09-0013' 으로 적었는데, 결재본(2025 올로원스점안액)은 붙여 쓴다.
    """
    return str(raw or "").strip()


# ERP 가 그대로 내려 주는 표는 머리행이 없고 열이 열아홉이다. 쓰는 칸은 다음과 같다.
RAW_CODE, RAW_NAME, RAW_PRODUCT, RAW_TEST, RAW_LOT = 0, 3, 7, 10, 16
CODE = re.compile(r"^[A-Z]{1,3}\d{3,6}$")                       # RBG201 · P17043
LOT = re.compile(r"^[A-Z0-9]{5,7}$")                            # OGY301


def _looks_raw(rows):
    """ERP 가 그대로 내려 준 표인지 — 머리행 없이 첫 칸이 원료 코드이고 열이 넉넉하다."""
    for row in rows[:3]:
        cells = [str(c or "").strip() for c in row]
        if len(cells) > RAW_LOT and CODE.match(cells[RAW_CODE]):
            return True
    return False


def read_material_tests(path, product=None):
    """원료/자재 코드 · 시험번호 · Lot 목록.  [(코드, 시험번호, Lot), ...] — Lot 없는 행은 뺀다.

    두 가지 꼴을 받는다.
      1) 사람이 추린 표: 머리행(원료 코드 · 시험번호 · Lot No.) + 세 칸
      2) ERP 가 그대로 내려 준 표: 머리행 없이 열아홉 칸. 담당자가 받는 것은 이쪽이다
         (2026-09 디겐타 자료). 이 꼴을 못 읽어 8.2 표에 지난해 값이 그대로 남았다.

    ERP 표는 원료 하나를 쓴 모든 제품이 함께 들어 있으므로(RSF101 은 후메론·톨론티 … 백여 줄)
    제품 이름으로 이 제품 줄만 가려낸다. Lot No. 는 열일곱째 칸에 그대로 있다 — 지어내지 않는다.
    """
    rows = _sheet_rows(path)
    if not _looks_raw(rows):
        out = []
        for row in rows[1:]:
            cells = [str(c or "").strip() for c in row[:3]]
            if len(cells) < 3 or not cells[2]:
                continue
            out.append((cells[0], test_no(cells[1]), cells[2]))
        return out

    key = _plain_product(product)
    out = []
    for row in rows:
        cells = [str(c or "").strip() for c in row] + [""] * (RAW_LOT + 1)
        code, lot, test = cells[RAW_CODE], cells[RAW_LOT], cells[RAW_TEST]
        if not CODE.match(code) or not LOT.match(lot) or not test:
            continue                       # 제품에 쓰이지 않은 입고·시험 줄
        if key and not _same_product(cells[RAW_PRODUCT], key):
            continue                       # 같은 원료를 쓰는 다른 제품 줄
        record = (code, test_no(test), lot)
        if record not in out:
            out.append(record)
    return out


def _plain_product(text):
    """제품 이름에서 빈칸과 괄호 설명을 뗀다 — '아이퓨어점안액[일회용]' → '아이퓨어점안액'.

    담당자 2026-09-08: ERP 표의 제품명은 '아이퓨어점안액[일회용]' 인데 제품 코드의 이름은
    '아이퓨어점안액' 이라 정확히 견주면 한 줄도 남지 않고 8.2 표가 통째로 사선이 됐다.
    """
    return re.sub(r"\s+", "", re.sub(r"[\[(（【][^\])）】]*[\])）】]", "", str(text or "")))


def _same_product(made, key):
    """ERP 표의 제품명이 이 제품인가 — 괄호 설명을 뗀 이름이 서로의 앞부분이면 같은 제품."""
    made = _plain_product(made)
    if not made or not key:
        return False
    return made == key or made.startswith(key) or key.startswith(made)


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


# ERP 제조번호 줄 — 'NJS2-2025-L0K2Y-2000101 2025.02.13 2027.02.12' (나조린 2026-09-10: 제품 코드·연도·
# Lot 글자와 달이 갈라져 적힌 꼴이라 LOT_LINE 에 걸리지 않아 '제조내역 0줄' 로 남았다).
# 셋째 마디 'L0K2Y' 는 Lot 글자(L·K·Y) 사이에 제조 달(02)이 끼어 있고, 넷째 마디 '2000101' 은
# 달 글자(2·N·D …) 뒤에 번호가 오며 마지막 자리가 그 달의 몇째 Lot 인지다 → LKY201.
ERP_LINE = re.compile(r"^\s*[A-Z0-9]+-(\d{4})-([A-Z])(\d)([A-Z])(\d)([A-Z])-([A-Z0-9])(\d+)\s+"
                      r"(\d{4}\.\d{2}\.\d{2})\s+(\d{4}\.\d{2}\.\d{2})\s*$", re.M)
MONTH_LETTER = {10: "O", 11: "N", 12: "D"}


def lot_from_erp(m):
    """ERP_LINE 의 match → Lot No.(LKY201) — 달이 맞지 않으면 None."""
    a, m1, b, m2, c, tail_month, tail = m.group(2), m.group(3), m.group(4), m.group(5), m.group(6), m.group(7), m.group(8)
    month = int(m1 + m2)
    if not 1 <= month <= 12:
        return None
    letter = MONTH_LETTER.get(month, str(month))
    if tail_month != letter:
        return None
    seq = int(tail[-1]) if tail else 0
    if not seq:
        return None
    return "%s%s%s%s%02d" % (a, b, c, letter, seq)


def read_manufacturing(path):
    """제조내역 PDF — [(Lot, 품명, 제조일자, 사용기한), ...]"""
    text = squash(read_text(path))
    out = [(m.group(1), m.group(2).strip(), m.group(3), m.group(4)) for m in LOT_LINE.finditer(text)]
    if out:
        return out
    for m in ERP_LINE.finditer(text):
        lot = lot_from_erp(m)
        if lot:
            out.append((lot, "", m.group(9), m.group(10)))
    return out


def read_material_names(path):
    """ERP 표의 {원료 코드: 원료명} — 8.2 표의 원료명 칸에 쓴다.

    담당자 2026-09-08: "8.2.2 표의 내용은 자재 코드와 자재명이 맞아, 점안제인데 자재가 튜브로
    잘못 기록되어 있어. 해당 첨부 파일을 제대로 확인해줘." 이름을 코드에 붙일 길이 없어
    안연고 문안('튜브 (내수용)')을 그대로 쓰고 있었다 — 이름은 ERP 표 넷째 칸에 있다.
    """
    rows = _sheet_rows(path)
    if not _looks_raw(rows):
        return {}
    out = {}
    for row in rows:
        cells = [str(c or "").strip() for c in row] + [""] * (RAW_LOT + 1)
        code, name = cells[RAW_CODE], cells[RAW_NAME]
        if CODE.match(code) and name and code not in out:
            out[code] = name
    return out
