# -*- coding: utf-8 -*-
"""보고서에 딸린 엑셀 — Cpk 계산 파일(HLF-QC-126-08/09) 과 안정성 경향 분석(HLF-QC-126-06).

Cpk 파일은 전년도 결재본 압축에 든 것을 그대로 물려받아 값만 갈아 끼운다(수식·서식 보존).
안정성 경향 파일은 서식(HLF-QC-126-06)이 제품 폴더·입력 폴더의 '서식' 폴더·프로그램에
있으면 채운다.
"""
import os
import re
import shutil
import tempfile
import zipfile

from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill
from openpyxl.comments import Comment

from . import convert

CPK_ITEMS = [   # (파일 이름에 든 낱말, 완제 성적서 항목)
    ("금속성이물(개개)", "metal_each"), ("금속성이물(합계)", "metal_total"),
    ("입자도", "particle"), ("함량", "assay"),
]


def _num(v):
    m = re.search(r"-?\d+(?:\.\d+)?", str(v or ""))
    return float(m.group()) if m else None


def _previous_xls(previous_path, workdir):
    """결재본(.zip 또는 폴더 옆)의 Cpk .xls 들. {낱말: 경로}"""
    found = {}
    cands = []
    if previous_path and previous_path.lower().endswith(".zip"):
        with zipfile.ZipFile(previous_path) as z:
            for info in z.infolist():
                name = info.filename
                try:
                    name = name.encode("cp437").decode("cp949")
                except Exception:
                    pass
                if name.lower().endswith((".xls", ".xlsx")) and not os.path.basename(name).startswith("~$"):
                    target = os.path.join(workdir, os.path.basename(name))
                    with open(target, "wb") as h:
                        h.write(z.read(info.filename))
                    cands.append(target)
    else:
        folder = os.path.dirname(previous_path) if previous_path else ""
        if folder and os.path.isdir(folder):
            cands = [os.path.join(folder, n) for n in os.listdir(folder) if n.lower().endswith((".xls", ".xlsx"))]
    for path in cands:
        for word, _ in CPK_ITEMS:
            if word in os.path.basename(path) and "Cpk" in os.path.basename(path):
                found[word] = path
    return found


def fill_cpk(src_xls, dst_xlsx, values, today):
    """B10~B44 에 값, K4 에 작성일. 남는 칸은 비운다."""
    tmp = dst_xlsx + ".tmp.xlsx"
    convert.to_xlsx(src_xls, tmp)
    wb = load_workbook(tmp)
    ws = wb.worksheets[0]
    ws["K4"] = today
    for i in range(35):
        ws.cell(row=10 + i, column=2).value = values[i] if i < len(values) else None
    wb.save(dst_xlsx)
    os.remove(tmp)
    return dst_xlsx


def write_cpk_files(folder, data, previous_path, today, lots=None):
    """내수용 Lot 의 완제 성적서 값으로 Cpk 파일 4종을 제품 폴더에 만든다. [(이름, 경로)]"""
    lots = lots or data.domestic
    work = tempfile.mkdtemp(prefix="pqr-cpk-")
    sources = _previous_xls(previous_path, work)
    out = []
    for word, key in CPK_ITEMS:
        src = sources.get(word)
        if not src:
            data.issues.append(("첨부", "", "전년도 결재본에 '%s Cpk 계산 파일' 이 없어 만들지 못함" % word))
            continue
        vals = [_num((data.coa.get(l) or {}).get("924", {}).get(key)) for l in lots]
        vals = [v for v in vals if v is not None]
        name = re.sub(r"\.xls$", ".xlsx", os.path.basename(src))
        dst = os.path.join(folder, name)
        try:
            fill_cpk(src, dst, vals, today)
            out.append((name, dst))
        except Exception as error:
            data.issues.append(("첨부", name, "Cpk 파일을 만들지 못함: %s" % error))
    shutil.rmtree(work, ignore_errors=True)
    return out


def _find_stability_form(folder, input_dir):
    for base in (folder, os.path.join(input_dir or "", "서식"), input_dir or "",
                 os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")):
        if not base or not os.path.isdir(base):
            continue
        for n in os.listdir(base):
            if n.lower().endswith(".xlsx") and ("12606" in n.replace("-", "") or "126-06" in n) and "결과" not in n:
                return os.path.join(base, n)
    return None


COL = {"Initial": "C", "초기": "C", "3M": "D", "6M": "E", "9M": "F", "12M": "G", "18M": "H", "24M": "I",
       "36M": "J", "48M": "K", "60M": "L"}


def write_stability_workbook(folder, data, product, today, input_dir=None):
    """HLF-QC-126-06 (함량 시트) — 비전 판독의 시점별 함량으로 채운다. 없으면 None."""
    stab = getattr(data, "stability", None)
    points = (stab or {}).get("points") or {}
    form = _find_stability_form(folder, input_dir)
    if not form:
        data.issues.append(("첨부", "", "안정성 경향 분석 서식(HLF-QC-126-06)을 찾지 못해 만들지 못함 — 제품 폴더나 입력 폴더의 '서식' 폴더에 두세요"))
        return None
    name = "HLF-QC-126-06 안정성 시험 경향 분석 결과 - %s.xlsx" % (product.get("name") or "")
    dst = os.path.join(folder, name)
    wb = load_workbook(form)
    ws = wb["함량"] if "함량" in wb.sheetnames else wb.worksheets[0]
    ws["C3"] = product.get("name") or ""
    ws["G3"] = "함량 (%)"
    ws["C4"] = "25±2℃, 60±5%RH"
    ws["K4"] = today
    thin = Side(style="thin")
    diag = Border(diagonal=thin, diagonalDown=True, left=thin, right=thin, top=thin, bottom=thin)
    yellow = PatternFill("solid", fgColor="FFFF00")
    r = 38
    for lot, vals in points.items():
        ws["A%d" % r] = r - 37
        ws["B%d" % r] = lot
        for label, col in COL.items():
            if label in vals:
                ws["%s%d" % (col, r)] = _num(vals[label])
        r += 1
    if not points:
        ws["B38"] = "확인 필요"
        ws["B38"].fill = yellow
        ws["B38"].comment = Comment("시험일지 판독값이 없어 비워 둠 — 담당자 기입", "PQR")
    for sname in wb.sheetnames:
        if sname != ws.title:
            wb[sname]["C3"] = product.get("name") or ""
    wb.save(dst)
    return (name, dst)
