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

from . import convert
from . import lotcode
from . import stability_xlsx
from .readers import trend as trend_reader
from . import xls_fill

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


def _fill_cpk_xlsx(src_xls, dst_xlsx, values, today, cells=None):
    """Excel·LibreOffice 가 없을 때의 대비책 — .xlsx 로 바꿔 값만 채운다(그래프는 사라진다)."""
    tmp = dst_xlsx + ".tmp.xlsx"
    convert.to_xlsx(src_xls, tmp)
    wb = load_workbook(tmp)
    ws = wb.worksheets[0]
    ws["K4"] = today
    for ref, value in (cells or {}).items():
        ws[ref] = value
    for i in range(35):
        ws.cell(row=10 + i, column=2).value = values[i] if i < len(values) else None
    wb.save(dst_xlsx)
    os.remove(tmp)
    return dst_xlsx


def fill_cpk(src_xls, dst_xls, values, today):
    """B10~B44 에 결과값, K4 에 작성일. 서식의 수식·그래프를 그대로 둔 채 .xls 로 저장한다."""
    xls_fill.fill(src_xls, dst_xls, {"K4": today}, values)
    return dst_xls


def cpk_jobs(data, lots, product_name=""):
    """만들 Cpk 파일 목록 — [{"word", "label", "values", "cells"}].

    - 주성분이 둘이면 함량은 성분마다 한 파일 (디겐타: 플루오로메톨론 90~110, 겐타마이신 90~120).
      2026-09 점검: 첫 성분 값만 한 파일에 들어가고 둘째 성분 파일이 없었다.
    - 시트 머리(제품명 C4 · 공정 C5 · 시험항목 K5)와 규격(양쪽 N6/P6, 한쪽 P6)은 올해 값으로
      적는다 — 전년도 파일을 물려받으므로 제품명·규격이 그대로 남아 있었다.
    """
    def rec(lot):
        return (data.coa.get(lot) or {}).get("924") or {}
    jobs = []
    head = {"C4": ("%s (내수용)" % product_name).strip() if product_name else None, "C5": "포장 완료 후"}
    head = {k: v for k, v in head.items() if v}
    for word, key in CPK_ITEMS:
        if key == "assay":
            parts = []
            for lot in lots:
                for a in rec(lot).get("assays") or []:
                    if a.get("part") and a["part"] not in [p for p, _ in parts]:
                        parts.append((a["part"], (_num(a.get("lo")), _num(a.get("hi")))))
            if len(parts) > 1:
                for part, (lo, hi) in parts:
                    vals = []
                    for lot in lots:
                        v = next((_num(a.get("value")) for a in rec(lot).get("assays") or []
                                  if a.get("part") == part), None)
                        if v is not None:
                            vals.append(v)
                    cells = dict(head, K5="함량(%%) - %s" % part)
                    if lo is not None and hi is not None:
                        cells.update(N6=lo, P6=hi)
                    jobs.append({"word": word, "label": "함량(%s)" % part, "values": vals, "cells": cells})
                continue
            vals = [_num(rec(l).get("assay")) for l in lots]
            cells = dict(head, K5="함량(%)")
            spec = [(_num(a.get("lo")), _num(a.get("hi"))) for l in lots for a in rec(l).get("assays") or []]
            if spec and spec[0][0] is not None and spec[0][1] is not None:
                cells.update(N6=spec[0][0], P6=spec[0][1])
            jobs.append({"word": word, "label": "함량", "values": [v for v in vals if v is not None], "cells": cells})
            continue
        vals = [_num(rec(l).get(key)) for l in lots]
        cells = dict(head, K5=word if key != "particle" else "입자도(㎛)")
        if key == "particle":
            usl = next((_num(rec(l).get("particle_spec")) for l in lots if rec(l).get("particle_spec")), None)
            if usl is not None:
                cells.update(O6=None, P6=usl)
        jobs.append({"word": word, "label": word, "values": [v for v in vals if v is not None], "cells": cells})
    return jobs


def _decimals(text):
    """'6.40' → 2, '7.0' → 1, '7' → 0, 숫자가 없으면 None."""
    m = re.search(r"-?\d+(?:\.(\d+))?", str(text if text is not None else ""))
    if not m:
        return None
    return len(m.group(1)) if m.group(1) else 0


def _fmt_code(decimals):
    return "0" if not decimals else "0." + "0" * int(decimals)


def _formats(cells, decimals, spec_decimals):
    """xls_fill 에 넘길 표시 형식 {칸: '0.00'} — 규격 칸(N6/O6/P6)과 결과값 열('B')."""
    out = {}
    if spec_decimals:
        for ref in ("N6", "O6", "P6"):
            if cells.get(ref) not in (None, "") and _num(cells.get(ref)) is not None:
                out[ref] = _fmt_code(spec_decimals)
    if decimals:
        out["B"] = _fmt_code(decimals)
    return out


SHEET_NAME = re.compile(r"경향\s*분석\s*Sheet", re.I)
STAGE_KEYS = {"조제": ("921", "922"), "충전": ("923",), "포장": ("924",)}
STAGE_WORD = {"조제": "조제 완료 후", "충전": "충전 완료 후", "포장": "포장 완료 후"}


def _parse_sheet_name(name):
    """'16. HLF-QC-126-09 …Sheet(양쪽 규격 용)(Rev.000)_포장 함량(말레인산페니라민) - 복사본.xls'
    → {"stage": "포장", "item": "함량", "part": "말레인산페니라민", "sided": "Bilateral", "prefix": "16"} — 아니면 None."""
    base = os.path.basename(name)
    if not SHEET_NAME.search(base) or base.startswith("~$"):
        return None
    m = re.match(r"^\s*(\d+)\.\s*", base)
    prefix = m.group(1) if m else ""
    tail = re.sub(r"^.*\(Rev\.?\s*[\d.\-]+\)", "", base)
    tail = re.sub(r"\s*-\s*복사본\s*", "", tail)
    tail = re.sub(r"\.xlsx?$", "", tail, flags=re.I).strip(" _-")
    stage = next((w for w in STAGE_KEYS if tail.startswith(w)), "")
    rest = tail[len(stage):].strip(" _-") if stage else tail
    pm = re.match(r"^(.*?)\s*[(（]([^)）]+)[)）]\s*$", rest)
    item, part = (pm.group(1).strip(), pm.group(2).strip()) if pm else (rest.strip(), "")
    if not item:
        return None
    sided = "Unilateral" if "한쪽" in base else "Bilateral"
    return {"stage": stage or "포장", "item": item, "part": part, "sided": sided, "prefix": prefix,
            "path": name}


def _norm(text):
    return re.sub(r"[^0-9a-z가-힣]", "", str(text or "").lower())


def _spec_cells(spec, sided):
    """허용기준 글('6.0 ~ 8.0', '판정값 15.0% 이하', '10CFU/100mL 이하') → {N6, P6}."""
    out = {}
    m = re.search(r"(-?\d+(?:\.\d+)?)\s*~\s*(-?\d+(?:\.\d+)?)", spec or "")
    if m:
        out["N6"], out["P6"] = float(m.group(1)), float(m.group(2))
        return out
    # 한쪽 규격은 **첫 숫자**가 한계다 — '10CFU/100mL 이하' 에서 100 을 집으면 안 된다
    first = re.search(r"-?\d+(?:\.\d+)?", spec or "")
    if first and re.search(r"이하|미만", spec or ""):
        out["P6"] = float(first.group())
    elif first and re.search(r"이상|초과", spec or ""):
        out["N6"] = float(first.group())
    return out


def _value_of(rec, item, part):
    """성적서 판독 한 건에서 (값, 허용기준, 원문 글) — 못 찾으면 (None, '', '').

    원문 글('6.40')은 소수 자릿수를 성적서대로 보이려고 함께 돌려준다(담당자 2026-09-10).
    """
    want = _norm(item)
    if "함량" in item:
        for a in rec.get("assays") or []:
            if not part or _norm(a.get("part")) == _norm(part) or _norm(part) in _norm(a.get("part")):
                lo, hi = a.get("lo"), a.get("hi")
                return _num(a.get("value")), ("%s ~ %s" % (lo, hi) if lo and hi else ""), str(a.get("value") or "")
        return _num(rec.get("assay")), str(rec.get("assay_spec") or ""), str(rec.get("assay") or "")
    if "바이오버든" in item or "생균수" in item:
        got = rec.get("bioburden")
        spec = rec.get("bioburden_spec") or ""
        if got is None:
            for name, one in (rec.get("items") or {}).items():
                if "생균" in name or "바이오버든" in name:
                    got, spec = one.get("value"), one.get("spec") or ""
        return _num(got), str(spec), str(got or "")
    best = None
    for name, one in (rec.get("items") or {}).items():
        flat = _norm(re.sub(r"[(（][^)）]*[)）]", "", name))
        if flat == want or (len(want) >= 2 and (want in flat or flat in want)):
            if best is None or len(flat) < len(best[0]):
                best = (flat, one)
    if best is None:
        return None, "", ""
    value = str(best[1].get("value") or "")
    if part:
        m = re.search(re.escape(re.sub(r"\s+", "", part)) + r"\s*[:：]\s*([^,，;；]+)", re.sub(r"\s+", "", value))
        if m:
            value = m.group(1)
    return _num(value), str(best[1].get("spec") or ""), value


def cpk_jobs_from_sheets(input_dir, data, lots, product_name=""):
    """제품 폴더의 '16. …경향분석 Sheet…' 파일마다 Cpk 일감 하나 — [{"src", "name", "values", "cells", "sided"}].

    담당자 2026-09-09: "앞으로는 16번 첨부 파일 중 엑셀 파일(안정성, Cpk)로 너가 생성하는 PQR 작성본
    '안정성, Cpk 엑셀 파일' 을 생성해줘 — 단 Cpk 는 10 로트 이상일 때만". 전년도에 어떤 항목의 Cpk
    Sheet 를 냈는지가 그 제품의 Cpk 항목 목록이다(나조린: 조제 pH·비중·바이오버든, 충전 pH, 포장 pH·
    함량×2·제제균일성×2). 같은 항목의 '0. … - 복사본'(올해 빈 서식)이 있으면 그것을 서식으로 쓴다.
    """
    if not input_dir or not os.path.isdir(input_dir):
        return []
    found = {}
    for name in sorted(os.listdir(input_dir)):
        if not name.lower().endswith((".xls", ".xlsx")):
            continue
        info = _parse_sheet_name(os.path.join(input_dir, name))
        if not info or info["prefix"] not in ("0", "16"):
            continue
        key = (info["stage"], _norm(info["item"]), _norm(info["part"]))
        slot = found.setdefault(key, {})
        slot.setdefault(info["prefix"], info)
    jobs = []
    for key, slot in found.items():
        info = slot.get("0") or slot.get("16")
        src = info["path"]
        stage, item, part = info["stage"], info["item"], info["part"]
        values, specs, raws, missing = [], [], [], []
        for lot in lots:
            # 조제는 IPC(921)와 바이오버든(922) 성적서가 따로다 — 항목이 있는 쪽을 쓴다
            found_one = False
            for k in STAGE_KEYS[stage]:
                rec = (data.coa.get(lot) or {}).get(k) or {}
                if not rec:
                    continue
                v, sp, raw = _value_of(rec, item, part)
                if v is not None:
                    values.append(v)
                    raws.append(raw)
                    if sp:
                        specs.append(re.sub(r"\s+", "", sp))
                    found_one = True
                    break
            if not found_one:
                missing.append(lot)
        label = "%s %s%s" % (stage, item, ("(%s)" % part) if part else "")
        out_name = re.sub(r"^\s*(?:0|16)\.\s*", "", os.path.basename(src))
        out_name = re.sub(r"\s*-\s*복사본\s*", "", out_name)
        # 머리(C4 제품명·C5 공정·K5 항목)는 16번 서식에 적힌 그대로 둔다 — 회사가 쓰는 표기다.
        # 규격(N6/P6)만 올해 성적서의 허용기준으로 적는다. 연중에 기준이 바뀌었으면(나조린 조제 pH:
        # 변경관리로 6.2~7.5 → 6.2~7.0) **가장 최근 Lot 의 기준**을 쓰고 문의로 알린다.
        cells = {}
        distinct = []
        for sp in specs:
            if sp not in distinct:
                distinct.append(sp)
        if distinct:
            cells.update(_spec_cells(distinct[-1], info["sided"]))
        # 소수 자릿수는 성적서대로 — 결과값은 가장 자세한 값('6.40' → 2), 규격은 기준 글('6.2~7.0' → 1)
        # (담당자 2026-09-10: "소숫점 자리수는 시험성적서와 맞춰서 작성해줘 — 다른 시험항목도").
        decimals = max([_decimals(r) for r in raws if _decimals(r) is not None] or [0])
        spec_decimals = max([_decimals(x) for x in re.findall(r"-?\d+(?:\.\d+)?", distinct[-1])] or [0]) if distinct else 0
        jobs.append({"src": src, "name": out_name, "label": label, "values": values, "cells": cells,
                     "decimals": decimals, "spec_decimals": spec_decimals, "missing": missing,
                     # 뒤쪽 Lot 만 줄줄이 비면 도중에 생략된 시험이다(변경관리) — 성적서 확인을 묻지 않는다
                     "trailing_gap": bool(missing) and lots[len(lots) - len(missing):] == missing,
                     "sided": info["sided"], "prefix": info["prefix"],
                     "stage": stage, "item": item, "part": part,
                     "spec_note": (" → ".join(distinct)) if len(distinct) > 1 else ""})
    return jobs


FORMS = {"Bilateral": "HLF-QC-126-09 경향분석 Sheet(양쪽 규격).xlsx",
         "Unilateral": "HLF-QC-126-08 경향분석 Sheet(한쪽 규격).xlsx"}


def bundled_form(cells):
    """프로그램이 지닌 빈 Cpk 서식 — 규격이 위·아래 모두 있으면 양쪽 규격(‑09), 하나면 한쪽 규격(‑08)."""
    lo, hi = cells.get("N6"), cells.get("P6")
    both = all(v is not None and str(v).strip() and str(v).strip().upper() != "N/A" for v in (lo, hi))
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data",
                        FORMS["Bilateral" if both else "Unilateral"])
    return path if os.path.isfile(path) else None


def write_cpk_files(folder, data, previous_path, today, lots=None, product_name="", log=None):
    """내수용 Lot 의 완제 성적서 값으로 Cpk 파일들을 만든다. [(이름, 경로)]

    전년도 결재본의 Cpk 파일을 서식으로 물려받아(수식·그래프 보존) 값·머리·규격을 갈아 끼운다.
    그 길(Excel → LibreOffice → .xlsx 변환)이 모두 막히거나 전년도 파일이 없어도 파일은 반드시 나온다 —
    openpyxl 로 서식·수식·그래프를 그린다(cpk_xlsx). 담당자 2026-09-06: "10 Lot 이상이라서 Cpk 도
    작성이 되었어야 하는데 안 됐네" — Excel 연결이 터지자 파일이 아예 없었다.
    """
    log = log or (lambda *a: None)
    lots = lots or data.domestic
    from . import qc, cpk_xlsx, cpk_form
    if not qc.cpk_applies(len(lots)):
        # QC-126: 평가 년도 생산 Lot 이 기준(10) 미만이면 Cpk 를 산출하지 않는다 — 계산 파일도 만들지 않는다
        data.issues.append(("첨부", "", "평가 년도 생산 %d Lot 으로 %d Lot 미만 — QC-126 에 따라 Cpk 계산 파일을 만들지 않음"
                            % (len(lots), qc.cpk_min_lots())))
        return []
    work = tempfile.mkdtemp(prefix="pqr-cpk-")
    out = []
    # 제품 폴더에 전년도 경향분석 Sheet(16번)가 있으면 **그 항목 그대로, 그 서식 그대로** 만든다
    sheet_jobs = cpk_jobs_from_sheets(getattr(data, "folder", "") or "", data, lots, product_name)
    if sheet_jobs:
        log("  Cpk: 16번 경향분석 Sheet %d개를 서식으로 씁니다 — %s"
            % (len(sheet_jobs), ", ".join(j["label"] for j in sheet_jobs)))
        for job in sheet_jobs:
            if not job["values"]:
                data.issues.append(("첨부", job["name"], "'%s' 성적서 값이 없어 Cpk 계산 파일을 만들지 않음" % job["label"]))
                continue
            if len(job["values"]) < len(lots):
                if job.get("trailing_gap"):
                    # 나조린 조제 비중: LKYD01 부터 변경관리(CC-250822-03)로 생략 — 본문 9.2 가 각주·문의를
                    # 이미 냈다 (담당자 2026-09-10: "문제 없어. 변경관리에 따라 비중 시험 삭제")
                    log("  %s: '%s' 는 %s 부터 시험이 없어 %d Lot 으로 만듭니다 (도중 생략 — 9.2 각주 참고)"
                        % (job["name"], job["label"], job["missing"][0], len(job["values"])))
                else:
                    data.issues.append(("첨부", job["name"], "'%s' 값이 %d/%d Lot 에만 있습니다(없는 Lot: %s) — 성적서를 확인하세요"
                                        % (job["label"], len(job["values"]), len(lots), ", ".join(job["missing"]))))
            if job.get("spec_note"):
                # 연중에 기준이 바뀌면 가장 최근 기준 — 담당자 확인 완료(2026-09-10: "가장 최근 기준으로 작성하면 돼").
                # 본문 9.1 이 기준 변경 각주와 문의를 따로 내므로 여기서는 기록만 남긴다.
                log("  %s: '%s' 허용기준이 Lot 에 따라 다름(%s) — 가장 최근 기준으로 규격 칸을 적음"
                    % (job["name"], job["label"], job["spec_note"]))
            cells = dict(job["cells"], K4=today)
            dst = free_path(os.path.join(folder, job["name"]), data, log)
            name = os.path.basename(dst)
            try:
                xls_fill.fill(job["src"], dst, cells, job["values"],
                              formats=_formats(cells, job.get("decimals"), job.get("spec_decimals")))
                out.append((name, dst))
                log("  %s: %s번 서식을 그대로 채움 (%d Lot)" % (name, job["prefix"], len(job["values"])))
                continue
            except Exception as error:
                why = str(error)
            from . import cpk_form
            dst2 = re.sub(r"\.xls$", ".xlsx", dst)
            try:
                converted = os.path.join(work, "form-%d.xlsx" % len(out))
                convert.to_xlsx(job["src"], converted)
                cpk_form.fill(converted, dst2, cells, job["values"], today,
                              decimals=job.get("decimals"), spec_decimals=job.get("spec_decimals"))
                out.append((os.path.basename(dst2), dst2))
                log("  %s: 서식을 .xlsx 로 바꿔 채움 — %s" % (os.path.basename(dst2), why))
            except Exception as error:
                data.issues.append(("첨부", job["name"], "Cpk 파일을 만들지 못함: %s / %s" % (why, error)))
        shutil.rmtree(work, ignore_errors=True)
        return out
    try:
        sources = _previous_xls(previous_path, work)
    except Exception as error:
        sources = {}
        log("  전년도 Cpk 파일을 꺼내지 못함: %s" % error)
    for job in cpk_jobs(data, lots, product_name):
        word = job["word"]
        src = sources.get(word)
        vals = job["values"]
        if not vals:
            data.issues.append(("첨부", "", "'%s' 완제 성적서 값이 없어 Cpk 계산 파일을 만들지 않음" % job["label"]))
            continue
        # 이름은 늘 같은 꼴 — 전년도 파일 이름('16. 전년도 PQR25함량 Cpk 계산 파일.xls')을 그대로 물려받지 않는다
        name = "a. %s Cpk 계산 파일.xls" % job["label"]
        cells = dict(job["cells"], K4=today)
        why = []
        if src:
            dst = free_path(os.path.join(folder, name), data, log)
            name = os.path.basename(dst)
            try:                                  # Excel·LibreOffice 가 있으면 .xls 서식을 그대로 채운다
                xls_fill.fill(src, dst, cells, vals)
                out.append((name, dst))
                log("  %s: 전년도 .xls 서식을 그대로 채움" % name)
                continue
            except Exception as error:            # FillError 든 COM 오류든 — 다음 길로
                why.append(str(error))
        else:
            why.append("전년도 결재본(16항 압축)에 '%s Cpk 계산 파일' 이 없음" % word)
        name = re.sub(r"\.xls$", ".xlsx", name)
        dst = free_path(os.path.join(folder, name), data, log)
        name = os.path.basename(dst)
        # 서식의 칸 값만 갈아 끼운다 — 그래프·로고·수식이 그대로 남는다(담당자 2026-09-06:
        # "Cpk 는 전년도 양식으로 작성하되 2026년 PQR 작성본 내용을 참고해서 업데이트하면 돼").
        # 전년도 파일을 .xlsx 로 바꿀 수 있으면 그것을, 아니면 프로그램이 지닌 빈 서식을 쓴다.
        forms = []
        if src:
            try:
                converted = os.path.join(work, "form-%s.xlsx" % re.sub(r"\W+", "", word))
                convert.to_xlsx(src, converted)
                forms.append(("전년도 서식", converted))
            except Exception as error:
                why.append(str(error))
        blank = bundled_form(cells)
        if blank:
            forms.append(("프로그램이 지닌 빈 서식", blank))
        for label, form in forms:
            try:
                kind = cpk_form.fill(form, dst, cells, vals, today)
                out.append((name, dst))
                log("  %s: %s(%s)을 채움 — %s" % (name, label, kind, "; ".join(why)))
                if label != "전년도 서식":
                    data.issues.append(("첨부", name, "전년도 Cpk 파일을 쓰지 못해 프로그램이 지닌 같은 서식(HLF-QC-126-%s)에 "
                                                     "올해 값을 채웠습니다 — 한 번 확인하세요 (%s)"
                                        % ("09" if kind == "Bilateral" else "08", "; ".join(why))))
                break
            except Exception as error:
                why.append("%s: %s" % (label, error))
        else:
            try:                                  # 마지막 길 — 프로그램이 서식·수식·그래프를 직접 그린다
                kind = cpk_xlsx.build(dst, vals, cells, today)
                out.append((name, dst))
                log("  %s: 서식을 직접 그림(%s) — %s" % (name, kind, "; ".join(why)))
                data.issues.append(("첨부", name, "전년도 Cpk 서식을 쓰지 못해 프로그램이 서식(HLF-QC-126-%s)을 직접 그렸습니다 — "
                                                  "값·Cpk·그래프는 같으나 모양을 한 번 확인하세요 (%s)"
                                    % ("09" if kind == "Bilateral" else "08", "; ".join(why))))
            except Exception as error:
                data.issues.append(("첨부", name, "Cpk 파일을 만들지 못함: %s / %s" % ("; ".join(why), error)))
                log("  %s: 만들지 못함 — %s / %s" % (name, "; ".join(why), error))
    shutil.rmtree(work, ignore_errors=True)
    return out


def _is_blank_form(path):
    """제품명 칸(C3)이 비어 있으면 아직 채우지 않은 서식이다.

    담당자가 올리는 서식 파일 이름에도 '결과'가 들어 있어(‘HLF-QC-126-06 안정성 시험 경향
    분석 결과 (Rev.001).xlsx’) 이름만으로는 채운 파일과 가릴 수 없다 — 안을 보고 가린다.
    """
    try:
        book = load_workbook(path, data_only=True, read_only=True)
        try:
            return not str(book.worksheets[0]["C3"].value or "").strip()
        finally:
            book.close()
    except Exception:
        return False


def _stability_form_candidates(folder, input_dir, product_dir=None):
    """빈 경향 분석 서식(HLF-QC-126-06) 후보 — 제품 폴더 → 작성 폴더 → 입력 폴더 '서식' → 입력 폴더 → 프로그램 서식."""
    out = []
    for base in (product_dir, folder, os.path.join(input_dir or "", "서식"), input_dir or "",
                 os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")):
        if not base or not os.path.isdir(base):
            continue
        found = [os.path.join(base, n) for n in sorted(os.listdir(base))
                 if n.lower().endswith(".xlsx") and not n.startswith("~$")
                 and ("12606" in n.replace("-", "") or "126-06" in n)]
        out += [p for p in found if _is_blank_form(p) and p not in out]
    return out


def _find_stability_form(folder, input_dir, product_dir=None):
    found = _stability_form_candidates(folder, input_dir, product_dir)
    return found[0] if found else None


def _form_for_parts(form, parts, folder, data):
    """서식에 시험항목 수만큼 시트가 없으면(제품 폴더 서식이 pH·함량·보존제 세 장인데 삼투압까지 네 항목)
    시트가 더 많은 다른 서식을 고른다 — 모자라면 뒤 항목(보존제)이 통째로 빠지고 삼투압이 함량 시트에 실렸다
    (올로원스 2026-09-10). 후보는 _stability_form_candidates 차례, 같은 점수면 앞엣것."""
    need = [(p, it) for p, it, _lo, _hi, _ol in parts]

    def score(path):
        try:
            names = stability_xlsx.sheet_names(path)
            picks = pick_form_sheets(need, names)
        except Exception:
            return -1
        # 항목마다 시트를 받았는가 + pH·삼투압·보존제가 제 이름 시트를 받았는가
        got = len(picks)
        same = sum(1 for (sheet, label) in picks if label in ("pH", "삼투압", "보존제") and sheet == label)
        return got * 10 + same

    best = form
    try:
        cur = score(form)
        if cur < 0 or cur >= len(need) * 10 + sum(1 for p, it in need if p in ("pH", "삼투압", "보존제")):
            return form                              # 못 여는 서식(시험의 가짜 경로)이거나 이미 넉넉하다
        for cand in _stability_form_candidates(folder, getattr(data, "input_dir", None) or "",
                                               getattr(data, "folder", None) or ""):
            if cand != best and score(cand) > score(best):
                best = cand
    except Exception:
        return form
    if best != form:
        try:
            note = ("첨부 경향표: 서식 '%s' 는 시트가 모자라(%s) '%s' 서식으로 만들었습니다"
                    % (os.path.basename(form), ", ".join(stability_xlsx.sheet_names(form)), os.path.basename(best)))
            data.issues.append(("첨부", "", note))
        except Exception:
            pass
    return best


COL = {"Initial": "C", "초기": "C", "3M": "D", "6M": "E", "9M": "F", "12M": "G", "18M": "H", "24M": "I",
       "36M": "J", "48M": "K", "60M": "L"}


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def written_by(report_path):
    """보고서 '검토 및 승인' 표에서 작 성 (Written by) 의 성 명을 읽는다.

    첨부 엑셀의 작성자를 본문과 같게 맞추기 위한 것이다. 못 찾으면 빈 문자열.
    """
    if not report_path or not os.path.isfile(report_path):
        return ""
    try:
        from lxml import etree
        with zipfile.ZipFile(report_path) as z:
            root = etree.fromstring(z.read("word/document.xml"))
    except Exception:
        return ""

    def text(el):
        return "".join(t.text or "" for t in el.iter(W + "t"))

    for tbl in root.iter(W + "tbl"):
        rows = tbl.findall(W + "tr")
        for i, tr in enumerate(rows):
            if "Written by" not in text(tr):
                continue
            for nxt in rows[i + 1:]:
                cells = nxt.findall(W + "tc")
                if len(cells) < 2:
                    continue
                name = re.sub(r"\s+", "", text(cells[1]))
                if name:
                    return name
            return ""
    return ""


def free_path(dst, data=None, log=None):
    """덮어쓸 수 없는 파일이면 다른 이름을 돌려준다 — 엑셀에서 열어 둔 채 재작성한 경우.

    담당자 2026-09-07: 워드는 새 판독으로 다시 쓰였는데 경향 엑셀에는 새 Lot(OEY301·302)이 없었다.
    파일이 엑셀에 열려 있으면 덮어쓰지 못하는데, 그때 조용히 지나가면 예전 파일이 그대로 남아
    본문과 첨부가 어긋난다. 그래서 다른 이름으로 만들고 문의 목록에 알린다.
    """
    if not os.path.exists(dst):
        return dst
    try:
        with open(dst, "r+b"):
            return dst
    except OSError:
        stem, ext = os.path.splitext(dst)
        alt = "%s (새로 만든 것)%s" % (stem, ext)
        why = ("이 파일이 엑셀에서 열려 있어 덮어쓰지 못했습니다 — '%s' 로 만들었습니다. "
               "엑셀을 닫고 다시 작성하면 원래 이름으로 만들어집니다." % os.path.basename(alt))
        if data is not None:
            data.issues.insert(0, ("첨부", os.path.basename(dst), "★ " + why))
        if log:
            log("  ★ %s — %s" % (os.path.basename(dst), why))
        return alt


def _grouped(points):
    """{Lot: {시점: 값}} 또는 {포장구분: {Lot: {시점: 값}}} 을 [(구분, [(Lot, 값)])] 로 편다."""
    if not points:
        return []
    inner = list(points.values())
    if inner and isinstance(inner[0], dict) and inner[0] and \
            all(isinstance(v, dict) for v in inner[0].values()):
        return [(k, list(v.items())) for k, v in points.items()]
    return [("", list(points.items()))]


def write_stability_workbook(folder, data, product, today, input_dir=None, report_path=None,
                             product_dir=None):
    """HLF-QC-126-06 — 시점별 함량을 채운다. 포장 규격이 나뉘어 있으면 파일도 나눈다.

    작성자(M3)·작성일(M4)은 보고서 본문의 작성자·작성일자와 같게 넣는다.
    반환값은 [(파일명, 경로), ...] 이며 만들지 못하면 빈 목록이다.
    """
    stab = getattr(data, "stability", None)
    points = (stab or {}).get("points") or {}
    logs = getattr(data, "stability_logs", None)
    form = _find_stability_form(folder, input_dir, product_dir)
    if not form:
        data.issues.append(("첨부", "", "안정성 경향 분석 서식(HLF-QC-126-06)을 찾지 못해 만들지 못함 — 제품 폴더나 입력 폴더의 '서식' 폴더에 두세요"))
        return []
    if logs or getattr(data, "stability_trend", None):
        return _write_trend(form, folder, data, product, today, report_path)
    if not points:
        data.issues.append(("첨부", "", "안정성 시험일지 판독값이 없어 경향 분석 파일을 만들지 못함 — 시험일지를 올리거나 담당자가 직접 기입"))
        return []
    base = product.get("name") or ""
    author = written_by(report_path)
    made = []
    for label, lots in _grouped(points):
        title = base if not label or label in base else "%s(%s)" % (base, label)
        name = "HLF-QC-126-06 안정성 시험 경향 분석 결과 - %s.xlsx" % title
        dst = os.path.join(folder, name)
        lots = [(lot, {k: _num(v) for k, v in vals.items()}) for lot, vals in lots]
        stability_xlsx.build(form, dst, title, lots,
                             lcl=(stab or {}).get("lcl", 90),
                             ucl=(stab or {}).get("ucl", 110),
                             prepared_by=author, prepared_on=today)
        made.append((name, dst))
    return made


FORM_SHEETS = ("함량", "A", "B", "기타", "총", "pH", "삼투압")


def pick_form_sheets(parts, names):
    """시험항목마다 (서식 시트, 새 시트 이름) — [(form_sheet, name)].

    담당자 PC 2026-09-07: pH 값을 '함량' 시트에 넣어 시트 이름이 '함량(pH)' 가 되고 축이 90~110 이라
    선이 안 보였다. 이름이 같은 시트(pH·보존제·삼투압)가 서식에 있으면 그것을, 함량 성분은 '함량'
    부터 차례로 쓴다. 한 시트를 두 번 쓰지 않는다.
    """
    names = [n for n in (names or []) if n]
    order = [n for n in FORM_SHEETS if n in names] + [n for n in names if n not in FORM_SHEETS]
    if not order:
        order = list(FORM_SHEETS)
    used, out = set(), []
    for part, item in parts:
        text = "%s %s" % (part or "", item or "")
        kind = None
        for word in ("pH", "보존제", "삼투압"):
            if word.lower() in text.lower():
                kind = word
                break
        pick = None
        if kind and kind in order and kind not in used:
            pick = kind
        if pick is None:
            for n in order:
                if n not in used and n not in ("pH", "보존제", "삼투압"):
                    pick = n
                    break
        if pick is None:
            pick = next((n for n in order if n not in used), None)
        if pick is None:
            break
        used.add(pick)
        label = kind if kind else ("함량" if part == "함량" else "함량(%s)" % part)
        if kind == "보존제" and part and part != "보존제":
            label = "보존제(%s)" % part
        out.append((pick, label))
    return out


def _assay_limits(data):
    """{성분: (하한, 상한)} — 완제 성적서의 함량 규격. 제품마다 다르므로 여기서 읽는다."""
    out = {}
    for lot in (data.coa or {}).values():
        rec = lot.get("924") or {}
        for a in rec.get("assays") or []:
            part = (a.get("part") or "").strip()
            if part and part not in out:
                lo, hi = _num(a.get("lo")), _num(a.get("hi"))
                if lo is not None and hi is not None:
                    out[part] = (lo, hi)
        # pH·삼투압·보존제도 경향표 항목이다 — 완제 성적서 시험항목 표의 허용기준('6.0 ~ 8.0')에서
        for name, one in (rec.get("items") or {}).items():
            plain = re.sub(r"[(（][^)）]*[)）]", "", name).strip()
            if plain in ("pH", "삼투압", "보존제") and plain not in out:
                m = re.search(r"(-?\d+(?:\.\d+)?)\s*~\s*(-?\d+(?:\.\d+)?)", str(one.get("spec") or ""))
                if m:
                    out[plain] = (float(m.group(1)), float(m.group(2)))
    return out


def points_by_part(one, part, year_to=None, shaky=None):
    """시험일지 판독 한 건에서 그 성분의 {시점: 값} — 평가 기간을 넘어선 시점은 뺀다.

    판독값의 성분 이름이 경향표와 다를 수 있다 — '트레할로스수화물' 과 '함량'
    (담당자 2026-09-08: 판독 3 Lot 이 있는데 '안정성 시험 결과값이 없어' 경향 파일을 못 만들었다).
    같은 계열 이름을 먼저 찾고, 성분이 하나뿐이면 이름이 달라도 그 값을 쓴다.
    """
    out = {}
    for point in one.get("points", []):
        got = re.findall(r"\d{4}", point.get("done") or "") or \
            re.findall(r"\d{4}", point.get("expected") or "")     # 일자를 못 읽었으면 예정 시기로
        if year_to and got and int(got[0]) > year_to:
            continue
        assays = point.get("assays") or {}
        value = _num(assays.get(part))
        if value is None and assays:
            같은것 = [v for k, v in assays.items() if _same_part(k, part)]
            if not 같은것 and len(assays) == 1:
                같은것 = list(assays.values())
            value = _num(같은것[0]) if 같은것 else None
        if shaky is not None and part in (point.get("unsure") or []):
            shaky.add(point["period"])         # 손글씨 판독이 애매한 시점 — 시트에서 주황
        if value is not None:
            out[point["period"]] = float(value)
    return out


def _same_part(a, b):
    a, b = re.sub(r"\s+", "", a or ""), re.sub(r"\s+", "", b or "")
    return bool(a) and bool(b) and (a in b or b in a)


def _write_trend(form, folder, data, product, today, report_path):
    """안정성 경향 분석(HLF-QC-126-06) — 내수용·수출용 시험일지가 섞여 있으면 시장마다 한 파일(2026-09,
    퀴노비드: '… - 퀴노비드안연고(내수용).xlsx' / '(수출용).xlsx'). 지난 경향표 시트는 제품명에 '수출' 이
    들어 있으면 수출용 파일로, 아니면 내수용 파일로 간다."""
    seed = list(getattr(data, "stability_trend", None) or [])
    logs = list(getattr(data, "stability_logs", None) or [])
    markets = []
    for one in logs:
        m = one.get("market") or "내수"
        if m not in markets:
            markets.append(m)
    if len(markets) <= 1:
        return _trend_file(form, folder, data, product, today, report_path, seed, logs, "")
    files = []
    for market in sorted(markets, key=lambda m: m != "내수"):
        mseed = [sh for sh in seed if ("수출" in (sh.get("product") or "")) == (market == "수출")]
        mlogs = [one for one in logs if (one.get("market") or "내수") == market]
        files += _trend_file(form, folder, data, product, today, report_path, mseed, mlogs, "(%s용)" % market)
    return files


def _trend_file(form, folder, data, product, today, report_path, seed, logs, suffix):
    """안정성 경향 분석(HLF-QC-126-06)을 새로 만든다 — 지난 경향표 + 올해 시험일지.

    담당자 지시(2026-09): "16. 안정성시험 경향표를 13항 안정성 시험 자료를 참고해서 최신
    파일로 신규 작성." 지난 경향표의 값은 담당자가 옮겨 적어 둔 것이라 그대로 이어받고,
    올해 시험일지에서 읽은 시점만 덧붙인다. 평가 기간을 넘어선 시점(다음 해에 끝난 시험)은
    넣지 않는다 — 보고서 13.3 의 경향 범위와 같은 값이어야 한 벌의 자료로 읽힌다.
    """
    year_to = lotcode.year_of((getattr(data, "period", None) or {}).get("to"))
    limits = _assay_limits(data)

    def points_of(one, part, shaky=None):
        return points_by_part(one, part, year_to, shaky)

    parts = []                                    # [(성분, 시험항목 글, 하한, 상한, 지난 Lot)]
    # 13항은 **올해 시험일지만으로** 적는다 — 지난 경향표(seed)의 Lot·값은 잇지 않고, 시트 서식만 빌린다
    # (담당자 2026-09-09: "안정성을 이전 PQR 에서 가져오지 말고 첨부된 자료로만 빈칸 기재해줘";
    #  나조린의 지난 경향표는 다른 제품 시트(HNV701·함량(총 유연물질))가 섞여 있어 pH 만 채워지고
    #  함량 시트는 엉뚱한 것이 남았다). 항목은 시험일지 판독의 숫자 값 가운데 pH·삼투압·보존제와
    #  완제 성적서에 규격이 있는 함량 성분만 — 제제균일성·불용성미립자는 경향표 항목이 아니다.
    keys = []
    for one in logs:
        for pt in one.get("points") or []:
            for k, v in (pt.get("assays") or {}).items():
                if _num(v) is not None and k not in keys:
                    keys.append(k)
    main_parts = [k for k in limits if k not in ("pH", "삼투압", "보존제")
                  and not any(w in k for w in ("벤잘코늄", "염화벤잘코늄", "보존제"))]
    for key in keys:
        kind = key if key in ("pH", "삼투압", "보존제") else None
        # 함량 성분은 이름이 **같아야** 한다 — '제제균일성 말레인산페니라민' 을 함량으로 잘못 잡지 않게
        flat = re.sub(r"[\s()（）·∙]", "", key)
        lim = limits.get(next((k for k in limits if re.sub(r"[\s()（）·∙]", "", k) == flat), ""))
        title = None
        if key == "함량" and lim is None:
            # 주성분이 하나인 제품의 판독 열쇠 '함량'(collect.normalize_stability_keys) — 그 성분의 규격,
            # 시트 이름은 지난 경향표대로 '함량' (퀴노비드 2026-09-10: 함량 시트가 통째로 빠졌다)
            if len(main_parts) == 1:
                lim = limits.get(main_parts[0])
            title = "함량"
        if kind is None and lim is None:
            continue
        lo, hi = lim or (None, None)
        if (kind or title) and lim is None:
            sh = next((x for x in seed if trend_reader.component_of(x.get("item")) == key), None)
            lo, hi = (sh.get("lcl"), sh.get("ucl")) if sh else (None, None)
        parts.append((key, title or (key if kind else "함량 - %s(%%)" % key), lo, hi, []))
    if not parts:                                 # 지난 경향표가 없는 첫해 — 성적서·시험일지에서
        names = list(limits)
        if not names:
            names = [part for one in logs for point in one.get("points", [])
                     for part in (point.get("assays") or {})]
            names = list(dict.fromkeys(names))
        for part in names:
            lo, hi = limits.get(part, (90, 110))
            parts.append((part, "함량 - %s(%%)" % part, lo, hi, []))
    if not parts:
        data.issues.append(("첨부", "", "안정성 경향 분석에 넣을 성분을 찾지 못해 만들지 못함 — "
                                        "지난 경향표(HLF-QC-126-06)나 시험일지 판독값이 필요합니다"))
        return []

    sheets, added = [], 0
    form = _form_for_parts(form, parts, folder, data)
    try:
        picks = pick_form_sheets([(p, it) for p, it, _lo, _hi, _ol in parts], stability_xlsx.sheet_names(form))
    except Exception:
        picks = [(FORM_SHEETS[i], "함량(%s)" % p) for i, (p, _it, _lo, _hi, _ol) in enumerate(parts[:len(FORM_SHEETS)])]
    for i, (part, item, lo, hi, old_lots) in enumerate(parts[:len(picks)]):
        rows = []
        seen, unsure = {}, {}
        for lot, values in old_lots:              # 지난 경향표의 차례를 지킨다
            seen[lot] = dict(values)
            rows.append(lot)
        for one in logs:
            shaky = set()
            values = points_of(one, part, shaky)
            if not values and not shaky:
                continue
            lot = one["lot"]
            if lot not in seen:
                seen[lot] = {}
                rows.append(lot)
            # 지난 경향표(담당자가 옮겨 적은 값)가 우선 — 판독값은 거기 없는 시점만 더한다
            known_before = set(seen[lot])
            for period, value in values.items():
                if period not in seen[lot]:
                    seen[lot][period] = value
                    added += 1
            # 경향표에 있던 시점만 애매에서 뺀다 — 방금 넣은 예상값은 주황으로 남아야 한다
            # (담당자 2026-09-06 미리 보기: 예상값 칸이 주황이 아니었다)
            shaky = set(shaky) - known_before
            if shaky:
                unsure[lot] = shaky
        sheets.append({"form_sheet": picks[i][0], "name": picks[i][1], "item": item,
                       "lots": [(lot, seen[lot]) for lot in rows], "unsure": unsure,
                       "lcl": float(lo if lo is not None else 90),
                       "ucl": float(hi if hi is not None else 110)})
    if not any(sheet["lots"] for sheet in sheets):
        data.issues.append(("첨부", "", "안정성 시험 결과값이 없어 경향 분석 파일을 만들지 못함 — "
                                        "13항에 시험일지 판독값(.json)이나 지난 경향표를 두세요"))
        return []
    if not added:
        data.issues.append(("첨부", "", "13항 시험일지에서 새 시점을 읽지 못해 지난 경향표 값만 "
                                        "옮겼습니다 — 올해 시점을 직접 채우세요"))

    name = "HLF-QC-126-06 안정성 시험 경향 분석 결과 - %s%s.xlsx" % (product.get("name") or "", suffix)
    dst = free_path(os.path.join(folder, name), data)
    name = os.path.basename(dst)
    stability_xlsx.build_multi(form, dst, (product.get("name") or "") + suffix, sheets,
                               prepared_by=written_by(report_path), prepared_on=today)
    return [(name, dst)]
