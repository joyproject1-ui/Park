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
    try:
        sources = _previous_xls(previous_path, work)
    except Exception as error:
        sources = {}
        log("  전년도 Cpk 파일을 꺼내지 못함: %s" % error)
    out = []
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


def _find_stability_form(folder, input_dir, product_dir=None):
    for base in (product_dir, folder, os.path.join(input_dir or "", "서식"), input_dir or "",
                 os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")):
        if not base or not os.path.isdir(base):
            continue
        found = [os.path.join(base, n) for n in sorted(os.listdir(base))
                 if n.lower().endswith(".xlsx") and not n.startswith("~$")
                 and ("12606" in n.replace("-", "") or "126-06" in n)]
        blank = [p for p in found if _is_blank_form(p)]
        if blank:
            return blank[0]
    return None


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
        label = kind if kind else "함량(%s)" % part
        if kind == "보존제" and part and part != "보존제":
            label = "보존제(%s)" % part
        out.append((pick, label))
    return out


def _assay_limits(data):
    """{성분: (하한, 상한)} — 완제 성적서의 함량 규격. 제품마다 다르므로 여기서 읽는다."""
    out = {}
    for lot in (data.coa or {}).values():
        for a in (lot.get("924") or {}).get("assays") or []:
            part = (a.get("part") or "").strip()
            if part and part not in out:
                lo, hi = _num(a.get("lo")), _num(a.get("hi"))
                if lo is not None and hi is not None:
                    out[part] = (lo, hi)
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
    for sheet in seed:
        part = trend_reader.component_of(sheet.get("item"))
        lo, hi = limits.get(next((k for k in limits if _same_part(k, part)), ""), (None, None))
        parts.append((part, sheet.get("item") or ("함량 - %s(%%)" % part),
                      lo if lo is not None else sheet.get("lcl"),
                      hi if hi is not None else sheet.get("ucl"),
                      list(sheet.get("lots") or [])))
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
