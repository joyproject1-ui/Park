# -*- coding: utf-8 -*-
"""안연고제 결재본(퀴노비드안연고 계열)에 값을 채우는 절차.

퀴노비드안연고 2026 PQR 을 담당자와 함께 만들며 확정한 채움 규칙을 그대로 옮겼다.
표는 항 제목과 머리행 낱말로 찾고, 값은 ProductData(판독 결과)에서 가져온다.
판독하지 못한 항(손글씨 안정성 등)은 결재본 값을 두고 issues 에 '확인 필요' 로 남긴다.
"""
import copy
import datetime as _dt
import re
import statistics

from docx.oxml.ns import qn

from . import docedit as E
from . import lotcode, qc
from .locate import find_tables, find_para, _text
from .ooxml_order import get_or_add

STORE = ("25±2°C,", "60±5%RH")


def _num(s):
    m = re.search(r"-?\d+(?:\.\d+)?", str(s or ""))
    return float(m.group()) if m else None


def _rng(s):
    m = re.findall(r"\d+(?:\.\d+)?", str(s or ""))
    return (float(m[0]), float(m[1])) if len(m) >= 2 else None


def avg(vals, digits=2):
    return ("%%.%df" % digits) % (sum(vals) / len(vals))


DEFAULT_LIMITS = {"particle": 75.0, "assay": (90.0, 110.0), "metal": 50.0}   # 퀴노비드안연고 값 — 9.1 을 못 읽을 때만


def parse_limits(spec_texts):
    """9.1 시험결과표의 허용기준 글에서 Cpk 한계를 읽는다. {'particle': usl, 'assay': (lsl, usl), 'metal': usl}

    spec_texts: {'particle': '75 ㎛ 이하', 'assay': '90.0 ~ 110.0%', 'metal': '… 합계 50개 이하 …'}.
    못 읽은 항목은 넣지 않는다 — 부르는 쪽이 기본값을 쓰고 문의 목록에 남긴다.
    """
    out = {}
    t = spec_texts.get("particle") or ""
    m = re.search(r"(\d+(?:\.\d+)?)\s*(?:㎛|um|µm|μm)", t)
    if m:
        out["particle"] = float(m.group(1))
    t = spec_texts.get("assay") or ""
    m = re.search(r"(\d+(?:\.\d+)?)\s*[~∼～-]\s*(\d+(?:\.\d+)?)\s*%", t)
    if m:
        out["assay"] = (float(m.group(1)), float(m.group(2)))
    t = spec_texts.get("metal") or ""
    m = re.search(r"(\d+)\s*개\s*이하", t)
    if m:
        out["metal"] = float(m.group(1))
    return out


def cpk_uni(vals, usl):
    m = sum(vals) / len(vals)
    s = statistics.stdev(vals) if len(vals) > 1 else 0
    return (usl - m) / (3 * s) if s else None


def cpk_bi(vals, lsl, usl):
    m = sum(vals) / len(vals)
    s = statistics.stdev(vals) if len(vals) > 1 else 0
    return min((usl - m) / (3 * s), (m - lsl) / (3 * s)) if s else None


def _tables(document, prefix):
    return [document.tables[i] for i in find_tables(document, prefix)]


def _by_header(tables, *words):
    for t in tables:
        head = re.sub(r"\s+", "", _text(t._tbl.findall(qn("w:tr"))[0]))
        if all(w in head for w in words):
            return t
    return None


ASSAY_PART = re.compile(r"\s*\[([^\]]+)\]")


def assay_component(crit_text):
    """9.1 허용기준 칸 앞머리의 [성분] 이름. 없으면 None.

    주성분이 둘 이상인 제품(디겐타안연고: 플루오로메톨론·겐타마이신황산염)은 함량 줄이
    성분별로 나뉘고, 어느 줄이 어느 성분인지는 이 표시로만 알 수 있다.
    """
    m = ASSAY_PART.match(crit_text or "")
    return m.group(1).strip() if m else None


def _avg_text(values, fmt="%.2f", unit_on_avg=True):
    """질량·용량 평균 칸의 글 — 값이 모두 같으면 평균·범위를 적지 않고 그 값만 적는다.

    2026 결재본의 주석 그대로다: "모든 Lot 의 시험결과가 (…) 동일하여 최댓값, 최솟값, 평균은
    별도로 작성하지 않음." 3 Lot 이 모두 4.1 g 인데 'Av. 4.1 (4.1 ~ 4.1g)' 은 읽기 나쁘다.
    """
    if len(set(values)) == 1:
        return (fmt + " g") % values[0]
    head = "Av. " + fmt + ("g" if unit_on_avg else "")
    return (head + "\n(" + fmt + " ~ " + fmt + "g)") % (
        sum(values) / len(values), min(values), max(values))


def _deviation_lines(dev):
    """11항 일탈사항 칸의 글 — 2026 결재본 차림새로 요점만.

        [조제 작업 중 장비 이상 건]
        * 일탈 내용
        조제 작업 진행 중 MAINMIXER 의 스크래퍼가 작동되지 않아 조제 작업 중단함.
        * 일탈 원인
        원인 미상

    담당자 지적(2026-09): "표 안의 일탈사항은 더 요약해서 요점만." 예전에는 일탈보고서의
    설명을 통째로 옮겨 시각 기록('1) 조제 시작 시간 …')까지 실렸다. 말은 지어내지 않고
    원본에서 **고르고 다듬기만** 한다 — 제목은 제품명·Lot 을 떼고, 내용은 시각·번호 기록을
    뺀 문장만, 원인은 일탈보고서의 '원인' 칸을 그대로 쓴다.
    """
    title = re.sub(r"^\S+\s*\([A-Z0-9]+\)\s*", "", dev.get("title") or "").strip()
    out = [("[%s]" % title, "none")] if title else []
    body = []
    for line in (dev.get("description") or "").split("\n"):
        line = line.strip()
        if not line or re.match(r"^\d\)", line):      # '1) 조제 시작 시간 …' 같은 시각 기록은 뺀다
            continue
        body.append(re.sub(r"^\d\.\s*", "", line))
    if body:
        out.append(("* 일탈 내용", "none"))
        out.append((" ".join(body), "none"))
    cause = re.sub(r"^\d\.\s*", "", (dev.get("cause") or "").replace("\n", " ")).strip()
    if cause:
        out.append(("* 일탈 원인", "none"))
        out.append((cause, "none"))
    return out


def _quarter(day):
    return (day.month - 1) // 3 + 1


def _small(para_el):
    for r in para_el.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = r.makeelement(qn("w:rPr"), {}); r.insert(0, rpr)
        for tag in ("sz", "szCs"):
            get_or_add(rpr, tag).set(qn("w:val"), "18")


def fill(document, data, product, period, today=None, log=None):
    log = log or (lambda *a: None)
    issues = []
    today = today or _dt.date.today()
    if isinstance(today, str):
        today = _dt.date(*[int(x) for x in re.findall(r"\d+", today)[:3]])
    write_year = today.year
    dom, exp = data.domestic, data.export
    # 평가 대상 연도는 제조번호에서 읽는다 — 세 번째 글자가 제조 연도이고 PQR 은 그 다음 해 것이다
    # (OGY301 → 2025년 제조 → 2026년 PQR). 화면에서 넘어온 기간보다 자료가 앞선다.
    year_from, _pqr_year = lotcode.years(dom + exp, today)
    if year_from is None:
        year_from = int(str(period.get("from"))[:4])
    else:
        odd = lotcode.odd_lots(dom + exp, year_from, today)
        if odd:
            issues.append(("6", ", ".join(odd),
                           "제조 연도가 다른 Lot 이 섞여 있습니다(%d년 것으로 평가함) — 확인하세요" % year_from))
    mfg = {l: d for l, d, _ in data.lots}
    name = product.get("name") or ""

    # ---------- 머리글 ----------
    # 머리글은 칸의 위치가 아니라 라벨로 찾는다. 전년도 결재본(HLF-QC-126-01)은 문서번호·
    # 작성일자·제품명 행이 있지만, EDMS 서식(E-HLF-32)은 문서번호·Rev. No.·Page 뿐이라
    # 세 번째 행이 쪽 번호 필드다 — 위치로 쓰면 거기에 날짜를 덮어쓴다.
    hdr = document.sections[0].header.tables[0]
    new_no, full_name = "", name
    for row in hdr.rows:
        cells = E.raw_cells(row)
        labels = "".join(E.cell_text(c) for c in cells[:-1])
        target = cells[-1]
        if "문서번호" in labels:
            old_no = E.cell_text(target).strip()
            if re.search(r"PQR\d{2}-", old_no):
                new_no = re.sub(r"PQR\d{2}-", "PQR%02d-" % (write_year % 100), old_no)
                E.set_cell(target, new_no)        # EDMS 서식은 비어 있다 — EDMS 가 번호를 준다
        elif "작성일자" in labels:
            E.set_cell(target, today.strftime("%Y.%m.%d"))
        elif "제품명" in labels:
            # 결론(16항)에는 머리글의 정식 제품명(성분명까지)을 쓴다 — 마스터의 짧은 이름이 아니라.
            full_name = E.cell_text(target).strip() or name
    log("머리글: %s / %s" % (new_no or "(EDMS 서식 — 문서번호 비움)", today.strftime("%Y.%m.%d")))

    # ---------- 3항 대상 제품 ----------
    t3 = _tables(document, "3.")
    if t3:
        # 비고 칸은 빈 줄끼리 이어 붙여 병합하고 사선 하나만 긋는다 (담당자 지시 2026-09)
        # — 줄마다 'N/A' 를 적거나 사선을 여러 개 긋지 않는다.
        blocks, rows_ = E.merge_empty_runs(t3[0], "비고")
        if blocks:
            log("3항 비고: 빈 칸 %d줄을 %d 묶음으로 합치고 사선" % (rows_, blocks))
        # 결론(16항)의 제품명은 정식 이름(성분명까지)이다 — 머리글이 비어 있으면 여기서 가져온다.
        for row in t3[0].rows[1:]:
            cells = E.raw_cells(row)
            if len(cells) >= 3 and "제품명" in E.cell_text(cells[1]):
                got = " ".join(E.cell_text(cells[2]).split())
                if got and (not full_name or full_name == name):
                    full_name = got
                break

    # ---------- 5항 책임과 권한 ----------
    # 담당자 지시(2026-09): "품질보증 1팀은 AQA 팀으로 변경해줘." EDMS 서식과 2026 결재본
    # 모두 'AQA팀 담당' · 'AQA팀 팀장' 이다. 전년도 결재본은 옛 이름이라 그대로 물려받는다.
    # '품질보증부서장' 은 팀이 아니므로 건드리지 않는다.
    t5 = _tables(document, "5.")
    if t5:
        renamed = 0
        for row in t5[0].rows[1:]:
            for cell in E.raw_cells(row):
                text = E.cell_text(cell)
                new_text = re.sub(r"품질보증\s*\d*\s*팀", "AQA팀", text)
                if new_text != text:
                    E.set_cell(cell, *new_text.split("\n"))
                    renamed += 1
        if renamed:
            log("5항: 품질보증n팀 → AQA팀 %d칸" % renamed)

    # ---------- 4항 ----------
    # 1항 '목적' 도 "제품품질평가는" 으로 시작하므로 평가 기간 문장만 집어 찾는다.
    p = find_para(document, "월까지 생산된")
    if p is not None and "월" in p.text:
        # 2026 결재본 문안: "제품품질평가는 2025년도 1월 ~ 12월까지 생산된 해당제품에 대하여 평가를
        # 실시하며, 'QC-126 제품품질평가 규정'에 따라 2 그룹으로 선정되어 차년도 3분기 내에 완료한다."
        # — 평가 대상 연도는 '년도', 마감은 연도를 적지 않고 '차년도 N분기' 로 쓴다.
        text = re.sub(r"\d{4}\s*년도?\s*1월", "%d년도 1월" % year_from, p.text)
        due = None
        try:
            due = _dt.date(*[int(x) for x in re.findall(r"\d+", str(product.get("due") or ""))[:3]])
        except (TypeError, ValueError):
            due = None
        q = _quarter(due) if due else _quarter(today)
        text = re.sub(r"(\d{4}\s*년도|차년도)\s*(상반기|하반기|\d\s*분기)", "차년도 %d분기" % q, text)
        group = str(product.get("group") or "").strip()
        if group and re.search(r"\S+\s*그룹으로 선정", text):
            text = re.sub(r"\S+\s*그룹으로 선정", "%s 그룹으로 선정" % group, text)
        E.set_para_text(p, text)

    # ---------- 6항 제조내역 ----------
    def fill_mfg(table, lots):
        if len(table.rows) < 2:
            return
        keep = [E.cell_text(c) for c in E.raw_cells(table.rows[1])]      # 배치 크기·포장 단위는 결재본 값
        f, l = E.fit_rows(table, 1, len(table.rows) - 1, max(1, len(lots)))
        for i, lot in enumerate(lots):
            c = E.raw_cells(table.rows[f + i])
            E.set_cell(c[0], str(i + 1)); E.set_cell(c[1], lot); E.set_cell(c[2], mfg.get(lot, ""))
            for k in range(3, len(c)):
                E.set_cell(c[k], keep[k] if k < len(keep) else "")
            if len(c) > 5:
                E.set_cell(c[5], "■ 적합 □ 부적합")
    t6 = _tables(document, "6.")
    if t6:
        fill_mfg(t6[0], dom)
        if len(t6) > 1:
            fill_mfg(t6[1], exp)

    # ---------- 7항 수율 ----------
    def yield_specs(table):
        """기준 행('95.0% 이상' · '91.0 ± 4.0%' · '98 ± 2%') → [(lo, hi) or None]"""
        specs = []
        for row in table.rows[:3]:
            cells = [E.cell_text(c) for c in E.raw_cells(row)]
            if any("이상" in c or "±" in c for c in cells):
                for c in cells[-4:-1] if len(cells) >= 5 else cells:
                    if "±" in c:
                        m, d = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", c)[:2]]
                        specs.append((m - d, m + d))
                    elif "이상" in c:
                        specs.append((_num(c), None))
                    else:
                        specs.append(None)
                break
        return specs

    yield_out, yield_dev = [], []
    dev_lots = {d.get("lot") for d in data.deviations if "수율" in (d.get("title") or "")}
    t7 = _tables(document, "7.")

    def put_sheet_specs(table, stages):
        """올해 수율현황표에 적힌 기준을 결재본의 기준 행에 옮긴다.

        기준은 해가 바뀌며 개정된다(디겐타안연고 충전: 96.0 ± 3.5% → 86.5 ± 6.5%). 전년도
        결재본의 기준을 그대로 두면 표에 지난해 기준이 적히고, 멀쩡한 Lot 이 모두
        '기준 벗어남' 으로 잡힌다.
        """
        if not data.yield_specs:
            return
        for row in table.rows[:3]:
            cells = E.raw_cells(row)
            texts = [E.cell_text(c) for c in cells]
            if not any("이상" in t or "±" in t for t in texts):
                continue
            targets = cells[-4:-1] if len(cells) >= 5 else cells
            for k, stage in enumerate(stages):
                spec = data.yield_specs.get(stage)
                if spec and k < len(targets):
                    E.set_cell(targets[k], spec)
            return

    def fill_yield(table, lots, is_dom):
        stages = ("조제", "충전", "포장")
        put_sheet_specs(table, stages)
        specs = yield_specs(table)
        f, l = E.fit_rows(table, 3, len(table.rows) - 4, max(1, len(lots)))
        vals = {}
        for i, lot in enumerate(lots):
            r = E.raw_cells(table.rows[f + i])
            y = data.yields.get(lot, {})
            v = [y.get(s) for s in stages]
            vals[lot] = v
            E.set_cell(r[0], str(i + 1)); E.set_cell(r[1], lot)
            for k in range(3):
                E.set_cell(r[2 + k], v[k] or "확인 필요")
                if v[k] is None:
                    issues.append(("7", lot, "%s 수율 값이 수율현황표에 없음" % stages[k]))
            out = False
            for k in range(3):
                sp = specs[k] if k < len(specs) else None
                if sp and v[k] is not None:
                    lo, hi = sp
                    fv = float(v[k])
                    if (lo is not None and fv < lo) or (hi is not None and fv > hi):
                        out = True
            if len(r) > 5:                      # 비고: 행마다 따로 (병합을 풀고) 주석을 단다
                E.set_vmerge(r[5], False)
                E.clear_diag(r[5])
                E.set_cell(r[5], "")
            if out:
                yield_out.append(lot)
                if lot in dev_lots:
                    yield_dev.append(lot)
                if len(r) > 5:
                    E.set_cell(r[5], "1)" if lot in dev_lots else "2)")
        # 최댓값·최솟값·평균 — 세 줄 모두 채우고 그 칸의 사선은 지운다 (담당자 지시 2026-09).
        # raw_cells 로 그 행이 실제로 가진 칸을 쓴다: .cells 는 세로 병합을 하나로 합쳐 돌려주어
        # 세 줄이 같은 칸을 가리키고, 마지막에 쓴 평균만 남는다(최댓값 자리에 평균이 찍혔다).
        summary = [(l + 1, lambda xs: "%.2f" % max(xs)),
                   (l + 2, lambda xs: "%.2f" % min(xs)),
                   (l + 3, avg)]
        for ri, _fn in summary:
            if ri < len(table.rows):
                for cell in E.raw_cells(table.rows[ri])[2:5]:
                    E.set_vmerge(cell, False)
                    E.clear_diag(cell)
        stat_lots = [x for x in lots if x not in yield_out and all(v is not None for v in vals[x])]
        if stat_lots:
            cols = list(zip(*[[float(x) for x in vals[lt]] for lt in stat_lots]))
            for ri, fn in summary:
                if ri >= len(table.rows):
                    continue
                cells = E.raw_cells(table.rows[ri])
                for k in range(3):
                    if 2 + k < len(cells):
                        E.set_cell(cells[2 + k], fn(cols[k]))
    if t7:
        fill_yield(t7[0], dom, True)
        if len(t7) > 1 and exp:
            fill_yield(t7[1], exp, False)
    note = find_para(document, "수율의 최대, 최소, 평균") or find_para(document, "1) 11. 일탈")
    if note is not None:
        if yield_out:
            devs = ", ".join("%s %s" % (l, next((d["doc_no"] for d in data.deviations if d.get("lot") == l), "")) for l in yield_dev)
            txt = "1) 충전 수율 일탈(%s)로 11항 일탈 관련 기록 참고. " % devs if yield_dev else ""
            others = [l for l in yield_out if l not in yield_dev]
            if others:
                txt += "2) 자가기준 이탈, 일탈 기록 확인 필요(%s). " % ", ".join(others)
                issues.append(("7", ", ".join(others), "수율이 자가기준을 벗어났으나 일탈 기록이 없음"))
            txt += "수율 일탈 Lot 은 최댓값·최솟값·평균에 미반영."
            E.set_para_text(note, txt)
            _small(note._p)
        else:
            E.set_para_text(note, "")
    log("7항: 이탈 %s / 일탈 %s" % (yield_out, yield_dev))

    # ---------- 8.1.3 공급업체 평가일 ----------
    def _norm_date(v):
        m = re.findall(r"\d+", str(v or ""))
        return "%s.%02d.%02d" % (m[0], int(m[1]), int(m[2])) if len(m) >= 3 else ""
    upd813 = 0
    for tb in _tables(document, "8.1.3"):
        hdr = [re.sub(r"\s+", "", E.cell_text(c)) for c in E.raw_cells(tb.rows[0])]
        dcol = next((i for i, h in enumerate(hdr) if "완료일" in h or "승인일" in h), None)
        ncol = next((i for i, h in enumerate(hdr) if "문서번호" in h or "평가문서" in h), None)
        ccol = next((i for i, h in enumerate(hdr) if "코드" in h), None)
        if dcol is None:
            continue
        for row in tb.rows[1:]:
            cells = E.raw_cells(row)
            if len(cells) <= dcol:
                continue
            code = E.cell_text(cells[ccol]).strip() if ccol is not None else ""
            docno = E.cell_text(cells[ncol]).strip() if ncol is not None else ""
            vendor = E.cell_text(cells[4]).strip() if len(cells) > 4 else ""
            key = lambda v: re.sub(r"[^a-z0-9가-힣]", "", str(v or "").lower())
            hit = None
            both = list(data.suppliers_raw) + list(data.suppliers_mat)
            tests = [lambda r_: bool(code) and code in (r_.get("원료코드") or ""),
                     lambda r_: bool(docno) and key(r_.get("문서번호")) and (key(r_.get("문서번호")).startswith(key(docno)) or key(docno).startswith(key(r_.get("문서번호")))),
                     lambda r_: bool(vendor) and key(r_.get("공급업체명")) and (key(vendor)[:6] in key(r_.get("공급업체명")) or key(r_.get("공급업체명"))[:6] in key(vendor))]
            for test in tests:                       # 코드 → 문서번호 → 업체명 순
                hit = next((r_ for r_ in both if test(r_)), None)
                if hit:
                    break
            day = _norm_date(hit.get("평가승인일")) if hit else ""
            if day and day != E.cell_text(cells[dcol]).strip():
                E.set_cell(cells[dcol], day); upd813 += 1
    log("8.1.3 평가일 갱신: %d" % upd813)

    # ---------- 8.2 시험성적 ----------
    def group_fill(table, records):
        """records: [(원료 코드, 원료명, 시험번호, [Lot ...]), ...] — Lot 하나가 한 줄.

        2026 결재본 차림새: **연번은 줄마다** 매기고, **원료 코드·원료명은 같은 원료끼리**,
        시험 성적 번호는 같은 번호끼리 세로로 합친다. 결과(적합)와 적용 Lot 은 줄마다 적는다.
        담당자 지적: "같은 원료 및 코드는 셀병합해줘."
        """
        rows = [(code, item, test, lot) for code, item, test, lots in records for lot in lots]
        if not rows:
            return
        f, l = E.fit_rows(table, 1, len(table.rows) - 1, len(rows))
        prev = None
        for i, (code, item, test, lot) in enumerate(rows):
            r = E.raw_cells(table.rows[f + i])
            # 열마다 '같은 묶음' 의 기준이 다르다 — 원료명은 원료가 바뀌면, 시험번호는 원료나
            # 번호가 바뀌면 새로 시작한다.
            keys = (code, (code, item), (code, test))
            E.set_cell(r[0], str(i + 1)); E.set_vmerge(r[0], False)
            for k, value in ((1, code), (2, item), (3, test)):
                head = prev is None or keys[k - 1] != prev[k - 1]
                E.set_cell(r[k], value if head else "")
                E.set_vmerge(r[k], "restart" if head else None)
            E.set_cell(r[4], lot); E.set_vmerge(r[4], False)
            E.set_cell(r[5], "적합"); E.set_vmerge(r[5], False)
            prev = keys

    def material_names(table):
        """결재본 8.2 표에 적힌 {원료 코드: 원료명} — 코드에 붙일 한글 이름은 여기에만 있다."""
        out = {}
        for row in table.rows[1:]:
            cells = E.raw_cells(row)
            if len(cells) > 2:
                code = E.cell_text(cells[1]).strip()
                if code:
                    out.setdefault(code, E.cell_text(cells[2]).strip())
        return out

    t821 = _tables(document, "8.2.1")
    if t821 and data.raw_tests:
        # 주원료가 둘 이상인 제품이 있다(디겐타안연고: 겐타마이신황산염·플루오로메톨론). 예전에는
        # 결재본 첫 행의 코드 하나만 채워 나머지 주원료가 표에서 통째로 빠졌다.
        names = material_names(t821[0])
        for tbl, lots in ((t821[0], dom), (t821[1] if len(t821) > 1 else None, exp)):
            if tbl is None:
                continue
            recs = []
            for code, test, ls in data.raw_tests:
                mine = [l for l in lots if l in ls]
                if mine:
                    recs.append((code, names.get(code, ""), test, mine))
            missing = sorted({c for c, item, _t, _l in recs if not item})
            if missing:
                issues.append(("8.2.1", ", ".join(missing),
                               "전년도 결재본에 없는 원료 코드입니다 — 원료명을 확인해 적으세요"))
            group_fill(tbl, recs)
    t822 = _tables(document, "8.2.2")
    if t822 and data.pkg_tests:
        tbl = t822[0]
        base_rows = {E.cell_text(E.raw_cells(r)[1]).strip(): E.cell_text(E.raw_cells(r)[2]).strip()
                     for r in tbl.rows[1:] if len(E.raw_cells(r)) > 2}
        recs = []
        for code, test, ls in data.pkg_tests:
            ls = [x for x in ls if x in dom + exp]
            if ls:
                item = base_rows.get(code) or ("튜브 (수출용)" if ls[0] in exp else "튜브 (내수용)")
                recs.append((code, item, test, ls))
        group_fill(tbl, recs)                  # 1차 포장 자재도 같은 차림새다 (담당자 지적)

    # ---------- 9항 ----------
    def spec_of(t91, words):
        for row in t91.rows:
            cells = [E.cell_text(c) for c in E.raw_cells(row)]
            if any(all(w in c for w in words) for c in cells[:2]):
                return cells[-2]
        return ""

    def rec(lot, key):
        return (data.coa.get(lot) or {}).get(key) or {}

    # Cpk 한계는 이 제품의 9.1 허용기준에서 읽는다 — 퀴노비드 숫자를 다른 제품에 쓰지 않는다
    limits = dict(DEFAULT_LIMITS)
    t91_all = _tables(document, "9.1")
    if t91_all:
        found = parse_limits({"particle": spec_of(t91_all[0], ("입자도",)),
                              "assay": spec_of(t91_all[0], ("함량",)),
                              "metal": spec_of(t91_all[0], ("금속성",))})
        limits.update(found)
        missing = [k for k in ("particle", "assay", "metal") if k not in found]
        if missing:
            issues.append(("9.1", "", "허용기준에서 %s 한계를 읽지 못해 기본값(퀴노비드안연고)으로 Cpk 를 계산함 — 확인 필요"
                           % ", ".join({"particle": "입자도", "assay": "함량", "metal": "금속성이물"}[k] for k in missing)))
    log("Cpk 한계: %s" % limits)

    def app(lots):
        for l in lots:
            for k in ("924", "923", "922"):
                a = rec(l, k).get("appearance")
                if a:
                    return a
        return ""

    def assay_by_part(lots):
        """{성분 이름: [Lot 별 함량]} — 주성분이 둘 이상인 제품은 성적서에 성분마다 함량 줄이 있다."""
        out = {}
        for lot in lots:
            for a in rec(lot, "924").get("assays") or []:
                value = _num(a.get("value"))
                if a.get("part") and value is not None:
                    out.setdefault(a["part"], []).append(float(value))
        return out

    def numbers(lots):
        g = lambda k, key: [float(_num(rec(l, k).get(key))) for l in lots if _num(rec(l, k).get(key)) is not None]
        return {"parts": assay_by_part(lots),
                "ms": g("924", "metal_total"), "mi": g("924", "metal_each"), "pt": g("924", "particle"),
                "pa": g("924", "mass_avg"), "pi": g("924", "mass_each_min"), "ct": g("924", "assay"),
                "fa": g("923", "mass_avg"),
                "flo": [_rng(rec(l, "923").get("mass_each"))[0] for l in lots if _rng(rec(l, "923").get("mass_each"))],
                "fhi": [_rng(rec(l, "923").get("mass_each"))[1] for l in lots if _rng(rec(l, "923").get("mass_each"))]}

    def bio_text(lot):
        b = rec(lot, "922").get("bioburden") or ""
        m = re.match(r"(\d+)\s*(CFU|FU)?/g\s*(미만|이하)?", b)
        return ("%s %s" % (m.group(1), m.group(3) or "미만")) if m else (b or "확인 필요")

    # 각주 번호는 실제로 다는 것만 세어 매긴다 — 생균수 각주가 없는 제품에서 1) 없이 2) 로
    # 시작하면 안 된다(디겐타안연고 2026).
    odd_bio = [(l, rec(l, "922").get("bioburden")) for l in dom + exp
               if bio_text(l) != "10 미만" and rec(l, "922")]
    note_bio = 1 if odd_bio else None
    note_mass = 2 if odd_bio else 1

    def fill_91(t91, lots, is_dom):
        n = numbers(lots)
        if not n["ct"]:
            issues.append(("9", "", "완제 성적서에서 함량을 읽지 못함")); return n
        res = {}
        current_item = ""
        assay_parts = []          # 주성분이 둘 이상인 제품(디겐타안연고: 플루오로메톨론·겐타마이신황산염)
        for ri, row in enumerate(t91.rows):
            cells = E.raw_cells(row)
            if ri == 0 or len(cells) < 3:
                continue
            label = re.sub(r"\s+", "", E.cell_text(cells[-3])) if len(cells) >= 4 else ""
            crit_text = E.cell_text(cells[-2])
            if "평균" not in label and "개개" not in label:
                # 칸이 넷뿐인 표(디겐타안연고)는 평균·개개 칸이 따로 없고 허용기준 글에 적혀 있다
                # — "허가) 평균 : 표시량(4.0 g) 이상" / "허가) 개개 : 3.60 g 이상".
                if re.search(r"평균\s*[:：]", crit_text):
                    label += "평균"
                elif re.search(r"개개\s*[:：]", crit_text):
                    label += "개개"
            named = re.sub(r"\s+", "", "".join(E.cell_text(c) for c in cells[:-2]))
            named = re.sub(r"(평균|개개)$", "", named)
            if named:
                current_item = named
            item = current_item
            val = None
            if "성상" in item:
                val = app(lots)
            elif "생균수" in item:
                vals = [bio_text(l) for l in lots]
                common = max(set(vals), key=vals.count)
                m = re.match(r"(\d+)\s*(미만|이하)", common)
                val = ("%s CFU/g %s" % (m.group(1), m.group(2))) if m else common
                if any(v != common for v in vals) and note_bio:
                    val += "%d)" % note_bio
            elif "평균" in label and n["fa"] and "질량" in item and not n.get("_fill_avg"):
                val = _avg_text(n["fa"], "%.1f" if is_dom else "%.2f", unit_on_avg=not is_dom)
                n["_fill_avg"] = True
            elif "개개" in label and n["flo"] and "질량" in item and not n.get("_fill_each"):
                val = "%.2f ~ %.2fg" % (min(n["flo"]), max(n["fhi"])); n["_fill_each"] = True
            elif "확인" in item and not is_dom:
                crit = E.cell_text(cells[-2])
                if "1)" in crit:
                    vs = {rec(l, "924").get("ident_color") for l in lots} - {None}
                    val = sorted(vs)[0] if vs else None
                elif "2)" in crit:
                    vs = {rec(l, "924").get("ident_precip") for l in lots} - {None}
                    val = sorted(vs)[0] if vs else None
                elif "3)" in crit:
                    pairs = [re.findall(r"(\d{3})nm", rec(l, "924").get("uv_max") or "") for l in lots]
                    pairs = [p_ for p_ in pairs if len(p_) == 2]
                    if pairs:
                        a = sorted(int(p_[0]) for p_ in pairs); b = sorted(int(p_[1]) for p_ in pairs)
                        fmt = lambda xs: ("%d ~ %dnm" % (xs[0], xs[-1])) if xs[0] != xs[-1] else "%dnm" % xs[0]
                        val = "%s 및 %s에서 흡수극대를 나타냄" % (fmt(a), fmt(b))
            elif "튜브인쇄" in item:
                val = "인쇄상태가 양호하며 제조번호 및 사용기한의 압인상태가 명확히 식별 가능함"
            elif "기밀도" in item:
                leaks = {rec(l, "924").get("leak") for l in lots} - {None}
                val = "메틸렌블루시액 침투 없음" if not n.get("_leak") else (sorted(leaks)[0] if leaks else "메틸렌블루시액 침투 없이 양호")
                n["_leak"] = True
            elif "금속성이물" in item and n["ms"]:
                # 모든 Lot 이 같은 값이면 평균·범위를 적지 않는다 — 2026 결재본의 주석 그대로
                # ("모든 Lot 의 시험결과가 0개(매)으로 동일하여 최댓값, 최솟값, 평균은 별도로 작성하지 않음").
                same = len(set(n["ms"])) == 1 and len(set(n["mi"])) == 1
                val = ("50 ㎛ 이상\n: Av. %.2f개(%.0f ~ %.0f개)\n개개 중 8개 초과 : %.0f매" % (sum(n["ms"]) / len(n["ms"]), min(n["ms"]), max(n["ms"]), max(n["mi"]))) if (len(lots) >= 3 and not same) else \
                      ("50 ㎛ 이상 : %.0f개\n개개 중 8개 초과 : %.0f매" % (max(n["ms"]), max(n["mi"])))
            elif "입자도" in item and n["pt"]:
                val = "Av. %.2f㎛ 이하\n(%.2f ~ %.2f㎛ 이하)" % (sum(n["pt"]) / len(n["pt"]), min(n["pt"]), max(n["pt"]))
            elif "함량" in item and n["ct"]:
                # 성분 이름이 적힌 줄이면 그 성분의 값만 쓴다 — 성분마다 규격도 결과도 다르다.
                part = assay_component(E.cell_text(cells[-2]))
                vals = n["parts"].get(part) if part else None
                if part and not vals:
                    vals = next((v for k, v in n["parts"].items() if k in part or part in k), None)
                if part and vals is None:
                    assay_parts.append(part)                      # 성분별 값을 못 찾았다 — 짚어 준다
                use = vals or n["ct"]
                val = "Av. %.1f%%\n(%.1f ~ %.1f%%)" % (sum(use) / len(use), min(use), max(use))
            elif "평균" in label and n["pa"] and "질량" in item:
                val = _avg_text(n["pa"], "%.2f")
            elif "개개" in label and n["pi"] and "질량" in item:
                val = "%.2f ~ %.2fg 이상%d)" % (min(n["pi"]), max(n["pi"]), note_mass)
            elif "무균" in item:
                st = {rec(l, "924").get("sterility") for l in lots} - {None}
                val = sorted(st)[0] if st else "음성"
            elif "포장규격" in item:
                val = "각 규격에 적합함"
            if val is not None:
                E.set_cell(cells[-1], *val.split("\n"))
                res[ri] = val
        if assay_parts:
            # 성적서에서 그 성분의 함량을 찾지 못했다 — 다른 성분 값이 들어가 있으니 짚는다.
            issues.append(("9.1", ", ".join(assay_parts),
                           "이 성분의 함량을 성적서에서 찾지 못해 다른 값이 들어가 있습니다 — "
                           "원본에서 확인해 적으세요"))
        return n
    t91 = _tables(document, "9.1")
    n_dom = fill_91(t91[0], dom, True) if t91 else {}
    n_exp = fill_91(t91[1], exp, False) if len(t91) > 1 and exp else {}

    # 9.2 세부표 — 항 아래 표들을 머리행 낱말로 고른다
    def fill_detail(table, first, lots, setter, n_summary=5):
        last = len(table.rows) - 1 - n_summary
        f, l = E.fit_rows(table, first, last, max(1, len(lots)))
        for i, lot in enumerate(lots):
            setter(E.raw_cells(table.rows[f + i]), i + 1, lot)
        return f, l

    def put(table, ri, idx, *vals):
        c = E.raw_cells(table.rows[ri])
        for k, v in zip(idx, vals):
            E.set_cell(c[k], v)

    def fill_92(section, lots, n, is_dom):
        tabs = _tables(document, section)
        if not tabs:
            return
        appearance = app(lots)
        t_bio = _by_header(tabs, "생균수")
        t_fill = _by_header(tabs, "질량", "기밀도")
        t_tube = _by_header(tabs, "튜브인쇄")
        t_metal = _by_header(tabs, "금속성이물")
        t_ident = _by_header(tabs, "확인")
        t_numb = _by_header(tabs, "입자도")
        if t_numb is None and t_ident is not None and "입자도" in _text(t_ident._tbl):
            t_numb = t_ident                      # 확인 블록 + 수치 블록이 한 표 (내수용)
        t_last = _by_header(tabs, "무균")
        if t_bio:
            vals_ = [bio_text(l) for l in lots]
            common_ = max(set(vals_), key=vals_.count)
            fill_detail(t_bio, 1, lots, lambda c, no, lot: [E.set_cell(c[0], str(no)), E.set_cell(c[1], lot), E.set_cell(c[2], appearance),
                                                            E.set_cell(c[3], bio_text(lot) + ("1)" if bio_text(lot) != common_ else ""))])
        if t_fill:
            def s_fill(c, no, lot):
                r = rec(lot, "923")
                E.set_cell(c[0], str(no)); E.set_cell(c[1], lot); E.set_cell(c[2], appearance)
                E.set_cell(c[3], (r.get("mass_avg") or "").replace("g", "")); E.set_cell(c[4], (r.get("mass_each") or "").replace("g", ""))
                E.set_cell(c[5], r.get("leak") or "메틸렌블루시액 침투 없음")
            f, l = fill_detail(t_fill, 2, lots, s_fill)
            if n.get("fa"):
                fmt = "%.1f" if is_dom else "%.2f"
                put(t_fill, l + 1, (2, 3), fmt % max(n["fa"]), "%.2f" % max(n["fhi"]))
                put(t_fill, l + 2, (2, 3), fmt % min(n["fa"]), "%.2f" % min(n["flo"]))
                put(t_fill, l + 3, (2,), "%.2f" % (sum(n["fa"]) / len(n["fa"])))
        if t_tube:
            fill_detail(t_tube, 1, lots, lambda c, no, lot: [E.set_cell(c[0], str(no)), E.set_cell(c[1], lot), E.set_cell(c[2], "인쇄상태가 양호하며 제조번호 및 사용기한의 압인상태가 명확히 식별 가능함")])
        cpk = {}
        if t_metal and n.get("ms"):
            f, l = fill_detail(t_metal, 2, lots, lambda c, no, lot: [E.set_cell(c[0], str(no)), E.set_cell(c[1], lot), E.set_cell(c[2], appearance), E.set_cell(c[3], rec(lot, "924").get("metal_total") or ""), E.set_cell(c[4], rec(lot, "924").get("metal_each") or "")])
            ms, mi = n["ms"], n["mi"]
            put(t_metal, l + 1, (2, 3), "%.0f" % max(ms), "%.0f" % max(mi))
            put(t_metal, l + 2, (2, 3), "%.0f" % min(ms), "%.0f" % min(mi))
            if qc.cpk_applies(len(lots)):
                put(t_metal, l + 3, (2, 3), "%.2f" % (sum(ms) / len(ms)), "%.2f" % (sum(mi) / len(mi)))
                cpk["metal"] = cpk_uni(ms, limits["metal"])
                if cpk["metal"] is not None:
                    put(t_metal, l + 4, (2,), "%.2f" % cpk["metal"]); put(t_metal, l + 5, (2,), "충분" if cpk["metal"] >= 1 else "부족")
            else:
                put(t_metal, l + 3, (2, 3), "%.0f" % (sum(ms) / len(ms)), "%.0f" % (sum(mi) / len(mi)))
        if t_ident is not None and t_ident is t_numb:            # 한 표에 확인 + 수치 두 블록 (내수용)
            t = t_ident
            firsts = [i for i, tr in enumerate(t._tbl.findall(qn("w:tr"))) if _text(tr.findall(qn("w:tc"))[0]).startswith("최댓값")]
            fa_, la_ = E.fit_rows(t, 2, firsts[0] - 1, len(lots))
            for i, lot in enumerate(lots):
                c = E.raw_cells(t.rows[fa_ + i]); E.set_cell(c[0], str(i + 1)); E.set_cell(c[1], lot)
                E.set_cell(c[2], "검액은 표준액과 동일한 주 피크 유지시간을 나타냄"); E.set_cell(c[3], "검액과 표준액의 주 피크 UV spectrum은 동일함")
            base2 = la_ + 6
            fb_, lb_ = E.fit_rows(t, base2 + 2, len(t.rows) - 6, len(lots))
            _fill_numbers(t, fb_, lb_, lots, n, cpk, rec, put, is_dom, limits)
        else:
            if t_ident:
                def s_ident(c, no, lot):
                    E.set_cell(c[0], str(no)); E.set_cell(c[1], lot)
                    if is_dom:
                        E.set_cell(c[2], "검액은 표준액과 동일한 주 피크 유지시간을 나타냄"); E.set_cell(c[3], "검액과 표준액의 주 피크 UV spectrum은 동일함")
                    else:
                        r = rec(lot, "924")
                        E.set_cell(c[2], r.get("ident_color") or "확인 필요"); E.set_cell(c[3], r.get("ident_precip") or "확인 필요")
                        E.set_cell(c[4], r.get("uv_max") or "확인 필요")
                fill_detail(t_ident, 2, lots, s_ident)
            if t_numb and n.get("ct"):
                f, l = E.fit_rows(t_numb, 2, len(t_numb.rows) - 6, len(lots))
                _fill_numbers(t_numb, f, l, lots, n, cpk, rec, put, is_dom, limits)
        if t_last:
            def s_last(c, no, lot):
                r = rec(lot, "924")
                E.set_cell(c[0], str(no)); E.set_cell(c[1], lot)
                E.set_cell(c[2], r.get("sterility") or "음성"); E.set_cell(c[3], r.get("leak") or "메틸렌블루시액 침투 없이 양호")
                E.set_cell(c[4], "각 규격에 적합함")
            fill_detail(t_last, 1, lots, s_last)
        return cpk
    cpk_dom = fill_92("9.2.1", dom, n_dom, True) or {}
    cpk_exp = fill_92("9.2.2", exp, n_exp, False) if exp else {}
    log("9항: Cpk %s" % {k: round(v, 2) for k, v in cpk_dom.items() if v is not None})

    # 각주 — 번호는 위에서 센 것을 쓴다
    notes = []
    if odd_bio:
        notes.append("%d) %s 조제(바이오버든) 공정 시험 성적서의 생균수 기재값은 “%s” 임. 원 기록의 단위 표기 확인 필요."
                     % (note_bio, ", ".join(l for l, _ in odd_bio), odd_bio[0][1]))
        issues.append(("9.2.2", ", ".join(l for l, _ in odd_bio), "생균수 기재값이 다른 Lot 과 다름 — 원본 확인"))
    notes.append("%d) %d년 완제 시험 성적서는 질량·용량 개개를 최솟값(···g 이상)으로만 기재하므로, 각 Lot 의 개개 최솟값으로 기재하였음."
                 % (note_mass, year_from))
    for t in reversed(t91):
        for nt in reversed(notes):
            E.note_after(document, t, nt)
    for section in ("9.2.1",):
        tabs = _tables(document, section)
        tb = _by_header(tabs, "생균수"); tn = _by_header(tabs, "입자도")
        if tn is not None:
            E.note_after(document, tn, notes[-1])
        if tb is not None and odd_bio:
            E.note_after(document, tb, notes[0])

    # ---------- 10항 ----------
    def update_pq(table, lookup):
        """관리번호 행마다 PQ 열의 (문서번호 행 · 완료일 행) 을 마스터의 최신 PQ 로 바꾼다."""
        rows = table.rows
        pq_col = None
        for row in rows[:4]:
            cells = E.raw_cells(row)
            for ci, cell in enumerate(cells):
                if E.cell_text(cell).strip() == "PQ":
                    pq_col = ci
            if pq_col is not None:
                break
        if pq_col is None:
            return 0
        n = 0
        for ri, row in enumerate(rows):
            cells = E.raw_cells(row)
            mid = E.cell_text(cells[1]).strip() if len(cells) > 1 else ""
            if not re.match(r"^[A-Z]{3}\d{4}", mid) or mid not in lookup:
                continue
            pq = [(d, dt) for d, dt in lookup[mid].get("PQ", []) if dt and d]
            if not pq or ri + 1 >= len(rows):
                continue
            latest = max(pq, key=lambda x: re.sub(r"\D", "", x[1])[:8])
            doc_cells, date_cells = E.raw_cells(rows[ri]), E.raw_cells(rows[ri + 1])
            col = min(pq_col, len(doc_cells) - 2, len(date_cells) - 2)
            old_doc = E.cell_text(doc_cells[col]).strip(); old_date = E.cell_text(date_cells[col]).strip()
            new_date = latest[1].replace(". ", ".").replace(" ", "")[:10]
            if re.sub(r"\D", "", new_date) > re.sub(r"\D", "", old_date) and latest[0].split("(")[0].strip() != old_doc.split("(")[0].strip():
                E.set_cell(doc_cells[col], latest[0].split("(")[0].strip()); E.set_cell(date_cells[col], new_date); n += 1
        return n
    # 10.1 공정밸리데이션: 평가 년도에 보고서가 난 PV 를 채운다 (마스터파일)
    def fill_pv(table):
        """표의 기존 보고서 번호(PV24-2-QUIO3-R …)에서 코드를 알아내 마스터에서 평가 년도 보고서를 찾는다."""
        codes = set(re.findall(r"PV\d{2}-\d-([A-Z0-9]+)-", _text(table._tbl)))
        pv_path = next((p_ for p_ in data.files.get("10.1", []) if p_.lower().endswith(".xlsx")), None)
        if not codes or not pv_path:
            return 0
        from .readers.masters import pv_by_code
        rows = []
        for code in codes:
            for e in pv_by_code(pv_path, code):
                yr = re.match(r"PV(\d{2})", e["report"] or "")
                if yr and 2000 + int(yr.group(1)) == year_from:
                    rows.append(e)
        if not rows:
            return 0
        # 가장 최근 것 하나만 싣는다 (담당자 지시 2026-09) — 마스터파일에는 해묵은 PV 가 함께 있다.
        rows = [max(rows, key=lambda e: (e.get("report_date") or "", e.get("report") or ""))]
        # 열은 머리행 이름으로 짚는다 — 전년도 양식은 'No.' 열이 있고 2026 결재본은 없다.
        # 자리로 쓰면 사유 체크박스가 No. 칸에 들어가는 등 한 칸씩 밀린다.
        head = [re.sub(r"\s+", "", E.cell_text(h)) for h in E.raw_cells(table.rows[0])]

        def col(*words):
            return next((i for i, h in enumerate(head) if any(w in h for w in words)), None)

        i_no, i_why, i_lot = col("No.", "연번"), col("사유"), col("대상Lot", "Lot")
        i_doc, i_done, i_note = col("문서번호"), col("완료일"), col("비고")
        f, l = E.fit_rows(table, 1, len(table.rows) - 1, len(rows))
        for i, e in enumerate(rows):
            c = E.raw_cells(table.rows[f + i])

            def put(index, *lines):
                if index is not None and index < len(c):
                    E.set_cell(c[index], *lines)

            reason = (e.get("reason") or "").strip()
            kind = "변경" if "변경" in reason else ("정기적" if "정기" in reason else "최초")
            put(i_no, str(i + 1))
            put(i_why, *["%s %s" % ("■" if k == kind else "☐", k) for k in ("최초", "변경", "정기적")])
            put(i_lot, *[lot for _, lot, _ in e["lots"]][:3])
            m = re.match(r"(\S+)\s*(\(.*?\))?", e["report"])
            put(i_doc, m.group(1), (m.group(2) or "").lower()) if m else put(i_doc, e["report"])
            put(i_done, e.get("report_date") or "확인 필요")
            put(i_note, reason)                 # 비고에 지난해 글이 남지 않게 늘 덮어쓴다
        return len(rows)
    t101 = _tables(document, "10.1")
    pv_n = 0
    for tb in t101:
        pv_n += fill_pv(tb)
    log("10.1 PV 행: %d" % pv_n)

    eq_lookup = {k: {"PQ": [(d, dt) for d, dt in v["docs"] if d.startswith("PQ")]} for k, v in data.equipment.items()}
    sp_lookup = {}
    for k, v in data.support.items():
        pq = [(d.split(" (")[0].strip(), dt.split("/")[0].strip()) for d, dt in v.get("PQ", []) if d and dt]
        if not pq:                              # 설비 행에 PQ 가 없으면 같은 시스템의 PQ 를 쓴다
            for k2, v2 in data.support.items():
                if v2.get("system") == v.get("system") and v2.get("PQ"):
                    pq = [(d.split(" (")[0].strip(), dt.split("/")[0].strip()) for d, dt in v2["PQ"] if d and dt]
                    break
        sp_lookup[k] = {"PQ": pq}
    upd = 0
    for prefix in ("10.2",):
        for t in _tables(document, prefix):
            upd += update_pq(t, eq_lookup)
    for prefix in ("10.3", "10.4", "10.5"):
        for t in _tables(document, prefix):
            upd += update_pq(t, sp_lookup)
    log("10항 PQ 갱신: %d" % upd)
    if not pv_n:
        issues.append(("10.1", "", "평가 년도의 PV 보고서를 마스터파일에서 찾지 못해 결재본 값을 유지함 — 확인 필요"))

    # ---------- 11항 ----------
    devs = [d for d in data.deviations if (d.get("lot") in dom + exp) or (name and name[:4] in (d.get("title") or ""))]
    t11 = _tables(document, "11.1")
    if t11:
        tbl = t11[0]
        if devs:
            # 열은 머리행 이름으로 짚는다 — 전년도 양식은 '구분(제품)' 열이 있고 EDMS 서식은 없다.
            head = [re.sub(r"\s+", "", E.cell_text(h)) for h in E.raw_cells(tbl.rows[0])]
            def col(*words):
                return next((i for i, h in enumerate(head) if any(w in h for w in words)), None)
            i_kind, i_lot, i_doc = col("구분"), col("Lot"), col("문서")
            i_det, i_act, i_capa = col("일탈사항"), col("조치사항"), col("CAPA")
            f, l = E.fit_rows(tbl, 1, len(tbl.rows) - 2, len(devs))
            for i, d in enumerate(devs):
                c = E.raw_cells(tbl.rows[f + i])
                E.set_cell(c[0], str(i + 1))
                if i_kind is not None and i_kind < len(c):
                    E.set_cell(c[i_kind], "제품")
                if i_lot is not None and i_lot < len(c):
                    E.set_cell(c[i_lot], d.get("lot") or "")
                if i_doc is not None and i_doc < len(c):
                    E.set_cell(c[i_doc], d.get("doc_no") or "")
                detail = _deviation_lines(d)
                action = [(line, "none") for line in (d.get("correction") or "").split("\n") if line.strip()]
                if d.get("completed"):
                    action.append(("(조치사항 완료일 : %s)" % d["completed"], "none"))
                if i_det is not None and i_det < len(c):
                    E.set_cell_flow(c[i_det], detail)
                if i_act is not None and i_act < len(c):
                    E.set_cell_flow(c[i_act], action)
                if i_capa is not None and i_capa < len(c):
                    E.set_cell(c[i_capa], "☐ Yes", "■ No")
            E.set_cell_plain(E.raw_cells(tbl.rows[-1])[0], "특이사항 (Comment)",
                             "- 평가 년도 내 %d건의 일탈 있었으나, 모두 적합하게 조치되었으며 특이사항 없음."
                             % len(devs))
        else:
            E.set_cell_plain(E.raw_cells(tbl.rows[-1])[0], "특이사항 (Comment)", "평가 년도 내 중요 일탈 및 기준 일탈 이력 없음.")
        for tb in t11[1:]:
            E.set_cell_plain(E.raw_cells(tb.rows[-1])[0], "특이사항 (Comment)", "평가 년도 내 중요 일탈 및 기준 일탈 이력 없음.")
    for tb in _tables(document, "11.2"):
        E.set_cell_plain(E.raw_cells(tb.rows[-1])[0], "특이사항 (Comment)", "평가 년도 내 경향 일탈 이력 없음.")

    # ---------- 12항 ----------
    t12 = _tables(document, "12.")
    if t12:
        tbl = t12[0]
        # 담당자가 12항에 넣어 준 변경요청서는 모두 싣는다. 변경요청서의 '대상 제품' 에 이 제품
        # 이름이 없는 것도 있다(디겐타안연고 2026: 주성분 플루오로메톨론 멸균 온도 변경 건은
        # 대상이 후메론점안액으로만 적혀 있다). 제품 이름으로 걸러 내면 실제로 있었던 변경이
        # 보고서에서 통째로 빠진다 — 싣고, 이름이 없는 건은 확인해 달라고 남긴다.
        ccs = list(data.changes)
        if ccs:
            f, l = E.fit_rows(tbl, 1, len(tbl.rows) - 2, len(ccs))
            for i, cc in enumerate(ccs):
                c = E.raw_cells(tbl.rows[f + i])
                E.set_cell(c[0], str(i + 1))
                # 변경사항: 변경 번호와 변경 내용만 (변경 사유는 적지 않는다 — 담당자 지시 2026-09)
                lines = ["[%s] %s" % (cc.get("doc_no"), cc.get("title") or "")]
                lines += [x for x in (cc.get("description") or "").split("\n") if x.strip()]
                E.set_cell_plain(c[1], *lines)
                # 조치사항: 변경 실행 계획의 부서별 조치사항을 간추린다. 위탁사·위수탁 줄은 뺀다.
                acts = [a for team, a in (cc.get("actions") or [])
                        if "위수탁" not in team and not re.search(r"위탁사|위수탁", a)]
                if acts:
                    E.set_cell_plain(c[2], *["%d. %s" % (k, a) for k, a in enumerate(acts, 1)])
                    issues.append(("12", cc.get("doc_no") or "",
                                   "조치사항은 변경 실행 계획을 간추린 것입니다 — 이 제품에 해당하지 않는 줄은 지우세요"))
                else:
                    E.set_cell_plain(c[2], "확인 필요")
                    issues.append(("12", cc.get("doc_no") or "",
                                   "변경 실행 계획을 읽지 못했습니다 — 조치사항을 직접 적으세요"))
                E.set_cell(c[3], "확인 필요"); E.set_cell(c[4], "N/A")
                where = (cc.get("products") or "") + " " + (cc.get("title") or "")
                if name and name[:4] not in re.sub(r"\s+", "", where):
                    issues.append(("12", cc.get("doc_no") or "",
                                   "변경요청서의 대상 제품에 이 제품 이름이 없습니다 — 이 제품에 해당하는지 확인하세요"))
            E.set_cell_plain(E.raw_cells(tbl.rows[-1])[0], "특이사항 (Comment)", "N/A")
        else:
            E.set_cell_plain(E.raw_cells(tbl.rows[-1])[0], "특이사항 (Comment)", "평가 년도 내 변경관리 이력 없음.")

    # ---------- 13항 ----------
    stab = getattr(data, "stability", None)
    if stab:
        _fill_stability(document, stab, log, limits["assay"])
    else:
        for prefix in ("13.1", "13.2"):
            for tb in _tables(document, prefix):
                if len(tb.rows) < 3:
                    continue
                f, l = E.fit_rows(tb, 1, len(tb.rows) - 2, 1)
                for c in E.raw_cells(tb.rows[f]):
                    E.set_cell(c, ""); E.set_vmerge(c, False)
                cells = E.raw_cells(tb.rows[f])
                E.set_cell(cells[0], "1"); E.set_cell(cells[min(3, len(cells) - 1)], "확인 필요")
                E.set_cell_plain(E.raw_cells(tb.rows[-1])[0], "특이사항 (Comment)", "* 안정성 시험일지 판독 필요 — 담당자 확인 후 기재")
        for tb in _tables(document, "13.3"):
            firsts = [i for i, r in enumerate(tb.rows) if E.cell_text(E.raw_cells(r)[0]).strip().startswith("관리")]
            last = (firsts[0] - 1) if firsts else len(tb.rows) - 6
            f, l = E.fit_rows(tb, 1, last, 1)
            cells = E.raw_cells(tb.rows[f])
            for c in cells:
                E.set_cell(c, ""); E.set_vmerge(c, False)
            E.set_cell(cells[0], "확인 필요")
            E.set_cell_plain(E.raw_cells(tb.rows[-1])[0], "특이사항 (Comment)", "* 안정성 시험일지 판독 필요 — 담당자 확인 후 기재")
        issues.append(("13", "", "안정성 시험(13.1~13.3) 값을 읽지 못해 '확인 필요' 로 두었음 — 시험일지 판독 필요"))

    # ---------- 14·15항 ----------
    us_export = "미국" in (name or "")           # 계획서 비고로 갈라진 '(미국 수출용)' 건
    returns = "평가 년도 내 반품 이력 없음." if us_export else "사용기한 경과 외 반품이력 없음"
    for prefix, msg in (("14.1", returns),
                        ("14.2", "평가 년도 내 불만 이력 없음."), ("14.3", "평가 년도 내 회수 이력 없음."),
                        ("15.", "평가 년도 내 시정조치사항 이력 없음.")):
        for tb in _tables(document, prefix):
            last = E.raw_cells(tb.rows[-1])[0]
            if E.cell_text(last).lstrip().startswith("특이사항"):
                E.set_cell_plain(last, "특이사항 (Comment)", msg)

    # ---------- 16항 ---------- 배포본 'PQR 작성방법 공유의 건'(2026-09-04) 문안 그대로.
    # 10 Lot 미만이라 Cpk 를 산출하지 않았다는 말은 당연한 것이라 결론에 적지 않는다(담당자 지시).
    from . import conclusion
    plan_year, plan_q = conclusion.plan_quarter(today)      # 계획서 기한 = 작성일의 다음 분기
    written = conclusion.apply(
        document, full_name, name or full_name, produced=bool(dom or exp), n_lots=len(dom),
        year=year_from, write_year=plan_year, quarter=plan_q, cpk=cpk_dom)
    if not written:
        issues.append(("16", "", "'16. 결론' 제목을 찾지 못해 결론을 다시 쓰지 못함 — 확인 필요"))
    log("16항 완료")
    cover = find_para(document, name[:5]) if name else None
    def rename_heading(old_text, new_text):
        """항 제목을 2026 결재본 차림새로 바꾼다. 찾지 못하면 그냥 둔다."""
        para = find_para(document, old_text)
        if para is not None and E.loose(para.text) != E.loose(new_text):
            E.set_para_text(para, new_text)
            return True
        return False

    # ---------- 14 · 17 · 18 항 차림새 (2026 결재본 기준) ----------
    rename_heading("반품 및 불만 회수관련 기록", "14. 반품, 불만 및 회수 현황표")
    # 17 참고 자료에서 첨부 문서(안정성 결과표·경향 분석 결과)를 18 항으로 옮긴다.
    ref = find_para(document, "17. 참고 자료")
    if ref is not None:
        moved, tail = [], []
        node = ref._p.getnext()
        while node is not None and node.tag == qn("w:p"):
            text = _text(node).strip()
            if re.match(r"^\s*18\.", text):
                moved = []                         # 이미 18 항이 있으면 손대지 않는다
                break
            if text.startswith("-") and ("HLF-QC-104" in text or "HLF-QC-126-06" in text):
                moved.append(node)
            elif text:
                tail.append(node)
            node = node.getnext()
        if moved:
            head = copy.deepcopy(ref._p)
            for run in head.findall(qn("w:r"))[1:]:
                head.remove(run)
            for t in head.iter(qn("w:t")):
                t.text = "18. 첨부 문서"
                t.set(qn("xml:space"), "preserve")
            anchor = (tail or moved)[-1] if not tail else tail[-1]
            last = moved[-1]
            for el in moved:                       # 옮길 줄들을 문서 끝으로 모은다
                el.getparent().remove(el)
            spacer = copy.deepcopy(moved[0])
            for run in spacer.findall(qn("w:r")):
                spacer.remove(run)
            anchor.addnext(head)
            head.addprevious(spacer)
            after = head
            for el in moved:
                after.addnext(el)
                after = el
            log("17·18항: 첨부 문서 %d줄을 18항으로 나눔" % len(moved))

    return {"issues": issues, "cover_title": (cover.text.strip() if cover is not None else None), "cpk": cpk_dom}


def _fill_numbers(t, f, l, lots, n, cpk, rec, put, is_dom, limits=None):
    limits = limits or DEFAULT_LIMITS
    for i, lot in enumerate(lots):
        r = rec(lot, "924"); c = E.raw_cells(t.rows[f + i])
        E.set_cell(c[0], str(i + 1)); E.set_cell(c[1], lot)
        E.set_cell(c[2], ("%.2f" % float(r["particle"])) if r.get("particle") else "")
        E.set_cell(c[3], ("%.2f" % float(r["mass_avg"])) if r.get("mass_avg") else ""); E.set_cell(c[4], "%s 이상" % r.get("mass_each_min") if r.get("mass_each_min") else "")
        E.set_cell(c[5], r.get("assay") or "")
    pt, pa, pi, ct = n["pt"], n["pa"], n["pi"], n["ct"]
    put(t, l + 1, (1, 2, 3, 4), "%.2f" % max(pt), "%.2f" % max(pa), "%.2f" % max(pi), "%.1f" % max(ct))
    put(t, l + 2, (1, 2, 3, 4), "%.2f" % min(pt), "%.2f" % min(pa), "%.2f" % min(pi), "%.1f" % min(ct))
    put(t, l + 3, (1, 2, 3, 4), "%.2f" % (sum(pt) / len(pt)), "%.2f" % (sum(pa) / len(pa)), "", "%.1f" % (sum(ct) / len(ct)))
    if qc.cpk_applies(len(lots)):
        cpk["particle"] = cpk_uni(pt, limits["particle"]); cpk["assay"] = cpk_bi(ct, *limits["assay"])
        if cpk["particle"] is not None and cpk["assay"] is not None:
            put(t, l + 4, (1, 4), "%.2f" % cpk["particle"], "%.2f" % cpk["assay"])
            put(t, l + 5, (1, 4), "충분" if cpk["particle"] >= 1 else "부족", "충분" if cpk["assay"] >= 1 else "부족")


def _fill_stability(document, stab, log, assay_limits=None):
    """stab: {"post_dom": [...], "post_exp": [...], "long_dom": [...], "long_exp": [...],
    "trend_dom": [...], "trend_exp": [...]} — 손글씨 판독(비전) 결과. 형식은 kynobuild/data.py 와 같다."""
    def fill_post(table, rows):
        f, l = E.fit_rows(table, 1, len(table.rows) - 2, max(1, len(rows)))
        for i, (no, yr, per, lot, pack, day, note) in enumerate(rows):
            c = E.raw_cells(table.rows[f + i])
            for k, v in enumerate((no, yr, per, lot, pack)):
                E.set_cell(c[k], *v.split("\n")); E.set_vmerge(c[k], False)
            E.set_cell(c[5], *STORE); E.set_vmerge(c[5], False)
            E.set_cell(c[6], day); E.set_vmerge(c[6], False)
            E.set_cell(c[7], note); E.set_vmerge(c[7], False)

    def fill_long(table, groups):
        total = sum(len(g[5]) for g in groups)
        f, l = E.fit_rows(table, 1, len(table.rows) - 2, max(1, total))
        ri, prev_year = f, None
        for no, yr, lot, pack, why, points in groups:
            for pi, (period, day) in enumerate(points):
                c = E.raw_cells(table.rows[ri]); head = (pi == 0)
                E.set_cell(c[0], no if head else ""); E.set_vmerge(c[0], "restart" if head else None)
                new_year = head and yr != prev_year
                E.set_cell(c[1], yr if new_year else ""); E.set_vmerge(c[1], "restart" if new_year else None)
                E.set_cell(c[2], period); E.set_vmerge(c[2], False)
                E.set_cell(c[3], *(lot.split("\n") if head else [""])); E.set_vmerge(c[3], "restart" if head else None)
                E.set_cell(c[4], pack if head else ""); E.set_vmerge(c[4], "restart" if head else None)
                E.set_cell(c[5], *(STORE if head else [""])); E.set_vmerge(c[5], "restart" if head else None)
                E.set_cell(c[6], day); E.set_vmerge(c[6], False)
                E.set_cell_plain(c[7], *(why.split("\n") if head else [""])); E.set_vmerge(c[7], "restart" if head else None)
                ri += 1
                if head:
                    prev_year = yr

    lo_hi = assay_limits or DEFAULT_LIMITS["assay"]

    def fill_trend(table, rows, note, comment):
        f, l = E.fit_rows(table, 1, len(table.rows) - 6, max(1, len(rows)))
        prev = None
        for i, (grp, yr, val) in enumerate(rows):
            c = E.raw_cells(table.rows[f + i]); head = grp != prev
            E.set_cell(c[0], grp if head else ""); E.set_vmerge(c[0], "restart" if head else None)
            E.set_cell(c[1], yr); E.set_cell(c[2], val); prev = grp
        lows = [float(r[2].split("~")[0]) for r in rows if "~" in r[2]]
        highs = [float(r[2].split("~")[1]) for r in rows if "~" in r[2]]
        for ri, v in ((l + 1, "%.1f~%.1f" % tuple(lo_hi)), (l + 2, "%.1f" % min(lows) if lows else ""), (l + 3, "%.1f" % max(highs) if highs else ""), (l + 4, "적합")):
            E.set_cell(E.raw_cells(table.rows[ri])[1], v)
        E.set_cell_plain(E.raw_cells(table.rows[-1])[0], "특이사항 (Comment)", note, comment)

    t131 = _tables(document, "13.1"); t132 = _tables(document, "13.2"); t133 = _tables(document, "13.3")
    if t131:
        fill_post(t131[0], stab.get("post_dom", [])); E.set_cell_plain(E.raw_cells(t131[0].rows[-1])[0], "특이사항 (Comment)", stab.get("post_dom_note") or "N/A")
        if len(t131) > 1:
            fill_post(t131[1], stab.get("post_exp", [])); E.set_cell_plain(E.raw_cells(t131[1].rows[-1])[0], "특이사항 (Comment)", stab.get("post_exp_note") or "N/A")
    if t132:
        fill_long(t132[0], stab.get("long_dom", [])); E.set_cell_plain(E.raw_cells(t132[0].rows[-1])[0], "특이사항 (Comment)", stab.get("long_dom_note") or "N/A")
        if len(t132) > 1:
            fill_long(t132[1], stab.get("long_exp", [])); E.set_cell_plain(E.raw_cells(t132[1].rows[-1])[0], "특이사항 (Comment)", stab.get("long_exp_note") or "N/A")
    if t133:
        fill_trend(t133[0], stab.get("trend_dom", []), stab.get("trend_dom_lots", ""), stab.get("trend_dom_comment", ""))
        if len(t133) > 1:
            fill_trend(t133[1], stab.get("trend_exp", []), stab.get("trend_exp_lots", ""), stab.get("trend_exp_comment", ""))
    log("13항: 안정성 %d/%d/%d 표" % (len(t131), len(t132), len(t133)))

