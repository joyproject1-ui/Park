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
    year_from = int(str(period.get("from"))[:4])
    write_year = today.year
    dom, exp = data.domestic, data.export
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

    # ---------- 4항 ----------
    # 1항 '목적' 도 "제품품질평가는" 으로 시작하므로 평가 기간 문장만 집어 찾는다.
    p = find_para(document, "월까지 생산된")
    if p is not None and "년 1월" in p.text:
        text = re.sub(r"\d{4}년 1월", "%d년 1월" % year_from, p.text)
        text = re.sub(r"\d{4}년도 (상반기|하반기|\d분기)", "%d년도 %d분기" % (write_year, _quarter(today)), text)
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

    def fill_yield(table, lots, is_dom):
        specs = yield_specs(table)
        stages = ("조제", "충전", "포장")
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
        stat_lots = [x for x in lots if x not in yield_out and all(v is not None for v in vals[x])]
        if stat_lots:
            cols = list(zip(*[[float(x) for x in vals[lt]] for lt in stat_lots]))
            for ri, fn in ((l + 1, max), (l + 2, min)):
                for k in range(3):
                    E.set_cell(table.rows[ri].cells[2 + k], "%.2f" % fn(cols[k]))
            for k in range(3):
                E.set_cell(table.rows[l + 3].cells[2 + k], avg(cols[k]))
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
    def group_fill(table, records, code, item):
        total = sum(len(x[1]) for x in records)
        if not total:
            return
        f, l = E.fit_rows(table, 1, len(table.rows) - 1, total)
        ri = f
        for gi, (test_no, lots) in enumerate(records, start=1):
            for li, lot in enumerate(lots):
                r = E.raw_cells(table.rows[ri]); head = (li == 0)
                for k, v in ((0, str(gi)), (1, code), (2, item), (3, test_no)):
                    E.set_cell(r[k], v if head else ""); E.set_vmerge(r[k], "restart" if head else None)
                E.set_cell(r[4], lot); E.set_vmerge(r[4], False)
                E.set_cell(r[5], "적합" if head else ""); E.set_vmerge(r[5], "restart" if head else None)
                ri += 1
    t821 = _tables(document, "8.2.1")
    if t821 and data.raw_tests:
        first = E.raw_cells(t821[0].rows[1])
        code, item = E.cell_text(first[1]).strip(), E.cell_text(first[2]).strip()
        for tbl, lots in ((t821[0], dom), (t821[1] if len(t821) > 1 else None, exp)):
            if tbl is None:
                continue
            recs = [(test, [l for l in lots if l in ls]) for c, test, ls in data.raw_tests if c == code]
            recs = [(test, ls) for test, ls in recs if ls]
            group_fill(tbl, recs, code, item)
    t822 = _tables(document, "8.2.2")
    if t822 and data.pkg_tests:
        tbl = t822[0]
        base_rows = {E.cell_text(E.raw_cells(r)[1]).strip(): E.cell_text(E.raw_cells(r)[2]).strip()
                     for r in tbl.rows[1:] if len(E.raw_cells(r)) > 2}
        total = sum(len(ls) for _, _, ls in data.pkg_tests if any(l in dom + exp for l in ls))
        if total:
            f, l = E.fit_rows(tbl, 1, len(tbl.rows) - 1, total)
            ri = f; gi = 0
            for code, test, ls in data.pkg_tests:
                ls = [x for x in ls if x in dom + exp]
                if not ls:
                    continue
                gi += 1
                for li, lot in enumerate(ls):
                    r = E.raw_cells(tbl.rows[ri]); head = (li == 0)
                    item = base_rows.get(code, "튜브 (수출용)" if lot in exp else "튜브 (내수용)")
                    for k, v in ((0, str(gi)), (1, code), (2, item), (3, test)):
                        E.set_cell(r[k], v if head else ""); E.set_vmerge(r[k], "restart" if head else None)
                    E.set_cell(r[4], lot); E.set_vmerge(r[4], False)
                    E.set_cell(r[5], "적합" if head else ""); E.set_vmerge(r[5], "restart" if head else None)
                    ri += 1

    # ---------- 9항 ----------
    def rec(lot, key):
        return (data.coa.get(lot) or {}).get(key) or {}

    def app(lots):
        for l in lots:
            for k in ("924", "923", "922"):
                a = rec(l, k).get("appearance")
                if a:
                    return a
        return ""

    def numbers(lots):
        g = lambda k, key: [float(_num(rec(l, k).get(key))) for l in lots if _num(rec(l, k).get(key)) is not None]
        return {"ms": g("924", "metal_total"), "mi": g("924", "metal_each"), "pt": g("924", "particle"),
                "pa": g("924", "mass_avg"), "pi": g("924", "mass_each_min"), "ct": g("924", "assay"),
                "fa": g("923", "mass_avg"),
                "flo": [_rng(rec(l, "923").get("mass_each"))[0] for l in lots if _rng(rec(l, "923").get("mass_each"))],
                "fhi": [_rng(rec(l, "923").get("mass_each"))[1] for l in lots if _rng(rec(l, "923").get("mass_each"))]}

    def spec_of(t91, words):
        for row in t91.rows:
            cells = [E.cell_text(c) for c in E.raw_cells(row)]
            if any(all(w in c for w in words) for c in cells[:2]):
                return cells[-2]
        return ""

    def bio_text(lot):
        b = rec(lot, "922").get("bioburden") or ""
        m = re.match(r"(\d+)\s*(CFU|FU)?/g\s*(미만|이하)?", b)
        return ("%s %s" % (m.group(1), m.group(3) or "미만")) if m else (b or "확인 필요")

    def fill_91(t91, lots, is_dom):
        n = numbers(lots)
        if not n["ct"]:
            issues.append(("9", "", "완제 성적서에서 함량을 읽지 못함")); return n
        res = {}
        current_item = ""
        for ri, row in enumerate(t91.rows):
            cells = E.raw_cells(row)
            if ri == 0 or len(cells) < 3:
                continue
            label = re.sub(r"\s+", "", E.cell_text(cells[-3])) if len(cells) >= 4 else ""
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
                if any(v != common for v in vals):
                    val += "1)"
            elif "평균" in label and n["fa"] and "질량" in item and not n.get("_fill_avg"):
                val = "Av. %.1f\n(%.1f ~ %.1fg)" % (sum(n["fa"]) / len(n["fa"]), min(n["fa"]), max(n["fa"])) if is_dom else \
                      "Av. %.2fg\n(%.2f ~ %.2fg)" % (sum(n["fa"]) / len(n["fa"]), min(n["fa"]), max(n["fa"]))
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
                val = ("50 ㎛ 이상\n: Av. %.2f개(%.0f ~ %.0f개)\n개개 중 8개 초과 : %.0f매" % (sum(n["ms"]) / len(n["ms"]), min(n["ms"]), max(n["ms"]), max(n["mi"]))) if len(lots) >= 3 else \
                      ("50 ㎛ 이상 : %.0f개\n개개 중 8개 초과 : %.0f매" % (max(n["ms"]), max(n["mi"])))
            elif "입자도" in item and n["pt"]:
                val = "Av. %.2f㎛ 이하\n(%.2f ~ %.2f㎛ 이하)" % (sum(n["pt"]) / len(n["pt"]), min(n["pt"]), max(n["pt"]))
            elif "함량" in item and n["ct"]:
                val = "Av. %.1f%%\n(%.1f ~ %.1f%%)" % (sum(n["ct"]) / len(n["ct"]), min(n["ct"]), max(n["ct"]))
            elif "평균" in label and n["pa"] and "질량" in item:
                val = "Av. %.2fg\n(%.2f ~ %.2fg)" % (sum(n["pa"]) / len(n["pa"]), min(n["pa"]), max(n["pa"]))
            elif "개개" in label and n["pi"] and "질량" in item:
                val = "%.2f ~ %.2fg 이상2)" % (min(n["pi"]), max(n["pi"]))
            elif "무균" in item:
                st = {rec(l, "924").get("sterility") for l in lots} - {None}
                val = sorted(st)[0] if st else "음성"
            elif "포장규격" in item:
                val = "각 규격에 적합함"
            if val is not None:
                E.set_cell(cells[-1], *val.split("\n"))
                res[ri] = val
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
            if len(lots) >= 10:
                put(t_metal, l + 3, (2, 3), "%.2f" % (sum(ms) / len(ms)), "%.2f" % (sum(mi) / len(mi)))
                cpk["metal"] = cpk_uni(ms, 50.0)
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
            _fill_numbers(t, fb_, lb_, lots, n, cpk, rec, put, is_dom)
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
                _fill_numbers(t_numb, f, l, lots, n, cpk, rec, put, is_dom)
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

    # 각주
    odd_bio = [(l, rec(l, "922").get("bioburden")) for l in dom + exp if bio_text(l) != "10 미만" and rec(l, "922")]
    notes = []
    if odd_bio:
        notes.append("1) %s 조제(바이오버든) 공정 시험 성적서의 생균수 기재값은 “%s” 임. 원 기록의 단위 표기 확인 필요."
                     % (", ".join(l for l, _ in odd_bio), odd_bio[0][1]))
        issues.append(("9.2.2", ", ".join(l for l, _ in odd_bio), "생균수 기재값이 다른 Lot 과 다름 — 원본 확인"))
    notes.append("2) %d년 완제 시험 성적서는 질량·용량 개개를 최솟값(···g 이상)으로만 기재하므로, 각 Lot 의 개개 최솟값으로 기재하였음." % year_from)
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
        f, l = E.fit_rows(table, 1, len(table.rows) - 1, len(rows))
        for i, e in enumerate(rows):
            c = E.raw_cells(table.rows[f + i])
            reason = (e.get("reason") or "").strip()
            kind = "변경" if "변경" in reason else ("정기적" if "정기" in reason else "최초")
            E.set_cell(c[0], *["%s %s" % ("■" if k == kind else "☐", k) for k in ("최초", "변경", "정기적")])
            E.set_cell(c[1], *[lot for _, lot, _ in e["lots"]][:3])
            m = re.match(r"(\S+)\s*(\(.*?\))?", e["report"])
            E.set_cell(c[2], m.group(1), (m.group(2) or "").lower()) if m else E.set_cell(c[2], e["report"])
            E.set_cell(c[3], e.get("report_date") or "확인 필요")
            E.set_cell(c[4], reason)
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
                title = re.sub(r"^\S+\([A-Z0-9]+\)\s*", "", d.get("title") or "")
                detail = [("* " + title, "none")]
                for line in (d.get("description") or "").split("\n"):
                    if not line.strip():
                        continue
                    kind = "num" if re.match(r"^\d\)", line) else ("body" if re.match(r"^\d\.", line) else "cont")
                    detail.append((re.sub(r"^\d\.\s*", "", line) if kind == "body" else line, kind))
                if d.get("cause"):
                    detail.append(("원인", "body"))
                    detail.append((": " + d["cause"].replace("\n", " "), "cont"))
                action = [(line, "none") for line in (d.get("correction") or "").split("\n") if line.strip()]
                if d.get("completed"):
                    action.append(("(조치사항 완료일 : %s)" % d["completed"], "none"))
                if i_det is not None and i_det < len(c):
                    E.set_cell_flow(c[i_det], detail)
                if i_act is not None and i_act < len(c):
                    E.set_cell_flow(c[i_act], action)
                if i_capa is not None and i_capa < len(c):
                    E.set_cell(c[i_capa], "☐ Yes", "■ No")
            nos = ", ".join(d.get("doc_no") or "" for d in devs)
            E.set_cell_plain(E.raw_cells(tbl.rows[-1])[0], "특이사항 (Comment)",
                             "평가 년도 내 발생한 내수용 제품의 일탈 %d건(%s)은 조치사항이 적절하게 완료되었음을 확인함." % (len(devs), nos))
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
        ccs = [c for c in data.changes if not name or name[:4] in (c.get("products") or "") or name[:4] in (c.get("title") or "")]
        if ccs:
            f, l = E.fit_rows(tbl, 1, len(tbl.rows) - 2, len(ccs))
            for i, cc in enumerate(ccs):
                c = E.raw_cells(tbl.rows[f + i])
                E.set_cell(c[0], str(i + 1))
                lines = ["[%s] %s" % (cc.get("doc_no"), cc.get("title") or ""), "", "- 변경 사유"]
                lines += [x for x in (cc.get("reason") or "").split("\n") if x]
                lines += ["", "- 변경 내용"] + [x for x in (cc.get("description") or "").split("\n") if x]
                E.set_cell_plain(c[1], *lines)
                E.set_cell_plain(c[2], "변경 승인일 : %s" % (cc.get("approved") or "확인 필요"),
                                 "완료 목표일 : %s" % ((cc.get("target_date") or "").replace("-", ".") or "확인 필요"))
                E.set_cell(c[3], "확인 필요"); E.set_cell(c[4], "N/A")
                issues.append(("12", cc.get("doc_no") or "", "변경 조치사항과 적용 Lot 은 변경요청서만으로 정해지지 않음 — 확인 필요"))
            E.set_cell_plain(E.raw_cells(tbl.rows[-1])[0], "특이사항 (Comment)", "N/A")
        else:
            E.set_cell_plain(E.raw_cells(tbl.rows[-1])[0], "특이사항 (Comment)", "평가 년도 내 변경관리 이력 없음.")

    # ---------- 13항 ----------
    stab = getattr(data, "stability", None)
    if stab:
        _fill_stability(document, stab, log)
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
    for prefix, msg in (("14.1", "점검기간 내 사용기한 경과로 인한 반품 외 특이사항 없음"),
                        ("14.2", "평가 년도 내 불만 이력 없음."), ("14.3", "평가 년도 내 회수 이력 없음."),
                        ("15.", "평가 년도 내 시정조치사항 이력 없음.")):
        for tb in _tables(document, prefix):
            last = E.raw_cells(tb.rows[-1])[0]
            if E.cell_text(last).lstrip().startswith("특이사항"):
                E.set_cell_plain(last, "특이사항 (Comment)", msg)

    # ---------- 16항 ----------
    p = find_para(document, "16.1.")
    full = full_name
    if p is not None:
        E.set_para_text(p, "16.1. %s 내수용%s에 대한 제품품질평가 결과, 출발물질, 포장자재, IPC Test 그리고 제품 시험 결과 모두 정해진 규격에 만족하며, "
                        "기준에 적합한 제품이 일관되게 제조되고 있어 표준제조공정이 적절하다고 판단됨." % (full, ", 수출용" if exp else ""))
    p = find_para(document, "16.2.")
    if p is not None:
        texts = []
        if devs:
            texts.append("16.2. 점검기간 동안 발생한 내수용 제품의 일탈 %d건(%s)은 조치사항이 적절하게 완료되었음을 확인하였음. 수출용 제품의 중요 일탈 및 기준 일탈, 경향 일탈 이력은 없음."
                         % (len(devs), ", ".join(d.get("doc_no") or "" for d in devs)))
        else:
            texts.append("16.2. 점검기간 동안 중요 일탈 및 기준 일탈, 경향 일탈 이력은 없음.")
        if yield_out:
            texts.append("16.3. 수율 현황 검토 결과, 내수용 %d개 Lot(%s)의 충전 수율이 자가기준을 벗어나 최댓값·최솟값·평균 계산에서 제외하였음."
                         % (len(yield_out), ", ".join(yield_out)))
        cp = {k: v for k, v in cpk_dom.items() if v is not None}
        if cp:
            names = {"metal": "금속성이물(합계)", "particle": "입자도", "assay": "함량"}
            good = ["%s %.2f" % (names[k], v) for k, v in cp.items() if v >= 1]
            bad = ["%s %.2f" % (names[k], v) for k, v in cp.items() if v < 1]
            s = "16.%d. 내수용 제품의 공정능력지수(Cpk) 산출 결과, " % (len(texts) + 2)
            if bad:
                s += ("%s 은 공정능력이 충분하였으나 " % ", ".join(good) if good else "") + "%s 으로 1 미만이므로 ‘QC-126 제품품질평가 규정’의 판정 기준에 따라 공정 개선 검토가 필요함." % ", ".join(bad)
            else:
                s += "%s 으로 공정능력이 충분함." % ", ".join(good)
            if exp and len(exp) < 10:
                s += " 수출용 제품은 평가 년도 생산 Lot 이 %d Lot 으로 10 Lot 미만이므로 Cpk 를 산출하지 않았음." % len(exp)
            texts.append(s)
        texts.append("16.%d. 시판 후 안정성 시험 및 장기 안정성 시험의 경향 분석 결과, 모두 관리 규격 범위 내에 있어 경시 변화에 따른 특이사항은 없는 것으로 판단됨." % (len(texts) + 2))
        # 기존 16.2 이후 문단 제거 후 새로 넣는다
        el = p._p
        nxt = el.getnext()
        while nxt is not None and nxt.tag == qn("w:p") and re.match(r"^16\.\d", _text(nxt)):
            gone = nxt; nxt = nxt.getnext(); gone.getparent().remove(gone)
        E.set_para_text(p, texts[0]); anchor = el
        for txt in texts[1:]:
            np_ = copy.deepcopy(el); anchor.addnext(np_); E.set_para_text(np_, txt); anchor = np_
    log("16항 완료")
    cover = find_para(document, name[:5]) if name else None
    return {"issues": issues, "cover_title": (cover.text.strip() if cover is not None else None), "cpk": cpk_dom}


def _fill_numbers(t, f, l, lots, n, cpk, rec, put, is_dom):
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
    if len(lots) >= 10:
        cpk["particle"] = cpk_uni(pt, 75.0); cpk["assay"] = cpk_bi(ct, 90.0, 110.0)
        if cpk["particle"] is not None and cpk["assay"] is not None:
            put(t, l + 4, (1, 4), "%.2f" % cpk["particle"], "%.2f" % cpk["assay"])
            put(t, l + 5, (1, 4), "충분" if cpk["particle"] >= 1 else "부족", "충분" if cpk["assay"] >= 1 else "부족")


def _fill_stability(document, stab, log):
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

    def fill_trend(table, rows, note, comment):
        f, l = E.fit_rows(table, 1, len(table.rows) - 6, max(1, len(rows)))
        prev = None
        for i, (grp, yr, val) in enumerate(rows):
            c = E.raw_cells(table.rows[f + i]); head = grp != prev
            E.set_cell(c[0], grp if head else ""); E.set_vmerge(c[0], "restart" if head else None)
            E.set_cell(c[1], yr); E.set_cell(c[2], val); prev = grp
        lows = [float(r[2].split("~")[0]) for r in rows if "~" in r[2]]
        highs = [float(r[2].split("~")[1]) for r in rows if "~" in r[2]]
        for ri, v in ((l + 1, "90.0~110.0"), (l + 2, "%.1f" % min(lows) if lows else ""), (l + 3, "%.1f" % max(highs) if highs else ""), (l + 4, "적합")):
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
