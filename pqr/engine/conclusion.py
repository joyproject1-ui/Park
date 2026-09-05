# -*- coding: utf-8 -*-
"""16항 결론 — 회사 배포본 'PQR 작성방법 공유의 건'(2026-09-04) 문안 그대로.

담당자: "(10 Lot 미만이라 Cpk 를 산출하지 않았다는 말은) 당연한 거라 결론에 기재하지 않아.
이 결론에 따라 작성하면 돼." 그래서 결론은 세 가지 꼴뿐이다.

  1) Cpk 1 미만 항목이 있는 제품 — 16.1 문장 + 표(항목·판정 기준·Cpk) + 16.2 + 16.3 + 16.4
  2) 생산이력이 있는 제품(그 밖) — 종합 문장 하나
  3) 생산이력이 없는 제품 — "생산 이력이 없어 … 추후 생산 시 평가할 예정임"

문단이 **하나뿐이면 번호를 붙이지 않는다**(담당자 확인 2026-09, 2026 디겐타 결재본이 그렇다).
여럿으로 나뉠 때만 16.1~16.4 로 매긴다.

일탈·수율·안정성 이야기는 결론이 아니라 각 항의 특이사항(Comment)에 적는다.
"""
import copy
import re

from docx.oxml.ns import qn

from . import docedit as E
from .ooxml_order import get_or_add

CPK_NAMES = {"metal": "금속성이물(합계)", "particle": "입자도", "assay": "함량"}
JUDGE = ("Cpk ≥ 1 : 공정능력 충분", "1 > Cpk : 공정능력 부족")
HEAD_SHADE = "E6E6E6"


def texts(full_name, short_name, produced, n_lots, year, write_year, quarter, low_cpk):
    """결론 문단 목록. low_cpk: [(항목 이름, Cpk)] — 1 미만인 것만.

    full_name 은 머리글의 정식 제품명(성분명까지), short_name 은 계획서 이름 — 16.3 은 짧은
    이름을 쓴다(배포본 예: '큐티스점안액 공정 개선'). produced 가 False 면 생산이력 없음 문구.
    """
    if not produced:
        return ["%s은 %d년 생산 이력이 없어 평가할 수 있는 항목이 한정적이므로, 추후 생산 시 수율, "
                "주요 공정 및 제품의 시험 결과 등 품질에 대한 평가를 실시할 예정임." % (full_name, year)]
    final = ("%s에 대한 제품품질평가 결과, 출발물질, 포장자재, IPC Test 그리고 제품 시험 결과 모두 정해진 규격에 "
             "만족하며, 기준에 적합한 제품이 일관되게 제조되고 있어 표준제조공정이 적절하다고 판단됨. 이에 따라, "
             "시정 및 예방조치 또는 재밸리데이션 진행 여부 검토 시, 해당되는 사항은 없음을 확인하였음." % full_name)
    if not low_cpk:
        return [final]                 # 문단이 하나면 번호를 붙이지 않는다 (담당자 확인 2026-09)
    return [
        "16.1 본 제품 품질 평가를 통해 공정능력을 평가한 결과 Cpk 값이 1 미만인 항목이 아래 표와 같이 "
        "검토되어 해당 시험항목에 대한 검토 진행함.",
        "16.2 %d년 생산된 %dLot에 사용된 원료 검토 결과, 모든 항목에서 기준 내 적합한 결과였으며 제품 시험 "
        "검토 결과 허가 및 자가 기준 내 적합함을 확인함." % (year, n_lots),
        "16.3 QC-126 제품 품질 평가 규정에 따라 Cpk 값이 1 미만인 항목 검토되었으므로, %d년도 %d분기 내 "
        "‘공정능력지수 검토 계획서(HLF-QC-126-10)’을 작성하여, %s 공정 개선에 대해 검토하도록 하겠음."
        % (write_year, quarter, short_name),
        "16.4 " + final,
    ]


def plan_quarter(today):
    """16.3 '공정능력지수 검토 계획서' 작성 기한 — 작성일의 다음 분기. (작성년도, 분기)

    배포본 예(2026-09 작성 → '2026년도 4분기 내')와 2026 퀴노비드 결재본(9월 작성 → 4분기)이
    모두 다음 분기다. 4분기에 쓰면 이듬해 1분기.
    """
    q = (today.month - 1) // 3 + 1
    return (today.year + 1, 1) if q == 4 else (today.year, q + 1)


def low_cpk_items(cpk):
    """{'assay': 0.93, 'particle': 0.79, 'metal': 24.1} → [('함량', 0.93), ('입자도', 0.79)] (1 미만만, 표 순서)."""
    out = []
    for key in ("assay", "particle", "metal"):
        for full, v in cpk.items():
            # 주성분이 둘이면 'assay/성분' 으로 들어온다 → '함량(성분)'
            if full != key and not full.startswith(key + "/"):
                continue
            if v is not None and v < 1:
                name = CPK_NAMES[key]
                if "/" in full:
                    name = "%s(%s)" % (name, full.split("/", 1)[1])
                out.append((name, v))
    return out


# ---------- 문서에 넣기 ----------
def _heading_16(document):
    """'16. 결론' 제목 문단과 그 뒤의 16.x 문단·표를 돌려준다. (제목, [뒤따르는 요소])"""
    head = None
    for p in document.paragraphs:
        if re.match(r"^\s*16\.\s*결\s*론", p.text):
            head = p._p
            break
    if head is None:
        return None, []
    tail, el = [], head.getnext()
    while el is not None:
        if el.tag == qn("w:p"):
            t = "".join(x.text or "" for x in el.iter(qn("w:t"))).strip()
            if re.match(r"^\s*1[78]\.\s", t) or re.match(r"^\s*17\.", t):
                break
            tail.append(el)
        elif el.tag == qn("w:tbl"):
            tail.append(el)
        else:
            break
        el = el.getnext()
    return head, tail


def _template_para(tail, head):
    """16.x 문단의 서식을 물려받을 본. 없으면 제목 문단을 복제해 쓴다."""
    for el in tail:
        if el.tag == qn("w:p") and "".join(x.text or "" for x in el.iter(qn("w:t"))).strip():
            return el
    return head


def _cpk_table(document, template_tbl, rows):
    """항목 | 판정 기준 | Cpk 표. template_tbl 이 있으면 그 서식을 그대로 쓰고 행만 바꾼다."""
    if template_tbl is not None:
        tbl = copy.deepcopy(template_tbl)
        trs = tbl.findall(qn("w:tr"))
        if len(trs) >= 2:
            body = trs[1]
            for tr in trs[2:]:
                tbl.remove(tr)
            anchor = body
            for i, (name, val) in enumerate(rows):
                tr = body if i == 0 else copy.deepcopy(body)
                if i:
                    anchor.addnext(tr); anchor = tr
                tcs = tr.findall(qn("w:tc"))
                E.set_cell(E._Cell(tcs[0], None) if hasattr(E, "_Cell") else _Shim(tcs[0]), name)
                E.set_cell(_Shim(tcs[1]), *JUDGE)
                E.set_cell(_Shim(tcs[2]), "%.2f" % val)
            return tbl
    table = document.add_table(rows=1 + len(rows), cols=3)
    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    for g, w in zip(grid.findall(qn("w:gridCol")), (3300, 3300, 3300)):
        g.set(qn("w:w"), str(w))
    pr = get_or_add(tbl, "tblPr")
    borders = get_or_add(pr, "tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = get_or_add(borders, side)
        b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "4"); b.set(qn("w:space"), "0"); b.set(qn("w:color"), "000000")
    heads = ("항목", "판정 기준", "Cpk")
    for ci, cell in enumerate(table.rows[0].cells):
        E.set_cell(cell, heads[ci])
        tcpr = E._tcpr(cell._tc)
        shd = get_or_add(tcpr, "shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), HEAD_SHADE)
        for r in cell._tc.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = r.makeelement(qn("w:rPr"), {}); r.insert(0, rpr)
            get_or_add(rpr, "b")
    for ri, (name, val) in enumerate(rows, 1):
        cells = table.rows[ri].cells
        E.set_cell(cells[0], name); E.set_cell(cells[1], *JUDGE); E.set_cell(cells[2], "%.2f" % val)
    for row in table.rows:
        for cell in row.cells:
            E.set_cell_align(cell, "center"); E.set_cell_valign(cell, "center")
    E.fix_table_widths(table)
    tbl.getparent().remove(tbl)
    return tbl


class _Shim(object):
    def __init__(self, tc):
        self._tc = tc


def apply(document, full_name, short_name, produced, n_lots, year, write_year, quarter, cpk):
    """16항을 배포본 문안으로 다시 쓴다. 넣은 문단 수를 돌려준다(제목을 못 찾으면 0)."""
    head, tail = _heading_16(document)
    if head is None:
        return 0
    low = low_cpk_items(cpk or {})
    lines = texts(full_name, short_name, produced, n_lots, year, write_year, quarter, low)
    template = _template_para(tail, head)
    old_tbl = next((el for el in tail if el.tag == qn("w:tbl")), None)
    new_tbl = _cpk_table(document, old_tbl, low) if low else None
    made = []
    for txt in lines:
        p = copy.deepcopy(template)
        for pr in p.findall(qn("w:pPr")):
            for tag in ("w:pageBreakBefore",):
                el = pr.find(qn(tag))
                if el is not None:
                    pr.remove(el)
        E.set_para_text(p, txt)
        made.append(p)
    for el in tail:
        el.getparent().remove(el)
    anchor = head
    for i, p in enumerate(made):
        anchor.addnext(p); anchor = p
        if i == 0 and new_tbl is not None:        # 16.1 문장 바로 아래에 표
            anchor.addnext(new_tbl); anchor = new_tbl
    return len(made)
