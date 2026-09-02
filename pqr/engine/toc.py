# -*- coding: utf-8 -*-
"""목차 쪽수를 Word 필드(PAGEREF)로 바꾼다 — LibreOffice 없이도 쪽수가 맞는다.

항 제목마다 책갈피를 두고, 목차 칸에는 그 책갈피의 쪽 번호 필드를 넣는다. 파일을 열 때
Word 가 필드를 다시 계산하도록 settings.xml 에 updateFields 를 켠다(polish 에서 처리).
"""
import copy
import re
from docx.oxml.ns import qn

from .locate import outline, _text
from .ooxml_order import get_or_add

_BM = [900]


def _bookmark(p, name):
    _BM[0] += 1
    bid = str(_BM[0])
    start = p.makeelement(qn("w:bookmarkStart"), {qn("w:id"): bid, qn("w:name"): name})
    end = p.makeelement(qn("w:bookmarkEnd"), {qn("w:id"): bid})
    ppr = p.find(qn("w:pPr"))
    (ppr.addnext(start) if ppr is not None else p.insert(0, start))
    p.append(end)


def _field_runs(template_run, instr, cached):
    """template_run 의 서식으로 [begin][instr][separate][cached][end] 런을 만든다."""
    def run(child):
        r = copy.deepcopy(template_run)
        for t in list(r):
            if t.tag != qn("w:rPr"):
                r.remove(t)
        r.append(child)
        return r
    b = template_run.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "begin"})
    it = template_run.makeelement(qn("w:instrText"), {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}); it.text = " %s " % instr
    s = template_run.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "separate"})
    t = template_run.makeelement(qn("w:t"), {}); t.text = str(cached)
    e = template_run.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "end"})
    return [run(b), run(it), run(s), run(t), run(e)]


def link_toc(document, toc_table):
    """목차 표의 각 행(제목 | 쪽수)을 본문 제목의 PAGEREF 필드로 잇는다. 이은 행 수를 돌려준다."""
    heads = [(re.match(r"^(\d{1,2})\.", v), el) for k, v, el in outline(document) if k == "h"]
    first_of = {}
    for m, el in heads:
        if m and m.group(1) not in first_of:
            first_of[m.group(1)] = el
    # 목차 표 위쪽의 제목(표지·개정내역)은 건너뛰기 위해 본문에서 '1.' 이 처음 나온 뒤만 쓴다
    n = 0
    for row in toc_table.rows:
        cells = row._tr.findall(qn("w:tc"))
        if len(cells) < 2:
            continue
        label = _text(cells[0])
        m = re.match(r"^\s*(\d{1,2})\.", label)
        if not m or m.group(1) not in first_of:
            continue
        target = first_of[m.group(1)]
        name = "_pqr_toc_%s" % m.group(1)
        if target.find(qn("w:bookmarkStart")) is None or not any(
                b.get(qn("w:name")) == name for b in target.findall(qn("w:bookmarkStart"))):
            _bookmark(target, name)
        cell = cells[-1]
        p = cell.find(qn("w:p"))
        runs = [r for r in p.findall(qn("w:r")) if r.find(qn("w:t")) is not None]
        cached = _text(cell).strip() or "0"
        template = runs[0] if runs else p.makeelement(qn("w:r"), {})
        for r in p.findall(qn("w:r")):
            p.remove(r)
        for r in _field_runs(template, "PAGEREF %s \\h" % name, cached):
            p.append(r)
        n += 1
    return n
