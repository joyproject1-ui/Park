# -*- coding: utf-8 -*-
"""목차 쪽수를 Word 필드(PAGEREF)로 바꾼다 — LibreOffice 없이도 쪽수가 맞는다.

항 제목마다 책갈피를 두고, 목차 칸에는 그 책갈피의 쪽 번호 필드를 넣는다. 파일을 열 때
Word 가 필드를 다시 계산하도록 settings.xml 에 updateFields 를 켠다(polish 에서 처리).
"""
import copy
import os
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
    # dirty="true" 를 붙여야 Word 가 파일을 열 때 이 필드를 **묻지 않고 다시 계산**한다.
    # 이것이 없으면 아래 cached 값(서식에 적혀 있던 옛 쪽수)이 그대로 보인다 — 담당자 화면에서
    # 목차가 전부 '1' 로 나온 까닭이다. LibreOffice 는 열 때 늘 다시 계산해 PDF 로는 맞아 보였다.
    b = template_run.makeelement(qn("w:fldChar"),
                                 {qn("w:fldCharType"): "begin", qn("w:dirty"): "true"})
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
        # 캐시 값은 Word 가 열면서 덮어쓴다. 서식에 적혀 있던 남의 쪽수를 남겨 두면 갱신 전에
        # 엉뚱한 수가 보이므로 비워 둔다.
        cached = ""
        template = runs[0] if runs else p.makeelement(qn("w:r"), {})
        for r in p.findall(qn("w:r")):
            p.remove(r)
        for r in _field_runs(template, "PAGEREF %s \\h" % name, cached):
            p.append(r)
        n += 1
    return n


# ---------- 쪽 번호를 직접 계산해 필드 결과에 적어 둔다 ----------
def _squeeze(text):
    return re.sub(r"[\s ]+", "", text or "")


def _pdf_pages(pdf_path):
    """PDF 쪽마다의 글자 (pdfplumber — Word 가 없는 PC 에서도 돈다)."""
    import pdfplumber
    out = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            out.append(_squeeze(page.extract_text() or ""))
    return out


def _heading_keys(document):
    """{항 번호: 제목 앞머리} — 본문 제목의 첫 12자 (번호 포함) 로 쪽을 찾는다."""
    keys = {}
    for kind, value, _el in outline(document):
        if kind != "h":
            continue
        m = re.match(r"^\s*(\d{1,2})\.\s*(.+)$", value or "")
        if m and m.group(1) not in keys:
            keys[m.group(1)] = _squeeze(m.group(1) + "." + m.group(2))[:12]
    return keys


def empty_results(docx_path):
    """캐시 값이 빈 PAGEREF 필드 수 — Word 제한된 보기에서는 이 값이 그대로 보인다."""
    import zipfile
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    n = 0
    for m in re.finditer(r'PAGEREF _pqr_toc_\d+[^<]*</w:instrText>.*?fldCharType="separate"/>'
                         r'</w:r>(.*?)<w:fldChar w:fldCharType="end"', xml, re.S):
        if not re.search(r"<w:t[^>]*>\s*\d+\s*</w:t>", m.group(1)):
            n += 1
    return n


def fill_page_numbers(docx_path, log=None):
    """목차 쪽 번호를 PDF 로 직접 세어 필드 결과에 적어 둔다. 적은 수를 돌려준다.

    담당자 2026-09: "목차의 페이지 번호는 여전히 적용이 안됐어". 필드에 dirty 표시만 달아
    두면 Word 가 열 때 다시 계산하리라 기대했지만, 제한된 보기(인터넷에서 받은 파일)에서는
    계산하지 않고 빈 캐시 값을 그대로 보인다. LibreOffice 는 늘 다시 계산해 PDF 로는 맞아
    보였다 — 그래서 여기서 눈에 보이는 값을 직접 적는다. PAGEREF 필드는 그대로 두므로
    Ctrl+A → F9 로 다시 계산할 수도 있다.
    """
    import docx as docx_module
    import shutil
    import tempfile
    from . import convert
    log = log or (lambda *a: None)
    document = docx_module.Document(docx_path)
    keys = _heading_keys(document)
    if not keys:
        return 0
    work = tempfile.mkdtemp(prefix="pqr-toc-")
    try:
        pdf = os.path.join(work, "pages.pdf")
        convert.to_pdf(docx_path, pdf)
        pages = _pdf_pages(pdf)
    finally:
        shutil.rmtree(work, ignore_errors=True)
    # 목차 쪽 자체에도 제목이 다 있다 — '목차' 가 든 쪽은 건너뛴다
    found = {}
    for number, key in keys.items():
        for i, text in enumerate(pages, 1):
            if "목차" in text or "TableofContents" in text:
                continue
            if key in text:
                found[number] = i
                break
    if not found:
        return 0
    n = 0
    body = document.element.body
    for p in body.iter(qn("w:p")):
        runs = p.findall(qn("w:r"))
        instr = next((r for r in runs if r.find(qn("w:instrText")) is not None), None)
        if instr is None:
            continue
        m = re.search(r"PAGEREF _pqr_toc_(\d+)", instr.find(qn("w:instrText")).text or "")
        if not m or m.group(1) not in found:
            continue
        # [begin][instr][separate][cached][end] — separate 다음 런의 w:t 가 캐시 값
        after = False
        for r in runs:
            fc = r.find(qn("w:fldChar"))
            if fc is not None and fc.get(qn("w:fldCharType")) == "separate":
                after = True
                continue
            if fc is not None and fc.get(qn("w:fldCharType")) == "end":
                break
            if after:
                t = r.find(qn("w:t"))
                if t is not None:
                    t.text = str(found[m.group(1)])
                    n += 1
                    break
        for r in runs:                       # dirty 표시를 떼야 Word 가 캐시 값을 그대로 보인다
            fc = r.find(qn("w:fldChar"))
            if fc is not None and fc.get(qn("w:dirty")):
                del fc.attrib[qn("w:dirty")]
    if n:
        document.save(docx_path)
        log("목차 쪽 번호 %d개를 PDF 에서 세어 적었습니다" % n)
    return n
