# -*- coding: utf-8 -*-
"""OOXML 자식 요소 순서 지키기.

Word 는 pPr·rPr·tcPr 안 자식의 순서가 스키마와 다르면 그 속성을 무시한다
(파일은 열리지만 정렬·사선·줄간격이 적용되지 않는다). LibreOffice 는
순서를 따지지 않아 변환 미리보기에서는 멀쩡해 보이므로 발견하기 어렵다.
"""
from docx.oxml.ns import qn

PPR = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
       "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
       "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
       "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
       "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
       "suppressOverlap", "jc", "textDirection", "textAlignment",
       "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
       "pPrChange"]

RPR = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
       "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
       "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
       "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
       "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
       "eastAsianLayout", "specVanish", "oMath", "rPrChange"]

TCPR = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd",
        "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
        "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"]

TRPR = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter",
        "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden",
        "ins", "del", "trPrChange"]

TCBORDERS = ["top", "start", "left", "bottom", "end", "right",
             "insideH", "insideV", "tl2br", "tr2bl"]

_SEQ = {"pPr": PPR, "rPr": RPR, "tcPr": TCPR, "trPr": TRPR,
        "tcBorders": TCBORDERS}


def _local(el):
    return el.tag.split("}")[-1]


def place(parent, child):
    """child 를 parent 의 스키마 순서에 맞는 자리에 넣는다(이미 붙어 있어도 재배치)."""
    seq = _SEQ[_local(parent)]
    name = _local(child)
    if child.getparent() is not None:
        child.getparent().remove(child)
    try:
        rank = seq.index(name)
    except ValueError:
        parent.append(child)
        return child
    for sib in parent:
        try:
            r = seq.index(_local(sib))
        except ValueError:
            continue
        if r > rank:
            sib.addprevious(child)
            return child
    parent.append(child)
    return child


def get_or_add(parent, name):
    """parent 안의 w:name 을 찾고, 없으면 올바른 자리에 만들어 넣는다."""
    el = parent.find(qn("w:" + name))
    if el is None:
        el = parent.makeelement(qn("w:" + name), {})
        place(parent, el)
    return el


def resort(parent):
    """이미 어긋난 자식들을 스키마 순서대로 다시 세운다."""
    seq = _SEQ[_local(parent)]
    kids = list(parent)
    order = []
    for i, k in enumerate(kids):
        n = _local(k)
        order.append((seq.index(n) if n in seq else len(seq) + i, i, k))
    order.sort(key=lambda x: (x[0], x[1]))
    after = [k for _, _, k in order]
    if after == kids:
        return False
    for k in kids:
        parent.remove(k)
    for k in after:
        parent.append(k)
    return True


def resort_all(document):
    """문서 전체(본문·머리글은 별도)에서 어긋난 pPr·rPr·tcPr·trPr·tcBorders 를 바로잡는다."""
    n = 0
    root = document.element
    for tag in ("pPr", "rPr", "tcPr", "trPr", "tcBorders"):
        for el in root.iter(qn("w:" + tag)):
            if resort(el):
                n += 1
    return n
