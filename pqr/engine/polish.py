# -*- coding: utf-8 -*-
"""LibreOffice 변환 흔적을 지워 결재본(원본 .doc)과 같이 보이게 다듬는다.

 1. 글꼴 이름  "굴림;Gulim" → "굴림"   (Word 는 세미콜론이 든 이름을 못 찾아 다른 글꼴로 대체함)
 2. LibreOffice 대체 글꼴  Liberation Serif/Sans, FreeSans … → 원래 글꼴
 3. 한글·숫자 사이 자동 간격(autoSpaceDE/DN) 을 꺼서 "품질보증 1 팀" 처럼 벌어지지 않게 함
 4. Strict OOXML 속성(w:start / w:end / jc="start") 을 Word 가 쓰는 이름으로 되돌림
"""
import re, shutil, sys, zipfile
from lxml import etree
from .ooxml_order import resort, _SEQ

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def fix_order(xml):
    """pPr·rPr·tcPr·trPr·tcBorders 안 자식 순서를 스키마대로 바로잡는다.
    순서가 어긋나면 Word 는 그 속성(정렬·사선·줄간격·윗첨자)을 통째로 무시한다."""
    root = etree.fromstring(xml.encode("utf-8"))
    n = 0
    for tag in _SEQ:
        for el in root.iter("{%s}%s" % (W, tag)):
            if resort(el):
                n += 1
    out = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True).decode("utf-8")
    return out, n


def check_order(xml):
    """어긋난 곳 개수(0 이어야 함)."""
    root = etree.fromstring(xml.encode("utf-8"))
    bad = 0
    for tag, seq in _SEQ.items():
        for el in root.iter("{%s}%s" % (W, tag)):
            ranks = [seq.index(k.tag.split("}")[-1]) for k in el if k.tag.split("}")[-1] in seq]
            if ranks != sorted(ranks):
                bad += 1
    return bad

FONT_MAP = {
    "Liberation Serif": "Times New Roman",
    "Liberation Sans": "Arial",
    "FreeSans": "Arial",
    "WenQuanYi Zen Hei": "굴림",
    "¢®Ii¢®E¢®©\xad¢®E?o": "Times New Roman",
}
FONT_ATTRS = ("ascii", "hAnsi", "eastAsia", "cs")


def fix_font_name(name):
    if ";" in name:
        head, tail = name.split(";", 1)
        # 앞쪽이 깨진 이름이면 뒤쪽(영문 이름)을 쓴다
        name = tail if re.search(r"[¢®?]", head) else head
    return FONT_MAP.get(name, name)


def fix_fonts(xml):
    def sub(m):
        return f'w:{m.group(1)}="{fix_font_name(m.group(2))}"'
    return re.sub(r'w:(ascii|hAnsi|eastAsia|cs)="([^"]*)"', sub, xml)


def fix_font_table(xml):
    def sub(m):
        return f'w:name="{fix_font_name(m.group(1))}"'
    return re.sub(r'w:name="([^"]*)"', sub, xml)


def fix_autospace(xml):
    xml = xml.replace('<w:autoSpaceDE w:val="true"/>', '<w:autoSpaceDE w:val="false"/>')
    xml = xml.replace('<w:autoSpaceDN w:val="true"/>', '<w:autoSpaceDN w:val="false"/>')
    return xml


def fix_strict(xml):
    xml = xml.replace('<w:jc w:val="start"/>', '<w:jc w:val="left"/>')
    xml = xml.replace('<w:jc w:val="end"/>', '<w:jc w:val="right"/>')
    def ind(m):
        attrs = dict(re.findall(r'(w:\w+)="([^"]*)"', m.group(2)))
        for a, b in (("w:start", "w:left"), ("w:end", "w:right"),
                     ("w:startChars", "w:leftChars"), ("w:endChars", "w:rightChars")):
            if a in attrs:
                attrs.setdefault(b, attrs[a])      # 이미 left 가 있으면 그것을 우선
                del attrs[a]
        body = "".join(f' {k}="{v}"' for k, v in attrs.items())
        return "<w:%s%s%s>" % (m.group(1), body, "/" if m.group(2).rstrip().endswith("/") else "")
    xml = re.sub(r'<w:(ind|tblInd)\b([^>]*)>', ind, xml)
    # 표 여백 <w:tblCellMar><w:start .../><w:end .../></w:tblCellMar>
    def cellmar(m):
        body = m.group(1).replace('<w:start ', '<w:left ').replace('<w:end ', '<w:right ')
        return "<w:tblCellMar>%s</w:tblCellMar>" % body
    xml = re.sub(r'<w:tblCellMar>(.*?)</w:tblCellMar>', cellmar, xml, flags=re.S)
    def cellmar2(m):
        body = m.group(1).replace('<w:start ', '<w:left ').replace('<w:end ', '<w:right ')
        return "<w:tcMar>%s</w:tcMar>" % body
    xml = re.sub(r'<w:tcMar>(.*?)</w:tcMar>', cellmar2, xml, flags=re.S)
    return xml


GULIM = '<w:rFonts w:ascii="굴림" w:hAnsi="굴림" w:eastAsia="굴림" w:cs="굴림"/>'


def fix_missing_fonts(xml):
    """글꼴이 지정되지 않은 런에 굴림을 박는다.

    글꼴을 비워 두면 문서 기본값(Times New Roman)이 적용돼 영문·숫자만
    다른 글꼴로 보인다(예: 'R-2025-05-08-0089').
    """
    def run(m):
        body = m.group(1)
        if "<w:rFonts" in body:
            return m.group(0)
        if "<w:rPr>" in body:
            body = body.replace("<w:rPr>", "<w:rPr>" + GULIM, 1)
        elif "<w:rPr/>" in body:
            body = body.replace("<w:rPr/>", "<w:rPr>" + GULIM + "</w:rPr>", 1)
        else:
            body = "<w:rPr>" + GULIM + "</w:rPr>" + body
        return "<w:r>" + body + "</w:r>"
    return re.sub(r"<w:r>(.*?)</w:r>", run, xml, flags=re.S)


def fix_default_font(xml):
    """docDefaults 의 기본 글꼴도 굴림으로 (영문 기본값이 Times New Roman 이라서)."""
    def fonts(m):                      # <w:rFonts .../> 만 바꾼다 (<w:lang> 은 건드리지 않음)
        return ('<w:rFonts w:ascii="굴림" w:hAnsi="굴림" '
                'w:eastAsia="굴림" w:cs="굴림"/>')

    def sub(m):
        return re.sub(r'<w:rFonts[^/>]*/>', fonts, m.group(0))
    return re.sub(r'<w:rPrDefault>.*?</w:rPrDefault>', sub, xml, flags=re.S)


def fix_footer(xml):
    """바닥글: '한림제약' 은 왼쪽 끝에, 문서번호는 오른쪽 끝에 붙인다.

    변환된 파일은 문단 전체가 오른쪽 맞춤이고 가운데를 공백으로 벌려 놓아,
    Word 에서 '한림제약' 이 왼쪽에 붙지 않는다. 왼쪽 맞춤으로 바꾸고
    공백 대신 오른쪽 탭을 써서 결재본과 같은 자리에 오게 한다.
    """
    if "한림제약" not in xml:
        return xml
    # 문단 맞춤: 오른쪽 → 왼쪽
    xml = xml.replace('<w:jc w:val="right"/>', '<w:jc w:val="left"/>')
    # 오른쪽 탭 추가 (표 칸 너비 10090 - 좌우 여백 99*2)
    tabs = '<w:tabs><w:tab w:val="right" w:leader="none" w:pos="9892"/></w:tabs>'
    if "<w:tabs>" not in xml:
        xml = xml.replace('<w:pStyle w:val="Style12"/>', '<w:pStyle w:val="Style12"/>' + tabs, 1)
    # '한림제약' 뒤의 공백 꼬리를 없앤다
    xml = re.sub(r'(<w:t xml:space="preserve">한림제약)\s+(</w:t>)', r'\1\2', xml)
    # 공백만 든 런은 첫 개를 탭으로 바꾸고 나머지는 지운다
    blank_run = re.compile(r'<w:r>(?:(?!</w:r>).)*?<w:t xml:space="preserve">\s+</w:t></w:r>', re.S)
    seen = {"n": 0}

    def swap(m):
        seen["n"] += 1
        if seen["n"] == 1:
            return re.sub(r'<w:t xml:space="preserve">\s+</w:t>', '<w:tab/>', m.group(0))
        return ""
    xml = blank_run.sub(swap, xml)
    return xml


DEFAULT_SPACING = ('<w:autoSpaceDE w:val="false"/><w:autoSpaceDN w:val="false"/>')


def add_doc_defaults(xml):
    """docDefaults 의 문단 기본값에 자동 간격 끄기를 넣는다."""
    m = re.search(r'(<w:pPrDefault>\s*<w:pPr>)', xml)
    if m:
        return xml[:m.end()] + DEFAULT_SPACING + xml[m.end():]
    m = re.search(r'(<w:pPrDefault\s*/>)', xml)
    if m:
        return xml[:m.start()] + "<w:pPrDefault><w:pPr>" + DEFAULT_SPACING + "</w:pPr></w:pPrDefault>" + xml[m.end():]
    m = re.search(r'(<w:docDefaults>)', xml)
    if m:
        return xml[:m.end()] + "<w:pPrDefault><w:pPr>" + DEFAULT_SPACING + "</w:pPr></w:pPrDefault>" + xml[m.end():]
    return xml


def set_update_fields(xml):
    """settings.xml 에 <w:updateFields w:val="true"/> 를 스키마가 정한 자리에 넣는다."""
    from lxml import etree
    from .ooxml_order import SETTINGS
    root = etree.fromstring(xml.encode("utf-8") if isinstance(xml, str) else xml)
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for el in root:
        if el.tag == W + "updateFields":
            el.set(W + "val", "true")
            break
    else:
        el = etree.Element(W + "updateFields")
        el.set(W + "val", "true")
        rank = {n: i for i, n in enumerate(SETTINGS)}
        mine = rank["updateFields"]
        for sib in root:
            if rank.get(etree.QName(sib).localname, 999) > mine:
                sib.addprevious(el)
                break
        else:
            root.append(el)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True).decode("utf-8")


def clear_update_fields(xml):
    """settings.xml 에서 <w:updateFields> 를 지운다.

    이 표시가 있으면 Word 는 파일을 열자마자, 쪽 나눔이 끝나기도 전에 모든 필드를 다시
    계산한다 — 그래서 PAGEREF 가 전부 '1' 이 된다(담당자 2026-09: "페이지 번호는 여전히
    안 나와", 목차가 모두 1). 목차 쪽 번호는 만들 때 PDF 로 세어 필드 결과에 적어 두므로
    (toc.fill_page_numbers) Word 가 캐시 값을 그대로 보이게 두는 것이 맞다. Ctrl+A → F9 는
    쪽 나눔 뒤에 도는 손 갱신이라 그때는 제대로 계산된다.
    """
    from lxml import etree
    root = etree.fromstring(xml.encode("utf-8") if isinstance(xml, str) else xml)
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for el in list(root):
        if el.tag == W + "updateFields":
            root.remove(el)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + etree.tostring(root, encoding="unicode")


def polish(path):
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            name = item.filename
            if name.endswith(".xml") and name.startswith("word/"):
                xml = data.decode("utf-8")
                if name == "word/fontTable.xml":
                    xml = fix_font_table(xml)
                else:
                    xml = fix_fonts(xml)
                    xml = fix_autospace(xml)
                    xml = fix_strict(xml)
                    if name in ("word/document.xml",) or name.startswith(("word/header", "word/footer")):
                        xml = fix_missing_fonts(xml)
                    if name.startswith("word/footer"):
                        xml = fix_footer(xml)
                    if name == "word/styles.xml":
                        xml = add_doc_defaults(xml)
                        xml = fix_default_font(xml)
                    if name == "word/settings.xml":
                        # 예전엔 updateFields 를 켜 Word 가 열 때 목차를 다시 계산하게 했는데,
                        # 쪽 나눔 전에 계산해 전부 '1' 이 됐다. 이제 쪽 번호는 만들 때 직접
                        # 적으므로(toc.fill_page_numbers) 이 표시는 있으면 지운다.
                        xml = clear_update_fields(xml)
                    xml, fixed = fix_order(xml)
                    if fixed:
                        print(f"  순서 바로잡음 {name}: {fixed}")
                data = xml.encode("utf-8")
            dst.writestr(item, data)
    shutil.move(tmp, path)


if __name__ == "__main__":
    target = sys.argv[1]
    polish(target)
    with zipfile.ZipFile(target) as z:
        x = z.read("word/document.xml").decode("utf-8")
        left = sorted(set(re.findall(r'w:(?:ascii|hAnsi|eastAsia|cs)="([^"]*;[^"]*)"', x)))
        print("남은 세미콜론 글꼴:", left)
        print("굴림 지정 개수:", x.count('"굴림"'))
        print("autoSpaceDE true:", x.count('autoSpaceDE w:val="true"'))
        print('jc start/end:', x.count('w:jc w:val="start"'), x.count('w:jc w:val="end"'))
        for name in z.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                b = check_order(z.read(name).decode("utf-8"))
                if b:
                    print("!! 순서 어긋남", name, b)
        print("순서 검사 끝")
