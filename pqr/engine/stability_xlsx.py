# -*- coding: utf-8 -*-
"""HLF-QC-126-06 안정성 시험 경향 분석 결과 — 서식을 그대로 두고 값과 그래프만 갱신한다.

openpyxl 로 저장하면 서식에 들어 있는 꺾은선 그래프가 사라지므로,
xlsx(zip) 안의 XML 을 직접 손봐서 그래프를 살린 채 값만 바꾼다.
"""
from __future__ import unicode_literals

import os
import re
import zipfile

from lxml import etree

NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
C = "{%s}" % NS_C

POINTS = ["Initial", "3M", "6M", "9M", "12M", "18M", "24M", "36M", "48M", "60M"]
COLS = "CDEFGHIJKL"
FIRST_ROW = 38            # 제조번호·결과 첫 행
ROWS = 30                 # 서식이 가진 행 수
MIRROR_ROW = 72           # 그래프가 참조하는 미러 행 (=B38 …)
LCL_ROW, UCL_ROW = 103, 104
DEFAULT_POINTS = 8        # 그래프에 기본으로 그리는 시점 수 (Initial ~ 36M)
LOT_STYLE = 5             # 제조번호 칸 서식
VALUE_STYLE = 6           # 결과값 칸 서식

COLORS = ['<a:schemeClr val="accent1"/>', '<a:schemeClr val="accent2"/>',
          '<a:schemeClr val="accent3"/>', '<a:schemeClr val="accent4"/>',
          '<a:schemeClr val="accent5"/>', '<a:schemeClr val="accent6"/>',
          '<a:schemeClr val="accent1"><a:lumMod val="60000"/></a:schemeClr>',
          '<a:schemeClr val="accent2"><a:lumMod val="60000"/></a:schemeClr>',
          '<a:schemeClr val="accent3"><a:lumMod val="60000"/></a:schemeClr>',
          '<a:schemeClr val="accent4"><a:lumMod val="60000"/></a:schemeClr>',
          '<a:schemeClr val="accent5"><a:lumMod val="60000"/></a:schemeClr>',
          '<a:schemeClr val="accent6"><a:lumMod val="60000"/></a:schemeClr>']


def _esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def set_cell(xml, ref, value, style=None):
    """sheet XML 문자열의 한 칸을 바꾼다. 없는 칸이면 그대로 둔다."""
    m = re.search(r'<c r="%s"(?P<attrs>[^>]*?)(?P<end>/>|>)' % re.escape(ref), xml)
    if not m:
        return xml
    start = m.start()
    end = m.end() if m.group("end") == "/>" else xml.index("</c>", m.end()) + 4
    s = re.search(r'\bs="(\d+)"', m.group("attrs"))
    sid = style if style is not None else (s.group(1) if s else None)
    sattr = ' s="%s"' % sid if sid is not None else ""
    if value is None or value == "":
        new = '<c r="%s"%s/>' % (ref, sattr)
    elif isinstance(value, (int, float)):
        new = '<c r="%s"%s><v>%s</v></c>' % (ref, sattr, value)
    else:
        new = ('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
               % (ref, sattr, _esc(value)))
    return xml[:start] + new + xml[end:]


_CELL = r'<c r="%s"(?P<attrs>[^>]*?)(?P<end>/>|>)'


def set_formula_cache(xml, ref, value):
    """수식은 그대로 두고 계산 결과(캐시값)만 바꾼다.

    72~101 행은 38~67 행을 그대로 비추는 수식이라 값 자체는 손대지 않지만,
    캐시값을 갱신하지 않으면 엑셀이 다시 계산하기 전까지 그래프가 지난해 값을 그린다.
    """
    m = re.search(_CELL % re.escape(ref), xml)
    if not m or m.group("end") == "/>":
        return xml
    start, end = m.start(), xml.index("</c>", m.end()) + 4
    body = xml[m.end():end - 4]
    f = re.search(r"<f[^>]*/>|<f[^>]*>.*?</f>", body, re.S)
    if not f:
        return xml
    s = re.search(r'\bs="(\d+)"', m.group("attrs"))
    sattr = ' s="%s"' % s.group(1) if s else ""
    if value is None:
        tattr, cached = ' t="e"', "#N/A"
    elif isinstance(value, (int, float)):
        tattr, cached = "", value
    else:
        tattr, cached = ' t="str"', _esc(value)
    new = '<c r="%s"%s%s>%s<v>%s</v></c>' % (ref, sattr, tattr, f.group(0), cached)
    return xml[:start] + new + xml[end:]


def strip_floating_lines(xml_bytes):
    """서식에 떠 있는 사선(직선 연결선) 도형을 모두 지운다.

    Lot 수가 달라지면 이 도형들이 자료 위를 가로지른다. 빈 칸 사선은 칸마다
    테두리로 넣으므로(GMP 공란 없음) 도형은 필요 없다. 차트와 로고는 남긴다.
    """
    xml = xml_bytes.decode("utf-8")
    out, i, removed = [], 0, 0
    for m in re.finditer(r"<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>|"
                         r"<xdr:oneCellAnchor.*?</xdr:oneCellAnchor>|"
                         r"<xdr:absoluteAnchor.*?</xdr:absoluteAnchor>", xml, re.S):
        if "<xdr:cxnSp" in m.group(0) or 'prst="line"' in m.group(0):
            out.append(xml[i:m.start()])
            i = m.end()
            removed += 1
    out.append(xml[i:])
    return "".join(out).encode("utf-8"), removed


def diagonal_styles(styles_xml, sources):
    """빈 칸에 쓸 '사선 테두리' 서식을 찾거나 만든다. {본래 s: 사선 s} 를 돌려준다."""
    xml = styles_xml.decode("utf-8")
    bs = re.search(r'<borders count="(\d+)">(.*?)</borders>', xml, re.S)
    borders = re.findall(r"<border(?: [^>]*)?>.*?</border>|<border[^>]*/>", bs.group(2), re.S)
    diag = next((i for i, b in enumerate(borders) if 'diagonalUp="1"' in b), None)

    cs = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', xml, re.S)
    xfs = re.findall(r"<xf [^>]*/>|<xf [^>]*>.*?</xf>", cs.group(2), re.S)

    if diag is None:                                   # 서식에 없으면 첫 원본을 본떠 만든다
        src = re.search(r'borderId="(\d+)"', xfs[sources[0]])
        base = borders[int(src.group(1))] if src else borders[0]
        new = base.replace("<border>", '<border diagonalUp="1">').replace(
            "<diagonal/>", '<diagonal style="thin"><color indexed="64"/></diagonal>')
        borders.append(new)
        diag = len(borders) - 1
        xml = xml.replace(bs.group(0), '<borders count="%d">%s</borders>'
                          % (len(borders), "".join(borders)), 1)

    mapping, added = {}, False
    for sid in sources:
        want = re.sub(r'borderId="\d+"', 'borderId="%d"' % diag, xfs[sid])
        found = next((i for i, x in enumerate(xfs) if x == want), None)
        if found is None:
            xfs.append(want)
            added = True
            found = len(xfs) - 1
        mapping[sid] = found
    if added:
        xml = xml.replace(cs.group(0), '<cellXfs count="%d">%s</cellXfs>'
                          % (len(xfs), "".join(xfs)), 1)
    return xml.encode("utf-8"), mapping


def points_shown(lots):
    """그래프에 그릴 시점 수 — 기본 36M 까지, 그 뒤 결과가 있으면 거기까지 늘린다."""
    last = DEFAULT_POINTS
    for _, vals in lots:
        for k, p in enumerate(POINTS):
            if vals.get(p) is not None:
                last = max(last, k + 1)
    return last


def _cat(sheet, n):
    pts = "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, p)
                  for i, p in enumerate(POINTS[:n]))
    return ('<c:cat><c:strRef><c:f>%s!$C$71:$%s$71</c:f><c:strCache>'
            '<c:ptCount val="%d"/>%s</c:strCache></c:strRef></c:cat>'
            % (sheet, COLS[n - 1], n, pts))


def _ser(sheet, i, row, name, values, n):
    color = COLORS[i % len(COLORS)]
    pts = []
    for k, p in enumerate(POINTS[:n]):
        v = values.get(p)
        pts.append('<c:pt idx="%d"><c:v>%s</c:v></c:pt>'
                   % (k, "#N/A" if not isinstance(v, (int, float)) else v))
    return (
        '<c:ser><c:idx val="%d"/><c:order val="%d"/>'
        '<c:tx><c:strRef><c:f>%s!$B$%d</c:f><c:strCache><c:ptCount val="1"/>'
        '<c:pt idx="0"><c:v>%s</c:v></c:pt></c:strCache></c:strRef></c:tx>'
        '<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill>%s</a:solidFill><a:round/></a:ln>'
        '<a:effectLst/></c:spPr>'
        '<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>'
        '<a:solidFill>%s</a:solidFill><a:ln w="9525"><a:solidFill>%s</a:solidFill></a:ln>'
        '<a:effectLst/></c:spPr></c:marker>%s'
        '<c:val><c:numRef><c:f>%s!$C$%d:$%s$%d</c:f><c:numCache>'
        '<c:formatCode>0.0_);[Red]\\(0.0\\)</c:formatCode><c:ptCount val="%d"/>%s'
        '</c:numCache></c:numRef></c:val><c:smooth val="0"/></c:ser>'
        % (i, i, sheet, row, _esc(name), color, color, color, _cat(sheet, n),
           sheet, row, COLS[n - 1], row, n, "".join(pts)))


def _limit(sheet, idx, row, name, value, rgb, n):
    pts = "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (k, value) for k in range(n))
    return (
        '<c:ser><c:idx val="%d"/><c:order val="%d"/>'
        '<c:tx><c:strRef><c:f>%s!$B$%d</c:f><c:strCache><c:ptCount val="1"/>'
        '<c:pt idx="0"><c:v>%s</c:v></c:pt></c:strCache></c:strRef></c:tx>'
        '<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:prstDash val="sysDash"/><a:round/></a:ln><a:effectLst/></c:spPr>'
        '<c:marker><c:symbol val="none"/></c:marker>%s'
        '<c:val><c:numRef><c:f>%s!$C$%d:$%s$%d</c:f><c:numCache>'
        '<c:formatCode>0;\\-0;;@</c:formatCode><c:ptCount val="%d"/>%s'
        '</c:numCache></c:numRef></c:val><c:smooth val="0"/></c:ser>'
        % (idx, idx, sheet, row, _esc(name), rgb, _cat(sheet, n),
           sheet, row, COLS[n - 1], row, n, pts))


def _rebuild_chart(xml_bytes, sheet, lots, lcl, ucl):
    root = etree.fromstring(xml_bytes)
    line = root.find(".//" + C + "plotArea/" + C + "lineChart")
    if line is None:
        return xml_bytes
    sers = line.findall(C + "ser")
    if not sers:
        return xml_bytes
    pos = list(line).index(sers[0])
    for s in sers:
        line.remove(s)
    for ext in line.findall(C + "extLst"):      # 숨겨 둔 계열 목록은 버린다
        line.remove(ext)
    n = points_shown(lots)
    blocks = [_ser(sheet, i, MIRROR_ROW + i, name, vals, n)
              for i, (name, vals) in enumerate(lots[:ROWS])]
    blocks.append(_limit(sheet, 30, LCL_ROW, "하한관리\n기준(LCL)", lcl, "C00000", n))
    blocks.append(_limit(sheet, 31, UCL_ROW, "상한관리\n기준(UCL)", ucl, "0000FF", n))
    for off, b in enumerate(blocks):
        el = etree.fromstring('<root xmlns:c="%s" xmlns:a="%s">%s</root>'
                              % (NS_C, NS_A, b))[0]
        line.insert(pos + off, el)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _sheet_part(data, sheet_name):
    """시트 이름 → xl/worksheets/sheetN.xml 경로. 못 찾으면 첫 시트."""
    wb = etree.fromstring(data["xl/workbook.xml"])
    rels = etree.fromstring(data["xl/_rels/workbook.xml.rels"])
    rid = {r.get("Id"): r.get("Target") for r in rels}
    first = None
    for sh in wb.iter("{%s}sheet" % NS_S):
        target = rid.get(sh.get("{%s}id" % NS_R), "")
        part = "xl/" + target.lstrip("/").replace("worksheets/", "worksheets/")
        part = part if part.startswith("xl/") else "xl/" + part
        first = first or (sh.get("name"), part)
        if sh.get("name") == sheet_name:
            return sh.get("name"), part
    return first or (sheet_name, "xl/worksheets/sheet1.xml")


def _drawing_part(data, sheet_part):
    rel = sheet_part.replace("worksheets/", "worksheets/_rels/") + ".rels"
    if rel not in data:
        return None
    for r in etree.fromstring(data[rel]):
        if r.get("Type", "").endswith("/drawing"):
            return os.path.normpath(os.path.join(os.path.dirname(sheet_part),
                                                 r.get("Target"))).replace("\\", "/")
    return None


def _chart_part(data, sheet_part):
    rel = sheet_part.replace("worksheets/", "worksheets/_rels/") + ".rels"
    if rel not in data:
        return None
    for r in etree.fromstring(data[rel]):
        if r.get("Type", "").endswith("/drawing"):
            drawing = os.path.normpath(os.path.join(os.path.dirname(sheet_part),
                                                    r.get("Target"))).replace("\\", "/")
            drel = drawing.replace("drawings/", "drawings/_rels/") + ".rels"
            if drel in data:
                for d in etree.fromstring(data[drel]):
                    if d.get("Type", "").endswith("/chart"):
                        return os.path.normpath(
                            os.path.join(os.path.dirname(drawing),
                                         d.get("Target"))).replace("\\", "/")
    return None


def build(form, out, product, lots, item="함량(%)",
          storage="25±2℃\n60±5%RH", lcl=90, ucl=110, remark="N/A",
          prepared_by="", prepared_on="", sheet_name="함량"):
    """서식(form)을 복제해 lots 를 채운 파일을 out 에 만든다.

    lots: [(제조번호, {"Initial": 100.5, "12M": 97.5, ...}), ...]
    작성자·작성일은 보고서 본문의 작성자·작성일자와 같게 넣는다(기본값은 빈칸).
    """
    zin = zipfile.ZipFile(form)
    names = zin.namelist()
    data = {n: zin.read(n) for n in names}
    zin.close()

    sheet, part = _sheet_part(data, sheet_name)
    xml = data[part].decode("utf-8")
    xml = set_cell(xml, "C3", product)
    xml = set_cell(xml, "G3", item)
    xml = set_cell(xml, "C4", storage)
    xml = set_cell(xml, "H4", lcl)
    xml = set_cell(xml, "J4", ucl)
    xml = set_cell(xml, "M3", prepared_by)
    xml = set_cell(xml, "M4", prepared_on)
    xml = set_cell(xml, "A27", remark)
    data["xl/styles.xml"], diag = diagonal_styles(data["xl/styles.xml"],
                                                  (LOT_STYLE, VALUE_STYLE))
    for i in range(ROWS):
        row = FIRST_ROW + i
        name = lots[i][0] if i < len(lots) else None
        vals = lots[i][1] if i < len(lots) else {}
        xml = set_cell(xml, "B%d" % row, name,
                       style=str(LOT_STYLE if name else diag[LOT_STYLE]))
        for col, p in zip(COLS, POINTS):               # 빈 칸에는 사선 (GMP 공란 없음)
            v = vals.get(p)
            xml = set_cell(xml, "%s%d" % (col, row), v,
                           style=str(VALUE_STYLE if v is not None else diag[VALUE_STYLE]))
    for i in range(ROWS):                              # 그래프가 보는 미러 행의 캐시값
        row = MIRROR_ROW + i
        name = lots[i][0] if i < len(lots) else None
        vals = lots[i][1] if i < len(lots) else {}
        xml = set_formula_cache(xml, "B%d" % row, name if name else 0)
        for col, p in zip(COLS, POINTS):
            xml = set_formula_cache(xml, "%s%d" % (col, row), vals.get(p))
    for col in COLS:
        xml = set_formula_cache(xml, "%s%d" % (col, LCL_ROW), lcl)
        xml = set_formula_cache(xml, "%s%d" % (col, UCL_ROW), ucl)
    data[part] = xml.encode("utf-8")

    chart = _chart_part(data, part)
    if chart and chart in data:
        data[chart] = _rebuild_chart(data[chart], sheet, lots, lcl, ucl)
    drawing = _drawing_part(data, part)
    if drawing and drawing in data:
        data[drawing], _ = strip_floating_lines(data[drawing])

    wb = data["xl/workbook.xml"].decode("utf-8")       # 열 때 수식을 다시 계산하게 한다
    if "fullCalcOnLoad" not in wb:
        wb = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wb)
        if "fullCalcOnLoad" not in wb:
            wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    data["xl/workbook.xml"] = wb.encode("utf-8")

    if "xl/calcChain.xml" in data:                      # 값이 바뀌었으니 캐시는 버린다
        del data["xl/calcChain.xml"]
        names = [n for n in names if n != "xl/calcChain.xml"]
        ct = data["[Content_Types].xml"].decode("utf-8")
        ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct)
        data["[Content_Types].xml"] = ct.encode("utf-8")

    zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    for n in names:
        zout.writestr(n, data[n])
    zout.close()
    return out
