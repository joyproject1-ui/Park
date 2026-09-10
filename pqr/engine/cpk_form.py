# -*- coding: utf-8 -*-
"""Cpk 계산 파일 서식(.xlsx)의 칸 값만 갈아 끼운다 — 그래프·로고·서식·수식은 손대지 않는다.

담당자 2026-09-06: "Cpk 는 전년도 양식으로 작성하되 2026년 PQR 작성본 내용을 참고해서 업데이트하면 돼."

openpyxl 로 열었다 저장하면 꺾은선 그래프와 한림 로고가 사라진다. 그래서 .xlsx(=zip) 안의
워크시트 XML 을 그대로 두고 칸 하나하나만 고친다. 수식 칸은 수식을 그대로 두고 계산값만
새로 넣어(Excel 이 '제한된 보기' 에서 다시 계산하지 않아도 값이 보이게) 둔다.

fill(form, dst, cells, values) — form: 서식 .xlsx, cells: {"C4": 제품명, …}, values: 결과값 목록
"""
import os
import re
import shutil
import zipfile

from . import cpk_xlsx

FIRST_DATA_ROW, ROWS = 10, 35
DATA_COL = "B"


def _letters(ref):
    return re.match(r"([A-Z]+)", ref).group(1)


def _rownum(ref):
    return int(re.search(r"(\d+)$", ref).group(1))


def _colnum(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _esc(text):
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


_ROW = re.compile(r"<row\b[^>]*?/>|<row\b[^>]*?>.*?</row>", re.S)
_CELL = re.compile(r"<c\b[^>]*?\br=\"(?P<ref>[A-Z]+\d+)\"[^>]*?(?:/>|>.*?</c>)", re.S)


def _cell_style(cell_xml):
    m = re.search(r'\bs="(\d+)"', cell_xml or "")
    return m.group(1) if m else None


def _make_cell(ref, style, kind, value, formula=None):
    """새 <c> 하나. kind: 'n' 숫자 · 'str' 글자 · 'blank' 빈 칸 · 'f' 수식(계산값만 새로)."""
    s = ' s="%s"' % style if style else ""
    if kind == "f":
        body = "<f>%s</f>" % formula if formula else ""      # 수식은 XML 에서 꺼낸 그대로 (이미 escape 되어 있다)
        if value is None:
            return '<c r="%s"%s>%s</c>' % (ref, s, body)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return '<c r="%s"%s>%s<v>%.10g</v></c>' % (ref, s, body, value)
        return '<c r="%s"%s t="str">%s<v>%s</v></c>' % (ref, s, body, _esc(value))
    if kind == "blank" or value is None:
        return '<c r="%s"%s/>' % (ref, s)
    if kind == "n":
        return '<c r="%s"%s><v>%.10g</v></c>' % (ref, s, float(value))
    return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, _esc(value))


def cell_number(xml, ref):
    """워크시트에 적힌 그 칸의 숫자 (글자·빈 칸이면 None)."""
    m = re.search(r'<c\b[^>]*?\br="%s"[^>]*?(?:/>|>(?P<body>.*?)</c>)' % ref, xml, re.S)
    if m is None or not m.group("body") or 't="s"' in m.group(0) or 't="inlineStr"' in m.group(0):
        return None
    v = re.search(r"<v>([^<]*)</v>", m.group("body"))
    try:
        return float(v.group(1)) if v else None
    except (TypeError, ValueError):
        return None


def _decimals(text):
    """'6.40' → 2, '7.0' → 1, '7' → 0, 못 읽으면 None."""
    m = re.search(r"-?\d+(?:\.(\d+))?", str(text if text is not None else ""))
    if not m:
        return None
    return len(m.group(1)) if m.group(1) else 0


def style_with_decimals(blobs, base, decimals):
    """styles.xml 에 base 서식(cellXfs 번호)을 복제하고 소수 자릿수 표시 형식('0.00')만 바꾼 서식을
    더한다 → 새 서식 번호. base 가 없으면 None.

    담당자 2026-09-10: "소숫점 자리수는 시험성적서와 맞춰서 작성해줘 — 다른 시험항목도" (성적서의
    '6.2 ~ 7.0' 이 규격 칸에 7 로, 결과값 '6.40' 이 6.4 로 보였다). 값은 숫자 그대로 두고 표시 형식만
    맞춘다 — 수식·그래프가 그대로 쓴다.
    """
    if base is None or decimals is None:
        return None
    xml = blobs.get("xl/styles.xml", b"").decode("utf-8")
    if not xml:
        return None
    code = "0" if decimals == 0 else "0." + "0" * decimals
    cache = blobs.setdefault("__fmt_cache__", {}) if isinstance(blobs, dict) else {}
    key = (str(base), code)
    if key in cache:
        return cache[key]
    # numFmt — 이미 같은 형식이 있으면 그 번호
    fm = re.search(r'<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="%s"' % re.escape(code), xml)
    if fm:
        fmt_id = int(fm.group(1))
    else:
        ids = [int(x) for x in re.findall(r'<numFmt\b[^>]*numFmtId="(\d+)"', xml)]
        fmt_id = max(ids + [163]) + 1
        tag = '<numFmt numFmtId="%d" formatCode="%s"/>' % (fmt_id, code)
        nm = re.search(r'<numFmts\b[^>]*count="(\d+)"[^>]*>', xml)
        if nm:
            xml = xml.replace(nm.group(0), re.sub(r'count="\d+"', 'count="%d"' % (int(nm.group(1)) + 1), nm.group(0)) + tag, 1)
        else:
            xml = re.sub(r'(<styleSheet\b[^>]*>)', r'\1<numFmts count="1">%s</numFmts>' % tag, xml, count=1)
    xm = re.search(r'<cellXfs\b[^>]*count="(\d+)"[^>]*>(.*?)</cellXfs>', xml, re.S)
    if not xm:
        return None
    xfs = re.findall(r'<xf\b[^>]*?(?:/>|>.*?</xf>)', xm.group(2), re.S)
    try:
        src = xfs[int(base)]
    except (ValueError, IndexError):
        return None
    new_xf = re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % fmt_id, src, count=1)
    if 'numFmtId=' not in new_xf:
        new_xf = new_xf.replace("<xf ", '<xf numFmtId="%d" ' % fmt_id, 1)
    if 'applyNumberFormat=' in new_xf:
        new_xf = re.sub(r'applyNumberFormat="\d"', 'applyNumberFormat="1"', new_xf, count=1)
    else:
        new_xf = new_xf.replace("<xf ", '<xf applyNumberFormat="1" ', 1)
    count = int(xm.group(1))
    new_block = xm.group(0).replace('count="%d"' % count, 'count="%d"' % (count + 1), 1)
    new_block = new_block[:-len("</cellXfs>")] + new_xf + "</cellXfs>"
    xml = xml.replace(xm.group(0), new_block, 1)
    blobs["xl/styles.xml"] = xml.encode("utf-8")
    cache[key] = str(count)
    return str(count)


def cell_styles(xml):
    """{칸 이름: 서식 번호} — 결과값 칸의 서식(노랑·사선)을 그대로 물려받으려고 미리 읽는다."""
    return {m.group("ref"): _cell_style(m.group(0)) for m in _CELL.finditer(xml)}


def _patch_sheet(xml, wanted, styles=None):
    """워크시트 XML 의 칸 값을 바꾼다. wanted: {ref: (kind, value)}.

    있는 칸은 그 자리에서 바꾸고, 없는 칸만 그 줄에 열 차례대로 끼워 넣는다.
    """
    styles = styles or {}
    done = set()

    def one_cell(m):
        ref = m.group("ref")
        if ref not in wanted:
            return m.group(0)
        done.add(ref)
        kind, value = wanted[ref]
        old = m.group(0)
        style = styles.get(ref) or _cell_style(old)
        formula = None
        if kind == "f":
            fm = re.search(r"<f\b[^>]*>(.*?)</f>", old, re.S)
            formula = fm.group(1) if fm else None
        return _make_cell(ref, style, kind, value, formula)

    xml = _CELL.sub(one_cell, xml)
    missing = [ref for ref in wanted if ref not in done]
    if not missing:
        return xml
    by_row = {}
    for ref in missing:
        by_row.setdefault(_rownum(ref), []).append(ref)

    def one_row(m):
        whole = m.group(0)
        rm = re.match(r"<row\b[^>]*?\br=\"(\d+)\"", whole)
        refs = by_row.get(int(rm.group(1))) if rm else None
        if not refs:
            return whole
        cells = {c.group("ref"): c.group(0) for c in _CELL.finditer(whole)}
        out = whole
        for ref in sorted(refs, key=lambda r: _colnum(_letters(r))):
            kind, value = wanted[ref]
            new = _make_cell(ref, styles.get(ref), kind, value)
            after = None
            for other in sorted(cells, key=lambda r: _colnum(_letters(r))):
                if _colnum(_letters(other)) < _colnum(_letters(ref)):
                    after = other
            if after:
                out = out.replace(cells[after], cells[after] + new, 1)
                cells[ref] = new
            elif out.endswith("</row>"):
                out = out[: -len("</row>")] + new + "</row>"
                cells[ref] = new
            else:                                   # <row .../> 처럼 칸이 하나도 없는 줄
                out = re.sub(r"/>$", ">" + new + "</row>", out, count=1)
                cells[ref] = new
        return out

    return _ROW.sub(one_row, xml)


_SER = re.compile(r"<c:ser>.*?</c:ser>", re.S)
_NUMCACHE = re.compile(r"<c:numCache>.*?</c:numCache>", re.S)
_STRCACHE = re.compile(r"<c:strCache>.*?</c:strCache>", re.S)


def _patch_chart(xml, columns, count):
    """그래프가 보는 범위를 올해 줄 수에 맞추고, 저장된 값(cache)도 새 값으로 바꾼다.

    columns: {열 글자: [값 …]} — 결과값 B 와 계산 열 R·S·T·U·V.
    """
    end = FIRST_DATA_ROW + max(count, 1) - 1

    def fix_range(m):
        return "$%s$%d:$%s$%d" % (m.group(1), FIRST_DATA_ROW, m.group(2), end)

    def one_ser(m):
        block = m.group(0)
        block = re.sub(r"\$([A-Z]+)\$%d:\$([A-Z]+)\$\d+" % FIRST_DATA_ROW, fix_range, block)
        col = None
        vm = re.search(r"<c:val>.*?<c:f>[^<]*?\$([A-Z]+)\$%d" % FIRST_DATA_ROW, block, re.S)
        if vm:
            col = vm.group(1)
        vals = columns.get(col)
        if vals is None:
            return block
        pts = "".join('<c:pt idx="%d"><c:v>%.10g</c:v></c:pt>' % (i, v)
                      for i, v in enumerate(vals[:end - FIRST_DATA_ROW + 1]) if v is not None)
        cache = ('<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="%d"/>%s</c:numCache>'
                 % (end - FIRST_DATA_ROW + 1, pts))

        def swap_val(vm2):
            return _NUMCACHE.sub(cache, vm2.group(0), count=1)

        return re.sub(r"<c:val>.*?</c:val>", swap_val, block, count=1, flags=re.S)

    xml = _SER.sub(one_ser, xml)
    # 가로 축(연번) — 1..n
    def one_cat(m):
        block = m.group(0)
        pts = "".join('<c:pt idx="%d"><c:v>%d</c:v></c:pt>' % (i, i + 1) for i in range(end - FIRST_DATA_ROW + 1))
        cache = ('<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="%d"/>%s</c:numCache>'
                 % (end - FIRST_DATA_ROW + 1, pts))
        block = _NUMCACHE.sub(cache, block, count=1)
        return _STRCACHE.sub("", block, count=1) if "<c:strCache>" in block and "<c:numCache>" in block else block

    return _legend_outside(re.sub(r"<c:cat>.*?</c:cat>", one_cat, xml, flags=re.S))


_LEGEND = re.compile(r"<c:legend>.*?</c:legend>", re.S)
_LAYOUT = re.compile(r"<c:layout>.*?</c:layout>", re.S)
_PLOT_LAYOUT = re.compile(r"(<c:plotArea>\s*)<c:layout>.*?</c:layout>", re.S)


def _legend_outside(xml):
    """범례를 그래프 밖(오른쪽)으로 — 그림을 가리지 않게.

    담당자 2026-09-07: "Cpk 범례표가 그래프를 가리네." 옛 .xls 를 .xlsx 로 바꾸면서 범례에
    '손으로 정한 자리'(manualLayout)가 붙어 그래프 한가운데에 얹혔다. 그 자리 지정을 지우면
    legendPos(오른쪽)대로 엑셀이 알아서 바깥에 놓고, 그림 영역도 그만큼 줄여 준다.
    """
    def one(m):
        block = _LAYOUT.sub("", m.group(0), count=1)          # 손으로 정한 범례 자리를 지운다
        block = re.sub(r'<c:legendPos val="[^"]*"/>', '<c:legendPos val="r"/>', block, count=1)
        if "<c:legendPos" not in block:
            block = block.replace("<c:legend>", '<c:legend><c:legendPos val="r"/>', 1)
        block = re.sub(r'<c:overlay val="[^"]*"/>', '<c:overlay val="0"/>', block, count=1)
        return block

    xml = _LEGEND.sub(one, xml, count=1)
    # 그림 영역이 판을 꽉 채우게 정해져 있으면(w 0.96) 범례 자리가 없다 — 그 지정도 지운다
    return _PLOT_LAYOUT.sub(lambda m: m.group(1), xml, count=1)


_ANCHOR = re.compile(r"<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>", re.S)
_FROM = re.compile(r"(<xdr:from>.*?<xdr:row>)(\d+)(</xdr:row>\s*<xdr:rowOff>)(\d+)(</xdr:rowOff>)", re.S)
_TO = re.compile(r"(<xdr:to>.*?<xdr:row>)(\d+)(</xdr:row>\s*<xdr:rowOff>)(\d+)(</xdr:rowOff>)", re.S)
_COL = re.compile(r"<xdr:(from|to)>\s*<xdr:col>(\d+)</xdr:col>", re.S)


def _patch_drawing(xml, count):
    """빈 결과값 칸을 가로지르는 사선을 올해 줄 수에 맞춘다 — 빈칸이 남지 않게.

    담당자 2026-09-07: "사선은 빈 공간을 모두 커버해야 돼." 서식의 사선은 그 해 줄 수에 맞춰
    그어져 있어, 올해 Lot 이 더 적으면 사선 위로 빈 줄이 남는다. 사선은 결과값 열(B)에 그어진
    직선이므로 그 도형의 자리만 '첫 빈 줄 ~ 마지막 줄' 로 다시 잡는다. 빈 줄이 없으면 지운다.
    """
    first_empty = FIRST_DATA_ROW + max(count, 0)          # 1-기준: 값이 끝난 다음 줄
    last_row = FIRST_DATA_ROW + ROWS - 1

    def one(m):
        block = m.group(0)
        if "graphicFrame" in block:                        # 그래프는 건드리지 않는다
            return block
        cols = [int(c) for _, c in _COL.findall(block)]
        if len(cols) < 2 or cols[0] > 1 or cols[1] > 2:    # 결과값 열(B)에 그은 도형만
            return block
        rows = [int(r) for r in re.findall(r"<xdr:row>(\d+)</xdr:row>", block)]
        if len(rows) < 2 or rows[1] - rows[0] < 2:         # 여러 줄을 가로지르는 사선만
            return block
        if first_empty > last_row:                         # 빈 줄이 없다 — 사선을 지운다
            return ""
        block = _FROM.sub(lambda f: "%s%d%s0%s" % (f.group(1), first_empty - 1, f.group(3), f.group(5)),
                          block, count=1)
        block = _TO.sub(lambda t: "%s%d%s0%s" % (t.group(1), last_row, t.group(3), t.group(5)),
                        block, count=1)
        return block

    return _ANCHOR.sub(one, xml)


def _drawing_names(names):
    return [n for n in names if re.match(r"xl/drawings/drawing\d+\.xml$", n)]


def _sheet_name(names):
    got = [n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
    return sorted(got)[0] if got else None


def _chart_names(names):
    return [n for n in names if re.match(r"xl/charts/chart\d+\.xml$", n)]


def fill(form, dst, cells, values, today=None, decimals=None, spec_decimals=None):
    """서식 form(.xlsx) 을 dst 로 복사하며 칸 값을 갈아 끼운다. 돌려주는 값은 'Bilateral'|'Unilateral'."""
    cells = dict(cells or {})
    if today:
        cells["K4"] = today
    values = [v for v in values if v is not None]
    with zipfile.ZipFile(form) as z:
        names = z.namelist()
        blobs = {n: z.read(n) for n in names}
    sheet = _sheet_name(names)
    if sheet is None:
        raise ValueError("서식에 워크시트가 없습니다: %s" % os.path.basename(form))
    xml = blobs[sheet].decode("utf-8")
    kind = "Bilateral" if re.search(r'name="Bilateral"', blobs["xl/workbook.xml"].decode("utf-8")) else "Unilateral"

    def spec(ref):
        """규격 — 올해 값이 있으면 그것, 없으면 서식(전년도)에 적힌 값 그대로."""
        if ref in cells:
            v = cells[ref]
            return None if v is None or str(v).strip() == "" else cpk_xlsx._num(v)
        return cell_number(xml, ref)

    lo = spec("N6") if kind == "Bilateral" else spec("O6")
    hi = spec("P6")
    st = cpk_xlsx.stats(values, lo, hi)
    # 1) 머리 칸과 규격
    wanted = {}
    for ref in ("C4", "C5", "K4", "K5"):
        if ref in cells:
            wanted[ref] = ("str", cells[ref])
    seen = cell_styles(xml)
    styles = {}
    for ref in ("N6", "O6", "P6"):
        if ref not in cells:
            continue
        v = cells[ref]
        if v is None or str(v).strip() == "":
            wanted[ref] = ("blank", None)
        elif cpk_xlsx._num(v) is None:
            wanted[ref] = ("str", v)                       # 'N/A'
        else:
            wanted[ref] = ("n", cpk_xlsx._num(v))
            # 규격 칸의 소수 자릿수는 성적서 글('6.2 ~ 7.0')대로 — 지정이 없으면 글의 자릿수
            d = spec_decimals if spec_decimals is not None else (_decimals(v) if isinstance(v, str) else None)
            if d:
                got = style_with_decimals(blobs, seen.get(ref), d)
                if got is not None:
                    styles[ref] = got
    # 2) 결과값 — 값이 있는 줄은 값 칸 서식, 없는 줄은 사선 칸 서식으로
    style_filled = seen.get("%s%d" % (DATA_COL, FIRST_DATA_ROW))
    style_blank = seen.get("%s%d" % (DATA_COL, FIRST_DATA_ROW + ROWS - 1))
    if decimals:
        style_filled = style_with_decimals(blobs, style_filled, decimals) or style_filled
    for i in range(ROWS):
        ref = "%s%d" % (DATA_COL, FIRST_DATA_ROW + i)
        if i < len(values):
            wanted[ref] = ("n", values[i])
            styles[ref] = style_filled
        else:
            wanted[ref] = ("blank", None)
            styles[ref] = style_blank
    # 3) 수식 칸 — 수식은 그대로, 계산값만 새로
    top_only = cpk_xlsx._num(hi) is not None and cpk_xlsx._num(lo) is None
    cached = {"G8": st["sd"], "G9": st["mean"], "P8": st["cp"], "P9": st["cpk"]}
    if kind == "Bilateral":
        cached["G10"], cached["G11"] = st["ucl"], st["lcl"]
        line = {"R": st["mean"], "S": st["ucl"], "T": st["lcl"], "U": cpk_xlsx._num(hi), "V": cpk_xlsx._num(lo)}
    else:
        cached["G10"] = st["ucl"] if top_only else st["lcl"]
        cached["O8"] = "Short-term Capability (Cpu) :" if top_only else "Short-term Capability (Cpl) :"
        cached["F10"] = "Upper Control Limit (UCL) :" if top_only else "Lower Control Limit (LCL) :"
        cached["S9"] = "UCL" if top_only else "LCL"
        cached["T9"] = "USL" if top_only else "LSL"
        line = {"R": st["mean"], "S": cached["G10"], "T": cpk_xlsx._num(hi) if top_only else cpk_xlsx._num(lo)}
    # 값을 못 내면 지난해 값이 남지 않게 비운다 (모든 결과가 같아 σ=0 인 금속성이물 등)
    cached["I11"] = None if st["cpk"] is None else (cpk_xlsx.JUDGE_OK if st["cpk"] >= 1 else cpk_xlsx.JUDGE_NO)
    for ref, v in cached.items():
        wanted[ref] = ("f", v)
    for col, v in line.items():
        for i in range(ROWS):
            wanted["%s%d" % (col, FIRST_DATA_ROW + i)] = ("f", v if i < len(values) else None)
    xml = _patch_sheet(xml, wanted, styles)
    blobs[sheet] = xml.encode("utf-8")
    # 4) 그래프 범위·저장값
    columns = {DATA_COL: values}
    for col, v in line.items():
        columns[col] = [v] * len(values) if v is not None else []
    for chart in _chart_names(names):
        blobs[chart] = _patch_chart(blobs[chart].decode("utf-8"), columns, len(values)).encode("utf-8")
    for drawing in _drawing_names(names):                  # 빈 줄을 가로지르는 사선 자리
        blobs[drawing] = _patch_drawing(blobs[drawing].decode("utf-8"), len(values)).encode("utf-8")
    # 5) 열 때 다시 계산하게
    book = blobs["xl/workbook.xml"].decode("utf-8")
    if "<calcPr" in book:
        book = re.sub(r"<calcPr\b[^>]*/>", '<calcPr calcId="191029" fullCalcOnLoad="1"/>', book)
    else:
        book = book.replace("</workbook>", '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    blobs["xl/workbook.xml"] = book.encode("utf-8")
    blobs.pop("__fmt_cache__", None)
    tmp = dst + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name in names:
            z.writestr(name, blobs[name])
    shutil.move(tmp, dst)
    return kind


def blank(form, dst):
    """서식에서 제품별 값(제품명·공정·일자·시험항목·규격·결과값)을 지운 빈 서식을 만든다.

    칸만 비우면 쓰이지 않는 공유 글자표(sharedStrings)에 제품 이름·날짜가 그대로 남는다 —
    프로그램에 함께 넣는 빈 서식이므로 그 글자와 만든 이 이름까지 지운다.
    """
    kind = fill(form, dst, {"C4": "", "C5": "", "K4": "", "K5": "", "N6": "", "O6": "", "P6": ""}, [])
    scrub(dst)
    return kind


_SI = re.compile(r"<si>.*?</si>|<si/>", re.S)


def scrub(path):
    """아무 칸도 가리키지 않는 공유 글자와 문서 속성(만든 이)을 비운다. 지운 글자 수를 돌려준다."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        blobs = {n: z.read(n) for n in names}
    used = set()
    for name in names:
        if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
            continue
        for m in _CELL.finditer(blobs[name].decode("utf-8")):
            if 't="s"' in m.group(0):
                v = re.search(r"<v>(\d+)</v>", m.group(0))
                if v:
                    used.add(int(v.group(1)))
    gone = 0
    if "xl/sharedStrings.xml" in blobs:
        xml = blobs["xl/sharedStrings.xml"].decode("utf-8")
        count = [0]

        def one(m):
            i = count[0]
            count[0] += 1
            if i in used or m.group(0) == "<si/>":
                return m.group(0)
            return "<si><t/></si>"

        xml, _ = _SI.subn(one, xml)
        gone = count[0] - len(used)
        blobs["xl/sharedStrings.xml"] = xml.encode("utf-8")
    for name in ("docProps/core.xml", "docProps/app.xml"):
        if name not in blobs:
            continue
        text = blobs[name].decode("utf-8")
        for tag in ("dc:creator", "cp:lastModifiedBy", "Company", "Manager", "dc:title", "dc:subject"):
            text = re.sub(r"<%s>.*?</%s>" % (tag, tag), "<%s></%s>" % (tag, tag), text, flags=re.S)
        blobs[name] = text.encode("utf-8")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name in names:
            z.writestr(name, blobs[name])
    shutil.move(tmp, path)
    return gone
