"""제출용 제품품질평가 보고서(.docx) 생성 — 표준 라이브러리만 씁니다.

담당자 PC에는 python-docx 가 없을 수 있고 사내망에서는 설치도 어렵습니다.
.docx 는 XML 몇 개를 담은 zip 파일이므로 zipfile 로 직접 씁니다.

원칙 — 값을 지어내지 않습니다.
    표로 읽은 자료가 있으면 그 값을 그대로 옮기고, 항 번호 파일(PDF·스캔)만 있으면
    "근거 자료는 확인됐고 값은 원본에서 옮겨 적어야 한다"고 적습니다. 자료가 아예 없으면
    '자료 미확보'로 남깁니다. '자료 미확보'와 '확인 결과 이력 없음'은 다른 것이므로
    이력 없음 문구는 마감 기록(해당없음 확인)이 있는 항에만 씁니다.
"""

import datetime as _dt
import decimal
import os
import re
import zipfile

from . import build as build_module

# 회사 통일문구 — 서술을 새로 지어내지 않고 정해진 문안을 씁니다.
NO_HISTORY = {
    "11": "평가 년도 내 중요 일탈 및 기준 일탈 이력 없음.",
    "14": "평가 년도 내 반품 · 불만 · 회수 이력 없음.",
    "15": "평가 년도 내 시정 및 예방조치 사항 이력 없음.",
}
CONCLUSION = (
    "%s에 대한 제품품질평가 결과, 출발물질, 포장자재, IPC Test 그리고 제품 시험 결과 "
    "모두 정해진 규격에 만족하며, 기준에 적합한 제품이 일관되게 제조되고 있어 표준제조공정이 "
    "적절하다고 판단됨. 이에 따라, 시정 및 예방조치 또는 재밸리데이션 진행 여부 검토 시, "
    "해당되는 사항은 없음을 확인하였음."
)
PENDING_NOTE = ("자료 미확보 항목이 남아 있어 결론을 확정하지 않았습니다. "
                "아래 '확인이 필요한 항목'을 채운 뒤 결론을 확정하십시오.")
DRAFT_NOTE = ("이 문서는 제품 폴더에 올라온 자료로 프로그램이 자동 작성한 제출용 초안입니다. "
              "값이 비어 있거나 '원본 확인'으로 표시된 칸은 담당자가 원본을 보고 채운 뒤, "
              "검토 · 승인을 거쳐 발행하십시오.")

_ILLEGAL = re.compile(r'[\\/:*?"<>|]')


# ---------------- OOXML 조각 ----------------

def _esc(text):
    return (str(text if text is not None else "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _runs(text, bold=False, size=None, color=None):
    props = []
    if bold:
        props.append("<w:b/>")
    if color:
        props.append('<w:color w:val="%s"/>' % color)
    if size:
        props.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (size * 2, size * 2))
    rpr = "<w:rPr>%s</w:rPr>" % "".join(props) if props else ""
    out = []
    for index, line in enumerate(str(text if text is not None else "").split("\n")):
        if index:
            out.append("<w:r>%s<w:br/></w:r>" % rpr)
        out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, _esc(line)))
    return "".join(out)


def _para(text="", style=None, bold=False, size=None, color=None, space_before=0,
          space_after=60, align=None):
    props = []
    if style:
        props.append('<w:pStyle w:val="%s"/>' % style)
    props.append('<w:spacing w:before="%d" w:after="%d"/>' % (space_before, space_after))
    if align:
        props.append('<w:jc w:val="%s"/>' % align)
    return "<w:p><w:pPr>%s</w:pPr>%s</w:p>" % ("".join(props), _runs(text, bold, size, color))


def _cell(text, width, bold=False, shade=None, align=None, span=1):
    props = ['<w:tcW w:w="%d" w:type="dxa"/>' % width]
    if span > 1:
        props.append('<w:gridSpan w:val="%d"/>' % span)
    if shade:
        props.append('<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % shade)
    props.append('<w:vAlign w:val="center"/>')
    body = _para(text, bold=bold, size=9, space_after=0,
                 align=align or ("center" if bold else None))
    return "<w:tc><w:tcPr>%s</w:tcPr>%s</w:tc>" % ("".join(props), body)


def _grid_widths(weights, total=9638):
    unit = total / float(sum(weights))
    widths = [int(round(weight * unit)) for weight in weights]
    widths[-1] += total - sum(widths)          # 반올림 오차를 마지막 열에 몰아 줍니다
    return widths


def _table(header, rows, weights=None, note=None):
    """머리글 한 줄 + 자료 행. rows 가 비면 '해당 자료 없음' 한 줄을 넣습니다."""
    columns = len(header)
    weights = weights or [1] * columns
    widths = _grid_widths(weights)
    borders = ("<w:tblBorders>"
               + "".join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="9AA0A6"/>' % side
                         for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
               + "</w:tblBorders>")
    xml = ['<w:tbl><w:tblPr><w:tblW w:w="9638" w:type="dxa"/>' + borders
           + '<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></w:tblPr>'
           + "<w:tblGrid>" + "".join('<w:gridCol w:w="%d"/>' % w for w in widths)
           + "</w:tblGrid>"]
    xml.append("<w:tr><w:trPr><w:tblHeader/></w:trPr>"
               + "".join(_cell(text, widths[i], bold=True, shade="EEF1F5")
                         for i, text in enumerate(header)) + "</w:tr>")
    if rows:
        for row in rows:
            cells = list(row) + [""] * (columns - len(row))
            xml.append("<w:tr>" + "".join(_cell(value, widths[i])
                                          for i, value in enumerate(cells[:columns])) + "</w:tr>")
    else:
        xml.append("<w:tr>" + _cell("해당 자료 없음", sum(widths), span=columns,
                                    align="center") + "</w:tr>")
    if note:
        xml.append("<w:tr>" + _cell(note, sum(widths), span=columns, align="left") + "</w:tr>")
    xml.append("</w:tbl>")
    return "".join(xml) + _para("", space_after=80)


# ---------------- 값 다듬기 ----------------

def _text(value, dash="—"):
    value = "" if value is None else str(value).strip()
    return value or dash


def _num(value, digits=3):
    """원자료 자릿수를 넘겨 짚지 않도록 뒤의 0 은 떼고 적습니다."""
    if value is None or value == "":
        return "—"
    if isinstance(value, float):
        text = ("%%.%df" % digits) % value
        return text.rstrip("0").rstrip(".") if "." in text else text
    return str(value)


def _decimals(*values):
    """원자료의 소수 자릿수를 되짚습니다 — 평균을 원자료보다 잘게 쓰지 않기 위해서입니다."""
    found = 0
    for value in values:
        if not isinstance(value, float):
            continue
        text = repr(value)
        if "e" in text or "E" in text:
            continue
        if "." in text:
            found = max(found, len(text.split(".")[1]))
    return min(found, 3)


def _round(value, digits):
    """사사오입(ROUND_HALF_UP). 파이썬 기본 반올림은 은행식이라 1.0165 를 1.016 으로 적습니다."""
    if value is None:
        return None
    quantum = decimal.Decimal(1).scaleb(-digits)
    return decimal.Decimal(repr(float(value))).quantize(quantum, rounding=decimal.ROUND_HALF_UP)


def _fixed(value, digits):
    rounded = _round(value, digits)
    return "—" if rounded is None else ("%%.%df" % digits) % rounded


def _spec(test):
    lsl, usl = test.get("lsl"), test.get("usl")
    unit = test.get("unit") or ""
    if lsl is None and usl is None:
        return "원본 확인"
    if lsl is None:
        return "%s 이하 %s" % (_num(usl), unit)
    if usl is None:
        return "%s 이상 %s" % (_num(lsl), unit)
    return "%s ~ %s %s" % (_num(lsl), _num(usl), unit)


def _result_range(test):
    """회사 서식대로 결과는 평균과 범위를 함께 적습니다: Av. 7.03(7.00 ~ 7.09).

    자릿수는 원자료와 같게 맞춥니다 — 평균만 소수 셋째 자리로 적으면 측정하지 않은
    정밀도를 적어 넣는 셈이 됩니다.
    """
    if test.get("mean") is None:
        return "원본 확인"
    digits = _decimals(test.get("min"), test.get("max"), test.get("lsl"), test.get("usl"))
    if test.get("min") is None or test.get("max") is None:
        return "Av. %s" % _fixed(test["mean"], digits)
    return "Av. %s(%s ~ %s)" % (_fixed(test["mean"], digits),
                                _fixed(test["min"], digits), _fixed(test["max"], digits))


# ---------------- 항목별 본문 ----------------

def _evidence_note(files):
    if not files:
        return None
    return "근거 자료: " + ", ".join(files)


def _section_heading(number, label):
    return _para("%s. %s" % (number, label), bold=True, size=12,
                 space_before=200, space_after=80)


def _status_line(state, files, closed):
    if closed:
        return _para("자료 상태: 대장 확인 결과 평가 년도 내 해당 이력 없음 (마감 기록 보유)",
                     size=9, color="1A7F37")
    if state == "y":
        note = _evidence_note(files) or "제출된 대장에서 집계"
        return _para("자료 상태: 수집 완료 — %s" % note, size=9, color="1A7F37")
    if state == "p":
        return _para("자료 상태: 진행 중 — 미결 건이 남아 있습니다", size=9, color="B26A00")
    return _para("자료 상태: 자료 미확보 — 원본을 받아 채워야 합니다", size=9, color="C0392B")


def _batches_table(records):
    rows = [[_text(row.get("batch_no")), _text(row.get("mfg_date")),
             _text(row.get("batch_size") or row.get("quantity")),
             _text(row.get("pack_size")), _text(row.get("result") or "적합"),
             _text(row.get("note"), "")]
            for row in records]
    return _table(["Lot No.", "제조일자", "제조단위", "포장단위", "결과", "비고"],
                  rows, [16, 16, 18, 18, 12, 20])


def _tests_table(tests):
    rows = [[_text(test["test_name"]), _spec(test), _result_range(test),
             str(test.get("n") or 0),
             "적합" if not test.get("oos_count") else "부적합 %d건" % test["oos_count"]]
            for test in tests]
    return _table(["시험항목", "허용기준", "결과", "n", "판정"], rows, [26, 24, 30, 8, 12])


def _cpk_table(tests, min_lots):
    """Cpk 는 양측 규격이면서 로트 수가 기준 이상일 때만 적습니다."""
    rows = []
    for test in tests:
        if test.get("lsl") is None or test.get("usl") is None:
            continue
        if (test.get("n") or 0) < min_lots:
            continue
        digits = _decimals(test.get("min"), test.get("max"), test.get("lsl"), test.get("usl"))
        rows.append([_text(test["test_name"]), str(test["n"]),
                     _fixed(test.get("mean"), digits), _fixed(test.get("sd"), digits),
                     _fixed(test.get("cp"), 2), _fixed(test.get("cpk"), 2),
                     _text(test.get("verdict"))])
    if not rows:
        return _para("양측 규격이면서 %d Lot 이상인 시험항목이 없어 Cpk 를 산출하지 않았습니다."
                     % min_lots, size=9)
    return _table(["시험항목", "n", "평균", "표준편차", "Cp", "Cpk", "판정"],
                  rows, [26, 8, 14, 14, 12, 12, 14])


def _records_table(records, columns, weights):
    rows = [[_text(record.get(field), "") for _, field in columns] for record in records]
    return _table([label for label, _ in columns], rows, weights)


def _stability_table(entries):
    rows = [[_text(entry.get("study") or entry.get("condition")),
             _text(entry.get("test_name")), _text(entry.get("lot") or entry.get("batch_no")),
             _num(entry.get("min")), _num(entry.get("max")),
             "부정적 경향" if entry.get("adverse") else "적합"]
            for entry in entries]
    return _table(["안정성 종류 · 보관조건", "시험항목", "대상 Lot", "최소", "최대", "경향분석"],
                  rows, [24, 20, 18, 12, 12, 14])


def _qualification_table(entries):
    rows = [[_text(entry.get("equipment") or entry.get("name")), _text(entry.get("type")),
             _text(entry.get("doc_no")), _text(entry.get("done_date")),
             _text(entry.get("next_date")), _text(entry.get("state"))]
            for entry in entries]
    return _table(["설비 · 대상", "구분", "문서번호", "완료일", "차기 예정일", "상태"],
                  rows, [24, 14, 20, 14, 14, 14])


def _body_for(number, product, quality, records, config):
    """항 번호별 본문. 표로 읽은 자료가 있을 때만 표를 넣습니다."""
    tests = quality.get("tests", [])
    if number == "6":
        return _batches_table(records.get("batches", []))
    # 같은 집계표를 여러 항에 반복해 넣지 않습니다 — 대표 항에만 싣고 나머지 항은
    # 어떤 근거 파일로 확인했는지만 적습니다. 원본은 항마다 서식이 다릅니다.
    if number == "8.2.1.1":
        return _records_table(records.get("materials", []),
                              [("자재명", "material_name"), ("Lot No.", "batch_no"),
                               ("공급처 · 제조원", "supplier"), ("시험항목", "test_name"),
                               ("결과", "value"), ("판정", "result")],
                              [24, 16, 20, 16, 12, 12])
    if number == "10.2":
        return _qualification_table(quality.get("qualification", []))
    if number == "11":
        deviations = records.get("deviations", []) + records.get("oos", [])
        if not deviations:
            return None
        return _records_table(deviations,
                              [("번호", "record_no"), ("발생일", "opened_date"),
                               ("제목", "title"), ("등급", "severity"),
                               ("상태", "status"), ("CAPA", "capa_no")],
                              [16, 14, 30, 12, 14, 14])
    if number == "12":
        changes = records.get("changes", [])
        if not changes:
            return None
        return _records_table(changes,
                              [("변경번호", "record_no"), ("발생일", "opened_date"),
                               ("변경 내용", "title"), ("상태", "status"),
                               ("종결일", "closed_date")],
                              [18, 14, 38, 14, 16])
    if number == "13":
        return _stability_table(quality.get("stability", []))
    if number == "14":
        complaints = records.get("complaints", [])
        if not complaints:
            return None
        return _records_table(complaints,
                              [("번호", "record_no"), ("접수일", "opened_date"),
                               ("구분 · 내용", "title"), ("상태", "status"),
                               ("종결일", "closed_date")],
                              [16, 14, 40, 14, 16])
    if number == "15":
        commitments = records.get("commitments", [])
        if not commitments:
            return None
        return _records_table(commitments,
                              [("번호", "record_no"), ("조치 내용", "title"),
                               ("상태", "status"), ("종결일", "closed_date")],
                              [16, 46, 18, 20])
    return None


# ---------------- 문서 조립 ----------------

def _cover(product, data, period_text):
    meta = product.get("license") or {}
    rows = [
        ["제품코드", _text(product.get("code")), "제형", _text(product.get("form"))],
        ["제품명", _text(product.get("name")), "제품 분류", _text(product.get("form_group"))],
        ["허가번호", _text(meta.get("license_no")), "허가일자", _text(meta.get("license_date"))],
        ["사용기한", _text(meta.get("shelf_life")), "보관조건", _text(meta.get("storage"))],
        ["평가 기간", period_text, "평가 그룹", _text(product.get("group"))],
        ["제조 Lot 수", _text(product.get("lots") or product.get("batches")),
         "작성 담당자", _text(product.get("owner"))],
        ["작성일", _text(data.get("today")), "자료 수집률",
         "%d / %d 항목 (%d%%)" % (product.get("collected", 0), len(product.get("checks") or []),
                                  product.get("pct", 0))],
    ]
    widths = _grid_widths([16, 34, 16, 34])
    borders = ("<w:tblBorders>"
               + "".join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="9AA0A6"/>' % side
                         for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
               + "</w:tblBorders>")
    xml = ['<w:tbl><w:tblPr><w:tblW w:w="9638" w:type="dxa"/>' + borders
           + '<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></w:tblPr>'
           + "<w:tblGrid>" + "".join('<w:gridCol w:w="%d"/>' % w for w in widths)
           + "</w:tblGrid>"]
    for row in rows:
        xml.append("<w:tr>"
                   + _cell(row[0], widths[0], bold=True, shade="EEF1F5")
                   + _cell(row[1], widths[1])
                   + _cell(row[2], widths[2], bold=True, shade="EEF1F5")
                   + _cell(row[3], widths[3]) + "</w:tr>")
    xml.append("</w:tbl>")
    return "".join(xml) + _para("", space_after=80)


def _schedule_sentence(period_text, product, today):
    """4항 일정 통일문구. 평가 년도 · 작성 년도 · 그룹은 실제 값으로 채웁니다."""
    year = (period_text.split("~")[0].strip()[:4] or "").strip()
    written = str(today or "")[:4]
    group = _text(product.get("group"), "해당")
    return ("제품품질평가는 %s 년 1월 ~ 12월까지 생산된 해당제품에 대하여 평가를 실시하며, "
            "'QC-126 제품품질평가 규정'에 따라 %s 으로 선정되어 %s 년도 내에 완료한다."
            % (year or "____", group, written or "____"))


def build_document_xml(data, code, config=None):
    """제품 하나의 제출용 보고서 본문 XML 을 만듭니다."""
    product = next((item for item in data["products"] if item["code"] == code), None)
    if product is None:
        raise KeyError("제품 코드를 찾을 수 없습니다: %s" % code)
    quality = data.get("quality", {}).get(code, {})
    records = quality.get("records", {})
    period = data.get("period", {}) or {}
    period_text = "%s ~ %s" % (_text(period.get("from")), _text(period.get("to")))
    items = data["items"]
    states = dict(zip([row[0] for row in items], product["checks"]))
    item_files = product.get("item_files") or {}
    thresholds = (config or {}).get("thresholds", {})

    body = [_para("제품품질평가 보고서", bold=True, size=18, align="center", space_after=40),
            _para(_text(product.get("name")), bold=True, size=13, align="center",
                  space_after=200),
            _cover(product, data, period_text),
            _para(DRAFT_NOTE, size=9, color="7A5900", space_after=160)]

    body.append(_section_heading("4", "평가 일정"))
    body.append(_para(_schedule_sentence(period_text, product, data.get("today"))))

    tests = quality.get("tests", [])
    pending = []
    for number, label, _hint in items:
        files = item_files.get(number) or []
        closed = any(build_module.CLOSE_MARKER in name for name in files)
        state = states.get(number, "n")

        if number == "9.2.1":
            # 완제품 시험결과와 공정능력은 9.2 세부 항 앞에 한 번만 싣습니다.
            body.append(_section_heading("9.1", "완제품 시험결과"))
            body.append(_tests_table(tests))
            body.append(_para("공정능력 (Cpk)", bold=True, size=10, space_before=60))
            body.append(_cpk_table(tests, thresholds.get("cpk_min_lots", 10)))

        body.append(_section_heading(number, label))
        body.append(_status_line(state, files, closed))
        table = _body_for(number, product, quality, records, config)
        if table:
            body.append(table)

        if closed and number in NO_HISTORY:
            body.append(_para(NO_HISTORY[number]))
        elif number in NO_HISTORY and any("해당없음" in name or "해당 없음" in name
                                          for name in files):
            # 담당자가 '해당없음' 이라 이름 붙인 대장이 근거로 올라온 경우입니다.
            # 파일 이름만으로 확정하지 않고, 원본을 확인한 뒤 쓸 문구를 함께 제시합니다.
            body.append(_para("근거 파일 이름이 '해당없음' 입니다. 대장 원본을 확인한 뒤 "
                              "아래 문구로 확정하십시오.", size=9, color="B26A00"))
            body.append(_para(NO_HISTORY[number], size=9, color="5F6368"))
        elif state == "n":
            pending.append("%s %s" % (number, label))
            body.append(_para("자료 미확보 — 원본을 받아 이 항을 채우십시오. "
                              "확인 결과 이력이 없다면 대시보드의 '이력 없음으로 마감'을 "
                              "눌러 확인 기록을 남기십시오.", size=9, color="C0392B"))
        elif table is None:
            note = _evidence_note(files)
            body.append(_para(("%s — 표 값은 원본을 보고 옮겨 적으십시오." % note) if note
                              else "제출된 대장에 해당 건이 없습니다. 원본을 확인해 주십시오.",
                              size=9))

    # 결론은 시험 결과 수치를 실제로 읽었을 때만 확정합니다. 근거 PDF 가 폴더에 있다는 것과
    # 그 안의 값이 규격에 맞는다는 것은 다른 이야기입니다 — 스캔 파일만 보고
    # "모두 규격에 만족" 이라고 적어 두면 확인하지 않은 판정이 보고서에 남습니다.
    if not tests:
        pending.append("완제품 · 공정 시험결과 원자료 (수치를 표로 읽지 못했습니다)")
    body.append(_section_heading("16", "결론"))
    if pending:
        body.append(_para(PENDING_NOTE, color="C0392B"))
        body.append(_para("확인이 필요한 항목", bold=True, size=10, space_before=120))
        for line in pending:
            body.append(_para("· " + line, size=9, space_after=20))
        body.append(_para("확인이 끝나면 아래 문안으로 결론을 확정하십시오.", size=9,
                          space_before=120))
        body.append(_para(CONCLUSION % _text(product.get("name")), size=9, color="5F6368"))
    else:
        body.append(_para(CONCLUSION % _text(product.get("name"))))

    body.append(_section_heading("별첨", "제품 폴더에서 확인한 근거 자료"))
    rows = [[number, label, ", ".join(item_files.get(number) or []) or "—"]
            for number, label, _hint in items]
    body.append(_table(["항", "평가항목", "파일"], rows, [8, 30, 62]))

    body.append(_para("검토 · 승인", bold=True, size=11, space_before=200))
    body.append(_table(["작성", "검토", "승인"], [["", "", ""]], [1, 1, 1]))

    section = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
               '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
               'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:body>" + "".join(body) + section + "</w:body></w:document>")


_STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:docDefaults><w:rPrDefault><w:rPr>"
    '<w:rFonts w:ascii="맑은 고딕" w:hAnsi="맑은 고딕" w:eastAsia="맑은 고딕" w:cs="맑은 고딕"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/>'
    "</w:rPr></w:rPrDefault>"
    "<w:pPrDefault><w:pPr><w:spacing w:after=\"60\" w:line=\"264\" w:lineRule=\"auto\"/>"
    "</w:pPr></w:pPrDefault></w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/><w:qFormat/></w:style>'
    "</w:styles>"
)

_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.wordprocessingml.styles+xml"/>'
    "</Types>"
)

_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/officeDocument" Target="word/document.xml"/></Relationships>'
)

_DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/styles" Target="styles.xml"/></Relationships>'
)


def report_filename(product, period=None):
    """완성본 인식 규칙에 맞게 '제출용' 을 이름에 넣습니다."""
    year = str((period or {}).get("from") or "")[:4]
    label = "%s년 " % year if year else ""
    name = "[%s] %s %s제품품질평가 (제출용).docx" % (
        product.get("code"), product.get("name"), label)
    return _ILLEGAL.sub("_", name)


def write_docx(data, code, out_dir, config=None, filename=None):
    """제출용 보고서를 out_dir 에 저장하고 경로를 돌려줍니다."""
    product = next((item for item in data["products"] if item["code"] == code), None)
    if product is None:
        raise KeyError("제품 코드를 찾을 수 없습니다: %s" % code)
    document = build_document_xml(data, code, config)
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    target = os.path.join(out_dir, filename or report_filename(product, data.get("period")))
    stamp = _dt.datetime.now().timetuple()[:6]
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in (("[Content_Types].xml", _CONTENT_TYPES),
                              ("_rels/.rels", _ROOT_RELS),
                              ("word/_rels/document.xml.rels", _DOC_RELS),
                              ("word/styles.xml", _STYLES),
                              ("word/document.xml", document)):
            info = zipfile.ZipInfo(name, date_time=stamp)
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, payload.encode("utf-8"))
    return target
