# -*- coding: utf-8 -*-
"""HLF-QC-126-06 안정성 경향 파일 생성기 — 칸 치환·계열 재구성 단위 시험."""
from __future__ import unicode_literals

import unittest

from lxml import etree

from pqr.engine import stability_xlsx as S

C = S.C

CHART = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<c:chartSpace xmlns:c="%s" xmlns:a="%s"><c:chart><c:plotArea><c:layout/>'
    '<c:lineChart><c:grouping val="standard"/>'
    '<c:ser><c:idx val="0"/><c:order val="0"/>'
    '<c:tx><c:strRef><c:f>함량!$B$72</c:f><c:strCache><c:ptCount val="1"/>'
    '<c:pt idx="0"><c:v>OLD</c:v></c:pt></c:strCache></c:strRef></c:tx>'
    '<c:val><c:numRef><c:f>함량!$C$72:$L$72</c:f></c:numRef></c:val></c:ser>'
    '<c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}"/></c:extLst>'
    '<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/>'
    '</c:lineChart></c:plotArea></c:chart></c:chartSpace>'
) % (S.NS_C, S.NS_A)

ROW = ('<row r="38"><c r="A38" s="4"><v>1</v></c><c r="B38" s="5" t="s"><v>29</v></c>'
       '<c r="C38" s="6"><v>101.7</v></c><c r="D38" s="6"/></row>')


class SetCellTest(unittest.TestCase):
    def test_문자열은_인라인으로_바뀐다(self):
        out = S.set_cell(ROW, "B38", "OEV301")
        self.assertIn('<c r="B38" s="5" t="inlineStr"><is><t xml:space="preserve">OEV301</t>', out)
        self.assertNotIn('t="s"', out)

    def test_숫자와_빈칸(self):
        out = S.set_cell(ROW, "C38", 100.5)
        self.assertIn('<c r="C38" s="6"><v>100.5</v></c>', out)
        out = S.set_cell(out, "C38", None)
        self.assertIn('<c r="C38" s="6"/>', out)

    def test_없는_칸은_그대로_둔다(self):
        self.assertEqual(S.set_cell(ROW, "Z99", 1), ROW)

    def test_스타일을_지정하면_바뀐다(self):
        self.assertIn('<c r="D38" s="9"><v>7</v></c>', S.set_cell(ROW, "D38", 7, style="9"))

    def test_특수문자는_이스케이프된다(self):
        self.assertIn("A&amp;B", S.set_cell(ROW, "B38", "A&B"))


class ChartTest(unittest.TestCase):
    def _sers(self, xml):
        root = etree.fromstring(xml)
        line = root.find(".//" + C + "plotArea/" + C + "lineChart")
        return [s.find(C + "tx//" + C + "v").text for s in line.findall(C + "ser")]

    def test_계열이_Lot수만큼_다시_만들어진다(self):
        lots = [("OEV301", {"Initial": 100.5}), ("OEW101", {"Initial": 98.7, "12M": 104.0})]
        out = S._rebuild_chart(CHART.encode("utf-8"), "함량", lots, 90, 110)
        names = self._sers(out)
        self.assertEqual(names[:2], ["OEV301", "OEW101"])
        self.assertEqual(len(names), 4)                       # Lot 2 + LCL + UCL
        self.assertTrue(names[2].startswith("하한관리"))
        self.assertTrue(names[3].startswith("상한관리"))

    def test_값이_없는_시점은_NA_로_들어간다(self):
        out = S._rebuild_chart(CHART.encode("utf-8"),
                               "함량", [("OEV301", {"Initial": 100.5})], 90, 110).decode("utf-8")
        self.assertIn("<c:v>100.5</c:v>", out)
        self.assertEqual(out.count("#N/A"), 7)                 # 8 시점 중 7 개가 빈칸

    def test_숨은_계열_목록은_지운다(self):
        out = S._rebuild_chart(CHART.encode("utf-8"),
                               "함량", [("A", {})], 90, 110).decode("utf-8")
        self.assertNotIn("02D57815-91ED-43cb-92C2-25804820EDAC", out)

    def test_꺾은선이_없으면_그대로_돌려준다(self):
        raw = b'<?xml version="1.0"?><c:chartSpace xmlns:c="%s"/>' % S.NS_C.encode()
        self.assertEqual(S._rebuild_chart(raw, "함량", [("A", {})], 90, 110), raw)


class GroupTest(unittest.TestCase):
    def test_평평한_묶음(self):
        from pqr.engine.excel_attach import _grouped
        self.assertEqual(_grouped({"OEV301": {"Initial": 100.5}}),
                         [("", [("OEV301", {"Initial": 100.5})])])

    def test_포장별_묶음(self):
        from pqr.engine.excel_attach import _grouped
        got = _grouped({"내수용": {"OEV301": {"Initial": 100.5}},
                        "수출용": {"OZW101": {"Initial": 97.3}}})
        self.assertEqual(sorted(k for k, _ in got), ["내수용", "수출용"])

    def test_빈_값(self):
        from pqr.engine.excel_attach import _grouped
        self.assertEqual(_grouped({}), [])


DOC = (
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
    '<w:tbl>'
    '<w:tr><w:tc><w:p><w:r><w:t>구 분</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:p><w:r><w:t>성 명</w:t></w:r></w:p></w:tc></w:tr>'
    '<w:tr><w:tc><w:p><w:r><w:t>작 성 (Written by)</w:t></w:r></w:p></w:tc></w:tr>'
    '<w:tr><w:tc><w:p><w:r><w:t>품질보증1팀</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:p><w:r><w:t>김 현</w:t></w:r><w:r><w:t> 수</w:t></w:r></w:p></w:tc></w:tr>'
    '</w:tbl></w:body></w:document>'
)


class WrittenByTest(unittest.TestCase):
    def _docx(self, document=DOC):
        import os
        import tempfile
        import zipfile
        path = os.path.join(tempfile.mkdtemp(prefix="pqr-doc-"), "r.docx")
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("word/document.xml", document)
        return path

    def test_작성자를_읽는다(self):
        from pqr.engine.excel_attach import written_by
        self.assertEqual(written_by(self._docx()), "김현수")

    def test_없는_파일은_빈_문자열(self):
        from pqr.engine.excel_attach import written_by
        self.assertEqual(written_by("/없는/경로.docx"), "")
        self.assertEqual(written_by(None), "")

    def test_작성_행이_없으면_빈_문자열(self):
        from pqr.engine.excel_attach import written_by
        doc = DOC.replace("작 성 (Written by)", "검 토 (Reviewed by)")
        self.assertEqual(written_by(self._docx(doc)), "")


class XlsFillTest(unittest.TestCase):
    def test_그래프_끝행은_머리행_더하기_Lot수(self):
        from pqr.engine import xls_fill
        self.assertEqual(xls_fill._last_row(17), 26)     # 9 행 머리 + 17 Lot
        self.assertEqual(xls_fill._last_row(1), 10)
        self.assertEqual(xls_fill._last_row(0), 10)      # Lot 이 없어도 한 행은 남긴다

    def test_세로축_최댓값_어림(self):
        from pqr.engine.xls_fill import _axis_max
        self.assertEqual(_axis_max(110.31), 120.0)      # 함량
        self.assertEqual(_axis_max(79.42), 80.0)        # 입자도
        self.assertEqual(_axis_max(1.0), 1.0)

    def test_결과선이_위쪽이면_가장_낮은_선_아래(self):
        from pqr.engine.xls_fill import _legend_layout as layout
        top, frac = layout([102.8, 107.4], (90, "N/A", 110, 110.31, 101.19))
        self.assertTrue(top)                            # 윗변 기준
        self.assertAlmostEqual(frac, 0.30, places=2)    # LSL 90 바로 아래
        top, frac = layout([45.0, 67.0], (None, None, 75, 79.42, 0.0))
        self.assertTrue(top)
        self.assertAlmostEqual(frac, 0.4875, places=3)  # 결과 최솟값 45 바로 아래

    def test_결과선이_바닥이면_빈_띠_가운데(self):
        from pqr.engine.xls_fill import _legend_layout as layout
        top, frac = layout([0.0, 1.0, 2.0], (None, None, 50, 2.35, 0.0))
        self.assertFalse(top)                           # 가운데 기준
        self.assertTrue(0.3 < frac < 0.7, frac)

    def test_범례_비율은_항상_0과_1_사이(self):
        from pqr.engine.xls_fill import _legend_layout as layout
        for values, levels in (([], ()), ([0.0], (0.0,)), ([5.0], (None, "N/A")),
                               ([1.0, 2.0, 3.0], (10,))):
            _, got = layout(values, levels)
            self.assertTrue(0.0 <= got <= 1.0, (values, levels, got))

    def test_계열_수식의_끝행만_바뀐다(self):
        import re
        f = "=SERIES(Bilateral!$R$9,,Bilateral!$R$10:$R$20,2)"
        got = re.sub(r"\$([A-Z]+)\$(\d+):\$([A-Z]+)\$\d+",
                     lambda m: "$%s$%s:$%s$%d" % (m.group(1), m.group(2), m.group(3), 26), f)
        self.assertEqual(got, "=SERIES(Bilateral!$R$9,,Bilateral!$R$10:$R$26,2)")

    def test_도구가_없으면_FillError(self):
        from pqr.engine import xls_fill
        real = xls_fill._with_uno
        xls_fill._with_uno = lambda *a, **k: False
        try:
            self.assertRaises(xls_fill.FillError, xls_fill.fill, "x.xls", "y.xls", {}, [])
        finally:
            xls_fill._with_uno = real


ROW72 = ('<row r="72"><c r="A72" s="4"><v>1</v></c>'
         '<c r="B72" s="5" t="str"><f>B38</f><v>OEU101</v></c>'
         '<c r="C72" s="6"><f t="shared" ref="C72:L87" si="0">IF(C38=0, NA(), C38)</f><v>101.7</v></c>'
         '<c r="D72" s="6" t="e"><f t="shared" si="0"/><v>#N/A</v></c>'
         '<c r="M72" s="6"/></row>')


class FormulaCacheTest(unittest.TestCase):
    def test_문자열_결과는_t_str(self):
        out = S.set_formula_cache(ROW72, "B72", "OEV301")
        self.assertIn('<c r="B72" s="5" t="str"><f>B38</f><v>OEV301</v></c>', out)

    def test_숫자_결과는_t_없이(self):
        out = S.set_formula_cache(ROW72, "C72", 100.5)
        self.assertIn('<f t="shared" ref="C72:L87" si="0">IF(C38=0, NA(), C38)</f><v>100.5</v>', out)
        self.assertNotIn('<c r="C72" s="6" t=', out)

    def test_값이_없으면_NA_오류로(self):
        out = S.set_formula_cache(ROW72, "C72", None)
        self.assertIn('<c r="C72" s="6" t="e">', out)
        self.assertIn("<v>#N/A</v>", out)

    def test_수식은_그대로_남는다(self):
        out = S.set_formula_cache(ROW72, "D72", 7.5)
        self.assertIn('<f t="shared" si="0"/><v>7.5</v>', out)

    def test_수식이_없는_칸은_건드리지_않는다(self):
        self.assertEqual(S.set_formula_cache(ROW72, "A72", 9), ROW72)
        self.assertEqual(S.set_formula_cache(ROW72, "M72", 9), ROW72)
        self.assertEqual(S.set_formula_cache(ROW72, "Z99", 9), ROW72)


DRAWING = (
    '<xdr:wsDr xmlns:xdr="x" xmlns:a="y">'
    '<xdr:twoCellAnchor><xdr:graphicFrame><xdr:nvGraphicFramePr>'
    '<xdr:cNvPr id="1" name="차트 1"/></xdr:nvGraphicFramePr></xdr:graphicFrame></xdr:twoCellAnchor>'
    '<xdr:twoCellAnchor><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="그림 2"/></xdr:nvPicPr>'
    '<xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic></xdr:twoCellAnchor>'
    '<xdr:twoCellAnchor><xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id="4" name="직선 연결선 4"/>'
    '</xdr:nvCxnSpPr><xdr:spPr><a:prstGeom prst="line"/></xdr:spPr></xdr:cxnSp></xdr:twoCellAnchor>'
    '</xdr:wsDr>'
)

STYLES = (
    '<styleSheet><borders count="2">'
    '<border><left/><right/><top/><bottom/><diagonal/></border>'
    '<border><left style="thin"/><right style="thin"/><top style="thin"/>'
    '<bottom style="thin"/><diagonal/></border>'
    '</borders><cellXfs count="2">'
    '<xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0"/>'
    '<xf numFmtId="176" fontId="0" fillId="2" borderId="1" xfId="0"/>'
    '</cellXfs></styleSheet>'
)


class DrawingTest(unittest.TestCase):
    def test_사선_도형만_지운다(self):
        out, removed = S.strip_floating_lines(DRAWING.encode("utf-8"))
        text = out.decode("utf-8")
        self.assertEqual(removed, 1)
        self.assertIn("차트 1", text)
        self.assertIn("그림 2", text)
        self.assertNotIn("직선 연결선", text)

    def test_지울_도형이_없으면_그대로(self):
        raw = '<xdr:wsDr xmlns:xdr="x"/>'.encode("utf-8")
        out, removed = S.strip_floating_lines(raw)
        self.assertEqual((out, removed), (raw, 0))


class DiagonalStyleTest(unittest.TestCase):
    def test_사선_테두리와_서식을_새로_만든다(self):
        out, mapping = S.diagonal_styles(STYLES.encode("utf-8"), (0, 1))
        text = out.decode("utf-8")
        self.assertIn('diagonalUp="1"', text)
        self.assertIn('<diagonal style="thin"><color indexed="64"/></diagonal>', text)
        self.assertIn('<borders count="3">', text)
        self.assertIn('<cellXfs count="4">', text)
        self.assertEqual(mapping, {0: 2, 1: 3})

    def test_이미_있는_사선_서식을_다시_쓴다(self):
        once, first = S.diagonal_styles(STYLES.encode("utf-8"), (0, 1))
        twice, second = S.diagonal_styles(once, (0, 1))
        self.assertEqual(first, second)
        self.assertEqual(once, twice)          # 두 번 돌려도 늘어나지 않는다


class PointsShownTest(unittest.TestCase):
    def test_기본은_36M_까지(self):
        self.assertEqual(S.points_shown([("A", {"Initial": 100.0})]), 8)
        self.assertEqual(S.points_shown([("A", {"Initial": 100.0, "36M": 101.0})]), 8)

    def test_36M_뒤에_결과가_있으면_늘린다(self):
        self.assertEqual(S.points_shown([("A", {"48M": 99.0})]), 9)
        self.assertEqual(S.points_shown([("A", {"Initial": 1.0}), ("B", {"60M": 2.0})]), 10)

    def test_Lot_이_없어도_36M_까지(self):
        self.assertEqual(S.points_shown([]), 8)

    def test_그래프_범위가_36M_열까지_잡힌다(self):
        lots = [("OEV301", {"Initial": 100.5, "36M": 102.8})]
        out = S._rebuild_chart(CHART.encode("utf-8"), "함량", lots, 90, 110).decode("utf-8")
        # 시트 이름은 늘 따옴표로 감싼다 — 괄호가 든 이름('함량(플루오로메톨론)')도 읽히게
        self.assertIn("'함량'!$C$72:$J$72", out)        # J = 36M
        self.assertIn('<c:ptCount val="8"/>', out)
        self.assertNotIn("$L$71", out)


class TrendSeedTest(unittest.TestCase):
    """지난 경향표 + 올해 시험일지 → 최신 경향표 (담당자 2026-09)."""

    def _seed(self):
        return [{"sheet": "함량(A성분)", "item": "함량 - A성분(%)", "product": "테스트연고",
                 "storage": "", "lcl": 90, "ucl": 110,
                 "lots": [("OGTD01", {"Initial": 100.4, "12M": 97.2}),
                          ("OGW701", {"Initial": 98.3, "12M": 98.8})]}]

    def _logs(self):
        return [{"lot": "OGW701", "points": [
                    {"period": "18M", "done": "2025.03.25", "assays": {"A성분": 97.9}},
                    {"period": "24M", "done": "2026.09.10", "assays": {"A성분": 98.5}}]},
                {"lot": "OGY301", "points": [
                    {"period": "Initial", "done": "2025.05.12", "assays": {"A성분": 97.7}}]}]

    def _data(self, seed, logs):
        class Data(object):
            pass
        data = Data()
        data.stability_trend, data.stability_logs = seed, logs
        data.coa, data.issues, data.period = {}, [], {"from": 2025, "to": 2025}
        return data

    def _sheets(self, seed, logs):
        from pqr.engine import excel_attach
        made = {}
        original = excel_attach.stability_xlsx.build_multi

        def spy(form, out, product, sheets, **kw):
            made["sheets"] = sheets
            return out
        excel_attach.stability_xlsx.build_multi = spy
        try:
            excel_attach._write_trend("form.xlsx", "/tmp", self._data(seed, logs),
                                      {"name": "테스트연고"}, "2026.09.05", None)
        finally:
            excel_attach.stability_xlsx.build_multi = original
        return made.get("sheets") or []

    def test_지난_값에_올해_시점만_덧붙인다(self):
        sheets = self._sheets(self._seed(), self._logs())
        self.assertEqual(len(sheets), 1)
        lots = dict(sheets[0]["lots"])
        self.assertEqual(lots["OGW701"], {"Initial": 98.3, "12M": 98.8, "18M": 97.9})
        self.assertEqual(lots["OGTD01"], {"Initial": 100.4, "12M": 97.2})   # 지난 값 그대로
        self.assertEqual(lots["OGY301"], {"Initial": 97.7})                 # 새 Lot 은 뒤에
        self.assertEqual([lot for lot, _ in sheets[0]["lots"]], ["OGTD01", "OGW701", "OGY301"])

    def test_시험일지가_없어도_지난_경향표로_만든다(self):
        sheets = self._sheets(self._seed(), [])
        self.assertEqual(len(dict(sheets[0]["lots"])), 2)

    def test_평가_기간을_넘어선_시점은_넣지_않는다(self):
        sheets = self._sheets(self._seed(), self._logs())
        self.assertNotIn("24M", dict(sheets[0]["lots"])["OGW701"])          # 2026 완료분
