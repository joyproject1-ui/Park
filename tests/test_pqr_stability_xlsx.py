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
        self.assertEqual(out.count("#N/A"), 9)                 # 10 시점 중 9 개가 빈칸

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
