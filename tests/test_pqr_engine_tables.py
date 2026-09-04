# -*- coding: utf-8 -*-
"""표 차림새 — 2026 디겐타 결재본을 보고 담당자가 짚어 준 것들.

  · 폴더와 압축에 같은 파일이 있으면 한 번만 읽는다 (8.2 표에 같은 줄이 두 벌 실렸다)
  · 같은 원료·코드는 시험번호가 달라도 세로로 합친다
  · 표는 모두 '창에 자동으로 맞춤' — 본문 폭에 딱 맞춘다
  · 내용이 없으면 빈 줄을 하나만 남긴다
"""
import os
import shutil
import tempfile
import unittest
import zipfile

import docx
from docx.oxml.ns import qn

from pqr.engine import collect, docedit as E


class 같은_파일은_한_번만(unittest.TestCase):
    """담당자는 자료를 압축으로 올린 뒤 제품 폴더에서 풀어 두는 일이 흔하다."""

    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="pqr-dup-")
        inner = os.path.join(self.root, "2026 필요 자료")
        os.makedirs(inner)
        for name in ("8.2.1 RBG201 - 주원료 ERP.xls", "7. 수율현황표.xlsx"):
            with open(os.path.join(inner, name), "wb") as handle:
                handle.write(b"same-bytes")
        with zipfile.ZipFile(os.path.join(self.root, "2026 필요 자료.zip"), "w") as archive:
            for name in ("8.2.1 RBG201 - 주원료 ERP.xls", "7. 수율현황표.xlsx"):
                archive.writestr("2026 필요 자료/" + name, b"same-bytes")

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    def test_폴더와_압축에_같은_파일이_있으면_한_번만(self):
        got = collect.discover(self.root)
        self.assertEqual(len(got.get("8.2.1", [])), 1)
        self.assertEqual(len(got.get("7", [])), 1)

    def test_남기는_것은_제품_폴더_안의_파일이다(self):
        # 담당자가 열어 고치는 것은 폴더 쪽이고, 압축 안의 것은 올릴 때 굳은 사본이다.
        got = collect.discover(self.root)
        self.assertTrue(os.path.abspath(got["8.2.1"][0]).startswith(os.path.abspath(self.root)))

    def test_크기가_다르면_다른_파일이다(self):
        with open(os.path.join(self.root, "2026 필요 자료", "8.2.1 다른 원료 ERP.xls"), "wb") as handle:
            handle.write(b"different-length-bytes")
        self.assertEqual(len(collect.discover(self.root).get("8.2.1", [])), 2)


def _table(rows, cols):
    d = docx.Document()
    t = d.add_table(rows=rows, cols=cols)
    return d, t


class 창에_자동으로_맞춤(unittest.TestCase):
    """전년도 결재본에서 물려받은 표는 본문 폭보다 넓어 오른쪽 여백을 넘는 것이 있었다."""

    def test_본문_폭에_딱_맞춘다(self):
        d, t = _table(2, 3)
        grid = t._tbl.find(qn("w:tblGrid"))
        for g, w in zip(grid.findall(qn("w:gridCol")), (2000, 4000, 4105)):
            g.set(qn("w:w"), str(w))
        self.assertTrue(E.fit_to_window(t, 9978))
        got = [int(g.get(qn("w:w"))) for g in grid.findall(qn("w:gridCol"))]
        self.assertEqual(sum(got), 9978)                      # 넘치지도 모자라지도 않는다
        self.assertLess(abs(got[0] / 9978.0 - 2000 / 10105.0), 0.01)   # 비율은 그대로

    def test_표_너비를_백_퍼센트로_둔다(self):
        d, t = _table(2, 2)
        E.fit_to_window(t, 9978)
        pr = t._tbl.find(qn("w:tblPr"))
        self.assertEqual(pr.find(qn("w:tblW")).get(qn("w:type")), "pct")
        self.assertEqual(pr.find(qn("w:tblW")).get(qn("w:w")), "5000")
        self.assertEqual(pr.find(qn("w:tblLayout")).get(qn("w:type")), "autofit")
        self.assertEqual(pr.find(qn("w:tblInd")).get(qn("w:w")), "0")

    def test_본문_폭을_모르면_손대지_않는다(self):
        d, t = _table(2, 2)
        self.assertFalse(E.fit_to_window(t, 0))

    def test_쪽_설정에서_본문_폭을_읽는다(self):
        d = docx.Document()
        sect = d.sections[0]._sectPr
        pg = sect.find(qn("w:pgSz")); pg.set(qn("w:w"), "11906")
        mar = sect.find(qn("w:pgMar")); mar.set(qn("w:left"), "1077"); mar.set(qn("w:right"), "851")
        self.assertEqual(E.text_width(d), 9978)


class 내용이_없으면(unittest.TestCase):
    """담당자 지시: 내용이 없으면 1행만 남겨두고 사선처리."""

    def test_빈_줄을_하나만_남긴다(self):
        d, t = _table(4, 3)
        t.rows[0].cells[0].text = "연번"
        self.assertEqual(E.collapse_empty_block(t, 1, 3), 2)
        self.assertEqual(len(t.rows), 2)

    def test_한_줄뿐이면_지우지_않는다(self):
        d, t = _table(2, 3)
        self.assertEqual(E.collapse_empty_block(t, 1, 1), 0)
        self.assertEqual(len(t.rows), 2)


if __name__ == "__main__":
    unittest.main()
