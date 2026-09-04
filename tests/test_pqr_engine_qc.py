# -*- coding: utf-8 -*-
"""QC-126 Cpk 적용 조건 — 10 Lot 미만은 산출하지 않는다(본문·첨부·문안 모두)."""
from __future__ import unicode_literals

import os
import tempfile
import unittest

from pqr.engine import qc


class 적용조건(unittest.TestCase):
    def test_기준은_설정에서_읽고_기본은_10(self):
        self.assertEqual(qc.cpk_min_lots(), 10)

    def test_10_Lot_미만이면_적용_안_함(self):
        self.assertFalse(qc.cpk_applies(9))
        self.assertFalse(qc.cpk_applies(2))
        self.assertFalse(qc.cpk_applies(0))

    def test_10_Lot_이상이면_적용(self):
        self.assertTrue(qc.cpk_applies(10))
        self.assertTrue(qc.cpk_applies(17))
        self.assertTrue(qc.cpk_applies(10, lsl=None, usl=75))      # 한쪽 규격도 회사 관행대로 허용

    def test_문안(self):
        s = qc.not_applied_sentence("수출용", 2)
        self.assertIn("수출용", s)
        self.assertIn("2 Lot", s)
        self.assertIn("10 Lot 미만", s)
        self.assertIn("QC-126", s)


class 첨부파일(unittest.TestCase):
    def test_10_Lot_미만이면_Cpk_계산_파일을_만들지_않는다(self):
        from pqr.engine import excel_attach

        class Data(object):
            domestic = ["OEY101", "OEY102", "OEY103"]
            coa = {}
            issues = []
        data = Data()
        out = excel_attach.write_cpk_files(tempfile.mkdtemp(), data, None, "2026.09.04")
        self.assertEqual(out, [])
        self.assertEqual(len(data.issues), 1)
        self.assertIn("3 Lot", data.issues[0][2])
        self.assertIn("QC-126", data.issues[0][2])


if __name__ == "__main__":
    unittest.main()
