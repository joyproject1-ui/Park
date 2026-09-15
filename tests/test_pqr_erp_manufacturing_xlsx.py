# -*- coding: utf-8 -*-
"""6항 제조내역을 ERP 엑셀로 올려도 읽는다 (올로원스 2026-09-16: '6. 제조내역- ERP.xlsx' 가 "못 읽음").

PDF 와 같은 네 칸(Lot · 품명 · 제조일자 · 사용기한)을 줄로 풀어 같은 규칙으로 읽는다.
"""
from __future__ import unicode_literals

import datetime
import os
import tempfile
import unittest

from pqr.engine.readers import erp


class 엑셀_제조내역(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()

    def make(self, rows, name="6. 제조내역- ERP.xlsx"):
        from openpyxl import Workbook
        wb = Workbook(); ws = wb.active
        for row in rows:
            ws.append(row)
        path = os.path.join(self.dir, name); wb.save(path)
        return path

    def test_네_칸을_읽는다(self):
        path = self.make([
            ["Lot No.", "품명", "제조일자", "사용기한"],
            ["GVY601", "올로원스점안액", "2025.06.10", "2027.06.09"],
            ["GVYD01", "올로원스점안액", "2025.12.11", "2027.12.10"],
        ])
        got = erp.read_manufacturing(path)
        self.assertEqual([g[0] for g in got], ["GVY601", "GVYD01"])
        self.assertEqual(got[0][2], "2025.06.10")
        self.assertEqual(got[1][3], "2027.12.10")

    def test_날짜_칸이_엑셀_날짜여도_읽는다(self):
        path = self.make([
            ["Lot No.", "품명", "제조일자", "사용기한"],
            ["GVY601", "올로원스점안액", datetime.datetime(2025, 6, 10), datetime.date(2027, 6, 9)],
        ])
        got = erp.read_manufacturing(path)
        self.assertEqual(got, [("GVY601", "올로원스점안액", "2025.06.10", "2027.06.09")])

    def test_표가_아니면_빈_목록(self):
        path = self.make([["아무 글", "없음"], ["x", "y"]])
        self.assertEqual(erp.read_manufacturing(path), [])

    def test_6항이_엑셀을_받아들인다(self):
        from pqr.engine import collect
        self.assertIn(".xlsx", collect.READABLE["6"][0])
        self.assertIn(".xls", collect.READABLE["6"][0])


if __name__ == "__main__":
    unittest.main()
