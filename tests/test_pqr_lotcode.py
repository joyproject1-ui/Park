# -*- coding: utf-8 -*-
"""제조번호에서 연도를 읽는다 — PQR 연도는 제조 연도의 다음 해다.

담당자 설명(2026-09): 세 번째 글자가 제조 연도다(X = 2024, Y = 2025, Z = 2026). 2025년에 만든
OGY301 은 **2026년 PQR** 이다. 예전에는 평가 기간의 연도를 그대로 써서 '2025년 제품품질평가'
로 나왔다. 전년도 결재본 이름이 'PQR25'(2024년 제조분)인 것과도 맞는다.
"""
import datetime
import unittest

from pqr import docx_report
from pqr.engine import lotcode


TODAY = datetime.date(2026, 9, 4)


class 제조_연도(unittest.TestCase):
    def test_세_번째_글자가_제조_연도다(self):
        self.assertEqual(lotcode.made_year("OGW701", TODAY), 2023)
        self.assertEqual(lotcode.made_year("OGX901", TODAY), 2024)
        self.assertEqual(lotcode.made_year("OGY301", TODAY), 2025)
        self.assertEqual(lotcode.made_year("OGZ101", TODAY), 2026)

    def test_PQR_은_다음_해_것이다(self):
        self.assertEqual(lotcode.pqr_year("OGY301", TODAY), 2026)      # 2025년 제조 → 2026 PQR
        self.assertEqual(lotcode.pqr_year("OGX901", TODAY), 2025)      # 2024년 제조 → 2025 PQR

    def test_네_번째_글자가_제조_월이다(self):
        # 1~9 는 그대로, O 는 10월, N 은 11월, D 는 12월
        self.assertEqual(lotcode.made_month("OGY301"), 3)
        self.assertEqual(lotcode.made_month("OEYO01"), 10)
        self.assertEqual(lotcode.made_month("OEYN01"), 11)
        self.assertEqual(lotcode.made_month("OEYD01"), 12)

    def test_읽을_수_없으면_없다고_한다(self):
        for bad in ("", None, "ABC", "1234567", "OG1301"):
            self.assertIsNone(lotcode.made_year(bad, TODAY), bad)

    def test_여러_Lot_에서_가장_많은_해를_고른다(self):
        self.assertEqual(lotcode.years(["OGY301", "OGY901", "OGY902"], TODAY), (2025, 2026))
        self.assertEqual(lotcode.years(["OGY301", "OGY901", "OGX901"], TODAY), (2025, 2026))
        self.assertEqual(lotcode.years([], TODAY), (None, None))

    def test_해가_다른_Lot_을_짚어_준다(self):
        self.assertEqual(lotcode.odd_lots(["OGY301", "OGX901"], 2025, TODAY), ["OGX901"])
        self.assertEqual(lotcode.odd_lots(["OGY301", "읽을수없음"], 2025, TODAY), [])


class 보고서_이름(unittest.TestCase):
    제품 = {"code": "QC1-7014", "name": "디겐타안연고"}

    def test_제조번호로_PQR_연도를_정한다(self):
        got = docx_report.report_filename(self.제품, {"from": "2025-01-01"},
                                          lots=["OGY301", "OGY901"], today=TODAY)
        self.assertIn("2026년 제품품질평가", got)

    def test_제조번호가_없으면_평가_기간의_다음_해(self):
        got = docx_report.report_filename(self.제품, {"from": "2025-01-01"})
        self.assertIn("2026년 제품품질평가", got)

    def test_기간도_없으면_연도를_적지_않는다(self):
        self.assertIn("디겐타안연고 제품품질평가", docx_report.report_filename(self.제품, {}))


if __name__ == "__main__":
    unittest.main()
