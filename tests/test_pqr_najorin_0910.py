# -*- coding: utf-8 -*-
"""2026-09-10 나조린점안액 문의 목록 답변으로 굳힌 규칙들."""
from __future__ import unicode_literals

import re
import unittest

from pqr.engine.readers import erp
from pqr.engine import recipe_ointment as R
from pqr.engine import excel_attach as XA


class 제조내역ERP(unittest.TestCase):
    def test_ERP_제조번호를_Lot_으로(self):
        text = ("제조번호 (제조) 제조일자 사용기한\n"
                "NJS2-2025-L0K2Y-2000101 2025.02.13 2027.02.12\n"
                "NJS2-2025-L1K1Y-N000011 2025.11.13 2027.11.12\n"
                "NJS2-2025-L1K2Y-D000033 2025.12.08 2027.12.07\n")
        got = [erp.lot_from_erp(m) for m in erp.ERP_LINE.finditer(text)]
        self.assertEqual(got, ["LKY201", "LKYN01", "LKYD03"])

    def test_달이_맞지_않으면_버린다(self):
        m = erp.ERP_LINE.search("NJS2-2025-L0K2Y-3000101 2025.02.13 2027.02.12\n")
        self.assertIsNone(erp.lot_from_erp(m))


class 변경관리해당(unittest.TestCase):
    def test_등_N품목은_해당한다(self):
        self.assertTrue(R.change_covers({"products": "브리딘티 점안액 0.15% 등 29품목", "title": "…"}, "나조린점안액"))

    def test_이름이_없고_무리도_아니면_아니다(self):
        self.assertFalse(R.change_covers({"products": "후메론점안액, 오클점안액", "title": "…"}, "나조린점안액"))


class 무균되찾기(unittest.TestCase):
    def test_포장_판독의_bioburden_음성은_무균이다(self):
        self.assertEqual(R._sterility_from_bioburden({"bioburden": "음성"}), "음성")
        self.assertIsNone(R._sterility_from_bioburden({"bioburden": "0"}))
        self.assertIsNone(R._sterility_from_bioburden({}))


class 소수자릿수(unittest.TestCase):
    def test_성적서_글의_자릿수(self):
        self.assertEqual(XA._decimals("6.40"), 2)
        self.assertEqual(XA._decimals("7.0"), 1)
        self.assertEqual(XA._decimals("7"), 0)
        self.assertIsNone(XA._decimals("음성"))

    def test_표시_형식(self):
        self.assertEqual(XA._formats({"N6": 6.2, "P6": 7.0}, 2, 1), {"N6": "0.0", "P6": "0.0", "B": "0.00"})
        self.assertEqual(XA._formats({"P6": 15.0}, 0, 0), {})


class 기준변경각주(unittest.TestCase):
    def test_범위_꼴만_본다(self):
        self.assertTrue(R.RANGE_SPEC.match("6.2 ~ 7.0"))
        self.assertTrue(R.RANGE_SPEC.match("0.906~1.108"))
        self.assertFalse(R.RANGE_SPEC.match("10CFU/100mL 이하"))
        self.assertFalse(R.RANGE_SPEC.match("무색 투명한 액"))



class 판독이상치(unittest.TestCase):
    def test_다른_Lot_과_크게_다르면_애매함(self):
        logs = [{"lot": "A", "kind": "장기", "market": "내수", "points": [{"period": "6M", "assays": {"X": 107.0, "pH": 6.4}}]},
                {"lot": "B", "kind": "장기", "market": "내수", "points": [{"period": "6M", "assays": {"X": 103.4, "pH": 6.4}}]},
                {"lot": "C", "kind": "장기", "market": "내수", "points": [{"period": "6M", "assays": {"X": 102.9, "pH": 7.5}}]}]
        issues = []
        self.assertEqual(R._flag_outliers(logs, issues, lambda *a: None), 1)
        self.assertEqual(logs[0]["points"][0]["unsure"], ["X"])
        self.assertEqual(len(issues), 1)
        self.assertNotIn("pH", str(issues))

    def test_비슷하면_그대로(self):
        logs = [{"lot": "A", "kind": "장기", "market": "내수", "points": [{"period": "6M", "assays": {"X": 103.0}}]},
                {"lot": "B", "kind": "장기", "market": "내수", "points": [{"period": "6M", "assays": {"X": 103.4}}]}]
        self.assertEqual(R._flag_outliers(logs, [], lambda *a: None), 0)


class 유효기간(unittest.TestCase):
    def test_허가증_유효기간(self):
        class D: license = {"shelf_life": "제조일로부터 36개월"}
        self.assertEqual(R._shelf_months(D()), 36)
        class E: license = None
        self.assertIsNone(R._shelf_months(E()))


if __name__ == "__main__":
    unittest.main()
