"""허용기준 문구 해석 — 결재본(PQR26-5-NJSE2)에 실제로 적힌 문구로 확인합니다.

담당자 규칙: 한쪽만 있으면 그 규격, 허가와 자가를 함께 관리하면 허가 규격.
"""

import unittest

from pqr import spec


class ParseTest(unittest.TestCase):

    def test_simple_range(self):
        self.assertEqual(spec.parse("자가) 6.2 ~ 6.7"), {"자가": (6.2, 6.7)})

    def test_at_most_and_at_least(self):
        self.assertEqual(spec.parse("자가) 10CFU/100mL 이하"), {"자가": (None, 10.0)})
        self.assertEqual(spec.parse("자가) 0.47mL 이상"), {"자가": (0.47, None)})

    def test_qualitative_criteria_have_no_limits(self):
        for text in ("자가) 무색 투명한 액", "허가) 음성(불검출)",
                     "자가) 메틸렌블루시액의 침투 없음"):
            self.assertEqual(spec.parse(text), {}, text)

    def test_parenthetical_is_not_a_limit(self):
        """'(말레인산페니라민: 3.0mg/mL)' 은 함량 규격이 아니라 설명입니다."""
        text = "허가) 90.0 ~ 110.0%\n(말레인산페니라민: 3.0mg/mL)\n자가) 92.0 ~ 108.0%"
        self.assertEqual(spec.parse(text), {"허가": (90.0, 110.0), "자가": (92.0, 108.0)})

    def test_colon_form_takes_the_judgement_value(self):
        """'300 ㎛이상/mL: 1 개 이하' 에서 기준은 300 이 아니라 1 입니다."""
        self.assertEqual(spec.parse("허가) 300 ㎛이상/mL: 1 개 이하"), {"허가": (None, 1.0)})
        self.assertEqual(spec.parse("허가) 판정값: 15.0% 이하"), {"허가": (None, 15.0)})

    def test_label_free_text_still_parses(self):
        self.assertEqual(spec.parse("0.906 ~ 1.108"), {"": (0.906, 1.108)})


class ChooseTest(unittest.TestCase):

    def test_only_one_criterion_uses_that_one(self):
        low, high, used = spec.choose("자가) 6.2 ~ 6.7")
        self.assertEqual((low, high, used), (6.2, 6.7, "자가"))

    def test_both_criteria_use_the_licensed_one(self):
        """함께 관리하면 허가 규격 — 결재본 포장 함량이 이렇게 산출됐습니다."""
        low, high, used = spec.choose("허가) 90.0 ~ 110.0%\n자가) 92.0 ~ 108.0%")
        self.assertEqual((low, high, used), (90.0, 110.0, "허가"))

    def test_empty_text_is_not_a_specification(self):
        self.assertEqual(spec.choose(""), (None, None, ""))
        self.assertEqual(spec.choose("무색 투명한 액"), (None, None, ""))

    def test_rule_changes_cpk_so_the_used_criterion_is_reported(self):
        """어느 기준을 썼는지 적지 않으면 숫자를 검증할 수 없습니다.

        결재본 포장 pH 는 자가(6.2~7.5)로 1.58 이고, 허가(6.0~8.0)면 2.89 입니다.
        """
        from pqr import metrics
        values = [6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.5,
                  6.5, 6.4, 6.5, 6.4, 6.5, 6.5, 6.5]
        both = "허가) 6.0 ~ 8.0\n자가) 6.2 ~ 7.5"
        low, high, used = spec.choose(both)
        self.assertEqual(used, "허가")
        self.assertEqual(round(metrics.capability(values, low, high, min_lots=10)["cpk"], 2),
                         2.89)
        low, high, used = spec.choose("자가) 6.2 ~ 7.5")
        self.assertEqual(round(metrics.capability(values, low, high, min_lots=10)["cpk"], 2),
                         1.58)


if __name__ == "__main__":
    unittest.main()
