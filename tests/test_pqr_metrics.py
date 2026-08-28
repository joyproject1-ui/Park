"""공정능력 · 경향 계산 테스트."""

import unittest

from pqr import metrics


class CapabilityTest(unittest.TestCase):
    def test_two_sided_cp_cpk(self):
        values = [100.0] * 10 + [98.0, 102.0]
        result = metrics.capability(values, 95, 105)
        self.assertIsNotNone(result["cp"])
        self.assertLessEqual(result["cpk"], result["cp"])
        self.assertEqual(result["n"], 12)

    def test_one_sided_has_no_cp(self):
        result = metrics.capability([92.0, 94.0, 96.0, 91.0], 80, None)
        self.assertIsNone(result["cp"])
        self.assertIsNotNone(result["cpk"])

    def test_verdict_thresholds(self):
        """사내 규정: Cpk 1 이상 공정능력 충분, 미만 부족."""
        tight = metrics.capability([100.0, 100.1, 99.9, 100.0, 100.05], 95, 105)
        self.assertEqual(tight["verdict"], metrics.VERDICT_SUFFICIENT)
        wide = metrics.capability([96.0, 104.0, 97.0, 103.0, 100.0], 95, 105)
        self.assertEqual(wide["verdict"], metrics.VERDICT_INSUFFICIENT)

    def test_minimum_lots_rule(self):
        """10 Lot 미만이면 계산하지 않고 이유를 남깁니다."""
        result = metrics.capability([7.0, 7.0, 7.09], 6.8, 7.2, min_lots=10)
        self.assertEqual(result["verdict"], metrics.VERDICT_NOT_APPLICABLE)
        self.assertIsNone(result["cpk"])
        self.assertIn("10 Lot", result["reason"])

    def test_two_sided_only_rule(self):
        """상·하한이 모두 있는 항목만 적용합니다."""
        one_sided = metrics.capability([0.04] * 12, None, 0.5,
                                       min_lots=10, two_sided_only=True)
        self.assertEqual(one_sided["verdict"], metrics.VERDICT_NOT_APPLICABLE)
        self.assertIn("상·하한", one_sided["reason"])

    def test_rules_pass_when_conditions_are_met(self):
        values = [100.0, 100.5, 99.5, 100.2, 99.8, 100.1, 99.9, 100.3, 99.7, 100.4,
                  100.0, 99.6]
        result = metrics.capability(values, 95, 105, min_lots=10, two_sided_only=True)
        self.assertEqual(result["verdict"], metrics.VERDICT_SUFFICIENT)
        self.assertIsNotNone(result["cpk"])

    def test_not_computable_cases(self):
        self.assertIn("2건 미만", metrics.capability([100.0], 95, 105)["reason"])
        self.assertEqual(metrics.capability([100.0, 101.0])["reason"], "규격이 없음")
        self.assertEqual(metrics.capability([100.0, 100.0], 95, 105)["reason"], "표준편차가 0")

    def test_none_values_are_ignored(self):
        result = metrics.capability([100.0, None, 101.0], 95, 105)
        self.assertEqual(result["n"], 2)


class TrendTest(unittest.TestCase):
    def test_out_of_spec_positions(self):
        self.assertEqual(metrics.out_of_spec([94.0, 99.0, 106.0], 95, 105), [0, 2])

    def test_single_large_excursion_masks_itself_under_3sigma(self):
        """이탈값 하나가 표준편차를 부풀려 평균±3σ 안에 숨는 성질을 확인합니다."""
        values = [100.0, 100.2, 99.8, 100.1, 99.9, 120.0]
        self.assertEqual(metrics.control_limits(values)["out_of_trend"], [])

    def test_robust_limits_catch_the_masked_excursion(self):
        values = [100.0, 100.2, 99.8, 100.1, 99.9, 120.0]
        self.assertIn(5, metrics.robust_limits(values)["out_of_trend"])

    def test_out_of_trend_combines_both_rules(self):
        values = [100.0, 100.2, 99.8, 100.1, 99.9, 120.0]
        _, flagged = metrics.out_of_trend(values)
        self.assertEqual(flagged, {5: "MAD"})

    def test_out_of_trend_is_empty_for_stable_data(self):
        _, flagged = metrics.out_of_trend([100.0, 100.2, 99.8, 100.1, 99.9, 100.3])
        self.assertEqual(flagged, {})

    def test_stability_downward_trend_is_adverse(self):
        result = metrics.stability_trend([0, 3, 6, 9, 12], [99.8, 99.1, 98.4, 97.9, 97.2],
                                         lsl=95, usl=105)
        self.assertTrue(result["adverse"])
        self.assertLess(result["slope"], 0)
        self.assertEqual(result["limit"], 95)

    def test_stable_data_is_not_adverse(self):
        result = metrics.stability_trend([0, 3, 6, 9, 12], [99.8, 99.9, 99.7, 99.8, 99.9],
                                         lsl=95, usl=105)
        self.assertFalse(result["adverse"])

    def test_trend_needs_three_points(self):
        result = metrics.stability_trend([0, 3], [99.0, 98.0], lsl=95)
        self.assertIsNone(result["slope"])
        self.assertFalse(result["adverse"])

    def test_round_all_handles_nested(self):
        rounded = metrics.round_all({"a": [1.23456, {"b": 2.34567}]}, 2)
        self.assertEqual(rounded, {"a": [1.23, {"b": 2.35}]})


if __name__ == "__main__":
    unittest.main()
