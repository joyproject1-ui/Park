"""회사 결재본(PQR26 나조린점안액(1회용), 문서번호 PQR26-5-NJSE2)으로 Cpk 계산법을 고정합니다.

결재가 끝난 보고서에 적힌 Cpk 를 그대로 재현하지 못하면, 우리가 쓰는 계산이
회사 시트(HLF-QC-126-08 · -09)와 다르다는 뜻입니다. 그 상태로 보고서를 내면
심사에서 숫자가 어긋납니다. 그래서 실제 승인된 값을 회귀 테스트로 박아 둡니다.

확인된 방법 — 표본표준편차(n-1), Cpk = min((USL-x̄)/3s, (x̄-LSL)/3s).
한쪽 규격(제제균일성: 판정값 15.0% 이하)에도 산출합니다.
"""

import unittest

from pqr import metrics

# (시험항목, 결과값, LSL, USL, 결재본에 적힌 Cpk)
APPROVED = [
    # 9.2.1 조제 완료 후 — 허용기준 자가) 6.2 ~ 6.7
    ("조제 pH", [6.44, 6.44, 6.43, 6.47, 6.47, 6.43, 6.43, 6.44,
                 6.43, 6.43, 6.41, 6.45, 6.44, 6.45, 6.44], 6.2, 6.7, 5.13),
    # 9.2.2 충전 완료 후 — 자가) 6.2 ~ 7.0 (CC-250822-03 으로 기준 변경)
    ("충전 pH", [6.43, 6.45, 6.48, 6.44, 6.45, 6.46, 6.45, 6.42,
                 6.42, 6.40, 6.38, 6.42, 6.45, 6.44, 6.45], 6.2, 7.0, 3.14),
    # 9.2.4 포장 완료 후 — pH 는 자가) 6.2 ~ 7.5 로 산출했습니다
    ("포장 pH", [6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.5,
                 6.5, 6.4, 6.5, 6.4, 6.5, 6.5, 6.5], 6.2, 7.5, 1.58),
    # 함량은 허가) 90.0 ~ 110.0 % 로 산출했습니다 (자가 92 ~ 108 이 아닙니다)
    #
    # 네 번째 값 99.8 은 결재본에 적힌 값 그대로입니다. 원본 시험성적서
    # (H-2025-04-11-0063, LKY403)에는 99.9 로 적혀 있어 옮겨 적는 과정에서
    # 어긋난 것으로 보입니다. 이 테스트는 '계산 방법' 을 붙잡는 것이므로 결재본의
    # 산술을 그대로 재현해야 합니다 — 값을 99.9 로 고치면 Cpk 가 4.80 이 되어
    # 결재본과 달라집니다. 보고서를 새로 쓸 때는 성적서 값(99.9)을 쓰고,
    # 이 차이는 담당자에게 확인받으세요.
    ("함량 말레인산페니라민",
     [100.3, 99.9, 100.2, 99.8, 101.5, 101.4, 100.9, 101.7,
      100.7, 101.7, 100.6, 101.6, 101.2, 101.1, 101.2], 90.0, 110.0, 4.71),
    ("함량 나파졸린염산염",
     [103.8, 101.3, 101.5, 101.2, 101.9, 101.9, 101.3, 101.1,
      101.2, 100.9, 100.1, 100.1, 100.5, 100.6, 100.6], 90.0, 110.0, 3.23),
    # 제제균일성 — 한쪽 규격(판정값 15.0% 이하)인데도 결재본은 Cpk 를 산출했습니다
    ("제제균일성 말레인산페니라민",
     [3.0, 3.3, 2.1, 1.9, 2.5, 1.9, 2.3, 1.6,
      2.4, 2.0, 4.4, 2.3, 3.1, 5.1, 2.8], None, 15.0, 4.24),
    ("제제균일성 나파졸린염산염",
     [5.4, 3.4, 2.1, 1.9, 1.1, 2.3, 2.3, 1.4,
      2.4, 1.8, 4.4, 2.1, 3.0, 5.1, 2.8], None, 15.0, 3.16),
]


class ApprovedReportCpkTest(unittest.TestCase):

    def test_reproduces_every_cpk_in_the_approved_report(self):
        for name, values, lsl, usl, expected in APPROVED:
            result = metrics.capability(values, lsl, usl, min_lots=10)
            self.assertIsNotNone(result["cpk"], "%s: Cpk 를 산출하지 못했습니다" % name)
            self.assertAlmostEqual(
                round(result["cpk"], 2), expected, places=2,
                msg="%s: 결재본 %.2f 인데 %.4f 가 나왔습니다" % (name, expected, result["cpk"]))

    def test_one_sided_items_still_get_a_cpk(self):
        """결재본은 판정값이 상한뿐인 제제균일성에도 Cpk 를 적었습니다."""
        name, values, lsl, usl, expected = APPROVED[-1]
        result = metrics.capability(values, lsl, usl, min_lots=10, two_sided_only=False)
        self.assertEqual(round(result["cpk"], 2), expected)

    def test_config_matches_the_approved_practice(self):
        from pqr import build
        thresholds = build.load_config()["thresholds"]
        self.assertEqual(thresholds["cpk_min_lots"], 10)      # 10 Lot 이상
        self.assertFalse(thresholds["cpk_two_sided_only"],
                         "결재본은 한쪽 규격 항목에도 Cpk 를 산출했습니다")

    def test_sample_standard_deviation_not_population(self):
        """모집단 표준편차(n)로 계산하면 결재본보다 큰 Cpk 가 나옵니다."""
        name, values, lsl, usl, expected = APPROVED[0]
        self.assertEqual(round(metrics.capability(values, lsl, usl, min_lots=10)["cpk"], 2),
                         expected)
        import statistics
        population = statistics.pstdev(values)
        mean = statistics.fmean(values)
        wrong = min((usl - mean) / (3 * population), (mean - lsl) / (3 * population))
        self.assertNotEqual(round(wrong, 2), expected)


if __name__ == "__main__":
    unittest.main()
