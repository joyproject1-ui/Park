"""공정 시험성적서 파서 테스트.

실제 성적서는 품질 기록이므로 저장소에 두지 않습니다. 대신 같은 서식의
줄 구성을 그대로 흉내 낸 가짜 자료로 검증합니다 (PDF 읽기 없이 동작).
"""

import unittest

from pqr.coa import (CoaError, cluster_columns, parse_criteria, parse_layout,
                     parse_lines, parse_result, to_rows)

# HLF-QC-121-04 서식에서 뽑히는 줄 순서 (라벨과 값이 번갈아 나옵니다)
SAMPLE = [
    "공 정 시 험 성 적 서", "AAA1",
    "제   품   명", "가상점안액[일회용]",
    "L O T _수 량", "500,000",
    "검체채취일", "2025 / 02 / 13",
    "제 조  번 호", "AB1201",
    "시험성적서번호", "E-2025-02-13-0023",
    "검체채취량",
    "공  정  명", "조제(점안제)",
    "의 뢰 번 호", "5-2025-02-13-0023",
    "검체채취자", "홍길동",
    "제 조  일 자", "2025/02/12",
    "의  뢰  자", "이몽룡",
    "시험방법기준", "자가기준",
    "시험항목", "시 험 기 준", "시 험 결 과", "시 험 자 시험일자 적부판정",
    "성상", "무색의 투명한 액", "무색의 투명한 액", "김시험 2025.02.13", "적합",
    "pH", "6.8 ~ 7.2", "7.00", "김시험 2025.02.13", "적합",
    "삼투압", "290 ~ 330 mOsm/kg", "Av. 304mOsm/kg (302 ~ 305", "mOsm/kg)",
    "김시험 2025.02.13", "적합",
    "비중", "참고치", "1.016", "김시험 2025.02.13", "적합",
    "비       고 N/A",
    "시험완료일자", "2025/ 02/ 13",
    "확 인 일 자", "확  인  자", "판 정 일 자", "판  정  자",
    "판 정 결 과", "합격",
    "1 / 1", "韓林製藥", "HLF-QC-121-04/Rev.004",
]


def box(x0, y1, text):
    return {"x0": x0, "y1": y1, "text": text}


# 셀 안에서 줄이 바뀌는 서식(충전 공정 성적서)을 좌표로 흉내 낸 것.
# 글자 순서만으로 읽으면 칸이 어긋나므로 좌표 기반 파서가 필요합니다.
LAYOUT = sorted([
    box(184.1, 784.5, "공 정 시 험 성 적 서"),
    box(28.1, 734.9, "제 품 명"), box(99.6, 734.9, "가상점안액[일회용]"),
    box(28.1, 713.2, "제 조 번 호"), box(145.2, 712.4, "AB1201"),
    box(231.6, 713.2, "시험성적서번호"), box(310.1, 713.2, "E-2025-02-15-0001"),
    box(33.1, 691.5, "공 정 명"), box(136.6, 691.5, "점안제충전"),
    box(28.1, 669.8, "제 조 일 자"), box(136.6, 669.8, "2025/02/12"),
    box(38.2, 641.9, "시험항목"), box(146.0, 641.9, "시 험 기 준"),
    box(298.0, 641.9, "시 험 결 과"), box(409.8, 641.9, "시 험 자 시험일자 적부판정"),
    # 성상 — 기준·결과가 두 줄로 나뉘어 있습니다
    box(49.2, 623.1, "성상"),
    box(99.6, 623.1, "무색투명한 액이 무색투명한 플라"), box(251.9, 623.1, "무색투명한 액이 무색투명한 플라"),
    box(414.8, 617.2, "김시험 2025.02.14"), box(530.6, 617.2, "적합"),
    box(99.6, 614.1, "스틱 1회용 용기에 든 점안제"), box(251.9, 614.1, "스틱 1회용 용기에 든 점안제"),
    # pH
    box(53.8, 576.7, "pH"), box(99.6, 576.7, "6.8 ~ 7.2"), box(251.9, 576.7, "6.93"),
    box(414.8, 570.8, "김시험 2025.02.14"), box(530.6, 570.8, "적합"),
    # 질량·용량 — 기준이 평균/개개 두 줄
    box(35.5, 530.2, "질량·용량"), box(99.6, 530.2, "평균 : 0.50 ~ 0.57mL"),
    box(251.9, 530.2, "0.52mL"),
    box(414.8, 524.3, "고시험 2025.02.15"), box(530.6, 524.3, "적합"),
    box(99.6, 483.8, "개개 : 0.47mL 이상"), box(251.9, 483.8, "0.50 ~ 0.55mL"),
    box(30.6, 151.6, "비 고 N/A"),
    box(28.1, 129.2, "시험완료일자"), box(130.3, 129.6, "2025/ 02/ 15"),
    box(30.6, 96.2, "판 정 결 과"), box(152.4, 96.6, "합격"),
], key=lambda b: (-b["y1"], b["x0"]))


class ColumnTest(unittest.TestCase):
    def test_clusters_left_edges_into_columns(self):
        boxes = [box(35.5, 1, "a"), box(38.2, 1, "b"), box(99.6, 1, "c"),
                 box(251.9, 1, "d"), box(414.8, 1, "e"), box(530.6, 1, "f")]
        self.assertEqual(len(cluster_columns(boxes)), 5)


class LayoutTest(unittest.TestCase):
    def setUp(self):
        self.parsed = parse_layout(LAYOUT)

    def test_batch_number_comes_from_the_nearest_value_on_the_row(self):
        """'제조번호' 오른쪽 첫 값이어야 합니다 — 같은 줄의 성적서번호가 아니라."""
        self.assertEqual(self.parsed["header"]["batch_no"], "AB1201")
        self.assertEqual(self.parsed["header"]["coa_no"], "E-2025-02-15-0001")

    def test_wrapped_cells_do_not_create_extra_tests(self):
        names = [t["test_name"] for t in self.parsed["tests"]]
        self.assertEqual(names, ["성상", "pH", "질량·용량"])

    def test_wrapped_criteria_are_joined_in_order(self):
        mass = next(t for t in self.parsed["tests"] if t["test_name"] == "질량·용량")
        self.assertIn("평균 : 0.50 ~ 0.57mL", mass["criteria"])
        self.assertIn("개개 : 0.47mL 이상", mass["criteria"])

    def test_tester_line_is_not_taken_as_a_result(self):
        ph = next(t for t in self.parsed["tests"] if t["test_name"] == "pH")
        self.assertEqual(ph["result"], "6.93")
        self.assertEqual(ph["verdict"], "적합")

    def test_descriptive_result_yields_no_number(self):
        """'1회용 용기' 의 1 을 측정값으로 뽑으면 안 됩니다."""
        rows = to_rows(self.parsed)
        appearance = next(r for r in rows if r["test_name"] == "성상")
        self.assertIsNone(appearance["value"])
        self.assertEqual(appearance["verdict"], "적합")

    def test_measured_values_are_parsed(self):
        rows = to_rows(self.parsed)
        values = {r["test_name"]: r["value"] for r in rows}
        self.assertEqual(values["pH"], 6.93)
        self.assertEqual(values["질량·용량"], 0.52)


class CriteriaTest(unittest.TestCase):
    def test_two_sided_range(self):
        self.assertEqual(parse_criteria("6.8 ~ 7.2"), (6.8, 7.2, ""))

    def test_range_with_unit(self):
        self.assertEqual(parse_criteria("290 ~ 330 mOsm/kg"), (290.0, 330.0, "mOsm/kg"))

    def test_reference_only_has_no_limits(self):
        self.assertEqual(parse_criteria("참고치"), (None, None, ""))

    def test_descriptive_criteria_has_no_limits(self):
        self.assertEqual(parse_criteria("무색의 투명한 액"), (None, None, ""))


class ResultTest(unittest.TestCase):
    def test_plain_number(self):
        self.assertEqual(parse_result("7.00"), 7.0)

    def test_average_wins_over_range_in_brackets(self):
        """'Av. 304 (302 ~ 305)' 에서 대표값은 304 여야 합니다."""
        self.assertEqual(parse_result("Av. 304mOsm/kg (302 ~ 305 mOsm/kg)"), 304.0)
        self.assertEqual(parse_result("Av.303 mOsm/kg (303 ~ 304 mOsm/kg)"), 303.0)

    def test_descriptive_result_has_no_number(self):
        self.assertIsNone(parse_result("무색의 투명한 액"))
        self.assertIsNone(parse_result("무색투명한 액이 무색투명한 플라 스틱 1회용 용기"))
        self.assertIsNone(parse_result("메틸렌블루 시액의 침투 없음."))

    def test_unit_containing_digits(self):
        """'0CFU/100mL' 의 단위 안 숫자에 속으면 안 됩니다."""
        self.assertEqual(parse_result("0CFU/100mL"), 0.0)
        self.assertEqual(parse_result("304 mOsm/kg"), 304.0)
        self.assertEqual(parse_result("0.52mL"), 0.52)

    def test_empty(self):
        self.assertIsNone(parse_result(""))


class ParseTest(unittest.TestCase):
    def setUp(self):
        self.parsed = parse_lines(SAMPLE)

    def test_header_fields(self):
        header = self.parsed["header"]
        self.assertEqual(header["batch_no"], "AB1201")
        self.assertEqual(header["product_name"], "가상점안액[일회용]")
        self.assertEqual(header["stage"], "조제(점안제)")
        self.assertEqual(header["mfg_date"], "2025-02-12")
        self.assertEqual(header["coa_no"], "E-2025-02-13-0023")
        self.assertEqual(header["lot_verdict"], "합격")

    def test_all_four_tests_are_found(self):
        names = [t["test_name"] for t in self.parsed["tests"]]
        self.assertEqual(names, ["성상", "pH", "삼투압", "비중"])

    def test_tester_and_date_are_not_mistaken_for_results(self):
        ph = next(t for t in self.parsed["tests"] if t["test_name"] == "pH")
        self.assertEqual(ph["result"], "7.00")
        self.assertEqual(ph["verdict"], "적합")

    def test_wrapped_result_is_joined(self):
        osmo = next(t for t in self.parsed["tests"] if t["test_name"] == "삼투압")
        self.assertIn("302 ~ 305", osmo["result"])

    def test_rows_match_the_batches_dataset_shape(self):
        rows = to_rows(self.parsed, product_code="AAA1")
        self.assertEqual(len(rows), 4)
        ph = next(r for r in rows if r["test_name"] == "pH")
        self.assertEqual(ph["product_code"], "AAA1")
        self.assertEqual(ph["batch_no"], "AB1201")
        self.assertEqual((ph["lsl"], ph["usl"], ph["value"]), (6.8, 7.2, 7.0))
        osmo = next(r for r in rows if r["test_name"] == "삼투압")
        self.assertEqual((osmo["lsl"], osmo["usl"], osmo["value"]), (290.0, 330.0, 304.0))
        self.assertEqual(osmo["unit"], "mOsm/kg")

    def test_missing_test_is_simply_absent(self):
        """변경관리로 생략된 시험항목은 그냥 빠진 채로 읽혀야 합니다."""
        trimmed = [line for line in SAMPLE
                   if line not in ("비중", "참고치", "1.016")]
        rows = to_rows(parse_lines(trimmed))
        self.assertEqual([r["test_name"] for r in rows], ["성상", "pH", "삼투압"])

    def test_missing_table_raises(self):
        with self.assertRaises(CoaError):
            parse_lines(["제   품   명", "가상점안액"])


if __name__ == "__main__":
    unittest.main()
