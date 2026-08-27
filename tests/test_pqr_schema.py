"""열 이름 매핑 · 값 정리 테스트."""

import datetime
import unittest

from pqr.schema import detect_dataset, normalize, parse_date, parse_number


class ParseTest(unittest.TestCase):
    def test_date_formats(self):
        expected = datetime.date(2026, 3, 11)
        for value in ["2026-03-11", "2026.03.11", "2026/03/11", "20260311",
                      "2026년 3월 11일", "2026-03-11 14:22:01", "2026-03-11T00:00:00"]:
            self.assertEqual(parse_date(value), expected, value)

    def test_date_excel_serial(self):
        self.assertEqual(parse_date("45700"), datetime.date(2025, 2, 12))

    def test_date_invalid(self):
        self.assertIsNone(parse_date(""))
        self.assertIsNone(parse_date("미정"))

    def test_number_formats(self):
        self.assertEqual(parse_number("1,234.5"), 1234.5)
        self.assertEqual(parse_number("99.4 %"), 99.4)
        self.assertEqual(parse_number("<0.05"), 0.05)
        self.assertEqual(parse_number(0), 0.0)
        self.assertIsNone(parse_number(""))
        self.assertIsNone(parse_number("적합"))

    def test_detect_dataset(self):
        self.assertEqual(detect_dataset("일탈대장_2026.xlsx"), "deviations")
        self.assertEqual(detect_dataset("batch_coa.csv"), "batches")
        self.assertEqual(detect_dataset("안정성.csv"), "stability")
        self.assertEqual(detect_dataset("qualification_적격성.csv"), "qualification")
        self.assertIsNone(detect_dataset("메모.txt"))


class NormalizeTest(unittest.TestCase):
    def test_korean_and_english_aliases(self):
        rows = [{"품목코드": "hp-101", "제조번호": "A1", "시험항목": "함량",
                 "결과": "99.4", "Lower limit": "95", "USL": "105"}]
        records, issues = normalize(rows, "batches")
        self.assertEqual(records[0]["product_code"], "HP-101")
        self.assertEqual(records[0]["batch_no"], "A1")
        self.assertEqual(records[0]["lsl"], 95.0)
        self.assertEqual(records[0]["usl"], 105.0)
        self.assertFalse([i for i in issues if i["level"] == "error"])

    def test_zero_timepoint_is_kept(self):
        """안정성 0개월 시점은 값이 있는 것이므로 버리면 안 됩니다."""
        rows = [{"제품코드": "HP-1", "시험항목": "함량", "시점": "0", "결과값": "99.8"}]
        records, issues = normalize(rows, "stability")
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["timepoint"], 0.0)
        self.assertFalse([i for i in issues if i["level"] == "error"])

    def test_missing_required_field_is_reported(self):
        rows = [{"제품코드": "", "시험항목": "함량", "배치번호": "A1"}]
        records, issues = normalize(rows, "batches", "x.csv")
        self.assertEqual(records, [])
        self.assertEqual(issues[0]["level"], "error")
        self.assertIn("product_code", issues[0]["field"])

    def test_folder_supplies_product_code(self):
        rows = [{"배치번호": "A1", "시험항목": "함량", "결과값": "99"}]
        records, _ = normalize(rows, "batches", "f.csv", default_product_code="hp-201")
        self.assertEqual(records[0]["product_code"], "HP-201")

    def test_folder_mismatch_warns_but_keeps_file_value(self):
        rows = [{"제품코드": "HP-999", "배치번호": "A1", "시험항목": "함량"}]
        records, issues = normalize(rows, "batches", "f.csv", default_product_code="HP-201")
        self.assertEqual(records[0]["product_code"], "HP-999")
        warnings = [i for i in issues if i["level"] == "warning"]
        self.assertEqual(len(warnings), 1)
        self.assertIn("HP-201", warnings[0]["message"])

    def test_note_row_below_the_table_is_skipped(self):
        """표 아래 안내 문장이 제품 한 건으로 잡히면 안 됩니다."""
        rows = [
            {"제품코드": "HP-110", "제품명": "레보클점안액"},
            {"제품코드": "노란 칸은 채우거나 확인해야 하는 값입니다. 제품을 추가할 때는 "
                     "3행부터 같은 형식으로 이어서 적으면 됩니다.", "제품명": ""},
        ]
        records, issues = normalize(rows, "products", "m.xlsx")
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["product_code"], "HP-110")
        skipped = [i for i in issues if "설명문" in i["message"]]
        self.assertEqual(len(skipped), 1)

    def test_short_single_field_row_is_kept(self):
        """코드만 적힌 행은 정상 자료이므로 남겨야 합니다."""
        records, _ = normalize([{"제품코드": "HP-110"}], "products", "m.xlsx")
        self.assertEqual(len(records), 1)

    def test_unmapped_columns_are_reported_as_info(self):
        rows = [{"제품코드": "HP-1", "배치번호": "A1", "시험항목": "함량", "비고": "메모"}]
        _, issues = normalize(rows, "batches", "f.csv")
        info = [i for i in issues if i["level"] == "info"]
        self.assertEqual(len(info), 1)
        self.assertIn("비고", info[0]["field"])


if __name__ == "__main__":
    unittest.main()
