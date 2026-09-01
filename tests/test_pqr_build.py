"""파일 적재 → 집계 → 판정 전체 흐름 테스트 (네트워크 불필요)."""

import os
import shutil
import tempfile
import unittest

from pqr import build as build_module
from pqr import build
from pqr.sample import write_samples

ITEM_IDS = [item[0] for item in build.load_config()["items"]]

TODAY = "2026-08-27"


class TreeLayoutTest(unittest.TestCase):
    """담당자가 제품 폴더에 자료를 올리는 방식."""

    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.mkdtemp()
        write_samples(cls.dir, layout="tree")
        cls.data = build_module.build(input_dir=cls.dir, today=TODAY)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.dir, ignore_errors=True)

    def test_detects_product_folders(self):
        self.assertTrue(build_module.has_product_folders(self.dir))

    def test_folder_name_supplies_product_code(self):
        """제품 폴더 안 파일에는 제품코드 열이 없지만 폴더 이름으로 채워집니다."""
        codes = {product["code"] for product in self.data["products"]}
        self.assertIn("HP-101", codes)
        self.assertEqual(len(self.data["products"]), 6)

    def test_no_ingest_errors_for_clean_input(self):
        errors = [issue for issue in self.data["issues"] if issue["level"] == "error"]
        self.assertEqual(errors, [])

    def test_sources_record_file_hash_for_audit(self):
        entry = self.data["sources"]["batches"][0]
        self.assertEqual(len(entry["sha256"]), 64)
        self.assertGreater(entry["rows"], 0)
        self.assertTrue(entry["product"])

    def test_every_product_has_twelve_checks(self):
        for product in self.data["products"]:
            self.assertEqual(len(product["checks"]), len(ITEM_IDS))
            self.assertTrue(set(product["checks"]) <= {"y", "p", "n"})

    def test_common_folder_data_applies_to_all_products(self):
        """설비 적격성은 공통 자료이므로 모든 제품의 (k) 항목이 채워집니다."""
        for product in self.data["products"]:
            self.assertNotEqual(product["checks"][ITEM_IDS.index("10.2")], "n", product["code"])

    def test_capability_is_computed_per_test(self):
        tests = self.data["quality"]["HP-201"]["tests"]
        names = {test["test_name"] for test in tests}
        self.assertIn("함량", names)
        content = next(test for test in tests if test["test_name"] == "함량")
        self.assertIsNotNone(content["cpk"])
        self.assertEqual(content["n"], content["n"])

    def test_deliberate_spec_excursion_is_detected(self):
        content = next(test for test in self.data["quality"]["HP-201"]["tests"]
                       if test["test_name"] == "함량")
        self.assertGreaterEqual(content["oos_count"], 1)
        self.assertTrue(content["oos_batches"][0]["batch_no"])

    def test_trend_has_twelve_months_and_three_series(self):
        self.assertEqual(len(self.data["trend"]["months"]), 12)
        self.assertEqual([s["key"] for s in self.data["trend"]["series"]],
                         ["일탈", "OOS/OOT", "불만"])

    def test_leadtime_uses_stage_log(self):
        planning = next(row for row in self.data["leadtime"] if row["stage"] == "계획 수립")
        self.assertIsNotNone(planning["actual"])
        self.assertGreater(planning["samples"], 0)


class MissingDataTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        write_samples(self.dir, layout="tree")
        self.folder = next(name for name in os.listdir(self.dir)
                           if name.startswith("HP-110"))
        os.remove(os.path.join(self.dir, self.folder, "안정성.csv"))

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_missing_dataset_marks_item_not_started(self):
        data = build_module.build(input_dir=self.dir, today=TODAY)
        product = next(p for p in data["products"] if p["code"] == "HP-110")
        other = next(p for p in data["products"] if p["code"] == "HP-101")
        self.assertEqual(product["checks"][ITEM_IDS.index("13")], "n")   # 13 안정성
        self.assertIn("stability", product["missing_datasets"])
        self.assertNotEqual(other["checks"][ITEM_IDS.index("13")], "n")  # 다른 제품은 영향 없음
        self.assertIn("입력 자료 1개 항목 미제출", product["reasons"])


class FlatLayoutTest(unittest.TestCase):
    """한 폴더에 모든 제품 자료를 넣는 방식도 계속 지원합니다."""

    def setUp(self):
        self.dir = tempfile.mkdtemp()
        write_samples(self.dir, layout="flat")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_flat_layout_builds_same_products(self):
        self.assertFalse(build_module.has_product_folders(self.dir))
        data = build_module.build(input_dir=self.dir, today=TODAY)
        self.assertEqual(len(data["products"]), 6)
        errors = [issue for issue in data["issues"] if issue["level"] == "error"]
        self.assertEqual(errors, [])

    def test_unknown_file_is_listed_not_fatal(self):
        with open(os.path.join(self.dir, "메모.csv"), "w", encoding="utf-8") as handle:
            handle.write("a,b\n1,2\n")
        data = build_module.build(input_dir=self.dir, today=TODAY)
        self.assertIn("메모.csv", data["unknown_files"])


if __name__ == "__main__":
    unittest.main()
