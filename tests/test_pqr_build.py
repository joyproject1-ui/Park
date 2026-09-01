"""파일 적재 → 집계 → 판정 전체 흐름 테스트 (네트워크 불필요)."""

import os
import shutil
import tempfile
import unittest

from pqr import build as build_module
from pqr import build
from pqr import schema
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


class ItemRuleTest(unittest.TestCase):
    """평가항목 규칙이 자료 출처와 어긋나지 않는지 봅니다."""

    def setUp(self):
        self.config = build.load_config()

    def test_pending_counters_come_from_the_item_own_datasets(self):
        """다른 자료의 카운터를 쓰면 그 자료가 없을 때 0 이 되어 '완료' 로 잘못 뜹니다."""
        sources = self.config["counter_sources"]
        for number, label, _hint in self.config["items"]:
            rule = self.config["item_rules"][number]
            for counter in rule["pending"]:
                self.assertIn(
                    sources.get(counter), rule["datasets"],
                    "%s(%s) 의 %s 은 %s 자료에서 나오는데 항목이 보는 자료는 %s 입니다"
                    % (number, label, counter, sources.get(counter), rule["datasets"]))


    def test_dashboard_sample_items_match_config(self):
        """대시보드 내장 샘플과 config 의 평가항목이 어긋나면 안 됩니다.

        data.js 가 없을 때(아티팩트·정적 파일) 화면은 SAMPLE_ITEMS 로 돕니다.
        config 만 고치면 그 화면은 옛 항목을 계속 보여 줍니다 — 실제로 겪은 일입니다.
        """
        import json
        import re
        page = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "docs", "pqr", "index.html")
        with open(page, encoding="utf-8") as handle:
            html = handle.read()
        block = re.search(r"const SAMPLE_ITEMS = (\[.*?\]);", html, re.S)
        self.assertIsNotNone(block, "index.html 에서 SAMPLE_ITEMS 를 찾지 못했습니다")
        sample = json.loads(block.group(1))
        expected = [list(item) for item in self.config["items"]]
        self.assertEqual(sample, expected,
                         "SAMPLE_ITEMS 가 config 의 items 와 다릅니다 — 둘을 함께 고치세요")

    def test_every_item_has_a_rule_and_every_counter_exists(self):
        numbers = [item[0] for item in self.config["items"]]
        self.assertEqual(len(numbers), len(set(numbers)), "항목 번호가 겹칩니다")
        for number in numbers:
            self.assertIn(number, self.config["item_rules"])
            self.assertIn(number, self.config["item_datasets"])
        for number, rule in self.config["item_rules"].items():
            if number.startswith("_"):
                continue
            self.assertIn(number, numbers, "쓰이지 않는 규칙: %s" % number)
            for counter in rule["pending"]:
                self.assertIn(counter, self.config["counter_sources"],
                              "counter_sources 에 없는 카운터: %s" % counter)
            for name in rule["datasets"]:
                self.assertIn(name, schema.DATASETS, "없는 자료 이름: %s" % name)
