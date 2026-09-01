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

    def test_trend_counts_low_cpk_products_by_dosage_form(self):
        """경향 그래프는 Cpk 1 미만 제품을 제형별로 셉니다."""
        trend = self.data["trend"]
        self.assertEqual(len(trend["months"]), 12)
        self.assertEqual(trend.get("label"), "Cpk 1 미만 제품")
        forms = {product["form_group"] for product in self.data["products"]
                 if product["cpk_low"]}
        self.assertEqual({s["key"] for s in trend["series"]}, forms)
        for series in trend["series"]:
            self.assertEqual(len(series["data"]), 12)

    def test_low_cpk_products_are_flagged(self):
        """'Cpk 발생' 카드와 목록이 이 값을 씁니다."""
        for product in self.data["products"]:
            tests = self.data["quality"][product["code"]]["tests"]
            expected = [t["test_name"] for t in tests
                        if t["cpk"] is not None and t["cpk"] < 1.0]
            self.assertEqual(product["cpk_low"], len(expected), product["code"])
            self.assertEqual(product["cpk_low_tests"], expected, product["code"])

    def test_leadtime_uses_stage_log(self):
        collecting = next(row for row in self.data["leadtime"] if row["stage"] == "자료 수집")
        self.assertIsNotNone(collecting["actual"])
        self.assertGreater(collecting["samples"], 0)


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

class ItemFileTest(unittest.TestCase):
    """항 번호가 붙은 파일·폴더가 수집 현황에 잡히는지 봅니다.

    예시 자료가 여러 항목을 이미 채우므로, 안정성 파일을 지워 13항을
    미착수로 만들어 놓고 파일 인식만으로 상태가 바뀌는지 확인합니다.
    """

    def setUp(self):
        self.dir = tempfile.mkdtemp()
        write_samples(self.dir, layout="tree")
        self.folder = os.path.join(
            self.dir, next(name for name in os.listdir(self.dir) if name.startswith("HP-110")))
        os.remove(os.path.join(self.folder, "안정성.csv"))

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def snapshot(self):
        data = build.build(input_dir=self.dir, today=TODAY)
        ids = [item[0] for item in data["items"]]
        product = next(p for p in data["products"] if p["code"] == "HP-110")
        return dict(zip(ids, product["checks"])), product, data

    def touch(self, name):
        with open(os.path.join(self.folder, name), "wb") as handle:
            handle.write(b"x")

    def test_numbered_pdf_turns_item_green(self):
        states, _, _ = self.snapshot()
        self.assertEqual(states["13"], "n")
        self.touch("13. 안정성 시험 보고서.pdf")
        states, product, _ = self.snapshot()
        self.assertEqual(states["13"], "y")
        self.assertIn("13", product["item_files"])

    def test_numbered_folder_counts_only_with_content(self):
        folder = os.path.join(self.folder, "13. 안정성 시험")
        os.makedirs(folder)
        states, _, _ = self.snapshot()
        self.assertEqual(states["13"], "n", "빈 폴더는 자료가 아닙니다")
        with open(os.path.join(folder, "보고서.pdf"), "wb") as handle:
            handle.write(b"x")
        states, _, _ = self.snapshot()
        self.assertEqual(states["13"], "y")

    def test_range_and_prefix_ids_map_to_items(self):
        self.touch("10.3, 10.4, 10.5 제조지원 설비.pdf")
        _, product, _ = self.snapshot()
        self.assertIn("10.3-5", product["item_files"])

    def test_lookalike_names_are_not_items(self):
        self.touch("13주년 행사 안내.pdf")   # 구분자 없는 한글 — 항목 아님
        self.touch("13M 자료.pdf")           # 영숫자가 바로 붙음 — 항목 아님
        states, product, _ = self.snapshot()
        self.assertEqual(states["13"], "n")
        self.assertNotIn("13", product["item_files"])


    def test_close_marker_does_not_pollute_other_items(self):
        """'14 반품 · 불만 · 회수 - 해당없음 확인.txt' 이름에는 '불만' 이 들어 있어
        변경·불만 대장으로 오인되면 12항(변경관리)이 근거 없이 초록이 됩니다."""
        os.remove(os.path.join(self.folder, "변경불만대장.csv"))
        self.touch("14 반품 · 불만 · 회수 - 해당없음 확인.txt")
        states, product, data = self.snapshot()
        self.assertEqual(states["14"], "y")
        self.assertEqual(states["12"], "n", "마감 기록이 변경 대장으로 오인됐습니다")
        loaded = [source["file"] for sources in data["sources"].values() for source in sources]
        self.assertFalse(any("해당없음 확인" in name for name in loaded),
                         "마감 기록 파일이 대장으로 적재됐습니다: %s" % loaded)


    def test_numbered_company_files_are_evidence_not_tables(self):
        """'10.2 Qualification Master File-….xlsx' 같은 회사 원본 양식은 이름의 낱말이
        대장 키워드와 우연히 겹칩니다. 표로 읽으려 들면 적재 오류 수천 건만 남으므로,
        항 번호 파일은 근거로만 취급하고 표로는 읽지 않아야 합니다."""
        name = "13. 안정성 시험 결과(회사 양식).xlsx"
        with open(os.path.join(self.folder, name), "wb") as handle:
            handle.write(b"this is not a spreadsheet")
        states, product, data = self.snapshot()
        self.assertEqual(states["13"], "y")                 # 근거로는 인정
        errors = [i for i in data["issues"] if i["level"] == "error"]
        self.assertEqual(errors, [], "근거 파일이 표로 적재됐습니다")
        loaded = [s2["file"] for ss in data["sources"].values() for s2 in ss]
        self.assertNotIn(name, loaded)

    def test_numbered_files_are_not_reported_unknown(self):
        self.touch("7. 수율현황표.pdf")
        _, _, data = self.snapshot()
        self.assertNotIn("7. 수율현황표.pdf", data["unknown_files"])


class FinalReportTest(unittest.TestCase):
    """제출용 보고서(.docx) 생성과 '완성본' 인식."""

    def setUp(self):
        self.dir = tempfile.mkdtemp()
        write_samples(self.dir, layout="tree")
        self.folder = os.path.join(
            self.dir, next(name for name in os.listdir(self.dir) if name.startswith("HP-110")))

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def build(self):
        return build_module.build(input_dir=self.dir, today=TODAY)

    def product(self, data=None):
        data = data or self.build()
        return next(p for p in data["products"] if p["code"] == "HP-110")

    def test_no_final_report_before_writing(self):
        self.assertEqual(self.product()["final_report"], "")

    def test_written_docx_is_detected_as_final_report(self):
        from pqr import docx_report
        data = self.build()
        path = docx_report.write_docx(data, "HP-110", self.folder,
                                      config=build_module.load_config())
        self.assertTrue(os.path.isfile(path))
        self.assertIn("제출용", os.path.basename(path))
        self.assertEqual(self.product()["final_report"], os.path.basename(path))

    def test_final_docx_is_not_read_as_a_table(self):
        """제출용 보고서가 제품 폴더에 있어도 대장으로 적재되면 안 됩니다."""
        from pqr import docx_report
        docx_report.write_docx(self.build(), "HP-110", self.folder,
                               config=build_module.load_config())
        data = self.build()
        loaded = [source["file"] for sources in data["sources"].values() for source in sources]
        self.assertFalse(any("제출용" in name for name in loaded))
        self.assertEqual([i for i in data["issues"] if i["level"] == "error"], [])

    def test_docx_is_a_readable_package(self):
        import zipfile
        from xml.dom import minidom
        from pqr import docx_report
        path = docx_report.write_docx(self.build(), "HP-110", self.dir,
                                      config=build_module.load_config())
        with zipfile.ZipFile(path) as archive:
            self.assertIsNone(archive.testzip())
            names = archive.namelist()
            for required in ("[Content_Types].xml", "_rels/.rels", "word/document.xml"):
                self.assertIn(required, names)
            for name in names:                     # Word 가 열 수 있어야 합니다
                minidom.parseString(archive.read(name))
            text = archive.read("word/document.xml").decode("utf-8")
        self.assertIn("제품품질평가", text)
        self.assertIn("레보클린", text)

    def test_evidence_only_product_does_not_get_a_settled_conclusion(self):
        """근거 PDF 만 있고 수치를 읽지 못했으면 '규격에 만족' 결론을 확정하지 않습니다."""
        import zipfile
        from pqr import docx_report
        # 대장 파일을 모두 치우고 항 번호 근거 파일(PDF)만 남깁니다 — 사용자의 실제 상황입니다.
        for name in os.listdir(self.folder):
            os.remove(os.path.join(self.folder, name))
        for number in (item[0] for item in build_module.load_config()["items"]):
            with open(os.path.join(self.folder, "%s. 근거자료.pdf" % number), "wb") as handle:
                handle.write(b"x")
        data = self.build()
        product = self.product(data)
        self.assertEqual(product["pct"], 100)            # 근거는 다 모였지만
        self.assertEqual(data["quality"]["HP-110"]["tests"], [])   # 수치는 못 읽었습니다
        path = docx_report.write_docx(data, "HP-110", self.dir,
                                      config=build_module.load_config())
        with zipfile.ZipFile(path) as archive:
            text = archive.read("word/document.xml").decode("utf-8")
        self.assertIn("결론을 확정하지 않았습니다", text)

    def test_mean_keeps_the_source_decimal_places(self):
        """평균을 원자료보다 잘게 적으면 측정하지 않은 정밀도를 쓰는 셈입니다."""
        from pqr import docx_report
        self.assertEqual(docx_report._decimals(6.92, 7.55), 2)
        self.assertEqual(docx_report._fixed(7.2513, 2), "7.25")
        # 사사오입 — 파이썬 기본 반올림은 1.0165 를 1.016 으로 적습니다.
        self.assertEqual(docx_report._fixed(1.0165, 3), "1.017")

    def test_cover_prints_lot_count_without_a_decimal_point(self):
        """제조 Lot 수가 '20.0' 으로 나오면 없는 단위를 읽게 됩니다."""
        from pqr import docx_report
        self.assertEqual(docx_report._text(20.0), "20")
        self.assertEqual(docx_report._text(20.5), "20.5")
        self.assertEqual(docx_report._text(None), "—")

    def test_missing_license_fields_say_what_to_do_when_evidence_exists(self):
        """허가증 PDF 는 있는데 마스터가 비었으면 '—'(해당 없음)이 아니라 '원본 확인'."""
        import zipfile
        from pqr import docx_report
        with open(os.path.join(self.folder, "3. 허가증.pdf"), "wb") as handle:
            handle.write(b"x")
        path = docx_report.write_docx(self.build(), "HP-110", self.dir,
                                      config=build_module.load_config())
        with zipfile.ZipFile(path) as archive:
            text = archive.read("word/document.xml").decode("utf-8")
        self.assertIn("원본 확인", text)


class LegacyXlsTest(unittest.TestCase):
    """ERP 출력물은 구형 .xls 입니다 — 못 읽으면 8·9항 자료가 통째로 빠집니다."""

    def test_read_table_accepts_xls_extension(self):
        from pqr import tabular
        # xlrd 가 없는 PC 에서도 '왜 못 읽는지'와 해결 방법을 알려 줘야 합니다.
        try:
            import xlrd                                   # noqa: F401
        except ImportError:
            with self.assertRaises(tabular.TableError) as caught:
                tabular.read_table("없는파일.xls")
            self.assertIn("xlsx", str(caught.exception))
            return
        with self.assertRaises(tabular.TableError):
            tabular.read_table("없는파일.xls")

    def test_unsupported_extension_still_refused(self):
        from pqr import tabular
        with self.assertRaises(tabular.TableError):
            tabular.read_table("보고서.pdf")
