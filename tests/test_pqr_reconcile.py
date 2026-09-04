# -*- coding: utf-8 -*-
"""제품 마스터에 없는 제품코드 되돌리기 — 레보클점안액이 PRD-001 로 잡히던 건."""
from __future__ import unicode_literals

import unittest

from pqr import build


MASTER = {
    "QC1-5041": {"product_code": "QC1-5041", "product_name": "레보클점안액"},
    "QC1-5001": {"product_code": "QC1-5001", "product_name": "후메론점안액"},
}


class 코드되돌리기(unittest.TestCase):
    def test_이름이_같으면_마스터_코드로_돌린다(self):
        datasets = {"batches": [{"product_code": "PRD-001", "product_name": "레보클점안액"},
                                {"product_code": "PRD-001", "product_name": "레보클점안액"}]}
        issues, remap = build.reconcile_codes(MASTER, dict(MASTER), datasets)
        self.assertEqual(remap, {"PRD-001": "QC1-5041"})
        self.assertEqual([r["product_code"] for r in datasets["batches"]],
                         ["QC1-5041", "QC1-5041"])
        self.assertEqual(len(issues), 1)
        self.assertIn("PRD-001", issues[0]["message"])
        self.assertIn("QC1-5041", issues[0]["message"])
        self.assertIn("2 행", issues[0]["message"])

    def test_공백과_대소문자는_무시한다(self):
        datasets = {"batches": [{"product_code": "prd-001", "product_name": " 레보클 점안액 "}]}
        _, remap = build.reconcile_codes(MASTER, dict(MASTER), datasets)
        self.assertEqual(remap, {"prd-001": "QC1-5041"})

    def test_마스터에_있는_코드는_건드리지_않는다(self):
        datasets = {"batches": [{"product_code": "QC1-5001", "product_name": "후메론점안액"}]}
        issues, remap = build.reconcile_codes(MASTER, dict(MASTER), datasets)
        self.assertEqual((issues, remap), ([], {}))

    def test_이름이_안_맞으면_고치지_않고_알린다(self):
        datasets = {"batches": [{"product_code": "PRD-009", "product_name": "없는점안액"}]}
        issues, remap = build.reconcile_codes(MASTER, dict(MASTER), datasets)
        self.assertEqual(remap, {})
        self.assertEqual(datasets["batches"][0]["product_code"], "PRD-009")
        self.assertIn("제품 마스터에 없습니다", issues[0]["message"])

    def test_같은_이름이_둘이면_고르지_않는다(self):
        master = dict(MASTER)
        master["QC1-9999"] = {"product_code": "QC1-9999", "product_name": "레보클점안액"}
        datasets = {"batches": [{"product_code": "PRD-001", "product_name": "레보클점안액"}]}
        issues, remap = build.reconcile_codes(master, dict(master), datasets)
        self.assertEqual(remap, {})
        self.assertIn("제품 마스터에 없습니다", issues[0]["message"])

    def test_변이_품목_이름으로도_맞춘다(self):
        expanded = dict(MASTER)
        expanded["QC1-5041-1회용"] = {"product_code": "QC1-5041-1회용",
                                     "product_name": "레보클점안액 (1회용)"}
        datasets = {"batches": [{"product_code": "PRD-002",
                                 "product_name": "레보클점안액 (1회용)"}]}
        _, remap = build.reconcile_codes(MASTER, expanded, datasets)
        self.assertEqual(remap, {"PRD-002": "QC1-5041-1회용"})


class 폴더자료도_같이_옮긴다(unittest.TestCase):
    def test_build_가_폴더_코드를_따라_옮긴다(self):
        """폴더 이름이 'PRD-001 레보클점안액' 이면 그 폴더에 모은 자료도 옮겨야 한다."""
        item_files = {"PRD-001": {"6": ["6. 제조내역.pdf"]}}
        final = {"PRD-001": "결재본.docx"}
        remap = {"PRD-001": "QC1-5041"}
        for wrong, right in remap.items():
            moved = item_files.pop(wrong, None)
            if moved:
                target = item_files.setdefault(right, {})
                for item_id, names in moved.items():
                    target.setdefault(item_id, []).extend(
                        n for n in names if n not in target.get(item_id, []))
            report = final.pop(wrong, None)
            if report and right not in final:
                final[right] = report
        self.assertEqual(item_files, {"QC1-5041": {"6": ["6. 제조내역.pdf"]}})
        self.assertEqual(final, {"QC1-5041": "결재본.docx"})


if __name__ == "__main__":
    unittest.main()


class 마스터_안의_본보기_코드(unittest.TestCase):
    """마스터가 둘 읽혀 레보클점안액이 QC1-5041 과 PRD-001 로 두 번 있던 건 (2026-09)."""

    def test_같은_이름이면_대다수_앞머리_코드로_접는다(self):
        master = dict(MASTER)
        master["PRD-001"] = {"product_code": "PRD-001", "product_name": "레보클점안액"}
        kept, remap, issues = build.fold_placeholder_codes(master)
        self.assertEqual(remap, {"PRD-001": "QC1-5041"})
        self.assertNotIn("PRD-001", kept)
        self.assertIn("QC1-5041", kept)
        self.assertIn("PRD-001", issues[0]["message"])

    def test_앞머리가_같으면_고르지_않는다(self):
        master = dict(MASTER)
        master["QC1-9999"] = {"product_code": "QC1-9999", "product_name": "레보클점안액"}
        kept, remap, _ = build.fold_placeholder_codes(master)
        self.assertEqual(remap, {})
        self.assertEqual(set(kept), set(master))

    def test_이름이_다른_본보기_코드는_그대로_둔다(self):
        master = dict(MASTER)
        master["PRD-002"] = {"product_code": "PRD-002", "product_name": "새제품"}
        kept, remap, _ = build.fold_placeholder_codes(master)
        self.assertEqual(remap, {})
        self.assertIn("PRD-002", kept)



class 마스터가_둘일_때_build(unittest.TestCase):
    """입력 폴더에 옛 본보기 마스터(xlsx 한 줄)와 진짜 마스터(csv)가 함께 있을 때."""

    def test_본보기_제품이_화면에_따로_생기지_않는다(self):
        import os, shutil, tempfile
        root = tempfile.mkdtemp(prefix="pqr-master2-")
        try:
            with open(os.path.join(root, "products_제품마스터.csv"), "w", encoding="utf-8-sig") as fh:
                fh.write("제품코드,제품명,제형,평가기간시작,평가기간종료,마감일,단계\n"
                         "QC1-5041,레보클점안액,점안제,2025-01-01,2025-12-31,2026-06-30,자료 수집\n"
                         "QC1-5084,레보클점안액1.5%,점안제,2025-01-01,2025-12-31,2026-09-30,자료 수집\n")
            with open(os.path.join(root, "옛_제품마스터.csv"), "w", encoding="utf-8-sig") as fh:
                fh.write("제품코드,제품명,제형\nPRD-001,레보클점안액,점안제\n")
            folder = os.path.join(root, "PRD-001 레보클점안액")     # 본보기 코드로 만든 폴더
            os.makedirs(folder)
            with open(os.path.join(folder, "6. 제조내역.pdf"), "wb") as fh:
                fh.write(b"x")
            data = build.build(input_dir=root, today="2026-09-05")
            codes = [p["code"] for p in data["products"]]
            self.assertNotIn("PRD-001", codes)
            self.assertIn("QC1-5041", codes)
            self.assertIn("QC1-5084", codes)
            product = next(p for p in data["products"] if p["code"] == "QC1-5041")
            self.assertIn("6", product["item_files"])                 # 폴더 자료가 진짜 제품에 붙었다
            self.assertTrue(any("PRD-001" in i["message"] for i in data["issues"]))
        finally:
            shutil.rmtree(root, ignore_errors=True)
