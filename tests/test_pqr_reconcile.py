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
