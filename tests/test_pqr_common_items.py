# -*- coding: utf-8 -*-
"""공통 자료와 자동 확인 항목 (담당자 2026-09-14).

· 6 · 8.1.1 · 8.1.2 · 8.1.3 · 10.1 · 10.2 · 10.3-5 는 공통 자료 — '공통' 폴더에 한 번
  올리면 모든 제품의 수집 현황에 들어온다.
· 3항 허가증은 식약처 API 로 프로그램이 대신 확인한다 — 올릴 것이 없으므로 'a' 로 두고
  수집률에서도 뺀다.
"""
from __future__ import unicode_literals

import os
import tempfile
import unittest

from pqr import build


class 설정(unittest.TestCase):
    def test_6항_이름은_제조내역_ERP(self):
        items = {row[0]: row for row in build.load_config()["items"]}
        self.assertEqual(items["6"][1], "제조내역")
        self.assertEqual(items["6"][2], "ERP")

    def test_공통_항목_일곱_가지(self):
        config = build.load_config()
        self.assertEqual(set(config["common_items"]),
                         {"6", "8.1.1", "8.1.2", "8.1.3", "10.1", "10.2", "10.3-5"})

    def test_3항은_자동_확인(self):
        self.assertEqual(build.load_config()["auto_items"], ["3"])


class 공통_폴더(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp()
        self.items = build.load_config()["items"]
        self.common = build.load_config()["common_items"]

    def _put(self, folder, name):
        os.makedirs(os.path.join(self.root, folder), exist_ok=True)
        with open(os.path.join(self.root, folder, name), "w", encoding="utf-8") as handle:
            handle.write("x")

    def test_공통_폴더_이름을_알아본다(self):
        self.assertTrue(build.is_common_folder("공통"))
        self.assertTrue(build.is_common_folder("common"))
        self.assertFalse(build.is_common_folder("QC1-5059 올로원스점안액"))

    def test_공통_폴더의_항_번호_파일을_모은다(self):
        self._put("공통", "6. 제조내역.xlsx")
        self._put("공통", "10.1 PV Validation Master File.xlsx")
        got = build.collect_common_files(self.root, self.items, self.common)
        self.assertIn("6", got)
        self.assertIn("10.1", got)

    def test_공통으로_정하지_않은_항목은_나눠_주지_않는다(self):
        """13항 안정성은 제품마다 다르다 — 공통 폴더에 있어도 모두에게 주면 안 된다."""
        self._put("공통", "13. 안정성 시험.xlsx")
        got = build.collect_common_files(self.root, self.items, self.common)
        self.assertNotIn("13", got)

    def test_제품_폴더는_공통으로_읽지_않는다(self):
        self._put("QC1-5059 올로원스점안액", "6. 제조내역.xlsx")
        self.assertEqual(build.collect_common_files(self.root, self.items, self.common), {})

    def test_모든_제품에_더해진다(self):
        per = {"7": ["7. 수율현황표.xlsx"]}
        merged = build.with_common_files(per, {"6": ["6. 제조내역.xlsx"]})
        self.assertEqual(sorted(merged), ["6", "7"])
        self.assertEqual(per, {"7": ["7. 수율현황표.xlsx"]})      # 원본은 그대로

    def test_같은_이름을_두_번_넣지_않는다(self):
        merged = build.with_common_files({"6": ["6. 제조내역.xlsx"]}, {"6": ["6. 제조내역.xlsx"]})
        self.assertEqual(merged["6"], ["6. 제조내역.xlsx"])


class 자동_확인_항목(unittest.TestCase):
    def _states(self, item_files=None):
        config = build.load_config()
        context = {"has": {}, "item_files": item_files or {}}
        return dict(zip([row[0] for row in config["items"]],
                        build._checks(context, config)))

    def test_3항은_자료가_없어도_a(self):
        self.assertEqual(self._states()["3"], "a")

    def test_3항은_자료가_있어도_a(self):
        self.assertEqual(self._states({"3": ["3. 허가증.pdf"]})["3"], "a")

    def test_다른_항목은_그대로_n(self):
        self.assertEqual(self._states()["7"], "n")

    def test_수집률에서_3항을_뺀다(self):
        config = build.load_config()
        self.assertIn("3", build.optional_items(config))


if __name__ == "__main__":
    unittest.main()
