# -*- coding: utf-8 -*-
"""보고서를 쓸 때 낸 Cpk 가 화면의 'Cpk 1 미만 제품' 으로 올라온다.

담당자 2026-09-14: "PQR 작성할 때 Cpk 1 미만 제품은 여기에 제품정보가 자동 업데이트
되도록 해줘." 성적서 대장이 없어도 보고서만 쓰면 화면이 채워져야 한다.
"""
from __future__ import unicode_literals

import datetime as _dt
import json
import os
import tempfile
import unittest

from pqr import build
from pqr.engine import writer


class 결과_남기기(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.mkdtemp()

    def _save(self, cpk):
        return writer.save_cpk_result(self.folder, {"code": "QC1-5059"}, cpk,
                                      _dt.date(2026, 9, 14))

    def test_1미만_항목을_골라_적는다(self):
        got = self._save({"assay": 0.93, "particle": 1.42, "metal": None})
        self.assertEqual(got["1미만"], [["함량", 0.93]])
        self.assertEqual(got["최저"], 0.93)
        self.assertEqual(got["제품"], "QC1-5059")
        self.assertTrue(os.path.isfile(os.path.join(self.folder, writer.CPK_RESULT_FILE)))

    def test_모두_1이상이면_1미만은_빈_목록(self):
        got = self._save({"assay": 1.33, "particle": 1.42})
        self.assertEqual(got["1미만"], [])
        self.assertEqual(got["최저"], 1.33)

    def test_산출하지_않았으면_빈_값으로_남긴다(self):
        got = self._save({})
        self.assertEqual(got["Cpk"], {})
        self.assertIsNone(got["최저"])

    def test_성분이_둘이면_이름에_성분을_적는다(self):
        got = self._save({"assay/올로파타딘": 0.88})
        self.assertEqual(got["1미만"], [["함량(올로파타딘)", 0.88]])


class 화면으로_읽기(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.mkdtemp()

    def _put(self, payload):
        with open(os.path.join(self.folder, build.CPK_RESULT_FILE), "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False)

    def test_파일이_없으면_None(self):
        self.assertIsNone(build.report_cpk(self.folder))

    def test_폴더가_없어도_터지지_않는다(self):
        self.assertIsNone(build.report_cpk(""))
        self.assertIsNone(build.report_cpk(os.path.join(self.folder, "없는폴더")))

    def test_망가진_파일은_None(self):
        with open(os.path.join(self.folder, build.CPK_RESULT_FILE), "w", encoding="utf-8") as handle:
            handle.write("{망가짐")
        self.assertIsNone(build.report_cpk(self.folder))

    def test_1미만_항목과_최저값을_돌려준다(self):
        self._put({"Cpk": {"assay": 0.93, "particle": 1.42}, "1미만": [["함량", 0.93]],
                   "최저": 0.93, "작성일": "2026-09-14"})
        got = build.report_cpk(self.folder)
        self.assertEqual(got["low"], [("함량", 0.93)])
        self.assertEqual(got["min"], 0.93)
        self.assertEqual(got["written"], "2026-09-14")

    def test_1미만이_없으면_빈_목록이되_None_은_아니다(self):
        """'아직 안 썼다' 와 '썼는데 1 미만이 없다' 는 화면에서 달리 보여야 한다."""
        self._put({"Cpk": {"assay": 1.33}, "1미만": [], "최저": 1.33})
        got = build.report_cpk(self.folder)
        self.assertIsNotNone(got)
        self.assertEqual(got["low"], [])
        self.assertEqual(got["min"], 1.33)


if __name__ == "__main__":
    unittest.main()
