# -*- coding: utf-8 -*-
"""공수 산출 — PQR 한 건의 시간을 생산 Lot 수 구간으로 정한다 (담당자 2026-09-16 공수표).

7-1 61~100 Lot 32h · 7-2 31~60 Lot 12h · 7-3 10~30 Lot 6h · 7-4 10 Lot 미만 4.5h · 7-5 생산 X 3h
"""
from __future__ import unicode_literals

import io
import json
import os
import re
import unittest

from pqr import build

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class 구간(unittest.TestCase):
    def setUp(self):
        self.config = build.load_config()

    def hours(self, lots):
        got = build.effort_of(lots, self.config)
        return got and got["hours"]

    def test_공수표_그대로(self):
        self.assertEqual(self.hours(100), 32)
        self.assertEqual(self.hours(61), 32)
        self.assertEqual(self.hours(60), 12)
        self.assertEqual(self.hours(31), 12)
        self.assertEqual(self.hours(30), 6)
        self.assertEqual(self.hours(10), 6)
        self.assertEqual(self.hours(9), 4.5)
        self.assertEqual(self.hours(1), 4.5)
        self.assertEqual(self.hours(0), 3)

    def test_구간_위는_맨_위_구간을_쓰고_표시한다(self):
        got = build.effort_of(140, self.config)
        self.assertEqual(got["hours"], 32)
        self.assertTrue(got["over"])
        self.assertFalse(build.effort_of(100, self.config)["over"])

    def test_Lot_수를_모르면_None(self):
        self.assertIsNone(build.effort_of(None, self.config))
        self.assertIsNone(build.effort_of("x", self.config))

    def test_소수_Lot_은_반올림(self):
        self.assertEqual(build.effort_of(9.6, self.config)["lots"], 10)

    def test_화면_샘플_공수표는_config_와_같다(self):
        html = io.open(os.path.join(HERE, "docs", "pqr", "index.html"), encoding="utf-8").read()
        block = re.search(r"const SAMPLE_EFFORT_TABLE = (\{.*?\n\});", html, re.S)
        self.assertIsNotNone(block, "index.html 에서 SAMPLE_EFFORT_TABLE 을 찾지 못했습니다")
        sample = json.loads(block.group(1))
        expected = {k: v for k, v in self.config["effort_table"].items() if not k.startswith("_")}
        self.assertEqual(sample, expected, "SAMPLE_EFFORT_TABLE 이 config 의 effort_table 과 다릅니다")


if __name__ == "__main__":
    unittest.main()
