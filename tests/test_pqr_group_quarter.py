# -*- coding: utf-8 -*-
"""4항 완료 기한 분기 — QC-126 그룹표 (담당자 2026-09-09 정정: 1(A) 1분기 · 1(B) 2분기 · 2·3 그룹 3분기)."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine.recipe_ointment import _group_quarter                  # noqa: E402


class 그룹별_분기(unittest.TestCase):
    def test_네_그룹_표(self):
        # 하나만 듣고 나머지를 넘겨짚지 않도록 네 그룹을 모두 못박는다 — 2026-09-08 에 1(B)를 3분기로 잘못 들었다.
        self.assertEqual(_group_quarter("1(A)"), 1)
        self.assertEqual(_group_quarter("1(B)"), 2)
        self.assertEqual(_group_quarter("2"), 3)
        self.assertEqual(_group_quarter("3"), 3)

    def test_표기가_달라도_같은_그룹(self):
        for g in ("1A", "1 (a)", "1-A", "1（A）"):
            self.assertEqual(_group_quarter(g), 1, g)
        for g in ("1B", "1 (b)", "1_b"):
            self.assertEqual(_group_quarter(g), 2, g)

    def test_모르는_그룹은_None(self):
        for g in ("", None, "4", "가"):
            self.assertIsNone(_group_quarter(g), g)


if __name__ == "__main__":
    unittest.main()
