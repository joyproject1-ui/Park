# -*- coding: utf-8 -*-
"""연간 계획서 제형별 표 읽기와 제품 마스터 반영."""
from __future__ import unicode_literals

import unittest

from pqr.plan import apply_to_master, entries, normalize_code, note_tokens


class CodeTest(unittest.TestCase):
    def test_사이_줄표와_공백을_정리한다(self):
        self.assertEqual(normalize_code("QC1 – 5002"), "QC1-5002")
        self.assertEqual(normalize_code("QC1-5002"), "QC1-5002")
        self.assertEqual(normalize_code(" QC1 - 7007 "), "QC1-7007")

    def test_코드가_아니면_빈_문자열(self):
        for text in ("", "누마렌점안액", "16", None):
            self.assertEqual(normalize_code(text), "")

    def test_비고를_구분으로_쪼갠다(self):
        self.assertEqual(note_tokens("1회용 / 다회용"), ["1회용", "다회용"])
        self.assertEqual(note_tokens("내수용/수출용"), ["내수용", "수출용"])
        self.assertEqual(note_tokens(""), [])


class EntriesTest(unittest.TestCase):
    def test_한_줄뿐이면_그대로(self):
        self.assertEqual(entries("QC1-7007", "퀴노비드안연고", [(19, "내수용/수출용")]),
                         [("QC1-7007", "퀴노비드안연고", 19, "내수용/수출용")])

    def test_미국_수출용_한_줄이면_이름에만_표기한다(self):
        got = entries("QC1-5031", "훼미리케어오리지날점안액", [(0, "미국 수출용")])
        self.assertEqual(got, [("QC1-5031", "훼미리케어오리지날점안액 (미국 수출용)", 0, "미국 수출용")])

    def test_1회용은_같은_Lot_을_나눠_쓰므로_품목_합계로_적는다(self):
        got = entries("QC1-5001", "후메론점안액", [(56, "1회용 / 다회용")])
        self.assertEqual([(g[0], g[2]) for g in got],
                         [("QC1-5001", 56), ("QC1-5001-1회용", 56)])
        self.assertIn("품목 합계", got[1][3])
        self.assertNotIn("품목 합계", got[0][3])

    def test_미국_수출용_줄은_그_줄의_Lot_을_쓴다(self):
        got = entries("QC1-5049", "데일리큐점안액",
                      [(3, "미국 수출용"), (1, "일회용 / 다회용")])
        self.assertEqual([(g[0], g[2]) for g in got],
                         [("QC1-5049", 1), ("QC1-5049-1회용", 1), ("QC1-5049-미국수출용", 3)])

    def test_내수용_수출용_줄이_본_품목이_된다(self):
        got = entries("QC1-5033", "이지드롭점안액",
                      [(0, "미국 수출용"), (1, "내수용 / 수출용")])
        self.assertEqual([(g[0], g[2]) for g in got],
                         [("QC1-5033", 1), ("QC1-5033-미국수출용", 0)])


class MasterTest(unittest.TestCase):
    def rows(self):
        return [{"제품코드": "QC1-5049", "제품명": "데일리큐점안액", "제형": "점안제",
                 "그룹": "1(B)", "생산수량": "4", "비고": "미국 수출용 / 일회용 / 다회용"},
                {"제품코드": "QC1-1022", "제품명": "메섹신정", "제형": "정제",
                 "그룹": "1(B)", "생산수량": "40", "비고": ""}]

    def test_계획서에_있는_품목만_바꾼다(self):
        got = apply_to_master({"QC1-5049": [(3, "미국 수출용"), (1, "일회용 / 다회용")]},
                              self.rows())
        codes = [r["제품코드"] for r in got]
        self.assertEqual(codes, ["QC1-5049", "QC1-5049-1회용", "QC1-5049-미국수출용",
                                 "QC1-1022"])
        self.assertEqual(got[0]["생산수량"], 1)          # 4 → 1 로 바로잡힘
        self.assertEqual(got[2]["생산수량"], 3)
        self.assertEqual(got[3], self.rows()[1])         # 계획서에 없는 품목은 그대로

    def test_갈라낸_행은_다른_열을_물려받는다(self):
        got = apply_to_master({"QC1-5049": [(3, "미국 수출용"), (1, "일회용 / 다회용")]},
                              self.rows())
        for row in got[:3]:
            self.assertEqual(row["제형"], "점안제")
            self.assertEqual(row["그룹"], "1(B)")

    def test_두_번_돌려도_행이_늘지_않는다(self):
        plan = {"QC1-5049": [(3, "미국 수출용"), (1, "일회용 / 다회용")]}
        once = apply_to_master(plan, self.rows())
        twice = apply_to_master(plan, once)
        self.assertEqual([r["제품코드"] for r in once], [r["제품코드"] for r in twice])
