# -*- coding: utf-8 -*-
"""9.2 표의 Cpk — 16번 경향분석 Sheet 가 있는 항목(pH·비중·제제균일성 …)은 그 Sheet 의 규격으로 센다.
나조린 2025 결재본은 조제 pH·비중, 충전 pH, 포장 pH·함량·제제균일성에 Cpk 를 적었는데 2026-09-10
작성본은 함량만 적고 나머지를 사선으로 남겼다."""
from __future__ import unicode_literals

import unittest

from pqr.engine import recipe_ointment as R
from pqr.engine import conclusion


PH = {"stage": "충전", "item": "pH", "part": "", "sided": "Bilateral", "cells": {"N6": 6.2, "P6": 7.5}}
UNI = {"stage": "포장", "item": "제제균일성", "part": "말레인산페니라민", "sided": "Unilateral", "cells": {"P6": 15.0}}
JOBS = [PH, UNI, {"stage": "조제", "item": "pH", "part": "", "sided": "Bilateral", "cells": {"N6": 6.2, "P6": 7.0}},
        {"stage": "포장", "item": "제제균일성", "part": "나파졸린염산염", "sided": "Unilateral", "cells": {"P6": 15.0}}]


class Sheet항목(unittest.TestCase):
    def test_공정과_머리글로_Sheet_를_고른다(self):
        self.assertIs(R.sheet_job_for(JOBS, "충전", "pH"), PH)
        self.assertIs(R.sheet_job_for(JOBS, "조제", "pH"), JOBS[2])
        self.assertIsNone(R.sheet_job_for(JOBS, "포장", "질량용량(mL) 평균"))

    def test_성분이_붙은_열은_같은_성분의_Sheet(self):
        self.assertIs(R.sheet_job_for(JOBS, "포장", "제제균일성(질량편차) 말레인산페니라민"), UNI)
        self.assertIs(R.sheet_job_for(JOBS, "포장", "제제균일성(질량편차) 나파졸린염산염"), JOBS[3])

    def test_각주_번호가_붙은_머리글(self):
        job = {"stage": "조제", "item": "비중", "part": "", "sided": "Bilateral", "cells": {"N6": 1.0, "P6": 1.02}}
        self.assertIs(R.sheet_job_for([job], "조제", "비중1)"), job)

    def test_양쪽_규격은_전년도_Sheet_값과_같다(self):
        # 나조린 2025 충전 pH: HLF-QC-126-09 Sheet 가 Cpk 4.68 (표본 표준편차)
        vals = [6.45, 6.45, 6.43, 6.44, 6.41, 6.46, 6.45, 6.43, 6.44, 6.41, 6.46, 6.43]
        self.assertAlmostEqual(R.cpk_from_sheet(vals, PH), 4.68, places=2)

    def test_한쪽_규격은_상한만(self):
        # 나조린 2025 포장 제제균일성(말레인산페니라민): 결재본 3.25
        vals = [2.4, 3.3, 3.8, 3.6, 4.2, 2.5, 3.3, 3.9, 2.6, 3.5, 3.9, 3.3]
        got = R.cpk_from_sheet(vals, UNI)
        self.assertIsNotNone(got)
        self.assertGreater(got, 3.0)

    def test_값이_모두_같으면_없음(self):
        self.assertIsNone(R.cpk_from_sheet([0, 0, 0, 0], UNI))       # 바이오버든 0·0·0 → 사선

    def test_규격을_못_읽으면_없음(self):
        self.assertIsNone(R.cpk_from_sheet([1, 2, 3], {"stage": "포장", "item": "pH", "sided": "Bilateral", "cells": {}}))

    def test_결론은_Sheet_항목도_1_미만이면_올린다(self):
        low = conclusion.low_cpk_items({"assay": 1.5, "item/조제 pH": 0.8, "item/포장 pH": 2.0})
        self.assertEqual(low, [("조제 pH", 0.8)])


if __name__ == "__main__":
    unittest.main()
