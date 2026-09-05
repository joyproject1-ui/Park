# -*- coding: utf-8 -*-
"""손글씨 시험일지 오프라인 판독 — 글 해석 규칙 (담당자 2026-09: "최대한 판독, 애매한 것만 노랑")."""
from __future__ import unicode_literals

import unittest

from pqr.engine import handwriting as H


class 함량_글_해석(unittest.TestCase):
    def test_깨끗한_값(self):
        self.assertEqual(H.parse_assay("100.5%.", 90, 110), (100.5, True))
        self.assertEqual(H.parse_assay("97.7%", 90, 110), (97.7, True))
        self.assertEqual(H.parse_assay("99.0%", 90, 120), (99.0, True))

    def test_흐트러진_값은_값이_있어도_애매(self):
        self.assertEqual(H.parse_assay("96-57.", 90, 110), (96.5, False))
        self.assertEqual(H.parse_assay("[02,1%", 90, 110), (102.1, False))
        self.assertEqual(H.parse_assay("qe.6%", 90, 110), (None, False))       # 글자가 섞이면 값도 두지 않는다

    def test_못_읽으면_없음(self):
        self.assertEqual(H.parse_assay("", 90, 110), (None, False))
        self.assertEqual(H.parse_assay("O7H,", 90, 110), (None, False))

    def test_규격에서_한참_벗어나면_없음(self):
        self.assertEqual(H.parse_assay("46.6%", 90, 120), (None, False))     # 98.6 을 잘못 읽은 것


class 날짜_글_해석(unittest.TestCase):
    def test_깨끗한_날짜(self):
        self.assertEqual(H.parse_date("2024.11.14", 2024, 2030), ("2024.11.14", True))
        self.assertEqual(H.parse_date("2025. 06. 30", 2024, 2030), ("2025.06.30", True))

    def test_흔한_오독을_연도_범위로_되살린다(self):
        self.assertEqual(H.parse_date("25250325", 2024, 2030), ("2025.03.25", False))
        self.assertEqual(H.parse_date("1225 11.18", 2024, 2030), ("2025.11.18", False))
        self.assertEqual(H.parse_date("202b ob. 29", 2024, 2030), ("2026.06.29", False))

    def test_말이_안_되면_없음(self):
        self.assertEqual(H.parse_date("D255202510", 2024, 2030), (None, False))
        self.assertEqual(H.parse_date("2025.13.40", 2024, 2030), (None, False))


if __name__ == "__main__":
    unittest.main()
