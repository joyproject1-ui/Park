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


def test_guess_assay_makes_a_plausible_value_from_messy_text():
    from pqr.engine.handwriting import guess_assay
    assert guess_assay("9a.4-1.", 90, 120) == 99.4
    assert guess_assay("qn.ay", 90, 110) == 97.9
    assert guess_assay("1013", 90, 110) == 101.3
    assert guess_assay("987", 90, 110) == 98.7
    assert guess_assay("lub", 90, 120) is None          # 숫자가 없다 — 지어내지 않는다
    assert guess_assay("", 90, 110) is None
    # 규격 밖 후보뿐이면 헷갈리는 숫자 하나를 바꿔 본다 — 손글씨 9 가 4·7 로 읽힌 것
    assert guess_assay("44.ny.", 90, 110) == 94.7
    assert guess_assay("79.9b", 90, 120) == 99.9
    assert guess_assay("12.3", 90, 110) is None         # 한 자리 바꿔도 규격 근처가 아니면 만들지 않는다


def test_refresh_cache_fills_only_unsure_cells_and_keeps_clean_ones(monkeypatch):
    from pqr.engine import handwriting as hw
    old = [{"lot": "OGX901", "points": [
        {"period": "3M", "done": "2025.03.17", "assays": {"A": 100.5}, "unsure": ["B"]},
        {"period": "6M", "done": "", "assays": {"A": 100.0, "B": 97.8}, "unsure": []}]}]
    new = [{"lot": "OGX901", "points": [
        {"period": "3M", "done": "2025.03.17", "assays": {"A": 100.4, "B": 90.8}, "unsure": ["B"]},
        {"period": "6M", "done": "2025.07.02", "assays": {"A": 99.0, "B": 96.0}, "unsure": ["A", "B"]}]}]
    monkeypatch.setattr(hw, "available", lambda: True)
    monkeypatch.setattr(hw, "read_folder", lambda paths, specs, log: new)
    changed = hw.refresh_cache(old, 0, ["/x/a.pdf"], None, None)
    p3, p6 = old[0]["points"]
    assert p3["assays"] == {"A": 100.5, "B": 90.8} and "B" in p3["unsure"]   # 애매한 B 만 예상값으로
    assert p6["assays"] == {"A": 100.0, "B": 97.8} and p6["done"] == "2025.07.02"  # 깨끗한 값은 그대로, 빈 완료일은 채움
    assert changed == 2
    assert hw.refresh_cache(old, hw.READER_VERSION, ["/x/a.pdf"], None, None) == 0    # 이미 새 판독기 결과
