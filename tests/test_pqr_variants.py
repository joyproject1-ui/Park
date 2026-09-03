# -*- coding: utf-8 -*-
"""연간 계획서 비고로 한 품목을 여러 PQR 건으로 나누는 규칙."""
from __future__ import unicode_literals

import unittest

from pqr.build import expand_product_variants


def meta(code, name, note, **extra):
    row = {"product_code": code, "product_name": name, "note": note,
           "form": "점안제", "group": "1(B)", "lots": 10}
    row.update(extra)
    return {code: row}


class VariantTest(unittest.TestCase):
    def test_비고가_없으면_그대로(self):
        rows = meta("QC1-5004", "칼테점안액", "")
        self.assertEqual(expand_product_variants(rows), rows)

    def test_미국_수출용_전용은_제품명에만_표기한다(self):
        got = expand_product_variants(meta("QC1-5031", "훼미리케어오리지날점안액", "미국 수출용"))
        self.assertEqual(list(got), ["QC1-5031"])
        self.assertEqual(got["QC1-5031"]["product_name"], "훼미리케어오리지날점안액 (미국 수출용)")

    def test_이미_표기된_이름은_두_번_붙이지_않는다(self):
        got = expand_product_variants(
            meta("QC1-5031", "훼미리케어오리지날점안액 (미국 수출용)", "미국 수출용"))
        self.assertEqual(got["QC1-5031"]["product_name"], "훼미리케어오리지날점안액 (미국 수출용)")

    def test_1회용_다회용은_기본이_다회용이고_1회용을_더한다(self):
        got = expand_product_variants(meta("QC1-5001", "후메론점안액", "1회용 / 다회용"))
        self.assertEqual(sorted(got), ["QC1-5001", "QC1-5001-1회용"])
        self.assertEqual(got["QC1-5001"]["product_name"], "후메론점안액")     # 기본 = 다회용
        self.assertEqual(got["QC1-5001-1회용"]["product_name"], "후메론점안액 (1회용)")

    def test_일회용_이라고_적혀도_같게_본다(self):
        got = expand_product_variants(meta("QC1-5049", "데일리큐점안액", "미국 수출용 / 일회용 / 다회용"))
        self.assertEqual(sorted(got),
                         ["QC1-5049", "QC1-5049-1회용", "QC1-5049-미국수출용"])

    def test_미국_수출용이_다른_구분과_같이_있으면_따로_만든다(self):
        got = expand_product_variants(meta("QC1-5033", "이지드롭점안액", "미국 수출용 / 내수용 / 수출용"))
        self.assertEqual(sorted(got), ["QC1-5033", "QC1-5033-미국수출용"])
        self.assertEqual(got["QC1-5033"]["product_name"], "이지드롭점안액")

    def test_내수용_수출용만_있으면_나누지_않는다(self):
        rows = meta("QC1-7007", "퀴노비드안연고", "내수용/수출용")
        self.assertEqual(expand_product_variants(rows), rows)

    def test_갈라낸_건은_제형과_그룹을_물려받는다(self):
        got = expand_product_variants(
            meta("QC1-5002", "누마렌점안액", "1회용 / 다회용 / 미국 수출용", owner="강민태"))
        for code in ("QC1-5002-1회용", "QC1-5002-미국수출용"):
            self.assertEqual(got[code]["form"], "점안제")
            self.assertEqual(got[code]["group"], "1(B)")
            self.assertEqual(got[code]["owner"], "강민태")
            self.assertEqual(got[code]["product_code"], code)

    def test_마스터에_이미_있는_코드는_손대지_않는다(self):
        rows = meta("QC1-5001", "후메론점안액", "1회용 / 다회용")
        rows["QC1-5001-1회용"] = {"product_code": "QC1-5001-1회용",
                                 "product_name": "후메론점안액(1회용) 담당자 기입",
                                 "note": "", "form": "점안제", "group": "2", "lots": 3}
        got = expand_product_variants(rows)
        self.assertEqual(got["QC1-5001-1회용"]["product_name"], "후메론점안액(1회용) 담당자 기입")
        self.assertEqual(got["QC1-5001-1회용"]["group"], "2")
