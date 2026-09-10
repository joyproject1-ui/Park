# -*- coding: utf-8 -*-
"""2026-09-10 올로원스점안액(QC1-5059) 작성 실패·검토에서 굳힌 규칙들."""
from __future__ import unicode_literals

import os
import unittest

from pqr.engine import recipe_ointment as R
from pqr.engine import claude_cli
from pqr.engine import excel_attach as XA
from pqr.engine import stability_read
from pqr.engine.readers import batch_record


class 성적서_같은이름줄(unittest.TestCase):
    """성적서에 '유연물질' 줄이 둘(유연물질 1·2 표)이면 뒤 줄을 버리지 않고 '유연물질 2)' 로 둔다."""

    def test_둘째_줄은_2로(self):
        got = claude_cli._coa_items({"items": [
            {"name": "확인시험", "spec": "RT", "value": "동일"},
            {"name": "확인시험", "spec": "UV", "value": "동일함"},
            {"name": "유연물질", "spec": "A", "value": "Olopatadine E-isomer : 불검출"},
            {"name": "유연물질", "spec": "C", "value": "Olopatadine related Compound C : 불검출"},
        ]})
        self.assertEqual(list(got), ["확인시험", "확인시험 1)", "확인시험 2)", "유연물질", "유연물질 1)", "유연물질 2)"])
        self.assertEqual(got["유연물질 2)"]["value"], "Olopatadine related Compound C : 불검출")
        self.assertEqual(got["확인시험 2)"]["spec"], "UV")


class 유연물질_묶음(unittest.TestCase):
    rec = {"items": {
        "유연물질": {"value": "Olopatadine E-isomer : 불검출, 기타 유연물질: 불검출"},
        "유연물질 1)": {"value": "Olopatadine E-isomer : 불검출, 기타 유연물질: 불검출"},
        "유연물질 2)": {"value": "Olopatadine related Compound C : 불검출, 기타유연물질: 0.0%이하, 총 유연물질: 보고농도수준(0.1%)미만"},
    }}

    def test_같은_묶음_줄을_먼저_본다(self):
        self.assertEqual(R._impurity_value(self.rec, "기타 유연물질", "2"), "0.0%이하")
        self.assertEqual(R._impurity_value(self.rec, "기타 유연물질", "1"), "불검출")
        self.assertEqual(R._impurity_value(self.rec, "기타 유연물질"), "불검출")

    def test_없는_성분은_None(self):
        self.assertIsNone(R._impurity_value(self.rec, "Olopatadine Carbaldehyde", "1"))
        self.assertEqual(R._impurity_value(self.rec, "총 유연물질", "2"), "보고농도수준(0.1%)미만")


class 경향표_서식_고르기(unittest.TestCase):
    def test_시트가_모자라면_뒤_항목이_빠진다(self):
        picks = XA.pick_form_sheets([("pH", "pH"), ("삼투압", "삼투압"), ("함량", "함량"), ("보존제", "보존제")],
                                    ["pH", "함량", "보존제"])
        self.assertEqual(len(picks), 3)            # 그래서 _form_for_parts 가 다른 서식을 고른다

    def test_시트가_더_많은_서식으로(self):
        names = {"/서식3.xlsx": ["pH", "함량", "보존제"], "/서식7.xlsx": ["pH", "삼투압", "함량", "A", "B", "기타", "총"]}
        orig_names, orig_cands = XA.stability_xlsx.sheet_names, XA._stability_form_candidates
        XA.stability_xlsx.sheet_names = lambda p: names[p]
        XA._stability_form_candidates = lambda *a, **k: list(names)
        try:
            class D:
                folder = ""
            parts = [("pH", "pH", 5, 8, []), ("삼투압", "삼투압", 260, 340, []),
                     ("함량", "함량", 90, 110, []), ("보존제", "보존제", 80, 120, [])]
            self.assertEqual(XA._form_for_parts("/서식3.xlsx", parts, "", D()), "/서식7.xlsx")
            # 세 항목뿐이면 제품 폴더 서식 그대로
            self.assertEqual(XA._form_for_parts("/서식3.xlsx", parts[:1] + parts[2:], "", D()), "/서식3.xlsx")
        finally:
            XA.stability_xlsx.sheet_names, XA._stability_form_candidates = orig_names, orig_cands


class 판독_covers_all(unittest.TestCase):
    def test_읽지_못한_장이_있으면_covers_all_아님(self):
        paths = [os.path.join("x", "GVV101.pdf"), os.path.join("x", "GVX501.pdf")]
        self.assertEqual(stability_read.unread_paths(paths, [{"source": "GVV101.pdf"}]), [paths[1]])
        self.assertEqual(stability_read.unread_paths(paths, [{"source": "GVV101.pdf"}, {"source": "GVX501.pdf"}]), [])


class 필터는_자재가_아니다(unittest.TestCase):
    """담당자 2026-09-10 올로원스 8.1.3: "필터정보는 추가하지 말아줘"."""

    def test_필터_카트리지는_뺀다(self):
        self.assertFalse(batch_record.is_material("ESF126", "Polyethersulfone 0.2㎛ (20inch) (SCS92SP72S)"))
        self.assertFalse(batch_record.is_material("EPX001", "멸균 필터 0.22um"))
        self.assertFalse(batch_record.is_material("EPX002", "PVDF membrane cartridge"))

    def test_부원료_포장자재는_그대로(self):
        self.assertTrue(batch_record.is_material("EPB111", "벤잘코늄염화물"))
        self.assertTrue(batch_record.is_material("P12078", "PE병"))
        self.assertTrue(batch_record.is_material("EPS118", "염화나트륨"))


class 적격성_평가기간(unittest.TestCase):
    """10항 IQ·OQ·PQ 는 평가 기간 끝까지 완료된 문서 — 다음 해 재적격성평가는 이번 PQR 것이 아니다."""

    def test_기간_안의_것만(self):
        got = [("PQ25-5-DAE5012-R", "2025.04.28"), ("PQ26-5-DAE5012-R", "2026.04.27")]
        self.assertEqual(R._within(got, "20251231"), [("PQ25-5-DAE5012-R", "2025.04.28")])
        self.assertEqual(R._within(got, None), got)

    def test_기간_안에_하나도_없으면_그대로(self):
        got = [("PQ26-5-DAE5012-R", "2026.04.27")]
        self.assertEqual(R._within(got, "20251231"), got)

    def test_IOQ_짝도_기간을_본다(self):
        entry = {"IQ": [("IOQ23-TS-X-R", "2023.01.01"), ("IOQ26-TS-X-R", "2026.02.01")],
                 "OQ": [("IOQ23-TS-X-R", "2023.01.01"), ("IOQ26-TS-X-R", "2026.02.01")]}
        self.assertEqual(R._latest_pair(entry, cutoff="20251231"), ("IOQ23-TS-X-R", "2023.01.01"))
        self.assertEqual(R._latest_pair(entry), ("IOQ26-TS-X-R", "2026.02.01"))


if __name__ == "__main__":
    unittest.main()
