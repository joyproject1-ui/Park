# -*- coding: utf-8 -*-
"""3항 '대상 제품' 을 식약처 허가정보와 대조한다 (담당자 2026-09-12: "자동으로 해").

회사 밖으로 나가지 않는다 — 응답을 흉내 낸 글로만 시험한다.
"""
from __future__ import unicode_literals

import json
import os
import tempfile
import unittest

from pqr.engine.readers import mfds


올로원스 = {
    "ITEM_NAME": "올로원스점안액(올로파타딘염산염)",
    "ENTP_NAME": "한림제약(주)",
    "ITEM_PERMIT_DATE": "20111130",
    "ITEM_SEQ": "201110503",
    "ETC_OTC_CODE": "전문의약품",
    "STORAGE_METHOD": "기밀용기, 2~25℃보관",
    "VALID_TERM": "제조일로부터 24 개월",
    "REEXAM_TARGET": "",
    "RMP_TARGET": "",
    "PACK_UNIT": "[다회용]3mL/병",
}

표3 = {"제품명": "올로원스점안액(올로파타딘염산염)", "허가일자": "2011년 11월 30일",
       "보관조건": "기밀용기, 2~25℃보관", "사용기한": "제조일로부터 24개월",
       "제품분류": "전문의약품"}


def _응답(items):
    return json.dumps({"body": {"items": [{"item": one} for one in items]}}).encode("utf-8")


class 받아오기(unittest.TestCase):
    def test_응답에서_품목을_꺼낸다(self):
        got = mfds.fetch("올로원스점안액", "열쇠",
                         opener=lambda url, timeout: _응답([올로원스]))
        self.assertEqual(len(got), 1)
        self.assertEqual(got[0]["허가일자"], "20111130")
        self.assertEqual(got[0]["업체명"], "한림제약(주)")

    def test_열쇠가_없으면_부르지_않는다(self):
        def 부르면안됨(url, timeout):
            raise AssertionError("열쇠 없이 불렀습니다")
        self.assertEqual(mfds.fetch("올로원스점안액", "", opener=부르면안됨), [])

    def test_이름이_가장_잘_맞는_것을_고른다(self):
        일회용 = dict(올로원스, ITEM_NAME="올로원스점안액(올로파타딘염산염)(1회용)")
        rows = mfds.fetch("x", "열쇠", opener=lambda u, t: _응답([일회용, 올로원스]))
        got = mfds.pick(rows, "올로원스점안액(올로파타딘염산염)")
        self.assertEqual(got["제품명"], "올로원스점안액(올로파타딘염산염)")


class 대조(unittest.TestCase):
    def _info(self, **바꿀것):
        rows = mfds.fetch("x", "열쇠", opener=lambda u, t: _응답([dict(올로원스, **바꿀것)]))
        return rows[0]

    def test_모두_같으면_어긋남_없음(self):
        self.assertEqual(mfds.compare(표3, self._info()), [])

    def test_허가일자가_다르면_올린다(self):
        got = mfds.compare(표3, self._info(ITEM_PERMIT_DATE="20111129"))
        self.assertEqual([one[0] for one in got], ["허가일자"])

    def test_보관조건_글이_다르면_올린다(self):
        got = mfds.compare(표3, self._info(STORAGE_METHOD="차광기밀용기, 실온보관"))
        self.assertEqual([one[0] for one in got], ["보관조건"])

    def test_허가정보_칸이_비면_견주지_않는다(self):
        """서비스가 그 항목을 주지 않는 것과 '해당 없음' 은 다르다."""
        self.assertEqual(mfds.compare(표3, self._info(STORAGE_METHOD="")), [])

    def test_취소_이력은_크게_알린다(self):
        got = mfds.compare(표3, self._info(CANCEL_DATE="20250401", CANCEL_NAME="자진취하"))
        self.assertIn("허가 상태", [one[0] for one in got])

    def test_재심사_RMP_가_비면_해당_없음으로_적는다(self):
        self.assertIn("재심사대상: 해당 없음", mfds.notes(self._info()))
        self.assertIn("RMP대상: 대상", mfds.notes(self._info(RMP_TARGET="대상")))


class 열쇠찾기(unittest.TestCase):
    def test_제품_폴더의_파일에서_읽는다(self):
        root = tempfile.mkdtemp()
        folder = os.path.join(root, "QC1-5059 올로원스점안액")
        os.makedirs(os.path.join(folder, "공통"))
        with open(os.path.join(folder, "공통", mfds.KEY_FILE), "w", encoding="utf-8") as handle:
            handle.write("  열쇠값  \n")
        old = os.environ.pop("MFDS_API_KEY", None)
        try:
            self.assertEqual(mfds.api_key(folder), "열쇠값")
        finally:
            if old is not None:
                os.environ["MFDS_API_KEY"] = old

    def test_없으면_None(self):
        old = os.environ.pop("MFDS_API_KEY", None)
        try:
            self.assertIsNone(mfds.api_key(tempfile.mkdtemp()))
        finally:
            if old is not None:
                os.environ["MFDS_API_KEY"] = old


class 노랑표시(unittest.TestCase):
    """어긋난 칸만 노랑으로 칠하고 문의에 올린다 (담당자 2026-09-12)."""

    def _표(self):
        import docx
        from pqr.engine import docedit as E
        doc = docx.Document()
        t = doc.add_table(rows=4, cols=4)
        for i, (이름, 값) in enumerate((("제품명", "올로원스점안액(올로파타딘염산염)"),
                                        ("허가일자", "2011년 11월 30일"),
                                        ("보관조건", "기밀용기, 2~25℃보관"))):
            cells = E.raw_cells(t.rows[i + 1])
            E.set_cell(cells[1], 이름)
            E.set_cell(cells[2], 값)
        return doc, t

    def _돌린다(self, **바꿀것):
        from pqr.engine import recipe_ointment as R, docedit as E
        doc, t = self._표()
        표3, 칸3 = {}, {}
        for row in t.rows[1:]:
            cells = E.raw_cells(row)
            이름 = E.cell_text(cells[1]).strip()
            if 이름:
                표3[이름] = E.cell_text(cells[2]).strip()
                칸3[이름] = cells[2]
        old_key, old_fetch = mfds.api_key, mfds.fetch
        mfds.api_key = lambda folder=None: "열쇠"
        mfds.fetch = lambda name, key, **kw: [dict(
            {k: mfds._value(dict(올로원스, **바꿀것), k) for k in mfds.FIELDS})]
        issues = []
        try:
            n = R._check_license(표3, "올로원스점안액(올로파타딘염산염)", "", issues, lambda *a: None, 칸3)
        finally:
            mfds.api_key, mfds.fetch = old_key, old_fetch
        칠한칸 = {이름 for 이름, 칸 in 칸3.items() if "highlight" in 칸._tc.xml}
        return n, issues, 칠한칸

    def test_어긋난_칸만_칠한다(self):
        n, issues, 칠한칸 = self._돌린다(ITEM_PERMIT_DATE="20111129")
        self.assertEqual(n, 1)
        self.assertEqual(칠한칸, {"허가일자"})
        self.assertEqual(len(issues), 1)
        self.assertIn("허가일자", issues[0][1])

    def test_모두_같으면_아무것도_칠하지_않는다(self):
        n, issues, 칠한칸 = self._돌린다()
        self.assertEqual((n, issues, 칠한칸), (0, [], set()))


if __name__ == "__main__":
    unittest.main()
