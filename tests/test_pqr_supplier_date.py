# -*- coding: utf-8 -*-
"""8.1.1 완료일 = 공급업체 목록의 최종 평가 승인일 (담당자 2026-09-28: RSP107 완료일이 공란 →
"다음부터는 최종 평가 승인일을 기재해주도록 해줘")."""
import os
import unittest

os.environ["PQR_REVIEW"] = "0"

from pqr.engine import recipe_ointment as R      # noqa: E402


class 최종_평가_승인일(unittest.TestCase):
    def test_평가승인일(self):
        self.assertEqual(R.approval_date({"평가승인일": "2024. 10. 10"}), "2024.10.10")

    def test_열_이름에_빈칸이나_줄바꿈이_섞여도(self):
        self.assertEqual(R.approval_date({"평가 승인일": "2024-10-10"}), "2024.10.10")
        self.assertEqual(R.approval_date({"평가승인일 Approval date": "2024.10.10"}), "2024.10.10")

    def test_최초승인일은_마지막_후보(self):
        rec = {"최초승인일": "2016.05.16", "평가승인일": "2024.10.10"}
        self.assertEqual(R.approval_date(rec), "2024.10.10")
        self.assertEqual(R.approval_date({"최초승인일": "2016.05.16"}), "2016.05.16")

    def test_여럿이면_가장_늦은_날(self):
        rec = {"평가승인일": "2022.03.01", "재평가 승인일": "2024.10.10"}
        self.assertEqual(R.approval_date(rec), "2024.10.10")

    def test_날짜가_아니면_빈_글(self):
        self.assertEqual(R.approval_date({"평가승인일": "예정"}), "")
        self.assertEqual(R.approval_date({"평가승인일": None, "원료코드": "RSP107"}), "")
        self.assertEqual(R.approval_date({"개정번호": "Rev.5"}), "")


if __name__ == "__main__":
    unittest.main()
