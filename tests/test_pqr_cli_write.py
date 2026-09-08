# -*- coding: utf-8 -*-
"""`pqr write <제품 폴더>` — 대시보드 없이 보고서를 만드는 명령.

담당자 2026-09-08: "너가 직접 Cowork 가 하는 것처럼 대시보드 폴더에 접근해서 PQR 파일을
직접 읽고 작성해 줘." PC 의 Claude 가 제품 폴더에서 이 명령을 돌리고 문의 목록을 읽는다.
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr import cli                                                   # noqa: E402


class write_명령(unittest.TestCase):
    def test_파서에_있다(self):
        args = cli.build_parser().parse_args(["write", "입력/QC1-5087 아이퓨어점안액", "--today", "2026-09-08"])
        self.assertIs(args.func, cli.cmd_write)
        self.assertEqual(args.today, "2026-09-08")

    def test_없는_폴더면_2(self):
        args = cli.build_parser().parse_args(["write", "/no/such/folder"])
        self.assertEqual(cli.cmd_write(args), 2)


if __name__ == "__main__":
    unittest.main()
