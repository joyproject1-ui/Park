# -*- coding: utf-8 -*-
"""압축 안의 한글 파일 이름 — 알집(CP949)이든 리눅스(UTF-8)든 제대로 읽는다.

담당자 2026-09: "PQR 입력폴더에 ZIP파일로 업로드 되어있으니 안정성 시험 파일을 찾아서 입력해".
이름이 깨지면 '13. 안정성 시험' 같은 폴더 이름으로 항을 가르지 못한다.

파이썬 zipfile 은 UTF-8 표시(0x800)가 없는 압축의 이름을 CP437 로 읽는다. 그래서
한글 이름이 '╡≡░╒┼╕' 처럼 온다 — 되돌린 바이트를 다시 제대로 읽어야 한다.
"""
import os
import shutil
import tempfile
import unittest
import zipfile

from pqr.engine import collect


class _표시없는압축(object):
    """UTF-8 표시가 꺼진 압축 항목 — 이름은 CP437 로 읽힌 상태."""

    def __init__(self, 바이트):
        self.filename = 바이트.decode("cp437")
        self.flag_bits = 0


class 이름_되돌리기(unittest.TestCase):
    이름 = "13. 안정성 시험/[LT-23-1]디겐타안연고_OGW701_24M.pdf"

    def test_알집이_만든_CP949_이름(self):
        got = collect._member_name(_표시없는압축(self.이름.encode("cp949")))
        self.assertEqual(got, self.이름)

    def test_리눅스가_만든_UTF8_이름(self):
        got = collect._member_name(_표시없는압축(self.이름.encode("utf-8")))
        self.assertEqual(got, self.이름)

    def test_표시가_켜져_있으면_그대로_둔다(self):
        info = _표시없는압축(b"plain.pdf")
        info.flag_bits = 0x800
        info.filename = self.이름
        self.assertEqual(collect._member_name(info), self.이름)


class 압축_속_안정성_시험(unittest.TestCase):
    """압축째 올린 13항 자료를 항 번호로 갈라 찾는다."""

    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="pqr-zip-")
        with zipfile.ZipFile(os.path.join(self.root, "13. 안정성 시험.zip"), "w") as z:
            for lot in ("OGW701", "OGY301"):
                z.writestr("13. 안정성 시험/[LT-23-1]디겐타안연고_%s_24M.pdf" % lot, b"%PDF-1.4\n")

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    def test_압축_안의_시험일지를_13항으로_찾는다(self):
        got = collect.discover(self.root)
        self.assertEqual(sorted(os.path.basename(p) for p in got.get("13", [])),
                         ["[LT-23-1]디겐타안연고_OGW701_24M.pdf",
                          "[LT-23-1]디겐타안연고_OGY301_24M.pdf"])


if __name__ == "__main__":
    unittest.main()
