# -*- coding: utf-8 -*-
"""이 PC 진단 — 무엇이 막혔는지 담당자가 두 번 눌러 알려 줄 수 있어야 한다."""
from __future__ import unicode_literals

import io
import os
import unittest

from pqr import doctor


class 보고서(unittest.TestCase):
    def test_사람이_읽을_수_있는_줄로_나온다(self):
        lines = doctor.report()
        text = "\n".join(lines)
        self.assertIn("PQR 프로그램 진단", text)
        self.assertIn("옛 워드(.doc) 결재본을 바꾸는 길", text)
        self.assertIn("첨부 Cpk 엑셀", text)
        self.assertIn("판정", text)
        self.assertIn("이 화면을 그대로 알려 주시면", text)

    def test_바꿀_길이_있으면_된다고_말한다(self):
        keep = doctor.shutil.which
        doctor.shutil.which = lambda name: "/usr/bin/soffice" if "office" in name else None
        try:
            text = "\n".join(doctor.report())
            self.assertIn("바꿀 수 있습니다", text)
        finally:
            doctor.shutil.which = keep

    def test_길이_없으면_무엇을_할지_말한다(self):
        from pqr.engine import convert
        keep = convert._soffice
        convert._soffice = lambda: None
        try:
            text = "\n".join(doctor.report())
            self.assertIn("바꿀 길이 없습니다", text)
            self.assertIn("Word 문서(*.docx)", text)
        finally:
            convert._soffice = keep

    def test_표시는_O_X_로_읽힌다(self):
        self.assertIn("[O]", doctor._line(True, "됨"))
        self.assertIn("[X]", doctor._line(False, "안 됨"))
        self.assertIn("[-]", doctor._line(None, "해당 없음"))


class 배치파일(unittest.TestCase):
    """cmd.exe 는 한글이 섞인 배치 파일을 읽다가 자기 텍스트를 명령으로 실행한다."""

    def path(self):
        return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "PQR-진단.bat")

    def test_ASCII_만_담는다(self):
        raw = io.open(self.path(), "rb").read()
        self.assertTrue(all(b < 128 for b in raw), "배치 파일에 한글이 들어가면 안 됩니다")

    def test_진단을_부른다(self):
        text = io.open(self.path(), encoding="ascii").read()
        self.assertIn("-m pqr doctor", text)
        self.assertIn("pause", text)


if __name__ == "__main__":
    unittest.main()


class 실제로_바꿔_본다(unittest.TestCase):
    """Word 를 띄웠다 닫는 것만으로는 모자랐다 — 담당자 PC 에서 그 검사는 모두 '됩니다' 였는데
    보고서는 안 나왔다(2026-09, 디겐타안연고)."""

    def setUp(self):
        self.platform = doctor.sys.platform
        self.make = doctor._make_doc

    def tearDown(self):
        doctor.sys.platform = self.platform
        doctor._make_doc = self.make

    def test_윈도우가_아니면_건너뛴다(self):
        doctor.sys.platform = "linux"
        ok, why = doctor.convert_check()
        self.assertIsNone(ok)

    def test_시험용_doc_를_못_만들면_그렇게_말한다(self):
        doctor.sys.platform = "win32"
        doctor._make_doc = lambda path: False
        ok, why = doctor.convert_check()
        self.assertFalse(ok)
        self.assertIn("시험용", why)

    def test_바꾸다_막히면_까닭을_그대로_보여_준다(self):
        from pqr.engine import convert
        doctor.sys.platform = "win32"
        doctor._make_doc = lambda path: io.open(path, "w").write("x") or True
        real = convert.to_docx

        def boom(src, dst):
            raise convert.ConvertError("파일 변환 대화상자에서 멈춤")
        convert.to_docx = boom
        try:
            ok, why = doctor.convert_check()
        finally:
            convert.to_docx = real
        self.assertFalse(ok)
        self.assertIn("파일 변환 대화상자", why)

    def test_되면_어느_길로_됐는지_말한다(self):
        from pqr.engine import convert
        doctor.sys.platform = "win32"
        doctor._make_doc = lambda path: io.open(path, "w").write("x") or True
        real = convert.to_docx
        convert.to_docx = lambda src, dst: "word"
        try:
            ok, why = doctor.convert_check()
        finally:
            convert.to_docx = real
        self.assertTrue(ok)
        self.assertIn("word", why)
