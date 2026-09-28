# -*- coding: utf-8 -*-
"""프린터 연결 대기 막기 — 담당자 PC 2026-09-28: 보고서 작성 22분, "프린터 연결을 기다리는 중…" 창.

Word·Excel 은 문서를 열 때 기본 프린터를 묻는다. 자동화로 띄운 Office 안에서만 조용한 프린터로 두고,
pywin32 길은 시간이 넘으면 끊는다. 이 PC 에는 Word 가 없으므로 흉내 개체로 시험한다.
"""
import os
import time
import unittest

os.environ["PQR_REVIEW"] = "0"

from pqr.engine import convert as C            # noqa: E402


class 흉내Word(object):
    def __init__(self, fail=()):
        self.calls = []; self.fail = set(fail)
        outer = self
        class Basic(object):
            def FilePrintSetup(self, *args, **kw):
                name = kw.get("Printer") or (args[0] if args else "")
                if name in outer.fail:
                    raise RuntimeError("no such printer")
                outer.calls.append((name, kw.get("DoNotSetAsSysDefault", args[1] if len(args) > 1 else None)))
        self.WordBasic = Basic()


class 흉내Excel(object):
    def __init__(self, accept):
        self.accept = accept; self.set = None
    @property
    def ActivePrinter(self): return self.set
    @ActivePrinter.setter
    def ActivePrinter(self, value):
        if value not in self.accept: raise RuntimeError("bad printer")
        self.set = value


class 조용한_프린터(unittest.TestCase):
    def test_Word_는_시스템_기본을_바꾸지_않고_PDF_프린터로(self):
        w = 흉내Word()
        self.assertEqual(C.quiet_printer_word(w), "Microsoft Print to PDF")
        self.assertEqual(w.calls[0], ("Microsoft Print to PDF", 1))

    def test_PDF_프린터가_없으면_XPS_로(self):
        w = 흉내Word(fail=("Microsoft Print to PDF",))
        self.assertEqual(C.quiet_printer_word(w), "Microsoft XPS Document Writer")

    def test_둘_다_없으면_그냥_넘어간다(self):
        w = 흉내Word(fail=C.QUIET_PRINTERS)
        self.assertEqual(C.quiet_printer_word(w), "")

    def test_Excel_은_포트가_붙은_이름이나_맨_이름(self):
        self.assertEqual(C.quiet_printer_excel(흉내Excel({"Microsoft Print to PDF"})), "Microsoft Print to PDF")
        self.assertEqual(C.quiet_printer_excel(흉내Excel(set())), "")

    def test_스크립트_길에도_들어간다(self):
        self.assertIn("FilePrintSetup", C.WORD_QUIET_VBS)
        self.assertIn("FilePrintSetup", C.WORD_QUIET_PS)
        self.assertIn("DoNotSetAsSysDefault", C.WORD_QUIET_VBS + "DoNotSetAsSysDefault") # 이름 대신 자리(1)로 넘긴다
        lines = C.excel_quiet_printer_line("$x.ActivePrinter = %s")
        self.assertTrue(any("Microsoft Print to PDF" in line for line in lines))


class 시간_제한(unittest.TestCase):
    def test_제때_끝나면_결과를_돌려준다(self):
        self.assertTrue(C.with_deadline(lambda: True, "x", seconds=2))
        self.assertFalse(C.with_deadline(lambda: False, "x", seconds=2))

    def test_넘으면_False_와_프린터_안내(self):
        del C.last_error[:]
        self.assertFalse(C.with_deadline(lambda: time.sleep(3), "pywin32 Word", seconds=0.3))
        self.assertIn("끝나지 않아", C.last_error[-1])
        self.assertIn("Microsoft Print to PDF", C.last_error[-1])

    def test_안에서_터진_오류는_그대로_올라온다(self):
        def boom(): raise ValueError("x")
        with self.assertRaises(ValueError):
            C.with_deadline(boom, "x", seconds=2)

    def test_tasklist_출력에서_PID_를_읽는다(self):
        text = '"WINWORD.EXE","1234","Console","1","150,000 K"\r\n"WINWORD.EXE","99","Console","1","1 K"\r\n'
        self.assertEqual(C._parse_tasklist(text), {1234, 99})
        self.assertEqual(C._parse_tasklist("정보: 실행 중인 작업이 없습니다."), set())


class 목차_계산_스크립트(unittest.TestCase):
    """예전에는 두 스크립트 길 모두 문자열 서식 오류로 터져 한 번도 돌지 못했다 (2026-09-28 발견)."""

    def test_스크립트를_만드는_데서_터지지_않는다(self):
        del C.last_error[:]
        # cscript·powershell 이 없는 PC 에서는 False 를 돌려주되 TypeError 는 나지 않아야 한다
        self.assertFalse(C._fields_via_vbscript("x.docx"))
        self.assertFalse(C._fields_via_powershell("x.docx"))
        self.assertFalse(any("TypeError" in e for e in C.last_error))


if __name__ == "__main__":
    unittest.main()
