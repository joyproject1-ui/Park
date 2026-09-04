# -*- coding: utf-8 -*-
"""옛 워드(.doc) 변환 경로 — 담당자 PC 에는 pywin32 도 LibreOffice 도 없다.

PowerShell 이 막혀 있어도 VBScript(cscript)로 Word 를 부를 수 있어야 하고, 둘 다 안 되면
무엇을 해야 하는지 말해 주는 오류가 나와야 한다. COM 은 이 환경에 없으므로 명령을 가로채
어떤 명령이 어떻게 불리는지 본다.
"""
from __future__ import unicode_literals

import io
import os
import unittest

from pqr.engine import convert


def setattr_pair(module, pair):
    module._word_via_powershell, module._word_via_vbscript = pair


class 스크립트파일로_넘긴다(unittest.TestCase):
    """명령줄에 통째로 넘기면 따옴표·한글 경로에서 잘리는 PC 가 있다."""

    def setUp(self):
        self.calls = []
        self.seen = []
        self._run = convert._run
        self._which = convert.shutil.which
        convert.last_error[:] = []

        def fake_run(cmd, why):
            self.calls.append((cmd, why))
            path = cmd[-1]
            self.seen.append(io.open(path, encoding="utf-8-sig").read())
            return self.ok

        convert._run = fake_run
        convert.shutil.which = lambda name: "/usr/bin/" + name
        self.ok = True

    def tearDown(self):
        convert._run = self._run
        convert.shutil.which = self._which

    def test_powershell_은_File_로_부른다(self):
        convert._powershell("$x = 1")
        cmd, why = self.calls[0]
        self.assertIn("-File", cmd)
        self.assertEqual(why, "ps1")
        self.assertEqual(self.seen[0], "$x = 1")

    def test_vbscript_는_cscript_로_부른다(self):
        convert._vbscript('WScript.Quit 0')
        cmd, why = self.calls[0]
        self.assertIn("//nologo", cmd)
        self.assertEqual(why, "vbs")

    def test_스크립트_임시파일은_지운다(self):
        convert._powershell("$x = 1")
        self.assertFalse(os.path.exists(self.calls[0][0][-1]))

    def test_한글_경로가_그대로_들어간다(self):
        convert._word_via_vbscript("/tmp/전년도 결재본.doc", "/tmp/나온 것.docx")
        self.assertIn("전년도 결재본.doc", self.seen[0])
        self.assertIn("나온 것.docx", self.seen[0])

    def _swap(self, ps, vbs):
        """두 경로를 바꿔 끼우고 원래 것을 되돌려 놓는다."""
        keep = (convert._word_via_powershell, convert._word_via_vbscript)
        convert._word_via_powershell, convert._word_via_vbscript = ps, vbs
        self.addCleanup(lambda: setattr_pair(convert, keep))

    def test_PowerShell_이_막히면_VBScript_로_넘어간다(self):
        order = []
        self._swap(lambda s, d, f=16: (order.append("ps"), False)[1],
                   lambda s, d, f=16: (order.append("vbs"), False)[1])
        self.assertFalse(convert._word_convert("a.doc", "b.docx"))
        self.assertEqual(order, ["ps", "vbs"])

    def test_VBScript_가_되면_거기서_끝(self):
        order = []
        self._swap(lambda s, d, f=16: (order.append("ps"), False)[1],
                   lambda s, d, f=16: (order.append("vbs"), True)[1])
        made = os.path.join(os.path.dirname(__file__), "..", "b.docx")
        io.open(made, "w").write("x")
        try:
            self.assertTrue(convert._word_convert("a.doc", made))
            self.assertEqual(order, ["ps", "vbs"])
        finally:
            os.remove(made)

    def test_한_방법이_터져도_다음_방법을_본다(self):
        def boom(s, d, f=16):
            raise RuntimeError("COM 막힘")
        self._swap(boom, lambda s, d, f=16: False)
        self.assertFalse(convert._word_convert("a.doc", "b.docx"))
        self.assertTrue(any("COM 막힘" in e for e in convert.last_error))


class 안내(unittest.TestCase):
    def test_둘_다_안_되면_무엇을_할지_말한다(self):
        soffice, with_word = convert._soffice, convert._with_word
        convert._soffice = lambda: None
        convert._with_word = lambda src, dst: False
        convert.last_error[:] = ["ps1: COM 을 만들지 못했습니다"]
        try:
            with self.assertRaises(convert.ConvertError) as caught:
                convert.to_docx("/tmp/전년도.doc", "/tmp/나온것.docx")
            message = str(caught.exception)
            self.assertIn("Word 문서(*.docx)", message)
            self.assertIn("COM 을 만들지 못했습니다", message)   # 마지막 오류도 함께
        finally:
            convert._soffice, convert._with_word = soffice, with_word

    def test_이미_docx_면_복사만(self):
        import tempfile
        src = os.path.join(tempfile.mkdtemp(), "a.docx")
        io.open(src, "w").write("x")
        dst = os.path.join(tempfile.mkdtemp(), "b.docx")
        self.assertEqual(convert.to_docx(src, dst), "copy")
        self.assertTrue(os.path.isfile(dst))


if __name__ == "__main__":
    unittest.main()


class pywin32가_터져도(unittest.TestCase):
    """담당자 PC 에서 pywin32·PowerShell·VBScript 가 다 되는데도 보고서가 안 나온 일이 있었다.

    pywin32 길만 쓰고 그 오류를 그대로 밖으로 던져, 진단에서 '됩니다' 로 나온 나머지 두 길을
    아예 시도하지 않았다.
    """

    def setUp(self):
        self.tried = []
        self._convert = convert._word_convert
        convert.last_error[:] = []

        def fake_convert(src, dst, fmt=16):
            self.tried.append((src, dst, fmt))
            io.open(dst, "w").write("made")
            return True
        convert._word_convert = fake_convert

        import sys
        self.fake = type(sys)("win32com.client")
        self.fake.DispatchEx = self.dispatch
        package = type(sys)("win32com")
        package.client = self.fake              # import win32com.client 뒤 win32com.client 로 닿는다
        sys.modules["win32com"] = package
        sys.modules["win32com.client"] = self.fake

    def tearDown(self):
        import sys
        convert._word_convert = self._convert
        sys.modules.pop("win32com.client", None)
        sys.modules.pop("win32com", None)

    def dispatch(self, name):
        raise OSError("COM 을 열지 못함")

    def test_스크립트_길로_넘어간다(self):
        import tempfile
        work = tempfile.mkdtemp(prefix="pqr-conv-test-")
        src, dst = os.path.join(work, "a.doc"), os.path.join(work, "a.docx")
        io.open(src, "w").write("x")
        self.assertTrue(convert._with_word(src, dst))
        self.assertEqual(len(self.tried), 1)                      # 다음 길을 실제로 시도했다
        self.assertTrue(any("pywin32" in e for e in convert.last_error))

    def test_문서를_못_열어도_스크립트_길로_넘어간다(self):
        import tempfile

        class Word(object):
            Visible = True
            DisplayAlerts = None

            class Documents(object):
                @staticmethod
                def Open(*args):
                    raise OSError("파일 변환 대화상자")
            def Quit(self):
                pass
        self.fake.DispatchEx = lambda name: Word()
        work = tempfile.mkdtemp(prefix="pqr-conv-test-")
        src, dst = os.path.join(work, "b.doc"), os.path.join(work, "b.docx")
        io.open(src, "w").write("x")
        self.assertTrue(convert._with_word(src, dst))
        self.assertEqual(len(self.tried), 1)

    def test_옛_워드를_열_때_변환_확인_대화상자를_끈다(self):
        """화면 없이 도는 자동화에서 '파일 변환' 대화상자가 뜨면 그대로 멈춘다."""
        import tempfile
        opened = {}

        class Doc(object):
            @staticmethod
            def SaveAs2(path, FileFormat=None):
                io.open(path, "w").write("made")
            @staticmethod
            def Close(flag):
                pass

        class Word(object):
            Visible = True
            DisplayAlerts = None

            class Documents(object):
                @staticmethod
                def Open(*args):
                    opened["args"] = args
                    return Doc()
            def Quit(self):
                pass
        word = Word()
        self.fake.DispatchEx = lambda name: word
        work = tempfile.mkdtemp(prefix="pqr-conv-test-")
        src, dst = os.path.join(work, "c.doc"), os.path.join(work, "c.docx")
        io.open(src, "w").write("x")
        self.assertTrue(convert._with_word(src, dst))
        # (파일, ConfirmConversions, ReadOnly, AddToRecentFiles)
        self.assertEqual(opened["args"][1:], (False, True, False))
        self.assertEqual(word.DisplayAlerts, 0)                   # wdAlertsNone
        self.assertEqual(self.tried, [])                          # 됐으니 다음 길은 안 쓴다


class 변환_실패는_언제나_ConvertError(unittest.TestCase):
    """COM 오류가 그대로 새어 나가면 부르는 쪽이 '변환 실패' 로 알아보지 못한다 —
    그러면 EDMS 빈 서식으로 물러나는 길까지 막힌다."""

    def setUp(self):
        self.pairs = (convert._with_word, convert._with_soffice)
        convert.last_error[:] = []

    def tearDown(self):
        convert._with_word, convert._with_soffice = self.pairs

    def test_COM_오류를_감싼다(self):
        import tempfile

        def boom(src, dst):
            raise OSError("Call was rejected by callee")
        convert._with_word = boom
        convert._with_soffice = boom
        work = tempfile.mkdtemp(prefix="pqr-conv-test-")
        src = os.path.join(work, "d.doc")
        io.open(src, "w").write("x")
        with self.assertRaises(convert.ConvertError) as caught:
            convert.to_docx(src, os.path.join(work, "d.docx"))
        self.assertIn("rejected by callee", str(caught.exception))    # 까닭을 그대로 보여 준다
        self.assertIn("Word 문서(*.docx)", str(caught.exception))     # 무엇을 하면 되는지도
