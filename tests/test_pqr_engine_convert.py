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
