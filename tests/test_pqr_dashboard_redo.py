# -*- coding: utf-8 -*-
"""대시보드 '보고서 완료' 단추 — 담당자 2026-09: "누르면 재작성 버튼이 뜨게 해서 재작성 or 취소".

완성본을 바로 열지 않고 작은 창에서 보고서 재작성·완성본 열기·취소를 고르게 합니다."""
from __future__ import unicode_literals

import io
import os
import re
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
INDEX = os.path.join(HERE, "..", "docs", "pqr", "index.html")


class 보고서완료_재작성창(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with io.open(INDEX, encoding="utf-8") as handle:
            cls.html = handle.read()

    def test_재작성_창이_있다(self):
        self.assertIn('id="redo-modal"', self.html)
        for button in ("rd-submit", "rd-cancel", "rd-open", "rd-close"):
            self.assertIn('id="%s"' % button, self.html)
        self.assertIn("보고서 재작성", self.html)

    def test_완료_단추는_재작성_창을_띄운다(self):
        """예전에는 누르자마자 완성본이 열렸다 — 이제 창에서 고른다."""
        handler = re.search(r'closest\("\.final-btn"\);\s*if \(final && !final\.disabled\)\{ (\w+)\(',
                            self.html)
        self.assertIsNotNone(handler)
        self.assertEqual(handler.group(1), "openRedo")

    def test_재작성은_보고서_작성과_같은_길을_쓴다(self):
        """재작성은 '보고서 작성' 과 같은 makeReport → api/report 로 같은 이름에 덮어씁니다."""
        submit = re.search(r'\$\("#rd-submit"\)\.addEventListener\("click", \(\)=>\{(.*?)\}\);',
                           self.html, re.S)
        self.assertIsNotNone(submit)
        self.assertIn("makeReport(code, button)", submit.group(1))
        self.assertIn("closeRedo()", submit.group(1))

    def test_esc_와_바탕_클릭으로_닫힌다(self):
        self.assertIn('if (redoModal.classList.contains("open")) closeRedo();', self.html)
        self.assertRegex(self.html, r'scrim\.addEventListener\("click", \(\)=>\{[^\n]*closeRedo\(\)')


if __name__ == "__main__":
    unittest.main()
