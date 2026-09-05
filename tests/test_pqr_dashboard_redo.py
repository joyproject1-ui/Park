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


class 옛_작성본_경고(unittest.TestCase):
    """담당자 2026-09: 업데이트 뒤 '완성본 열기' 를 눌렀는데 옛 프로그램이 만든 파일이 열려
    "지시한 것이 하나도 반영 안 됐다". 작성본이 프로그램보다 오래되면 '재작성 필요' 로 보인다."""
    @classmethod
    def setUpClass(cls):
        with io.open(INDEX, encoding="utf-8") as handle:
            cls.html = handle.read()

    def test_작성_시각과_프로그램_버전을_견준다(self):
        self.assertIn("function staleReport(d)", self.html)
        self.assertIn("made < version", self.html)

    def test_오래된_작성본은_단추와_창에서_경고한다(self):
        self.assertIn('id="rd-stale"', self.html)
        self.assertIn("재작성 필요 ⚠", self.html)
        self.assertIn("옛 프로그램이 만든 작성본입니다", self.html)

    def test_완성본_열기는_폴더와_엑셀도_연다고_알린다(self):
        self.assertIn("작성본 폴더", self.html)
        self.assertIn("payload.folder_opened", self.html)

    def test_재작성_창에_첨부_엑셀도_보인다(self):
        self.assertIn("product.final_attachments", self.html)
        self.assertIn("첨부 엑셀: <code>", self.html)
