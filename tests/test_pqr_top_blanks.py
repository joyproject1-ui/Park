"""쪽 맨 위에 홀로 남은 빈 줄은 지운다 (담당자 2026-09-07: "del 키 눌러서 빈 공간 없게 만들어 줘")."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import convert                                     # noqa: E402


class _Range(object):
    def __init__(self, para):
        self.para = para

    @property
    def Text(self):
        return self.para.text + "\r"

    def Information(self, what):
        return self.para.page if what == convert.WD_PAGE else self.para.in_table

    def Delete(self):
        self.para.doc.gone.append(self.para.name)
        self.para.doc.paras.remove(self.para)


class _Para(object):
    def __init__(self, doc, name, text, page, in_table=False):
        self.doc, self.name, self.text, self.page, self.in_table = doc, name, text, page, in_table
        self.Range = _Range(self)


class _Paragraphs(object):
    def __init__(self, doc):
        self.doc = doc

    def __call__(self, i):
        return self.doc.paras[i - 1]

    @property
    def Count(self):
        return len(self.doc.paras)


class _Doc(object):
    """Word 문서 흉내 — (이름, 글, 쪽, 표 안인가) 차례로."""

    def __init__(self, rows):
        self.paras, self.gone = [], []
        self.Paragraphs = _Paragraphs(self)
        for name, text, page, in_table in rows:
            self.paras.append(_Para(self, name, text, page, in_table))


class 쪽_맨_위_빈_줄(unittest.TestCase):
    def _run(self, rows):
        doc = _Doc(rows)
        convert._drop_top_blanks(doc)
        return doc.gone

    def test_쪽_첫_줄이_비어_있으면_지운다(self):
        self.assertEqual(self._run([("a", "9.1 시험 결과", 1, False),
                                    ("blank", "", 2, False),
                                    ("b", "9.2 세부", 2, False)]), ["blank"])

    def test_쪽_한가운데_빈_줄은_그대로(self):
        self.assertEqual(self._run([("a", "글", 2, False), ("blank", "  ", 2, False),
                                    ("b", "글", 2, False)]), [])

    def test_표_안의_빈_문단은_건드리지_않는다(self):
        self.assertEqual(self._run([("a", "글", 1, False), ("blank", "", 2, True)]), [])

    def test_표와_표_사이는_남긴다(self):
        # 지우면 두 표가 하나로 붙는다
        self.assertEqual(self._run([("t1", "표 끝", 1, True), ("blank", "", 2, False),
                                    ("t2", "표 머리", 2, True)]), [])

    def test_표_다음이면서_글_앞이면_지운다(self):
        self.assertEqual(self._run([("t1", "표 끝", 1, True), ("blank", "", 2, False),
                                    ("p", "10. 적격성", 2, False)]), ["blank"])

    def test_맨_끝_빈_줄도_쪽_첫_줄이면_지운다(self):
        self.assertEqual(self._run([("a", "글", 1, False), ("blank", "", 2, False)]), ["blank"])

    def test_첫_문단은_건드리지_않는다(self):
        self.assertEqual(self._run([("blank", "", 1, False), ("a", "글", 1, False)]), [])


class 세_갈래_모두(unittest.TestCase):
    """VBScript·PowerShell·pywin32 세 길 모두에서 같은 손질을 한다."""

    def test_vbscript_와_powershell_에도_들어_있다(self):
        import inspect
        vbs = inspect.getsource(convert._fields_via_vbscript)
        ps = inspect.getsource(convert._fields_via_powershell)
        win = inspect.getsource(convert._fields_via_pywin32)
        self.assertIn("TOP_BLANKS_VBS", vbs)
        self.assertIn("TOP_BLANKS_PS", ps)
        self.assertIn("_drop_top_blanks", win)
        for source in (vbs, ps, win):                 # 지운 뒤에는 쪽을 다시 나누고 목차를 다시 계산한다
            self.assertIn("Repaginate", source)


if __name__ == "__main__":
    unittest.main()
