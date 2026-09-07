"""글자 없는 스캔 변경요청서는 Claude 로 읽는다 (담당자 2026-09-07:
"변경요청서 읽으면 돼 PDF 라서 못 읽는 거야?", "못 읽으면 다른 방법을 사용해서라도 읽게 해야지")."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import collect as C                                   # noqa: E402


class _Fake(object):
    """vision / claude_cli 흉내."""

    def __init__(self, on, answer=None, error=None):
        self.on, self.answer, self.error, self.calls = on, answer, error, 0

    def available(self):
        return self.on

    def _read(self, *a, **kw):
        self.calls += 1
        if self.error is not None:
            raise self.error
        return self.answer


class 스캔_변경요청서(unittest.TestCase):
    def setUp(self):
        self.notes = []
        self.note = lambda item, path, why: self.notes.append((item, why))

    def _run(self, api=None, cli=None):
        # `from . import vision` 은 이미 불러 둔 꾸러미 속성을 쓴다 — sys.modules 만 바꾸면 안 듣는다
        import types
        import pqr.engine as pkg
        mods = {}
        if api is not None:
            mods["vision"] = types.SimpleNamespace(available=api.available)
            mods["vision_claude"] = types.SimpleNamespace(read_change=api._read)
        if cli is not None:
            mods["claude_cli"] = types.SimpleNamespace(available=cli.available,
                                                       read_change=cli._read)
        old = {}
        for name, mod in mods.items():
            old[name] = getattr(pkg, name, None)
            setattr(pkg, name, mod)
            sys.modules["pqr.engine." + name] = mod
        try:
            return C._change_by_claude("/x/12.CC-240723-08.pdf", "/제품", None, self.note)
        finally:
            for name, mod in old.items():
                if mod is None:
                    delattr(pkg, name)
                    sys.modules.pop("pqr.engine." + name, None)
                else:
                    setattr(pkg, name, mod)
                    sys.modules["pqr.engine." + name] = mod

    def _answer(self, **kw):
        base = {"doc_no": "", "title": "포장자재 변경", "description": "", "reason": "",
                "products": "", "approved": "", "actions": [("QA", "규격서 개정")]}
        base.update(kw)
        return base

    def test_api_키가_있으면_그것으로_먼저_읽는다(self):
        api, cli = _Fake(True, self._answer()), _Fake(True, self._answer())
        got = self._run(api, cli)
        self.assertEqual(got["title"], "포장자재 변경")
        self.assertEqual((api.calls, cli.calls), (1, 0))
        self.assertEqual(got["doc_no"], "CC-240723-08")      # 파일 이름에서 문서번호를 채운다

    def test_api_가_없으면_이_PC_의_claude_code_로(self):
        api, cli = _Fake(False), _Fake(True, self._answer())
        got = self._run(api, cli)
        self.assertEqual(got["title"], "포장자재 변경")
        self.assertEqual(cli.calls, 1)

    def test_앞_갈래가_터지면_뒤_갈래로_이어_간다(self):
        api = _Fake(True, error=RuntimeError("판독 거부"))
        cli = _Fake(True, self._answer())
        self.assertIsNotNone(self._run(api, cli))
        self.assertEqual(cli.calls, 1)

    def test_모두_읽지_못하면_까닭을_남긴다(self):
        api = _Fake(True, error=RuntimeError("판독 거부"))
        cli = _Fake(True, self._answer(title="", actions=[]))
        self.assertIsNone(self._run(api, cli))
        self.assertEqual(len(self.notes), 1)
        self.assertIn("판독 거부", self.notes[0][1])
        self.assertIn("읽어 낸 것이 없음", self.notes[0][1])

    def test_판독기가_하나도_없으면_설치를_알려_준다(self):
        self.assertIsNone(self._run(_Fake(False), _Fake(False)))
        self.assertIn("PQR-Claude설치.bat", self.notes[0][1])


if __name__ == "__main__":
    unittest.main()
