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


class OCR_판독(unittest.TestCase):
    """담당자 2026-09-07: "OCR 로 변환해서 읽으면 안 되는 거야?" — 된다, 한글 인식 모델이 있으면."""

    def test_이름표_뒤의_글을_값으로_삼는다(self):
        from pqr.engine.readers import change_ocr as X
        got = X.parse(["변경관리 요청서", "문서번호 CC-240723-08", "변경명 포장자재 규격 변경",
                       "변경사유", "원가 절감", "변경내용 PE병 두께 변경",
                       "관련제품 한림포비돈점안액", "QA 규격서 개정", "QC팀 시험방법 확인"])
        self.assertEqual(got["doc_no"], "CC-240723-08")
        self.assertEqual(got["title"], "포장자재 규격 변경")
        self.assertEqual(got["reason"], "원가 절감")          # 다음 줄에 있어도 찾는다
        self.assertEqual(got["products"], "한림포비돈점안액")
        self.assertEqual(got["actions"], [("QA", "규격서 개정"), ("QC", "시험방법 확인")])

    def test_못_읽은_칸은_비워_둔다(self):
        from pqr.engine.readers import change_ocr as X
        got = X.parse(["변경관리 요청서", "...", "___"])
        self.assertEqual((got["title"], got["reason"], got["actions"]), ("", "", []))

    def test_한글_모델이_없으면_그_길은_쓰지_않는다(self):
        from pqr.engine.readers import change_ocr as X
        from pqr.engine import handwriting
        old = handwriting.korean_engine
        handwriting.korean_engine = lambda: None
        try:
            with self.assertRaises(RuntimeError):
                X.read("/x/a.pdf")
        finally:
            handwriting.korean_engine = old


class 판독_결과_되쓰기(unittest.TestCase):
    """판독은 몇 분이 걸린다 — 한 번 읽은 것은 제품 폴더에 남겨 재작성 때 다시 읽지 않는다
    (담당자 2026-09-08: "보고서 작성 시간이 기존보다 오래 걸리는 이유가?")."""

    def setUp(self):
        import tempfile
        self.dir = tempfile.mkdtemp(prefix="pqr-cc-")
        self.pdf = os.path.join(self.dir, "12.CC-240723-08.pdf")
        with open(self.pdf, "wb") as handle:
            handle.write(b"%PDF-1.4 scan")
        self.notes = []

    def _note(self, item, path, why):
        self.notes.append((item, why))

    def _once(self, reader):
        import types
        import pqr.engine as pkg
        old = getattr(pkg, "claude_cli", None)
        pkg.claude_cli = types.SimpleNamespace(available=lambda: True, read_change=reader)
        sys.modules["pqr.engine.claude_cli"] = pkg.claude_cli
        try:
            return C._change_by_claude(self.pdf, self.dir, None, self._note)
        finally:
            if old is None:
                delattr(pkg, "claude_cli"); sys.modules.pop("pqr.engine.claude_cli", None)
            else:
                pkg.claude_cli = old; sys.modules["pqr.engine.claude_cli"] = old

    def test_두_번째부터는_다시_읽지_않는다(self):
        calls = {"n": 0}

        def reader(*a, **kw):
            calls["n"] += 1
            return {"doc_no": "", "title": "포장자재 변경", "description": "", "reason": "",
                    "products": "", "approved": "", "actions": [("QA", "규격서 개정")]}
        first = self._once(reader)
        self.assertEqual(first["title"], "포장자재 변경")
        self.assertTrue(os.path.isfile(os.path.join(self.dir, C.CHANGE_CACHE)))
        second = self._once(reader)                      # 두 번째 — 판독기를 부르지 않는다
        self.assertEqual(calls["n"], 1)
        self.assertEqual(second["title"], "포장자재 변경")
        self.assertEqual(second["actions"], [("QA", "규격서 개정")])

    def test_파일이_바뀌면_다시_읽는다(self):
        calls = {"n": 0}

        def reader(*a, **kw):
            calls["n"] += 1
            return {"doc_no": "", "title": "바뀐 변경명", "description": "", "reason": "",
                    "products": "", "approved": "", "actions": []}
        self._once(reader)
        with open(self.pdf, "ab") as handle:
            handle.write(b" more")                       # 크기가 달라지면 다른 파일로 본다
        self._once(reader)
        self.assertEqual(calls["n"], 2)


class 스캔_시험성적서(unittest.TestCase):
    """담당자 2026-09-08: "시험 성적을 제대로 못 읽는 것 같아" — 9.2.4 완제 성적서가
    세 쪽 모두 글자 0자인 스캔본이었다."""

    def setUp(self):
        import tempfile
        self.dir = tempfile.mkdtemp(prefix="pqr-coa-")
        self.pdf = os.path.join(self.dir, "LWY201.pdf")
        with open(self.pdf, "wb") as handle:
            handle.write(b"%PDF-1.4 scan")
        self.notes = []

    def _note(self, item, path, why):
        self.notes.append((item, why))

    def _run(self, reader):
        import types
        import pqr.engine as pkg
        old = getattr(pkg, "claude_cli", None)
        pkg.claude_cli = types.SimpleNamespace(available=lambda: True, read_coa=reader)
        sys.modules["pqr.engine.claude_cli"] = pkg.claude_cli
        try:
            return C._coa_by_claude(self.pdf, self.dir, None, self._note, "9.2.4")
        finally:
            if old is None:
                delattr(pkg, "claude_cli"); sys.modules.pop("pqr.engine.claude_cli", None)
            else:
                pkg.claude_cli = old; sys.modules["pqr.engine.claude_cli"] = old

    def _answer(self):
        return {"file": "LWY201.pdf", "lot": "LWY201", "appearance": "무색의 투명한 액",
                "assays": [{"part": "트레할로스수화물", "lo": "90.0", "hi": "110.0", "value": "99.8"}]}

    def test_읽고_되쓴다(self):
        calls = {"n": 0}

        def reader(*a, **kw):
            calls["n"] += 1
            return self._answer()
        first = self._run(reader)
        self.assertEqual(first["assays"][0]["value"], "99.8")
        self.assertTrue(os.path.isfile(os.path.join(self.dir, C.COA_CACHE)))
        self._run(reader)                                 # 두 번째는 판독기를 부르지 않는다
        self.assertEqual(calls["n"], 1)

    def test_읽어_낸_것이_없으면_까닭을_남긴다(self):
        got = self._run(lambda *a, **kw: {"file": "LWY201.pdf", "assays": []})
        self.assertIsNone(got)
        self.assertEqual(self.notes[0][0], "9.2.4")
        self.assertIn("읽어 낸 것이 없음", self.notes[0][1])

    def test_판독_파일은_자료로_세지_않는다(self):
        from pqr import build
        self.assertTrue(any(C.COA_CACHE.startswith(side) for side in build.SIDE_FILES))
