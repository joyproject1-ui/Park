"""이 PC 의 Claude Code 로 13항 손글씨를 읽는 길 (담당자 2026-09-07: "이것을 자동화하면 안 돼?")."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import claude_cli as C                                # noqa: E402


class claude_찾기(unittest.TestCase):
    def test_담당자가_적어_둔_경로를_먼저_쓴다(self):
        # 담당자 2026-09-07: "이 PC 로 Claude Code 로 작업하고 있는데 없다는 게 무슨 말이냐"
        # — PATH 에 없어도 'CLAUDE 경로.txt' 나 PQR_CLAUDE_EXE 로 짚어 줄 수 있다.
        import tempfile
        fake = os.path.join(tempfile.mkdtemp(prefix="pqr-claude-"), "claude.exe")
        with open(fake, "w", encoding="utf-8") as handle:
            handle.write("")
        old = os.environ.get("PQR_CLAUDE_EXE")
        os.environ["PQR_CLAUDE_EXE"] = fake
        try:
            self.assertEqual(C._exe(), fake)
            self.assertTrue(C.available())
        finally:
            if old is None:
                del os.environ["PQR_CLAUDE_EXE"]
            else:
                os.environ["PQR_CLAUDE_EXE"] = old

    def test_찾아본_자리를_알려_준다(self):
        places = C.places()
        self.assertTrue(any(p.endswith("claude.exe") for p in places))
        self.assertTrue(any(".local" in p for p in places))


class 답에서_JSON_꺼내기(unittest.TestCase):
    def test_코드_표시와_설명이_붙어_있어도_꺼낸다(self):
        text = '읽었습니다.\n```json\n{"logs": [{"lot": "OEX101"}]}\n```\n확인하세요.'
        self.assertEqual(C._json_object(text), {"logs": [{"lot": "OEX101"}]})

    def test_글_안의_중괄호에_속지_않는다(self):
        text = '{"logs": [{"lot": "OEX101", "note": "값 { 이상 }"}]}'
        self.assertEqual(C._json_object(text)["logs"][0]["note"], "값 { 이상 }")

    def test_JSON_이_없으면_알린다(self):
        with self.assertRaises(ValueError):
            C._json_object("못 읽었습니다")


class 받은_기록_다듬기(unittest.TestCase):
    PATHS = ["/x/13 장기 안정성시험일지(내수용).pdf"]

    def _one(self, **kw):
        base = {"lot": "oex101", "year": "2024", "kind": "장기", "market": "내수",
                "pack": "5g tube/갑", "store": "25±2°C", "mfg": "2024-01-10", "expiry": "2027.01.09",
                "source": "13 장기 안정성시험일지(내수용).pdf",
                "points": [{"period": "12개월", "done": "2025-03-10",
                            "assays": {"함량": "101.2 %"}, "unsure": []},
                           {"period": "초기", "done": "", "assays": {"오플록사신": 99.0}, "unsure": []}]}
        base.update(kw)
        return base

    def test_판독_파일_꼴로_바꾼다(self):
        got = C._clean([self._one()], self.PATHS, {"오플록사신": (90.0, 110.0)})[0]
        self.assertEqual(got["lot"], "OEX101")                       # 대문자로
        self.assertEqual(got["mfg"], "2024.01.10")                   # 날짜 꼴 통일
        self.assertEqual([p["period"] for p in got["points"]], ["Initial", "12M"])   # 초기가 먼저
        self.assertEqual(got["points"][1]["assays"], {"오플록사신": 101.2})           # 성분 이름·숫자로
        self.assertEqual(got["points"][0]["unsure"], ["done"])       # 완료 일자가 없으면 애매
        self.assertEqual(got["source"], "13 장기 안정성시험일지(내수용).pdf")

    def test_제조번호나_시점이_없으면_버린다(self):
        self.assertEqual(C._clean([self._one(lot="")], self.PATHS), [])
        self.assertEqual(C._clean([self._one(points=[])], self.PATHS), [])
        self.assertEqual(C._clean(["글자"], self.PATHS), [])

    def test_모르는_구분과_시장은_비워_둔다(self):
        got = C._clean([self._one(kind="가속", market="해외")], self.PATHS)[0]
        self.assertEqual((got["kind"], got["market"]), ("장기", ""))
        self.assertFalse(got["market_hint"])

    def test_파일_이름을_달리_적어_와도_한_개면_바로잡는다(self):
        got = C._clean([self._one(source="엉뚱한 이름.pdf")], self.PATHS)[0]
        self.assertEqual(got["source"], "13 장기 안정성시험일지(내수용).pdf")


class 설치_확인(unittest.TestCase):
    def test_없으면_읽지_않는다(self):
        old = C._exe
        C._exe = lambda: ""
        try:
            self.assertFalse(C.available())
            with self.assertRaises(RuntimeError):
                C.read_logs(["/x/a.pdf"])
        finally:
            C._exe = old


if __name__ == "__main__":
    unittest.main()
