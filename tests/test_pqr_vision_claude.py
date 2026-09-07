"""Claude 판독 — PC 판독과 같은 꼴의 판독 결과를 만든다 (담당자 2026-09-07: "PC 로 판독하면 시간이 너무 오래 걸려")."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import vision_claude as V                            # noqa: E402


def _rec(**kw):
    base = {"product_name": "퀴노비드안연고(내수용)", "lot": "OEX101", "test_type": "장기", "market": "",
            "mfg_date": "2024-01-10", "expiry_date": "2027.01.09", "package": "5g/Tube",
            "storage": "25±2°C, 60±5%RH", "uncertain": [],
            "points": [
                {"label": "초기", "test_date": "2024.01.15", "reviewer_date": "2024.01.20",
                 "tested": True, "date_confidence": 0.95,
                 "assays": [{"name": "오플록사신", "value": "101.2", "confidence": 0.93}]},
                {"label": "3개월", "test_date": "2024.04.15", "reviewer_date": "",
                 "tested": True, "date_confidence": 0.4,
                 "assays": [{"name": "함량", "value": "99.8", "confidence": 0.5}]},
                {"label": "6M", "test_date": "", "reviewer_date": "", "tested": False,
                 "date_confidence": 0.0, "assays": []},
            ]}
    base.update(kw)
    return base


class 판독결과_변환(unittest.TestCase):
    def test_pc_판독과_같은_꼴로_바꾼다(self):
        got = V.to_log(_rec(), "/x/13 장기 안정성시험일지(내수용).pdf", {"오플록사신": (90.0, 110.0)})
        self.assertEqual(got["lot"], "OEX101")
        self.assertEqual((got["kind"], got["market"], got["year"]), ("장기", "내수", "2024"))
        self.assertEqual((got["mfg"], got["expiry"]), ("2024.01.10", "2027.01.09"))
        self.assertEqual(got["source"], "13 장기 안정성시험일지(내수용).pdf")
        self.assertEqual([p["period"] for p in got["points"]], ["Initial", "3M"])   # 시험 안 한 시점은 빼고
        self.assertEqual(got["points"][0]["done"], "2024.01.20")                    # 결재일이 완료 일자
        self.assertEqual(got["points"][0]["assays"], {"오플록사신": 101.2})
        self.assertEqual(got["points"][0]["unsure"], [])

    def test_확신이_낮으면_애매로_남긴다(self):
        got = V.to_log(_rec(), "/x/a.pdf", {"오플록사신": (90.0, 110.0)})
        second = got["points"][1]
        self.assertEqual(second["assays"], {"오플록사신": 99.8})       # 이름 없는 '함량' 도 성분에 맞춘다
        self.assertEqual(second["unsure"], ["done", "오플록사신"])     # 일자·함량 모두 애매
        self.assertEqual(second["done"], "2024.04.15")

    def test_제조번호가_없으면_버린다(self):
        self.assertIsNone(V.to_log(_rec(lot=""), "/x/a.pdf"))

    def test_읽은_시점이_없으면_버린다(self):
        rec = _rec(points=[{"label": "3M", "test_date": "", "reviewer_date": "", "tested": False,
                            "date_confidence": 0.0, "assays": []}])
        self.assertIsNone(V.to_log(rec, "/x/a.pdf"))

    def test_시장은_제품명이나_파일_이름에서(self):
        self.assertEqual(V.to_log(_rec(product_name="퀴노비드안연고(수출용)"), "/x/a.pdf")["market"], "수출")
        self.assertEqual(V.to_log(_rec(product_name=""), "/x/13 수출용 일지.pdf")["market"], "수출")
        self.assertEqual(V.to_log(_rec(product_name="", market="내수"), "/x/a.pdf")["market"], "내수")

    def test_시점_이름_고르기(self):
        self.assertEqual([V._period(x) for x in ("초기", "Initial", "3개월", "3M", " 12 M ", "")],
                         ["Initial", "Initial", "3M", "3M", "12M", ""])

    def test_writer_훅은_아무것도_하지_않는다(self):
        self.assertIsNone(V.read_stability_into(object()))


class 여러쪽_합치기(unittest.TestCase):
    def test_한_lot_이_두_쪽에_걸쳐_있으면_합친다(self):
        class _FakeClient(object):
            pass

        pages = [_rec(points=[{"label": "초기", "test_date": "2024.01.15", "reviewer_date": "2024.01.20",
                               "tested": True, "date_confidence": 0.9,
                               "assays": [{"name": "함량", "value": "101.2", "confidence": 0.9}]}]),
                 _rec(points=[{"label": "12M", "test_date": "2025.01.15", "reviewer_date": "2025.01.20",
                               "tested": True, "date_confidence": 0.9,
                               "assays": [{"name": "함량", "value": "100.1", "confidence": 0.9}]}])]
        calls = {"n": 0}

        def fake_page(client, png):
            rec = pages[calls["n"]]
            calls["n"] += 1
            return rec

        old_page, old_png, old_client = V.read_page, V._png, V._client
        V.read_page, V._png, V._client = fake_page, (lambda p, n: ""), (lambda: _FakeClient())
        try:
            from pqr.engine import handwriting
            old_count = handwriting.page_count
            handwriting.page_count = lambda path: 2
            try:
                logs = V.read_logs(["/x/a.pdf"], None, None, workers=1)
            finally:
                handwriting.page_count = old_count
        finally:
            V.read_page, V._png, V._client = old_page, old_png, old_client
        self.assertEqual(len(logs), 1)
        self.assertEqual([p["period"] for p in logs[0]["points"]], ["Initial", "12M"])


if __name__ == "__main__":
    unittest.main()


class API_키(unittest.TestCase):
    """키는 환경 변수나 프로그램 폴더의 API-KEY.txt 로 준다 (담당자 2026-09-07)."""

    def setUp(self):
        import tempfile
        from pqr.engine import vision
        self.vision = vision
        self.dir = tempfile.mkdtemp(prefix="pqr-key-")
        self.old = os.environ.pop("ANTHROPIC_API_KEY", None)

    def tearDown(self):
        os.environ.pop("ANTHROPIC_API_KEY", None)
        if self.old is not None:
            os.environ["ANTHROPIC_API_KEY"] = self.old

    def _write(self, text):
        with open(os.path.join(self.dir, self.vision.KEY_FILE), "w", encoding="utf-8") as fh:
            fh.write(text)

    def test_파일에서_키를_읽고_공백과_따옴표를_뗀다(self):
        self._write('﻿  "sk-ant-test123"  \n')
        self.assertEqual(self.vision.key_from_file(self.dir), "sk-ant-test123")

    def test_설명_줄과_빈_줄은_건너뛴다(self):
        self._write("# 여기에 키를 적으세요\n\nsk-ant-real\n")
        self.assertEqual(self.vision.key_from_file(self.dir), "sk-ant-real")

    def test_파일이_없으면_빈_값(self):
        self.assertEqual(self.vision.key_from_file(self.dir), "")

    def test_환경_변수가_있으면_켜진다(self):
        os.environ["ANTHROPIC_API_KEY"] = "sk-ant-env"
        self.assertTrue(self.vision.available())
