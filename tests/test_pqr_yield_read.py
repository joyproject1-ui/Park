# -*- coding: utf-8 -*-
"""7항 수율 판독 — 서식 프로필, 판독값 검산, 수율현황표 초안, 화면에서 부르는 길.

담당자 2026-09-27: "우리 팀은 직접 작성하는 것이 편하다는 의견이고 다른 팀은 기록서 스캔본에서
수율이 기재된 페이지를 직접 AI 가 읽어서 정리해줬으면 좋겠다는 의견이야." → 두 길을 다 둔다.
쪽수는 제품마다가 아니라 "수탁사 × 기록서 서식마다 한 번 … 버전마다 관리".
"""
import json
import os
import shutil
import tempfile
import threading
import unittest
import urllib.error
import urllib.request

os.environ["PQR_REVIEW"] = "0"

from pqr import server as server_module                     # noqa: E402
from pqr.engine import claude_cli, yield_profile as P, yield_read as Y   # noqa: E402
from pqr.engine.readers import yield_sheet                  # noqa: E402
from pqr.sample import write_samples                        # noqa: E402

TODAY = "2026-08-27"

READING = {
    "lots": [
        {"lot": "GVY601", "values": {"조제": 99.85, "충전": 95.31, "포장": 99.99},
         "production": {}, "page": {"조제": 7, "충전": 12, "포장": 18}, "unsure": []},
        {"lot": "GVYD01", "values": {"조제": 99.85, "충전": 98.13},
         "production": {}, "page": {"조제": 7, "충전": 12}, "unsure": ["포장"]},
    ],
    "specs": {"조제": "99.5±0.45%", "충전": "98.0±3.0", "포장": "98.0±2.0%"},
}


class 쪽수_적기(unittest.TestCase):
    def test_공정_이름과_쪽을_함께_적는다(self):
        self.assertEqual(P.parse_pages("조제 7, 충전 12, 포장 18-19"),
                         {"조제": [7], "충전": [12], "포장": [18, 19]})

    def test_콜론도_슬래시도_받는다(self):
        self.assertEqual(P.parse_pages("조제: 7 / 충전: 12"), {"조제": [7], "충전": [12]})

    def test_이름_없이_쪽만_적어도_된다(self):
        self.assertEqual(P.parse_pages("7,12,18-19"), {"": [7, 12, 18, 19]})

    def test_쪽이_겹쳐도_한_번만(self):
        self.assertEqual(P.page_list({"pages": {"조제": [7], "충전": [12, 7]}}), [7, 12])


class 서식_프로필(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.dir, True)

    def test_판이_다르면_다른_프로필이다(self):
        a = P.make_id("코오롱제약", "정산기록서", "MBRM-00004", "4.0")
        b = P.make_id("코오롱제약", "정산기록서", "MBRM-00004", "5.0")
        self.assertNotEqual(a, b)
        self.assertIn("4.0", a)

    def test_저장하고_다시_읽는다(self):
        data = P.empty()
        one = P.put(data, {"maker": "코오롱제약", "record": "정산기록서",
                           "form_no": "MBRM-00004", "rev": "4.0",
                           "pages": P.parse_pages("조제 1, 충전 2")})
        P.assign(data, "QC1-2080", one["id"])
        P.save(self.dir, data)
        again = P.load(self.dir)
        self.assertEqual(len(again["profiles"]), 1)
        self.assertEqual(P.for_product(again, "QC1-2080")["id"], one["id"])
        self.assertEqual(P.page_list(P.for_product(again, "qc1-2080")), [1, 2])

    def test_공통_폴더에_둔다(self):
        P.save(self.dir, P.empty())
        self.assertTrue(os.path.isfile(os.path.join(self.dir, "공통", P.FILE_NAME)))

    def test_서식_판이_바뀌면_알아챈다(self):
        one = {"form_no": "MBRM-00004", "rev": "4.0"}
        self.assertTrue(P.rev_changed(one, "MBRM-00004", "5.0"))
        self.assertFalse(P.rev_changed(one, "MBRM 00004", "4.0"))
        self.assertFalse(P.rev_changed(one, "", ""))       # 못 읽었으면 바뀌었다고 하지 않는다

    def test_판을_올리면_옛_프로필은_남는다(self):
        data = P.empty()
        old = P.put(data, {"maker": "동화약품", "record": "제조총괄표", "rev": "1.0",
                           "pages": {"조제": [3]}})
        P.put(data, P.bump(old, "2.0", {"조제": [4]}))
        self.assertEqual(len(data["profiles"]), 2)
        self.assertEqual(P.find(data, old["id"])["pages"], {"조제": [3]})


class 판독값_검산(unittest.TestCase):
    def test_숫자가_아니면_버리고_확인_필요로_남긴다(self):
        got = claude_cli._yield_clean({"lots": [{"lot": "gvy601",
                                                 "values": {"조제": "99.85%", "충전": "흐림"}}]})
        self.assertEqual(got["lots"][0]["values"], {"조제": 99.85})
        self.assertIn("충전", got["lots"][0]["unsure"])
        self.assertEqual(got["lots"][0]["lot"], "GVY601")

    def test_수율일_수_없는_값은_버린다(self):
        got = claude_cli._yield_clean({"lots": [{"lot": "A1234", "values": {"조제": "150"}}]})
        self.assertEqual(got["lots"][0]["values"], {})

    def test_기준을_벗어나면_짚는다(self):
        notes = Y.check({"lots": [{"lot": "A1234", "values": {"조제": 80.0}, "unsure": []}],
                         "specs": {"조제": "99.5±0.45%"}})
        self.assertTrue(any("기준" in n for n in notes), notes)

    def test_제조내역에_없는_제조번호를_짚는다(self):
        notes = Y.check(READING, known_lots=["GVY601"])
        self.assertTrue(any("GVYD01" in n and "제조내역" in n for n in notes), notes)

    def test_흐린_칸을_짚는다(self):
        notes = Y.check(READING)
        self.assertTrue(any("GVYD01 포장" in n for n in notes), notes)

    def test_기준_읽기(self):
        self.assertEqual(Y._spec_range("99.5±0.45%"), (99.05, 99.95))
        self.assertEqual(Y._spec_range("98.0~102.0"), (98.0, 102.0))
        self.assertEqual(Y._spec_range("98.0 이상"), (98.0, None))
        self.assertEqual(Y._spec_range("적합"), (None, None))


class 수율현황표_초안(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.dir, True)
        self.path = Y.write_sheet(self.dir, READING, Y.check(READING))

    def test_보고서가_읽는_모양으로_쓴다(self):
        """이 파일은 '보고서 작성' 이 그대로 읽어야 한다 — 읽는 쪽으로 되읽어 확인한다."""
        got = dict(yield_sheet.read_yields(self.path))
        self.assertEqual(got["GVY601"]["조제"], "99.85")
        self.assertEqual(got["GVY601"]["충전"], "95.31")
        self.assertEqual(yield_sheet.read_specs(self.path)["조제"], "99.5±0.45%")

    def test_못_읽은_칸은_확인_필요로_둔다(self):
        got = dict(yield_sheet.read_yields(self.path))
        self.assertEqual(got["GVYD01"]["포장"], "확인 필요")

    def test_근거_쪽수를_비고에_남긴다(self):
        import openpyxl
        ws = openpyxl.load_workbook(self.path).active
        text = "\n".join(str(c.value or "") for row in ws.iter_rows() for c in row)
        self.assertIn("7, 12, 18쪽", text)

    def test_초안이라고_적어_둔다(self):
        import openpyxl
        ws = openpyxl.load_workbook(self.path).active
        text = "\n".join(str(c.value or "") for row in ws.iter_rows() for c in row)
        self.assertIn("초안", text)

    def test_판독_결과를_제품_폴더에_남긴다(self):
        Y.save_reading(self.dir, READING, ["확인 필요"], {"profile": "x"})
        got = Y.load_reading(self.dir)
        self.assertEqual(len(got["lots"]), 2)
        self.assertEqual(got["profile"], "x")


class 공정_이름_대응(unittest.TestCase):
    def test_수탁사_말을_우리_말로_바꾼다(self):
        got = Y.apply_names({"캡슐충전": 99.1, "조제": 98.0},
                            {"names": {"캡슐충전": "충전"}})
        self.assertEqual(got, {"충전": 99.1, "조제": 98.0})


class 담당자가_만든_표가_먼저다(unittest.TestCase):
    def test_직접_만든_수율현황표를_알아본다(self):
        folder = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, folder, True)
        open(os.path.join(folder, "7. 수율현황표 - 개인.xlsx"), "wb").close()
        self.assertTrue(Y.own_sheet(folder))

    def test_판독본은_담당자_표로_치지_않는다(self):
        folder = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, folder, True)
        open(os.path.join(folder, Y.SHEET_NAME), "wb").close()
        self.assertIsNone(Y.own_sheet(folder))


class 화면에서_부르는_길(unittest.TestCase):
    """단추를 누르는 길은 HTTP 로 끝까지 밟아 본다 (규정집: 안쪽 함수만 시험하면 놓친다)."""

    def setUp(self):
        self.dir = tempfile.mkdtemp()
        write_samples(self.dir, layout="tree")
        self.folder = next(n for n in os.listdir(self.dir) if n.startswith("HP-110"))
        self.out = tempfile.mkdtemp()
        self.httpd = server_module.serve(self.dir, host="127.0.0.1", port=0, out_dir=self.out,
                                         today=TODAY, log=lambda *a: None)
        threading.Thread(target=self.httpd.serve_forever, daemon=True).start()
        self.base = "http://127.0.0.1:%d" % self.httpd.server_port

    def tearDown(self):
        self.httpd.shutdown()
        self.httpd.server_close()
        shutil.rmtree(self.dir, ignore_errors=True)
        shutil.rmtree(self.out, ignore_errors=True)

    def call(self, path, body):
        request = urllib.request.Request(
            self.base + path, data=json.dumps(body).encode("utf-8"),
            headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            return json.loads(error.read().decode("utf-8"))

    def test_프로필을_만들고_제품에_배정한다(self):
        got = self.call("/api/yield-profile", {
            "profile": {"maker": "코오롱제약", "record": "정산기록서",
                        "form_no": "MBRM-00004", "rev": "4.0",
                        "pages": "조제 1, 충전 2", "rules": {"note": "공정수율 행만"}}})
        self.assertTrue(got.get("ok"), got.get("error"))
        pid = got["saved"]
        got = self.call("/api/yield-profile", {"product": "HP-110", "profile_id": pid})
        self.assertEqual(got["assigned"]["HP-110"], pid)
        again = self.call("/api/yield-profiles", {})
        self.assertEqual([p["id"] for p in again["profiles"]], [pid])
        self.assertTrue(os.path.isfile(os.path.join(self.dir, "공통", P.FILE_NAME)))

    def test_쪽수가_없으면_거절한다(self):
        got = self.call("/api/yield-profile", {
            "profile": {"maker": "동화약품", "record": "제조총괄표", "pages": ""}})
        self.assertFalse(got.get("ok"))
        self.assertIn("쪽", got.get("error", ""))

    def test_프로필_없이_판독하면_무엇이_필요한지_알려_준다(self):
        got = self.call("/api/read-yield", {"product": "HP-110"})
        self.assertFalse(got.get("ok"))
        self.assertEqual(got.get("need"), "profile")
        self.assertIn("프로필", got.get("why", ""))

    def test_판독_기록을_제품_폴더에_남긴다(self):
        self.call("/api/read-yield", {"product": "HP-110"})
        self.assertTrue(os.path.isfile(os.path.join(self.dir, self.folder, Y.LOG_NAME)))

    def test_제품_목록을_내려보낸다(self):
        got = self.call("/api/yield-profiles", {})
        self.assertTrue(got.get("ok"))
        self.assertIsInstance(got.get("rows"), list)


if __name__ == "__main__":
    unittest.main()
