# -*- coding: utf-8 -*-
"""9.2 공정관리 — 화면의 칸은 하나, 그 안에서 공정별 세부 항을 다룬다.

담당자 2026-09-27: "팀마다 9.2항 상세 항이 달라 예를들어 타팀은 공정에 따라 9.2.6항까지
값이 기재되는데 어떻게하면 좋을까?" → 칸을 9.2 하나로 합치고, 기대 목록은 제형 군마다
config(item_subs)에 둔다. 목록을 정하지 않은 제형은 올라온 파일이 곧 목록이다.
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

from pqr import build as B, server as server_module          # noqa: E402
from pqr.engine import collect as C                          # noqa: E402
from pqr.sample import write_samples                         # noqa: E402

TODAY = "2026-08-27"
CONFIG = B.load_config()


class 항목_목록(unittest.TestCase):
    def test_9_2_는_한_칸이다(self):
        ids = [row[0] for row in CONFIG["items"]]
        self.assertIn("9.2", ids)
        for gone in ("9.2.1", "9.2.2", "9.2.3", "9.2.4"):
            self.assertNotIn(gone, ids)

    def test_세부_목록은_제형마다_다르다(self):
        self.assertEqual(len(B.item_subs(CONFIG, "9.2", "점안제")), 4)
        self.assertEqual(B.item_subs(CONFIG, "9.2", "정제"), [])   # 아직 정하지 않은 제형

    def test_세부_이름을_찾는다(self):
        self.assertEqual(B.sub_label(CONFIG, "9.2", "9.2.3"), "충전 완료 후")
        self.assertEqual(B.sub_label(CONFIG, "9.2", "9.2.9"), "")


class 세부_항_판정(unittest.TestCase):
    def state(self, form, names):
        return B.sub_status(CONFIG, "9.2", form, names)[1]

    def test_다_차야_완료(self):
        names = ["9.2.%d x.pdf" % n for n in (1, 2, 3, 4)]
        self.assertEqual(self.state("점안제", names), "y")

    def test_하나라도_비면_진행_중(self):
        self.assertEqual(self.state("점안제", ["9.2.1 x.pdf", "9.2.3 y.pdf"]), "p")

    def test_아무것도_없으면_미착수(self):
        self.assertEqual(self.state("점안제", []), "n")

    def test_목록을_정하지_않은_제형은_올라온_파일이_목록이다(self):
        self.assertEqual(self.state("정제", ["9.2.6 선별.pdf"]), "y")
        self.assertEqual(self.state("정제", []), "n")

    def test_9_2_6_처럼_목록에_없는_번호도_보여_준다(self):
        rows, _ = B.sub_status(CONFIG, "9.2", "점안제", ["9.2.6 선별.pdf"])
        self.assertIn("9.2.6", [row["id"] for row in rows])

    def test_한_파일에_담아_올려도_된다(self):
        """'9.2 공정관리 시험성적.pdf' 한 장이면 나눠 올리라고 막지 않는다."""
        self.assertEqual(self.state("점안제", ["9.2 공정관리 시험성적.pdf"]), "y")


class 파일_이름이_공정을_지킨다(unittest.TestCase):
    def test_화면은_9_2_로_모으고(self):
        m = B.item_matcher(CONFIG["items"])
        self.assertEqual(m("9.2.1 조제 완료 후.pdf"), "9.2")
        self.assertEqual(m("9.2.6 선별 완료 후.pdf"), "9.2")

    def test_보고서는_세부_번호를_그대로_본다(self):
        """engine 은 어느 공정인지 알아야 표를 채운다 — collect 는 세부 번호를 지킨다."""
        self.assertEqual(C._item_of("9.2.1 조제 완료 후 IPC.pdf"), "9.2.1")
        self.assertEqual(C._item_of("9.2.6 선별.pdf"), "9.2.6")
        self.assertEqual(C._item_of("9.2 공정관리.pdf"), "9.2")

    def test_다른_항의_깊은_번호는_예전대로_위_항으로(self):
        self.assertEqual(C._item_of("8.2.2.1 P12060 자재.xlsx"), "8.2.2")


class 화면에_내려보내는_세부_상태(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.dir, True)
        write_samples(self.dir, layout="tree")
        self.folder = os.path.join(
            self.dir, next(n for n in os.listdir(self.dir) if n.startswith("HP-110")))

    def product(self):
        data = B.build(input_dir=self.dir, today=TODAY)
        return next(p for p in data["products"] if p["code"] == "HP-110")

    def test_제품마다_세부_상태가_실린다(self):
        open(os.path.join(self.folder, "9.2.1 조제 완료 후.pdf"), "wb").close()
        rows = self.product()["subs"]["9.2"]
        got = {row["id"]: row["ok"] for row in rows}
        self.assertTrue(got["9.2.1"])
        self.assertFalse(got["9.2.3"])

    def test_세부가_덜_차면_칸은_진행_중(self):
        open(os.path.join(self.folder, "9.2.1 조제 완료 후.pdf"), "wb").close()
        product = self.product()
        ids = [row[0] for row in CONFIG["items"]]
        self.assertEqual(product["checks"][ids.index("9.2")], "p")

    def test_세부를_다_채우면_완료(self):
        for n in (1, 2, 3, 4):
            open(os.path.join(self.folder, "9.2.%d 자료.pdf" % n), "wb").close()
        product = self.product()
        ids = [row[0] for row in CONFIG["items"]]
        self.assertEqual(product["checks"][ids.index("9.2")], "y")


class 세부_항으로_올리기(unittest.TestCase):
    """올리기 창에서 공정을 고르면 그 번호로 저장된다 — HTTP 로 끝까지 밟아 본다."""

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

    def upload(self, item, filename, sub=""):
        import uuid
        boundary = uuid.uuid4().hex
        parts = []
        fields = [("product", "HP-110"), ("item", item)] + ([("sub", sub)] if sub else [])
        for name, value in fields:
            parts.append(("--%s\r\nContent-Disposition: form-data; name=\"%s\"\r\n\r\n%s\r\n"
                          % (boundary, name, value)).encode())
        parts.append(("--%s\r\nContent-Disposition: form-data; name=\"file\"; "
                      "filename=\"%s\"\r\n\r\n" % (boundary, filename)).encode())
        parts.append(b"x" * 40 + b"\r\n" + ("--%s--\r\n" % boundary).encode())
        request = urllib.request.Request(
            self.base + "/api/upload", data=b"".join(parts),
            headers={"Content-Type": "multipart/form-data; boundary=" + boundary})
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            return json.loads(error.read().decode("utf-8"))

    def test_고른_공정_번호로_저장한다(self):
        got = self.upload("9.2", "충전 성적서.pdf", sub="9.2.3")
        self.assertTrue(got.get("ok"), got.get("error"))
        self.assertTrue(got["name"].startswith("9.2.3 충전 완료 후 - "), got["name"])

    def test_목록에_없는_공정도_받는다(self):
        got = self.upload("9.2", "선별 성적서.pdf", sub="9.2.6")
        self.assertTrue(got.get("ok"), got.get("error"))
        self.assertTrue(got["name"].startswith("9.2.6"), got["name"])

    def test_이미_번호로_시작하면_그대로_둔다(self):
        got = self.upload("9.2", "9.2.3 충전 완료 후.pdf", sub="9.2.3")
        self.assertEqual(got["name"], "9.2.3 충전 완료 후.pdf")

    def test_세부를_고르지_않으면_항_번호로_저장한다(self):
        got = self.upload("9.2", "공정관리.pdf")
        self.assertTrue(got["name"].startswith("9.2 공정관리 시험성적 - "), got["name"])

    def test_한_칸에_세부_파일이_모두_보인다(self):
        self.upload("9.2", "조제.pdf", sub="9.2.1")
        self.upload("9.2", "충전.pdf", sub="9.2.3")
        request = urllib.request.Request(
            self.base + "/api/item-files",
            data=json.dumps({"product": "HP-110", "item": "9.2"}).encode("utf-8"),
            headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(request, timeout=20) as response:
            got = json.loads(response.read().decode("utf-8"))
        names = sorted(row["name"] for row in got["files"])
        self.assertEqual(len(names), 2, names)
        self.assertTrue(names[0].startswith("9.2.1"))
        self.assertTrue(names[1].startswith("9.2.3"))


if __name__ == "__main__":
    unittest.main()
