"""업로드 서버 테스트 — 실제 포트를 열되 외부 통신은 하지 않습니다."""

import json
import os
import shutil
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
import uuid

from pqr import server as server_module
from pqr.sample import write_samples

TODAY = "2026-08-27"

STABILITY_CSV = """배치번호,보관조건,개월,시험항목,측정값,단위,하한,상한
110-S1,장기(25℃/60%RH),0,함량,99.8,%,95,105
110-S1,장기(25℃/60%RH),6,함량,99.1,%,95,105
110-S1,장기(25℃/60%RH),12,함량,98.4,%,95,105
"""


def multipart(fields, files):
    """테스트용 multipart 본문을 만듭니다."""
    boundary = uuid.uuid4().hex
    parts = []
    for name, value in fields.items():
        parts.append(("--%s\r\nContent-Disposition: form-data; name=\"%s\"\r\n\r\n%s\r\n"
                      % (boundary, name, value)).encode("utf-8"))
    for name, (filename, payload) in files.items():
        parts.append(("--%s\r\nContent-Disposition: form-data; name=\"%s\"; filename=\"%s\"\r\n"
                      "Content-Type: application/octet-stream\r\n\r\n"
                      % (boundary, name, filename)).encode("utf-8"))
        parts.append(payload + b"\r\n")
    parts.append(("--%s--\r\n" % boundary).encode("utf-8"))
    return b"".join(parts), "multipart/form-data; boundary=%s" % boundary


class ServerTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        write_samples(self.dir, layout="tree")
        self.folder = next(name for name in os.listdir(self.dir) if name.startswith("HP-110"))
        os.remove(os.path.join(self.dir, self.folder, "안정성.csv"))
        self.out = tempfile.mkdtemp()
        self.httpd = server_module.serve(self.dir, host="127.0.0.1", port=0, out_dir=self.out,
                                         today=TODAY, log=lambda *args: None)
        self.thread = threading.Thread(target=self.httpd.serve_forever, daemon=True)
        self.thread.start()
        self.base = "http://127.0.0.1:%d" % self.httpd.server_port

    def tearDown(self):
        self.httpd.shutdown()
        self.httpd.server_close()
        shutil.rmtree(self.dir, ignore_errors=True)
        shutil.rmtree(self.out, ignore_errors=True)

    def get(self, path):
        with urllib.request.urlopen(self.base + path, timeout=10) as response:
            return response.status, response.read().decode("utf-8")

    def upload(self, dataset, code, filename, payload):
        body, content_type = multipart({"product": code, "dataset": dataset},
                                       {"file": (filename, payload)})
        request = urllib.request.Request(self.base + "/api/upload", data=body,
                                         headers={"Content-Type": content_type})
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                return response.status, json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            return error.code, json.loads(error.read().decode("utf-8"))

    # ---------------- 기본 ----------------

    def test_serves_dashboard_and_data(self):
        status, body = self.get("/")
        self.assertEqual(status, 200)
        self.assertIn("PQR", body)
        status, body = self.get("/data.js")
        self.assertEqual(status, 200)
        self.assertTrue(body.startswith("window.PQR_DATA = "))

    def test_data_payload_advertises_upload(self):
        _, body = self.get("/api/data")
        payload = json.loads(body)
        self.assertTrue(payload["upload"]["enabled"])
        self.assertIn("g", payload["upload"]["item_datasets"])
        self.assertEqual(payload["upload"]["item_datasets"]["g"], ["stability"])

    def test_path_traversal_is_refused(self):
        """urllib 은 경로를 정리해 버리므로 원문 그대로 보내 확인합니다."""
        import http.client
        connection = http.client.HTTPConnection("127.0.0.1", self.httpd.server_port, timeout=10)
        for path in ("/../../etc/passwd", "/..%2f..%2fetc%2fpasswd", "/../pqr/server.py"):
            connection.putrequest("GET", path, skip_host=False, skip_accept_encoding=True)
            connection.endheaders()
            response = connection.getresponse()
            body = response.read()
            self.assertEqual(response.status, 404, path)
            self.assertNotIn(b"root:", body)
        connection.close()

    # ---------------- 업로드 ----------------

    def test_upload_lands_in_product_folder_and_updates_state(self):
        before = json.loads(self.get("/api/data")[1])
        product = next(p for p in before["products"] if p["code"] == "HP-110")
        self.assertEqual(product["checks"][6], "n")           # (g) 안정성 미착수

        status, result = self.upload("stability", "HP-110", "안정성.csv",
                                     STABILITY_CSV.encode("utf-8"))
        self.assertEqual(status, 200, result)
        self.assertTrue(result["ok"])
        self.assertEqual(result["rows"], 3)
        self.assertTrue(os.path.exists(os.path.join(self.dir, self.folder, "안정성.csv")))

        product = next(p for p in result["data"]["products"] if p["code"] == "HP-110")
        self.assertNotEqual(product["checks"][6], "n")
        self.assertNotIn("stability", product["missing_datasets"])

    def test_upload_without_product_code_column_uses_folder(self):
        _, result = self.upload("stability", "HP-110", "안정성.csv", STABILITY_CSV.encode("utf-8"))
        self.assertEqual(result["rows"], 3)

    def test_mismatched_file_is_refused_and_not_saved(self):
        status, result = self.upload("changes", "HP-110", "안정성.csv",
                                     STABILITY_CSV.encode("utf-8"))
        self.assertEqual(status, 400)
        self.assertIn("안정성 모니터링", result["error"])       # 어느 자료인지 알려 줍니다
        remaining = os.listdir(os.path.join(self.dir, self.folder))
        self.assertNotIn("변경불만대장_안정성.csv", remaining)

    def test_unsupported_extension_is_refused(self):
        status, result = self.upload("stability", "HP-110", "메모.pdf", b"%PDF-1.4")
        self.assertEqual(status, 400)
        self.assertIn("올릴 수 있는 형식", result["error"])

    def test_empty_file_is_refused(self):
        status, result = self.upload("stability", "HP-110", "빈파일.csv", b"")
        self.assertEqual(status, 400)

    def test_replacing_keeps_one_file(self):
        self.upload("stability", "HP-110", "안정성.csv", STABILITY_CSV.encode("utf-8"))
        self.upload("stability", "HP-110", "안정성.csv", STABILITY_CSV.encode("utf-8"))
        files = [name for name in os.listdir(os.path.join(self.dir, self.folder))
                 if "안정성" in name]
        self.assertEqual(files, ["안정성.csv"])
        self.assertFalse([name for name in os.listdir(os.path.join(self.dir, self.folder))
                          if name.endswith(".bak")])

    def test_common_scope_lands_in_common_folder(self):
        payload = ("설비명,구분,최종적격성일,차기예정일,상태\n"
                   "AHU-09,HVAC,2026-01-01,2027-01-01,유효\n").encode("utf-8")
        status, result = self.upload("qualification", "HP-110", "적격성.csv", payload)
        self.assertEqual(status, 200, result)
        self.assertTrue(result["saved"].startswith("공통"), result["saved"])

    def test_unknown_product_code_is_refused(self):
        status, result = self.upload("stability", "../etc", "안정성.csv",
                                     STABILITY_CSV.encode("utf-8"))
        self.assertEqual(status, 400)
        self.assertIn("제품 코드", result["error"])

    def test_export_writes_reports(self):
        request = urllib.request.Request(self.base + "/api/export", data=b"")
        with urllib.request.urlopen(request, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8"))
        self.assertTrue(payload["ok"])
        self.assertTrue(os.path.exists(os.path.join(self.out, "pqr.json")))
        self.assertIn("PQR_요약.md", payload["files"])


class FilenameTest(unittest.TestCase):
    def test_strips_path_and_checks_extension(self):
        self.assertEqual(server_module.safe_filename("../../etc/passwd.csv"), "passwd.csv")
        with self.assertRaises(server_module.UploadError):
            server_module.safe_filename("보고서.exe")
        with self.assertRaises(server_module.UploadError):
            server_module.safe_filename("")

    def test_removes_unsafe_characters(self):
        self.assertEqual(server_module.safe_filename('a<b>c:"d.csv'), "a_b_c__d.csv")


if __name__ == "__main__":
    unittest.main()
