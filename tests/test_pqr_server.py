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
from pqr.build import load_config

ITEM_IDS = [item[0] for item in load_config()["items"]]
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
        self.assertIn("13", payload["upload"]["item_datasets"])
        self.assertEqual(payload["upload"]["item_datasets"]["13"], ["stability"])

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
        self.assertEqual(product["checks"][ITEM_IDS.index("13")], "n")   # 13 안정성 미착수

        status, result = self.upload("stability", "HP-110", "안정성.csv",
                                     STABILITY_CSV.encode("utf-8"))
        self.assertEqual(status, 200, result)
        self.assertTrue(result["ok"])
        self.assertEqual(result["rows"], 3)
        self.assertTrue(os.path.exists(os.path.join(self.dir, self.folder, "안정성.csv")))

        product = next(p for p in result["data"]["products"] if p["code"] == "HP-110")
        self.assertNotEqual(product["checks"][ITEM_IDS.index("13")], "n")
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


    # ---------------- 폴더 열기 · 보고서 작성 ----------------

    def post_json(self, path, payload):
        request = urllib.request.Request(self.base + path, data=json.dumps(payload).encode(),
                                         headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(request, timeout=30) as response:
            return response.status, json.loads(response.read().decode("utf-8"))

    def test_open_creates_product_folder(self):
        status, payload = self.post_json("/api/open", {"product": "HP-201"})
        self.assertEqual(status, 200)
        self.assertTrue(payload["ok"])
        self.assertTrue(os.path.isdir(payload["path"]))
        self.assertIn("HP-201", os.path.basename(payload["path"]))

    def test_report_refuses_until_collection_complete(self):
        status, payload = self.post_json("/api/report", {"product": "HP-110"})
        self.assertEqual(status, 200)
        self.assertFalse(payload["ok"])          # 안정성 파일을 지워 두어 100% 가 아닙니다
        self.assertIn("%", payload["error"])

    def test_report_writes_files_when_forced(self):
        status, payload = self.post_json("/api/report", {"product": "HP-110", "force": True})
        self.assertTrue(payload["ok"])
        for path in payload["files"]:
            self.assertTrue(os.path.exists(path))

    def test_report_writes_submission_docx_and_shows_final_button(self):
        """보고서 작성 → 제출용 .docx 가 제품 폴더에 생기고, 그 자리에서 완성본 단추가 섭니다."""
        status, payload = self.post_json("/api/report", {"product": "HP-110", "force": True})
        self.assertTrue(payload["ok"])
        self.assertTrue(os.path.isfile(payload["final"]), payload.get("final"))
        self.assertIn("제출용", payload["final_name"])
        # 제품 폴더에 두어야 근거 자료 옆에서 검토할 수 있습니다.
        self.assertEqual(os.path.dirname(payload["final"]), payload["folder"])
        product = next(p for p in payload["data"]["products"] if p["code"] == "HP-110")
        self.assertEqual(product["final_report"], payload["final_name"])

    def test_final_endpoint_opens_the_submission_report(self):
        self.post_json("/api/report", {"product": "HP-110", "force": True})
        status, payload = self.post_json("/api/final", {"product": "HP-110"})
        self.assertEqual(status, 200)
        self.assertTrue(payload["ok"], payload.get("error"))
        self.assertTrue(os.path.isfile(payload["path"]))

    def test_final_endpoint_reports_when_nothing_is_written_yet(self):
        status, payload = self.post_json("/api/final", {"product": "HP-201"})
        self.assertFalse(payload["ok"])
        self.assertIn("완성본", payload["error"])

    def test_close_as_none_creates_record_and_marks_item(self):
        status, payload = self.post_json("/api/close", {"product": "HP-110", "item": "13"})
        self.assertEqual(status, 200)
        self.assertTrue(payload["ok"])
        record = os.path.join(self.dir, payload["saved"])
        self.assertTrue(os.path.exists(record))
        with open(record, encoding="utf-8") as handle:
            body = handle.read()
        self.assertIn("해당 이력 없음", body)
        product = next(p for p in payload["data"]["products"] if p["code"] == "HP-110")
        ids = [item[0] for item in payload["data"]["items"]]
        self.assertEqual(dict(zip(ids, product["checks"]))["13"], "y")




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


class ItemUploadTest(ServerTest):
    """평가항목 자료는 제품 폴더에 항 번호 이름으로 저장됩니다.

    표준 대장으로 읽으려 들면 '원료 공급업체 List' 가 '설비 적격성' 으로 저장되는
    일이 생깁니다 — 실제로 그렇게 보였습니다.
    """

    def upload_item(self, product, item, filename, payload=b"x" * 40):
        import uuid
        boundary = uuid.uuid4().hex
        parts = []
        for name, value in (("product", product), ("item", item)):
            parts.append(("--%s\r\nContent-Disposition: form-data; name=\"%s\"\r\n\r\n%s\r\n"
                          % (boundary, name, value)).encode())
        parts.append(("--%s\r\nContent-Disposition: form-data; name=\"file\"; "
                      "filename=\"%s\"\r\n\r\n" % (boundary, filename)).encode())
        parts.append(payload + b"\r\n" + ("--%s--\r\n" % boundary).encode())
        request = urllib.request.Request(
            self.base + "/api/upload", data=b"".join(parts),
            headers={"Content-Type": "multipart/form-data; boundary=" + boundary})
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            return json.loads(error.read().decode("utf-8"))

    def test_saved_into_the_product_folder_with_the_item_number(self):
        result = self.upload_item("HP-110", "13", "안정성 결과표.xlsx")
        self.assertTrue(result["ok"], result.get("error"))
        self.assertIn("HP-110", result["saved"])
        self.assertTrue(os.path.basename(result["saved"]).startswith("13 "),
                        result["saved"])
        product = next(p for p in result["data"]["products"] if p["code"] == "HP-110")
        ids = [item[0] for item in result["data"]["items"]]
        self.assertEqual(dict(zip(ids, product["checks"]))["13"], "y")

    def test_company_original_keeps_its_own_name(self):
        """이미 항 번호로 시작하면 번호를 또 붙이지 않습니다."""
        name = "8.1.1 원료 공급업체 List_(Rev.26).xlsx"
        result = self.upload_item("HP-110", "8.1.1", name)
        self.assertTrue(result["ok"], result.get("error"))
        self.assertEqual(os.path.basename(result["saved"]), name)

    def test_pdf_and_legacy_xls_are_accepted(self):
        """근거 자료는 성적서 PDF 와 ERP 의 구형 .xls 로 옵니다."""
        for filename in ("6. 제조내역 - ERP.pdf", "RSN101 주원료.xls", "안정성.docx"):
            result = self.upload_item("HP-110", "6", filename)
            self.assertTrue(result["ok"], "%s: %s" % (filename, result.get("error")))

    def test_executable_is_refused(self):
        result = self.upload_item("HP-110", "6", "악성.exe")
        self.assertFalse(result.get("ok"))
        self.assertIn("형식", result["error"])

    def test_unknown_item_is_refused(self):
        result = self.upload_item("HP-110", "9.9", "무엇.pdf")
        self.assertFalse(result.get("ok"))
        self.assertIn("평가항목", result["error"])


class BulkUploadTest(ItemUploadTest):
    """'파일 한번에 올리기' — 원본 이름 그대로 제품 폴더에 저장하고 항은 이름으로 인식."""

    def upload_bulk(self, product, filename, last="1", payload=b"x" * 40):
        import uuid
        boundary = uuid.uuid4().hex
        parts = []
        for name, value in (("product", product), ("bulk", "1"), ("last", last)):
            parts.append(("--%s\r\nContent-Disposition: form-data; name=\"%s\"\r\n\r\n%s\r\n"
                          % (boundary, name, value)).encode())
        parts.append(("--%s\r\nContent-Disposition: form-data; name=\"file\"; "
                      "filename=\"%s\"\r\n\r\n" % (boundary, filename)).encode())
        parts.append(payload + b"\r\n" + ("--%s--\r\n" % boundary).encode())
        request = urllib.request.Request(
            self.base + "/api/upload", data=b"".join(parts),
            headers={"Content-Type": "multipart/form-data; boundary=" + boundary})
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            return json.loads(error.read().decode("utf-8"))

    def test_keeps_original_name_and_detects_item(self):
        name = "7. 수율현황표 - 개인.xlsx"
        result = self.upload_bulk("HP-110", name)
        self.assertTrue(result["ok"], result.get("error"))
        self.assertEqual(os.path.basename(result["saved"]), name)
        self.assertEqual(result["item"], "7")
        product = next(p for p in result["data"]["products"] if p["code"] == "HP-110")
        ids = [item[0] for item in result["data"]["items"]]
        self.assertEqual(dict(zip(ids, product["checks"]))["7"], "y")

    def test_unrecognized_name_is_still_saved(self):
        result = self.upload_bulk("HP-110", "메모.pdf")
        self.assertTrue(result["ok"], result.get("error"))
        self.assertEqual(result["item"], "")

    def test_executable_refused_in_bulk_too(self):
        result = self.upload_bulk("HP-110", "악성.exe")
        self.assertFalse(result.get("ok"))
        self.assertIn("형식", result["error"])

    def test_submitted_report_becomes_the_final_report(self):
        """작성해 준 제출본을 한번에 올리기로 넣으면 완성본 버튼이 그 파일을 엽니다."""
        name = "[HP-110] 히알루론점안액 2025년 제품품질평가 (제출용).docx"
        result = self.upload_bulk("HP-110", name)
        self.assertTrue(result["ok"], result.get("error"))
        product = next(p for p in result["data"]["products"] if p["code"] == "HP-110")
        self.assertEqual(product["final_report"], name)
