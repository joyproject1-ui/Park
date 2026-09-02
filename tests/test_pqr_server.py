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


class PriorReportTest(ServerTest):
    """보고서 작성 — 전년도 PQR 이 있으면 그 서식으로 새 연도 보고서를 만듭니다."""

    def make_previous(self, year=2025):
        """전년도 PQR 을 흉내낸 워드 파일을 제품 폴더에 둡니다."""
        import zipfile
        from pqr import prior_report
        folder = os.path.join(self.dir, self.folder)
        path = os.path.join(folder, "0. 전년도 PQR 히알루론점안액.docx")
        document = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body>'
            '<w:p><w:r><w:t>제품품질평가는 %d 년 1월 ~ 12월까지 생산된 제품을 평가한다.</w:t></w:r></w:p>'
            '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>%d.02.13</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
            '</w:body></w:document>' % (year, year))
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("word/document.xml", document)
            archive.writestr("[Content_Types].xml", "<Types/>")
        return path, folder

    def report(self, code="HP-110"):
        request = urllib.request.Request(
            self.base + "/api/report",
            data=json.dumps({"product": code, "force": True}).encode("utf-8"),
            headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            return json.loads(error.read().decode("utf-8"))

    def test_uses_previous_pqr_as_the_base(self):
        previous, folder = self.make_previous(2024)
        result = self.report()
        self.assertTrue(result["ok"], result.get("error"))
        self.assertIsNotNone(result["based_on"], "전년도 PQR 을 기준 본으로 써야 합니다")
        self.assertEqual(result["based_on"]["previous"], os.path.basename(previous))
        self.assertEqual(result["based_on"]["previous_year"], 2024)

    def test_body_years_move_but_table_values_stay(self):
        """평가 기간이 2025년이면 2024년 기준 본의 본문 연도는 2025로 옮깁니다."""
        import zipfile
        self.make_previous(2024)
        result = self.report()
        with zipfile.ZipFile(result["final"]) as archive:
            body = archive.read("word/document.xml").decode("utf-8")
        self.assertIn("2025 년 1월", body)      # 본문 연도는 새 평가 연도로
        self.assertIn("2024.02.13", body)       # 표 안 날짜는 그대로 — 담당자가 채웁니다

    def test_zip_base_is_unpacked_and_attachments_come_along(self):
        """전년도 PQR 을 압축으로 올려도 안의 워드를 기준 본으로 씁니다."""
        import io, zipfile
        folder = os.path.join(self.dir, self.folder)
        document = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p><w:r><w:t>제품품질평가는 2024 년 1월 ~ 12월까지 생산된 제품이다.</w:t>'
            '</w:r></w:p></w:body></w:document>')
        inner = io.BytesIO()
        with zipfile.ZipFile(inner, "w") as docx:
            docx.writestr("word/document.xml", document)
            docx.writestr("[Content_Types].xml", "<Types/>")
        bundle = os.path.join(folder, "0. 전년도 PQR 자료.zip")
        with zipfile.ZipFile(bundle, "w") as archive:
            archive.writestr("PQR24 히알루론점안액.docx", inner.getvalue())
            archive.writestr("HLF-QC-126-06 안정성 경향 2024.xlsx", b"excel")
        result = self.report()
        self.assertTrue(result["ok"], result.get("error"))
        self.assertIn(".docx", result["based_on"]["previous"])
        with zipfile.ZipFile(result["final"]) as archive:
            body = archive.read("word/document.xml").decode("utf-8")
        self.assertIn("2025 년 1월", body)
        # 첨부 엑셀도 새 연도 이름으로 옆에 놓입니다.
        self.assertIn("HLF-QC-126-06 안정성 경향 2025.xlsx", result["based_on"]["attachments"])
        self.assertTrue(os.path.isfile(
            os.path.join(folder, "HLF-QC-126-06 안정성 경향 2025.xlsx")))

    def test_without_previous_pqr_it_still_makes_a_summary(self):
        result = self.report()
        self.assertTrue(result["ok"], result.get("error"))
        self.assertIsNone(result["based_on"])


class ItemFileListTest(ItemUploadTest):
    """올리기 창에서 그 항목의 첨부를 보고, 잘못 올린 것은 지웁니다."""

    def call(self, path, body):
        request = urllib.request.Request(
            self.base + path, data=json.dumps(body).encode("utf-8"),
            headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(request, timeout=20) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            return json.loads(error.read().decode("utf-8"))

    def test_list_shows_only_that_items_files(self):
        self.upload_item("HP-110", "13", "안정성 결과표.xlsx")
        self.upload_item("HP-110", "12", "변경관리 대장.xlsx")
        listing = self.call("/api/item-files", {"product": "HP-110", "item": "13"})
        self.assertTrue(listing["ok"], listing.get("error"))
        names = [row["name"] for row in listing["files"]]
        self.assertTrue(all(name.startswith("13") for name in names), names)

    def test_delete_removes_the_file_and_updates_the_screen(self):
        saved = self.upload_item("HP-110", "13", "안정성 결과표.xlsx")
        name = os.path.basename(saved["saved"])
        result = self.call("/api/item-delete",
                           {"product": "HP-110", "item": "13", "name": name})
        self.assertTrue(result["ok"], result.get("error"))
        self.assertFalse(os.path.isfile(os.path.join(self.dir, self.folder, name)))
        product = next(p for p in result["data"]["products"] if p["code"] == "HP-110")
        ids = [item[0] for item in result["data"]["items"]]
        self.assertEqual(dict(zip(ids, product["checks"]))["13"], "n")

    def test_delete_refuses_a_file_from_another_item(self):
        saved = self.upload_item("HP-110", "12", "변경관리 대장.xlsx")
        name = os.path.basename(saved["saved"])
        result = self.call("/api/item-delete",
                           {"product": "HP-110", "item": "13", "name": name})
        self.assertFalse(result.get("ok"))
        self.assertTrue(os.path.isfile(os.path.join(self.dir, self.folder, name)))

    def test_delete_refuses_a_path_outside_the_folder(self):
        result = self.call("/api/item-delete",
                           {"product": "HP-110", "item": "13", "name": "../config.json"})
        self.assertFalse(result.get("ok"))


class ReferenceDocTest(ServerTest):
    """'PQR 작성 시 참고 사항' — 올린 문서를 목록에서 눌러 읽습니다."""

    def upload_reference(self, filename, payload=b"x" * 30):
        body, content_type = multipart({"reference": "1"}, {"file": (filename, payload)})
        request = urllib.request.Request(self.base + "/api/upload", data=body,
                                         headers={"Content-Type": content_type})
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            return json.loads(error.read().decode("utf-8"))

    def test_uploaded_document_shows_up_in_the_list(self):
        name = "PQR 작성방법.docx"
        result = self.upload_reference(name)
        self.assertTrue(result["ok"], result.get("error"))
        files = [row["name"] for row in result["data"]["reference"]["files"]]
        self.assertIn(name, files)
        # 제품 폴더가 아니라 참고 폴더에 들어갑니다.
        folder = os.path.join(self.dir, "PQR 작성 시 참고 사항")
        self.assertTrue(os.path.isfile(os.path.join(folder, name)))

    def test_open_refuses_a_path_outside_the_folder(self):
        request = urllib.request.Request(
            self.base + "/api/reference-open",
            data=json.dumps({"name": "../config.json"}).encode("utf-8"),
            headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(request, timeout=10) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            payload = json.loads(error.read().decode("utf-8"))
        self.assertFalse(payload.get("ok"))

    def test_empty_name_opens_the_folder(self):
        """이름 없이 부르면 참고 폴더를 엽니다 — 탐색기로 파일을 직접 넣을 수 있게."""
        request = urllib.request.Request(
            self.base + "/api/reference-open", data=b"{}",
            headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(request, timeout=10) as response:
            payload = json.loads(response.read().decode("utf-8"))
        self.assertTrue(payload["ok"])
        self.assertTrue(payload["folder"])
        self.assertTrue(os.path.isdir(payload["path"]))

    def test_reference_files_are_not_counted_as_product_data(self):
        """참고 폴더는 제품 폴더가 아니므로 수집 현황에 섞이지 않습니다."""
        result = self.upload_reference("0. 전년도 PQR 참고.docx")
        codes = [p["code"] for p in result["data"]["products"]]
        self.assertNotIn("PQR", codes)


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

    def test_authored_report_wins_over_generated_draft(self):
        """프로그램이 만든 초안이 있어도 완성본 단추는 담당자 제출본을 엽니다."""
        from pqr import build as build_module
        folder = os.path.join(self.dir, self.folder)   # setUp 이 만든 제품 폴더
        draft = os.path.join(folder, "[HP-110] 히알루론점안액 2025년 제품품질평가 (제출용).docx")
        with open(draft, "wb") as handle:
            handle.write(b"draft")
        build_module.mark_auto_draft(folder, draft)
        authored = os.path.join(folder, "[HP-110] 히알루론점안액 2025년 제품품질평가 (제출용) v4.docx")
        with open(authored, "wb") as handle:
            handle.write(b"authored")
        os.utime(authored, (1, 1))              # 초안보다 오래된 파일이어도
        found = build_module.find_final_report(folder)
        self.assertEqual(os.path.basename(found), os.path.basename(authored))

    def test_generated_draft_is_recognized_by_its_own_wording(self):
        """표시 파일이 없어도(예전에 만든 초안) 문서 안 문구로 초안임을 알아봅니다."""
        import os
        from pqr import build as build_module, docx_report
        folder = os.path.join(self.dir, self.folder)
        data = build_module.build(input_dir=self.dir, today=TODAY)
        draft = docx_report.write_docx(data, "HP-110", folder)
        os.remove(os.path.join(folder, build_module.AUTO_DRAFT_MARKER))   # 표시 파일 삭제
        self.assertTrue(build_module._looks_like_auto_draft(draft))
        authored = os.path.join(folder, "[HP-110] 히알루론점안액 2025년 제품품질평가 (제출용) v4.docx")
        with open(draft, "rb") as source:
            payload = source.read()
        import zipfile, io, re
        buffer = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(payload)) as src, zipfile.ZipFile(buffer, "w") as dst:
            for info in src.infolist():                    # 초안 문구만 지운 사본 = 담당자 작성본
                blob = src.read(info.filename)
                if info.filename == "word/document.xml":
                    blob = blob.replace(build_module.AUTO_DRAFT_SIGNATURE.encode("utf-8"), b"")
                dst.writestr(info, blob)
        with open(authored, "wb") as handle:
            handle.write(buffer.getvalue())
        os.utime(authored, (1, 1))                         # 초안보다 오래된 파일이어도
        found = build_module.find_final_report(folder)
        self.assertEqual(os.path.basename(found), os.path.basename(authored))

    def test_submitted_report_becomes_the_final_report(self):
        """작성해 준 제출본을 한번에 올리기로 넣으면 완성본 버튼이 그 파일을 엽니다."""
        name = "[HP-110] 히알루론점안액 2025년 제품품질평가 (제출용).docx"
        result = self.upload_bulk("HP-110", name)
        self.assertTrue(result["ok"], result.get("error"))
        product = next(p for p in result["data"]["products"] if p["code"] == "HP-110")
        self.assertEqual(product["final_report"], name)
