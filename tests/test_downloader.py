"""End-to-end tests against a local stand-in for the regulator websites."""

import json
import shutil
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from gmpai.catalog import load_catalog
from gmpai.downloader import (
    FAILED,
    OK,
    SKIPPED,
    UNCHANGED,
    UPDATED,
    download_all,
    download_document,
    load_manifest,
    sha256_bytes,
)
from gmpai.fetcher import FetchError, fetch, find_matching_links

PDF_A = b"%PDF-1.7\n% annex 22 draft\n%%EOF\n"
PDF_B = b"%PDF-1.7\n% annex 22 revised\n%%EOF\n"

LANDING_HTML = """
<html><body>
  <a href="/other/newsletter.pdf">Newsletter</a>
  <a href="/document/download/abc_en?filename=mp_vol4_chap4_annex22_consultation_guideline_en.pdf">Annex 22 draft</a>
  <a href="/media/184830/download">Download the guidance PDF</a>
</body></html>
"""

# Mutated by tests to change what the "regulator" serves.
STATE = {"annex22": PDF_A, "direct_ok": True}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):  # silence test output
        pass

    def _send(self, code, content_type, body):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = self.path.split("?")[0]
        if path == "/landing":
            self._send(200, "text/html; charset=utf-8", LANDING_HTML.encode())
        elif path == "/direct.pdf":
            if STATE["direct_ok"]:
                self._send(200, "application/pdf", STATE["annex22"])
            else:
                self._send(404, "text/plain", b"gone")
        elif path == "/document/download/abc_en":
            self._send(200, "application/pdf", STATE["annex22"])
        elif path == "/media/184830/download":
            self._send(200, "application/octet-stream", PDF_B)
        elif path == "/other/newsletter.pdf":
            self._send(200, "application/pdf", b"%PDF-1.4\nnewsletter\n")
        elif path == "/consent.pdf":
            self._send(200, "text/html", b"<!doctype html><html><body>cookie wall</body></html>")
        elif path == "/flaky.pdf":
            STATE["flaky_hits"] = STATE.get("flaky_hits", 0) + 1
            if STATE["flaky_hits"] < 2:
                self._send(503, "text/plain", b"try later")
            else:
                self._send(200, "application/pdf", PDF_A)
        else:
            self._send(404, "text/plain", b"not found")


class ServerTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        cls.base = f"http://127.0.0.1:{cls.server.server_port}"
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()

    def setUp(self):
        STATE.update({"annex22": PDF_A, "direct_ok": True, "flaky_hits": 0})
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def write_catalog(self, documents):
        path = self.tmp / "catalog.json"
        path.write_text(
            json.dumps({"schema_version": 1, "last_reviewed": "2026-08-11", "documents": documents}),
            encoding="utf-8",
        )
        return load_catalog(path)

    def base_doc(self, **overrides):
        doc = {
            "id": "annex22",
            "title": "Annex 22",
            "title_ko": "부속서 22",
            "authority": "EU",
            "category": "gmp-ai",
            "status": "draft",
            "document_date": "2025-07-07",
            "landing_page": f"{self.base}/landing",
            "direct_url": f"{self.base}/direct.pdf",
            "discover": {"pattern": "annex22"},
            "filename": "annex22.pdf",
        }
        doc.update(overrides)
        return doc


class DownloadTests(ServerTestCase):
    def test_downloads_pdf_and_records_manifest(self):
        catalog = self.write_catalog([self.base_doc()])
        out = self.tmp / "downloads"
        results = download_all(list(catalog), out)

        self.assertEqual([r.status for r in results], [OK])
        saved = out / "EU" / "annex22.pdf"
        self.assertTrue(saved.exists())
        self.assertEqual(saved.read_bytes(), PDF_A)

        manifest = load_manifest(out)
        entry = manifest["documents"]["annex22"]
        self.assertEqual(entry["sha256"], sha256_bytes(PDF_A))
        self.assertEqual(entry["bytes"], len(PDF_A))
        self.assertTrue(entry["downloaded_at"].endswith("Z"))
        self.assertFalse(list(out.rglob("*.part")), "temporary files must not be left behind")

    def test_second_run_skips_already_downloaded_file(self):
        catalog = self.write_catalog([self.base_doc()])
        out = self.tmp / "downloads"
        download_all(list(catalog), out)
        again = download_all(list(catalog), out)
        self.assertEqual(again[0].status, SKIPPED)

    def test_force_detects_a_changed_regulation(self):
        catalog = self.write_catalog([self.base_doc()])
        out = self.tmp / "downloads"
        download_all(list(catalog), out)

        STATE["annex22"] = PDF_B
        results = download_all(list(catalog), out, force=True)
        self.assertEqual(results[0].status, UPDATED)
        self.assertEqual(results[0].previous_sha256, sha256_bytes(PDF_A))
        self.assertEqual(results[0].sha256, sha256_bytes(PDF_B))
        self.assertEqual((out / "EU" / "annex22.pdf").read_bytes(), PDF_B)

    def test_force_reports_unchanged_when_content_is_identical(self):
        catalog = self.write_catalog([self.base_doc()])
        out = self.tmp / "downloads"
        download_all(list(catalog), out)
        results = download_all(list(catalog), out, force=True)
        self.assertEqual(results[0].status, UNCHANGED)

    def test_falls_back_to_landing_page_when_direct_url_is_dead(self):
        STATE["direct_ok"] = False
        catalog = self.write_catalog([self.base_doc()])
        results = download_all(list(catalog), self.tmp / "downloads", retries=1)
        self.assertEqual(results[0].status, OK)
        self.assertIn("/document/download/abc_en", results[0].resolved_url)

    def test_document_without_direct_url_is_resolved_by_discovery(self):
        doc = self.base_doc(id="fda-doc", direct_url=None, discover={"pattern": r"/media/\d+/download"})
        catalog = self.write_catalog([doc])
        results = download_all(list(catalog), self.tmp / "downloads")
        self.assertEqual(results[0].status, OK)
        self.assertEqual(results[0].sha256, sha256_bytes(PDF_B))

    def test_html_response_is_rejected_rather_than_saved_as_pdf(self):
        doc = self.base_doc(direct_url=f"{self.base}/consent.pdf", discover=None)
        catalog = self.write_catalog([doc])
        out = self.tmp / "downloads"
        results = download_all(list(catalog), out, retries=1)
        self.assertEqual(results[0].status, FAILED)
        self.assertIn("HTML", results[0].error)
        self.assertFalse((out / "EU" / "annex22.pdf").exists())

    def test_failure_is_recorded_without_erasing_a_previous_download(self):
        catalog = self.write_catalog([self.base_doc()])
        out = self.tmp / "downloads"
        download_all(list(catalog), out)

        broken = self.write_catalog([self.base_doc(direct_url=f"{self.base}/missing.pdf", discover=None)])
        download_all(list(broken), out, force=True, retries=1)

        entry = load_manifest(out)["documents"]["annex22"]
        self.assertEqual(entry["sha256"], sha256_bytes(PDF_A))
        self.assertIn("last_error", entry)

    def test_dry_run_writes_nothing(self):
        catalog = self.write_catalog([self.base_doc()])
        out = self.tmp / "downloads"
        results = download_all(list(catalog), out, dry_run=True)
        self.assertEqual(results[0].status, OK)
        self.assertFalse(out.exists())

    def test_prefer_discovery_reverses_url_order(self):
        catalog = self.write_catalog([self.base_doc(discover={"pattern": r"/media/\d+/download"})])
        results = download_all(list(catalog), self.tmp / "downloads", prefer_discovery=True)
        self.assertIn("/media/184830/download", results[0].resolved_url)

    def test_transient_server_error_is_retried(self):
        doc = self.base_doc(direct_url=f"{self.base}/flaky.pdf", discover=None)
        catalog = self.write_catalog([doc])
        result = download_document(
            catalog.get("annex22"), self.tmp / "downloads", retries=3, timeout=5
        )
        self.assertEqual(result.status, OK)
        self.assertGreaterEqual(STATE["flaky_hits"], 2)


class FetcherTests(ServerTestCase):
    def test_404_is_not_retried_forever(self):
        calls = []

        def sleep(seconds):
            calls.append(seconds)

        with self.assertRaises(FetchError):
            fetch(f"{self.base}/missing.pdf", retries=3, sleep=sleep)
        self.assertEqual(calls, [], "client errors must not trigger backoff")

    def test_find_matching_links_matches_href_and_text(self):
        html = '<a href="/x/1.pdf">Annex 22 draft</a><a href="/annex22.pdf">Other</a>'
        by_text = find_matching_links(html, "https://example.org/p", "(?i)annex 22")
        self.assertEqual(by_text, ["https://example.org/x/1.pdf"])
        by_href = find_matching_links(html, "https://example.org/p", "annex22")
        self.assertEqual(by_href, ["https://example.org/annex22.pdf"])

    def test_relative_links_are_absolutised(self):
        links = find_matching_links(LANDING_HTML, f"{self.base}/landing", r"/media/\d+/download")
        self.assertEqual(links, [f"{self.base}/media/184830/download"])


if __name__ == "__main__":
    unittest.main()
