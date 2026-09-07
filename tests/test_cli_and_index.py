import io
import os
import shutil
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

from gmpai.catalog import load_catalog
from gmpai.cli import main
from gmpai.index import render_html, render_markdown, write_indexes


class IndexTests(unittest.TestCase):
    def setUp(self):
        self.catalog = load_catalog()
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_markdown_lists_every_document(self):
        md = render_markdown(self.catalog)
        for doc in self.catalog:
            self.assertIn(doc.title_ko, md)
        self.assertIn("| 문서 | 분류 | 상태 |", md)

    def test_html_is_self_contained_and_escaped(self):
        page = render_html(self.catalog)
        self.assertTrue(page.startswith("<!doctype html>"))
        self.assertNotIn("<script", page.lower())
        self.assertNotIn("http://", page.replace("http://www.w3.org", ""))
        self.assertIn("prefers-color-scheme", page)
        for doc in self.catalog:
            self.assertIn(doc.landing_page, page)

    def test_write_indexes_creates_both_files(self):
        md = self.tmp / "INDEX.md"
        html = self.tmp / "docs" / "index.html"
        written = write_indexes(self.catalog, markdown_path=md, html_path=html)
        self.assertEqual(written, [md, html])
        self.assertTrue(md.stat().st_size > 0)
        self.assertTrue(html.stat().st_size > 0)


class CliTests(unittest.TestCase):
    def run_cli(self, argv):
        buffer = io.StringIO()
        with redirect_stdout(buffer):
            code = main(argv)
        return code, buffer.getvalue()

    def test_list_prints_documents(self):
        code, out = self.run_cli(["list"])
        self.assertEqual(code, 0)
        self.assertIn("eu-gmp-annex22-ai", out)
        self.assertIn("fda-ai-regulatory-decision-making", out)

    def test_list_filters_by_authority(self):
        code, out = self.run_cli(["list", "--authority", "FDA"])
        self.assertEqual(code, 0)
        self.assertNotIn("eu-gmp-annex22-ai", out)
        self.assertIn("fda-ai-drug-manufacturing", out)

    def test_list_with_unknown_id_exits_with_error(self):
        code, _ = self.run_cli(["list", "--id", "nope"])
        self.assertEqual(code, 2)

    def test_index_command_writes_files(self):
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp, True)
        code, out = self.run_cli(
            ["index", "--markdown", str(tmp / "INDEX.md"), "--html", str(tmp / "index.html")]
        )
        self.assertEqual(code, 0)
        self.assertIn("생성:", out)
        self.assertTrue((tmp / "INDEX.md").exists())

    def test_status_without_downloads_reports_missing(self):
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp, True)
        code, out = self.run_cli(["status", "-o", str(tmp)])
        self.assertEqual(code, 1)
        self.assertIn("download", out)


if __name__ == "__main__":
    unittest.main()


class 작성_주의사항_화면(unittest.TestCase):
    """담당자 2026-09-07: "오늘 수정 요청한 모든 PQR 관련 내용도 PQR 작성 주의사항에 반영해야 돼."
    지시는 지시 대장 한 곳에만 적고, 대시보드가 그것을 읽어 보여 준다."""

    def test_지시_대장을_장별로_읽어_준다(self):
        from pqr.server import Handler
        got = Handler._handle_notes(Handler)
        self.assertTrue(got["ok"])
        self.assertGreater(got["count"], 50)
        titles = [g["title"] for g in got["groups"]]
        self.assertIn("0. 늘 지키는 것", titles)
        rules = [r["rule"] for g in got["groups"] for r in g["rules"]]
        self.assertTrue(any("금속성이물" in r for r in rules))      # 오늘 지시가 들어 있다
        self.assertTrue(any("안정성 판독" in r for r in rules))
        self.assertFalse(any(set(r) <= set("-: ") for r in rules))  # 표 구분선은 섞이지 않는다

    def test_화면이_그_목록을_그린다(self):
        here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        text = io.open(os.path.join(here, "docs", "pqr", "index.html"), encoding="utf-8").read()
        for mark in ("api/notes", "note-body", "note-find", "renderNotes"):
            self.assertIn(mark, text)
