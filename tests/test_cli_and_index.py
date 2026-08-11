import io
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
