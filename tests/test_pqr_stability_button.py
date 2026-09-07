"""'안정성 판독' 단추 — 13 폴더를 읽어 판독 파일(json)을 따로 만든다 (담당자 2026-09-07)."""
import json
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import handwriting, stability_read                     # noqa: E402


class 판독_단추(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.mkdtemp(prefix="pqr-read-")
        self.thirteen = os.path.join(self.folder, "13. 안정성 시험")
        os.makedirs(self.thirteen)

    def _scan(self, name):
        with open(os.path.join(self.thirteen, name), "wb") as handle:
            handle.write(b"%PDF-1.4 scanned")     # 글자 없는 스캔본으로 읽힌다

    def test_시험일지가_없으면_까닭을_알려_준다(self):
        got = stability_read.make_reading(self.folder)
        self.assertFalse(got["ok"])
        self.assertIn("13 폴더", got["why"])

    def test_읽은_결과를_covers_all_로_저장한다(self):
        self._scan("퀴노비드안연고 장기 안정성시험일지 OEX101.pdf")
        logs = [{"lot": "OEX101", "year": "2024", "kind": "장기", "market": "내수",
                 "points": [{"period": "12M", "done": "2025.03.10", "assays": {"오플록사신": 101.2},
                             "unsure": ["done"]}]}]
        old_how, old_read = stability_read.how, handwriting.read_folder
        stability_read.how = lambda folder=None: ("pc", "시험용 판독기")
        handwriting.read_folder = lambda paths, specs=None, log=None: logs
        try:
            got = stability_read.make_reading(self.folder, "QC1-7007")
        finally:
            stability_read.how, handwriting.read_folder = old_how, old_read
        self.assertEqual((got["ok"], got["lots"], got["unsure"], got["pages"]), (True, 1, 1, 1))
        with open(os.path.join(self.folder, handwriting.CACHE_NAME), encoding="utf-8") as handle:
            saved = json.load(handle)
        self.assertTrue(saved["covers_all"])        # '보고서 작성' 이 시험일지를 다시 읽지 않게
        self.assertEqual(saved["logs"][0]["lot"], "OEX101")

    def test_담당자가_고쳐_둔_판독_파일_값이_먼저다(self):
        self._scan("퀴노비드안연고 장기 안정성시험일지 OEX101.pdf")
        handwriting.save_cache(self.folder, [{"lot": "OEX101", "year": "2024", "kind": "장기",
                                              "points": [{"period": "12M", "done": "2025.01.02",
                                                          "assays": {"오플록사신": 99.9}, "unsure": []}]}])
        old_how, old_read = stability_read.how, handwriting.read_folder
        stability_read.how = lambda folder=None: ("pc", "시험용 판독기")
        handwriting.read_folder = lambda paths, specs=None, log=None: [
            {"lot": "OEX101", "year": "2024", "kind": "장기",
             "points": [{"period": "12M", "done": "2025.09.09", "assays": {"오플록사신": 50.0}, "unsure": []}]}]
        try:
            stability_read.make_reading(self.folder)
        finally:
            stability_read.how, handwriting.read_folder = old_how, old_read
        got = handwriting.load_cache(self.folder)
        self.assertEqual(got[0]["points"][0]["done"], "2025.01.02")
        self.assertEqual(got[0]["points"][0]["assays"]["오플록사신"], 99.9)


if __name__ == "__main__":
    unittest.main()
