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
            got = stability_read.make_reading(self.folder, "QC1-7007", allow_pc=True)
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
            stability_read.make_reading(self.folder, allow_pc=True)
        finally:
            stability_read.how, handwriting.read_folder = old_how, old_read
        got = handwriting.load_cache(self.folder)
        self.assertEqual(got[0]["points"][0]["done"], "2025.01.02")
        self.assertEqual(got[0]["points"][0]["assays"]["오플록사신"], 99.9)


    def test_이_PC_판독기밖에_없으면_묻고_묶음을_만들어_둔다(self):
        # 담당자 2026-09-07: "시간이 너무 걸려서 PC 로 안 읽기로 한 거 아냐?"
        self._scan("퀴노비드안연고 장기 안정성시험일지 OEX101.pdf")
        old_how = stability_read.how
        stability_read.how = lambda folder=None: ("pc", "이 PC 의 판독기")
        try:
            got = stability_read.make_reading(self.folder, "퀴노비드안연고")
        finally:
            stability_read.how = old_how
        self.assertEqual((got["ok"], got["need"], got["pages"]), (False, "pc", 1))
        self.assertEqual(got["minutes"], stability_read.MINUTES_PER_PAGE)
        self.assertTrue(got["zip"] and got["zip"][0].startswith("13. Claude 판독 요청"))
        self.assertTrue(os.path.isfile(os.path.join(self.folder, got["zip"][0])))


if __name__ == "__main__":
    unittest.main()


class 첨부_점검(unittest.TestCase):
    """담당자 2026-09-07: "안정성 Lot No 가 워드 파일과 엑셀 파일이 상이해 — 다시는 실수하지 않게"."""

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="pqr-attach-")

    def _touch(self, name, when=None):
        path = os.path.join(self.dir, name)
        with open(path, "wb") as handle:
            handle.write(b"x")
        if when is not None:
            os.utime(path, (when, when))
        return path

    def test_새로_쓰이지_않은_첨부와_남은_예전_첨부를_알린다(self):
        import time
        from pqr.engine import writer

        class Data(object):
            issues = None
        data = Data(); data.issues = []
        started = time.time()
        old = self._touch("HLF-QC-126-06 안정성 시험 경향 분석 결과 - 가.xlsx", started - 3600)
        self._touch("a. 함량 Cpk 계산 파일.xlsx")                  # 이번에 만들지 않은 예전 첨부
        writer._check_attachments(self.dir, [(os.path.basename(old), old)], started, data, lambda *a: None)
        말 = [why for _, _, why in data.issues]
        self.assertTrue(any("새로 쓰이지 않았습니다" in w for w in 말))
        self.assertTrue(any("예전 첨부가 폴더에 남아" in w for w in 말))

    def test_새로_쓴_첨부만_있으면_조용하다(self):
        import time
        from pqr.engine import writer

        class Data(object):
            issues = None
        data = Data(); data.issues = []
        started = time.time() - 1
        path = self._touch("a. 함량 Cpk 계산 파일.xlsx")
        writer._check_attachments(self.dir, [(os.path.basename(path), path)], started, data, lambda *a: None)
        self.assertEqual(data.issues, [])


class 덮어쓰지_못한_첨부(unittest.TestCase):
    """엑셀에서 열어 둔 채 재작성하면 다른 이름으로 만들고 ★ 로 알린다 (담당자 2026-09-07)."""

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="pqr-free-")
        from pqr.engine import excel_attach
        self.excel_attach = excel_attach

    def _data(self):
        class Data(object):
            issues = None
        data = Data(); data.issues = []
        return data

    def test_없는_파일은_그대로(self):
        dst = os.path.join(self.dir, "가.xlsx")
        self.assertEqual(self.excel_attach.free_path(dst, self._data()), dst)

    def test_열_수_있으면_그대로(self):
        dst = os.path.join(self.dir, "가.xlsx")
        with open(dst, "wb") as handle:
            handle.write(b"x")
        data = self._data()
        self.assertEqual(self.excel_attach.free_path(dst, data), dst)
        self.assertEqual(data.issues, [])

    def test_잠겨_있으면_다른_이름과_문의(self):
        import builtins
        dst = os.path.join(self.dir, "가.xlsx")
        with open(dst, "wb") as handle:
            handle.write(b"x")
        data, said = self._data(), []
        real = builtins.open

        def locked(path, mode="r", *args, **kw):
            if path == dst and "+" in mode:
                raise OSError(13, "다른 프로그램이 쓰고 있습니다")
            return real(path, mode, *args, **kw)

        builtins.open = locked
        try:
            got = self.excel_attach.free_path(dst, data, said.append)
        finally:
            builtins.open = real
        self.assertEqual(got, os.path.join(self.dir, "가 (새로 만든 것).xlsx"))
        self.assertEqual(len(data.issues), 1)
        self.assertTrue(data.issues[0][2].startswith("★"))
        self.assertTrue(said)
