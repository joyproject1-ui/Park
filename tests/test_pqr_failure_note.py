# -*- coding: utf-8 -*-
"""결재본 양식으로 만들지 못했을 때 — 원인 파일이 작성본 폴더에 남는다 (담당자 2026-09)."""
from __future__ import unicode_literals

import io
import os
import shutil
import tempfile
import unittest

from pqr import build as build_module
from pqr import server as server_module


class 실패_원인_파일(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_작성본_폴더에_까닭과_멈춘_자리를_남긴다(self):
        made = os.path.join(self.dir, build_module.OUTPUT_DIR)
        path = server_module.write_failure_note(made, {"code": "QC1-7014", "name": "디겐타안연고"},
                                                "전년도 결재본을 읽지 못함", "Traceback\n  File x", ["한 일 1"])
        self.assertTrue(path and os.path.isfile(path))
        self.assertEqual(os.path.basename(path), "★ 보고서 작성 실패 원인 (QC1-7014).txt")
        with io.open(path, encoding="utf-8") as handle:
            text = handle.read()
        self.assertIn("전년도 결재본을 읽지 못함", text)
        self.assertIn("File x", text)
        self.assertIn("한 일 1", text)
        self.assertIn("프로그램 버전", text)


if __name__ == "__main__":
    unittest.main()


class 보고서가_열려_있으면(unittest.TestCase):
    """담당자 PC 2026-09-08: 워드로 보고서를 열어 둔 채 재작성해 PermissionError 로 작성 전체가 실패했다."""

    def setUp(self):
        import tempfile
        self.dir = tempfile.mkdtemp(prefix="pqr-open-")
        self.path = os.path.join(self.dir, "[QC1-5087] 아이퓨어점안액 2026년 제품품질평가 (제출용).docx")

    class _Data(object):
        def __init__(self):
            self.issues = []

    def test_없는_파일은_그대로(self):
        from pqr.engine import writer
        data = self._Data()
        self.assertEqual(writer._writable(self.path, data, lambda *a: None), self.path)
        self.assertEqual(data.issues, [])

    def test_열_수_있으면_그대로(self):
        from pqr.engine import writer
        with open(self.path, "wb") as handle:
            handle.write(b"x")
        data = self._Data()
        self.assertEqual(writer._writable(self.path, data, lambda *a: None), self.path)
        self.assertEqual(data.issues, [])

    def test_열려_있으면_다른_이름으로_저장하고_알린다(self):
        import builtins
        from pqr.engine import writer
        with open(self.path, "wb") as handle:
            handle.write(b"x")
        real, data, said = builtins.open, self._Data(), []

        def locked(path, mode="r", *a, **kw):
            if path == self.path and "+" in mode:
                raise PermissionError(13, "Permission denied")
            return real(path, mode, *a, **kw)

        builtins.open = locked
        try:
            got = writer._writable(self.path, data, said.append)
        finally:
            builtins.open = real
        self.assertTrue(got.endswith("(새로 만든 것).docx"))
        self.assertEqual(len(data.issues), 1)
        self.assertTrue(data.issues[0][2].startswith("★"))
        self.assertIn("워드에서 열려 있어", data.issues[0][2])
        self.assertTrue(said)
