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
