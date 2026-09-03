# -*- coding: utf-8 -*-
"""입력 폴더에서 평가항목 자료를 찾아내는 규칙."""
from __future__ import unicode_literals

import os
import shutil
import tempfile
import unittest

from pqr.engine.collect import discover


class DiscoverTest(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="pqr-discover-")

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    def touch(self, *parts):
        path = os.path.join(self.root, *parts)
        directory = os.path.dirname(path)
        if directory and not os.path.isdir(directory):
            os.makedirs(directory)
        open(path, "w").close()
        return path

    def test_이름이_항번호로_시작하는_파일(self):
        self.touch("7. 수율현황표.xlsx")
        self.assertEqual(list(discover(self.root)), ["7"])

    def test_번호_붙은_폴더_안의_파일은_이름에_번호가_없어도_된다(self):
        self.touch("13. 안정성 시험", "퀴노비드안연고(내수용) 시판후 안정성시험일지.pdf")
        self.touch("13. 안정성 시험", "퀴노비드안연고(수출용) 장기 안정성시험일지.pdf")
        got = discover(self.root)
        self.assertEqual(sorted(os.path.basename(p) for p in got["13"]),
                         ["퀴노비드안연고(내수용) 시판후 안정성시험일지.pdf",
                          "퀴노비드안연고(수출용) 장기 안정성시험일지.pdf"])

    def test_번호_없는_중간_폴더는_지나쳐_들어간다(self):
        self.touch("필요 자료", "13. 안정성 시험", "내수용 장기.pdf")
        self.touch("필요 자료", "7. 수율현황표.xlsx")
        got = discover(self.root)
        self.assertEqual(sorted(got), ["13", "7"])

    def test_너무_깊으면_그만_들어간다(self):
        self.touch("a", "b", "c", "d", "7. 수율현황표.xlsx")
        self.assertEqual(discover(self.root, depth=2), {})

    def test_임시파일과_숨김파일은_건너뛴다(self):
        self.touch("~$7. 수율현황표.xlsx")
        self.touch(".7. 수율현황표.xlsx")
        self.assertEqual(discover(self.root), {})


class StabilityGroupTest(unittest.TestCase):
    def _record(self, lot, name, package):
        return {"lot": lot, "product_name": name, "test_type": "시판후",
                "mfg_date": "2023.01.20", "package": package, "uncertain": [],
                "points": [{"label": "초기", "assay": "97.3", "tested": True, "confidence": 0.9,
                            "test_date": "2023.03.06", "reviewer_date": "2023.03.06"},
                           {"label": "24M", "assay": "103.5", "tested": True, "confidence": 0.9,
                            "test_date": "2025.02.12", "reviewer_date": "2025.02.17"}]}

    def test_판독값은_포장_규격별로_나뉜다(self):
        from pqr.engine.vision_claude import build_tables

        class Data(object):
            export = ["OZW101"]
            issues = []

        stab, _ = build_tables([self._record("OEV301", "퀴노비드안연고", "5g x Tube/갑"),
                                self._record("OZW101", "퀴노비드안연고", "3.5g x Tube/갑")],
                               Data(), {"from": "2025-01-01", "to": "2025-12-31"})
        self.assertEqual(sorted(stab["points"]), ["내수용", "수출용"])
        self.assertEqual(list(stab["points"]["내수용"]), ["OEV301"])
        self.assertEqual(stab["points"]["수출용"]["OZW101"]["24M"], "103.5")

    def test_나뉜_판독값은_파일_두_개로_이어진다(self):
        from pqr.engine.excel_attach import _grouped
        got = _grouped({"내수용": {"OEV301": {"Initial": 100.5}},
                        "수출용": {"OZW101": {"Initial": 97.3}}})
        self.assertEqual(sorted(k for k, _ in got), ["내수용", "수출용"])
