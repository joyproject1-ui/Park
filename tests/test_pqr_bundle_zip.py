# -*- coding: utf-8 -*-
"""폴더째 묶은 압축 — 디겐타안연고처럼 항 번호 없는 ZIP 하나로 올린 자료."""
from __future__ import unicode_literals

import os
import shutil
import tempfile
import unittest
import zipfile

from pqr import build, server
from pqr.engine import collect

ITEMS = [["3", "허가증", ""], ["6", "제조내역", ""], ["9.2.1", "조제 완료 후", ""],
         ["13", "안정성", ""]]


class _Cp949Info(zipfile.ZipInfo):
    """Windows 탐색기가 만드는 압축 — 이름을 cp949 로 적고 UTF-8 표시를 안 켠다."""

    def _encodeFilenameFlags(self):
        return self.filename.encode("cp949"), 0


def make_bundle(path, wrapper="디겐타 안연고 2026년 PQR 필요 자료", cp949=False):
    """담당자가 Windows 에서 폴더째 묶은 압축을 흉내 냅니다."""
    names = ["3. 허가증.pdf", "6. 제조내역 - ERP.pdf", "9.2.1 조제 완료 후/OZ101.pdf",
             "필요 자료/13 안정성 시험일지.pdf", "메모.txt"]
    with zipfile.ZipFile(path, "w") as archive:
        for name in names:
            full = "%s/%s" % (wrapper, name) if wrapper else name
            if cp949:
                archive.writestr(_Cp949Info(full), b"x")
            else:
                archive.writestr(full, b"x")


class 수집현황(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp()
        self.folder = os.path.join(self.root, "QC1-7014 디겐타안연고")
        os.makedirs(self.folder)

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    def test_번호_없는_압축_안의_항을_본다(self):
        make_bundle(os.path.join(self.folder, "디겐타 안연고 2026년 PQR 필요 자료.zip"))
        found = build.collect_item_files(self.root, ITEMS)["QC1-7014"]
        self.assertEqual(sorted(found), ["13", "3", "6", "9.2.1"])
        self.assertTrue(found["3"][0].endswith("/3. 허가증.pdf"))
        self.assertEqual(len(found["9.2.1"]), 1)           # 폴더 하나로 한 번만

    def test_한글이_cp437_로_깨진_압축도_읽는다(self):
        make_bundle(os.path.join(self.folder, "묶음.zip"), cp949=True)
        found = build.collect_item_files(self.root, ITEMS)["QC1-7014"]
        self.assertIn("3", found)
        self.assertIn("허가증", found["3"][0])

    def test_번호_없는_중간_폴더도_지나간다(self):
        inner = os.path.join(self.folder, "필요 자료", "9.2.1 조제 완료 후")
        os.makedirs(inner)
        open(os.path.join(inner, "a.pdf"), "w").close()
        found = build.collect_item_files(self.root, ITEMS)["QC1-7014"]
        self.assertEqual(found, {"9.2.1": ["필요 자료/9.2.1 조제 완료 후"]})

    def test_번호_붙은_압축은_전처럼_그_항이다(self):
        with zipfile.ZipFile(os.path.join(self.folder, "9.2.1 조제 완료 후 - ERP.zip"), "w") as z:
            z.writestr("a.pdf", b"x")
        found = build.collect_item_files(self.root, ITEMS)["QC1-7014"]
        self.assertEqual(found, {"9.2.1": ["9.2.1 조제 완료 후 - ERP.zip"]})


class 압축풀기(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    def test_겉_폴더를_벗기고_제품_폴더에_푼다(self):
        archive = os.path.join(self.root, "묶음.zip")
        make_bundle(archive)
        folder = os.path.join(self.root, "QC1-7014 디겐타안연고")
        os.makedirs(folder)
        written = server.extract_bundle(archive, folder)
        self.assertIn("3. 허가증.pdf", written)
        self.assertIn("9.2.1 조제 완료 후/OZ101.pdf", written)
        self.assertTrue(os.path.exists(os.path.join(folder, "9.2.1 조제 완료 후", "OZ101.pdf")))
        self.assertFalse(os.path.exists(os.path.join(folder, "디겐타 안연고 2026년 PQR 필요 자료")))

    def test_겉_폴더가_없으면_그대로_푼다(self):
        archive = os.path.join(self.root, "묶음.zip")
        make_bundle(archive, wrapper="")
        folder = os.path.join(self.root, "p")
        os.makedirs(folder)
        written = server.extract_bundle(archive, folder)
        self.assertIn("3. 허가증.pdf", written)

    def test_상위_폴더로_빠져나가는_항목은_버린다(self):
        archive = os.path.join(self.root, "나쁜.zip")
        with zipfile.ZipFile(archive, "w") as z:
            z.writestr("../탈출.txt", b"x")
            z.writestr("정상.pdf", b"x")
        folder = os.path.join(self.root, "p")
        os.makedirs(folder)
        self.assertEqual(server.extract_bundle(archive, folder), ["정상.pdf"])
        self.assertFalse(os.path.exists(os.path.join(self.root, "탈출.txt")))

    def test_빈_압축은_거부한다(self):
        archive = os.path.join(self.root, "빈.zip")
        with zipfile.ZipFile(archive, "w"):
            pass
        with self.assertRaises(server.UploadError):
            server.extract_bundle(archive, self.root)


class 보고서엔진(unittest.TestCase):
    def test_discover_가_번호_없는_압축_안을_본다(self):
        root = tempfile.mkdtemp()
        try:
            make_bundle(os.path.join(root, "디겐타 안연고 2026년 PQR 필요 자료.zip"))
            items = collect.discover(root, workdir=tempfile.mkdtemp(dir=root))
            self.assertEqual(sorted(items), ["13", "3", "6", "9.2.1"])
            self.assertTrue(items["9.2.1"][0].endswith("OZ101.pdf"))
        finally:
            shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
