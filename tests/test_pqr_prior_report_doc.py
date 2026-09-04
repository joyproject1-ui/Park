# -*- coding: utf-8 -*-
"""옛 워드(.doc)를 워드 문서로 착각하지 않는다.

.doc 는 OLE 파일인데 꼬리에 테마(.thmx) zip 이 박혀 있어 zipfile 이 그것을 열어 준다.
그대로 복제하면 테마 조각만 든 2 KB 파일이 나오고 Word 가 거부한다.
"""
from __future__ import unicode_literals

import io
import os
import tempfile
import unittest
import zipfile

from pqr import prior_report


def theme_only_doc(path):
    """실제 .doc 처럼 OLE 머리 + 꼬리에 테마 zip 이 붙은 파일을 만든다."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("theme/theme/theme1.xml", "<a:theme/>")
        z.writestr("theme/theme/themeManager.xml", "<a:themeManager/>")
    with open(path, "wb") as h:
        h.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 2048)
        h.write(buf.getvalue())
    return path


def real_docx(path):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("word/document.xml", "<w:document/>")
    return path


class 진짜문서인지(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()

    def test_doc_속_테마는_문서가_아니다(self):
        p = theme_only_doc(os.path.join(self.dir, "PQR25.doc"))
        payload, _ = prior_report.read_document(p)
        self.assertTrue(zipfile.is_zipfile(io.BytesIO(payload)))   # zip 으로는 열린다
        self.assertFalse(prior_report.is_docx(payload))            # 그래도 문서는 아니다

    def test_진짜_docx_는_통과(self):
        p = real_docx(os.path.join(self.dir, "a.docx"))
        payload, _ = prior_report.read_document(p)
        self.assertTrue(prior_report.is_docx(payload))

    def test_비어_있거나_zip_이_아니면_False(self):
        self.assertFalse(prior_report.is_docx(None))
        self.assertFalse(prior_report.is_docx(b""))
        self.assertFalse(prior_report.is_docx(b"not a zip"))

    def test_doc_로는_사본을_만들지_않는다(self):
        """전에는 테마 조각만 든 파일을 '제출용' 으로 내놓았다."""
        src = theme_only_doc(os.path.join(self.dir, "PQR25.doc"))
        target = os.path.join(self.dir, "제출용.docx")
        self.assertIsNone(prior_report.write_from_previous(src, target, 2026))
        self.assertFalse(os.path.exists(target))


if __name__ == "__main__":
    unittest.main()
