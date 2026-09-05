# -*- coding: utf-8 -*-
"""내보내기 전 검사 — Word 가 복구를 묻게 하는 문제를 잡는다."""
from __future__ import unicode_literals

import os
import tempfile
import unittest
import zipfile

from pqr.engine import polish, verify

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-'
      'officedocument.wordprocessingml.document.main+xml"/></Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
             'relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOC = '<w:document %s><w:body><w:p/></w:body></w:document>' % W


def make(parts):
    path = os.path.join(tempfile.mkdtemp(), "t.docx")
    base = {"[Content_Types].xml": CT, "_rels/.rels": ROOT_RELS, "word/document.xml": DOC}
    base.update(parts)
    with zipfile.ZipFile(path, "w") as z:
        for n, data in base.items():
            z.writestr(n, data)
    return path


class 설정파일차례(unittest.TestCase):
    def test_updateFields_를_끝에_붙이면_잡는다(self):
        bad = ('<w:settings %s><w:zoom w:percent="100"/><w:compat><w:noLeading/></w:compat>'
               '<w:updateFields w:val="true"/></w:settings>' % W)
        got = verify.check_docx(make({"word/settings.xml": bad}))
        self.assertTrue(got)
        self.assertIn("차례가 뒤집혔", got[0])
        self.assertIn("updateFields", got[0])

    def test_제자리에_있으면_통과(self):
        good = ('<w:settings %s><w:zoom w:percent="100"/><w:updateFields w:val="true"/>'
                '<w:compat><w:noLeading/></w:compat></w:settings>' % W)
        self.assertEqual(verify.check_docx(make({"word/settings.xml": good})), [])

    def test_polish_가_제자리에_넣는다(self):
        import re
        x = ('<?xml version="1.0"?><w:settings %s><w:zoom w:percent="100"/>'
             '<w:compat><w:noLeading/></w:compat><w:rsids><w:rsid w:val="1"/></w:rsids>'
             '</w:settings>' % W)
        out = polish.set_update_fields(x)
        order = re.findall(r"<w:([A-Za-z]+)[ />]", out)[1:]
        self.assertLess(order.index("updateFields"), order.index("compat"))
        self.assertEqual(verify.check_docx(make({"word/settings.xml": out})), [])

    def test_두_번_불러도_하나만_넣는다(self):
        x = '<?xml version="1.0"?><w:settings %s><w:zoom w:percent="100"/></w:settings>' % W
        twice = polish.set_update_fields(polish.set_update_fields(x))
        self.assertEqual(twice.count("updateFields"), 1)

    def test_clear_update_fields_가_지운다(self):
        """담당자 2026-09: 목차가 전부 '1' — updateFields 가 켜져 있으면 Word 가 쪽 나눔
        전에 필드를 다시 계산한다. 만든 파일에는 이 표시가 없어야 한다."""
        x = ('<?xml version="1.0"?><w:settings %s><w:zoom w:percent="100"/>'
             '<w:updateFields w:val="true"/><w:compat><w:noLeading/></w:compat></w:settings>' % W)
        out = polish.clear_update_fields(x)
        self.assertNotIn("updateFields", out)
        self.assertIn("compat", out)
        self.assertEqual(verify.check_docx(make({"word/settings.xml": out})), [])

    def test_polish_는_settings_의_updateFields_를_없앤다(self):
        import zipfile
        path = make({"word/settings.xml":
                     '<?xml version="1.0"?><w:settings %s><w:zoom w:percent="100"/>'
                     '<w:updateFields w:val="true"/></w:settings>' % W})
        polish.polish(path)
        with zipfile.ZipFile(path) as z:
            self.assertNotIn("updateFields", z.read("word/settings.xml").decode("utf-8"))


class 패키지(unittest.TestCase):
    def test_형식_등록이_없으면_잡는다(self):
        got = verify.check_docx(make({"word/media/logo.png": "x"}))
        self.assertTrue(any("부품 형식" in g for g in got), got)

    def test_관계가_없는_부품을_가리키면_잡는다(self):
        rels = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/'
                'package/2006/relationships"><Relationship Id="rId9" Type="http://x" '
                'Target="header9.xml"/></Relationships>')
        got = verify.check_docx(make({"word/_rels/document.xml.rels": rels}))
        self.assertTrue(any("없는 부품을 가리킵니다" in g for g in got), got)

    def test_스타일_이름이_겹치면_잡는다(self):
        st = ('<w:styles %s><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/>'
              '</w:style><w:style w:type="paragraph" w:styleId="a"><w:name w:val="Normal"/>'
              '</w:style></w:styles>' % W)
        got = verify.check_docx(make({"word/styles.xml": st}))
        self.assertTrue(any("이름이 겹치는 스타일" in g for g in got), got)

    def test_멀쩡한_문서는_통과(self):
        self.assertEqual(verify.check_docx(make({})), [])


if __name__ == "__main__":
    unittest.main()
