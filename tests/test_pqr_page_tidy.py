"""빈 쪽·쪽 맨 위 빈 줄, 14.2 한 줄, 16항 내어쓰기, 판독 파일에 없는 시험일지 읽기 (담당자 2026-09-06)."""
import os
import sys
import unittest

import docx
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import docedit as E, collect                       # noqa: E402


def _break_para(d):
    p = d.add_paragraph()
    r = p.add_run()
    r._r.append(r._r.makeelement(qn("w:br"), {qn("w:type"): "page"}))
    return p


def _pbb(el):
    pr = el.find(qn("w:pPr"))
    return pr is not None and pr.find(qn("w:pageBreakBefore")) is not None


class 쪽나눔_정리(unittest.TestCase):
    def test_쪽나눔_뒤_빈_문단을_건너뛰어_제목에_앞에서_쪽나눔(self):
        d = docx.Document()
        d.add_paragraph("각주 1)")
        _break_para(d); d.add_paragraph(""); d.add_paragraph("")
        d.add_paragraph("9.2.2.3 포장 완료 후")
        n = E.hard_breaks_to_page_break_before(d)
        self.assertEqual(n, 1)
        texts = [p.text for p in d.paragraphs]
        self.assertEqual(texts, ["각주 1)", "9.2.2.3 포장 완료 후"])
        self.assertTrue(_pbb(d.paragraphs[1]._p))

    def test_표_앞_쪽나눔은_표_첫_문단의_앞에서_쪽나눔으로(self):
        d = docx.Document()
        d.add_paragraph("앞 글")
        _break_para(d); d.add_paragraph("")
        t = d.add_table(rows=1, cols=1); t.rows[0].cells[0].text = "표"
        n = E.hard_breaks_to_page_break_before(d)
        self.assertEqual(n, 1)
        self.assertEqual([p.text for p in d.paragraphs], ["앞 글"])
        self.assertTrue(_pbb(t._tbl.find(".//" + qn("w:p"))))


class 불만_없음_한_줄(unittest.TestCase):
    def test_빈_줄_셋을_한_줄로(self):
        d = docx.Document()
        t = d.add_table(rows=5, cols=3)
        for j, h in enumerate(("연번", "Lot No.", "문서번호")):
            t.rows[0].cells[j].text = h
        t.rows[1].cells[0].text = "1"
        for j in (0, 1):
            E.set_vmerge(t.rows[1].cells[j], "restart"); E.set_vmerge(t.rows[2].cells[j], None); E.set_vmerge(t.rows[3].cells[j], None)
        t.rows[4].cells[0].text = "특이사항 (Comment)"
        self.assertEqual(E.single_blank_row(t), 2)
        self.assertEqual(len(t.rows), 3)
        # 내역이 없으면 연번도 비운다 (담당자 2026-09-07: "불만이 없는데 연번에 1 표시됨")
        self.assertEqual(E.cell_text(E.raw_cells(t.rows[1])[0]), "")
        self.assertIsNone(E.raw_cells(t.rows[1])[0]._tc.find(qn("w:tcPr")).find(qn("w:vMerge")))

    def test_글이_있는_줄은_줄이지_않는다(self):
        d = docx.Document()
        t = d.add_table(rows=4, cols=2)
        t.rows[0].cells[0].text = "연번"; t.rows[1].cells[1].text = "불만 1건"; t.rows[3].cells[0].text = "특이사항 (Comment)"
        self.assertEqual(E.single_blank_row(t), 0)
        self.assertEqual(len(t.rows), 4)


class 결론_내어쓰기(unittest.TestCase):
    def test_번호_문단은_번호_폭만큼_내어쓴다(self):
        d = docx.Document()
        d.add_paragraph("16. 결론")
        d.add_paragraph("16.1 본 제품 품질 평가를 통해")
        d.add_paragraph("보통 문단")
        d.add_paragraph("17. 참고")
        self.assertEqual(E.set_section_indent(d, 16, 2), 2)
        ind = d.paragraphs[1]._p.find(qn("w:pPr")).find(qn("w:ind"))
        # 왼쪽 들여쓰기는 2칸 그대로이고 번호만 그만큼 내어쓴다 — 첫 줄이 여백 밖으로 나가지 않게
        # (담당자 2026-09-07: "16.1~16.4 는 단락 들여쓰기 왼쪽 2글자로 변경해 줘")
        self.assertEqual(ind.get(qn("w:leftChars")), "200")
        self.assertEqual(ind.get(qn("w:hangingChars")), "200")
        ind = d.paragraphs[2]._p.find(qn("w:pPr")).find(qn("w:ind"))
        self.assertEqual(ind.get(qn("w:leftChars")), "200")
        self.assertIsNone(ind.get(qn("w:hangingChars")))


class 판독_파일에_없는_일지(unittest.TestCase):
    def test_source_가_없는_일지만_새로_읽는다(self):
        logs = [{"lot": "OEV301", "source": "13 시판후 안정성시험일지.pdf"}]
        scanned = ["/x/13 시판후 안정성시험일지.pdf", "/x/13 장기 안정성시험일지(내수용).pdf"]
        self.assertEqual(collect.unread_scans(logs, scanned), ["/x/13 장기 안정성시험일지(내수용).pdf"])

    def test_어느_일지를_읽었는지_모르는_옛_판독_파일은_그대로(self):
        self.assertEqual(collect.unread_scans([{"lot": "OEV301"}], ["/x/a.pdf"]), [])


if __name__ == "__main__":
    unittest.main()


class 제목_뒤_쪽_나눔(unittest.TestCase):
    """담당자 2026-09-07: "16항 결론은 왜 다음 페이지로 밀렸는지? 21페이지에 작성되도록 해"."""

    def _doc(self):
        import docx
        from docx.oxml.ns import qn
        from pqr.engine.docedit import get_or_add
        d = docx.Document()
        d.add_paragraph("16. 결론")
        d.add_paragraph("")
        body = d.add_paragraph("한림포비돈점안액에 대한 제품품질평가 결과 …")
        get_or_add(body._p.get_or_add_pPr(), "pageBreakBefore")
        d.add_paragraph("17. 참고 자료")
        other = d.add_paragraph("- 제품표준서")
        d.add_paragraph("아무 글")
        late = d.add_paragraph("떨어진 문단")
        get_or_add(late._p.get_or_add_pPr(), "pageBreakBefore")
        return d, body, late

    def test_제목_바로_뒤_본문의_쪽_나눔만_뗀다(self):
        from docx.oxml.ns import qn
        from pqr.engine import docedit as E
        d, body, late = self._doc()
        self.assertEqual(E.drop_break_after_headings(d), 1)
        self.assertIsNone(body._p.pPr.find(qn("w:pageBreakBefore")))
        self.assertIsNotNone(late._p.pPr.find(qn("w:pageBreakBefore")))   # 제목 뒤가 아닌 것은 그대로
