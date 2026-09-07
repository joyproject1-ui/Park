# -*- coding: utf-8 -*-
"""전년도 결재본에서 이어받기 — 해마다 바뀌지 않는 값만, 빈 칸에만.

담당자 지시(2026-09): "전년도 PQR 결재본은 장비 및 원료 등 정보를 참고하는거고, 2026년 PQR
필요 자료를 참고해서 새롭게 작성해줘야 하는거야."
"""
import unittest

import docx

from pqr.engine import carry


def _doc(rows, heading="8.1.1 주원료"):
    document = docx.Document()
    document.add_paragraph(heading)
    table = document.add_table(rows=len(rows), cols=len(rows[0]))
    for ri, cells in enumerate(rows):
        for ci, text in enumerate(cells):
            table.cell(ri, ci).text = text
    return document


HEAD = ["연번", "관리번호", "원/자재명", "규격", "제조원", "완료일"]


class 빈_칸에만_이어받는다(unittest.TestCase):
    def test_같은_코드의_규격을_가져온다(self):
        old = _doc([HEAD, ["1", "RBG201", "겐타마이신황산염", "KP", "예전 제조원", "2023.12.08"]])
        new = _doc([HEAD, ["1", "RBG201", "겐타마이신황산염", "", "올해 제조원", ""]])
        carry.carry(new, old)
        got = [c.text for c in new.tables[0].rows[1].cells]
        self.assertEqual(got[3], "KP")                 # 규격은 이어받고
        self.assertEqual(got[4], "올해 제조원")          # 올해 자료로 채운 칸은 그대로 두고
        self.assertEqual(got[5], "")                   # 완료일은 해마다 바뀌므로 이어받지 않는다

    def test_코드가_다르면_가져오지_않는다(self):
        old = _doc([HEAD, ["1", "RSF101", "플루오로메톨론", "USP", "", ""]])
        new = _doc([HEAD, ["1", "RBG201", "겐타마이신황산염", "", "", ""]])
        carry.carry(new, old)
        self.assertEqual([c.text for c in new.tables[0].rows[1].cells][3], "")


class 줄_열쇠가_해마다_다른_표(unittest.TestCase):
    """6항은 Lot No. 가 해마다 달라 짝이 없다 — 전년도 값이 한 가지면 그 값을 쓴다."""

    HEAD6 = ["연번", "Lot No.", "제조일자", "제조단위", "포장단위"]

    def test_제조단위는_전년도_값이_하나면_가져온다(self):
        old = _doc([self.HEAD6, ["1", "OGX901", "2024.09.24", "164,000g", "4g x 1Tube/Case"]],
                   "6. 제조내역 확인")
        new = _doc([self.HEAD6, ["1", "OGY301", "2025.03.11", "", ""]], "6. 제조내역 확인")
        carry.carry(new, old)
        got = [c.text for c in new.tables[0].rows[1].cells]
        self.assertEqual(got[3], "164,000g")
        self.assertEqual(got[4], "4g x 1Tube/Case")
        self.assertEqual(got[1], "OGY301")            # Lot 은 올해 것 그대로

    def test_전년도_값이_여럿이면_고르지_않는다(self):
        old = _doc([self.HEAD6,
                    ["1", "OGX901", "2024.09.24", "164,000g", "4g x 1Tube/Case"],
                    ["2", "OGX902", "2024.10.24", "82,000g", "4g x 1Tube/Case"]], "6. 제조내역 확인")
        new = _doc([self.HEAD6, ["1", "OGY301", "2025.03.11", "", ""]], "6. 제조내역 확인")
        carry.carry(new, old)
        self.assertEqual([c.text for c in new.tables[0].rows[1].cells][3], "")


class 밸리데이션_사유(unittest.TestCase):
    """안정성 시험 '실시 사유' 는 그 Lot 의 공정밸리데이션 사유다 (담당자 2026-09)."""

    def _doc(self, rows):
        document = docx.Document()
        document.add_paragraph("10.1 공정밸리데이션")
        table = document.add_table(rows=len(rows), cols=len(rows[0]))
        for ri, cells in enumerate(rows):
            for ci, text in enumerate(cells):
                lines = text.split("\n")
                table.cell(ri, ci).text = lines[0]
                for extra in lines[1:]:            # 칸 안의 줄바꿈은 문단으로 — 결재본과 같은 꼴
                    table.cell(ri, ci).add_paragraph(extra)
        return document

    def test_대상_Lot_마다_비고를_돌려준다(self):
        got = carry.pv_reasons(self._doc([
            ["No.", "사유", "대상Lot", "문서번호", "완료일", "비고"],
            ["1", "변경", "OGW701\nOGWN01", "PV24-2-DGTO1-FR", "2024.10.29", "Lot size 축소,\n용액 멸균 조건 변경"]]))
        self.assertEqual(got, {"OGW701": "Lot size 축소, 용액 멸균 조건 변경",
                               "OGWN01": "Lot size 축소, 용액 멸균 조건 변경"})

    def test_비고가_비었거나_NA_면_넣지_않는다(self):
        got = carry.pv_reasons(self._doc([
            ["No.", "사유", "대상Lot", "문서번호", "완료일", "비고"],
            ["1", "변경", "OGY301", "PV25", "2025.05.22", "N/A"]]))
        self.assertEqual(got, {})


if __name__ == "__main__":
    unittest.main()


class 전년도_값은_노랑(unittest.TestCase):
    """담당자 2026-09-07: "기록서가 없어 내용이 없거나 이상하면 전년도 PQR 정보로 가져오고 노랑 MARK 해"."""

    def _table(self, rows):
        import docx
        from pqr.engine import docedit as E
        d = docx.Document()
        t = d.add_table(rows=len(rows), cols=len(rows[0]))
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                E.set_cell(t.rows[i].cells[j], v)
        return t

    def _yellow(self, table):
        from docx.oxml.ns import qn
        from pqr.engine import docedit as E
        out = []
        for row in table.rows[1:]:
            for cell in E.raw_cells(row):
                if list(cell._tc.iter(qn("w:highlight"))):
                    out.append(E.cell_text(cell))
        return out

    def test_전년도_값_그대로인_칸만_노랑으로(self):
        from pqr.engine.recipe_ointment import _mark_carried_cells
        old = [["연번", "관리번호", "원/자재명", "제조원"],
               ["1", "EPB116", "벤잘코늄염화물", "Hubei"],
               ["2", "EPB105", "붕산", "삼전순약"]]
        table = self._table([["연번", "관리번호", "원/자재명", "제조원"],
                             ["1", "EPB116", "벤잘코늄염화물", "Hubei"],          # 그대로 — 노랑
                             ["2", "EPB105", "붕산", "올해 제조원㈜"]])           # 갱신됨 — 그대로 둔다
        self.assertEqual(_mark_carried_cells(table, old), 3)
        self.assertEqual(sorted(self._yellow(table)), sorted(["벤잘코늄염화물", "Hubei", "붕산"]))

    def test_전년도에_없는_줄은_건드리지_않는다(self):
        from pqr.engine.recipe_ointment import _mark_carried_cells
        old = [["연번", "관리번호", "원/자재명"], ["1", "EPB116", "벤잘코늄염화물"]]
        table = self._table([["연번", "관리번호", "원/자재명"], ["1", "P12061", "PE 병 10mL"]])
        self.assertEqual(_mark_carried_cells(table, old), 0)
        self.assertEqual(self._yellow(table), [])

    def test_관리번호_열이_없으면_아무것도_하지_않는다(self):
        from pqr.engine.recipe_ointment import _mark_carried_cells
        old = [["연번", "이름"], ["1", "가"]]
        table = self._table([["연번", "이름"], ["1", "가"]])
        self.assertEqual(_mark_carried_cells(table, old), 0)


class 올해_이력만(unittest.TestCase):
    """담당자 2026-09-07: "일탈이 없는데 일탈 문서번호는 왜 작성한 거야?" """

    def _doc(self, rows):
        import docx
        from pqr.engine import docedit as E
        d = docx.Document()
        d.add_paragraph("11.1 중요 일탈 및 기준일탈 내역")
        t = d.add_table(rows=len(rows), cols=len(rows[0]))
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                E.set_cell(t.rows[i].cells[j], v)
        return d, t

    def test_일탈_문서번호는_전년도에서_옮기지_않는다(self):
        from pqr.engine import carry as C, docedit as E
        old, _ = self._doc([["연번", "Lot No.", "문서번호", "일탈사항"],
                            ["1", "OEX101", "DR-240509-06", "충전 수율"]])
        new, table = self._doc([["연번", "Lot No.", "문서번호", "일탈사항"],
                                ["1", "", "", ""]])
        C.carry(new, old)
        self.assertEqual(E.cell_text(table.rows[1].cells[2]), "")

    def test_다른_항의_규격은_그대로_옮긴다(self):
        import docx
        from pqr.engine import carry as C, docedit as E
        def make(rows):
            d = docx.Document()
            d.add_paragraph("8.1.3 부원료 및 포장자재")
            t = d.add_table(rows=len(rows), cols=len(rows[0]))
            for i, row in enumerate(rows):
                for j, v in enumerate(row):
                    E.set_cell(t.rows[i].cells[j], v)
            return d, t
        old, _ = make([["연번", "관리번호", "규격"], ["1", "EPB116", "KP"]])
        new, table = make([["연번", "관리번호", "규격"], ["1", "EPB116", ""]])
        C.carry(new, old)
        self.assertEqual(E.cell_text(table.rows[1].cells[2]), "KP")
