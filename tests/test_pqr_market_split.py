"""내수용·수출용이 따로인 서식(퀴노비드) — 전년도 결재본 읽기, 13항 표 가르기, 시판 후 완료 판정, 시험일지 여러 쪽 (2026-09)."""
import os
import sys
import unittest

import docx

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import carry, docedit as E, handwriting     # noqa: E402
from pqr.engine.recipe_ointment import _stability_tables, _post_completed, _has_equipment, _seed_equipment_rows   # noqa: E402


def _doc(headings_tables):
    d = docx.Document()
    for heading, table in headings_tables:
        d.add_paragraph(heading)
        if table:
            t = d.add_table(rows=len(table), cols=len(table[0]))
            for i, row in enumerate(table):
                for j, v in enumerate(row):
                    E.set_cell(t.rows[i].cells[j], v)
    return d


class 표_가르기(unittest.TestCase):
    def test_내수용_수출용_시판후_장기_경향을_제목으로_가른다(self):
        d = _doc([("13. 안정성시험", None), ("13.1 장기 안정성 시험", None),
                  ("13.1.1 내수용", [["연번", "제조 번호"]]), ("13.1.2 수출용 (베트남)", [["연번", "제조 번호"]]),
                  ("13.2 시판 후 안정성 시험", None), ("13.2.1 내수용", [["연번", "비고"]]),
                  ("13.3 안정성 시험 경향 분석 결과", None), ("13.3.1 내수용", [["시험항목", "함량(%)"]]),
                  ("14. 반품", [["a"]])])
        got = [(k, m) for k, m, _ in _stability_tables(d)]
        self.assertEqual(got, [("장기", "내수"), ("장기", "수출"), ("시판후", "내수"), ("경향", "내수")])

    def test_디겐타_서식은_시장_없이_장기와_경향뿐(self):
        d = _doc([("13.1 장기 안정성 시험", [["연번"]]), ("13.2 시판 후 안정성 시험 이력 없음.", None),
                  ("13.3 안정성 시험 경향 분석 결과", [["시험항목"]])])
        self.assertEqual([(k, m) for k, m, _ in _stability_tables(d)], [("장기", ""), ("경향", "")])


class 시판후_완료(unittest.TestCase):
    def test_사용기한까지_시험했으면_완료(self):
        one = {"mfg": "2022.03.18", "expiry": "2025.03.17", "points": [{"period": "36M"}]}
        self.assertTrue(_post_completed(one, []))
        one = {"mfg": "2023.01.06", "expiry": "2026.01.05", "points": [{"period": "24M"}]}
        self.assertFalse(_post_completed(one, []))


class 전년도_결재본(unittest.TestCase):
    def test_포장과_설비줄과_10_1_표를_읽는다(self):
        d = _doc([("10.1 공정밸리데이션", None), ("10.1.1 내수용", [["사유", "대상Lot", "문서번호"], ["변경", "OEX101", "PV24-2-QUIO3-R"]]),
                  ("10.1.2 수출용", [["사유", "대상Lot", "문서번호"], ["변경", "OAX101", "PV24-2-QUIO2-R"]]),
                  ("10.3 공기조화장치", [["No.", "관리번호", "설비명", "IQ", "OQ", "PQ"], ["", "", "", "문서번호", "문서번호", "문서번호"],
                                     ["", "", "", "완료일", "완료일", "완료일"], ["", "", "", "", "", ""],
                                     ["1", "HBA5089", "A.H.U 26호", "IQ18-HA-HBA5089-R", "OQ18-HA-HBA5089-R", "PQ18-HA-HBA5089-FR"],
                                     ["1", "", "", "2018.09.03", "2018.11.01", "2019.12.04"]]),
                  ("13.1 시판 후 안정성 시험", None),
                  ("13.1.1 내수용", [["연번", "제조 번호", "포장 형태"], ["1", "OEV301", "5g tube/갑"]]),
                  ("13.1.2 수출용 (베트남)", [["연번", "제조 번호", "포장 형태"], ["1", "OZW101", "3.5g tube/갑"]])])
        packs = carry.stability_packs(d)
        self.assertEqual(packs["by_lot"], {"OEV301": "5g tube/갑", "OZW101": "3.5g tube/갑"})
        self.assertEqual(packs["by_market"], {"내수": "5g tube/갑", "수출": "3.5g tube/갑"})
        self.assertEqual(packs["market_by_lot"], {"OEV301": "내수", "OZW101": "수출"})
        self.assertEqual(packs["market_by_prefix"], {"OE": "내수", "OZ": "수출"})
        eq = carry.equipment_rows(d)
        self.assertEqual(eq["10.3"][0]["mid"], "HBA5089")
        self.assertEqual(eq["10.3"][0]["docs"]["OQ"], ("OQ18-HA-HBA5089-R", "2018.11.01"))
        grids = carry.section_grids_all(d, "10.1")
        self.assertEqual(len(grids), 2)
        self.assertIn("PV24-2-QUIO2-R", grids[1][1][2])

    def test_빈_설비_표에_전년도_줄을_세운다(self):
        d = _doc([("10.3 공기조화장치", [["No.", "관리번호", "설비명", "완료일", "", "", "비고"], ["No.", "관리번호", "설비명", "IQ", "OQ", "PQ", "비고"],
                                     ["", "", "", "문서번호", "문서번호", "문서번호", ""], ["", "", "", "완료일", "완료일", "완료일", ""],
                                     ["1", "", "", "", "", "", ""], ["1", "", "", "", "", "", ""]])])
        t = d.tables[0]
        self.assertFalse(_has_equipment(t))
        n = _seed_equipment_rows(t, [{"mid": "HBA5089", "name": "A.H.U 26호", "docs": {"IQ": ("IQ18-HA-HBA5089-R", "2018. 09. 03")}},
                                     {"mid": "HBA5033", "name": "A.H.U 4호", "docs": {}}])
        self.assertEqual(n, 2)
        self.assertTrue(_has_equipment(t))
        rows = [[E.cell_text(c).strip() for c in E.raw_cells(r)] for r in t.rows[4:]]
        self.assertEqual(rows[0][:4], ["1", "HBA5089", "A.H.U 26호", "IQ18-HA-HBA5089-R"])
        self.assertEqual(rows[1][3], "2018.09.03")
        self.assertEqual(rows[2][:3], ["2", "HBA5033", "A.H.U 4호"])
        self.assertEqual(len(rows), 4)


class 시험일지_여러_쪽(unittest.TestCase):
    def test_read_folder_는_쪽마다_한_Lot(self):
        import types
        fake = [{"lot": "OEV301", "year": "2022", "page": 1, "points": []}, {"lot": "OEW101", "year": "2023", "page": 2, "points": []}]
        old = handwriting.read_pages
        handwriting.read_pages = lambda p, specs, log: list(fake)
        try:
            out = handwriting.read_folder(["/x/a.pdf"], None, None)
        finally:
            handwriting.read_pages = old
        self.assertEqual([o["lot"] for o in out], ["OEV301", "OEW101"])
