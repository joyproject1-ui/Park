"""IQ·OQ 를 하나로 합친 문서(IOQ20-UT-HEA5029-R)는 두 칸을 합쳐 한 번만 적는다 —
담당자 PC(2026-09-06)에서 공양식의 합친 칸이 사선만 남은 채 비어 있었다."""
import os
import sys
import unittest

import docx
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import docedit as E                        # noqa: E402
from pqr.engine.recipe_ointment import update_qualification, _spans   # noqa: E402


def _표(합침=False, 글=""):
    doc = docx.Document()
    t = doc.add_table(rows=6, cols=7)
    for j, v in enumerate(["No.", "관리번호", "설비명", "완료일", "", "", "비고"]):
        E.set_cell(t.rows[0].cells[j], v)
    for j, v in enumerate(["", "", "", "IQ", "OQ", "PQ", ""]):
        E.set_cell(t.rows[1].cells[j], v)
    for ri, 이름 in ((2, "문서번호"), (3, "완료일")):
        for j, v in enumerate(["", "", "", 이름, 이름, 이름, ""]):
            E.set_cell(t.rows[ri].cells[j], v)
    for j, v in enumerate(["1", "HEA5029", "질소 분배라인", "", "", "", ""]):
        E.set_cell(t.rows[4].cells[j], v)
    if 합침:
        for ri in (4, 5):
            row = t.rows[ri]
            E.set_cell(row.cells[3], 글)
            E.add_diag(row.cells[3])
            E.merge_right(row.cells[3], row.cells[4])
    else:
        for ri in (4, 5):
            for j in (3, 4):
                E.add_diag(t.rows[ri].cells[j])
    return t


IOQ = {"HEA5029": {"IQ": [("IOQ20-UT-HEA5029-R", "2020.05.05")],
                   "OQ": [("IOQ20-UT-HEA5029-R", "2020.05.05")],
                   "PQ": [("QM26-F1,2,3-N2-R", "2026.05.23")]}}
SPLIT = {"HEA5029": {"IQ": [("IQ20-2-HEA5030-R", "2020.04.29")],
                     "OQ": [("OPQ20-2-HEA5030-R", "2020.05.06")],
                     "PQ": [("QM26-F1,F2,-CA-R", "2026.08.25")]}}


def _cells(row):
    return [(E.cell_text(c).strip(), _spans(c)) for c in E.raw_cells(row)]


class 합친_칸(unittest.TestCase):
    def test_비어_있는_합친_칸에_IOQ_문서를_적는다(self):
        t = _표(합침=True)
        update_qualification(t, IOQ)
        self.assertEqual(_cells(t.rows[4])[3], ("IOQ20-UT-HEA5029-R", 2))
        self.assertEqual(_cells(t.rows[5])[3], ("2020.05.05", 2))
        self.assertEqual(_cells(t.rows[4])[4], ("QM26-F1,2,3-N2-R", 1))
        self.assertFalse(E.has_diag(t.rows[4].cells[3]))
        self.assertFalse(E.has_diag(t.rows[5].cells[3]))

    def test_따로_있는_빈_칸은_합쳐서_한_번만_적는다(self):
        t = _표()
        update_qualification(t, IOQ)
        self.assertEqual(_cells(t.rows[4])[3:5], [("IOQ20-UT-HEA5029-R", 2), ("QM26-F1,2,3-N2-R", 1)])
        self.assertEqual(_cells(t.rows[5])[3:5], [("2020.05.05", 2), ("2026.05.23", 1)])
        self.assertEqual(len(E.raw_cells(t.rows[4])), 6)
        self.assertFalse(E.has_diag(t.rows[4].cells[3]))

    def test_합친_칸이_비었는데_IQ_OQ_가_따로면_칸을_나눠_적는다(self):
        t = _표(합침=True)
        update_qualification(t, SPLIT)
        self.assertEqual(_cells(t.rows[4])[3:6], [("IQ20-2-HEA5030-R", 1), ("OPQ20-2-HEA5030-R", 1), ("QM26-F1,F2,-CA-R", 1)])
        self.assertEqual(_cells(t.rows[5])[3:6], [("2020.04.29", 1), ("2020.05.06", 1), ("2026.08.25", 1)])

    def test_값이_적힌_합친_칸은_다른_문서로_나누지_않는다(self):
        t = _표(합침=True, 글="IOQ21-WS-HCA5152_1,HCA5152_2")
        update_qualification(t, SPLIT)
        self.assertEqual(_cells(t.rows[4])[3], ("IOQ21-WS-HCA5152_1,HCA5152_2", 2))

    def test_이미_적힌_IOQ_는_그대로(self):
        t = _표(합침=True, 글="IOQ20-UT-HEA5029-R")
        E.set_cell(t.rows[5].cells[3], "2020.05.05")
        n = update_qualification(t, IOQ)
        self.assertEqual(_cells(t.rows[4])[3], ("IOQ20-UT-HEA5029-R", 2))
        self.assertEqual(n, 1)                       # PQ 만 새로 적힌다


if __name__ == "__main__":
    unittest.main()


class 검토(unittest.TestCase):
    def test_마스터에_문서가_있는데_빈_칸이면_잡아낸다(self):
        from pqr.engine.recipe_ointment import blank_qualification_cells
        t = _표(합침=True)
        self.assertEqual(blank_qualification_cells(t, IOQ), [("HEA5029", "IQ"), ("HEA5029", "PQ")])
        update_qualification(t, IOQ)
        self.assertEqual(blank_qualification_cells(t, IOQ), [])

    def test_마스터에도_없는_칸은_잡지_않는다(self):
        from pqr.engine.recipe_ointment import blank_qualification_cells
        t = _표()
        self.assertEqual(blank_qualification_cells(t, {"HEA5029": {"IQ": [], "OQ": [], "PQ": []}}), [])
