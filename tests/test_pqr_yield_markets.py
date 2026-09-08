"""수율현황표의 포장이 시장별로 갈린 표 (담당자 2026-09-07~08:
"ELYN01 과 ELYN02 는 내수가 아니고 베트남 포장했네")."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine.readers import yield_sheet as Y                       # noqa: E402


def _sheet(path, merge_spec=False):
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active
    rows = [[None] * 7,
            [None, None, "중요공정 별 수율 현황(%)", None, None, None, None],
            [None, "공정", "조제", "충전", "포장", None, "비고"],
            ["연번", "기준", "99.5±0.48%", "97.0±3.0%", "내수", "베트남(한비돈점안액)", None],
            [None, "Lot No.", None, None, "98.0 ± 2.0%", None, None],
            [1, "ELYO01", 99.97, 97.7, 99.92, None, None],
            [2, "ELYN01", 99.97, 98.52, None, 99.75, None]]
    for r in rows:
        ws.append(r)
    ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=6)   # 표 제목이 값 열 전체에 걸친다
    ws.merge_cells(start_row=3, start_column=5, end_row=3, end_column=6)   # '포장' 이 두 열에 걸친다
    if merge_spec:
        ws.merge_cells(start_row=5, start_column=5, end_row=5, end_column=6)
    wb.save(path)
    return path


class 병합된_머리행(unittest.TestCase):
    """머리가 여러 줄이고 칸이 병합된 표 — 표 제목 줄은 이름에서 뺀다."""

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="pqr-yield-")

    def test_포장이_시장별로_갈리면_이름을_나눠_읽는다(self):
        path = _sheet(os.path.join(self.dir, "y.xlsx"))
        got = dict(Y.read_yields(path))
        self.assertEqual(got["ELYO01"]["포장(내수)"], "99.92")
        self.assertEqual(got["ELYN01"]["포장(베트남)"], "99.75")     # 예전에는 통째로 버려졌다
        self.assertNotIn("포장(내수)", got["ELYN01"])                 # 그 시장으로 포장하지 않았다

    def test_기준이_하나면_두_시장이_나눠_쓴다(self):
        path = _sheet(os.path.join(self.dir, "y2.xlsx"))
        specs = Y.read_specs(path)
        self.assertEqual(specs["포장(내수)"], "98.0 ± 2.0%")
        self.assertEqual(specs["포장(베트남)"], "98.0 ± 2.0%")
        self.assertEqual(specs["조제"], "99.5±0.48%")

    def test_병합된_기준_줄도_읽는다(self):
        path = _sheet(os.path.join(self.dir, "y3.xlsx"), merge_spec=True)
        specs = Y.read_specs(path)
        self.assertEqual(specs["포장(베트남)"], "98.0 ± 2.0%")


if __name__ == "__main__":
    unittest.main()


class 머리가_두_줄인_표(unittest.TestCase):
    """아이퓨어 수율현황표 — '포장' 아래에 '내수'·'온누리에이치엔씨' 가 따로 있다
    (담당자 2026-09-08: 세 Lot 모두 '포장 수율 값이 수율현황표에 없음')."""

    def _sheet(self, path):
        from openpyxl import Workbook
        wb = Workbook(); ws = wb.active
        for r in [[None, "조제", "충전", "포장", "포장"],
                  [None, "조제", "충전", "내수", "온누리에이치엔씨"],
                  [None, "99.5 ± 0.5%", "92.0 ± 8.0%", "98.0 ± 2.0%", "98.0 ± 2.0%"],
                  ["LWY201", 99.97, 93.47, None, 99.12],
                  ["LWY501", 99.93, 94.93, 99.41, None]]:
            ws.append(r)
        wb.save(path)
        return path

    def test_두_줄을_이어_포장_내수로_읽는다(self):
        path = self._sheet(os.path.join(tempfile.mkdtemp(prefix="pqr-y2-"), "y.xlsx"))
        got = dict(Y.read_yields(path))
        self.assertEqual(got["LWY201"]["포장(온누리에이치엔씨)"], "99.12")
        self.assertEqual(got["LWY501"]["포장(내수)"], "99.41")
        self.assertEqual(got["LWY201"]["조제"], "99.97")     # 한 줄뿐인 공정은 그대로
        specs = Y.read_specs(path)
        self.assertEqual(specs["포장(내수)"], "98.0 ± 2.0%")
        self.assertEqual(specs["충전"], "92.0 ± 8.0%")


class 보고서_표의_두_줄_머리(unittest.TestCase):
    """담당자 2026-09-08: "온누리 값을 잘못 넣었어" — 표 머리가 '포장' 아래 '내수·온누리' 두 칸인데
    한 줄만 읽어 열이 셋으로 잡히고 값이 한 칸 밀렸다."""

    def _table(self):
        import docx
        from pqr.engine import docedit as E
        d = docx.Document()
        t = d.add_table(rows=4, cols=7)
        rows = [["연번", "중요공정 별 수율 현황 (%)", "", "", "", "", "비고"],
                ["", "공정", "조제", "충전", "포장", "", ""],
                ["", "기준\nLot No.", "99.5 ± 0.5%", "92.0 ± 8.0%", "내수", "온누리\n에이치엔씨", ""],
                ["1", "LWY201", "", "", "", "", ""]]
        for i, r in enumerate(rows):
            for j, v in enumerate(r):
                E.set_cell(t.rows[i].cells[j], v)
        E.merge_right(E.raw_cells(t.rows[1])[4], E.raw_cells(t.rows[1])[5])   # '포장' 이 두 열에 걸친다
        return t

    def test_걸친_머리_칸을_두_열에_모두_편다(self):
        from pqr.engine.recipe_ointment import _yield_columns
        self.assertEqual(_yield_columns(self._table()),
                         [(2, "조제"), (3, "충전"), (4, "포장(내수)"), (5, "포장(온누리에이치엔씨)")])

    def test_시장이_적힌_열은_그_시장_값만_쓴다(self):
        from pqr.engine.recipe_ointment import _yield_value
        vals = {"조제": "99.97", "포장(온누리에이치엔씨)": "99.12"}
        self.assertIsNone(_yield_value(vals, "포장(내수)"))       # 예전에는 99.12 가 내수로 들어갔다
        self.assertEqual(_yield_value(vals, "포장(온누리에이치엔씨)"), "99.12")
        self.assertEqual(_yield_value(vals, "조제"), "99.97")

    def test_시장이_없는_열은_괄호를_뗀_이름으로도_찾는다(self):
        from pqr.engine.recipe_ointment import _yield_value
        self.assertEqual(_yield_value({"포장(내수)": "99.41"}, "포장"), "99.41")

    def test_그_시장으로_포장하지_않은_칸은_사선(self):
        from pqr.engine.recipe_ointment import _yield_columns, _yield_value, _other_market
        cols = _yield_columns(self._table())
        vals = [_yield_value({"포장(온누리에이치엔씨)": "99.12"}, n) for _, n in cols]
        self.assertTrue(_other_market(cols, vals, 2))            # 내수 칸 — 사선
        self.assertFalse(_other_market(cols, vals, 3))           # 온누리 칸 — 값이 있다
