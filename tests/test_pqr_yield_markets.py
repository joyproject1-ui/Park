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
    ws.merge_cells(start_row=3, start_column=5, end_row=3, end_column=6)   # '포장' 이 두 열에 걸친다
    if merge_spec:
        ws.merge_cells(start_row=5, start_column=5, end_row=5, end_column=6)
    wb.save(path)
    return path


class 병합된_머리행(unittest.TestCase):
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
