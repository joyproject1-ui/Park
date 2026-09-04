# -*- coding: utf-8 -*-
"""9.1·9.2 를 자리가 아니라 **머리글과 허용기준**으로 짚는다 (2026 최신 양식).

담당자 지시(2026-09): "올해부터 오른쪽 docx 최신양식으로 변경됐어."
2026 양식은 9.2 가 공정별(조제·충전·포장) 표로 나뉘고 함량이 성분마다 한 열씩이라, 예전처럼
열 자리를 정해 두면 입자도 값이 함량 칸에 들어간다 — 실제로 그렇게 나왔다.
"""
import unittest

import docx
from docx.oxml.ns import qn
from docx.table import _Cell

from pqr.engine import detail92 as D


def _table(document, rows):
    """rows: [[(글, gridSpan), …], …] — 첫 줄이 머리행."""
    width = sum(span for _, span in rows[0])
    table = document.add_table(rows=0, cols=width)
    for cells in rows:
        tr = table._tbl.add_tr()
        for text, span in cells:
            tc = tr.add_tc()
            if span > 1:
                pr = tc.get_or_add_tcPr()
                el = pr.makeelement(qn("w:gridSpan"), {qn("w:val"): str(span)})
                pr.append(el)
            _Cell(tc, table).text = text
    return table


class 열_이름으로_짚기(unittest.TestCase):
    def setUp(self):
        self.doc = docx.Document()
        self.table = _table(self.doc, [
            [("연번", 1), ("Lot No.", 1), ("함량(%)", 2), ("입자도(㎛ 이하)", 1)],
            [("", 1), ("", 1), ("플루오로메톨론", 1), ("겐타마이신황산염", 1), ("", 1)],
            [("1", 1), ("", 1), ("", 1), ("", 1), ("", 1)],
            [("최댓값", 2), ("", 1), ("", 1), ("", 1)],
            [("최솟값", 2), ("", 1), ("", 1), ("", 1)],
            [("평균", 2), ("", 1), ("", 1), ("", 1)]])

    def test_머리행_두_줄을_이어_붙여_열_이름을_만든다(self):
        self.assertEqual(D.labels(self.table),
                         ["연번", "LotNo.", "함량(%)플루오로메톨론", "함량(%)겐타마이신황산염", "입자도(㎛이하)"])

    def test_자료_행과_요약_행을_가른다(self):
        first, last, summary = D.data_range(self.table)
        self.assertEqual((first, last), (2, 2))
        self.assertEqual(summary, [3, 4, 5])

    def test_값은_열_이름으로_들어가고_요약은_그_열에서_나온다(self):
        want = {"함량(%)플루오로메톨론": ["97.7", "98.2", "100.2"],
                "함량(%)겐타마이신황산염": ["99.3", "99.5", "100.5"],
                "입자도(㎛이하)": ["15", "13", "35"]}
        D.fill(self.table, ["OGY301", "OGY901", "OGY902"],
               lambda label, lot, i: want[label][i] if label in want else None)
        rows = [[c.text for c in row.cells] for row in self.table.rows]
        self.assertEqual(rows[2][:5], ["1", "OGY301", "97.7", "99.3", "15"])
        self.assertEqual(rows[4][:5], ["3", "OGY902", "100.2", "100.5", "35"])
        self.assertEqual(rows[5][2:5], ["100.2", "100.5", "35"])      # 최댓값
        # 입자도는 '이하' 항목이라 최댓값만 적는다 — 한림 2026 결재본 그대로
        self.assertEqual(rows[6][2:5], ["97.7", "99.3", ""])          # 최솟값
        self.assertEqual(rows[7][2:5], ["98.7", "99.8", ""])          # 평균


class 한쪽만_뜻이_있는_값(unittest.TestCase):
    """'4.58 이상' 은 최솟값만, '이하' 열은 최댓값만 — 한림 2026 결재본이 그렇게 쓴다."""

    def test_이상은_최솟값만(self):
        top, bottom, mean = D._stats("질량용량(g)개개", ["4.58 이상", "4.11 이상", "4.06 이상"])
        self.assertEqual((top, bottom, mean), (None, "4.06", None))

    def test_이하_열은_최댓값만(self):
        top, bottom, mean = D._stats("입자도(㎛이하)", ["15", "13", "35"])
        self.assertEqual((top, bottom, mean), ("35", None, None))

    def test_범위는_평균이_없다(self):
        top, bottom, mean = D._stats("질량용량(g)개개", ["4.0 ~ 4.2", "4.1 ~ 4.2", "4.0 ~ 4.2"])
        self.assertEqual((top, bottom, mean), ("4.2", "4.0", None))


class 허용기준에서_결과_문구(unittest.TestCase):
    def test_자가_허가_와_성분_표시를_뗀다(self):
        self.assertEqual(D.criterion_text("자가) 튜브개봉 시 튜브 막힘이 없음"),
                         "튜브개봉 시 튜브 막힘이 없음")
        self.assertEqual(D.criterion_text("[플루오로메톨론]허가) 검액은 표준액과 동일한 주피크\n유지시간을 나타낸다."),
                         "검액은 표준액과 동일한 주피크 유지시간을 나타낸다.")

    def test_해야_한다는_함_으로_바꾼다(self):
        self.assertEqual(D.criterion_text("자가) 인쇄상태가 양호하며 압인상태가 명확히 식별 가능해야 한다."),
                         "인쇄상태가 양호하며 압인상태가 명확히 식별 가능함")


if __name__ == "__main__":
    unittest.main()
