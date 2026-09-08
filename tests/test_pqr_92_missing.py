# -*- coding: utf-8 -*-
"""9.2 표에서 통째로 빈 열은 문의 목록에 남긴다.

담당자 2026-09-08: "값을 왜 입력하지 못해? 오른쪽 파일로 작성하면 돼 뭐가 문제지?
PQR 문의목록에도 해당내용이 없네" — 못 채운 칸은 공양식의 사선이 그대로 남아 말없이 비어 있었다.
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import docx                                                            # noqa: E402

from pqr.engine import detail92 as D, docedit as E                     # noqa: E402


def _table():
    d = docx.Document()
    t = d.add_table(rows=6, cols=5)
    rows = [["연번", "Lot No.", "성상", "제제균일성(%)", "비고"],
            ["1", "LWY201", "", "", ""],
            ["2", "LWY501", "", "", ""],
            ["최댓값", "", "", "", ""],
            ["최솟값", "", "", "", ""],
            ["평균", "", "", "", ""]]
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            E.set_cell(t.rows[i].cells[j], v)
    return t


class 못_채운_열을_알린다(unittest.TestCase):
    def test_한_Lot도_값이_없으면_알린다(self):
        got = []
        D.fill(_table(), ["LWY201", "LWY501"],
               lambda lab, lot, i: "무색투명한 액" if "성상" in lab else None,
               miss=got.append)
        self.assertEqual(got, [["제제균일성(%)"]])          # 성상은 채웠고, 연번·LotNo·비고는 뺀다

    def test_한_칸이라도_채우면_알리지_않는다(self):
        got = []
        D.fill(_table(), ["LWY201", "LWY501"],
               lambda lab, lot, i: "102.4" if "제제균일성" in lab or "성상" in lab else None,
               miss=got.append)
        self.assertEqual(got, [])

    def test_miss를_주지_않아도_그대로_동작한다(self):
        t = _table()
        D.fill(t, ["LWY201"], lambda lab, lot, i: "적합" if "성상" in lab else None)
        self.assertEqual(E.cell_text(E.raw_cells(t.rows[1])[2]).strip(), "적합")


class Claude가_읽은_시험항목(unittest.TestCase):
    """스캔 성적서를 Claude 로 읽을 때 함량만 받아 오면 9.2 표의 그 열이 통째로 빈다."""

    def test_items를_readers_coa_와_같은_꼴로_바꾼다(self):
        from pqr.engine.claude_cli import _coa_items
        got = _coa_items({"items": [{"name": "제제균일성", "spec": "85.0 ~ 115.0%", "value": "102.4"},
                                    {"name": "확인시험", "spec": "", "value": "적합"},
                                    {"name": "제제균일성", "spec": "", "value": "99.9"},   # 같은 이름은 하나만
                                    {"name": "", "spec": "", "value": "x"}]})
        self.assertEqual(got["제제균일성"], {"spec": "85.0 ~ 115.0%", "value": "102.4"})
        self.assertEqual(got["확인시험"]["value"], "적합")
        self.assertEqual(len(got), 2)

    def test_items가_없어도_빈_사전(self):
        from pqr.engine.claude_cli import _coa_items
        self.assertEqual(_coa_items({}), {})


if __name__ == "__main__":
    unittest.main()
