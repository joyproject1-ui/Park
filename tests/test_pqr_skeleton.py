"""공양식이 바탕일 때 전년도 결재본의 표 뼈대를 옮겨 심는다 (담당자 PC 2026-09-07 한림포비돈점안액)."""
import os
import sys
import unittest

import docx
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import carry, docedit as E                        # noqa: E402


def _doc(items):
    """[(제목 또는 문단 글, 표 또는 None)] — 글이 빈 문자열이면 빈 문단."""
    d = docx.Document()
    for text, table in items:
        d.add_paragraph(text)
        if table:
            t = d.add_table(rows=len(table), cols=len(table[0]))
            for i, row in enumerate(table):
                for j, v in enumerate(row):
                    E.set_cell(t.rows[i].cells[j], v)
    return d


H91 = ["공 정", "시험항목", "허용기준", "결과"]
CPK = [["Cpk", "판정", "비고"], ["Cpk ≥ 1", "공정능력 충분", "현재 공정 유지"]]


def _form():
    return _doc([("4. 제품품질평가 일정계획", None), ("", None), ("", None),
                 ("5. 책임과 권한", [["직 책", "업 무"], ["작성자", "작성한다"]]),
                 ("9.1 시험결과표", [H91, ["", "", "", ""]]),
                 ("9.2 세부 시험결과", None),
                 ("9.2.1. 조제 완료 후(점안제)", [["연번", "Lot No.", "", ""], ["1", "", "", ""],
                                              ["최댓값", "", "", ""], ["최솟값", "", "", ""], ["평균", "", "", ""]]),
                 ("", None),
                 ("※ 공정능력지수(Cpk) 판정표", CPK),
                 ("10.2 제조설비", [["No.", "관리번호", "설비명", "IQ", "비고"], ["1", "", "", "", "N/A"]]),
                 ("10.3 공기조화장치", [["No.", "관리번호", "설비명", "IQ"], ["1", "", "", ""]]),
                 ("13.3 안정성 시험 경향 분석 결과", [["시험항목", "", ""], ["", "", ""], ["관리 규격", "", ""],
                                                ["특이사항 (Comment)", "", ""]]),
                 ("14. 반품 및 불만 회수관련 기록", None)])


def _prev():
    return _doc([("4. 제품품질평가 일정계획", None),
                 ("제품품질평가는 2024년 1월 ~ 12월까지 생산된 해당제품에 대하여 평가를 실시하며 2025년도 1분기 내에 완료한다.", None),
                 ("", None),
                 ("5. 책임과 권한", [["직 책", "업 무"], ["작성자", "작성한다"]]),
                 ("9.1 시험결과표", [H91, ["조제", "pH", "자가) 6.5 ~ 7.5", "Av. 7.21"]]),
                 ("", None), ("", None),
                 ("", [H91, ["포장", "함량", "허가) 90.0 ~ 110.0%", "Av. 101.0"]]),
                 ("9.2 세부 시험결과", None),
                 ("9.2.1 조제(점안제)", [["연번", "Lot No.", "성상", "pH"], ["1", "ELX101", "무색의 액", "7.38"],
                                      ["2", "ELX501", "무색의 액", "7.21"],
                                      ["최댓값", "", "", "7.38"], ["최솟값", "", "", "7.21"], ["평균", "", "", "7.30"]]),
                 ("", None),
                 ("※ 공정능력지수(Cpk) 판정표", CPK),
                 ("10.2 제조설비", None),
                 ("10.2.1 점안제 충전기 라인", [["No.", "관리번호", "설비명", "IQ", "비고"],
                                            ["1", "DAA5040", "2,000L 조제탱크", "IQ09-5-DAA5040-R", ""]]),
                 ("", None),
                 ("10.2.2 Mar 충전기 2호 라인", [["No.", "관리번호", "설비명", "IQ", "비고"],
                                             ["1", "DAE5016", "고압증기멸균기 5호", "IQ18-5-DAE5016-R", ""]]),
                 ("", None),
                 ("10.3 공기조화장치", [["No.", "관리번호", "설비명", "IQ"], ["1", "AHU-1", "공조기", "IQ-1"]]),
                 ("13.3 안정성 시험 경향 분석 결과", [["시험항목", "pH", "함량 (%)"], ["시판 후 (2024)", "7.40", "99.7"],
                                                ["관리 규격", "6.0 ~ 8.0", "90.0~110.0"],
                                                ["특이사항 (Comment)", "", ""]]),
                 ("14. 반품 및 불만 회수관련 기록", None)])


def _grid(table):
    return [[E.cell_text(c) for c in E.raw_cells(r)] for r in table.rows]


def _body_texts(document):
    out = []
    for el in document.element.body:
        if el.tag == qn("w:p"):
            out.append(("p", "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()))
        elif el.tag == qn("w:tbl"):
            out.append(("t", ""))
    return out


class 뼈대_옮겨_심기(unittest.TestCase):
    def setUp(self):
        self.form, self.prev = _form(), _prev()
        self.moved = carry.adopt_skeleton(self.form, self.prev, {"from": "2025-01-01", "to": "2025-12-31"})
        self.by = carry._tables_by_section(self.form)

    def test_빈_뼈대인_항만_옮긴다(self):
        self.assertEqual(self.moved, ["4", "9.1", "9.2.1", "10.2", "13.3"])   # 5·10.3 은 값이 있어 그대로

    def test_9_1_은_하나로_합치고_결과_열은_비운다(self):
        tables = self.by["9.1"]
        self.assertEqual(len(tables), 1)
        grid = _grid(tables[0])
        self.assertEqual([r[1] for r in grid[1:]], ["pH", "함량"])
        self.assertEqual([r[2] for r in grid[1:]], ["자가) 6.5 ~ 7.5", "허가) 90.0 ~ 110.0%"])
        self.assertEqual([r[3] for r in grid[1:]], ["", ""])                  # 작년 Av. 7.21 은 남기지 않는다

    def test_9_2_는_열_머리글만_남기고_lot_줄은_하나로(self):
        grid = _grid(self.by["9.2.1"][0])
        self.assertEqual(grid[0], ["연번", "Lot No.", "성상", "pH"])
        self.assertEqual(grid[1], ["1", "", "", ""])
        self.assertEqual([r[0] for r in grid[2:]], ["최댓값", "최솟값", "평균"])
        self.assertTrue(all(c == "" for r in grid[2:] for c in r[1:]))

    def test_공양식의_cpk_안내_블록은_남고_하나뿐이다(self):
        texts = _body_texts(self.form)
        stars = [t for k, t in texts if k == "p" and t.startswith("※")]
        self.assertEqual(len(stars), 1)
        cpk = [t for t in self.form.tables if _grid(t)[0][0] == "Cpk"]
        self.assertEqual(len(cpk), 1)

    def test_10_2_는_라인_소제목과_설비_목록이_들어온다(self):
        texts = [t for k, t in _body_texts(self.form) if k == "p"]
        self.assertIn("10.2.1 점안제 충전기 라인", texts)
        self.assertIn("10.2.2 Mar 충전기 2호 라인", texts)
        # 옮겨 온 표는 소제목(10.2.1·10.2.2) 아래에 놓인다 — 엔진의 _tables(document, "10.2") 는 앞글자로 찾는다
        mids = [_grid(t)[1][1] for k in sorted(self.by) if k.startswith("10.2.") for t in self.by[k]]
        self.assertEqual(mids, ["DAA5040", "DAE5016"])
        self.assertEqual(_grid(self.by["10.3"][0])[1][1], "")               # 값 없던 10.3 은 건드리지 않는다

    def test_13_3_은_시험항목_머리글이_들어온다(self):
        self.assertEqual(_grid(self.by["13.3"][0])[0], ["시험항목", "pH", "함량 (%)"])

    def test_4항_문안은_연도를_올해로(self):
        texts = [t for k, t in _body_texts(self.form) if k == "p"]
        hit = [t for t in texts if t.startswith("제품품질평가는")]
        self.assertEqual(len(hit), 1)
        self.assertIn("2025년 1월 ~ 12월", hit[0])
        self.assertIn("2026년도 1분기", hit[0])

    def test_두_번_불러도_더_옮기지_않는다(self):
        self.assertEqual(carry.adopt_skeleton(self.form, _prev()), [])

    def test_다른_문서의_스타일_이름은_떼어_낸다(self):
        for el in self.form.element.body.iter(qn("w:pStyle")):
            self.assertIn(el.get(qn("w:val")), {s.style_id for s in self.form.styles})


class 연도_옮기기(unittest.TestCase):
    def test_평가_연도에_맞춰_모든_연도를_같이_민다(self):
        got = carry._shift_years("2024년 1월 ~ 12월 … 2025년도 1분기", {"from": "2025-01-01"})
        self.assertEqual(got, "2025년 1월 ~ 12월 … 2026년도 1분기")

    def test_기간이_없으면_그대로(self):
        self.assertEqual(carry._shift_years("2024년", None), "2024년")


if __name__ == "__main__":
    unittest.main()
