# -*- coding: utf-8 -*-
"""담당자가 받는 진짜 자료 꼴로 읽는다 — 2026 디겐타안연고 자료에서 드러난 것들.

전년도 결재본은 장비·원료 같은 '틀' 을 참고하는 것이고, 값은 그해 자료에서 새로 읽어야 한다.
그 자료가 사람이 추린 표가 아니라 ERP·마스터파일이 그대로 내려 준 꼴이라, 못 읽으면 표에
지난해 값이 그대로 남는다 — GMP 기록에서 가장 나쁜 결과다.
"""
import os
import tempfile
import unittest

from openpyxl import Workbook

from pqr.engine.readers import erp, masters, yield_sheet


def _xlsx(rows, sheet="Sheet1"):
    book = Workbook()
    ws = book.active
    ws.title = sheet
    for row in rows:
        ws.append(list(row))
    path = os.path.join(tempfile.mkdtemp(prefix="pqr-real-"), "t.xlsx")
    book.save(path)
    return path


class 수율현황표(unittest.TestCase):
    """제품마다 머리 부분이 다르다 — 자리를 정해 두면 다른 제품에서 값을 통째로 놓친다."""

    회사서식 = [
        ["연번", "중요공정 별 수율 현황 (%)", None, None, None, "비고"],
        [None, "공정", "조제", "충전", "포장", None],
        [None, "기준", "94% 이상", "86.5±6.5%", "98 ± 2%", None],
        [None, "Lot No.", None, None, None, None],
        [1, "OGY301", 96.89, 80.82, 99.5, None],
        [2, "OGY901", 98.84, 83.16, 99.68, None],
        ["최댓값", None, 98.84, 83.16, 99.68, None],
    ]
    단순표 = [
        [None, "조제", "충전", "포장"],
        ["OEY101", 97.45, 87.91, 99.04],
        ["OEY102", 98.09, 89.28, 99.85],
    ]

    def test_회사_서식에서_Lot_과_값을_읽는다(self):
        got = yield_sheet.read_yields(_xlsx(self.회사서식))
        self.assertEqual(got[0], ("OGY301", {"조제": "96.89", "충전": "80.82", "포장": "99.50"}))
        self.assertEqual([lot for lot, _ in got], ["OGY301", "OGY901"])   # 최댓값 줄은 자료가 아니다

    def test_단순한_표도_그대로_읽는다(self):
        got = yield_sheet.read_yields(_xlsx(self.단순표))
        self.assertEqual(got[0], ("OEY101", {"조제": "97.45", "충전": "87.91", "포장": "99.04"}))

    def test_그해_기준을_읽는다(self):
        # 기준은 해마다 개정된다(디겐타안연고 충전: 96.0 ± 3.5% → 86.5 ± 6.5%).
        # 전년도 결재본의 기준으로 견주면 멀쩡한 Lot 이 모두 '기준 벗어남' 으로 잡힌다.
        self.assertEqual(yield_sheet.read_specs(_xlsx(self.회사서식)),
                         {"조제": "94% 이상", "충전": "86.5±6.5%", "포장": "98 ± 2%"})

    def test_기준_줄이_없으면_빈_것을_돌려준다(self):
        self.assertEqual(yield_sheet.read_specs(_xlsx(self.단순표)), {})


class ERP_시험번호표(unittest.TestCase):
    """ERP 가 그대로 내려 준 표는 머리행이 없고 열이 열아홉이며, 그 원료를 쓴 모든 제품이 들어 있다."""

    def _raw(self, *rows):
        out = []
        for code, product, test, lot in rows:
            row = [code, "20250310", "4", "Gentamicin Sulfate", "KP", "G",
                   "DGTO1", product, "DGTO2025030001", "M711", test, "7100",
                   "", "", "", "789.0", lot, "", "Y"]
            out.append(row)
        return _xlsx(out)

    def test_제품_이름으로_이_제품_줄만_가려낸다(self):
        path = self._raw(("RSF101", "디겐타안연고", "R202406130032", "OGY301"),
                         ("RSF101", "후메론점안액0.1%", "R202406130032", "EAY101"),
                         ("RSF101", "", "R202406130032", ""))          # 입고·시험 줄
        self.assertEqual(erp.read_material_tests(path, product="디겐타안연고"),
                         [("RSF101", "R-2024-06-13-0032", "OGY301")])

    def test_Lot_은_표에_있는_것을_쓴다(self):
        # ERP 제조번호(DGTO2025030001)가 아니라 열일곱째 칸의 Lot No. 를 쓴다 — 지어내지 않는다.
        path = self._raw(("RBG201", "디겐타안연고", "R202401190012", "OGY301"))
        self.assertEqual(erp.read_material_tests(path, product="디겐타안연고")[0][2], "OGY301")

    def test_사람이_추린_세_칸_표도_그대로_읽는다(self):
        path = _xlsx([["원료 코드", "시험번호", "Lot No."],
                      ["RBO101", "R202405130070", "OEY101"],
                      ["RBO101", "R202405130070", "OEY102"]])
        self.assertEqual(erp.read_material_tests(path, product="퀴노비드안연고"),
                         [("RBO101", "R-2024-05-13-0070", "OEY101"),
                          ("RBO101", "R-2024-05-13-0070", "OEY102")])


class 적격성_마스터파일(unittest.TestCase):
    """회사가 서식을 개정하면 시트 이름과 열 이름이 바뀐다 — 정해 두면 그 항이 통째로 빈다."""

    def _sheet(self, name, date_header):
        return _xlsx([["( 생산장비 ) Qualification"], [], [],
                      ["No.", "부서", "라인", "관리번호", "설비명", "보고서 문서번호",
                       "Rev No.", date_header],
                      [1, "생산2부", "MAR1", "DAA5114", "MAINMIXER", "PQ24-2-DAA5114-R",
                       "0", "2025.01.06"]], sheet=name)

    def test_시트_이름이_바뀌어도_읽는다(self):
        got = masters.equipment_docs(self._sheet("생산장비", "승인일자"))
        self.assertEqual(got["DAA5114"]["name"], "MAINMIXER")
        self.assertEqual(got["DAA5114"]["docs"], [("PQ24-2-DAA5114-R", "2025.01.06")])

    def test_완료일이_승인일자로_바뀌어도_읽는다(self):
        got = masters.equipment_docs(self._sheet("제조시설적격성평가현황", "완료일"))
        self.assertEqual(got["DAA5114"]["docs"], [("PQ24-2-DAA5114-R", "2025.01.06")])

    def test_제조지원설비가_제조설비_꼴로_적혀_있어도_읽는다(self):
        got = masters.support_docs(self._sheet("생산장비", "승인일자"))
        self.assertEqual(got["DAA5114"]["PQ"], [("PQ24-2-DAA5114-R", "2025.01.06")])


if __name__ == "__main__":
    unittest.main()
