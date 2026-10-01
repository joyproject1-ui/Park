# -*- coding: utf-8 -*-
"""6항 공 기록서가 글자 있는 PDF 면 읽어 제조단위·포장단위·수율 기준을 얻는다.

담당자 2026-09-30: "포장기록서에 포장단위가 표시되어있는데 본문 포장단위가 사선처리되어 있어 조치해줘."
나조린점안액(1회용)의 포장기록서는 PDF 로 올라와 있고 2쪽에 '포장 단위  0.5mL X 10Strip/Case' 가 적혀
있는데, 프로그램이 "제조·충전·포장 기록서는 .docx 만 읽습니다" 로 지나쳐 6항 포장단위 칸이 사선(빈칸)으로
남았다. 값은 기록서에 적힌 **그대로** 옮긴다 ("기록서 그대로 `0.5mL X 10Strip/Case`").
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import pdftext
from pqr.engine.readers import batch_record as B

# pdftotext -layout 꼴 — 칸 사이가 두 칸 이상 벌어진다 (실제 PK-NJSBEE-1 Rev.008 에서 옮김)
포장기록서 = ["""  지시서 번호  PK-NJSBEE-1
  포장  지시  및  기록서
  개정  번호  008""", """  제 품 명  나조린점안액(1 회용)  지시서 번호  PK-NJSBEE-1
  사 용 기 간  24 개월  지시 포장 수량
  제 조 일 자  사용 기한
  포장 단위  0.5mL X 10Strip/Case
  성 상  플라스틱 용기에 든 무색투명한 액체""", """  8  수율  : ------------------------------------ × 100 = %
  수율  인수량  (  ) Strip
  * 지시 : 수율은 98.0±2.0%로 관리한다."""]

제조기록서 = ["""  제조 지시 및 기록서  MO-NJSE2-1
  제 품 명  나조린점안액(1 회용)
  제조단위  160L
  * 지시 : 수율은 99.0±1.0%로 관리한다."""]


class 이름으로_공기록서를_가린다(unittest.TestCase):
    def test_기록서_지시는_공기록서다(self):
        for name in ("6. PK-NJSBEE-1 나조린점안액(일회용) (Rev.008) [포장 기록서].pdf",
                     "6. EHLF-22 MO-NJSE2-1 나조린점안액(1회용) 제개정이력 [제조 기록서].pdf",
                     "6 제조내역 - PK-NJSBEE-1 나조린점안액(일회용) (Rev.008) [포장 지시].pdf",
                     "6. 퀴노비드점안액(이라크) 포장기록서 rev0.pdf"):
            self.assertTrue(B.looks_like_record(name), name)

    def test_ERP_제조내역은_공기록서가_아니다(self):
        for name in ("6. 제조내역 - ERP.pdf", "6. 2025년 전체 제조 내역.xlsx", "6. 제조내역- ERP.xlsx"):
            self.assertFalse(B.looks_like_record(name), name)


class PDF_공기록서_읽기(unittest.TestCase):
    def setUp(self):
        self._real = pdftext.read_layout
        self.addCleanup(setattr, pdftext, "read_layout", self._real)

    def _read(self, pages, name):
        pdftext.read_layout = lambda path, pages_=None, **kw: pages
        return B.read_pdf(os.path.join("/tmp", name))

    def test_포장기록서에서_포장단위를_그대로_읽는다(self):
        got = self._read(포장기록서, "6. PK-NJSBEE-1 나조린점안액(일회용) (Rev.008) [포장 기록서].pdf")
        self.assertEqual(got["kind"], "포장")
        self.assertEqual(got["pack_unit"], "0.5mL X 10Strip/Case")   # 띄어쓰기·대소문자 그대로
        self.assertEqual(got["yield_spec"], "98.0±2.0%")
        self.assertEqual(got["materials"], [])                       # 글자만으로는 원/자재 열을 믿지 않는다

    def test_제조기록서에서_제조단위를_읽는다(self):
        got = self._read(제조기록서, "6. 제조 MO-NJSE2-1 나조린점안액(1회용).pdf")
        self.assertEqual(got["kind"], "조제")
        self.assertEqual(got["batch_size"], "160L")
        self.assertEqual(got["pack_unit"], "")

    def test_제조단위_칸에_포장이_붙어_있으면_제조단위로_보지_않는다(self):
        got = self._read(["  포장 단위  5g x 1Tube/Case"], "6. 포장기록서.pdf")
        self.assertEqual(got["batch_size"], "")
        self.assertEqual(got["pack_unit"], "5g x 1Tube/Case")

    def test_포장단위_칸이_비면_빈_채로_둔다(self):
        got = self._read(["  포장 단위", "  성 상  무색투명한 액체"], "6. 포장기록서.pdf")
        self.assertEqual(got["pack_unit"], "")


class 합치기(unittest.TestCase):
    def test_포장단위는_포장기록서_제조단위는_제조기록서(self):
        recs = [{"kind": "조제", "batch_size": "160L", "pack_unit": "", "yield_spec": "", "materials": [],
                 "file": "제조.pdf"},
                {"kind": "포장", "batch_size": "", "pack_unit": "0.5mL X 10Strip/Case", "yield_spec": "", "materials": [],
                 "file": "포장.pdf"}]
        got = B.merge(recs)
        self.assertEqual(got["batch_size"], "160L")
        self.assertEqual(got["pack_unit"], "0.5mL X 10Strip/Case")
        self.assertEqual(got["pack_unit_from"], "포장.pdf")
        self.assertEqual(got["fallback"], [])


if __name__ == "__main__":
    unittest.main()
