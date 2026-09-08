"""PV 마스터는 라인별 시트로 갈라져 있고 문서 코드에 밑줄이 들어간다
(담당자 2026-09-08: "해당 라인은 충전기 기준 BFS 2호야")."""
import os
import re
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine.readers import masters                                # noqa: E402


def _book(path):
    from openpyxl import Workbook
    wb = Workbook()
    first = wb.active
    first.title = "Mar 1호"
    for r in [["No.", "제품명", "(PV) 계획서", "(PV) 사유", "(PV) 종류", "Lot. No", "제조일자", "(PV) 보고서"],
              ["1", "후메론점안액", "PV18-5-FMRE8_1-P", "신규 제정", "예측적", "1st", "2018.08.20",
               "PV18-5-FMRE8_1-R"]]:
        first.append(r)
    bfs = wb.create_sheet("BFS 2호 ")            # 이름 끝에 빈칸이 있다 — 실제 파일이 그렇다
    for r in [["No.", "제품명", "(PV) 계획서", "(PV) 사유", "(PV) 종류", "Lot. No", "제조일자", "(PV) 보고서"],
              ["38", "아이퓨어점안액", "PV25-5-TRHE1_1-P(Rev.2)", "변경관리", "동시적", "1st", "2025.05.12",
               "PV25-5-TRHE1_1-R-1 (Rev.2)"],
              ["", "", "", "", "", "2nd", "2025.10.16", "PV25-5-TRHE1_1-R-2 (Rev.2)"]]:
        bfs.append(r)
    wb.save(path)
    return path


class 라인별_시트(unittest.TestCase):
    def setUp(self):
        self.path = _book(os.path.join(tempfile.mkdtemp(prefix="pqr-pv-"), "pv.xlsx"))

    def test_첫_시트가_아니어도_찾는다(self):
        got = masters.pv_by_code(self.path, "TRHE1_1")
        self.assertEqual(len(got), 1)
        self.assertEqual(got[0]["report"], "PV25-5-TRHE1_1-R-1 (Rev.2)")

    def test_첫_시트의_제품도_그대로_찾는다(self):
        got = masters.pv_by_code(self.path, "FMRE8_1")
        self.assertEqual(len(got), 1)

    def test_없는_코드는_빈_목록(self):
        self.assertEqual(masters.pv_by_code(self.path, "ZZZZ9"), [])


class 문서_코드_뽑기(unittest.TestCase):
    """PV 번호에 밑줄이 든다 — 밑줄을 빼면 한 건도 못 찾는다."""

    def test_밑줄이_든_코드를_뽑는다(self):
        found = re.findall(r"PV\d{2}-\d-([A-Z0-9_]+)-", "PV20-5-TRHE1_1-P(Rev.1) PV24-2-QUIO3-R")
        self.assertEqual(sorted(found), ["QUIO3", "TRHE1_1"])


if __name__ == "__main__":
    unittest.main()
