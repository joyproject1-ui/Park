"""마스터파일 해석 테스트.

실제 마스터파일은 사내 자료이므로 저장소에 두지 않고, 같은 구조를 흉내 낸
행 목록으로 검증합니다 (엑셀 읽기 없이 동작).
"""

import unittest

from pqr.master import (MasterError, find_header, parse_equipment_sheet,
                        parse_pv_sheet, parse_support_sheet)

# 공정밸리데이션(PV) 마스터 — 제품 한 건에 검증 Lot 3개가 병합 셀로 딸려 있습니다.
PV_ROWS = [
    ["( PV ) Validation Master File", None, None, None, None, None, None, None, None, None],
    [None] * 10,
    ["①", None, ": 밸리데이션 진행 중", None, None, None, None, None, None, None],
    [None] * 10,
    ["No.", "제품명 or 장비명(관리번호)", "(PV) 계획서", "(PV) 사유", "(PV) 종류",
     "Lot. No", None, "제조일자", "(PV) 보고서", "재밸리데이션 연도"],
    [None, None, "문서번호(Rev. No.)", None, None, None, None, None, None, None],
    ["1", "가상점안액", "PV25-P (Rev.2)", "변경관리에 따른 변경", "동시적",
     "1st", "AB1201", "2025.05.12", "PV25-R-1 (Rev.2)", "2031"],
    [None, None, None, None, None, "2nd", "AB1202", "2025.10.16", "PV25-R-2 (Rev.2)", None],
    [None, None, None, None, None, "3rd", "AB1203", "2026.01.28", "PV26-FR (Rev.2)", None],
    ["2", "다른점안액", "PV24-P (Rev.1)", "신규", "예측적",
     "1st", "ZZ0101", "2024.03.03", "PV24-R (Rev.1)", "2030"],
]

# 설비 적격성 마스터 — 설비 하나에 개정 이력이 여러 행 쌓입니다.
EQUIPMENT_ROWS = [
    ["( 생산장비 ) Qualification Master File"] + [None] * 10,
    [None] * 11,
    [None, None, ": 최종 성능 적격성평가"] + [None] * 8,
    [None] * 11,
    ["No.", "부서", "라인", "관리번호", "설비명", "보고서 문서번호", "Rev No.",
     "승인일자", "재적격성평가 연도", "개정 사유", "비고"],
    [None] * 11,
    ["1", "생산1부", "BFS2", "DAA5090", "800L 조제탱크", "PQ22-R", "3",
     "2022.10.06", "N/A", "정기적격성평가", None],
    [None, None, None, None, None, "PQ25-R", "5", "2025.09.29", "2028", "정기적격성평가", None],
    ["2", "생산1부", "MAR1", "DAA5044", "조제탱크(100L)", "PQ23-R", "5",
     "2023.10.06", "2026", "정기적격성평가", None],
    ["3", "생산1부", "BFS2", "DAF5058", "BFS2호 충전기", "PQ25-R", "2",
     "2025.12.22", "2026", "정기적격성평가", None],
]


# 제조지원 설비 표 — 라인 열이 없고, 설비명에 라인이 적혀 있기도 합니다.
SUPPORT_ROWS = [
    ["적격성평가 현황표(제조지원 설비)"] + [None] * 11,
    [None] * 12,
    [None, None, "1. 제조지원설비 적격성평가"] + [None] * 9,
    [None, None, "시스템", "구분", "설비명", "관리번호", None,
     "DQ 문서번호", "DQ 승인일", "사유", "IQ 문서번호", "IQ 승인일"],
    [None, None, "수처리시스템", "제조 시스템", "정제수 제조 시스템 (1층 액제, 안구이식제)",
     "PWG2301", None, "DQ23-R", "2023.04.12", "신규", "IQ23-R", "2023.09.08"],
    [None, None, None, None, "주사용수 제조 시스템 (2t - 2층)", "HJA5037",
     None, "DQ12-R", "2012.06.07", "신규", "IQ12-R", "2012.08.11"],
    [None, None, "공조기 시스템", "공조기 (1층)", "A.H.U 2호 (MAR1 라인, BFS3호 라인)",
     "HBA5031", None, "DQ15-R", "2015.05.01", "신규", "IQ15-R", "2015.07.01"],
    [None, None, None, None, "A.H.U 3호 (BFS2호 라인)", "HBA5032",
     None, "DQ15-R", "2015.05.02", "신규", "IQ15-R", "2015.07.24"],
    [None, None, None, "공조기 (2층)", "A.H.U 13호 (2층 포장실)", "HCA5074",
     None, "DQ16-R", "2016.01.01", "신규", "IQ16-R", "2016.03.01"],
]


class SupportTest(unittest.TestCase):
    def test_line_written_into_the_asset_name_is_matched(self):
        """'A.H.U 3호 (BFS2호 라인)' 처럼 설비명에 라인이 적힌 것을 찾습니다."""
        found = parse_support_sheet(SUPPORT_ROWS, line="BFS2호")
        self.assertEqual([f["asset_id"] for f in found], ["HBA5032"])
        self.assertEqual(found[0]["iq"]["date"], "2015.07.24")

    def test_a_different_line_is_not_matched(self):
        found = parse_support_sheet(SUPPORT_ROWS, line="BFS1호")
        self.assertEqual(found, [])

    def test_floor_filter_when_the_line_is_not_written(self):
        found = parse_support_sheet(SUPPORT_ROWS, floor=1, system="수처리")
        self.assertEqual([f["asset_id"] for f in found], ["PWG2301"])

    def test_explicit_asset_ids_win(self):
        """관리번호를 직접 주면 이름·층과 무관하게 그것만 씁니다."""
        found = parse_support_sheet(SUPPORT_ROWS, asset_ids=["HJA5037", "HBA5032"])
        self.assertEqual(sorted(f["asset_id"] for f in found), ["HBA5032", "HJA5037"])

    def test_system_and_kind_carry_down(self):
        found = parse_support_sheet(SUPPORT_ROWS, asset_ids=["HBA5032"])
        self.assertEqual(found[0]["system"], "공조기 시스템")
        self.assertEqual(found[0]["kind"], "공조기 (1층)")


class HeaderTest(unittest.TestCase):
    def test_finds_header_below_title_rows(self):
        index, cells = find_header(PV_ROWS, ("제품명", "lot"))
        self.assertEqual(index, 4)
        self.assertIn("no.", cells)

    def test_missing_header_raises(self):
        with self.assertRaises(MasterError):
            find_header([["a", "b"], ["c", "d"]], ("제품명", "lot"))


class PvTest(unittest.TestCase):
    def setUp(self):
        self.entries = parse_pv_sheet(PV_ROWS, "가상점안액")

    def test_absent_product_returns_none(self):
        self.assertIsNone(parse_pv_sheet(PV_ROWS, "없는제품"))

    def test_other_products_are_not_included(self):
        self.assertEqual(len(self.entries), 1)
        self.assertEqual(self.entries[0]["plan"], "PV25-P (Rev.2)")

    def test_merged_cells_are_carried_down(self):
        """계획서·사유·종류는 첫 행에만 있고 아래 행은 병합으로 비어 있습니다."""
        entry = self.entries[0]
        self.assertEqual(entry["kind"], "동시적")
        self.assertEqual(entry["revalidation_year"], "2031")

    def test_all_three_validation_lots_are_collected(self):
        lots = self.entries[0]["lots"]
        self.assertEqual([l["lot_no"] for l in lots], ["AB1201", "AB1202", "AB1203"])
        self.assertEqual(lots[0]["mfg_date"], "2025-05-12")
        self.assertEqual(lots[2]["mfg_date"], "2026-01-28")

    def test_each_lot_keeps_its_own_report(self):
        lots = self.entries[0]["lots"]
        self.assertEqual(lots[1]["report"], "PV25-R-2 (Rev.2)")
        self.assertEqual(len(self.entries[0]["reports"]), 3)


class EquipmentTest(unittest.TestCase):
    def test_only_the_requested_line_is_returned(self):
        items = parse_equipment_sheet(EQUIPMENT_ROWS, "BFS2")
        self.assertEqual([i["asset_id"] for i in items], ["DAA5090", "DAF5058"])

    def test_latest_revision_wins(self):
        """같은 설비의 여러 개정 이력 중 승인일이 가장 늦은 것을 씁니다."""
        items = parse_equipment_sheet(EQUIPMENT_ROWS, "BFS2")
        tank = next(i for i in items if i["asset_id"] == "DAA5090")
        self.assertEqual(tank["approved"], "2025-09-29")
        self.assertEqual(tank["revalidation_year"], "2028")

    def test_line_and_asset_are_carried_into_revision_rows(self):
        items = parse_equipment_sheet(EQUIPMENT_ROWS, "BFS2")
        tank = next(i for i in items if i["asset_id"] == "DAA5090")
        self.assertEqual(tank["asset"], "800L 조제탱크")
        self.assertEqual(tank["dept"], "생산1부")

    def test_unknown_line_returns_nothing(self):
        self.assertEqual(parse_equipment_sheet(EQUIPMENT_ROWS, "BFS9"), [])

    def test_no_line_filter_returns_all(self):
        items = parse_equipment_sheet(EQUIPMENT_ROWS, "")
        self.assertEqual(len(items), 3)


if __name__ == "__main__":
    unittest.main()
