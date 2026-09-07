"""13항 — 올해 시험일지를 못 읽은 Lot 은 전년도 결재본·서식 각주에서 옮긴다; 쪽 나눔으로 갈라진 표 잇기 (2026-09-06)."""
import os
import sys
import unittest

import docx

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import carry, docedit as E, handwriting            # noqa: E402
from pqr.engine.recipe_ointment import _fill_stability26, _declared_lots, _tables   # noqa: E402


def _doc(headings_tables):
    d = docx.Document()
    for heading, table in headings_tables:
        d.add_paragraph(heading)
        if table:
            t = d.add_table(rows=len(table), cols=len(table[0]))
            for i, row in enumerate(table):
                for j, v in enumerate(row):
                    E.set_cell(t.rows[i].cells[j], v)
    return d


STORE = "25±2°C,\n60±5%RH"
HEAD_L = ["연번", "해당 연도", "시험 기간", "제조 번호", "포장 형태", "보관 조건", "완료 일자", "실시 사유"]
HEAD_P = ["연번", "해당 연도", "시험 기간", "제조 번호", "포장 형태", "보관 조건", "완료 일자", "비고"]


def _prev():
    """2025 결재본 꼴 — 13.1 시판 후, 13.2 장기(시점마다 한 줄), 13.3 경향."""
    d = _doc([("13. 안정성 시험", None), ("13.1. 시판 후 안정성 시험", None),
              ("13.1.1. 내수용", [HEAD_P, ["1", "2021", "36M", "OEU101", "5g tube/갑", STORE, "2024.03.18", "완료"],
                                ["2", "2022", "24M", "OEV301", "5g tube/갑", STORE, "2024.05.07", "진행중"]]),
              ("13.1.2. 수출용 (베트남)", [HEAD_P, ["1", "2023", "12M", "OZW101", "3.5g tube/갑", STORE, "2024.01.31", "진행중"]]),
              ("13.2. 장기 안정성 시험", None),
              ("13.2.1. 내수용", [HEAD_L, ["1", "2024", "3M", "OEX101", "5g tube/갑", STORE, "2024.05.07", "제품 멸균 조건 변경"],
                                ["", "", "6M", "", "", "", "2024.07.24", ""],
                                ["2", "2024", "Initial", "OEXO01", "5g tube/갑", STORE, "2024.11.05", "주원료 멸균 조건 변경"]]),
              ("13.3. 안정성 시험 경향 분석 결과", None),
              ("13.3.1. 내수용", [["시험항목", "시험항목", "함량(%)"], ["시판 후", "2021", "101.7 ~ 107.4"],
                                ["시판 후", "2022", "97.5 ~ 100.5"], ["장기", "20241)", "103.4 ~ 107.5"],
                                ["장기", "20242)", "105.3"], ["관리 규격", "", "90.0~110.0"], ["최소", "", "97.5"],
                                ["최대", "", "107.5"], ["경향분석 결과", "", "적합"], ["특이사항 (Comment)\n1) OEX101 2) OEXO01", "", ""]])])
    # 6M 줄의 세로 병합 — 제조번호·연도·실시 사유는 위 칸에 이어진다
    t = d.tables[2]
    for j in (0, 1, 3, 4, 5, 7):
        E.set_vmerge(t.rows[1].cells[j], "restart"); E.set_vmerge(t.rows[2].cells[j], None)
    return d


class 전년도_실시내역(unittest.TestCase):
    def test_lot_마다_하나로_읽고_이어지는_시험을_가린다(self):
        got = {(e["kind"], e["market"], e["lot"]): e for e in carry.stability_entries(_prev())}
        self.assertEqual(sorted(got), [("시판후", "내수", "OEU101"), ("시판후", "내수", "OEV301"), ("시판후", "수출", "OZW101"),
                                       ("장기", "내수", "OEX101"), ("장기", "내수", "OEXO01")])
        self.assertFalse(got[("시판후", "내수", "OEU101")]["ongoing"])       # 완료
        self.assertTrue(got[("시판후", "내수", "OEV301")]["ongoing"])
        ex = got[("장기", "내수", "OEX101")]
        self.assertEqual(ex["periods"], ["3M", "6M"])                          # 세로 병합 줄을 합쳤다
        self.assertEqual(ex["dones"], ["2024.05.07", "2024.07.24"])
        self.assertEqual(ex["last"], "제품 멸균 조건 변경")
        self.assertTrue(ex["ongoing"])
        self.assertEqual(ex["range"], {"함량(%)": "103.4 ~ 107.5"})           # 13.3 의 같은 차례 줄
        self.assertEqual(got[("장기", "내수", "OEXO01")]["range"], {"함량(%)": "105.3"})
        self.assertEqual(got[("시판후", "내수", "OEV301")]["range"], {"함량(%)": "97.5 ~ 100.5"})


def _form():
    return _doc([("13. 안정성시험", None), ("13.1 장기 안정성 시험", None),
                 ("13.1.1 내수용", [HEAD_L, ["1", "", "", "", "", "", "", ""], ["2", "", "", "", "", "", "", ""],
                                  ["특이사항 (Comment)\nN/A", "", "", "", "", "", "", ""]]),
                 ("13.2 시판 후 안정성 시험", None),
                 ("13.2.1 내수용", [HEAD_P, ["1", "", "", "", "", "", "", ""], ["특이사항 (Comment)\nN/A", "", "", "", "", "", "", ""]]),
                 ("13.3 안정성 시험 경향 분석 결과", None),
                 ("13.3.1 내수용", [["시험항목", "시험항목", "함량(%)"], ["장기", "20241)", ""], ["장기", "20242)", ""],
                                  ["장기", "20253)", ""], ["시판후", "2022", ""], ["관리 규격", "", ""], ["최소", "", ""],
                                  ["최대", "", ""], ["경향분석 결과", "", ""],
                                  ["특이사항 (Comment)\n- 1) OEX101  2) OEXO01  3) OEY301\n- 경향을 분석하였음.", "", ""]])])


class 서식_각주(unittest.TestCase):
    def test_각주의_lot_과_연도_표시를_읽는다(self):
        got = _declared_lots(_tables(_form(), "13.3")[0])
        self.assertEqual(got["lots"], [("1)", "OEX101"), ("2)", "OEXO01"), ("3)", "OEY301")])
        self.assertEqual(got["year"], {"1)": "2024", "2)": "2024", "3)": "2025"})
        self.assertEqual(got["label"]["3)"], "장기")


class 옮겨_채우기(unittest.TestCase):
    def _run(self, logs):
        d = _form()
        prev = carry.stability_entries(_prev())
        issues, log = [], []
        _fill_stability26(d, logs, {"from": "2025-01-01", "to": "2025-12-31"}, {"오플록사신": "90.0 ~ 110.0%"},
                          log.append, issues, {}, carry.stability_packs(_prev()), prev, "PQR25.docx")
        return d, issues

    def test_시험일지가_없는_장기_lot_은_전년도와_각주에서_세우고_노랑으로_남긴다(self):
        logs = [{"lot": "OEV301", "year": "2022", "kind": "시판후", "market": "내수", "market_hint": True,
                 "pack": "5g tube/갑", "store": STORE, "mfg": "2022.03.18", "expiry": "2025.03.17",
                 "points": [{"period": "36M", "done": "2025.03.10", "assays": {"오플록사신": 99.1}}]}]
        d, issues = self._run(logs)
        long = _tables(d, "13.1.1")[0]
        rows = [[E.cell_text(c) for c in E.raw_cells(r)] for r in long.rows]
        self.assertEqual([r[3] for r in rows[1:4]], ["OEX101", "OEXO01", "OEY301"])
        self.assertEqual([r[1] for r in rows[1:4]], ["2024", "2024", "2025"])
        self.assertEqual(rows[1][2], "확인 필요"); self.assertEqual(rows[1][6], "확인 필요")
        self.assertEqual(rows[1][7], "제품 멸균 조건 변경")                    # 실시 사유는 전년도 것
        self.assertEqual(rows[3][4], "5g tube/갑")                                # 각주 Lot 의 포장은 그 시장 전년도 표기
        self.assertTrue(E.cell_text(E.raw_cells(long.rows[-1])[0]).startswith("특이사항"))
        self.assertIn("옮긴 것입니다", E.cell_text(E.raw_cells(long.rows[-1])[0]))
        self.assertEqual(len(E.raw_cells(long.rows[-1])), 1)
        # 시판 후는 올해 읽은 값 그대로
        post = _tables(d, "13.2.1")[0]
        self.assertEqual([E.cell_text(c) for c in E.raw_cells(post.rows[1])][3:4], ["OEV301"])
        # 13.3 — 전년도 범위는 노랑, 각주 Lot 은 '확인 필요', 번호는 각주대로, 장기 줄이 먼저(서식 차례)
        trend = _tables(d, "13.3")[0]
        body = [[E.cell_text(c) for c in E.raw_cells(r)] for r in trend.rows[1:5]]
        self.assertEqual([b[0] for b in body], ["장기", "", "", "시판 후"])
        self.assertEqual([b[1] for b in body], ["20241)", "20242)", "20253)", "2022"])
        self.assertEqual([b[2] for b in body], ["103.4 ~ 107.5", "105.3", "확인 필요", "99.1"])
        from docx.oxml.ns import qn
        self.assertTrue(list(E.raw_cells(trend.rows[1])[2]._tc.iter(qn("w:highlight"))))      # 전년도 범위는 노랑
        self.assertEqual(len([i for i in issues if i[0] == "13" and "옮겼고" in i[2] and "OEY301" in i[2]]), 1)

    def test_올해_읽은_lot_은_옮기지_않는다(self):
        logs = [{"lot": "OEX101", "year": "2024", "kind": "장기", "market": "내수", "market_hint": True,
                 "pack": "5g tube/갑", "store": STORE, "why": "제품 멸균 조건 변경",
                 "points": [{"period": "12M", "done": "2025.05.02", "assays": {"오플록사신": 104.0}}]}]
        d, issues = self._run(logs)
        long = _tables(d, "13.1.1")[0]
        rows = [[E.cell_text(c) for c in E.raw_cells(r)] for r in long.rows]
        self.assertEqual([r[3] for r in rows[1:4]], ["OEX101", "OEXO01", "OEY301"])
        self.assertEqual(rows[1][2], "12M"); self.assertEqual(rows[1][6], "2025.05.02")
        self.assertEqual(rows[2][2], "확인 필요")


class 이어진_표(unittest.TestCase):
    def _split(self):
        return _doc([("11.1 중요 일탈 및 기준일탈 내역", [["연번", "Lot No.", "CAPA"], ["1", "", "☐ Yes"], ["2", "", "☐ Yes"]]),
                     ("", [["연번", "Lot No.", "CAPA"], ["3", "", "☐ Yes"], ["특이사항 (Comment)\n-", "", ""]]),
                     ("11.2 경향일탈 내역", [["연번", "Lot No.", "CAPA"], ["", "", ""], ["특이사항 (Comment)", "", ""]])])

    def test_같은_머리행으로_시작하는_뒤_표를_앞_표에_잇는다(self):
        d = self._split()
        kept = E.join_continuations([d.tables[0], d.tables[1]])
        self.assertEqual(len(kept), 1)
        self.assertEqual(len(d.tables), 2)                      # 11.1 하나 + 11.2
        first = [E.cell_text(E.raw_cells(r)[0]) for r in d.tables[0].rows]
        self.assertEqual(first, ["연번", "1", "2", "3", "특이사항 (Comment)\n-"])

    def test_다른_표는_잇지_않는다(self):
        d = self._split()
        kept = E.join_continuations([d.tables[0], d.tables[2]])
        self.assertEqual(len(kept), 2)
        self.assertEqual(len(d.tables), 3)

    def test_특이사항_칸은_갈라져_있으면_한_칸으로_합친다(self):
        d = self._split()
        t = d.tables[0]                                          # 마지막 줄이 세 칸짜리 데이터 줄
        cell = E.comment_cell(t)
        self.assertEqual(len(E.raw_cells(t.rows[-1])), 1)
        E.set_cell_plain(cell, "특이사항 (Comment)", "없음")
        self.assertEqual(E.cell_text(E.raw_cells(t.rows[-1])[0]), "특이사항 (Comment)\n없음")
        self.assertEqual(E.grid_width(t), 3)


if __name__ == "__main__":
    unittest.main()


class 같은_lot_합치기(unittest.TestCase):
    """판독 파일에 이미 있는 Lot 을 다시 읽어도 두 줄이 되지 않는다 (담당자 2026-09-07: 13 폴더에 일지를 모두 올림)."""

    def _one(self, lot, kind, periods, sure=True):
        return {"lot": lot, "kind": kind, "kind_sure": sure, "market": "내수", "pack": "", "store": "", "year": "2023",
                "points": [{"period": p, "done": "2025.03.0%d" % i, "assays": {}} for i, p in enumerate(periods, 1)]}

    def test_같은_lot_은_시점만_보탠다(self):
        old = [self._one("OEV301", "시판후", ["12M", "24M"])]
        added = handwriting.merge_logs(old, [self._one("OEV301", "시판후", ["24M", "36M"])])
        self.assertEqual(added, 0)
        self.assertEqual(len(old), 1)
        self.assertEqual([p["period"] for p in old[0]["points"]], ["12M", "24M", "36M"])

    def test_구분을_어림한_일지는_이미_읽은_구분을_따른다(self):
        old = [self._one("OEV301", "시판후", ["12M"])]
        handwriting.merge_logs(old, [self._one("OEV301", "장기", ["24M"], sure=False)])
        self.assertEqual(len(old), 1)
        self.assertEqual(old[0]["kind"], "시판후")

    def test_다른_lot_은_새_줄로(self):
        old = [self._one("OEV301", "시판후", ["12M"])]
        added = handwriting.merge_logs(old, [self._one("OEX101", "장기", ["3M"])])
        self.assertEqual((added, len(old)), (1, 2))


class 사람이_만든_판독_파일(unittest.TestCase):
    """covers_all 이 적힌 판독 파일이면 PDF 를 다시 읽지 않는다 (담당자 2026-09-07: API 키 없이 Claude 로 판독)."""

    def _folder(self, payload):
        import json
        import tempfile
        folder = tempfile.mkdtemp(prefix="pqr-13-")
        with open(os.path.join(folder, "13. 안정성시험일지 판독.json"), "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False)
        return folder

    def _collect(self, folder):
        from pqr.engine import collect
        return collect.collect(folder, product_name="퀴노비드안연고", log=lambda *a: None)

    def test_covers_all_이면_더_읽지_않는다(self):
        one = {"lot": "OEX101", "kind": "장기", "market": "내수", "year": "2024",
               "points": [{"period": "12M", "done": "2025.03.10", "assays": {"오플록사신": 101.2}, "unsure": []}]}
        data = self._collect(self._folder({"logs": [one], "reader_version": 3, "covers_all": True}))
        self.assertTrue(data.stability_all_read)
        self.assertEqual([r["lot"] for r in data.stability_logs], ["OEX101"])

    def test_covers_all_이_없으면_평소대로(self):
        one = {"lot": "OEX101", "kind": "장기", "points": [], "source": "a.pdf"}
        data = self._collect(self._folder({"logs": [one], "reader_version": 3}))
        self.assertFalse(data.stability_all_read)


class 같은_일지_한_번만(unittest.TestCase):
    """이름만 다른 같은 시험일지는 한 번만 읽는다 (담당자 PC 2026-09-07: 24장 가운데 여덟 가지뿐)."""

    def setUp(self):
        import tempfile
        self.dir = tempfile.mkdtemp(prefix="pqr-same-")

    def _pdf(self, name, body):
        path = os.path.join(self.dir, name)
        with open(path, "wb") as fh:
            fh.write(body)
        return path

    def test_내용이_같으면_하나만(self):
        from pqr.engine import collect
        a = self._pdf("시험일지 OEV301.pdf", b"%PDF-1.4 same")
        b = self._pdf("[OG-22-1]_QEV301_36M.pdf", b"%PDF-1.4 same")
        c = self._pdf("다른 일지.pdf", b"%PDF-1.4 other")
        self.assertEqual(collect.same_files_once([a, b, c]), [a, c])

    def test_없는_파일도_버리지_않는다(self):
        from pqr.engine import collect
        self.assertEqual(collect.same_files_once(["/없는/파일.pdf"]), ["/없는/파일.pdf"])


class 구분이_없는_옛_기록(unittest.TestCase):
    """kind 가 없는 판독 기록이 섞여도 넘어지지 않는다 (담당자 PC 2026-09-07 멈춤)."""

    def test_kind_가_없어도_작성이_이어진다(self):
        import json
        import tempfile
        from pqr.engine import collect
        folder = tempfile.mkdtemp(prefix="pqr-nokind-")
        logs = [{"lot": "OEV301", "points": [{"period": "12M", "done": "2025.03.10",
                                              "assays": {"오플록사신": 101.2}, "unsure": []}]},
                {"lot": "OEV301", "kind": "장기",
                 "points": [{"period": "24M", "done": "2025.04.10", "assays": {}, "unsure": []}]}]
        with open(os.path.join(folder, "13. 안정성시험일지 판독.json"), "w", encoding="utf-8") as fh:
            json.dump({"logs": logs, "reader_version": 3, "covers_all": True}, fh, ensure_ascii=False)
        data = collect.collect(folder, product_name="퀴노비드안연고", log=lambda *a: None)
        self.assertEqual(len(data.stability_logs), 2)
