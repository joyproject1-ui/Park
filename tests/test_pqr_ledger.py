# -*- coding: utf-8 -*-
"""자료 판독 대장 — 입력 폴더의 파일마다 읽었는지 (담당자 2026-09-08: "PC 입력폴더에 있는 원본
파일을 빠짐없이 찾아서 읽었는지 검증하는 기능을 강화하라")."""
import os
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import collect as C                                    # noqa: E402


class 대장(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.mkdtemp(prefix="pqr-ledger-")
        def put(name, data=b"x"):
            path = os.path.join(self.folder, name)
            with open(path, "wb") as h:
                h.write(data)
            return path
        self.yield_x = put("7. 수율현황표.xlsx", b"y" * 10)
        self.erp = put("6. 제조내역 - ERP.pdf", b"e" * 20)
        put("6. 제조지시서.hwp", b"h" * 5)                     # 6항이 읽지 않는 형식
        put("작성 참고 문구.docx", b"d" * 7)                    # 항 번호 없음
        put("PQR 문의 목록 - X.txt", b"p")                       # 우리 파일
        put("PQR 작성 시 통일 문구.docx", b"q" * 3)              # 'PQR ' 로 시작하지만 우리 것이 아니다
        with zipfile.ZipFile(os.path.join(self.folder, "9.2.4 포장 완료 후.zip"), "w") as z:
            z.writestr("LWY201.pdf", b"a" * 30)
            z.writestr("LWY501.pdf", b"b" * 31)

    def _rows(self, read):
        work = tempfile.mkdtemp(prefix="pqr-ledger-w-")
        got = C.discover(self.folder, work)
        by_name = {os.path.basename(p): p for ps in got.values() for p in ps}
        return {name: (item, status, detail)
                for item, name, status, detail in C.build_ledger(self.folder, got, {by_name[k]: v for k, v in read.items()})}

    def test_상태를_넷으로_가른다(self):
        rows = self._rows({"7. 수율현황표.xlsx": "수율 Lot 3개", "LWY201.pdf": "LWY201 성적서 · 시험항목 16건",
                           "6. 제조내역 - ERP.pdf": "제조내역 0줄"})
        self.assertEqual(rows["7. 수율현황표.xlsx"][1], "읽음")
        self.assertEqual(rows["9.2.4 포장 완료 후.zip/LWY201.pdf"][1], "읽음")            # 압축 안도 한 줄씩
        self.assertEqual(rows["9.2.4 포장 완료 후.zip/LWY501.pdf"][1], "안 읽음")         # 읽을 수 있는데 안 읽었다
        self.assertEqual(rows["6. 제조내역 - ERP.pdf"][1], "안 읽음")                     # 0줄 = 못 읽은 것
        self.assertEqual(rows["6. 제조지시서.hwp"][1], "못 읽음")
        self.assertEqual(rows["작성 참고 문구.docx"][1], "항 없음")
        self.assertEqual(rows["PQR 작성 시 통일 문구.docx"][1], "항 없음")
        self.assertEqual(rows["PQR 문의 목록 - X.txt"][1], "우리 파일")

    def test_안_읽음이_맨_앞이고_파일로_남는다(self):
        work = tempfile.mkdtemp(prefix="pqr-ledger-w-")
        got = C.discover(self.folder, work)
        rows = C.build_ledger(self.folder, got, {})
        self.assertEqual(rows[0][2], "안 읽음")
        data = C.ProductData(); data.ledger = rows
        path = C.write_ledger(self.folder, "QC1-0001", data)
        with open(path, encoding="utf-8") as h:
            text = h.read()
        self.assertIn("★ [7] 7. 수율현황표.xlsx — 안 읽음", text)
        self.assertIn("요약:", text)
        self.assertNotIn("노랑으로 둔 칸", text)

    def test_파일이_아예_없는_항은_자료_없음_으로_대장에_적는다(self):
        """담당자 2026-09-28: "메모장에 설명도 없고" — 파일이 없으면 대장에 줄이 없어 왜 비었는지 알 수 없었다."""
        data = C.ProductData()
        got = {"7": [self.yield_x]}
        rows = C.missing_sources(data, got)
        items = [i for i, _ in rows]
        self.assertIn("10.2", items)
        self.assertIn("8.1.1", items)
        self.assertNotIn("7", items)                       # 파일이 있는 항은 여기서 다루지 않는다
        self.assertNotIn("11", items)                      # 일탈·변경관리는 없는 것이 정상
        self.assertTrue(all("노랑" in why for _, why in rows))
        data.ledger = [("7", "7. 수율현황표.xlsx", "읽음", "수율 Lot 3개"),
                       ("10.2", "(파일 없음)", "자료 없음", "제품 폴더·공통 폴더 어디에도 이 항의 파일이 없음 — 그 항의 빈 칸은 노랑")]
        path = C.write_ledger(self.folder, "QC1-0001", data)
        with open(path, encoding="utf-8") as h:
            text = h.read()
        self.assertIn("★ [10.2] (파일 없음) — 자료 없음", text)
        self.assertIn("자료 없음 1", text)

    def test_노랑으로_둔_칸의_까닭도_대장_끝에(self):
        """담당자 2026-09-28: "9.2.1에 노랑마크가 표시되어 있는데 메모장 PQR 자료 판독대장에 관련 내용이 없네"."""
        data = C.ProductData()
        data.ledger = [("7", "7. 수율현황표.xlsx", "읽음", "수율 Lot 3개")]
        data.issues = [("6", "제조단위", "올해 공 기록서가 없어 전년도 결재본 값(300 L)을 옮겼습니다(노랑) — 확인하세요"),
                       ("13", "", "시험 기간을 적지 못함")]
        data.issues_for_ledger = data.issues + [("8.1.1", "RSN101", "공급업체 목록('실시간' 시트 3줄)에서 관리번호 RSN101 을(를) 찾지 못했습니다 — 노랑으로 두었으니 확인")]
        path = C.write_ledger(self.folder, "QC1-0001", data)
        with open(path, encoding="utf-8") as h:
            text = h.read()
        self.assertIn("노랑으로 둔 칸 — 까닭", text)
        self.assertIn("[6] 제조단위 — 올해 공 기록서가 없어", text)
        self.assertIn("[8.1.1] RSN101 — 공급업체 목록('실시간' 시트 3줄)", text)
        self.assertNotIn("시험 기간을 적지 못함", text)               # 노랑과 상관없는 문의는 대장에 넣지 않는다


if __name__ == "__main__":
    unittest.main()
