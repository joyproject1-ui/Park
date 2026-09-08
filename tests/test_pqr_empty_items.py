"""자료는 올렸는데 한 칸도 못 읽은 항은 작성 전에 알린다
(담당자 2026-09-08: "다른 제품 작성할 때 동일한 문제가 발생 안 되도록 조치해 줘")."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import collect as C                                   # noqa: E402


def _data(**kw):
    d = C.ProductData()
    for k, v in kw.items():
        setattr(d, k, v)
    return d


class 못_읽은_항(unittest.TestCase):
    def test_완제_성적서를_못_읽으면_알린다(self):
        got = {"9.2.4": ["/x/LWY201.pdf", "/x/LWY501.pdf", "/x/LWYO01.pdf"]}
        rows = C.empty_after_read(_data(), got)
        self.assertEqual([r[0] for r in rows], ["9.2.4"])
        self.assertIn("파일 3개", rows[0][1])
        self.assertIn("완제 성적서", rows[0][1])
        self.assertTrue(rows[0][1].startswith("★"))

    def test_읽혔으면_알리지_않는다(self):
        got = {"9.2.4": ["/x/LWY201.pdf"]}
        data = _data(coa={"LWY201": {"924": {"appearance": "무색"}}})
        self.assertEqual(C.empty_after_read(data, got), [])

    def test_안_올린_항은_다루지_않는다(self):
        self.assertEqual(C.empty_after_read(_data(), {}), [])

    def test_마감_글만_있는_항은_다루지_않는다(self):
        got = {"11": ["/x/11 일탈 - 해당없음 확인.txt"]}
        self.assertEqual(C.empty_after_read(_data(), got), [])

    def test_우리가_남긴_글은_세지_않는다(self):
        got = {"12": ["/x/PQR 변경요청서 판독.json"]}
        self.assertEqual(C.empty_after_read(_data(), got), [])

    def test_못_읽은_변경요청서는_읽은_것으로_치지_않는다(self):
        got = {"12": ["/x/12.CC-240723-08.pdf"]}
        data = _data(changes=[{"doc_no": "CC-240723-08", "unread": True}])
        self.assertEqual([r[0] for r in C.empty_after_read(data, got)], ["12"])

    def test_여러_항을_한꺼번에(self):
        got = {"7": ["/x/7. 수율현황표.xlsx"], "10.2": ["/x/10.2 마스터.xlsx"]}
        self.assertEqual(sorted(r[0] for r in C.empty_after_read(_data(), got)), ["10.2", "7"])


if __name__ == "__main__":
    unittest.main()
