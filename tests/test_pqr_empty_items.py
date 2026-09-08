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


class 경향_성분_이름(unittest.TestCase):
    """판독값의 성분 이름이 경향표와 달라도 값을 잃지 않는다
    (담당자 2026-09-08: 판독 3 Lot 이 있는데 '결과값이 없어' 경향 파일을 못 만들었다)."""

    def _log(self, assays, period="12M", done="2025.05.02", unsure=()):
        return {"lot": "LWY201", "points": [{"period": period, "done": done,
                                             "assays": assays, "unsure": list(unsure)}]}

    def test_이름이_같으면_그대로(self):
        from pqr.engine.excel_attach import points_by_part
        got = points_by_part(self._log({"트레할로스수화물": 99.8}), "트레할로스수화물", 2025)
        self.assertEqual(got, {"12M": 99.8})

    def test_같은_계열_이름이면_찾는다(self):
        from pqr.engine.excel_attach import points_by_part
        got = points_by_part(self._log({"트레할로스": 99.8}), "트레할로스수화물", 2025)
        self.assertEqual(got, {"12M": 99.8})

    def test_성분이_하나뿐이면_이름이_달라도_쓴다(self):
        from pqr.engine.excel_attach import points_by_part
        got = points_by_part(self._log({"함량": 99.8}), "트레할로스수화물", 2025)
        self.assertEqual(got, {"12M": 99.8})

    def test_성분이_여럿이고_이름이_다르면_쓰지_않는다(self):
        from pqr.engine.excel_attach import points_by_part
        got = points_by_part(self._log({"겐타마이신": 99.8, "플루오로메톨론": 101.2}),
                             "트레할로스수화물", 2025)
        self.assertEqual(got, {})

    def test_평가_기간을_넘어선_시점은_뺀다(self):
        from pqr.engine.excel_attach import points_by_part
        got = points_by_part(self._log({"함량": 99.8}, done="2026.01.20"), "트레할로스수화물", 2025)
        self.assertEqual(got, {})

    def test_애매한_시점은_따로_알린다(self):
        from pqr.engine.excel_attach import points_by_part
        shaky = set()
        points_by_part(self._log({"함량": 99.8}, unsure=["트레할로스수화물"]),
                       "트레할로스수화물", 2025, shaky)
        self.assertEqual(shaky, {"12M"})
