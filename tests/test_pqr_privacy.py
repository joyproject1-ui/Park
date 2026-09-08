# -*- coding: utf-8 -*-
"""외부 전송 금지 스위치 — 담당자 2026-09-08: "PQR 내용 Claude 외부로 유출되는 것은 아닌지?"."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import privacy, claude_cli, vision                    # noqa: E402


class 외부_전송_금지(unittest.TestCase):
    def setUp(self):
        os.environ.pop(privacy.ENV, None)
        self._review = os.environ.pop("PQR_REVIEW", None)     # 다른 시험 모듈이 꺼 두었을 수 있다

    def tearDown(self):
        os.environ.pop(privacy.ENV, None)
        if self._review is not None:
            os.environ["PQR_REVIEW"] = self._review

    def test_환경_변수로_막는다(self):
        os.environ[privacy.ENV] = "1"
        self.assertTrue(privacy.blocked())
        self.assertFalse(claude_cli.available())          # Claude Code 가 깔려 있어도 부르지 않는다
        self.assertFalse(vision.available())
        self.assertIn("금지", privacy.note())

    def test_표시_파일로_막는다(self):
        folder = tempfile.mkdtemp(prefix="pqr-priv-")
        self.assertFalse(privacy.blocked(folder))
        with open(os.path.join(folder, privacy.OFFLINE_FILE), "w", encoding="utf-8") as h:
            h.write("")
        self.assertTrue(privacy.blocked(folder))

    def test_검토_단계도_건너뛴다(self):
        from pqr.engine import review
        os.environ[privacy.ENV] = "1"
        log = []
        self.assertEqual(review.review("/no/such", {"code": "X"}, "/no/such/r.docx", [], log.append), [])
        self.assertTrue(any("외부 전송 금지" in l for l in log))


if __name__ == "__main__":
    unittest.main()
