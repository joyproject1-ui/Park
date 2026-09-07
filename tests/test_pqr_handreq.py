"""13항 판독 요청 묶음 — 이 PC 로 읽지 않고 Claude 대화에 올릴 zip 을 만든다 (담당자 2026-09-07)."""
import os
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr.engine import handreq                                       # noqa: E402


class 판독_요청_묶음(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="pqr-req-")

    def _pdf(self, name, size=1000):
        path = os.path.join(self.dir, name)
        with open(path, "wb") as fh:
            fh.write(b"%PDF-1.4\n" + b"0" * size)
        return path

    def test_시험일지와_안내와_서식을_한_묶음으로(self):
        paths = [self._pdf("13 장기(내수용).pdf"), self._pdf("13 장기(수출용).pdf")]
        made = handreq.make_request(self.dir, paths, "퀴노비드안연고")
        self.assertEqual(len(made), 1)
        name, path = made[0]
        self.assertEqual(name, "13. Claude 판독 요청 - 퀴노비드안연고.zip")
        with zipfile.ZipFile(path) as z:
            self.assertEqual(sorted(z.namelist()),
                             ["13 장기(내수용).pdf", "13 장기(수출용).pdf", "읽어주세요.txt"])
            guide = z.read("읽어주세요.txt").decode("utf-8")
        self.assertIn("13. 안정성시험일지 판독.json", guide)
        self.assertIn('"covers_all": true', guide)
        self.assertIn("13 장기(수출용).pdf", guide)
        self.assertIn(os.path.abspath(self.dir), guide)

    def test_너무_크면_나눠_담는다(self):
        big = handreq.PART_LIMIT
        paths = [self._pdf("a.pdf", big), self._pdf("b.pdf", big)]
        made = handreq.make_request(self.dir, paths, "퀴노비드안연고")
        self.assertEqual([n for n, _ in made],
                         ["13. Claude 판독 요청 - 퀴노비드안연고 (1 of 2).zip",
                          "13. Claude 판독 요청 - 퀴노비드안연고 (2 of 2).zip"])

    def test_지난_묶음은_지운다(self):
        old = os.path.join(self.dir, "13. Claude 판독 요청 - 옛것.zip")
        with zipfile.ZipFile(old, "w") as z:
            z.writestr("x.txt", "x")
        handreq.make_request(self.dir, [self._pdf("a.pdf")], "퀴노비드안연고")
        self.assertFalse(os.path.exists(old))

    def test_이름이_같은_시험일지는_번호를_붙인다(self):
        sub = os.path.join(self.dir, "안쪽")
        os.makedirs(sub)
        second = os.path.join(sub, "13 장기.pdf")
        with open(second, "wb") as fh:
            fh.write(b"%PDF-1.4\n")
        made = handreq.make_request(self.dir, [self._pdf("13 장기.pdf"), second], "")
        with zipfile.ZipFile(made[0][1]) as z:
            self.assertEqual(sorted(n for n in z.namelist() if n.endswith(".pdf")),
                             ["13 장기 (1).pdf", "13 장기.pdf"])

    def test_시험일지가_없으면_만들지_않는다(self):
        self.assertEqual(handreq.make_request(self.dir, [], "가"), [])


class PC_판독_사용(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="pqr-opt-")
        self.root = tempfile.mkdtemp(prefix="pqr-root-")

    def test_기본은_꺼짐(self):
        self.assertFalse(handreq.pc_reading_on(self.dir, self.root))

    def test_제품_폴더에_표시_파일이_있으면_켜짐(self):
        open(os.path.join(self.dir, handreq.PC_OPT_IN), "w").close()
        self.assertTrue(handreq.pc_reading_on(self.dir, self.root))

    def test_프로그램_폴더에_두어도_켜짐(self):
        open(os.path.join(self.root, handreq.PC_OPT_IN), "w").close()
        self.assertTrue(handreq.pc_reading_on(self.dir, self.root))


if __name__ == "__main__":
    unittest.main()
