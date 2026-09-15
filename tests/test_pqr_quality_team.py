# -*- coding: utf-8 -*-
"""5항 '책임과 권한' 의 담당 팀 — 제형에 따라 AQA팀·SQA팀으로 갈린다.

담당자 2026-09-15: "점안제, 주사제, 안연고제, 안구이식제는 AQA팀이고 나머지 제품은
SQA팀으로 작성해주면 돼." 전에는 모든 제품을 AQA팀으로 바꾸어, 고형제 결재본까지
AQA팀으로 나갔다 (담당자: "해당 제품은 고형제 제품인데 AQA로 작성이 되었네").
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr import build                                       # noqa: E402
from pqr.engine import recipe_ointment as R                 # noqa: E402


class 제형별_담당_팀(unittest.TestCase):
    def setUp(self):
        self.config = build.load_config()

    def team(self, form, name=""):
        return build.quality_team(form, self.config, name)

    def test_점안제와_주사제는_AQA팀(self):
        self.assertEqual(self.team("점안제", "히아루론맥스점안액0.15%"), "AQA팀")
        self.assertEqual(self.team("주사제", "한림덱스토민주"), "AQA팀")

    def test_안연고와_안구이식제는_제형_군이_달라도_AQA팀(self):
        # 안연고는 '연고제', 안구이식제는 '기타제' 군이라 군만으로는 가릴 수 없다
        self.assertEqual(self.team("연고제", "토브라덱스안연고"), "AQA팀")
        self.assertEqual(self.team("기타제", "안구이식제"), "AQA팀")

    def test_그_밖의_제품은_SQA팀(self):
        self.assertEqual(self.team("정제", "베넥사엑스알서방캡슐75mg"), "SQA팀")
        self.assertEqual(self.team("캡슐제", "자이자핀정5mg"), "SQA팀")
        self.assertEqual(self.team("연고제", "케토톱크림"), "SQA팀")
        self.assertEqual(self.team("내용액제", "한림시럽"), "SQA팀")

    def test_제형_칸이_비어_있으면_제품명으로_본다(self):
        self.assertEqual(self.team("", "히아루론점안액"), "AQA팀")
        self.assertEqual(self.team("", "다파로엠서방정10/1000밀리그램"), "SQA팀")


class 보고서가_쓰는_팀(unittest.TestCase):
    def test_build_가_정해_둔_값을_그대로_쓴다(self):
        self.assertEqual(R.quality_team_of({"quality_team": "AQA팀", "form": "정제"}), "AQA팀")

    def test_없으면_제형으로_다시_푼다(self):
        self.assertEqual(R.quality_team_of({"form": "정제", "name": "카세핀정12.5밀리그램"}), "SQA팀")
        self.assertEqual(R.quality_team_of({"form": "점안제", "name": "디쿠아린점안액3%"}), "AQA팀")

    def test_옛_이름도_AQA도_이_제품의_팀으로_고친다(self):
        self.assertEqual(R._TEAM_RE.sub("SQA팀", "품질보증 1팀 담당"), "SQA팀 담당")
        self.assertEqual(R._TEAM_RE.sub("SQA팀", "AQA팀 팀장"), "SQA팀 팀장")
        self.assertEqual(R._TEAM_RE.sub("AQA팀", "SQA팀 담당"), "AQA팀 담당")

    def test_품질보증부서장은_건드리지_않는다(self):
        self.assertEqual(R._TEAM_RE.sub("SQA팀", "품질보증부서장"), "품질보증부서장")


class 제품_목록에_팀이_실린다(unittest.TestCase):
    def test_build_가_만든_제품에_quality_team_이_있다(self):
        import tempfile, shutil
        from pqr.sample import write_samples
        folder = tempfile.mkdtemp()
        try:
            write_samples(folder, layout="tree")
            data = build.build(folder, today="2026-03-31")
            self.assertTrue(data["products"], "제품이 없습니다")
            for product in data["products"]:
                self.assertIn(product["quality_team"], ("AQA팀", "SQA팀"), product["code"])
        finally:
            shutil.rmtree(folder, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
