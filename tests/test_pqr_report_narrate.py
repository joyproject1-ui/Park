"""보고서 생성과 서술 문안 모듈 테스트 (API 호출 없음)."""

import json
import shutil
import tempfile
import unittest

from pqr import narrate as narrate_module
from pqr import report as report_module
from pqr.build import build
from pqr.sample import write_samples

TODAY = "2026-08-27"


class ReportTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.mkdtemp()
        write_samples(cls.dir, layout="tree")
        cls.data = build(input_dir=cls.dir, today=TODAY)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.dir, ignore_errors=True)

    def test_report_covers_all_twelve_items(self):
        text = report_module.product_report(self.data, "HP-201")
        for key, title, _ in report_module.ITEM_SECTIONS:
            self.assertIn("## (%s) %s" % (key, title), text)

    def test_report_states_it_is_a_draft(self):
        text = report_module.product_report(self.data, "HP-201")
        self.assertIn("초안", text)
        self.assertIn("검토", text)

    def test_report_includes_source_hashes(self):
        text = report_module.product_report(self.data, "HP-201")
        self.assertIn("SHA-256", text)
        self.assertIn("부록 — 데이터 출처", text)

    def test_report_without_narrative_says_so(self):
        text = report_module.product_report(self.data, "HP-201")
        self.assertIn("서술 문안이 생성되지 않았습니다", text)

    def test_unknown_product_raises(self):
        with self.assertRaises(KeyError):
            report_module.product_report(self.data, "없는코드")

    def test_write_reports_creates_summary_and_products(self):
        out = tempfile.mkdtemp()
        try:
            written = report_module.write_reports(self.data, out)
            self.assertEqual(len(written), len(self.data["products"]) + 1)
            self.assertTrue(written[0].endswith("PQR_요약.md"))
        finally:
            shutil.rmtree(out, ignore_errors=True)


class FakeBlock(object):
    type = "text"

    def __init__(self, text):
        self.text = text


class FakeResponse(object):
    def __init__(self, payload):
        self.content = [FakeBlock(json.dumps(payload, ensure_ascii=False))]


class FakeMessages(object):
    def __init__(self, payload, recorder):
        self.payload = payload
        self.recorder = recorder

    def create(self, **kwargs):
        self.recorder.append(kwargs)
        return FakeResponse(self.payload)


class FakeClient(object):
    """anthropic.Anthropic 을 대신하는 가짜 클라이언트 (네트워크 없음)."""

    def __init__(self, payload):
        self.calls = []
        self.messages = FakeMessages(payload, self.calls)


ANSWER = {
    "overview": "개요", "capability_assessment": "공정능력", "deviation_assessment": "일탈",
    "change_assessment": "변경", "stability_assessment": "안정성",
    "qualification_assessment": "적격성", "conclusion": "결론",
    "recommendations": ["조치 1"], "open_questions": ["확인 필요 1"],
}


class NarrateTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.mkdtemp()
        write_samples(cls.dir, layout="tree")
        cls.data = build(input_dir=cls.dir, today=TODAY)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.dir, ignore_errors=True)

    def test_payload_has_summaries_not_raw_records(self):
        payload = narrate_module.build_payload(self.data, "HP-201")
        self.assertIn("시험결과요약", payload)
        self.assertIn("평가항목_자료상태", payload)
        text = json.dumps(payload, ensure_ascii=False)
        # 배치별 원본 값이나 일탈 제목 같은 원자료는 전송하지 않습니다.
        self.assertNotIn("records", text)
        self.assertNotIn("충전 중 정지", text)

    def test_payload_stays_small(self):
        payload = narrate_module.build_payload(self.data, "HP-201")
        self.assertLess(len(json.dumps(payload, ensure_ascii=False)), 8000)

    def test_dry_run_does_not_call_the_api(self):
        preview = narrate_module.narrate(self.data, codes=["HP-201"], dry_run=True,
                                         log=lambda *args: None)
        self.assertIn("HP-201", preview)
        self.assertEqual(self.data["narrative"], {})

    def test_narrate_stores_result_and_marks_review_required(self):
        client = FakeClient(ANSWER)
        result = narrate_module.narrate(self.data, codes=["HP-201"], client=client,
                                        log=lambda *args: None)
        self.assertEqual(result["HP-201"]["conclusion"], "결론")
        self.assertTrue(result["HP-201"]["_meta"]["review_required"])

    def test_request_uses_structured_output_and_named_model(self):
        client = FakeClient(ANSWER)
        narrate_module.narrate(self.data, codes=["HP-201"], client=client,
                               model="claude-opus-5", log=lambda *args: None)
        request = client.calls[0]
        self.assertEqual(request["model"], "claude-opus-5")
        self.assertEqual(request["output_config"]["format"]["type"], "json_schema")
        self.assertIn("추정하지 마십시오", request["system"])

    def test_narrative_appears_in_report(self):
        client = FakeClient(ANSWER)
        narrate_module.narrate(self.data, codes=["HP-201"], client=client,
                               log=lambda *args: None)
        text = report_module.product_report(self.data, "HP-201")
        self.assertIn("결론", text)
        self.assertIn("조치 1", text)
        self.assertIn("확인 필요 1", text)

    def test_unknown_product_raises(self):
        with self.assertRaises(KeyError):
            narrate_module.build_payload(self.data, "없는코드")


if __name__ == "__main__":
    unittest.main()
