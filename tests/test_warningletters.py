"""FDA 경고장 수집·선별·메일 생성 테스트 — 네트워크 없이 로컬 모의 서버로 검증."""

import json
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from gmpai.mailer import MailConfig, MailError, build_message, render_html, render_text, send
from gmpai.warningletters import (
    Source,
    WarningLetter,
    WarningLetterError,
    classify,
    collect,
    enrich,
    load_state,
    new_letters,
    only_relevant,
    parse_date,
    parse_datatables,
    parse_landing_html,
    parse_rss,
    save_state,
    sort_letters,
    within_days,
)

RSS = """<?xml version="1.0" encoding="utf-8"?>
<rss version="2.0"><channel>
  <item>
    <title>Sterile Pharma Co., Ltd. - MARCS-CMS 690123 - September 16, 2026</title>
    <link>https://example.test/letter/sterile-pharma</link>
    <description>Aseptic processing deficiencies</description>
    <pubDate>Tue, 23 Sep 2026 14:05:00 -0400</pubDate>
  </item>
  <item>
    <title>Bulk Chemicals Inc. - MARCS-CMS 690124 - 09/02/2026</title>
    <link>https://example.test/letter/bulk-chemicals</link>
    <description>Labeling</description>
    <pubDate>Tue, 09 Sep 2026 10:00:00 -0400</pubDate>
  </item>
</channel></rss>
"""

DATATABLES = json.dumps(
    {
        "recordsTotal": 1,
        "data": [
            {
                "field_company_name_text": '<a href="/inspections/warning-letters/acme-260910">Acme Ophthalmics</a>',
                "field_letter_issue_datetime": "09/10/2026",
                "field_change_date_2": "09/17/2026",
                "field_building_issuing_office_ta": "Division of Pharmaceutical Quality Operations II",
                "field_subject_text": "CGMP/Finished Pharmaceuticals/Adulterated",
            }
        ],
    }
)

LANDING = """
<html><body>
  <a href="/inspections/warning-letters">Warning Letters</a>
  <a href="/inspections/warning-letters/nova-sterile-260915">Nova Sterile LLC - 09/15/2026</a>
  <a href="/about-fda/contact">Contact</a>
</body></html>
"""

LETTER_BODY = """
<html><body>
<p>Your firm failed to establish an adequate aseptic process. Smoke studies were not
performed. Environmental monitoring data showed recurring excursions, and sterility
test failures were invalidated without justification.</p>
<p>This is a violation of 21 CFR 211.113(b) and 21 CFR 211.192.</p>
<p>Division of Pharmaceutical Quality Operations I</p>
</body></html>
"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def _send(self, code, content_type, body):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = self.path.split("?")[0]
        if path == "/rss.xml":
            self._send(200, "application/rss+xml", RSS.encode())
        elif path == "/data.json":
            self._send(200, "application/json", DATATABLES.encode())
        elif path == "/landing":
            self._send(200, "text/html; charset=utf-8", LANDING.encode())
        elif path.startswith("/letter/"):
            self._send(200, "text/html; charset=utf-8", LETTER_BODY.encode())
        elif path == "/empty.xml":
            self._send(200, "application/rss+xml", b'<?xml version="1.0"?><rss><channel/></rss>')
        else:
            self._send(404, "text/plain", b"not found")


class ServerTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        cls.base = f"http://127.0.0.1:{cls.server.server_address[1]}"
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()


class ParseTests(unittest.TestCase):
    def test_rss_title_is_split_into_company_and_date(self):
        letters = parse_rss(RSS)
        self.assertEqual(len(letters), 2)
        first = letters[0]
        self.assertEqual(first.company, "Sterile Pharma Co., Ltd.")
        self.assertEqual(first.letter_date, "2026-09-16")
        self.assertEqual(first.posted_date, "2026-09-23")
        self.assertEqual(first.source, "rss")

    def test_rss_handles_slash_dates(self):
        self.assertEqual(parse_rss(RSS)[1].letter_date, "2026-09-02")

    def test_rss_rejects_malformed_xml(self):
        with self.assertRaises(WarningLetterError):
            parse_rss("<rss><channel>")

    def test_datatables_finds_fields_by_partial_key(self):
        letters = parse_datatables(DATATABLES)
        self.assertEqual(len(letters), 1)
        letter = letters[0]
        self.assertEqual(letter.company, "Acme Ophthalmics")
        self.assertEqual(letter.letter_date, "2026-09-10")
        self.assertEqual(letter.posted_date, "2026-09-17")
        self.assertIn("Division of Pharmaceutical Quality", letter.office)
        self.assertEqual(letter.url, "https://www.fda.gov/inspections/warning-letters/acme-260910")

    def test_datatables_rejects_non_json(self):
        with self.assertRaises(WarningLetterError):
            parse_datatables("<html>nope</html>")

    def test_landing_html_keeps_only_individual_letters(self):
        letters = parse_landing_html(LANDING)
        self.assertEqual(len(letters), 1)
        self.assertEqual(letters[0].letter_date, "2026-09-15")
        self.assertTrue(letters[0].url.endswith("/nova-sterile-260915"))

    def test_parse_date_formats(self):
        for value, expected in [
            ("09/16/2026", "2026-09-16"),
            ("September 16, 2026", "2026-09-16"),
            ("2026-09-16", "2026-09-16"),
            ("Tue, 23 Sep 2026 14:05:00 -0400", "2026-09-23"),
            ("Acme - 690123 - September 16, 2026", "2026-09-16"),
            ("", ""),
            ("no date here", ""),
        ]:
            with self.subTest(value=value):
                self.assertEqual(parse_date(value), expected)


class CollectTests(ServerTestCase):
    def test_uses_first_working_source(self):
        seen = []
        letters = collect(
            [
                Source("rss", f"{self.base}/rss.xml", parse_rss),
                Source("datatables", f"{self.base}/data.json", parse_datatables),
            ],
            retries=1,
            on_source=lambda n, o, c: seen.append((n, o, c)),
        )
        self.assertEqual(len(letters), 2)
        self.assertEqual(seen, [("rss", "성공", 2)])

    def test_falls_back_when_source_is_unreachable(self):
        seen = []
        letters = collect(
            [
                Source("rss", f"{self.base}/missing.xml", parse_rss),
                Source("datatables", f"{self.base}/data.json", parse_datatables),
            ],
            retries=1,
            on_source=lambda n, o, c: seen.append((n, o, c)),
        )
        self.assertEqual([l.company for l in letters], ["Acme Ophthalmics"])
        self.assertEqual([s[1] for s in seen], ["실패", "성공"])

    def test_falls_back_when_source_is_empty(self):
        letters = collect(
            [
                Source("rss", f"{self.base}/empty.xml", parse_rss),
                Source("html", f"{self.base}/landing", parse_landing_html),
            ],
            retries=1,
        )
        self.assertEqual(len(letters), 1)

    def test_raises_when_every_source_fails(self):
        with self.assertRaises(WarningLetterError) as ctx:
            collect([Source("rss", f"{self.base}/missing.xml", parse_rss)], retries=1)
        self.assertIn("rss", str(ctx.exception))

    def test_enrich_reads_body_for_keywords_and_cfr(self):
        letters = enrich(parse_rss(RSS.replace("https://example.test", self.base)), retries=1)
        first = letters[0]
        self.assertIn("무균조작", first.sterile_hits)
        self.assertIn("환경모니터링", first.sterile_hits)
        self.assertIn("기류시험", first.sterile_hits)
        self.assertEqual(first.cfr_citations, ["211.113(b)", "211.192"])
        self.assertIn("Division of Pharmaceutical Quality", first.office)

    def test_enrich_survives_a_dead_link(self):
        errors = []
        letter = WarningLetter(company="Gone Inc", url=f"{self.base}/missing", subject="sterile injection")
        result = enrich([letter], retries=1, on_error=lambda l, e: errors.append(l.company))
        self.assertEqual(errors, ["Gone Inc"])
        # 본문을 못 읽어도 제목만으로 분류는 해 둔다.
        self.assertIn("무균", result[0].sterile_hits)


class SelectionTests(unittest.TestCase):
    def _letters(self):
        return [
            WarningLetter(company="A", url="u1", posted_date="2026-09-20", sterile_hits=["무균"]),
            WarningLetter(company="B", url="u2", posted_date="2026-09-16"),
            WarningLetter(company="C", url="u3", posted_date="2026-08-01", cgmp_hits=["CGMP"]),
            WarningLetter(company="D", url="u4", posted_date=""),
        ]

    def test_within_days_keeps_recent_and_undated(self):
        from datetime import datetime, timezone

        now = datetime(2026, 9, 22, tzinfo=timezone.utc)
        kept = [l.company for l in within_days(self._letters(), 7, now=now)]
        self.assertEqual(kept, ["A", "B", "D"])

    def test_only_relevant_and_sterile_only(self):
        letters = self._letters()
        self.assertEqual([l.company for l in only_relevant(letters)], ["A", "C"])
        self.assertEqual([l.company for l in only_relevant(letters, sterile_only=True)], ["A"])

    def test_sort_puts_sterile_first_then_newest(self):
        letters = [
            WarningLetter(company="old-sterile", url="a", posted_date="2026-09-01", sterile_hits=["무균"]),
            WarningLetter(company="new-cgmp", url="b", posted_date="2026-09-20", cgmp_hits=["CGMP"]),
            WarningLetter(company="new-sterile", url="c", posted_date="2026-09-18", sterile_hits=["무균"]),
            WarningLetter(company="undated", url="d"),
        ]
        self.assertEqual(
            [l.company for l in sort_letters(letters)],
            ["new-sterile", "old-sterile", "new-cgmp", "undated"],
        )

    def test_classify_matches_keywords_case_insensitively(self):
        letter = classify(WarningLetter(company="X", url="u"), "STERILE ophthalmic Data Integrity 21 CFR 211.42(c)")
        self.assertIn("무균", letter.sterile_hits)
        self.assertIn("점안제", letter.sterile_hits)
        self.assertIn("데이터 완전성", letter.cgmp_hits)
        self.assertEqual(letter.cfr_citations, ["211.42(c)"])


class StateTests(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())
        self.path = self.dir / "state.json"

    def test_new_letters_excludes_already_notified(self):
        letters = [WarningLetter(company="A", url="u1"), WarningLetter(company="B", url="u2")]
        save_state(self.path, {"seen": {}}, letters[:1])
        state = load_state(self.path)
        self.assertEqual([l.company for l in new_letters(letters, state)], ["B"])

    def test_state_accumulates_across_runs(self):
        save_state(self.path, load_state(self.path), [WarningLetter(company="A", url="u1")])
        save_state(self.path, load_state(self.path), [WarningLetter(company="B", url="u2")])
        self.assertEqual(sorted(load_state(self.path)["seen"]), ["u1", "u2"])

    def test_missing_or_corrupt_state_is_treated_as_empty(self):
        self.assertEqual(load_state(self.dir / "nope.json"), {"seen": {}})
        self.path.write_text("{not json", encoding="utf-8")
        self.assertEqual(load_state(self.path), {"seen": {}})


class _FakeSMTP:
    instances = []

    def __init__(self, host, port):
        self.host, self.port = host, port
        self.logged_in = None
        self.sent = []
        _FakeSMTP.instances.append(self)

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def login(self, user, password):
        self.logged_in = (user, password)

    def send_message(self, message):
        self.sent.append(message)


class MailTests(unittest.TestCase):
    ENV = {
        "GMPAI_SMTP_USER": "sender@example.test",
        "GMPAI_SMTP_PASSWORD": "app-password",
        "GMPAI_MAIL_TO": "a@hanlim.test, b@hanlim.test",
    }

    def setUp(self):
        _FakeSMTP.instances = []
        self.letters = [
            WarningLetter(
                company="Sterile Pharma <Co>",
                url="https://example.test/letter/1",
                letter_date="2026-09-16",
                posted_date="2026-09-23",
                office="Division of Pharmaceutical Quality Operations I",
                sterile_hits=["무균조작"],
                cfr_citations=["211.113(b)"],
            )
        ]

    def test_from_env_reports_every_missing_variable(self):
        with self.assertRaises(MailError) as ctx:
            MailConfig.from_env({})
        for name in ("GMPAI_SMTP_USER", "GMPAI_SMTP_PASSWORD", "GMPAI_MAIL_TO"):
            self.assertIn(name, str(ctx.exception))

    def test_from_env_defaults_and_recipient_split(self):
        config = MailConfig.from_env(self.ENV)
        self.assertEqual(config.host, "smtp.gmail.com")
        self.assertEqual(config.port, 587)
        self.assertTrue(config.starttls)
        self.assertEqual(config.sender, "sender@example.test")
        self.assertEqual(config.recipients, ["a@hanlim.test", "b@hanlim.test"])

    def test_html_escapes_company_and_keeps_link(self):
        body = render_html(self.letters, "2026-09-16", source="rss")
        self.assertIn("Sterile Pharma &lt;Co&gt;", body)
        self.assertNotIn("<Co>", body)
        self.assertIn('href="https://example.test/letter/1"', body)
        self.assertIn("21 CFR 211.113(b)", body)
        self.assertIn("무균·주사제·점안제 관련 <b>1건</b>", body)

    def test_text_alternative_lists_the_letter(self):
        text = render_text(self.letters, "2026-09-16")
        self.assertIn("[무균] Sterile Pharma <Co>", text)
        self.assertIn("https://example.test/letter/1", text)

    def test_message_is_multipart_with_expected_subject(self):
        message = build_message(self.letters, MailConfig.from_env(self.ENV), "2026-09-16", "rss")
        self.assertEqual(message["Subject"], "[FDA] 신규 Warning Letter 1건")
        self.assertEqual(message["To"], "a@hanlim.test, b@hanlim.test")
        self.assertEqual(
            sorted(part.get_content_type() for part in message.iter_parts()),
            ["text/html", "text/plain"],
        )

    def test_send_logs_in_and_delivers(self):
        config = MailConfig.from_env(self.ENV)
        send(build_message(self.letters, config, "2026-09-16"), config, smtp_factory=_FakeSMTP)
        server = _FakeSMTP.instances[0]
        self.assertEqual((server.host, server.port), ("smtp.gmail.com", 587))
        self.assertEqual(server.logged_in, ("sender@example.test", "app-password"))
        self.assertEqual(len(server.sent), 1)

    def test_send_wraps_transport_errors(self):
        class Broken(_FakeSMTP):
            def send_message(self, message):
                raise OSError("connection reset")

        config = MailConfig.from_env(self.ENV)
        with self.assertRaises(MailError):
            send(build_message(self.letters, config, "2026-09-16"), config, smtp_factory=Broken)


if __name__ == "__main__":
    unittest.main()
