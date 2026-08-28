import io
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date, datetime, time, timedelta
from pathlib import Path

from stock.analysis import analyze, best_pairs, worst_pairs
from stock.cli import main
from stock.loader import LoaderError, load_csv, load_csv_text, write_csv
from stock.model import Bar, DEFAULT_SLOTS, build_slots, infer_bar_minutes, normalize_stamp
from stock.momentum import profile
from stock.naver import FetchBlocked, TICKERS, fetch_minutes, parse_chart_xml, resolve_code
from stock.report import render, render_matrix, render_peak

TRADING_DAYS = [date(2026, 6, 1) + timedelta(days=n) for n in range(28) if (date(2026, 6, 1) + timedelta(days=n)).weekday() < 5]


def make_bars(step_by_minute):
    """분마다 step_by_minute(hhmm) 만큼 곱해 오르는(내리는) 결정론적 분봉.

    날짜마다 아주 작은 흔들림을 더해 표본 분산이 0 이 되지 않게 한다 (t값 계산용).
    """

    bars = []
    for index, day in enumerate(TRADING_DAYS):
        price = 10000.0
        wobble = 1.0 + (index % 5) * 0.02
        cursor = datetime.combine(day, time(9, 1))
        end = datetime.combine(day, time(15, 30))
        while cursor <= end:
            open_ = price
            price = open_ * (1.0 + step_by_minute(cursor.time()) * wobble)
            bars.append(Bar(ts=cursor, open=open_, high=max(open_, price), low=min(open_, price), close=price, volume=100))
            cursor += timedelta(minutes=1)
    return bars


def morning_down_close_up(t):
    if t < time(9, 31):  # end-stamp 기준 09:00~09:30 구간
        return -0.0004
    if t > time(15, 0):
        return +0.0006
    return 0.0


class ModelTests(unittest.TestCase):
    def test_default_slots_cover_the_regular_session(self):
        self.assertEqual(DEFAULT_SLOTS[0].start, time(9, 0))
        self.assertEqual(DEFAULT_SLOTS[-1].end, time(15, 30))
        for prev, cur in zip(DEFAULT_SLOTS, DEFAULT_SLOTS[1:]):
            self.assertEqual(prev.end, cur.start, "시간대 사이에 빈 구간이 있으면 안 됩니다")

    def test_build_slots_always_isolates_the_closing_auction(self):
        for minutes in (5, 15, 30, 60, 90):
            slots = build_slots(minutes)
            self.assertEqual(slots[-1].start, time(15, 20))
            self.assertEqual(slots[-1].end, time(15, 30))
            self.assertEqual(slots[0].start, time(9, 0))

    def test_build_slots_rejects_non_positive_length(self):
        with self.assertRaises(ValueError):
            build_slots(0)

    def test_infer_bar_minutes_uses_the_most_common_gap(self):
        bars = make_bars(lambda t: 0.0)[:100]
        self.assertEqual(infer_bar_minutes(bars), 1)

    def test_normalize_stamp_shifts_end_stamped_bars_back(self):
        bar = Bar(ts=datetime(2026, 6, 1, 9, 1), open=1, high=1, low=1, close=1, volume=0)
        moved = normalize_stamp([bar], stamp="end", minutes=1)[0]
        self.assertEqual(moved.ts, datetime(2026, 6, 1, 9, 0))
        kept = normalize_stamp([bar], stamp="start", minutes=1)[0]
        self.assertEqual(kept.ts, bar.ts)

    def test_normalize_stamp_rejects_unknown_mode(self):
        with self.assertRaises(ValueError):
            normalize_stamp([], stamp="middle", minutes=1)


class LoaderTests(unittest.TestCase):
    def test_reads_korean_headers_and_thousand_separators(self):
        text = "날짜,시간,시가,고가,저가,종가,거래량\n20260828,090100,25200,25400,25100,25300,\"1,234\"\n"
        bars = load_csv_text(text)
        self.assertEqual(len(bars), 1)
        self.assertEqual(bars[0].ts, datetime(2026, 8, 28, 9, 1))
        self.assertEqual(bars[0].volume, 1234)

    def test_close_only_file_synthesises_ohlc(self):
        bars = load_csv_text("datetime,close\n2026-08-28 09:01,25300\n")
        self.assertEqual(bars[0].open, bars[0].close)
        self.assertEqual(bars[0].high, bars[0].close)

    def test_missing_close_column_is_an_error(self):
        with self.assertRaises(LoaderError):
            load_csv_text("datetime,volume\n2026-08-28 09:01,10\n")

    def test_missing_timestamp_column_is_an_error(self):
        with self.assertRaises(LoaderError):
            load_csv_text("close\n25300\n")

    def test_round_trip_through_cp949_file(self):
        bars = make_bars(lambda t: 0.0)[:5]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "cp949.csv"
            path.write_bytes("날짜,시간,종가\n20260828,0901,25300\n".encode("cp949"))
            self.assertEqual(load_csv(path)[0].close, 25300)

            utf8 = Path(tmp) / "out.csv"
            write_csv(utf8, bars)
            self.assertEqual(len(load_csv(utf8)), len(bars))


class AnalysisTests(unittest.TestCase):
    def setUp(self):
        bars = normalize_stamp(make_bars(morning_down_close_up), stamp="end", minutes=1)
        self.analysis = analyze(bars, DEFAULT_SLOTS, code="TEST")

    def test_every_trading_day_is_complete(self):
        self.assertEqual(self.analysis.n_days, len(TRADING_DAYS))
        self.assertEqual(self.analysis.incomplete_days, [])

    def test_opening_slot_is_negative_and_closing_slot_is_positive(self):
        stats = {s.label.split()[0]: s for s in self.analysis.slot_stats}
        self.assertLess(stats["09:00-09:30"].mean_bp, -100)
        self.assertGreater(stats["15:00-15:20"].mean_bp, 100)
        self.assertEqual(stats["11:00-11:30"].mean_bp, 0.0)
        self.assertGreater(stats["15:00-15:20"].win_rate, 0.9)

    def test_a_deterministic_drift_is_flagged_significant(self):
        opening = self.analysis.slot_stats[0]
        self.assertTrue(opening.significant)
        self.assertEqual(opening.win_rate, 0.0)

    def test_flat_midday_slot_is_not_flagged(self):
        midday = [s for s in self.analysis.slot_stats if s.label.startswith("11:00")][0]
        self.assertFalse(midday.significant)

    def test_volume_shares_sum_to_one(self):
        total = sum(s.volume_share for s in self.analysis.slot_stats)
        self.assertAlmostEqual(total, 1.0, places=6)

    def test_best_pair_avoids_the_falling_opening_slot(self):
        best = best_pairs(self.analysis, top=1, min_days=5)[0]
        self.assertGreater(best.buy, 0)
        self.assertGreaterEqual(best.sell, len(DEFAULT_SLOTS) - 2, "상승 구간 끝에서 팔아야 합니다")
        worst = worst_pairs(self.analysis, top=1, min_days=5)[0]
        self.assertEqual(worst.buy, 0)

    def test_pair_matrix_is_upper_triangular(self):
        for (buy, sell) in self.analysis.pair_stats:
            self.assertLessEqual(buy, sell)

    def test_overnight_and_intraday_summaries_exist(self):
        self.assertIsNotNone(self.analysis.overnight)
        self.assertIsNotNone(self.analysis.intraday)
        self.assertEqual(self.analysis.overnight.days, self.analysis.n_days - 1)

    def test_partial_day_is_dropped_when_full_days_are_required(self):
        bars = normalize_stamp(make_bars(morning_down_close_up), stamp="end", minutes=1)
        trimmed = [b for b in bars if not (b.day == TRADING_DAYS[0] and b.clock > time(11, 0))]
        analysis = analyze(trimmed, DEFAULT_SLOTS)
        self.assertNotIn(TRADING_DAYS[0], analysis.days)
        self.assertIn(TRADING_DAYS[0], analysis.incomplete_days)

        loose = analyze(trimmed, DEFAULT_SLOTS, require_full_day=False)
        self.assertIn(TRADING_DAYS[0], loose.days)


def pump_then_fade(t):
    """09:30~11:00 에 올리고 13:00 이후 흘러내리는, 전형적인 '올리고 빠지는' 하루."""

    if time(9, 30) < t <= time(11, 0):
        return +0.0010
    if t > time(13, 0):
        return -0.0006
    return 0.0


class MomentumTests(unittest.TestCase):
    def setUp(self):
        bars = normalize_stamp(make_bars(pump_then_fade), stamp="end", minutes=1)
        self.analysis = analyze(bars, DEFAULT_SLOTS, code="PUMP")
        self.profile = profile(self.analysis, surge_threshold_bp=100.0)

    def test_high_is_found_in_the_slot_where_the_rally_ends(self):
        top = max(self.profile.shapes, key=lambda s: s.high_share)
        self.assertTrue(top.label.startswith("10:30-11:00"), top.label)
        self.assertGreater(top.high_share, 0.9)

    def test_low_is_found_at_the_end_of_the_fade(self):
        bottom = max(self.profile.shapes, key=lambda s: s.low_share)
        self.assertTrue(bottom.label.startswith("15:20"), bottom.label)

    def test_every_day_is_classified_as_pump_and_fade(self):
        self.assertEqual(self.profile.pump_and_fade_rate, 1.0)
        self.assertEqual(self.profile.n_days, len(TRADING_DAYS))

    def test_runup_is_positive_and_fade_is_negative(self):
        self.assertGreater(self.profile.median_runup_bp, 0)
        self.assertLess(self.profile.median_fade_bp, 0)

    def test_surge_filter_selects_days_above_the_threshold(self):
        loose = profile(self.analysis, surge_threshold_bp=100.0)
        strict = profile(self.analysis, surge_threshold_bp=100_000.0)
        self.assertEqual(len(loose.surge_days), len(TRADING_DAYS))
        self.assertEqual(strict.surge_days, [])
        for day in loose.surge_days:
            self.assertGreaterEqual(day.runup_bp, 100.0)

    def test_surge_high_shares_sum_to_one(self):
        self.assertAlmostEqual(sum(self.profile.surge_high_shares()), 1.0, places=6)

    def test_exit_suggestion_points_at_the_rally_top(self):
        exits = self.profile.best_exit_slots(top=1)
        self.assertTrue(self.profile.slots[exits[0]].startswith("10:30-11:00"))

    def test_follow_through_is_one_inside_a_continuous_rally(self):
        rally = [s for s in self.profile.shapes if s.label.startswith("09:30-10:00")][0]
        self.assertEqual(rally.follow_through, 1.0)
        self.assertGreater(rally.follow_days, 0)

    def test_last_slot_has_no_follow_through_sample(self):
        self.assertEqual(self.profile.shapes[-1].follow_days, 0)

    def test_falls_back_to_overall_highs_when_no_surge_day(self):
        strict = profile(self.analysis, surge_threshold_bp=100_000.0)
        self.assertEqual(strict.surge_high_shares(), [0.0] * len(DEFAULT_SLOTS))
        self.assertTrue(strict.slots[strict.best_exit_slots(top=1)[0]].startswith("10:30-11:00"))


class ReportTests(unittest.TestCase):
    def setUp(self):
        bars = normalize_stamp(make_bars(morning_down_close_up), stamp="end", minutes=1)
        self.analysis = analyze(bars, DEFAULT_SLOTS, code="243070")

    def test_text_report_mentions_sample_size_and_disclaimer(self):
        text = render(self.analysis, min_days=5)
        self.assertIn("243070", text)
        self.assertIn("{}거래일".format(self.analysis.n_days), text)
        self.assertIn("투자 판단과 그 결과는 본인 책임", text)

    def test_markdown_report_is_a_table(self):
        text = render(self.analysis, markdown=True, min_days=5)
        self.assertIn("| 시간대 |", text)
        self.assertIn("| --- |", text)

    def test_peak_report_reports_shape_and_caveat(self):
        prof = profile(self.analysis)
        text = render_peak(self.analysis, prof)
        self.assertIn("[시간대별 고점·저점·연속성]", text)
        self.assertIn("매수 후보", text)
        self.assertIn("사후 통계", text)
        self.assertNotIn("%%", text)

    def test_matrix_lower_triangle_is_blank(self):
        rows = render_matrix(self.analysis).splitlines()
        last_cell = rows[-1].split()
        self.assertEqual(len(last_cell), 2, "마지막 행에는 자기 자신 조합 하나만 남아야 합니다")


class NaverTests(unittest.TestCase):
    def test_resolve_code_accepts_name_or_code(self):
        self.assertEqual(resolve_code("휴온스"), "243070")
        self.assertEqual(resolve_code("243070"), "243070")
        self.assertEqual(resolve_code(" 084110 "), "084110")
        with self.assertRaises(ValueError):
            resolve_code("없는종목")

    def test_every_registered_ticker_is_six_digits(self):
        for name, code in TICKERS.items():
            self.assertRegex(code, r"^\d{6}$", name)

    def test_parse_chart_xml(self):
        payload = (
            '<protocol><chartdata symbol="243070" timeframe="minute">'
            '<item data="202608280901|25200|25400|25100|25300|1234"/>'
            '<item data="202608280902|25300|25350|25250|25280|987"/>'
            '<item data="깨진줄"/>'
            "</chartdata></protocol>"
        )
        bars = parse_chart_xml(payload)
        self.assertEqual(len(bars), 2)
        self.assertEqual(bars[0].ts, datetime(2026, 8, 28, 9, 1))
        self.assertEqual(bars[1].volume, 987)

    def test_unreachable_host_raises_fetch_blocked(self):
        with self.assertRaises(FetchBlocked):
            fetch_minutes("243070", count=1, timeout=1)


class CliTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.csv = Path(self.tmp.name) / "bars.csv"
        write_csv(self.csv, make_bars(morning_down_close_up))

    def _run(self, argv):
        buffer = io.StringIO()
        with redirect_stdout(buffer):
            code = main(argv)
        return code, buffer.getvalue()

    def test_slots_command_prints_a_table(self):
        code, out = self._run(["slots", str(self.csv), "--min-days", "5"])
        self.assertEqual(code, 0)
        self.assertIn("09:00-09:30 개장", out)
        self.assertIn("[시간대별 손익]", out)

    def test_slot_minutes_option_changes_the_grid(self):
        code, out = self._run(["slots", str(self.csv), "--slot-minutes", "60", "--min-days", "5"])
        self.assertEqual(code, 0)
        self.assertIn("09:00-10:00", out)

    def test_date_filters_narrow_the_sample(self):
        code, out = self._run(["slots", str(self.csv), "--from", "2026-06-15", "--min-days", "5"])
        self.assertEqual(code, 0)
        self.assertIn("2026-06-15", out)

    def test_peak_command(self):
        code, out = self._run(["peak", str(self.csv), "--surge", "0.5"])
        self.assertEqual(code, 0)
        self.assertIn("고점확률", out)
        self.assertIn("급등일", out)

    def test_matrix_command(self):
        code, out = self._run(["matrix", str(self.csv)])
        self.assertEqual(code, 0)
        self.assertIn("매수\\매도", out)

    def test_report_command_writes_markdown(self):
        out_path = Path(self.tmp.name) / "report.md"
        code, _ = self._run(["report", str(self.csv), "-o", str(out_path), "--min-days", "5"])
        self.assertEqual(code, 0)
        text = out_path.read_text(encoding="utf-8")
        self.assertIn("# 시간대별 분석", text)
        self.assertIn("매수→매도 매트릭스", text)
        self.assertIn("시간대별 고점·저점·연속성", text)

    def test_empty_selection_reports_an_error(self):
        code, _ = self._run(["slots", str(self.csv), "--from", "2030-01-01"])
        self.assertEqual(code, 1)

    def test_blocked_network_exits_with_code_two(self):
        code, _ = self._run(["slots", "--code", "243070", "--timeout", "1"])
        self.assertEqual(code, 2)


if __name__ == "__main__":
    unittest.main()
