"""명령행 인터페이스: fetch / slots / matrix / report."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Sequence

from .analysis import analyze
from .loader import LoaderError, load_csv, write_csv
from .model import Bar, build_slots, infer_bar_minutes, normalize_stamp
from .momentum import profile
from .naver import DEFAULT_TIMEOUT, FetchBlocked, TICKERS, fetch_minutes, resolve_code
from .report import render, render_matrix, render_peak

DEFAULT_CODE = "243070"  # 휴온스


def _add_common(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("source", nargs="?", help="분봉 CSV 경로 (생략하면 --code 로 네이버 금융 조회)")
    parser.add_argument("--code", default=DEFAULT_CODE, help="종목코드 또는 종목명 (기본 243070 휴온스)")
    parser.add_argument("--count", type=int, default=2000, help="조회할 분봉 개수 (기본 2000)")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help="요청 타임아웃 초")
    parser.add_argument("--slot-minutes", type=int, default=30, help="시간대 길이(분), 기본 30")
    parser.add_argument("--stamp", choices=("end", "start"), default="end", help="분봉 시각 표기 기준 (기본 end)")
    parser.add_argument("--from", dest="date_from", help="시작일 YYYY-MM-DD")
    parser.add_argument("--to", dest="date_to", help="종료일 YYYY-MM-DD")
    parser.add_argument("--partial-days", action="store_true", help="일부 시간대가 비어 있는 날도 포함")
    parser.add_argument("--markdown", action="store_true", help="마크다운 표로 출력")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="stock",
        description="분봉을 시간대별로 접어서 매수/매도 시점 편향을 확인합니다. 기본 종목은 휴온스(243070).",
        epilog="등록 종목명: " + ", ".join("{}({})".format(k, v) for k, v in TICKERS.items()),
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_fetch = sub.add_parser("fetch", help="네이버 금융 분봉을 CSV 로 저장")
    p_fetch.add_argument("--code", default=DEFAULT_CODE, help="종목코드 또는 종목명")
    p_fetch.add_argument("--count", type=int, default=2000, help="조회할 분봉 개수")
    p_fetch.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help="요청 타임아웃 초")
    p_fetch.add_argument("-o", "--out", required=True, help="저장할 CSV 경로")

    p_slots = sub.add_parser("slots", help="시간대별 손익·거래량·승률 표")
    _add_common(p_slots)
    p_slots.add_argument("--top", type=int, default=5, help="매수/매도 조합 상·하위 개수")
    p_slots.add_argument("--min-days", type=int, default=10, help="조합 통계에 요구할 최소 표본일수")

    p_matrix = sub.add_parser("matrix", help="매수 시간대 × 매도 시간대 평균수익률 매트릭스")
    _add_common(p_matrix)

    p_peak = sub.add_parser("peak", help="고점·저점이 찍히는 시간대와 되돌림(올리고 빠지는 패턴)")
    _add_common(p_peak)
    p_peak.add_argument("--surge", type=float, default=3.0, help="급등일 기준 시가 대비 고가 상승률(%%), 기본 3.0")
    p_peak.add_argument("--top", type=int, default=3, help="요약에 보여줄 구간 수")

    p_report = sub.add_parser("report", help="표 전체를 파일로 저장")
    _add_common(p_report)
    p_report.add_argument("-o", "--out", required=True, help="저장할 마크다운 경로")
    p_report.add_argument("--top", type=int, default=5, help="매수/매도 조합 상·하위 개수")
    p_report.add_argument("--min-days", type=int, default=10, help="조합 통계에 요구할 최소 표본일수")
    p_report.add_argument("--surge", type=float, default=3.0, help="급등일 기준 시가 대비 고가 상승률(%%), 기본 3.0")

    return parser


def _parse_day(text: Optional[str]) -> Optional[datetime]:
    if not text:
        return None
    return datetime.strptime(text.strip(), "%Y-%m-%d")


def _collect_bars(args: argparse.Namespace) -> List[Bar]:
    if args.source:
        bars = load_csv(args.source)
    else:
        bars = fetch_minutes(args.code, count=args.count, timeout=args.timeout)

    start = _parse_day(args.date_from)
    end = _parse_day(args.date_to)
    if start:
        bars = [b for b in bars if b.ts >= start]
    if end:
        bars = [b for b in bars if b.ts.date() <= end.date()]
    if not bars:
        raise LoaderError("조건에 맞는 분봉이 없습니다")

    minutes = infer_bar_minutes(bars)
    return normalize_stamp(bars, stamp=args.stamp, minutes=minutes)


def _run_analysis(args: argparse.Namespace):
    bars = _collect_bars(args)
    slots = build_slots(args.slot_minutes)
    code = args.code if not args.source else Path(args.source).stem
    return analyze(bars, slots, code=code, require_full_day=not args.partial_days)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "fetch":
            bars = fetch_minutes(args.code, count=args.count, timeout=args.timeout)
            write_csv(args.out, bars)
            print(
                "{} 분봉 {}개 저장: {} ({} ~ {})".format(
                    resolve_code(args.code), len(bars), args.out, bars[0].ts, bars[-1].ts
                )
            )
            return 0

        analysis = _run_analysis(args)
        if args.command == "slots":
            print(render(analysis, markdown=args.markdown, top=args.top, min_days=args.min_days))
        elif args.command == "matrix":
            print(render_matrix(analysis, markdown=args.markdown))
        elif args.command == "peak":
            prof = profile(analysis, surge_threshold_bp=args.surge * 100)
            print(render_peak(analysis, prof, markdown=args.markdown, top=args.top))
        elif args.command == "report":
            prof = profile(analysis, surge_threshold_bp=args.surge * 100)
            text = render(analysis, markdown=True, top=args.top, min_days=args.min_days)
            text += "\n\n## 매수→매도 매트릭스 (평균 bp)\n\n" + render_matrix(analysis, markdown=True) + "\n"
            text += "\n\n" + render_peak(analysis, prof, markdown=True, top=args.top).split("\n", 1)[1] + "\n"
            out = Path(args.out)
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(text, encoding="utf-8")
            print("저장했습니다: {}".format(out))
        return 0
    except FetchBlocked as exc:
        print("[시세 조회 실패] {}".format(exc), file=sys.stderr)
        return 2
    except (LoaderError, ValueError) as exc:
        print("[오류] {}".format(exc), file=sys.stderr)
        return 1
