"""분석 결과를 표로 출력한다 (텍스트 / 마크다운 공용)."""

from __future__ import annotations

import unicodedata
from typing import List, Sequence

from .analysis import Analysis, PairStat, SlotStat, best_pairs, worst_pairs
from .momentum import MomentumProfile


def _width(text: str) -> int:
    """한글·전각 문자를 2칸으로 세어 터미널 표 정렬을 맞춘다."""

    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in text)


def _pad(text: str, width: int) -> str:
    return text + " " * max(0, width - _width(text))


def _fmt_bp(value: float) -> str:
    return "{:+.1f}".format(value)


def _pct(value: float) -> str:
    return "{:.1f}%".format(value * 100)


def _table(headers: Sequence[str], rows: Sequence[Sequence[str]], *, markdown: bool) -> List[str]:
    widths = [_width(h) for h in headers]
    for row in rows:
        for i, cell in enumerate(row):
            widths[i] = max(widths[i], _width(cell))
    out = []
    if markdown:
        out.append("| " + " | ".join(headers) + " |")
        out.append("| " + " | ".join("---" for _ in headers) + " |")
        for row in rows:
            out.append("| " + " | ".join(row) + " |")
    else:
        out.append("  ".join(_pad(h, widths[i]) for i, h in enumerate(headers)))
        out.append("  ".join("-" * widths[i] for i in range(len(headers))))
        for row in rows:
            out.append("  ".join(_pad(row[i], widths[i]) for i in range(len(headers))))
    return out


def _slot_rows(stats: Sequence[SlotStat]) -> List[List[str]]:
    rows = []
    for s in stats:
        rows.append(
            [
                s.label,
                str(s.days),
                _fmt_bp(s.mean_bp),
                _fmt_bp(s.median_bp),
                _pct(s.win_rate),
                "{:.1f}".format(s.stdev_bp),
                "{:+.2f}".format(s.t_stat),
                _fmt_bp(s.cumulative_bp),
                _pct(s.volume_share),
                "{:.1f}".format(s.range_bp),
                "*" if s.significant else "",
            ]
        )
    return rows


SLOT_HEADERS = ("시간대", "일수", "평균bp", "중앙bp", "승률", "변동성", "t값", "누적bp", "거래비중", "폭bp", "유의")


def render(analysis: Analysis, *, markdown: bool = False, top: int = 5, min_days: int = 10) -> str:
    lines: List[str] = []
    title = "시간대별 분석 {}".format(analysis.code or "").strip()
    if markdown:
        lines.append("# {}".format(title))
    else:
        lines.append(title)
        lines.append("=" * _width(title))
    lines.append("")

    if analysis.days:
        lines.append(
            "표본: {}거래일 ({} ~ {}), 시간대 {}구간".format(
                analysis.n_days, analysis.days[0], analysis.days[-1], len(analysis.slots)
            )
        )
    else:
        lines.append("표본 없음")
    if analysis.incomplete_days:
        lines.append("데이터가 불완전해 제외한 날: {}일".format(len(analysis.incomplete_days)))
    lines.append("수익률 단위 bp = 0.01%. t값 |t|>=2 이고 10거래일 이상이면 '유의' 표시.")
    lines.append("")

    lines.append("## 시간대별 손익" if markdown else "[시간대별 손익]")
    lines.extend(_table(SLOT_HEADERS, _slot_rows(analysis.slot_stats), markdown=markdown))
    lines.append("")

    extras = [s for s in (analysis.overnight, analysis.intraday) if s is not None]
    if extras:
        lines.append("## 오버나이트 vs 장중" if markdown else "[오버나이트 vs 장중]")
        lines.extend(_table(SLOT_HEADERS, _slot_rows(extras), markdown=markdown))
        lines.append("")

    best = best_pairs(analysis, top=top, min_days=min_days)
    worst = worst_pairs(analysis, top=top, min_days=min_days)
    if best:
        lines.append("## 매수→매도 조합 상위" if markdown else "[매수→매도 조합 상위]")
        lines.extend(_pair_table(analysis, best, markdown=markdown))
        lines.append("")
    if worst:
        lines.append("## 매수→매도 조합 하위" if markdown else "[매수→매도 조합 하위]")
        lines.extend(_pair_table(analysis, worst, markdown=markdown))
        lines.append("")

    lines.append(
        "주의: 과거 통계이며 수수료·세금·슬리피지·호가 스프레드가 반영되어 있지 않습니다. "
        "투자 판단과 그 결과는 본인 책임입니다."
    )
    return "\n".join(lines)


def _pair_table(analysis: Analysis, pairs: Sequence[PairStat], *, markdown: bool) -> List[str]:
    headers = ("매수(시가)", "매도(종가)", "일수", "평균bp", "승률", "변동성", "t값")
    rows = []
    for p in pairs:
        rows.append(
            [
                analysis.slots[p.buy].label,
                analysis.slots[p.sell].label,
                str(p.days),
                _fmt_bp(p.mean_bp),
                _pct(p.win_rate),
                "{:.1f}".format(p.stdev_bp),
                "{:+.2f}".format(p.t_stat),
            ]
        )
    return _table(headers, rows, markdown=markdown)


def render_matrix(analysis: Analysis, *, markdown: bool = False) -> str:
    """매수 시간대(행) × 매도 시간대(열) 평균수익률 bp 매트릭스."""

    labels = [s.label.split()[0] for s in analysis.slots]
    headers = ["매수\\매도"] + labels
    rows = []
    for i in range(len(analysis.slots)):
        row = [labels[i]]
        for j in range(len(analysis.slots)):
            stat = analysis.pair_stats.get((i, j))
            row.append("" if stat is None or j < i else "{:+.0f}".format(stat.mean_bp))
        rows.append(row)
    return "\n".join(_table(headers, rows, markdown=markdown))


PEAK_HEADERS = ("시간대", "고점확률", "저점확률", "상승비율", "다음구간 연속상승", "표본", "거래비중")


def render_peak(analysis: Analysis, prof: MomentumProfile, *, markdown: bool = False, top: int = 3) -> str:
    """'어디서 올리고 어디서 빠지는가' 를 보는 표."""

    lines: List[str] = []
    head = "상승 타이밍 프로파일 {}".format(analysis.code or "").strip()
    if markdown:
        lines.append("# {}".format(head))
    else:
        lines.append(head)
        lines.append("=" * _width(head))
    lines.append("")
    lines.append("표본 {}거래일. 고점확률 = 그 시간대에 당일 고가가 찍힌 날의 비율.".format(prof.n_days))
    lines.append("")

    rows = []
    for s in prof.shapes:
        rows.append(
            [
                s.label,
                _pct(s.high_share),
                _pct(s.low_share),
                _pct(s.up_rate),
                _pct(s.follow_through) if s.follow_days else "-",
                str(s.follow_days),
                _pct(s.volume_share),
            ]
        )
    lines.append("## 시간대별 고점·저점·연속성" if markdown else "[시간대별 고점·저점·연속성]")
    lines.extend(_table(PEAK_HEADERS, rows, markdown=markdown))
    lines.append("")

    lines.append("## 하루의 모양" if markdown else "[하루의 모양]")
    shape_rows = [
        ["고점을 저점보다 먼저 찍은 날 (올리고 빠지는 형태)", _pct(prof.pump_and_fade_rate)],
        ["시가→당일고가 상승폭 (중앙값)", _fmt_bp(prof.median_runup_bp) + "bp"],
        ["당일고가→종가 되돌림 (중앙값)", _fmt_bp(prof.median_fade_bp) + "bp"],
        [
            "급등일 (시가 대비 +{:.1f}% 이상 터치)".format(prof.surge_threshold_bp / 100),
            "{}일 / {}일 ({})".format(
                len(prof.surge_days), prof.n_days, _pct(len(prof.surge_days) / prof.n_days if prof.n_days else 0.0)
            ),
        ],
    ]
    lines.extend(_table(("항목", "값"), shape_rows, markdown=markdown))
    lines.append("")

    if prof.surge_days:
        shares = prof.surge_high_shares()
        surge_rows = [
            [prof.slots[i], _pct(shares[i])]
            for i in sorted(range(len(shares)), key=lambda i: shares[i], reverse=True)
            if shares[i] > 0
        ]
        lines.append("## 급등일의 고점 시간대" if markdown else "[급등일의 고점 시간대]")
        lines.extend(_table(("시간대", "고점확률"), surge_rows, markdown=markdown))
        lines.append("")

    entries = ", ".join(prof.slots[i] for i in prof.best_entry_slots(top))
    exits = ", ".join(prof.slots[i] for i in prof.best_exit_slots(top))
    lines.append("## 요약" if markdown else "[요약]")
    lines.append("- 저가가 가장 자주 나온 구간(매수 후보): {}".format(entries))
    lines.append("- 고가가 가장 자주 나온 구간(매도 후보): {}".format(exits))
    lines.append("")
    lines.append(
        "주의: 고점·저점 확률은 사후 통계입니다. 실제로는 고점이 찍히기 전에 그 사실을 알 수 없고, "
        "확률이 가장 높은 구간이라도 대개 30~40%를 넘지 않습니다. 수수료·세금·슬리피지 미반영. "
        "투자 판단과 그 결과는 본인 책임입니다."
    )
    return "\n".join(lines)
