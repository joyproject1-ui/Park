"""시간대별 통계와 매수/매도 시간대 조합 분석."""

from __future__ import annotations

import math
import statistics
from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Optional, Sequence, Tuple

from .model import Bar, Slot

BP = 10_000.0  # 1 = 10,000bp. 수익률은 bp(0.01%) 단위로 보고한다.


@dataclass(frozen=True)
class SlotBar:
    """하루 중 한 시간대를 하나의 봉으로 압축한 값."""

    day: date
    index: int
    open: float
    high: float
    low: float
    close: float
    volume: int

    @property
    def ret(self) -> float:
        return self.close / self.open - 1.0

    @property
    def range_pct(self) -> float:
        return (self.high - self.low) / self.open if self.open else 0.0


@dataclass
class SlotStat:
    """한 시간대의 여러 날짜에 걸친 요약 통계."""

    index: int
    label: str
    days: int
    mean_bp: float
    median_bp: float
    stdev_bp: float
    win_rate: float
    t_stat: float
    cumulative_bp: float
    volume_share: float
    range_bp: float

    @property
    def significant(self) -> bool:
        """표본이 최소한 있고 |t| >= 2 일 때만 '의미 있는 편향'으로 표시한다."""

        return self.days >= 10 and abs(self.t_stat) >= 2.0


@dataclass
class PairStat:
    """buy 시간대 시가에 사서 sell 시간대 종가에 판 결과."""

    buy: int
    sell: int
    days: int
    mean_bp: float
    win_rate: float
    stdev_bp: float
    t_stat: float


@dataclass
class Analysis:
    code: str
    slots: Sequence[Slot]
    days: List[date]
    slot_stats: List[SlotStat]
    grid: Dict[date, Dict[int, SlotBar]] = field(default_factory=dict)
    pair_stats: Dict[Tuple[int, int], PairStat] = field(default_factory=dict)
    overnight: Optional[SlotStat] = None
    intraday: Optional[SlotStat] = None
    incomplete_days: List[date] = field(default_factory=list)

    @property
    def n_days(self) -> int:
        return len(self.days)


def _mean(values: Sequence[float]) -> float:
    return statistics.fmean(values) if values else 0.0


def _stdev(values: Sequence[float]) -> float:
    return statistics.stdev(values) if len(values) > 1 else 0.0


def _t_stat(values: Sequence[float]) -> float:
    """단순 1표본 t 값. 표본이 적으면 그대로 0 을 돌려준다."""

    if len(values) < 3:
        return 0.0
    sd = _stdev(values)
    if sd == 0:
        return 0.0
    return _mean(values) / (sd / math.sqrt(len(values)))


def _slot_index(bar: Bar, slots: Sequence[Slot]) -> Optional[int]:
    last = len(slots) - 1
    for i, slot in enumerate(slots):
        if slot.contains(bar.clock, last=(i == last)):
            return i
    return None


def aggregate_slots(bars: Sequence[Bar], slots: Sequence[Slot]) -> Dict[date, Dict[int, SlotBar]]:
    """분봉을 (날짜, 시간대) 격자로 접는다."""

    buckets: Dict[date, Dict[int, List[Bar]]] = {}
    for bar in bars:
        idx = _slot_index(bar, slots)
        if idx is None:  # 시간외·장전 데이터는 버린다.
            continue
        buckets.setdefault(bar.day, {}).setdefault(idx, []).append(bar)

    grid: Dict[date, Dict[int, SlotBar]] = {}
    for day, per_slot in buckets.items():
        row: Dict[int, SlotBar] = {}
        for idx, group in per_slot.items():
            group.sort(key=lambda b: b.ts)
            row[idx] = SlotBar(
                day=day,
                index=idx,
                open=group[0].open,
                high=max(b.high for b in group),
                low=min(b.low for b in group),
                close=group[-1].close,
                volume=sum(b.volume for b in group),
            )
        grid[day] = row
    return grid


def analyze(
    bars: Sequence[Bar],
    slots: Sequence[Slot],
    *,
    code: str = "",
    require_full_day: bool = True,
) -> Analysis:
    grid = aggregate_slots(bars, slots)
    all_days = sorted(grid)

    complete = [d for d in all_days if len(grid[d]) == len(slots)]
    incomplete = [d for d in all_days if d not in set(complete)]
    days = complete if require_full_day else all_days
    if require_full_day and not days:
        # 데이터가 짧아 완전한 날이 없으면, 최소한 절반 이상 채워진 날을 쓴다.
        days = [d for d in all_days if len(grid[d]) >= len(slots) // 2]
        incomplete = [d for d in all_days if d not in set(days)]

    slot_stats: List[SlotStat] = []
    total_volume_per_day = {d: sum(sb.volume for sb in grid[d].values()) for d in days}

    for i, slot in enumerate(slots):
        rets = [grid[d][i].ret for d in days if i in grid[d]]
        ranges = [grid[d][i].range_pct for d in days if i in grid[d]]
        shares = [
            grid[d][i].volume / total_volume_per_day[d]
            for d in days
            if i in grid[d] and total_volume_per_day[d] > 0
        ]
        cumulative = 1.0
        for r in rets:
            cumulative *= 1.0 + r
        slot_stats.append(
            SlotStat(
                index=i,
                label=slot.label,
                days=len(rets),
                mean_bp=_mean(rets) * BP,
                median_bp=(statistics.median(rets) if rets else 0.0) * BP,
                stdev_bp=_stdev(rets) * BP,
                win_rate=(sum(1 for r in rets if r > 0) / len(rets)) if rets else 0.0,
                t_stat=_t_stat(rets),
                cumulative_bp=(cumulative - 1.0) * BP,
                volume_share=_mean(shares),
                range_bp=_mean(ranges) * BP,
            )
        )

    pair_stats: Dict[Tuple[int, int], PairStat] = {}
    for i in range(len(slots)):
        for j in range(i, len(slots)):
            rets = [
                grid[d][j].close / grid[d][i].open - 1.0
                for d in days
                if i in grid[d] and j in grid[d] and grid[d][i].open
            ]
            if not rets:
                continue
            pair_stats[(i, j)] = PairStat(
                buy=i,
                sell=j,
                days=len(rets),
                mean_bp=_mean(rets) * BP,
                win_rate=sum(1 for r in rets if r > 0) / len(rets),
                stdev_bp=_stdev(rets) * BP,
                t_stat=_t_stat(rets),
            )

    last = len(slots) - 1
    overnight_rets = []
    for prev_day, next_day in zip(days, days[1:]):
        prev_row, next_row = grid[prev_day], grid[next_day]
        if last in prev_row and 0 in next_row and prev_row[last].close:
            overnight_rets.append(next_row[0].open / prev_row[last].close - 1.0)
    intraday_rets = [
        grid[d][last].close / grid[d][0].open - 1.0 for d in days if 0 in grid[d] and last in grid[d] and grid[d][0].open
    ]

    return Analysis(
        code=code,
        slots=slots,
        days=days,
        slot_stats=slot_stats,
        grid={d: grid[d] for d in days},
        pair_stats=pair_stats,
        overnight=_summary_stat(-1, "오버나이트 (전일 종가→당일 시가)", overnight_rets),
        intraday=_summary_stat(-2, "장중 전체 (09:00 시가→15:30 종가)", intraday_rets),
        incomplete_days=incomplete,
    )


def _summary_stat(index: int, label: str, rets: Sequence[float]) -> Optional[SlotStat]:
    if not rets:
        return None
    cumulative = 1.0
    for r in rets:
        cumulative *= 1.0 + r
    return SlotStat(
        index=index,
        label=label,
        days=len(rets),
        mean_bp=_mean(rets) * BP,
        median_bp=statistics.median(rets) * BP,
        stdev_bp=_stdev(rets) * BP,
        win_rate=sum(1 for r in rets if r > 0) / len(rets),
        t_stat=_t_stat(rets),
        cumulative_bp=(cumulative - 1.0) * BP,
        volume_share=0.0,
        range_bp=0.0,
    )


def best_pairs(analysis: Analysis, *, top: int = 5, min_days: int = 10) -> List[PairStat]:
    pool = [p for p in analysis.pair_stats.values() if p.days >= min_days and p.buy != p.sell]
    return sorted(pool, key=lambda p: p.mean_bp, reverse=True)[:top]


def worst_pairs(analysis: Analysis, *, top: int = 5, min_days: int = 10) -> List[PairStat]:
    pool = [p for p in analysis.pair_stats.values() if p.days >= min_days and p.buy != p.sell]
    return sorted(pool, key=lambda p: p.mean_bp)[:top]
