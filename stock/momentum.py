"""'올려놓고 빠지는' 하루를 정량화한다.

시간대 평균 수익률만 보면 상승 타이밍이 지워진다. 아침에 올리고 오후에 흘리는 날과
그 반대인 날이 섞여서 평균이 0 근처로 눌리기 때문이다. 그래서 여기서는
 - 하루 고점이 실제로 몇 시에 찍히는지(고점 시간대 분포),
 - 시가→고점 상승폭 대비 고점→종가 되돌림(페이드)이 얼마나 되는지,
 - 한 시간대가 오르면 다음 시간대도 오르는지(모멘텀 지속성)
를 따로 센다.
"""

from __future__ import annotations

import statistics
from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Optional, Sequence

from .analysis import BP, Analysis, SlotBar


@dataclass
class SlotShape:
    """시간대별 '고점/저점이 여기서 나올 확률'과 모멘텀 지속성."""

    index: int
    label: str
    high_share: float          # 이 시간대에 당일 고가가 찍힌 날의 비율
    low_share: float           # 이 시간대에 당일 저가가 찍힌 날의 비율
    up_rate: float             # 이 시간대 수익률이 양수였던 비율
    follow_through: float      # 이 시간대가 올랐을 때 다음 시간대도 오른 비율
    follow_days: int           # follow_through 의 표본 수
    volume_share: float


@dataclass
class DayShape:
    """하루의 모양: 어디서 올리고 어디서 빠졌는가."""

    day: date
    open_price: float
    close_price: float
    high_slot: int
    low_slot: int
    runup_bp: float            # 시가 → 당일 고가
    fade_bp: float             # 당일 고가 → 종가 (보통 음수)
    day_bp: float              # 시가 → 종가

    @property
    def pump_and_fade(self) -> bool:
        """고점을 먼저 찍고 그 뒤에 저점이 온, 전형적인 '올리고 빠지는' 날."""

        return self.high_slot < self.low_slot


@dataclass
class MomentumProfile:
    slots: Sequence[str]
    shapes: List[SlotShape]
    days: List[DayShape]
    surge_days: List[DayShape]
    surge_threshold_bp: float

    @property
    def n_days(self) -> int:
        return len(self.days)

    @property
    def pump_and_fade_rate(self) -> float:
        if not self.days:
            return 0.0
        return sum(1 for d in self.days if d.pump_and_fade) / len(self.days)

    @property
    def median_runup_bp(self) -> float:
        return statistics.median([d.runup_bp for d in self.days]) if self.days else 0.0

    @property
    def median_fade_bp(self) -> float:
        return statistics.median([d.fade_bp for d in self.days]) if self.days else 0.0

    def surge_high_shares(self) -> List[float]:
        """급등일만 골랐을 때의 고점 시간대 분포."""

        if not self.surge_days:
            return [0.0] * len(self.slots)
        counts = [0] * len(self.slots)
        for day in self.surge_days:
            counts[day.high_slot] += 1
        return [c / len(self.surge_days) for c in counts]

    def best_exit_slots(self, top: int = 3) -> List[int]:
        """급등일 기준으로 고점이 가장 자주 찍히는 시간대 (없으면 전체 기준)."""

        shares = self.surge_high_shares() if self.surge_days else [s.high_share for s in self.shapes]
        ranked = sorted(range(len(shares)), key=lambda i: shares[i], reverse=True)
        return ranked[:top]

    def best_entry_slots(self, top: int = 3) -> List[int]:
        """저점이 가장 자주 찍히는 시간대."""

        ranked = sorted(self.shapes, key=lambda s: s.low_share, reverse=True)
        return [s.index for s in ranked[:top]]


def _day_shape(day: date, row: Dict[int, SlotBar], n_slots: int) -> Optional[DayShape]:
    present = [i for i in range(n_slots) if i in row]
    if not present:
        return None
    open_price = row[present[0]].open
    close_price = row[present[-1]].close
    if not open_price:
        return None
    high_slot = max(present, key=lambda i: row[i].high)
    low_slot = min(present, key=lambda i: row[i].low)
    high = row[high_slot].high
    low_of_day = row[low_slot].low
    return DayShape(
        day=day,
        open_price=open_price,
        close_price=close_price,
        high_slot=high_slot,
        low_slot=low_slot,
        runup_bp=(high / open_price - 1.0) * BP,
        fade_bp=(close_price / high - 1.0) * BP if high else 0.0,
        day_bp=(close_price / open_price - 1.0) * BP,
    )


def profile(analysis: Analysis, *, surge_threshold_bp: float = 300.0) -> MomentumProfile:
    """surge_threshold_bp: 시가 대비 고가가 이만큼 이상 오른 날을 '급등일'로 본다 (기본 +3%)."""

    n_slots = len(analysis.slots)
    shapes_days: List[DayShape] = []
    for day in analysis.days:
        shape = _day_shape(day, analysis.grid[day], n_slots)
        if shape is not None:
            shapes_days.append(shape)

    total = len(shapes_days)
    high_counts = [0] * n_slots
    low_counts = [0] * n_slots
    for day in shapes_days:
        high_counts[day.high_slot] += 1
        low_counts[day.low_slot] += 1

    shapes: List[SlotShape] = []
    for i in range(n_slots):
        ups = 0
        seen = 0
        follow_hits = 0
        follow_seen = 0
        for day in analysis.days:
            row = analysis.grid[day]
            if i not in row or not row[i].open:
                continue
            seen += 1
            rose = row[i].close > row[i].open
            ups += 1 if rose else 0
            if rose and (i + 1) in row and row[i + 1].open:
                follow_seen += 1
                follow_hits += 1 if row[i + 1].close > row[i + 1].open else 0
        shapes.append(
            SlotShape(
                index=i,
                label=analysis.slots[i].label,
                high_share=high_counts[i] / total if total else 0.0,
                low_share=low_counts[i] / total if total else 0.0,
                up_rate=ups / seen if seen else 0.0,
                follow_through=follow_hits / follow_seen if follow_seen else 0.0,
                follow_days=follow_seen,
                volume_share=analysis.slot_stats[i].volume_share,
            )
        )

    surge = [d for d in shapes_days if d.runup_bp >= surge_threshold_bp]
    return MomentumProfile(
        slots=[s.label for s in analysis.slots],
        shapes=shapes,
        days=shapes_days,
        surge_days=surge,
        surge_threshold_bp=surge_threshold_bp,
    )
