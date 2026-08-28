"""데이터 모델과 KRX 정규장 시간대 정의."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Iterable, Sequence

# KRX 정규장: 09:00 시작, 15:20~15:30 장 마감 동시호가, 15:30 종가 확정.
MARKET_OPEN = time(9, 0)
MARKET_CLOSE = time(15, 30)
CLOSING_AUCTION_START = time(15, 20)


@dataclass(frozen=True)
class Bar:
    """분봉 하나. ts 는 봉의 '시작 시각'(KST, naive)으로 정규화해 보관한다."""

    ts: datetime
    open: float
    high: float
    low: float
    close: float
    volume: int

    @property
    def day(self) -> date:
        return self.ts.date()

    @property
    def clock(self) -> time:
        return self.ts.time()


@dataclass(frozen=True)
class Slot:
    """[start, end) 로 정의되는 시간대. 마지막 시간대만 end 를 포함한다."""

    start: time
    end: time
    label: str

    def contains(self, t: time, *, last: bool = False) -> bool:
        if last:
            return self.start <= t <= self.end
        return self.start <= t < self.end


def _label(start: time, end: time, note: str = "") -> str:
    base = "{:02d}:{:02d}-{:02d}:{:02d}".format(start.hour, start.minute, end.hour, end.minute)
    return "{} {}".format(base, note).strip()


# 30분 기준 기본 구간. 마지막 두 구간은 KRX 마감 구조(15:20 동시호가)에 맞춰 쪼갠다.
DEFAULT_SLOTS: Sequence[Slot] = (
    Slot(time(9, 0), time(9, 30), _label(time(9, 0), time(9, 30), "개장")),
    Slot(time(9, 30), time(10, 0), _label(time(9, 30), time(10, 0))),
    Slot(time(10, 0), time(10, 30), _label(time(10, 0), time(10, 30))),
    Slot(time(10, 30), time(11, 0), _label(time(10, 30), time(11, 0))),
    Slot(time(11, 0), time(11, 30), _label(time(11, 0), time(11, 30))),
    Slot(time(11, 30), time(12, 0), _label(time(11, 30), time(12, 0))),
    Slot(time(12, 0), time(12, 30), _label(time(12, 0), time(12, 30), "점심")),
    Slot(time(12, 30), time(13, 0), _label(time(12, 30), time(13, 0), "점심")),
    Slot(time(13, 0), time(13, 30), _label(time(13, 0), time(13, 30))),
    Slot(time(13, 30), time(14, 0), _label(time(13, 30), time(14, 0))),
    Slot(time(14, 0), time(14, 30), _label(time(14, 0), time(14, 30))),
    Slot(time(14, 30), time(15, 0), _label(time(14, 30), time(15, 0))),
    Slot(time(15, 0), time(15, 20), _label(time(15, 0), time(15, 20))),
    Slot(time(15, 20), time(15, 30), _label(time(15, 20), time(15, 30), "동시호가")),
)


def build_slots(minutes: int) -> Sequence[Slot]:
    """지정한 분 단위로 09:00~15:30 을 균등 분할한다 (동시호가 구간은 항상 분리)."""

    if minutes <= 0:
        raise ValueError("시간대 길이는 1분 이상이어야 합니다")
    if minutes == 30:
        return DEFAULT_SLOTS

    slots = []
    cursor = datetime.combine(date(2000, 1, 1), MARKET_OPEN)
    auction = datetime.combine(date(2000, 1, 1), CLOSING_AUCTION_START)
    while cursor < auction:
        end = min(cursor + timedelta(minutes=minutes), auction)
        slots.append(Slot(cursor.time(), end.time(), _label(cursor.time(), end.time())))
        cursor = end
    slots.append(Slot(CLOSING_AUCTION_START, MARKET_CLOSE, _label(CLOSING_AUCTION_START, MARKET_CLOSE, "동시호가")))
    return tuple(slots)


def infer_bar_minutes(bars: Sequence[Bar]) -> int:
    """연속 봉 사이 간격의 최빈값으로 봉 주기를 추정한다 (기본 1분)."""

    gaps: dict[int, int] = {}
    for prev, cur in zip(bars, bars[1:]):
        if prev.day != cur.day:
            continue
        delta = int((cur.ts - prev.ts).total_seconds() // 60)
        if delta > 0:
            gaps[delta] = gaps.get(delta, 0) + 1
    if not gaps:
        return 1
    return max(gaps.items(), key=lambda kv: (kv[1], -kv[0]))[0]


def normalize_stamp(bars: Iterable[Bar], *, stamp: str, minutes: int) -> list[Bar]:
    """봉 시각 표기를 '시작 시각' 기준으로 맞춘다.

    네이버·대부분의 HTS 분봉은 봉의 '끝 시각'을 찍기 때문에(09:00~09:01 봉 → 09:01),
    그대로 두면 개장 30분 구간이 한 봉씩 밀린다.
    """

    if stamp == "start":
        return list(bars)
    if stamp != "end":
        raise ValueError("stamp 는 'start' 또는 'end' 여야 합니다")
    shift = timedelta(minutes=minutes)
    out = []
    for bar in bars:
        out.append(
            Bar(
                ts=bar.ts - shift,
                open=bar.open,
                high=bar.high,
                low=bar.low,
                close=bar.close,
                volume=bar.volume,
            )
        )
    return out
