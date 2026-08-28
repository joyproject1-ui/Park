"""CSV 분봉 읽기 (HTS·네이버·직접 저장 파일 모두 수용)."""

from __future__ import annotations

import csv
import io
import re
from datetime import date, datetime, time
from pathlib import Path
from typing import Iterable, Sequence

from .model import Bar

# 컬럼 이름 후보. 소문자·공백제거 후 비교한다.
_ALIASES = {
    "datetime": ("datetime", "일시", "체결시각", "시각", "timestamp"),
    "date": ("date", "날짜", "일자", "거래일"),
    "time": ("time", "시간", "체결시간"),
    "open": ("open", "시가", "시작가"),
    "high": ("high", "고가"),
    "low": ("low", "저가"),
    "close": ("close", "종가", "현재가", "체결가", "price"),
    "volume": ("volume", "거래량", "체결량", "vol"),
}

_NUM_RE = re.compile(r"[^0-9.\-]")


class LoaderError(Exception):
    """CSV 를 분봉으로 해석할 수 없을 때."""


def _key(name: str) -> str:
    return name.strip().lstrip("﻿").replace(" ", "").lower()


def _resolve_columns(header: Sequence[str]) -> dict:
    normalized = {_key(h): h for h in header if h is not None}
    found = {}
    for field, aliases in _ALIASES.items():
        for alias in aliases:
            if alias in normalized:
                found[field] = normalized[alias]
                break
    if "close" not in found:
        raise LoaderError("종가(close) 컬럼을 찾지 못했습니다. 헤더: {}".format(list(header)))
    if "datetime" not in found and not ("date" in found and "time" in found):
        raise LoaderError("일시(datetime) 또는 날짜+시간 컬럼이 필요합니다. 헤더: {}".format(list(header)))
    return found


def _num(raw: str) -> float:
    text = _NUM_RE.sub("", (raw or "").strip())
    if not text or text in {"-", "."}:
        raise LoaderError("숫자로 해석할 수 없는 값: {!r}".format(raw))
    return float(text)


def parse_datetime(raw: str) -> datetime:
    text = (raw or "").strip().replace("T", " ")
    digits = re.sub(r"\D", "", text)
    if len(digits) >= 12:
        return datetime.strptime(digits[:12], "%Y%m%d%H%M")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M", "%Y.%m.%d %H:%M"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise LoaderError("일시를 해석할 수 없습니다: {!r}".format(raw))


def _parse_date(raw: str) -> date:
    digits = re.sub(r"\D", "", (raw or "").strip())
    if len(digits) == 8:
        return datetime.strptime(digits, "%Y%m%d").date()
    if len(digits) == 6:  # YYMMDD
        return datetime.strptime(digits, "%y%m%d").date()
    raise LoaderError("날짜를 해석할 수 없습니다: {!r}".format(raw))


def _parse_time(raw: str) -> time:
    digits = re.sub(r"\D", "", (raw or "").strip())
    if len(digits) == 6:
        return time(int(digits[0:2]), int(digits[2:4]), int(digits[4:6]))
    if len(digits) == 4:
        return time(int(digits[0:2]), int(digits[2:4]))
    raise LoaderError("시간을 해석할 수 없습니다: {!r}".format(raw))


def parse_rows(rows: Iterable[dict], columns: dict) -> list[Bar]:
    bars: list[Bar] = []
    for row in rows:
        if not any((v or "").strip() for v in row.values()):
            continue
        if "datetime" in columns:
            ts = parse_datetime(row[columns["datetime"]])
        else:
            ts = datetime.combine(_parse_date(row[columns["date"]]), _parse_time(row[columns["time"]]))
        close = _num(row[columns["close"]])
        open_ = _num(row[columns["open"]]) if "open" in columns else close
        high = _num(row[columns["high"]]) if "high" in columns else max(open_, close)
        low = _num(row[columns["low"]]) if "low" in columns else min(open_, close)
        try:
            volume = int(_num(row[columns["volume"]])) if "volume" in columns else 0
        except LoaderError:
            volume = 0
        bars.append(Bar(ts=ts, open=open_, high=high, low=low, close=close, volume=volume))
    bars.sort(key=lambda b: b.ts)
    return bars


def load_csv_text(text: str) -> list[Bar]:
    reader = csv.DictReader(io.StringIO(text))
    if reader.fieldnames is None:
        raise LoaderError("빈 CSV 입니다")
    columns = _resolve_columns(reader.fieldnames)
    bars = parse_rows(reader, columns)
    if not bars:
        raise LoaderError("분봉 데이터가 한 줄도 없습니다")
    return bars


def load_csv(path: Path | str) -> list[Bar]:
    """UTF-8(BOM 포함) 우선, 실패하면 CP949 로 다시 읽는다 (HTS 내보내기 대응)."""

    raw = Path(path).read_bytes()
    for encoding in ("utf-8-sig", "cp949", "utf-8"):
        try:
            return load_csv_text(raw.decode(encoding))
        except UnicodeDecodeError:
            continue
    raise LoaderError("파일 인코딩을 해석하지 못했습니다: {}".format(path))


def write_csv(path: Path | str, bars: Sequence[Bar]) -> None:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["datetime", "open", "high", "low", "close", "volume"])
        for bar in bars:
            writer.writerow(
                [
                    bar.ts.strftime("%Y-%m-%d %H:%M"),
                    "{:g}".format(bar.open),
                    "{:g}".format(bar.high),
                    "{:g}".format(bar.low),
                    "{:g}".format(bar.close),
                    bar.volume,
                ]
            )
