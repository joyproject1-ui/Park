"""네이버 금융 분봉 조회 (표준 라이브러리만 사용).

사내망·에이전트 프록시 등에서 외부 접속이 막혀 있으면 FetchBlocked 로 명확히 알린다.
그 경우 HTS 에서 내보낸 CSV 를 `--csv` 로 넘겨 동일한 분석을 그대로 돌릴 수 있다.
"""

from __future__ import annotations

import re
import urllib.error
import urllib.request
from datetime import datetime
from typing import Dict, List

from .model import Bar

MINUTE_CHART_URL = "https://fchart.stock.naver.com/sise.nhn?symbol={code}&timeframe=minute&count={count}&requestType=0"
USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0 Safari/537.36 park-stock/1.0"
)
DEFAULT_TIMEOUT = 30

# 휴온스 그룹 상장사. 코드를 직접 넘겨도 된다.
TICKERS: Dict[str, str] = {
    "휴온스": "243070",
    "휴온스글로벌": "084110",
    "휴메딕스": "200670",
    "휴엠앤씨": "044060",
}

_ITEM_RE = re.compile(r'data="([^"]+)"')


class FetchBlocked(Exception):
    """네트워크 정책 등으로 시세 서버에 닿지 못했을 때."""


def resolve_code(name_or_code: str) -> str:
    text = name_or_code.strip()
    if re.fullmatch(r"\d{6}", text):
        return text
    if text in TICKERS:
        return TICKERS[text]
    raise ValueError("6자리 종목코드 또는 {} 중 하나를 넘기세요".format(", ".join(TICKERS)))


def parse_chart_xml(payload: str) -> List[Bar]:
    """<item data="202608280901|25200|25400|25100|25300|1234"/> 형태를 파싱한다."""

    bars: List[Bar] = []
    for raw in _ITEM_RE.findall(payload):
        parts = raw.split("|")
        if len(parts) < 6:
            continue
        stamp = parts[0].strip()
        if len(stamp) < 12 or not stamp[:12].isdigit():
            continue
        try:
            values = [float(p) for p in parts[1:6]]
        except ValueError:
            continue
        bars.append(
            Bar(
                ts=datetime.strptime(stamp[:12], "%Y%m%d%H%M"),
                open=values[0],
                high=values[1],
                low=values[2],
                close=values[3],
                volume=int(values[4]),
            )
        )
    bars.sort(key=lambda b: b.ts)
    return bars


def fetch_minutes(code: str, *, count: int = 2000, timeout: int = DEFAULT_TIMEOUT) -> List[Bar]:
    url = MINUTE_CHART_URL.format(code=resolve_code(code), count=max(1, count))
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Referer": "https://finance.naver.com/"})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            payload = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:  # pragma: no cover - 네트워크 의존
        raise FetchBlocked("네이버 금융이 {} 를 돌려줬습니다: {}".format(exc.code, url)) from exc
    except OSError as exc:  # pragma: no cover - 네트워크 의존
        raise FetchBlocked(
            "시세 서버에 접속하지 못했습니다({}). 사내망·프록시 차단이면 HTS CSV 를 내려받아 "
            "`python -m stock slots <파일.csv>` 로 분석하세요.".format(exc)
        ) from exc
    bars = parse_chart_xml(payload)
    if not bars:
        raise FetchBlocked("응답에 분봉이 없습니다. 종목코드나 접속 정책을 확인하세요: {}".format(url))
    return bars
