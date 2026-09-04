# -*- coding: utf-8 -*-
"""제조번호(Lot No.)에서 제조 연도와 월을 읽는다.

담당자 설명(2026-09): 세 번째 글자가 제조 연도다 — X 는 2024년, Y 는 2025년, Z 는 2026년.
그리고 **PQR 은 그 다음 해 것**이다. 2025년에 만든 OGY301 은 2026년 PQR 이고, 2024년에 만든
OGX901 은 2025년 PQR(전년도 결재본 이름이 'PQR25' 인 까닭)이다.

네 번째 글자는 제조 월이다 — 1~9 는 그대로, O 는 10월, N 은 11월, D 는 12월
(OEYO01·OEYN01·OEYD01 로 확인).

글자는 스물여섯 해를 한 바퀴 돈다(Z = 2026 을 기준으로 거슬러 A = 2001). 한 바퀴를 넘어가면
같은 글자가 두 해를 가리키므로, 오늘에 가장 가까운 해를 고른다.
"""
import datetime
import re

ANCHOR_LETTER, ANCHOR_YEAR = "Z", 2026          # 담당자 확인: Z = 2026년 제조
CYCLE = 26
MONTHS = {"O": 10, "N": 11, "D": 12}
LOT = re.compile(r"^[A-Z0-9]{2}([A-Z])([1-9OND])\d{2}$")


def _parts(lot):
    m = LOT.match(str(lot or "").strip().upper())
    return (m.group(1), m.group(2)) if m else (None, None)


def made_year(lot, today=None):
    """이 Lot 을 만든 해. 읽을 수 없으면 None."""
    letter, _month = _parts(lot)
    if letter is None:
        return None
    base = ANCHOR_YEAR - (ord(ANCHOR_LETTER) - ord(letter))
    near = (today or datetime.date.today()).year - 1        # 평가하는 것은 대개 지난해 제조분
    return min((base + CYCLE * k for k in range(-2, 3)), key=lambda y: abs(y - near))


def made_month(lot):
    """이 Lot 을 만든 달. 읽을 수 없으면 None."""
    _letter, month = _parts(lot)
    if month is None:
        return None
    return MONTHS.get(month) or int(month)


def pqr_year(lot, today=None):
    """이 Lot 이 들어가는 PQR 의 연도 — 제조 연도의 다음 해. 읽을 수 없으면 None."""
    year = made_year(lot, today)
    return (year + 1) if year else None


def years(lots, today=None):
    """[(제조 연도, PQR 연도)] 가운데 가장 많이 나온 것. 읽을 수 없으면 (None, None).

    한 보고서의 Lot 은 모두 같은 해 것이어야 한다 — 섞여 있으면 가장 많은 해를 따르고,
    부르는 쪽이 나머지를 짚을 수 있게 한다.
    """
    seen = {}
    for lot in lots or []:
        year = made_year(lot, today)
        if year:
            seen[year] = seen.get(year, 0) + 1
    if not seen:
        return None, None
    best = max(sorted(seen), key=lambda y: seen[y])
    return best, best + 1


def odd_lots(lots, year, today=None):
    """제조 연도가 year 가 아닌 Lot 들 — 자료가 섞여 들어왔는지 짚어 준다."""
    return [lot for lot in lots or [] if made_year(lot, today) not in (None, year)]
