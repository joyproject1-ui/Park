"""허용기준 문구를 상·하한 값으로 읽습니다.

성적서와 보고서의 허용기준은 한 칸에 두 기준이 함께 오는 일이 많습니다.

    허가) 90.0 ~ 110.0%
    자가) 92.0 ~ 108.0%

담당자 규칙 — **한쪽만 있으면 그 규격을, 허가와 자가를 함께 관리하면 허가 규격을**
Cpk 산출 기준으로 씁니다. 어느 기준으로 냈는지는 보고서에 함께 적습니다.
같은 자료에서 기준이 갈리면 Cpk 가 달라지기 때문입니다
(포장 pH: 자가 6.2~7.5 면 1.58, 허가 6.0~8.0 이면 2.89).
"""

import re

LICENSE = "허가"
IN_HOUSE = "자가"
PRIORITY = (LICENSE, IN_HOUSE)

_LABEL = re.compile(r"(허가|자가)\s*\)")
_NUMBER = r"[-+]?\d+(?:\.\d+)?"
_RANGE = re.compile(r"(%s)\s*(?:~|∼|-|―|—)\s*(%s)" % (_NUMBER, _NUMBER))
_AT_MOST = re.compile(r"(%s)\s*[^\s]*\s*이하" % _NUMBER)
_AT_LEAST = re.compile(r"(%s)\s*[^\s]*\s*이상" % _NUMBER)
_PAREN = re.compile(r"[(（][^)）]*[)）]")


def _limits(text):
    """한 문구에서 (하한, 상한). 정성 기준이면 (None, None)."""
    if not text:
        return None, None
    cleaned = _PAREN.sub(" ", text)          # "(말레인산페니라민: 3.0mg/mL)" 는 기준이 아닙니다
    # "300 ㎛이상/mL: 1 개 이하" 처럼 콜론이 있으면 뒤쪽이 실제 판정 기준입니다.
    if ":" in cleaned:
        cleaned = cleaned.rsplit(":", 1)[1]
    found = _RANGE.search(cleaned)
    if found:
        low, high = float(found.group(1)), float(found.group(2))
        return (low, high) if low <= high else (high, low)
    found = _AT_MOST.search(cleaned)
    if found:
        return None, float(found.group(1))
    found = _AT_LEAST.search(cleaned)
    if found:
        return float(found.group(1)), None
    return None, None


def parse(text):
    """허용기준 문구를 {기준 이름: (하한, 상한)} 으로 나눕니다.

    라벨이 없으면 {"": (하한, 상한)} 한 칸으로 돌려줍니다.
    """
    text = str(text or "").strip()
    if not text:
        return {}
    marks = list(_LABEL.finditer(text))
    if not marks:
        low, high = _limits(text)
        return {"": (low, high)} if (low, high) != (None, None) else {}
    found = {}
    for index, mark in enumerate(marks):
        end = marks[index + 1].start() if index + 1 < len(marks) else len(text)
        low, high = _limits(text[mark.end():end])
        if (low, high) != (None, None):
            found[mark.group(1)] = (low, high)
    return found


def choose(text, priority=PRIORITY):
    """Cpk 를 낼 기준을 고릅니다. 반환값: (하한, 상한, 쓴 기준 이름).

    허가와 자가가 함께 있으면 허가를, 하나만 있으면 그것을 씁니다.
    """
    found = parse(text)
    if not found:
        return None, None, ""
    if len(found) == 1:
        name, (low, high) = next(iter(found.items()))
        return low, high, name
    for name in priority:
        if name in found:
            low, high = found[name]
            return low, high, name
    name, (low, high) = next(iter(found.items()))
    return low, high, name
