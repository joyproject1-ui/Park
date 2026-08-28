"""PQR 판정에 쓰는 통계 계산 (표준 라이브러리만 사용).

여기서 계산하는 값은 모두 결정론적입니다 — 같은 입력이면 항상 같은 결과가 나오고,
계산에 쓰인 값(n, 평균, 표준편차, 규격)을 함께 돌려주므로 재현·검증할 수 있습니다.
"""

import math
import statistics

# 관리한계는 평균 ± 3σ 를 기본으로 봅니다 (내부 기준이 있으면 config 로 바꿉니다).
DEFAULT_SIGMA = 3.0
# Cpk 판정 기준 (사내 규정 QC-126): Cpk ≥ 1 공정능력 충분, 1 미만 공정능력 부족.
# 적용 대상은 10 Lot 이상 생산했고 상·하한 규격이 모두 설정된 시험항목입니다.
CPK_SUFFICIENT = 1.00
CPK_MIN_LOTS = 10
VERDICT_SUFFICIENT = "공정능력 충분"
VERDICT_INSUFFICIENT = "공정능력 부족"
VERDICT_NOT_APPLICABLE = "산출 대상 아님"


def describe(values):
    """표본 통계. 값이 없으면 n=0 만 채워 돌려줍니다."""
    clean = [float(v) for v in values if v is not None]
    result = {"n": len(clean), "mean": None, "sd": None, "min": None, "max": None}
    if not clean:
        return result
    result["mean"] = statistics.fmean(clean)
    result["min"] = min(clean)
    result["max"] = max(clean)
    result["sd"] = statistics.stdev(clean) if len(clean) > 1 else 0.0
    return result


def capability(values, lsl=None, usl=None, min_lots=1, two_sided_only=False,
               threshold=CPK_SUFFICIENT):
    """공정능력 Cp · Cpk.

    min_lots · two_sided_only 는 사내 규정의 적용 조건입니다. 조건을 채우지
    못하면 계산하지 않고 '산출 대상 아님' 과 그 이유를 돌려줍니다 — 조건 밖의
    값을 굳이 내놓으면 보고서에서 잘못 인용될 수 있기 때문입니다.
    """
    stats = describe(values)
    out = {"cp": None, "cpk": None, "verdict": VERDICT_NOT_APPLICABLE, "reason": ""}
    out.update(stats)
    if stats["n"] < max(2, min_lots):
        out["reason"] = ("%d Lot — 기준 %d Lot 미만" % (stats["n"], min_lots)
                         if min_lots > 1 else "표본이 2건 미만")
        return out
    if two_sided_only and (lsl is None or usl is None):
        out["reason"] = "상·하한 규격이 모두 설정된 항목만 적용"
        return out
    if lsl is None and usl is None:
        out["reason"] = "규격이 없음"
        return out
    sd = stats["sd"]
    if not sd:
        out["reason"] = "표준편차가 0"
        return out
    mean = stats["mean"]
    if lsl is not None and usl is not None:
        out["cp"] = (usl - lsl) / (6 * sd)
        out["cpk"] = min((usl - mean) / (3 * sd), (mean - lsl) / (3 * sd))
    elif usl is not None:
        out["cpk"] = (usl - mean) / (3 * sd)
    else:
        out["cpk"] = (mean - lsl) / (3 * sd)
    out["verdict"] = (VERDICT_SUFFICIENT if out["cpk"] >= threshold
                      else VERDICT_INSUFFICIENT)
    out["reason"] = ""
    return out


def control_limits(values, sigma=DEFAULT_SIGMA):
    """평균 ± nσ 관리한계와 이를 벗어난 값의 위치를 돌려줍니다."""
    stats = describe(values)
    if stats["n"] < 2 or not stats["sd"]:
        return {"lcl": None, "ucl": None, "out_of_trend": [], "n": stats["n"]}
    lcl = stats["mean"] - sigma * stats["sd"]
    ucl = stats["mean"] + sigma * stats["sd"]
    outliers = [index for index, value in enumerate(values)
                if value is not None and (value < lcl or value > ucl)]
    return {"lcl": lcl, "ucl": ucl, "out_of_trend": outliers, "n": stats["n"]}


def robust_limits(values, k=DEFAULT_SIGMA):
    """중앙값 ± k·MAD 기반 관리한계.

    평균 ± 3σ 는 큰 이탈값이 표준편차를 부풀려 자기 자신을 한계 안에 숨기는
    성질(masking)이 있습니다. 중앙값·MAD 는 이 영향을 거의 받지 않으므로
    같은 데이터로 한 번 더 확인하는 용도로 씁니다. 1.4826 은 정규분포에서
    MAD 를 표준편차와 같은 눈금으로 맞추는 상수입니다.
    """
    clean = [float(v) for v in values if v is not None]
    if len(clean) < 3:
        return {"lcl": None, "ucl": None, "out_of_trend": [], "n": len(clean)}
    center = statistics.median(clean)
    mad = statistics.median([abs(value - center) for value in clean])
    if not mad:
        return {"lcl": None, "ucl": None, "out_of_trend": [], "n": len(clean)}
    spread = k * 1.4826 * mad
    lcl, ucl = center - spread, center + spread
    outliers = [index for index, value in enumerate(values)
                if value is not None and (value < lcl or value > ucl)]
    return {"lcl": lcl, "ucl": ucl, "out_of_trend": outliers, "n": len(clean)}


def out_of_trend(values, sigma=DEFAULT_SIGMA):
    """두 관리한계(평균±3σ, 중앙값±3MAD) 중 하나라도 벗어난 값을 모읍니다.

    반환값: (관리한계 dict, {위치: 잡아낸 규칙 이름})
    """
    classic = control_limits(values, sigma)
    robust = robust_limits(values, sigma)
    flagged = {}
    for index in classic["out_of_trend"]:
        flagged[index] = "3σ"
    for index in robust["out_of_trend"]:
        flagged[index] = "3σ · MAD" if index in flagged else "MAD"
    return classic, flagged


def out_of_spec(values, lsl=None, usl=None):
    """규격을 벗어난 값의 위치."""
    hits = []
    for index, value in enumerate(values):
        if value is None:
            continue
        if lsl is not None and value < lsl:
            hits.append(index)
        elif usl is not None and value > usl:
            hits.append(index)
    return hits


def linear_slope(xs, ys):
    """최소제곱 기울기와 결정계수. 점이 3개 미만이면 None."""
    pairs = [(float(x), float(y)) for x, y in zip(xs, ys) if x is not None and y is not None]
    if len(pairs) < 3:
        return {"slope": None, "r2": None, "n": len(pairs)}
    xs_ = [p[0] for p in pairs]
    ys_ = [p[1] for p in pairs]
    mean_x = statistics.fmean(xs_)
    mean_y = statistics.fmean(ys_)
    sxx = sum((x - mean_x) ** 2 for x in xs_)
    if not sxx:
        return {"slope": None, "r2": None, "n": len(pairs)}
    sxy = sum((x - mean_x) * (y - mean_y) for x, y in pairs)
    slope = sxy / sxx
    syy = sum((y - mean_y) ** 2 for y in ys_)
    r2 = (sxy ** 2) / (sxx * syy) if syy else None
    return {"slope": slope, "r2": r2, "n": len(pairs)}


def stability_trend(timepoints, values, lsl=None, usl=None, horizon=36):
    """안정성 경향. 기울기를 규격까지 외삽해 유효기간 내 이탈 여부를 봅니다.

    부정적 경향(adverse)은 기울기가 규격 쪽으로 향하고, 관측 구간을 넘어
    horizon(기본 36개월) 안에 규격을 벗어날 것으로 계산될 때만 표시합니다.
    """
    fit = linear_slope(timepoints, values)
    result = {
        "slope": fit["slope"], "r2": fit["r2"], "n": fit["n"],
        "adverse": False, "months_to_limit": None, "limit": None,
    }
    slope = fit["slope"]
    if slope is None or abs(slope) < 1e-12:
        return result
    clean = [(float(t), float(v)) for t, v in zip(timepoints, values)
             if t is not None and v is not None]
    last_t, last_v = max(clean, key=lambda pair: pair[0])
    limit = usl if slope > 0 else lsl
    if limit is None:
        return result
    months = (limit - last_v) / slope
    if months <= 0:
        result["months_to_limit"] = 0.0
        result["limit"] = limit
        result["adverse"] = True
        return result
    result["months_to_limit"] = months
    result["limit"] = limit
    result["adverse"] = (last_t + months) <= horizon
    return result


def round_all(value, digits=3):
    """보고서·JSON 출력용 반올림 (dict · list 도 그대로 훑습니다)."""
    if isinstance(value, dict):
        return {key: round_all(item, digits) for key, item in value.items()}
    if isinstance(value, list):
        return [round_all(item, digits) for item in value]
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        return round(value, digits)
    return value
