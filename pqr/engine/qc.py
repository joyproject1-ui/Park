# -*- coding: utf-8 -*-
"""QC-126 적용 조건 — 모든 PQR 에 같은 규칙.

담당자: "10 로트 미만은 Cpk 를 계산하지 않아. 이것도 모든 PQR 에 적용이 돼."
기준은 설정(`pqr/data/config.json` → thresholds.cpk_min_lots, 기본 10)에 한 번만 두고,
보고서 본문(9항 요약·16항 문안)·첨부 Cpk 계산 파일·대시보드가 모두 여기서 읽는다.
"""


def _thresholds():
    try:
        from ..build import load_config
        return load_config().get("thresholds") or {}
    except Exception:
        return {}


def cpk_min_lots():
    try:
        return int(_thresholds().get("cpk_min_lots", 10))
    except (TypeError, ValueError):
        return 10


def cpk_applies(n_lots, lsl=None, usl=None):
    """평가 년도 생산 Lot 이 기준 이상이면 True. 설정이 양쪽 규격만 허용하면 한쪽 규격 항목은 False."""
    if n_lots < cpk_min_lots():
        return False
    if _thresholds().get("cpk_two_sided_only") and (lsl is None or usl is None):
        return False
    return True


def not_applied_sentence(kind, n_lots):
    """16항 결론에 쓰는 표준 문안. kind: '내수용' / '수출용'."""
    return ("%s 제품은 평가 년도 생산 Lot 이 %d Lot 으로 %d Lot 미만이므로 ‘QC-126 제품품질평가 규정’에 "
            "따라 공정능력지수(Cpk)를 산출하지 않았음." % (kind, n_lots, cpk_min_lots()))
