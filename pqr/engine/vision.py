# -*- coding: utf-8 -*-
"""손글씨 안정성시험일지 판독 (Claude). API 키가 있을 때만 켜진다.

키는 두 가지 길로 준다 — 둘 다 이 PC 안에만 있는다.
  · 환경 변수 ANTHROPIC_API_KEY
  · 프로그램 폴더의 `API-KEY.txt` (키 한 줄만 적어 둔다)

담당자 2026-09-07: "PC 로 판독하면 시간이 너무 오래 걸려" — 키가 있으면 Claude 가 읽어
24장이 40분에서 2~3분이 된다. 키가 없으면 PC 가 오프라인(RapidOCR)으로 읽는다.
"""
import os

KEY_FILE = "API-KEY.txt"


def _root():
    """프로그램 폴더 (pqr/engine/vision.py → …/)."""
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def key_from_file(root=None):
    """API-KEY.txt 에 적힌 키 — 없으면 빈 문자열. 앞뒤 공백·따옴표·BOM 을 떼어 낸다."""
    path = os.path.join(root or _root(), KEY_FILE)
    try:
        with open(path, encoding="utf-8-sig") as handle:
            for line in handle:
                key = line.strip().strip('"').strip("'")
                if key and not key.startswith("#"):
                    return key
    except OSError:
        pass
    return ""


def available():
    """Claude 로 읽을 수 있는가. 파일에 키가 있으면 이 실행에 한해 환경 변수로 세워 둔다."""
    if os.environ.get("ANTHROPIC_API_KEY", "").strip():
        return True
    key = key_from_file()
    if key:
        os.environ["ANTHROPIC_API_KEY"] = key
        return True
    return False


def hook():
    from .vision_claude import read_stability_into
    return read_stability_into
