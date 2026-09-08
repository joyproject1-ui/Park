# -*- coding: utf-8 -*-
"""외부 전송 금지 스위치 — PQR 자료가 이 PC 를 떠나지 않게 한다.

담당자 2026-09-08: "내부 PC 로 자료를 읽으면 그럼 PQR 내용 Claude 외부로 유출되는 것은 아닌지?
보안에서 강력하다고 강조하면 좋을 듯해."

평소 흐름에서 PC 밖으로 나가는 것은 두 가지뿐이다 — ③ 글자 없는 스캔본을 Claude(PC 의
Claude Code 나 API 키)로 읽을 때 그 쪽 그림, ⑤ 검토 단계에서 보고서 본문(글)과 첨부.
글자 있는 자료(ERP·엑셀·성적서·기록서)는 파이썬이 PC 안에서 읽어 어디로도 보내지 않는다.

프로그램 폴더(또는 입력 폴더)에 `PQR 외부 전송 금지.txt` 를 두거나 환경 변수 PQR_OFFLINE=1 을
세우면 그 두 길도 막힌다: Claude 판독·검토를 부르지 않고, 스캔본은 이 PC 의 OCR 로 읽거나
(한글 모델이 있을 때) 담당자가 판독 파일을 직접 만들어 넣는다. 보고서는 그대로 나온다 —
못 읽은 칸이 '확인 필요' 로 남을 뿐이다.
"""
import os

OFFLINE_FILE = "PQR 외부 전송 금지.txt"
ENV = "PQR_OFFLINE"


def _program_root():
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def blocked(folder=None):
    """외부 전송을 막아 두었는가 — 표시 파일이나 환경 변수."""
    if os.environ.get(ENV, "").strip() in ("1", "true", "yes", "on"):
        return True
    for where in (folder, os.getcwd(), _program_root()):
        if where and os.path.isfile(os.path.join(where, OFFLINE_FILE)):
            return True
    return False


def note():
    """기록·화면에 적을 한 줄."""
    return ("외부 전송 금지 모드 — Claude 판독·검토를 부르지 않습니다 (%s)" % OFFLINE_FILE) if blocked() \
        else "Claude 로 가는 것: 글자 없는 스캔본 판독, 검토 단계의 본문·첨부 — 나머지는 PC 안에서만"
