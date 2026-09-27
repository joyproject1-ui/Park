"""수집한 경고장을 식약처 행정처분 메일과 같은 표 형식으로 만들어 SMTP로 보냅니다.

설정은 전부 환경변수로 받습니다. 비밀번호를 코드나 저장소에 두지 않기 위해서입니다.

    GMPAI_SMTP_HOST      기본 smtp.gmail.com
    GMPAI_SMTP_PORT      기본 587
    GMPAI_SMTP_USER      계정 (예: joyproject1@gmail.com)
    GMPAI_SMTP_PASSWORD  Gmail은 2단계 인증 후 발급한 '앱 비밀번호'
    GMPAI_SMTP_STARTTLS  기본 1 (0이면 465 SSL로 접속)
    GMPAI_MAIL_FROM      기본값은 GMPAI_SMTP_USER
    GMPAI_MAIL_TO        쉼표로 구분한 수신자
"""

from __future__ import annotations

import html
import os
import smtplib
import ssl
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from typing import Sequence

from .warningletters import WarningLetter

KST = timezone(timedelta(hours=9))

STYLE_TABLE = "border-collapse:collapse;width:100%;margin:4px 0 18px 0;font-size:13px;"
STYLE_TH = (
    "border:1px solid #d0d7de;background:#f2f5f8;padding:7px 10px;"
    "text-align:left;width:140px;vertical-align:top;white-space:nowrap;"
)
STYLE_TD = "border:1px solid #d0d7de;padding:7px 10px;vertical-align:top;line-height:1.55;"
STYLE_H3 = "font-size:15px;margin:26px 0 6px 0;color:#111;"
STYLE_SECTION = "color:#1f5f8b;font-size:13px;font-weight:bold;margin:12px 0 2px 0;"


class MailError(Exception):
    """메일 설정이 없거나 발송에 실패했을 때."""


@dataclass
class MailConfig:
    host: str
    port: int
    user: str
    password: str
    sender: str
    recipients: "list[str]"
    starttls: bool = True

    @classmethod
    def from_env(cls, env: "dict | None" = None) -> "MailConfig":
        source = os.environ if env is None else env
        user = (source.get("GMPAI_SMTP_USER") or "").strip()
        password = source.get("GMPAI_SMTP_PASSWORD") or ""
        recipients = [a.strip() for a in (source.get("GMPAI_MAIL_TO") or "").split(",") if a.strip()]
        missing = [
            name
            for name, value in (
                ("GMPAI_SMTP_USER", user),
                ("GMPAI_SMTP_PASSWORD", password),
                ("GMPAI_MAIL_TO", recipients),
            )
            if not value
        ]
        if missing:
            raise MailError("환경변수가 비어 있습니다: " + ", ".join(missing))
        return cls(
            host=(source.get("GMPAI_SMTP_HOST") or "smtp.gmail.com").strip(),
            port=int(source.get("GMPAI_SMTP_PORT") or 587),
            user=user,
            password=password,
            sender=(source.get("GMPAI_MAIL_FROM") or user).strip(),
            recipients=recipients,
            starttls=(source.get("GMPAI_SMTP_STARTTLS") or "1").strip() not in ("0", "false", "False"),
        )


# --------------------------------------------------------------------------- 본문

def _row(label: str, value: str, raw: bool = False) -> str:
    shown = value if raw else html.escape(value or "-")
    return f'<tr><th style="{STYLE_TH}">{html.escape(label)}</th><td style="{STYLE_TD}">{shown or "-"}</td></tr>'


def _letter_block(letter: WarningLetter, number: int) -> str:
    tag = "🔴 무균 관련" if letter.is_sterile_related else "· CGMP"
    link = f'<a href="{html.escape(letter.url)}">{html.escape(letter.url)}</a>' if letter.url else "-"
    keywords = ", ".join(letter.sterile_hits + letter.cgmp_hits)
    cfr = ", ".join(f"21 CFR {c}" for c in letter.cfr_citations)

    basic = "".join(
        [
            _row("업체명", letter.company),
            _row("발행 사무소", letter.office),
            _row("서한일자", letter.letter_date),
            _row("게시일자", letter.posted_date),
            _row("원문 링크", link, raw=True),
        ]
    )
    detail = "".join(
        [
            _row("해당 키워드", keywords),
            _row("인용 조항", cfr),
        ]
    )
    return (
        f'<h3 style="{STYLE_H3}">{number}. {html.escape(letter.company)} '
        f'<span style="font-weight:normal;color:#666;font-size:12px;">({tag})</span></h3>'
        f'<div style="{STYLE_SECTION}">■ 기본정보</div>'
        f'<table style="{STYLE_TABLE}">{basic}</table>'
        f'<div style="{STYLE_SECTION}">■ 지적 관련성</div>'
        f'<table style="{STYLE_TABLE}">{detail}</table>'
    )


def render_html(letters: "Sequence[WarningLetter]", since_date: str, source: str = "") -> str:
    sterile = [l for l in letters if l.is_sterile_related]
    header = (
        f'<h2 style="font-size:17px;margin:0 0 4px 0;">FDA Warning Letter 신규 {len(letters)}건 '
        f'({since_date.replace("-", "")} 이후)</h2>'
        f'<p style="font-size:13px;color:#444;margin:0 0 14px 0;">'
        f'무균·주사제·점안제 관련 <b>{len(sterile)}건</b> 포함. '
        f'분류는 본문 키워드 자동 판정이므로 원문 확인이 필요합니다.</p>'
    )
    body = "".join(_letter_block(letter, i) for i, letter in enumerate(letters, start=1))
    footer = (
        '<hr style="border:none;border-top:1px solid #e1e4e8;margin:26px 0 10px 0;">'
        f'<p style="font-size:11px;color:#888;line-height:1.6;">'
        f'출처: <a href="https://www.fda.gov/inspections-compliance-enforcement-and-criminal-investigations'
        f'/compliance-actions-and-activities/warning-letters">FDA Warning Letters</a>'
        f'{" (수집 경로: " + html.escape(source) + ")" if source else ""}<br>'
        f'생성: {datetime.now(KST).strftime("%Y-%m-%d %H:%M")} KST · gmpai letters</p>'
    )
    return (
        '<html><body style="font-family:\'Malgun Gothic\',AppleGothic,sans-serif;color:#222;">'
        f"{header}{body}{footer}</body></html>"
    )


def render_text(letters: "Sequence[WarningLetter]", since_date: str) -> str:
    lines = [f"FDA Warning Letter 신규 {len(letters)}건 ({since_date} 이후)", ""]
    for i, letter in enumerate(letters, start=1):
        tag = "[무균]" if letter.is_sterile_related else "[CGMP]"
        lines += [
            f"{i}. {tag} {letter.company}",
            f"   서한일자: {letter.letter_date or '-'} / 게시일자: {letter.posted_date or '-'}",
            f"   사무소: {letter.office or '-'}",
            f"   키워드: {', '.join(letter.sterile_hits + letter.cgmp_hits) or '-'}",
            f"   조항: {', '.join('21 CFR ' + c for c in letter.cfr_citations) or '-'}",
            f"   원문: {letter.url}",
            "",
        ]
    return "\n".join(lines)


def build_message(letters: "Sequence[WarningLetter]", config: MailConfig, since_date: str, source: str = "") -> EmailMessage:
    message = EmailMessage()
    message["Subject"] = f"[FDA] 신규 Warning Letter {len(letters)}건"
    message["From"] = config.sender
    message["To"] = ", ".join(config.recipients)
    message.set_content(render_text(letters, since_date))
    message.add_alternative(render_html(letters, since_date, source), subtype="html")
    return message


# --------------------------------------------------------------------------- 발송

def send(message: EmailMessage, config: MailConfig, smtp_factory=None) -> None:
    """SMTP로 발송합니다. smtp_factory는 테스트에서 갈아끼우기 위한 자리입니다."""
    try:
        if smtp_factory is not None:
            with smtp_factory(config.host, config.port) as server:
                _login_and_send(server, message, config)
            return
        if config.starttls:
            with smtplib.SMTP(config.host, config.port, timeout=60) as server:
                server.starttls(context=ssl.create_default_context())
                _login_and_send(server, message, config)
        else:
            with smtplib.SMTP_SSL(config.host, config.port, timeout=60, context=ssl.create_default_context()) as server:
                _login_and_send(server, message, config)
    except (smtplib.SMTPException, OSError) as exc:
        raise MailError(f"발송 실패: {exc}") from exc


def _login_and_send(server, message: EmailMessage, config: MailConfig) -> None:
    if config.password:
        server.login(config.user, config.password)
    server.send_message(message)
