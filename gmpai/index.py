"""Generate human-readable indexes (Markdown and a standalone HTML page)."""

from __future__ import annotations

import html
from datetime import datetime, timezone
from pathlib import Path
from typing import Sequence

from .catalog import Catalog, Document, group_by_authority

STATUS_LABELS = {
    "draft": "초안 (draft)",
    "final": "최종 (final)",
    "in-force": "발효 (in force)",
    "discussion-paper": "논의문서 (discussion paper)",
}

CATEGORY_LABELS = {
    "gmp-ai": "GMP × AI",
    "computerised-systems": "전산화 시스템",
    "data-integrity": "데이터 무결성",
    "medical-device-ai": "의료기기 AI",
    "horizontal-law": "일반 법령",
}


def _status_label(status: str) -> str:
    return STATUS_LABELS.get(status, status)


def _category_label(category: str) -> str:
    return CATEGORY_LABELS.get(category, category)


def _timestamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")


def render_markdown(catalog: Catalog, documents: Sequence[Document] | None = None) -> str:
    docs = list(documents if documents is not None else catalog.documents)
    lines = [
        "# GMP × AI 규정 문서 목록",
        "",
        f"카탈로그 최종 검토일: {catalog.last_reviewed} · 문서 {len(docs)}건 · 생성 {_timestamp()}",
        "",
        "각 문서는 규제기관 공식 페이지에서 직접 내려받습니다. 링크가 바뀌면 `landing page`를 먼저 확인하세요.",
        "",
    ]
    for authority, group in group_by_authority(docs).items():
        lines.append(f"## {authority}")
        lines.append("")
        lines.append("| 문서 | 분류 | 상태 | 일자 | 다운로드 | 출처 페이지 |")
        lines.append("| --- | --- | --- | --- | --- | --- |")
        for doc in group:
            download = f"[PDF]({doc.direct_url})" if doc.direct_url else "출처 페이지에서 자동 탐색"
            lines.append(
                f"| **{doc.title_ko}**<br>{doc.title} | {_category_label(doc.category)} | "
                f"{_status_label(doc.status)} | {doc.document_date or '-'} | {download} | "
                f"[landing page]({doc.landing_page}) |"
            )
        lines.append("")
    lines.append("## 참고")
    lines.append("")
    for doc in docs:
        if doc.notes:
            lines.append(f"- **{doc.id}** ({doc.reference or doc.authority}): {doc.notes}")
    lines.append("")
    return "\n".join(lines)


_HTML_TEMPLATE = """<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>GMP × AI 규정 다운로드</title>
<style>
  :root {{
    color-scheme: light dark;
    --bg: #f6f7f9;
    --card: #ffffff;
    --text: #16181d;
    --muted: #5c6270;
    --line: #e2e5ea;
    --accent: #1a56c4;
    --chip: #eef1f6;
  }}
  @media (prefers-color-scheme: dark) {{
    :root {{
      --bg: #14161a;
      --card: #1c1f25;
      --text: #e8eaee;
      --muted: #9aa1ae;
      --line: #2c3038;
      --accent: #7aa7ff;
      --chip: #262a32;
    }}
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0; padding: 2rem 1rem 4rem;
    background: var(--bg); color: var(--text);
    font: 16px/1.6 -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans KR", sans-serif;
  }}
  main {{ max-width: 1040px; margin: 0 auto; }}
  h1 {{ font-size: 1.6rem; margin: 0 0 .35rem; }}
  .sub {{ color: var(--muted); font-size: .9rem; margin-bottom: 2rem; }}
  h2 {{ font-size: 1.15rem; margin: 2.2rem 0 .9rem; padding-bottom: .4rem; border-bottom: 1px solid var(--line); }}
  .doc {{
    background: var(--card); border: 1px solid var(--line); border-radius: 10px;
    padding: 1rem 1.1rem; margin-bottom: .75rem;
  }}
  .doc h3 {{ font-size: 1rem; margin: 0 0 .2rem; }}
  .doc .en {{ color: var(--muted); font-size: .85rem; margin: 0 0 .6rem; }}
  .chips {{ display: flex; flex-wrap: wrap; gap: .4rem; margin-bottom: .6rem; }}
  .chip {{
    background: var(--chip); border-radius: 999px; padding: .12rem .6rem;
    font-size: .75rem; color: var(--muted); white-space: nowrap;
  }}
  .links {{ display: flex; flex-wrap: wrap; gap: 1rem; font-size: .9rem; }}
  a {{ color: var(--accent); }}
  .notes {{ color: var(--muted); font-size: .82rem; margin: .6rem 0 0; }}
  footer {{ margin-top: 3rem; color: var(--muted); font-size: .8rem; }}
  code {{ background: var(--chip); padding: .1rem .35rem; border-radius: 4px; font-size: .85em; }}
</style>
</head>
<body>
<main>
<h1>GMP × AI 규정 다운로드</h1>
<p class="sub">EU(EC · EMA · EUR-Lex)와 미국 FDA의 공식 원문 링크 · 카탈로그 검토일 {reviewed} · 생성 {generated}</p>
{sections}
<footer>
<p>모든 링크는 규제기관 공식 사이트를 가리킵니다. 일괄 내려받기: <code>python -m gmpai download</code></p>
</footer>
</main>
</body>
</html>
"""


def render_html(catalog: Catalog, documents: Sequence[Document] | None = None) -> str:
    docs = list(documents if documents is not None else catalog.documents)
    sections: list[str] = []
    for authority, group in group_by_authority(docs).items():
        cards: list[str] = []
        for doc in group:
            chips = "".join(
                f'<span class="chip">{html.escape(text)}</span>'
                for text in (
                    _category_label(doc.category),
                    _status_label(doc.status),
                    doc.document_date or "",
                    doc.reference or "",
                    doc.issuer or "",
                )
                if text
            )
            links = [f'<a href="{html.escape(doc.landing_page)}" rel="noreferrer">출처 페이지</a>']
            if doc.direct_url:
                links.insert(0, f'<a href="{html.escape(doc.direct_url)}" rel="noreferrer">PDF 직접 다운로드</a>')
            notes = f'<p class="notes">{html.escape(doc.notes)}</p>' if doc.notes else ""
            cards.append(
                '<article class="doc">'
                f"<h3>{html.escape(doc.title_ko)}</h3>"
                f'<p class="en">{html.escape(doc.title)}</p>'
                f'<div class="chips">{chips}</div>'
                f'<div class="links">{" ".join(links)}</div>'
                f"{notes}"
                "</article>"
            )
        sections.append(f"<h2>{html.escape(authority)}</h2>\n" + "\n".join(cards))

    return _HTML_TEMPLATE.format(
        reviewed=html.escape(catalog.last_reviewed),
        generated=_timestamp(),
        sections="\n".join(sections),
    )


def write_indexes(
    catalog: Catalog,
    documents: Sequence[Document] | None = None,
    markdown_path: Path | str = "INDEX.md",
    html_path: Path | str = "docs/index.html",
) -> list[Path]:
    written: list[Path] = []
    md = Path(markdown_path)
    md.parent.mkdir(parents=True, exist_ok=True)
    md.write_text(render_markdown(catalog, documents), encoding="utf-8")
    written.append(md)

    page = Path(html_path)
    page.parent.mkdir(parents=True, exist_ok=True)
    page.write_text(render_html(catalog, documents), encoding="utf-8")
    written.append(page)
    return written
