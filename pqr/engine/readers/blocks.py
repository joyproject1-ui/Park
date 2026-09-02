# -*- coding: utf-8 -*-
"""칸으로 짜인 서식(일탈보고서·변경요청서)을 '행 = (라벨, 내용 줄들)' 로 읽는다.

글자 순서로는 어느 칸의 글인지 알 수 없어(라벨이 칸 세로 가운데에 놓인다) 좌표를 쓴다:
표의 가로 괘선으로 행을 가르고, 문서번호 값이 시작하는 x 를 내용 열의 시작으로 삼아
그 왼쪽 낱말은 라벨, 오른쪽 낱말은 내용으로 나눈다. 영문 설명 낱말은 버린다.
"""
import re

_ENGLISH = re.compile(r"^[A-Za-z0-9()/,&.'\-]+$")
_DOC = re.compile(r"^(DR|CC|CAPA|OOS)-\d")


def _squash(s):
    return re.sub(r"\s+", "", s or "")


def _lines_of(words, tol=3.0):
    rows = []
    for w in sorted(words, key=lambda w: (round(w["top"]), w["x0"])):
        if rows and abs(rows[-1][0] - w["top"]) <= tol:
            rows[-1][1].append(w)
        else:
            rows.append([w["top"], [w]])
    return [(top, sorted(ws, key=lambda w: w["x0"])) for top, ws in rows]


def form_rows(page, content_x=None):
    """[(라벨, [내용 줄, ...]), ...] — 가로 괘선 사이가 한 행."""
    words = page.extract_words(x_tolerance=1.5, keep_blank_chars=False)
    if not words:
        return []
    if content_x is None:
        doc = next((w for w in words if _DOC.match(w["text"])), None)
        content_x = (doc["x0"] - 4) if doc else page.width * 0.25
    # 라벨 열만 가르는 짧은 괘선도 행 경계다 — 너비 12% 이상이면 모두 모으고 2pt 안은 하나로
    raw = sorted(e["top"] for e in page.edges
                 if e.get("orientation") == "h" and (e["x1"] - e["x0"]) > page.width * 0.12)
    edges = []
    for y in raw:
        if not edges or y - edges[-1] > 2:
            edges.append(y)
    if len(edges) < 2:
        edges = [0, page.height]
    bands = list(zip(edges, edges[1:]))
    rows = []
    for top, bottom in bands:
        inside = [w for w in words if top - 1 <= (w["top"] + w["bottom"]) / 2 <= bottom + 1]
        if not inside:
            continue
        def _line_text(ws):
            text = " ".join(w["text"] for w in ws).strip()
            # 영문만 있는 줄은 서식의 영문 설명(Description of …)이다. 한글·숫자가 섞인 줄은 내용.
            return "" if text and not re.search(r"[가-힣0-9]", text) else text
        label = " ".join(_line_text(ws) for _, ws in _lines_of([w for w in inside if w["x1"] <= content_x + 2]))
        content = []
        for _, ws in _lines_of([w for w in inside if w["x0"] >= content_x - 2]):
            text = _line_text(ws)
            if text:
                content.append(text)
        rows.append((label.strip(), content))
    return rows


def pick(rows, prefixes):
    """라벨이 prefixes 중 하나로 시작하는 행들의 내용 줄을 모아 돌려준다."""
    out = []
    for label, content in rows:
        if any(_squash(label).startswith(_squash(p)) for p in prefixes):
            out.extend(content)
    return out
