# -*- coding: utf-8 -*-
"""시험성적서의 '시험항목' 표를 **열 위치로** 읽는다.

정규식으로 '질량·용량 평균 : … g … g' 처럼 항목과 단위를 코드에 박아 두면 제품이 바뀔 때
통째로 빈다 — 담당자 2026-09-08 아이퓨어점안액에서 실제로 그랬다. 연고용 'g' 로 박혀 있어
점안액의 'mL' 을 못 읽었고, 9.2.3 표의 질량·용량 두 열이 세 Lot 모두 비었다.

그래서 항목 이름을 모른 채로도 읽는다: 글자의 x 자리를 보아 표의 세로 칸(시험항목·시험기준·
시험결과 …)을 찾고, 줄마다 그 칸에 넣는다. 시험항목 칸이 빈 줄은 윗 항목의 이어지는 줄이고,
시험기준 칸이 '평균 :' · '개개 :' 처럼 시작하면 그 항목의 하위 줄로 본다.
"""
import re

from ..pdftext import PdfTextError

# 칸과 칸 사이는 이만큼 넓게 벌어져 있다 — 한 칸 안 낱말 사이는 이보다 좁다
GAP = 8.0
ROW = 2.5                       # 같은 줄로 볼 세로 차이(pt)
SUB = re.compile(r"^([가-힣A-Za-z]{1,8})\s*[:：]\s*(.*)$")
HEAD_ITEM = ("시험항목", "항목")
HEAD_SPEC = ("시험기준", "기준", "규격")
HEAD_VALUE = ("시험결과", "결과")
STOP = ("비고", "시험완료일자", "판정결과", "확인일자")


def _lines(words):
    """[(top, [낱말 …])] — 세로 자리가 가까운 낱말끼리 한 줄로."""
    out = []
    for word in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if out and abs(word["top"] - out[-1][0]) <= ROW:
            out[-1][1].append(word)
        else:
            out.append((word["top"], [word]))
    return [(top, sorted(ws, key=lambda w: w["x0"])) for top, ws in out]


def _bounds(words):
    """낱말이 하나도 지나가지 않는 넓은 빈 띠를 찾아 칸 경계 x 목록을 만든다."""
    spans = sorted((w["x0"], w["x1"]) for w in words)
    if not spans:
        return []
    out, edge = [], spans[0][1]
    for x0, x1 in spans[1:]:
        if x0 - edge >= GAP:
            out.append((edge + x0) / 2.0)
        edge = max(edge, x1)
    return out


def _column(x, bounds):
    k = 0
    while k < len(bounds) and x >= bounds[k]:
        k += 1
    return k


def _flat(text):
    return re.sub(r"\s+", "", text or "")


def _head_index(cells, names):
    for k, text in cells.items():
        if any(n in _flat(text) for n in names):
            return k
    return None


def read_table(path, page_no=1):
    """{항목: {"spec": 기준, "value": 결과}} — 못 읽으면 {}.

    하위 줄이 있는 항목은 '질량·용량(평균)' · '질량·용량(개개)' 로 나눠 담고,
    나눠 담은 뒤에도 통째 이름('질량·용량')으로 첫 하위 줄을 찾을 수 있게 둔다.
    """
    try:
        import pdfplumber
    except ImportError:
        raise PdfTextError("PDF 를 읽으려면 pdfplumber 가 필요합니다: pip install pdfplumber")
    try:
        with pdfplumber.open(path) as pdf:
            if len(pdf.pages) < page_no:
                return {}
            words = pdf.pages[page_no - 1].extract_words(x_tolerance=1.5)
    except Exception:
        return {}
    lines = _lines(words)
    # 표 머리 줄 — '시험항목' 과 '시험기준'(또는 '시험결과')이 함께 있는 줄
    head = None
    for i, (_top, ws) in enumerate(lines):
        flat = _flat("".join(w["text"] for w in ws))
        if any(n in flat for n in HEAD_ITEM) and (any(n in flat for n in HEAD_SPEC)
                                                 or any(n in flat for n in HEAD_VALUE)):
            head = i
            break
    if head is None:
        return {}
    # 표 끝 — '비고'·'시험완료일자' 줄 앞까지
    stop = len(lines)
    for i in range(head + 1, len(lines)):
        flat = _flat("".join(w["text"] for w in lines[i][1]))
        if any(flat.startswith(n) for n in STOP):
            stop = i
            break
    body = [w for _top, ws in lines[head:stop] for w in ws]
    bounds = _bounds(body)
    if len(bounds) < 2:
        return {}

    def cells(ws):
        got = {}
        for w in ws:
            k = _column((w["x0"] + w["x1"]) / 2.0, bounds)
            got[k] = (got.get(k, "") + " " + w["text"]).strip()
        return got

    head_cells = cells(lines[head][1])
    ci = _head_index(head_cells, HEAD_ITEM)
    cs = _head_index(head_cells, HEAD_SPEC)
    cv = _head_index(head_cells, HEAD_VALUE)
    if ci is None or cs is None or cv is None or not (ci < cs < cv):
        return {}
    out, name, key = {}, "", ""
    for _top, ws in lines[head + 1:stop]:
        got = cells(ws)
        item, spec, value = got.get(ci, ""), got.get(cs, ""), got.get(cv, "")
        if item:
            name, key = item.strip(), item.strip()
        if not name or not (spec or value):
            continue                       # 시험자·시험일자만 있는 줄
        m = SUB.match(spec)
        if m:                              # '평균 : 0.50 ~ 0.57mL' — 그 항목의 하위 줄
            key = "%s(%s)" % (name, m.group(1))
            spec = m.group(2)
        one = out.setdefault(key, {"spec": "", "value": ""})
        one["spec"] = (one["spec"] + " " + spec).strip()
        one["value"] = (one["value"] + " " + value).strip()
    # 하위 줄만 있는 항목은 통째 이름으로도 찾을 수 있게 첫 하위 줄을 걸어 둔다
    for k in list(out):
        base = k.split("(")[0]
        if base != k and base not in out:
            out[base] = dict(out[k])
    return {k: v for k, v in out.items() if v["spec"] or v["value"]}
