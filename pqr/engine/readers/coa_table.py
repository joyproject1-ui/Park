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


def _cell_text(page, bbox):
    """칸 상자 안의 글 — 위아래 1pt 여유를 둔다(마지막 낱말이 괘선에 걸쳐 잘리곤 한다)."""
    if bbox is None:
        return ""
    x0, y0, x1, y1 = bbox
    try:
        text = page.crop((x0, max(0, y0 - 1), x1, y1 + 1)).extract_text() or ""
    except Exception:
        return ""
    # 낱말 가운데서 줄이 접힌 것('무색투명한 액' / '체')은 붙인다 — 다음 줄이 한두 글자뿐이면 앞 낱말의 꼬리다
    text = re.sub(r"\s*\n\s*(?=[가-힣]\s*(?:\n|$))", "", text)
    return re.sub(r"\s*\n\s*", " ", text).strip()


def _read_by_grid(page):
    """괘선 표를 칸 단위로 읽는다 — {항목: {"spec", "value"}}. 표를 못 찾으면 {}.

    시험항목 칸이 빈 줄은 윗 항목의 이어지는 줄이다. 시험기준이 '평균 :'·'개개 :' 로 시작하면
    하위 줄('질량·용량(평균)'), 아니면 같은 항목의 둘째 기준(이물검사)이라 '1) … 2) …' 로 잇는다
    (담당자 2026-09-09: "2가지를 기재할 때는 1) 육안으로 … 2) 이물이 … 이렇게 기재하면 돼").
    """
    for table in page.find_tables():
        rows = [[_cell_text(page, b) for b in row.cells] for row in table.rows]
        head = None
        for i, cells in enumerate(rows):
            flat = [_flat(c) for c in cells]
            if any(any(n in c for n in HEAD_ITEM) for c in flat) and \
               (any(any(n in c for n in HEAD_SPEC) for c in flat) or any(any(n in c for n in HEAD_VALUE) for c in flat)):
                head = i
                break
        if head is None:
            continue
        flat = [_flat(c) for c in rows[head]]
        ci = next((k for k, c in enumerate(flat) if any(n in c for n in HEAD_ITEM)), None)
        cs = next((k for k, c in enumerate(flat) if any(n in c for n in HEAD_SPEC)), None)
        cv = next((k for k, c in enumerate(flat) if any(n in c for n in HEAD_VALUE)), None)
        if ci is None or cs is None or cv is None or not (ci < cs < cv):
            continue
        out, name, lines_of = {}, "", {}
        for cells in rows[head + 1:]:
            item = cells[ci] if ci < len(cells) else ""
            spec = cells[cs] if cs < len(cells) else ""
            value = cells[cv] if cv < len(cells) else ""
            if any(_flat(item).startswith(n) for n in STOP) or any(_flat(c).startswith(n) for n in STOP for c in cells[:1]):
                break
            if item:
                name = item.strip()
            if not name or not (spec or value):
                continue
            m = SUB.match(spec)
            if m:
                key = "%s(%s)" % (name, m.group(1))
                out[key] = {"spec": m.group(2).strip(), "value": value}
                continue
            lines_of.setdefault(name, []).append((spec, value))
        for name, pairs in lines_of.items():
            if len(pairs) == 1:
                out[name] = {"spec": pairs[0][0], "value": pairs[0][1]}
            else:
                out[name] = {"spec": " ".join("%d) %s" % (k + 1, sp) for k, (sp, _) in enumerate(pairs)),
                             "value": " ".join("%d) %s" % (k + 1, v) for k, (_, v) in enumerate(pairs))}
                for k, (sp, v) in enumerate(pairs):        # 9.1 처럼 줄이 갈라진 표를 위해 따로도 둔다
                    out["%s %d)" % (name, k + 1)] = {"spec": sp, "value": v}
        for k in list(out):
            base = k.split("(")[0]
            if base != k and base not in out:
                out[base] = dict(out[k])
        return out
    return {}


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
            page = pdf.pages[page_no - 1]
            # 괘선이 있는 글자 PDF(ERP 공정시험성적서)는 **칸 상자 그대로** 읽는 것이 정확하다.
            # 낱말을 줄로 묶어 읽으면 두 줄로 접힌 칸('플라스틱 용기에 든 무색투명한 액 / 체')과
            # 시험자·시험일자 줄이 끼어들어 글이 뒤섞였다 (담당자 2026-09-09 나조린 충전 성적서:
            # 기밀도 '침투 없음 메틸렌블루시액의 침투 없음', 이물검사 '때 맑으며, 쉽게 육안으로 …').
            try:
                got = _read_by_grid(page)
            except Exception:
                got = {}
            if len(got) >= 2:
                return got
            words = page.extract_words(x_tolerance=1.5)
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
