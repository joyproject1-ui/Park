# -*- coding: utf-8 -*-
"""손글씨 안정성 시험일지(스캔 PDF)를 Claude 비전으로 읽어 13항 표 값으로 만든다.

담당자 지시: "손글씨도 최대한 작성해 주고 헷갈리는 것만 문의해줘". 그래서 판독은 하되
확신이 낮은 값은 issues(문의 목록)에 올리고 보고서에는 값 옆에 아무 표시도 하지 않는다 —
문의 목록이 곧 확인 대상이다. ANTHROPIC_API_KEY 가 있어야 돌아간다.
"""
import base64
import io
import json
import os
import re

MODEL = "claude-opus-5"

SCHEMA = {
    "type": "object",
    "properties": {
        "product_name": {"type": "string"},
        "lot": {"type": "string", "description": "제조번호 (예: OEV301)"},
        "test_type": {"type": "string", "enum": ["시판후", "장기", "가속", "기타"]},
        "mfg_date": {"type": "string", "description": "제조일자 YYYY.MM.DD, 모르면 빈 문자열"},
        "package": {"type": "string", "description": "포장형태 (예: 5g x Tube/갑)"},
        "storage": {"type": "string"},
        "points": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "label": {"type": "string", "description": "시점: 초기, 3M, 6M, 9M, 12M, 18M, 24M, 36M …"},
                    "test_date": {"type": "string", "description": "시험일자 YYYY.MM.DD, 없으면 빈 문자열"},
                    "assay": {"type": "string", "description": "함량 결과 (%), 없으면 빈 문자열"},
                    "reviewer_date": {"type": "string", "description": "확인자/팀장 서명일 YYYY.MM.DD, 없으면 빈 문자열"},
                    "tested": {"type": "boolean", "description": "그 시점에 시험 결과가 적혀 있으면 true"},
                    "confidence": {"type": "number", "description": "0~1, 손글씨 판독 확신도"},
                },
                "required": ["label", "test_date", "assay", "reviewer_date", "tested", "confidence"],
                "additionalProperties": False,
            },
        },
        "uncertain": {"type": "array", "items": {"type": "string"},
                      "description": "읽기 애매했던 곳을 사람이 확인할 수 있게 짧게 적는다"},
    },
    "required": ["product_name", "lot", "test_type", "mfg_date", "package", "storage", "points", "uncertain"],
    "additionalProperties": False,
}

PROMPT = (
    "이 이미지는 제약회사 안정성 시험 결과 기록지(손글씨 포함)입니다. 표의 머리(제품명·제조번호·시험구분·제조일자·"
    "포장형태·보관조건)와, 시험일자 행·함량(%) 행·결재(담당자/팀장/확인자 서명일) 행을 시점(초기, 3M, 6M … 36M)별로 읽어 "
    "JSON 으로 주세요. 사선으로 지워진 칸은 tested=false 입니다. 숫자는 보이는 그대로 적고, 확신이 낮은 값은 confidence 를 "
    "낮게 주고 uncertain 에 이유를 적으세요. 값을 지어내지 마세요."
)


def _pages_png(path, resolution=150):
    import pdfplumber
    out = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            im = page.to_image(resolution=resolution).original
            buf = io.BytesIO()
            im.convert("RGB").save(buf, format="PNG", optimize=True)
            out.append(base64.standard_b64encode(buf.getvalue()).decode("ascii"))
    return out


def read_page(client, png_b64):
    response = client.messages.create(
        model=MODEL, max_tokens=16000,
        output_config={"format": {"type": "json_schema", "schema": SCHEMA}},
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": png_b64}},
            {"type": "text", "text": PROMPT},
        ]}],
    )
    if response.stop_reason == "refusal":
        raise RuntimeError("판독 거부: %s" % getattr(response.stop_details, "explanation", ""))
    text = next(b.text for b in response.content if b.type == "text")
    return json.loads(text)


def _year(d):
    m = re.match(r"(\d{4})", d or "")
    return m.group(1) if m else ""


def _in_period(day, period):
    y = _year(day)
    return bool(y) and str(period.get("from", ""))[:4] <= y <= str(period.get("to", ""))[:4]


def build_tables(records, data, period):
    """판독 결과 → recipe._fill_stability 가 받는 stab 사전 + 문의 목록."""
    issues, stab = [], {"post_dom": [], "post_exp": [], "long_dom": [], "long_exp": [],
                        "trend_dom": [], "trend_exp": [], "points": {}}
    export_lots = set(data.export)
    by_lot = {}
    for r in records:
        by_lot.setdefault((r.get("lot") or "").strip(), []).append(r)
    trend = {"dom": {}, "exp": {}}
    for lot, recs in by_lot.items():
        if not lot:
            continue
        rec = recs[0]
        is_exp = lot in export_lots or "수출" in (rec.get("product_name") or "")
        side = "exp" if is_exp else "dom"
        pts = [p for r in recs for p in r.get("points", [])]
        for p in pts:
            if p.get("confidence", 1) < 0.7 and p.get("tested"):
                issues.append(("13", lot, "%s 시점 판독 확신 낮음 (함량 %s, 일자 %s)" % (p.get("label"), p.get("assay"), p.get("test_date") or p.get("reviewer_date"))))
        for u in rec.get("uncertain") or []:
            issues.append(("13", lot, u))
        done = [p for p in pts if p.get("tested")]
        in_year = [p for p in done if _in_period(p.get("reviewer_date") or p.get("test_date"), period)]
        year = _year(rec.get("mfg_date")) or "확인 필요"
        pack = rec.get("package") or "확인 필요"
        # 안정성 경향 파일(126-06)은 포장 규격마다 따로 만든다 — 그래서 여기서 나눠 담는다.
        pack_key = "수출용" if is_exp else "내수용"
        stab["points"].setdefault(pack_key, {})[lot] = {
            p["label"]: p.get("assay") for p in done if p.get("assay")}
        vals = [float(re.sub(r"[^\d.]", "", p["assay"])) for p in done if re.sub(r"[^\d.]", "", p.get("assay") or "")]
        if vals:
            key = (rec.get("test_type") or "기타", year)
            trend[side].setdefault(key, []).extend(vals)
        if rec.get("test_type") == "시판후":
            last = in_year[-1] if in_year else None
            if last:
                labels = [p["label"] for p in pts]
                final = labels and last["label"] == labels[-1] and all(p.get("tested") for p in pts)
                stab["post_%s" % side].append((str(len(stab["post_%s" % side]) + 1), year, last["label"], lot, pack,
                                               last.get("reviewer_date") or last.get("test_date") or "확인 필요",
                                               "완료" if final else "진행중"))
        elif rec.get("test_type") == "장기":
            if in_year:
                stab["long_%s" % side].append((str(len(stab["long_%s" % side]) + 1), year, lot, pack, "확인 필요",
                                               [(p["label"], p.get("reviewer_date") or p.get("test_date") or "") for p in in_year]))
                issues.append(("13.2", lot, "장기 안정성 실시 사유는 시험일지에 없음 — 확인 필요"))
    notes = {"dom": [], "exp": []}
    for side in ("dom", "exp"):
        rows = []
        for (kind, year), vals in sorted(trend[side].items(), key=lambda kv: (kv[0][0] != "시판후", kv[0][1])):
            rows.append(("시판 후" if kind == "시판후" else "장기", year, "%.1f ~ %.1f" % (min(vals), max(vals))))
        stab["trend_%s" % side] = rows
    return stab, issues


def read_stability_into(data, log=None):
    """writer 가 부르는 훅: data.stability_files 의 스캔 PDF 를 읽어 data.stability 에 넣는다."""
    log = log or (lambda *a: None)
    files = [p for p, scanned in getattr(data, "stability_files", []) if scanned]
    if not files:
        return None
    import anthropic
    client = anthropic.Anthropic()
    records = []
    for path in files:
        for i, png in enumerate(_pages_png(path), 1):
            try:
                rec = read_page(client, png)
                records.append(rec)
                log("  손글씨 판독: %s p%d → %s %s %d시점" % (os.path.basename(path), i, rec.get("lot"), rec.get("test_type"), len(rec.get("points", []))))
            except Exception as error:
                data.issues.append(("13", os.path.basename(path), "p%d 판독 실패: %s" % (i, error)))
    period = getattr(data, "period", None) or {"from": "", "to": ""}
    stab, issues = build_tables(records, data, period)
    data.stability = stab
    data.issues.extend(issues)
    return stab
