# -*- coding: utf-8 -*-
"""손글씨 안정성 시험일지(스캔 PDF)를 Claude 로 읽는다 — PC 판독(RapidOCR)보다 훨씬 빠르고 정확하다.

담당자 2026-09-07: "PC 로 판독하면 시간이 너무 오래 걸려 (24장에 40분) — Claude 로 판독하면 안 돼?"
ANTHROPIC_API_KEY 가 있으면 이 길을 쓰고, 없으면 PC 판독으로 간다.

read_logs() 가 돌려주는 꼴은 handwriting.read_folder() 와 똑같다 —
[{"lot", "year", "pack", "store", "kind", "market", "mfg", "expiry",
  "points": [{"period", "done", "assays": {성분: 값}, "unsure": [...]}], "source"}]
그래서 13.1·13.3·경향 엑셀을 채우는 길과 판독 파일(json)이 PC 판독과 완전히 같다.

값을 지어내지 않는다 — 확신이 낮은 칸은 unsure 에 담아 워드 노랑·엑셀 주황으로 표시된다.
"""
import base64
import io
import json
import os
import re

MODEL = "claude-opus-5"
DPI = 200
WORKERS = 4                       # 쪽마다 따로 물어보므로 몇 장을 한꺼번에 — 24장이 2~3분
LOW = 0.7                         # 이보다 낮은 확신은 '애매' 로 본다

SCHEMA = {
    "type": "object",
    "properties": {
        "product_name": {"type": "string", "description": "제품명. 없으면 빈 문자열"},
        "lot": {"type": "string", "description": "제조번호 (예: OEV301). 없으면 빈 문자열"},
        "test_type": {"type": "string", "enum": ["시판후", "장기", "가속", "기타"],
                      "description": "시험구분. 표에 적힌 그대로 고른다"},
        "market": {"type": "string", "enum": ["내수", "수출", ""],
                   "description": "제품명에 '수출용' 이 있으면 수출, '내수용' 이면 내수, 아니면 빈 문자열"},
        "mfg_date": {"type": "string", "description": "제조일자 YYYY.MM.DD, 모르면 빈 문자열"},
        "expiry_date": {"type": "string", "description": "사용기한 YYYY.MM.DD, 모르면 빈 문자열"},
        "package": {"type": "string", "description": "포장형태 (예: 5g/Tube). 없으면 빈 문자열"},
        "storage": {"type": "string", "description": "보관조건 (예: 25±2°C, 60±5%RH). 없으면 빈 문자열"},
        "points": {
            "type": "array",
            "description": "시점(세로 열) 하나가 하나. 시험하지 않은(사선·빈) 시점도 tested=false 로 담는다",
            "items": {
                "type": "object",
                "properties": {
                    "label": {"type": "string", "description": "시점: 초기, 3M, 6M, 9M, 12M, 18M, 24M, 36M …"},
                    "test_date": {"type": "string", "description": "시험일자 YYYY.MM.DD, 없으면 빈 문자열"},
                    "reviewer_date": {"type": "string", "description": "그 시점 결재(확인자·팀장) 일자 YYYY.MM.DD, 없으면 빈 문자열"},
                    "tested": {"type": "boolean", "description": "그 시점에 시험 결과가 적혀 있으면 true"},
                    "date_confidence": {"type": "number", "description": "0~1, 일자 판독 확신도"},
                    "assays": {
                        "type": "array",
                        "description": "함량 등 숫자 결과. 성분이 둘이면 둘 다 담는다",
                        "items": {
                            "type": "object",
                            "properties": {
                                "name": {"type": "string", "description": "성분 이름. 표에 없으면 '함량'"},
                                "value": {"type": "string", "description": "숫자 그대로 (예: 99.8). 없으면 빈 문자열"},
                                "confidence": {"type": "number", "description": "0~1, 손글씨 판독 확신도"},
                            },
                            "required": ["name", "value", "confidence"],
                            "additionalProperties": False,
                        },
                    },
                },
                "required": ["label", "test_date", "reviewer_date", "tested", "date_confidence", "assays"],
                "additionalProperties": False,
            },
        },
        "uncertain": {"type": "array", "items": {"type": "string"},
                      "description": "읽기 애매했던 곳을 사람이 확인할 수 있게 짧게 적는다"},
    },
    "required": ["product_name", "lot", "test_type", "market", "mfg_date", "expiry_date",
                 "package", "storage", "points", "uncertain"],
    "additionalProperties": False,
}

PROMPT = (
    "이 이미지는 제약회사 안정성 시험 결과 기록지(손글씨 포함)입니다. 표의 머리(제품명·제조번호·시험구분·제조일자·"
    "사용기한·포장형태·보관조건)와, 시점(초기·3M·6M·9M·12M·18M·24M·36M)마다 시험일자 행·함량(%) 행·결재 서명일 행을 "
    "읽어 JSON 으로 주세요. 사선으로 지워졌거나 비어 있는 시점은 tested=false 로 담습니다. "
    "숫자는 보이는 그대로 적고, 지어내지 마세요. 확신이 낮으면 confidence 를 낮게 주고 uncertain 에 이유를 적으세요."
)


def _client():
    import anthropic
    return anthropic.Anthropic()


def _png(pdf_path, page_no):
    from . import handwriting
    image = handwriting.render(pdf_path, DPI, page_no)
    buf = io.BytesIO()
    image.save(buf, format="PNG", optimize=True)
    return base64.standard_b64encode(buf.getvalue()).decode("ascii")


def read_page(client, png_b64):
    """한 쪽을 읽어 판독 결과(dict)를 돌려준다."""
    body = dict(model=MODEL, max_tokens=16000,
                output_config={"format": {"type": "json_schema", "schema": SCHEMA}},
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": png_b64}},
                    {"type": "text", "text": PROMPT},
                ]}])
    try:      # 안전 판정으로 거절되면 다른 모델이 이어 받게 한다 (서버 쪽 대체)
        response = client.beta.messages.create(betas=["server-side-fallback-2026-07-01"],
                                               fallbacks="default", **body)
    except TypeError:                       # 라이브러리가 오래되어 그 값을 모르면 그냥 부른다
        response = client.messages.create(**body)
    if getattr(response, "stop_reason", "") == "refusal":
        raise RuntimeError("판독 거부: %s" % getattr(getattr(response, "stop_details", None), "explanation", ""))
    text = next(b.text for b in response.content if b.type == "text")
    return json.loads(text)


# ---------------------------------------------------------------- 판독 결과를 판독 파일 꼴로
_DATE = re.compile(r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})")


def _date(text):
    m = _DATE.search(str(text or ""))
    return "%s.%02d.%02d" % (m.group(1), int(m.group(2)), int(m.group(3))) if m else ""


def _period(label):
    """'초기'·'Initial' → 'Initial', '3개월'·'3M' → '3M'."""
    text = re.sub(r"\s+", "", str(label or "")).upper()
    if not text:
        return ""
    if "초기" in text or text.startswith("INITIAL"):
        return "Initial"
    m = re.search(r"(\d{1,3})\s*(?:M|개월|MONTH)", text)
    return "%dM" % int(m.group(1)) if m else text


def _part(name, specs):
    """성분 이름을 그 제품 성적서(COA)의 이름에 맞춘다 — 못 맞추면 읽은 이름 그대로."""
    got = re.sub(r"[\s()（）]", "", str(name or ""))
    if not specs:
        return got or "함량"
    for part in specs:
        bare = re.sub(r"[\s()（）]", "", part)
        if bare and (bare in got or got in bare):
            return part
    if len(specs) == 1 and got in ("함량", "", "ASSAY"):
        return list(specs)[0]
    return got or "함량"


def to_log(rec, source, specs=None):
    """read_page 결과 하나 → 판독 파일(handwriting) 꼴의 기록 하나. 읽은 시점이 없으면 None."""
    lot = re.sub(r"\s+", "", str(rec.get("lot") or "")).upper()
    if not lot:
        return None
    points = []
    for p in rec.get("points") or []:
        if not p.get("tested"):
            continue
        period = _period(p.get("label"))
        if not period:
            continue
        assays, unsure = {}, []
        for a in p.get("assays") or []:
            value = re.sub(r"[^\d.]", "", str(a.get("value") or ""))
            if not value:
                continue
            part = _part(a.get("name"), specs)
            try:
                assays[part] = float(value)
            except ValueError:
                continue
            if float(a.get("confidence") or 0) < LOW:
                unsure.append(part)
        done = _date(p.get("reviewer_date")) or _date(p.get("test_date"))
        if not done or float(p.get("date_confidence") or 0) < LOW:
            unsure.append("done")
        if not assays and not done:
            continue
        points.append({"period": period, "done": done, "assays": assays, "unsure": sorted(set(unsure))})
    if not points:
        return None
    from . import handwriting
    points.sort(key=lambda p: handwriting.period_order(p["period"]))
    mfg = _date(rec.get("mfg_date"))
    name = str(rec.get("product_name") or "") + " " + os.path.basename(source)
    market = rec.get("market") or ("수출" if "수출" in name else ("내수" if "내수" in name else ""))
    kind = rec.get("test_type") if rec.get("test_type") in ("시판후", "장기") else "장기"
    return {"lot": lot, "year": mfg[:4], "pack": (rec.get("package") or "").strip(),
            "store": (rec.get("storage") or "").strip(), "kind": kind, "kind_sure": True,
            "market": market, "market_hint": bool(market), "mfg": mfg,
            "expiry": _date(rec.get("expiry_date")), "why": "", "points": points,
            "notes": list(rec.get("uncertain") or []), "source": os.path.basename(source)}


def read_logs(paths, specs=None, log=None, workers=WORKERS):
    """스캔 PDF 들을 Claude 로 읽어 handwriting.read_folder 와 같은 꼴의 목록을 돌려준다."""
    from concurrent.futures import ThreadPoolExecutor
    from . import handwriting
    say = log or (lambda *a: None)
    client = _client()
    jobs = []
    for path in paths:
        for page_no in range(handwriting.page_count(path)):
            jobs.append((path, page_no))
    say("    Claude 판독: %d장 (한 번에 %d장씩)" % (len(jobs), workers))
    out, done = [], [0]

    def one(job):
        path, page_no = job
        rec = read_page(client, _png(path, page_no))
        return path, page_no, rec

    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        for path, page_no, rec in pool.map(one, jobs):
            done[0] += 1
            try:
                got = to_log(rec, path, specs)
            except Exception as error:
                say("    판독 실패 %s p%d — %s" % (os.path.basename(path), page_no + 1, error))
                continue
            if got is None:
                say("    %s p%d: 읽을 시점이 없어 건너뜀" % (os.path.basename(path), page_no + 1))
                continue
            say("    Claude 판독 %d/%d: %s p%d → %s %s·%s 시점 %d개 (애매 %d칸)"
                % (done[0], len(jobs), os.path.basename(path), page_no + 1, got["lot"], got["kind"],
                   got["market"] or "구분 없음", len(got["points"]),
                   sum(len(p["unsure"]) for p in got["points"])))
            out.append(got)
    merged = []
    handwriting.merge_logs(merged, out, log)          # 한 Lot 이 여러 쪽에 걸쳐 있으면 합친다
    merged.sort(key=lambda r: (r.get("year") or "", r.get("lot") or ""))
    return merged


def read_stability_into(data, log=None):
    """writer 훅 — 13항 판독은 collect 에서 이미 끝냈으므로 여기서는 아무것도 하지 않는다.

    옛 판(2026-09 이전)은 이 훅이 data.stability(옛 꼴)를 만들었지만, 그 꼴로는 내수·수출이 갈린
    2026 서식의 13.1.1·13.1.2·13.3.x 를 채우지 못한다. 지금은 collect 가 PC 판독과 같은 길로
    Claude 판독을 받아 data.stability_logs 를 만든다.
    """
    return None

# ---------------------------------------------------------------- 변경요청서(스캔) 판독
CHANGE_PROMPT = """이 변경요청서(스캔 이미지)를 읽고 JSON 만 출력하세요.
보이는 대로만 적고, 안 보이면 빈 값으로 둡니다 — 지어내지 않습니다.
"actions" 는 '변경 실행 계획' 표의 부서별 조치사항입니다."""

CHANGE_SCHEMA = {
    "type": "object",
    "properties": {
        "doc_no": {"type": "string"},
        "title": {"type": "string"},
        "description": {"type": "string"},
        "reason": {"type": "string"},
        "products": {"type": "string"},
        "approved": {"type": "string"},
        "actions": {"type": "array", "items": {
            "type": "object",
            "properties": {"team": {"type": "string"}, "action": {"type": "string"}},
            "required": ["team", "action"], "additionalProperties": False}},
    },
    "required": ["doc_no", "title", "description", "reason", "products", "approved", "actions"],
    "additionalProperties": False,
}


def _change_page(client, png_b64):
    body = dict(model=MODEL, max_tokens=8000,
                output_config={"format": {"type": "json_schema", "schema": CHANGE_SCHEMA}},
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": png_b64}},
                    {"type": "text", "text": CHANGE_PROMPT},
                ]}])
    try:
        response = client.beta.messages.create(betas=["server-side-fallback-2026-07-01"],
                                               fallbacks="default", **body)
    except TypeError:
        response = client.messages.create(**body)
    if getattr(response, "stop_reason", "") == "refusal":
        raise RuntimeError("판독 거부")
    return json.loads(next(b.text for b in response.content if b.type == "text"))


def read_change(path, log=None, pages=3):
    """글자 없는 스캔 변경요청서를 Claude(API 키)로 읽는다 — readers.change.read_change 와 같은 꼴.

    담당자 2026-09-07: "못 읽으면 다른 방법을 사용해서라도 읽게 해야지, 공란으로 두면 안 돼."
    변경명·변경내용은 앞쪽에, 실행 계획은 그다음 쪽에 있으므로 앞 몇 쪽만 읽는다.
    """
    say = log or (lambda *a: None)
    from . import handwriting
    client = _client()
    n = min(pages, handwriting.page_count(path) or 1)
    out = {"doc_no": "", "title": "", "description": "", "reason": "", "products": "",
           "approved": "", "attachments": "", "target_date": "", "all_dates": [], "actions": []}
    seen = set()
    for page_no in range(1, n + 1):
        try:
            got = _change_page(client, _png(path, page_no))
        except Exception as error:
            say("    [12] %d쪽을 읽지 못했습니다 — %s" % (page_no, error))
            continue
        for key in ("doc_no", "title", "description", "reason", "products", "approved"):
            if not out[key] and str(got.get(key) or "").strip():
                out[key] = str(got[key]).strip()
        for one in got.get("actions") or []:
            team = str((one or {}).get("team") or "").strip()
            act = str((one or {}).get("action") or "").strip()
            if act and (team, act) not in seen:
                seen.add((team, act))
                out["actions"].append((team, act))
    say("    [12] %s — Claude(API 키)로 읽음: %s (조치 %d건)"
        % (os.path.basename(path), out["title"] or "제목 못 읽음", len(out["actions"])))
    return out

# ---------------------------------------------------------------- 완제/공정 시험성적서(스캔) 판독
COA_PROMPT = """이 시험성적서(스캔 이미지)를 읽고 JSON 만 출력하세요.
보이는 대로만 적고, 안 보이면 빈 값으로 둡니다 — 지어내지 않습니다.
· "assays" 는 함량 시험입니다. 성분 이름과 규격(하한~상한), 결과값을 짝으로 적습니다.
· "items" 는 시험항목 표를 줄마다 하나씩 그대로 옮긴 것입니다 — 확인시험·제제균일성·
  불용성미립자·불용성이물·질량·용량·기밀도·pH·삼투압·비중처럼 표에 있는 항목을 하나도
  빠뜨리지 말고 적습니다. "name" 은 표에 적힌 항목 이름 그대로, "spec" 은 기준 칸,
  "value" 는 결과 칸입니다.
· "sterility" 는 무균 시험(무균 | 음성) 결과, "bioburden" 은 생균수·바이오버든(CFU 숫자) 결과입니다 —
  둘을 바꿔 적지 않습니다. 무균 줄도 "items" 에 반드시 넣습니다.
· 숫자는 단위를 빼고 숫자만 적습니다."""

COA_SCHEMA = {
    "type": "object",
    "properties": {
        "lot": {"type": "string"}, "mfg_date": {"type": "string"}, "expiry": {"type": "string"},
        "appearance": {"type": "string"}, "verdict": {"type": "string"},
        "particle": {"type": "string"}, "particle_spec": {"type": "string"},
        "metal_total": {"type": "string"}, "metal_each": {"type": "string"},
        "bioburden": {"type": "string"}, "bioburden_spec": {"type": "string"},
        "sterility": {"type": "string"}, "sterility_spec": {"type": "string"},
        "assays": {"type": "array", "items": {
            "type": "object",
            "properties": {"part": {"type": "string"}, "lo": {"type": "string"},
                           "hi": {"type": "string"}, "value": {"type": "string"}},
            "required": ["part", "lo", "hi", "value"], "additionalProperties": False}},
        "items": {"type": "array", "items": {
            "type": "object",
            "properties": {"name": {"type": "string"}, "spec": {"type": "string"},
                           "value": {"type": "string"}},
            "required": ["name", "spec", "value"], "additionalProperties": False}},
    },
    "required": ["lot", "mfg_date", "expiry", "appearance", "verdict", "particle", "particle_spec",
                 "metal_total", "metal_each", "bioburden", "bioburden_spec", "sterility", "sterility_spec",
                 "assays", "items"],
    "additionalProperties": False,
}


def _coa_page(client, png_b64):
    body = dict(model=MODEL, max_tokens=8000,
                output_config={"format": {"type": "json_schema", "schema": COA_SCHEMA}},
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": png_b64}},
                    {"type": "text", "text": COA_PROMPT},
                ]}])
    try:
        response = client.beta.messages.create(betas=["server-side-fallback-2026-07-01"],
                                               fallbacks="default", **body)
    except TypeError:
        response = client.messages.create(**body)
    if getattr(response, "stop_reason", "") == "refusal":
        raise RuntimeError("판독 거부")
    return json.loads(next(b.text for b in response.content if b.type == "text"))


def read_coa(path, log=None, pages=3):
    """글자 없는 스캔 시험성적서를 Claude(API 키)로 읽는다 — readers.coa 와 같은 꼴."""
    say = log or (lambda *a: None)
    from . import handwriting
    client = _client()
    n = min(pages, handwriting.page_count(path) or 1)
    out = {"file": os.path.basename(path), "assays": [], "items": {}}
    본 = set()
    for page_no in range(1, n + 1):
        try:
            got = _coa_page(client, _png(path, page_no))
        except Exception as error:
            say("    [9.2] %d쪽을 읽지 못했습니다 — %s" % (page_no, error))
            continue
        for key in ("lot", "mfg_date", "expiry", "appearance", "verdict", "particle", "particle_spec",
                    "metal_total", "metal_each", "bioburden", "bioburden_spec", "sterility", "sterility_spec"):
            value = str(got.get(key) or "").strip()
            if value and not out.get(key):
                out[key] = value
        for one in got.get("assays") or []:
            part = str((one or {}).get("part") or "").strip()
            if part and part not in 본:
                본.add(part)
                out["assays"].append({"part": part, "lo": str(one.get("lo") or "").strip(),
                                      "hi": str(one.get("hi") or "").strip(),
                                      "value": str(one.get("value") or "").strip()})
        # 시험항목 표 — 확인시험·제제균일성·불용성미립자처럼 코드에 박아 두지 않은 항목이
        # 여기에 있다. 함량만 받아 오면 9.2 표의 그 열이 통째로 빈다(담당자 2026-09-08).
        for one in got.get("items") or []:
            name = str((one or {}).get("name") or "").strip()
            if name and name not in out["items"]:
                out["items"][name] = {"spec": str(one.get("spec") or "").strip(),
                                      "value": str(one.get("value") or "").strip()}
    if out["assays"]:
        out["assay"] = out["assays"][0]["value"]
        out["assay_spec"] = "%s ~ %s%%" % (out["assays"][0]["lo"], out["assays"][0]["hi"])
    say("    [9.2] %s — Claude(API 키)로 읽음: 함량 %d건 · 시험항목 %d건 (%s)"
        % (out["file"], len(out["assays"]), len(out["items"]),
           ", ".join(list(out["items"])[:8]) or "없음"))
    return out
