# -*- coding: utf-8 -*-
"""식약처 의약품 제품 허가정보를 받아 3항 '대상 제품' 표와 대조한다.

담당자 2026-09-12: "PQR 작성할 때 3항 내용 검토할 때 의약품안전나라에 접속해서 해당 제품명으로
검색 후 '재심사, RMP, 보험, 기타정보' 와 3항 표 내용을 비교해 줄 수 있나?" → "자동으로 해".

화면을 긁지 않고 공공데이터포털의 조회 서비스를 쓴다 — 화면 모양이 바뀌어도 깨지지 않는다.
열쇠(서비스 키)가 없거나 회사 밖으로 나가지 못하면 **아무 일도 하지 않고** 넘어간다. 허가사항은
담당자가 판단할 일이므로 값을 고치지 않는다. 어긋난 칸만 문의 목록에 올린다.

열쇠 두는 곳 (앞에 것부터 찾는다):
  · 환경 변수 `MFDS_API_KEY`
  · 입력 폴더의 `공통/식약처-허가정보-키.txt` 또는 제품 폴더의 같은 이름 파일
"""
import json
import os
import re

BASE = "http://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService06/getDrugPrdtPrmsnDtlInq05"
KEY_FILE = "식약처-허가정보-키.txt"
TIMEOUT = 20
MAX_ROWS = 20

# 응답 항목 이름은 서비스 판마다 조금씩 다르다 — 아는 이름을 모두 적어 두고 먼저 잡히는 것을 쓴다.
FIELDS = {
    "제품명": ("ITEM_NAME", "itemName"),
    "업체명": ("ENTP_NAME", "entpName"),
    "허가일자": ("ITEM_PERMIT_DATE", "itemPermitDate"),
    "품목기준코드": ("ITEM_SEQ", "itemSeq"),
    "전문일반": ("ETC_OTC_CODE", "ETC_OTC_NAME", "etcOtcCode"),
    "저장방법": ("STORAGE_METHOD", "storageMethod"),
    "사용기한": ("VALID_TERM", "validTerm"),
    "재심사대상": ("REEXAM_TARGET", "reexamTarget"),
    "재심사기간": ("REEXAM_DATE", "reexamDate"),
    "RMP대상": ("RMP_TARGET", "rmpTarget"),
    "성상": ("CHART", "chart"),
    "제형": ("PRDLST_STDR_CODE", "FORM_CODE_NAME", "formCodeName"),
    "포장정보": ("PACK_UNIT", "packUnit"),
    "주성분": ("MAIN_ITEM_INGR", "MATERIAL_NAME", "materialName"),
    "취소일자": ("CANCEL_DATE", "cancelDate"),
    "취소사유": ("CANCEL_NAME", "cancelName"),
}


def api_key(folder=None):
    """서비스 키 — 없으면 None."""
    got = (os.environ.get("MFDS_API_KEY") or "").strip()
    if got:
        return got
    for root in [folder, os.path.dirname(os.path.abspath(folder))] if folder else []:
        for path in (os.path.join(root or "", KEY_FILE), os.path.join(root or "", "공통", KEY_FILE)):
            try:
                with open(path, encoding="utf-8-sig") as handle:
                    got = handle.read().strip()
            except OSError:
                continue
            if got:
                return got
    return None


def _value(item, name):
    for key in FIELDS.get(name, ()):
        got = item.get(key)
        if got not in (None, ""):
            return str(got).strip()
    return ""


def _rows(payload):
    """응답에서 품목 목록만 꺼낸다 — 서비스 판마다 껍데기가 다르다.

    {"body":{"items":[…]}} · {"body":{"items":[{"item":{…}}]}} · {"items":{"item":[…]}} ·
    맨 위에 {"response":…} 가 한 겹 더 있는 것까지 모두 받는다.
    """
    got = payload
    for key in ("response", "body", "items", "item"):
        if isinstance(got, dict) and key in got:
            got = got[key]
    if isinstance(got, dict):
        got = [got]
    if not isinstance(got, list):
        return []
    out = []
    for one in got:                         # 목록 안에 다시 {"item": {…}} 로 싸인 판이 있다
        if isinstance(one, dict) and len(one) == 1 and "item" in one:
            inner = one["item"]
            out.extend(inner if isinstance(inner, list) else [inner])
        else:
            out.append(one)
    return out


def fetch(name, key, base=None, timeout=TIMEOUT, opener=None):
    """제품명으로 허가정보를 받아 [{항목: 값}] — 못 받으면 []."""
    import urllib.parse
    import urllib.request
    if not name or not key:
        return []
    query = urllib.parse.urlencode({
        "serviceKey": key, "item_name": name, "type": "json",
        "pageNo": "1", "numOfRows": str(MAX_ROWS)})
    url = "%s?%s" % (base or BASE, query)
    get = opener or (lambda u, t: urllib.request.urlopen(u, timeout=t).read())
    raw = get(url, timeout)
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", "replace")
    payload = json.loads(raw)
    out = []
    for item in _rows(payload):
        if not isinstance(item, dict):
            continue
        one = {k: _value(item, k) for k in FIELDS}
        if one.get("제품명"):
            out.append(one)
    return out


def _flat(text):
    return re.sub(r"[\s,./()（）~∼-]", "", str(text or "")).lower()


def _digits(text):
    return re.sub(r"\D", "", str(text or ""))


def pick(rows, name):
    """제품명이 가장 잘 맞는 한 건 — 없으면 None."""
    want = _flat(name)
    if not want:
        return None
    best = None
    for one in rows:
        got = _flat(one.get("제품명"))
        if not got:
            continue
        if got == want:
            return one
        if want in got or got in want:
            if best is None or len(got) < len(_flat(best.get("제품명"))):
                best = one
    return best


# 3항 표의 '점검 항목' → 허가정보 항목
MATCH = (("제품명", "제품명"), ("허가일자", "허가일자"), ("보관조건", "저장방법"),
         ("사용기한", "사용기한"), ("제품분류", "전문일반"), ("제형", "제형"))


def compare(section3, info):
    """3항 표({점검 항목: 내용})와 허가정보를 견줘 [(항목, 보고서 값, 허가정보 값, 까닭)].

    날짜는 숫자만, 그 밖은 빈칸·기호를 떼고 견준다. 허가정보 칸이 비어 있으면 견주지 않는다 —
    서비스가 그 항목을 주지 않는 것과 '해당 없음' 은 다르다.
    """
    out = []
    for 표항목, 허가항목 in MATCH:
        mine = str(section3.get(표항목) or "").strip()
        theirs = str(info.get(허가항목) or "").strip()
        if not mine or not theirs:
            continue
        if 표항목 == "허가일자":
            if _digits(mine) and _digits(theirs) and _digits(mine) != _digits(theirs):
                out.append((표항목, mine, theirs, "날짜가 다릅니다"))
            continue
        a, b = _flat(mine), _flat(theirs)
        if a and b and a not in b and b not in a:
            out.append((표항목, mine, theirs, "글이 다릅니다"))
    # 취소된 품목이면 크게 알린다
    if info.get("취소일자"):
        out.append(("허가 상태", "-", "%s %s" % (info.get("취소일자"), info.get("취소사유") or ""),
                    "허가정보에 취소 이력이 있습니다"))
    return out


def notes(info):
    """대조 결과와 함께 적어 둘 것 — 재심사·RMP 는 값이 없으면 '해당 없음' 으로 읽는다."""
    out = []
    for name in ("재심사대상", "RMP대상"):
        got = info.get(name) or ""
        out.append("%s: %s" % (name, got if got else "해당 없음"))
    if info.get("재심사기간"):
        out.append("재심사기간: %s" % info["재심사기간"])
    return out
