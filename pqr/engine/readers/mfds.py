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
# 포털이 http 주소를 막거나 서비스 판이 바뀌면 HTTP 400/404 가 난다 (담당자 PC 2026-09-16:
# "식약처 허가정보를 받지 못했습니다 — HTTP Error 400: Bad Request"). 담당자가 포털의
# '요청 주소' 를 `공통/식약처-허가정보-주소.txt` 에 넣으면 그것을 먼저 쓰고, 없으면 아래를
# 차례로 두드려 통한 것을 기억한다.
LICENSE_URL_FILE = "식약처-허가정보-주소.txt"
LICENSE_BASES = (
    BASE,
    BASE.replace("http://", "https://", 1),
    "https://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07/getDrugPrdtPrmsnDtlInq06",
    "http://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService05/getDrugPrdtPrmsnDtlInq04",
)
KEY_FILE = "식약처-허가정보-키.txt"
TIMEOUT = 20
PROBE_TIMEOUT = 6          # 주소 후보를 두드릴 때는 짧게 기다린다 — 아닌 주소에 20초씩 매달리지 않는다
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


def normalize_url(text):
    """파일에 적힌 글에서 요청 주소를 꺼낸다 — 없으면 빈 글.

    · `http://`·`https://` 로 시작하면 그대로. `apis.data.go.kr/…` 처럼 앞을 떼고 적었으면
      `http://` 를 붙인다 (담당자 PC 2026-09-16: 주소를 넣었는데 "http:// 로 시작하는 줄" 이
      없다고 나왔다).
    · 포털의 샘플 주소(`…?serviceKey=…&…`)를 통째로 붙여 넣었으면 `?` 뒤는 뗀다.
    · 줄이 여럿이면 주소로 보이는 첫 줄을 쓴다 — 열쇠 줄과 섞여 있어도 된다.
    """
    for line in str(text or "").replace("\ufeff", "").splitlines():
        got = line.strip().strip('"').strip("'")
        if not got:
            continue
        low = got.lower()
        if low.startswith(("http://", "https://")):
            return got.split("?")[0].strip()
        if re.match(r"^(apis?|www)\.data\.go\.kr/", low) or low.startswith("data.go.kr/"):
            return ("http://" + got).split("?")[0].strip()
    return ""


def _looks_like_url(text):
    """주소인지 열쇠인지 가린다.

    파일 두 개가 나란히 있어 서로 바꿔 넣기 쉽다 — 담당자 2026-09-14: "둘다 키주소는
    동일하네". 인증키는 계정당 하나라 두 서비스가 같은 키를 쓰지만, 행정처분 파일에는
    키가 아니라 요청 주소가 들어가야 한다.
    """
    return bool(normalize_url(text))


KEY_LABELS = ("인증키", "서비스키", "service key", "servicekey", "end point", "endpoint", "주소", "key")


def extract_key(text):
    """열쇠 파일 글에서 열쇠만 — 없으면 빈 글.

    담당자 PC 2026-09-16: 파일에 'End Point / https://…Service07 / 일반 인증키 / 44fc…' 네 줄을
    넣어 두었다. 전에는 파일 전체를 열쇠로 보내 포털이 HTTP 400 을 돌려줬다. 이름표 줄과 주소
    줄은 빼고, 열쇠처럼 생긴 줄(빈칸 없이 길고 글자·숫자·+/=% 뿐)을 고른다. 여럿이면 가장 긴 것.
    """
    best = ""
    for line in str(text or "").replace("\ufeff", "").splitlines():
        got = line.strip().strip('"').strip("'")
        if not got or "://" in got or normalize_url(got):
            continue
        if ":" in got and not re.search(r"%[0-9A-Fa-f]{2}", got):
            got = got.split(":", 1)[1].strip()        # '인증키: 44fc…' 꼴
        low = got.lower()
        if not got or any(low == w or low == w.replace(" ", "") for w in KEY_LABELS):
            continue                                  # '인증키' 같은 이름표만 있는 줄
        if re.search(r"\s", got):
            continue                                  # 'End Point' · '일반 인증키' — 빈칸이 든 줄은 열쇠가 아니다
        if len(got) > len(best):
            best = got
    return best


def api_key(folder=None):
    """서비스 키 — 없으면 None."""
    got = extract_key(os.environ.get("MFDS_API_KEY") or "")
    if got:
        return got
    for root in [folder, os.path.dirname(os.path.abspath(folder))] if folder else []:
        for path in (os.path.join(root or "", KEY_FILE), os.path.join(root or "", "공통", KEY_FILE)):
            try:
                with open(path, encoding="utf-8-sig") as handle:
                    got = extract_key(handle.read())
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


def plain_key(key):
    """공공데이터포털이 주는 두 가지 인증키 어느 쪽을 넣어도 되게 맞춰 준다.

    마이페이지에는 '일반 인증키(Encoding)' 와 '(Decoding)' 이 나란히 있다. 우리는 주소를
    만들 때 다시 인코딩하므로 Decoding 키가 맞는데, 담당자가 어느 것을 복사했는지 알 수
    없다. Encoding 키(`%2B`·`%3D` 가 섞인 것)면 되돌려 쓴다 — 두 번 인코딩되면 인증이
    조용히 실패한다 (담당자 2026-09-14 열쇠 넣기 안내).
    """
    import urllib.parse
    got = (key or "").strip()
    if not got:
        return got
    if re.search(r"%[0-9A-Fa-f]{2}", got):
        back = urllib.parse.unquote(got)
        if back != got:
            return back.strip()
    return got


ERROR_TAGS = ("returnAuthMsg", "errMsg", "resultMsg")

# 포털이 자주 주는 오류를 담당자가 알아볼 말로 바꾼다 — 영어 코드만 보면 무엇을 고칠지 모른다.
ERROR_KOREAN = {
    "SERVICE_KEY_IS_NOT_REGISTERED_ERROR":
        "등록되지 않은 서비스 키입니다 — 마이페이지의 '일반 인증키' 를 다시 복사해 넣어 주세요",
    "SERVICE ERROR": "서비스 쪽 오류입니다 — 잠시 뒤 다시 해 보세요",
    "LIMITED_NUMBER_OF_SERVICE_REQUESTS_EXCEEDS_ERROR":
        "오늘 쓸 수 있는 횟수를 다 썼습니다 — 내일 다시 되거나, 포털에서 한도를 늘려야 합니다",
    "DEADLINE_HAS_EXPIRED_ERROR": "서비스 활용 기간이 끝났습니다 — 포털에서 연장해 주세요",
    "UNREGISTERED_IP_ERROR": "등록되지 않은 IP 입니다 — 포털의 활용 정보에서 IP 를 지우거나 이 PC 것으로 바꿔 주세요",
    "NO_OPENAPI_SERVICE_ERROR": "주소가 맞지 않습니다 — 서비스 주소를 확인해 주세요",
}


def service_error(raw):
    """포털이 돌려준 오류 글 — 없으면 빈 글. 아는 오류는 우리말로 바꾼다."""
    text = raw if isinstance(raw, str) else (raw or b"").decode("utf-8", "replace")
    found = []
    for tag in ERROR_TAGS:
        for m in re.finditer(r"<%s>(.*?)</%s>" % (tag, tag), text, re.S):
            one = " ".join(m.group(1).split())
            if one and one not in found:
                found.append(one)
    if not found:
        return ""
    out = []
    for one in found[:2]:
        korean = ERROR_KOREAN.get(one.upper().replace(" ", "_")) or ERROR_KOREAN.get(one.upper())
        out.append("%s (%s)" % (korean, one) if korean else one)
    return " · ".join(out)


_LICENSE_WORKING = ""       # 이번 실행에서 통한 허가정보 주소


def with_operation(url):
    """서비스 주소(End Point)만 적혀 있으면 조회 이름을 붙인 후보들 — 이미 붙어 있으면 그대로.

    포털의 End Point 는 `…/DrugPrdtPrmsnInfoService07` 까지고 실제 부르는 주소는 그 뒤에
    `/getDrugPrdtPrmsnDtlInq06` 이 붙는다. 판 번호 N 이면 조회 이름은 대개 N-1 (06→Inq05, 07→Inq06).
    확실치 않으니 N-1 · N · 05 를 차례로 둔다.
    """
    url = (url or "").rstrip("/")
    if not url or re.search(r"/get[A-Za-z]+\d*$", url):
        return [url] if url else []
    m = re.search(r"Service(\d+)$", url)
    if not m:
        return [url]
    n = int(m.group(1))
    names = []
    for k in (n - 1, n, 5):
        one = "%s/getDrugPrdtPrmsnDtlInq%02d" % (url, k)
        if one not in names:
            names.append(one)
    return names


def license_bases(folder=None):
    """허가정보 서비스 주소 후보 — 담당자가 넣어 둔 것이 있으면 그것(들)부터, 아니면 기본 후보들."""
    got = (os.environ.get("MFDS_LICENSE_URL") or "").strip()
    if got and normalize_url(got):
        return with_operation(normalize_url(got))
    for root in [folder, os.path.dirname(os.path.abspath(folder))] if folder else []:
        # 주소 파일이 있으면 그것, 없으면 열쇠 파일에 함께 적힌 End Point 를 쓴다
        for name in (LICENSE_URL_FILE, KEY_FILE):
            for path in (os.path.join(root or "", name), os.path.join(root or "", "공통", name)):
                try:
                    with open(path, encoding="utf-8-sig") as handle:
                        got = normalize_url(handle.read())
                except OSError:
                    continue
                if got:
                    mine = with_operation(got)
                    # 주소 파일에 적은 것은 그것만 — 열쇠 파일의 End Point 에서 얻은 것은 기본 후보를 뒤에 둔다
                    rest = [] if name == LICENSE_URL_FILE else [b for b in LICENSE_BASES if b not in mine]
                    if _LICENSE_WORKING and _LICENSE_WORKING in mine + rest:
                        return [_LICENSE_WORKING] + [b for b in mine + rest if b != _LICENSE_WORKING]
                    return mine + rest
    if _LICENSE_WORKING:
        return [_LICENSE_WORKING] + [b for b in LICENSE_BASES if b != _LICENSE_WORKING]
    return list(LICENSE_BASES)


def _http_error_text(error):
    """HTTPError 를 담당자가 읽을 말로 — 포털이 본문에 준 까닭까지."""
    body = ""
    try:
        body = error.read().decode("utf-8", "replace")
    except Exception:
        pass
    why = service_error(body) if body else ""
    return "HTTP %s %s%s" % (getattr(error, "code", "?"), getattr(error, "reason", ""),
                             (" — " + why) if why else "")


def fetch(name, key, base=None, timeout=TIMEOUT, opener=None, folder=None):
    """제품명으로 허가정보를 받아 [{항목: 값}] — 못 받으면 [].

    주소 후보를 차례로 두드린다 — 4xx 면 다음 후보로, 통한 주소는 기억한다. 모두 안 되면
    마지막 오류를 포털 본문의 까닭과 함께 올린다.
    """
    import urllib.error
    import urllib.parse
    import urllib.request
    global _LICENSE_WORKING
    key = plain_key(key)
    if not name or not key:
        return []
    query = urllib.parse.urlencode({
        "serviceKey": key, "item_name": name, "type": "json",
        "pageNo": "1", "numOfRows": str(MAX_ROWS)})
    get = opener or (lambda u, t: urllib.request.urlopen(u, timeout=t).read())
    bases = [base] if base else license_bases(folder)
    raw, last = None, None
    for candidate in bases:
        url = "%s?%s" % (candidate, query)
        try:
            raw = get(url, timeout if len(bases) == 1 else min(timeout, PROBE_TIMEOUT * 2))
            _LICENSE_WORKING = candidate
            break
        except urllib.error.HTTPError as error:
            last = ValueError("%s (주소 %s)" % (_http_error_text(error), candidate.split("?")[0]))
            if error.code in (400, 404, 405):
                continue                     # 주소 문제 — 다음 후보
            raise last
        except Exception as error:
            last = error
            continue
    if raw is None:
        raise last or ValueError("허가정보 주소에 닿지 못했습니다")
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", "replace")
    try:
        payload = json.loads(raw)
    except ValueError:
        # 인증이 안 되면 포털은 HTTP 200 에 XML 오류를 준다. 그 글을 그대로 올려야
        # 담당자가 '열쇠가 틀렸구나' 를 안다 — 안 그러면 까닭 없이 넘어간다.
        raise ValueError(service_error(raw) or "응답을 읽지 못했습니다 (JSON 이 아닙니다)")
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


# ── 의약품 행정처분 정보 ────────────────────────────────────────────────
# 담당자 2026-09-14: "행정처분 정보 API 도 함께 승인되어 있습니다 … 같이 넣어줘".
# 3항 6번 '시판 후 준수사항 이행하여 행정 처분 이력 없음을 확인함' 을 사람 대신 확인한다.
#
# 이 서비스는 판마다 주소가 달라서 하나로 못 박는다. 담당자가 포털 '요청 주소' 를 그대로
# 복사해 `공통/식약처-행정처분-주소.txt` 에 넣으면 그것을 쓰고, 없으면 아래 후보를 차례로
# 두드린다. 어느 것도 안 되면 조용히 넘어가고 작성 기록에 주소를 넣어 달라고 적는다.
PENALTY_URL_FILE = "식약처-행정처분-주소.txt"
# 주소를 모를 때 '자동' 이라고 적어 두면 아래 후보를 차례로 두드린다. 그냥 두면(파일이 없으면)
# 건너뛴다 — 매 실행마다 아닌 주소를 기다리느라 시간을 버리지 않기 위해서다
# (담당자 2026-09-14: "너무 오래 걸리면 생략할까 고민하고 있어").
AUTO_WORDS = ("자동", "auto", "찾기")
PENALTY_BASES = (
    "http://apis.data.go.kr/1471000/MdcinPrmisnAdmDsposInfoService/getMdcinPrmisnAdmDsposInq",
    "http://apis.data.go.kr/1471000/AdmDsposInfoService/getAdmDsposInq",
    "http://apis.data.go.kr/1471000/DrugAdmDsposInfoService/getDrugAdmDsposInq",
)
# 담당자 PC 2026-09-16: 승인된 서비스는 '식품의약품안전처_의약품 행정처분 정보' —
# End Point https://apis.data.go.kr/1471000/MdcinExaathrService04 (참고문서 IROS_50). 그 응답 항목
# (ADM_DISPS_NAME · LAST_SETTLE_DATE · EXPOSE_CONT · BEF_APPLY_LAW · DISPS_TERM_DATE)을 앞에 둔다.
PENALTY_FIELDS = {
    "업체명": ("ENTP_NAME", "entpName", "BSSH_NM", "bsshNm", "COMPANY_NAME"),
    "제품명": ("ITEM_NAME", "PRDUCT", "itemName", "PRDLST_NM", "prdlstNm", "PRODUCT_NAME"),
    "처분일자": ("LAST_SETTLE_DATE", "DISPOS_DATE", "disposDate", "ADM_DISPOS_DE", "DSPS_DT", "PROCESS_DATE"),
    "처분내용": ("ADM_DISPS_NAME", "DISPOS_CONT", "disposCont", "ADM_DISPOS_CN", "DSPS_CN", "PROCESS_CONTENT"),
    "처분기간": ("DISPS_TERM_DATE", "DISPOS_PERIOD", "disposPeriod", "DSPS_PD"),
    "근거법령": ("BEF_APPLY_LAW", "VIOLATION_LAW", "violationLaw", "LAW_NM", "BASIS_LAW"),
    "위반내용": ("EXPOSE_CONT", "VIOLATION_CONT", "violationCont", "VILT_CN"),
}


_WORKING_BASE = ""          # 이번 실행에서 통한 행정처분 주소


def remember_base(base):
    """통한 주소를 기억해 다음 제품부터 바로 쓴다."""
    global _WORKING_BASE
    if base:
        _WORKING_BASE = base


def forget_base():
    """시험에서 기억을 지운다."""
    global _WORKING_BASE
    _WORKING_BASE = ""


def penalty_with_operation(url):
    """행정처분 End Point 만 적혀 있으면(조회 이름 /get… 이 없으면) 아는 조회 이름들을 붙인 후보 —
    담당자 PC 2026-09-16: 열쇠 파일처럼 포털 화면의 End Point 만 붙여 넣는 것이 자연스럽다.
    이미 조회 이름이 붙어 있으면 그것만."""
    url = (url or "").rstrip("/")
    if not url:
        return []
    if re.search(r"/get[A-Za-z]+\d*$", url) or "data.go.kr" not in url.lower():
        return [url]                               # 조회 이름이 있거나 포털 밖 주소면 그대로
    ops = []
    # 'MdcinExaathrService04' 처럼 판 번호가 붙은 서비스는 조회 이름도 번호가 붙는다 —
    # 허가정보처럼 N-1 이 흔하고(06→05), 번호 없는 옛 꼴도 있어 차례로 둔다.
    m = re.search(r"/([A-Za-z]+?)(?:Info)?Service(\d+)$", url)
    if m:
        stem, n = m.group(1), int(m.group(2))
        for k in (n - 1, n, None, n - 2, n + 1):
            if k is not None and k < 1:
                continue
            ops.append("get%sList%s" % (stem, ("%02d" % k) if k is not None else ""))
    for base in PENALTY_BASES:
        op = base.rsplit("/", 1)[-1]
        if op not in ops:
            ops.append(op)
    return ["%s/%s" % (url, op) for op in ops]


def penalty_base(folder=None):
    """행정처분 서비스 주소 — 담당자가 넣어 둔 것이 있으면 그것을 먼저 쓴다."""
    got = (os.environ.get("MFDS_PENALTY_URL") or "").strip()
    if got:
        return [got] if _looks_like_url(got) else []
    for root in [folder, os.path.dirname(os.path.abspath(folder))] if folder else []:
        for path in (os.path.join(root or "", PENALTY_URL_FILE),
                     os.path.join(root or "", "공통", PENALTY_URL_FILE)):
            try:
                with open(path, encoding="utf-8-sig") as handle:
                    raw_text = handle.read()
                    got = raw_text.strip().split("?")[0].strip()
            except OSError:
                continue
            url = normalize_url(raw_text)
            if url:
                return penalty_with_operation(url)
            if got.strip() in AUTO_WORDS:          # '자동' 이라고 적으면 후보를 두드린다
                return list(PENALTY_BASES)
            if got:                                # 주소 자리에 열쇠를 넣은 것 — 못 쓴다
                return []
    return []                                      # 주소 파일이 없으면 건너뛴다


def _penalty_value(item, name):
    for key in PENALTY_FIELDS.get(name, ()):
        value = item.get(key)
        if value not in (None, ""):
            return " ".join(str(value).split())
    return ""


def fetch_penalties(name, key, bases=None, timeout=TIMEOUT, opener=None):
    """제품명으로 행정처분 이력을 받아 [{항목: 값}].

    주소 후보를 차례로 두드려 **처음으로 제대로 답한 것**을 쓴다. 못 받으면 마지막 까닭을
    올린다 — 조용히 '이력 없음' 으로 넘어가면 안 되는 항목이다.
    """
    import urllib.parse
    import urllib.request
    key = plain_key(key)
    if not name or not key:
        return [], ""
    get = opener or (lambda u, t: urllib.request.urlopen(u, timeout=t).read())
    last = ""
    candidates = list(bases or PENALTY_BASES)
    # 한 번 통한 주소를 기억한다 — 제품마다 후보를 다시 두드리면 그만큼 느려진다.
    if _WORKING_BASE and _WORKING_BASE in candidates:
        candidates = [_WORKING_BASE] + [b for b in candidates if b != _WORKING_BASE]
    probing = len(candidates) > 1
    for base in candidates:
        query = urllib.parse.urlencode({
            "serviceKey": key, "type": "json", "pageNo": "1",
            "numOfRows": str(MAX_ROWS), "PRDUCT": name, "item_name": name})
        wait = PROBE_TIMEOUT if (probing and base != _WORKING_BASE) else timeout
        try:
            raw = get("%s?%s" % (base, query), wait)
        except Exception as error:                     # 주소가 아예 없으면 다음 후보로
            last = str(error)
            continue
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8", "replace")
        try:
            payload = json.loads(raw)
        except ValueError:
            last = service_error(raw) or "응답을 읽지 못했습니다 (JSON 이 아닙니다)"
            continue
        out = []
        for item in _rows(payload):
            if not isinstance(item, dict):
                continue
            one = {k: _penalty_value(item, k) for k in PENALTY_FIELDS}
            if any(one.values()):
                out.append(one)
        remember_base(base)
        return out, base
    return [], last


def penalties_for(rows, name, entp=""):
    """우리 제품(또는 우리 회사)에 걸린 것만 추린다."""
    want, company = _flat(name), _flat(entp)
    out = []
    for one in rows:
        product = _flat(one.get("제품명"))
        maker = _flat(one.get("업체명"))
        if product and want and (want in product or product in want):
            out.append(one)
        elif company and maker and not product and company in maker:
            out.append(one)
    return out
