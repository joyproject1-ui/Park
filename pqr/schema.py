"""입력 파일의 열 이름을 표준 필드로 맞추고 값을 정리합니다.

사내 시스템(LIMS · MES · ERP)마다 열 이름이 달라서, 한국어 · 영어 별칭을 모두
받아들이고 표준 필드명으로 바꿉니다. 별칭은 `pqr/data/aliases.json` 에서 확장할 수 있습니다.
"""

import datetime as _dt
import json
import os
import re

from .tabular import excel_serial_to_date

_HERE = os.path.dirname(os.path.abspath(__file__))

# 데이터셋별 표준 필드: (필드명, 필수 여부)
DATASETS = {
    "products": {
        "label": "제품 마스터",
        "fields": ["product_code", "product_name", "form", "site", "owner",
                   "period_from", "period_to", "due", "stage", "group", "lots",
                   "product_class", "license_no", "license_date", "shelf_life", "storage"],
        "required": ["product_code"],
    },
    "batches": {
        "label": "배치 시험성적서 · 공정관리",
        "fields": ["product_code", "product_name", "batch_no", "mfg_date", "stage",
                   "test_name", "value", "unit", "lsl", "usl", "verdict"],
        "required": ["product_code", "batch_no", "test_name"],
    },
    "deviations": {
        "label": "일탈 · OOS/OOT · CAPA",
        "fields": ["product_code", "record_no", "type", "severity", "opened_date",
                   "title", "status", "closed_date", "capa_no", "capa_status"],
        "required": ["product_code", "type"],
    },
    "changes": {
        "label": "변경관리 · 불만 · 회수",
        "fields": ["product_code", "record_no", "type", "opened_date", "title",
                   "status", "closed_date"],
        "required": ["product_code", "type"],
    },
    "stability": {
        "label": "안정성 모니터링",
        "fields": ["product_code", "batch_no", "condition", "timepoint",
                   "test_name", "value", "unit", "lsl", "usl"],
        "required": ["product_code", "test_name", "timepoint"],
    },
    "stagelog": {
        "label": "단계 진행 이력 (선택)",
        "fields": ["product_code", "stage", "entered_date"],
        "required": ["product_code", "stage", "entered_date"],
    },
    "qualification": {
        "label": "설비 적격성",
        "fields": ["asset", "type", "product_code", "last_qualified", "next_due", "status"],
        "required": ["asset"],
    },
}

DATE_FIELDS = frozenset([
    "mfg_date", "opened_date", "closed_date", "last_qualified", "next_due",
    "due", "period_from", "period_to", "entered_date",
])
NUMBER_FIELDS = frozenset(["value", "lsl", "usl", "timepoint", "lots"])

# 표준 필드 -> 받아들이는 열 이름들 (소문자·공백제거 후 비교)
_BUILTIN_ALIASES = {
    "product_code": ["제품코드", "품목코드", "제품번호", "코드", "productcode", "itemcode", "code"],
    "product_name": ["제품명", "품목명", "productname", "itemname", "name"],
    "form": ["제형", "dosageform", "form"],
    # 연간 계획서가 정해 주는 값 — 평가 그룹과 그 해 생산 Lot 수
    "group": ["그룹", "평가그룹", "pqr그룹", "group"],
    "lots": ["생산수량", "생산lot", "lot수", "생산수량lot", "lots"],
    # 보고서 3항(대상 제품)이 요구하는 값 — 허가증에서 옮겨 적습니다.
    "product_class": ["제품분류", "분류", "productclass"],
    "license_no": ["허가번호", "품목허가번호", "licenseno"],
    "license_date": ["허가일자", "허가일", "licensedate"],
    "shelf_life": ["사용기한", "유효기간", "shelflife"],
    "storage": ["보관조건", "저장방법", "storage"],
    "site": ["공장", "제조소", "사이트", "라인", "site", "plant", "line"],
    "owner": ["담당자", "담당", "작성자", "owner", "assignee"],
    "period_from": ["평가시작일", "평가기간시작", "기간시작", "periodfrom", "from"],
    "period_to": ["평가종료일", "평가기간종료", "기간종료", "periodto", "to"],
    "due": ["마감일", "제출기한", "기한", "duedate", "due"],
    "stage": ["단계", "진행단계", "공정구분", "구분단계", "stage", "step"],
    "batch_no": ["배치번호", "제조번호", "로트", "로트번호", "batchno", "lot", "lotno", "batch"],
    "mfg_date": ["제조일", "제조일자", "생산일", "mfgdate", "manufacturedate"],
    "test_name": ["시험항목", "항목", "관리항목", "testname", "test", "parameter"],
    "value": ["결과값", "측정값", "결과", "값", "value", "result"],
    "unit": ["단위", "unit"],
    "lsl": ["규격하한", "하한", "하한규격", "lsl", "lowerlimit", "min"],
    "usl": ["규격상한", "상한", "상한규격", "usl", "upperlimit", "max"],
    "verdict": ["판정", "적합여부", "verdict", "judgement", "judgment"],
    "record_no": ["번호", "관리번호", "문서번호", "recordno", "no", "id"],
    "type": ["구분", "유형", "종류", "type", "category"],
    "severity": ["등급", "중요도", "위험도", "severity", "grade"],
    "opened_date": ["발생일", "등록일", "접수일", "발생일자", "openeddate", "date"],
    "title": ["제목", "내용", "사유", "title", "subject", "description"],
    "status": ["상태", "진행상태", "처리상태", "status"],
    "closed_date": ["종결일", "완료일", "조치완료일", "closeddate", "closed"],
    "capa_no": ["capa번호", "capano", "capa"],
    "capa_status": ["capa상태", "capa진행", "capastatus"],
    "condition": ["조건", "보관조건", "시험조건", "condition"],
    "timepoint": ["시점", "개월", "시험시점", "timepoint", "month", "months"],
    "asset": ["설비", "설비명", "시스템", "시스템명", "asset", "equipment", "system"],
    "entered_date": ["진입일", "단계진입일", "시작일", "entereddate", "startdate"],
}


def _key(text):
    return re.sub(r"[\s_\-()·.]", "", str(text or "")).lower()


def load_aliases():
    """내장 별칭에 `pqr/data/aliases.json` 의 사용자 정의를 더합니다."""
    aliases = {field: list(names) for field, names in _BUILTIN_ALIASES.items()}
    path = os.path.join(_HERE, "data", "aliases.json")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as handle:
            extra = json.load(handle)
        for field, names in extra.items():
            aliases.setdefault(field, [])
            aliases[field].extend(names)
    lookup = {}
    for field, names in aliases.items():
        lookup[_key(field)] = field
        for name in names:
            lookup[_key(name)] = field
    return lookup


def alias_lookup(dataset):
    """그 자료 종류에 있는 필드만으로 별칭을 풉니다.

    같은 낱말을 두 필드가 쓰기도 합니다 — '보관조건' 은 제품 마스터에서는 storage,
    안정성 자료에서는 condition 입니다. 전체 별칭표 하나로 풀면 나중에 등록된 쪽이
    이겨서, 다른 자료에서는 그 열이 통째로 버려집니다.
    """
    fields = set(DATASETS[dataset]["fields"])
    aliases = {field: list(names) for field, names in _BUILTIN_ALIASES.items()}
    path = os.path.join(_HERE, "data", "aliases.json")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as handle:
            for field, names in json.load(handle).items():
                aliases.setdefault(field, []).extend(names)
    lookup = {}
    for field, names in aliases.items():
        if field not in fields:
            continue
        lookup[_key(field)] = field
        for name in names:
            lookup.setdefault(_key(name), field)
    return lookup


def parse_date(text):
    """ISO · 한국식 · 엑셀 일련번호를 date 로 바꿉니다. 실패하면 None."""
    if text is None:
        return None
    if isinstance(text, _dt.datetime):
        return text.date()
    if isinstance(text, _dt.date):
        return text
    raw = str(text).strip()
    if not raw:
        return None
    raw = raw.split("T")[0]
    raw = re.sub(r"\s+\d{1,2}:\d{2}(:\d{2})?.*$", "", raw).strip()
    normalized = raw.replace("/", "-").replace(".", "-")
    normalized = re.sub(r"년|월", "-", normalized).replace("일", "")
    parts = [part.strip() for part in normalized.split("-") if part.strip() != ""]
    if len(parts) == 3:
        try:
            year, month, day = (int(part) for part in parts)
            if year < 100:
                year += 2000
            return _dt.date(year, month, day)
        except ValueError:
            pass
    if len(parts) == 1 and len(parts[0]) == 8 and parts[0].isdigit():
        try:
            return _dt.date(int(parts[0][:4]), int(parts[0][4:6]), int(parts[0][6:]))
        except ValueError:
            pass
    return excel_serial_to_date(raw)


def parse_number(text):
    """숫자를 float 로 바꿉니다. 쉼표·단위·비교기호는 무시합니다."""
    if text is None or text == "":
        return None
    if isinstance(text, (int, float)):
        return float(text)
    cleaned = str(text).replace(",", "").strip()
    match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    if not match:
        return None
    try:
        return float(match.group(0))
    except ValueError:
        return None


def normalize(rows, dataset, source="", default_product_code=None):
    """원본 행을 표준 필드로 바꾸고, 필수값이 빠진 행은 문제로 보고합니다.

    default_product_code 를 주면(제품 폴더에서 읽은 경우) 제품코드 열이 없거나 비어 있는
    행에 그 값을 채웁니다. 행에 적힌 제품코드가 폴더와 다르면 경고로 남깁니다.

    반환값: (정규화된 행 목록, 문제 목록)
    """
    spec = DATASETS[dataset]
    lookup = alias_lookup(dataset)
    known = set(spec["fields"])
    issues = []
    mapped_columns = {}
    normalized_rows = []

    for line_no, row in enumerate(rows, start=2):
        record = {}
        for column, value in row.items():
            field = lookup.get(_key(column))
            if field is None or field not in known:
                continue
            mapped_columns[column] = field
            if field in DATE_FIELDS:
                parsed = parse_date(value)
                if value not in ("", None) and parsed is None:
                    issues.append({
                        "source": source, "row": line_no, "field": field,
                        "level": "warning",
                        "message": "날짜를 해석하지 못했습니다: %r" % value,
                    })
                record[field] = parsed.isoformat() if parsed else None
            elif field in NUMBER_FIELDS:
                record[field] = parse_number(value)
            else:
                record[field] = str(value).strip()
        if default_product_code and "product_code" in known:
            written = (record.get("product_code") or "").strip().upper()
            folder_code = default_product_code.strip().upper()
            if not written:
                record["product_code"] = folder_code
            elif written != folder_code:
                issues.append({
                    "source": source, "row": line_no, "field": "product_code",
                    "level": "warning",
                    "message": "폴더(%s)와 파일에 적힌 제품코드(%s)가 다릅니다. 파일 값을 씁니다."
                               % (folder_code, written),
                })

        # 표 아래에 적어 둔 안내·비고 문장이 한 건의 자료로 잡히지 않게 걸러 냅니다.
        filled = {key: value for key, value in record.items() if value not in (None, "")}
        if len(filled) == 1:
            only = list(filled.values())[0]
            if isinstance(only, str) and len(only) > 40:
                issues.append({
                    "source": source, "row": line_no, "field": list(filled)[0],
                    "level": "info",
                    "message": "설명문으로 보여 건너뛴 행입니다",
                })
                continue

        # 숫자 0(예: 안정성 0개월 시점)은 값이 있는 것이므로 '비어 있음'으로 보면 안 됩니다.
        missing = [field for field in spec["required"]
                   if record.get(field) is None or record.get(field) == ""]
        if missing:
            issues.append({
                "source": source, "row": line_no, "field": ",".join(missing),
                "level": "error",
                "message": "필수 항목이 비어 있어 행을 건너뜁니다",
            })
            continue
        if record.get("product_code"):
            record["product_code"] = record["product_code"].upper()
        normalized_rows.append(record)

    if rows:
        unmapped = [column for column in rows[0] if column and column not in mapped_columns]
        if unmapped:
            issues.append({
                "source": source, "row": 1, "field": ",".join(unmapped),
                "level": "info",
                "message": "표준 필드로 연결되지 않아 무시한 열입니다",
            })
    return normalized_rows, issues


def detect_dataset(filename):
    """파일 이름으로 데이터셋 종류를 추정합니다."""
    stem = _key(os.path.splitext(os.path.basename(filename))[0])
    table = [
        ("products", ["products", "product", "제품마스터", "제품목록", "품목"]),
        ("batches", ["batches", "batch", "배치", "시험성적서", "성적서", "공정관리", "coa"]),
        ("deviations", ["deviations", "deviation", "일탈", "oos", "oot", "capa"]),
        ("changes", ["changes", "change", "변경", "불만", "회수", "반품", "complaint", "recall"]),
        ("stability", ["stability", "안정성", "장기보존", "가속"]),
        ("stagelog", ["stagelog", "단계이력", "진행이력", "단계로그"]),
        ("qualification", ["qualification", "적격성", "밸리데이션", "hvac", "용수", "설비"]),
    ]
    for dataset, keys in table:
        for key in keys:
            if _key(key) in stem:
                return dataset
    return None
