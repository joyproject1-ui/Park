"""공정 시험성적서(PDF)에서 시험결과를 뽑아냅니다.

한림제약 서식 `HLF-QC-121-04` 처럼 라벨과 값이 번갈아 나오는 성적서를 대상으로,
제조번호 · 공정 · 시험항목별 기준/결과/판정을 표 형태로 만들어 줍니다.
결과 행은 `batches` 데이터셋과 같은 모양이라 그대로 집계에 넣을 수 있습니다.

PDF 읽기에는 `pdfminer.six` 가 필요합니다 (`pip install pdfminer.six`).
설치하지 않아도 나머지 기능은 그대로 동작합니다.
"""

import re

# 성적서 머리말의 라벨 — 공백을 없앤 형태로 비교합니다.
HEADER_LABELS = {
    "제품명": "product_name",
    "제조번호": "batch_no",
    "공정명": "stage",
    "제조일자": "mfg_date",
    "검체채취일": "sampled_date",
    "시험성적서번호": "coa_no",
    "LOT_수량": "lot_size",
    "시험방법기준": "test_basis",
    "판정결과": "lot_verdict",
    "시험완료일자": "tested_date",
}

# 시험결과 표의 머리글 (이 줄들이 나오면 그 뒤부터 항목이 시작됩니다)
TABLE_HEADER = ("시험항목", "시험기준", "시험결과")
# 표의 끝을 알리는 라벨
TABLE_END = ("비고", "시험완료일자", "확인일자", "판정일자", "판정결과")

VERDICTS = ("적합", "부적합", "합격", "불합격")
_RANGE = re.compile(r"(-?\d+(?:\.\d+)?)\s*[~∼-]\s*(-?\d+(?:\.\d+)?)")
_NUMBER = re.compile(r"-?\d+(?:\.\d+)?")
_TESTER_DATE = re.compile(r"^\S+\s+\d{4}[./-]\d{1,2}[./-]\d{1,2}$")


class CoaError(Exception):
    """성적서를 읽지 못했을 때."""


def _squash(text):
    return re.sub(r"\s+", "", text or "")


def read_pdf_lines(path):
    """PDF 에서 빈 줄을 뺀 텍스트 줄을 순서대로 돌려줍니다."""
    try:
        from pdfminer.high_level import extract_text
    except ImportError:
        raise CoaError(
            "PDF 성적서를 읽으려면 pdfminer.six 가 필요합니다. "
            "`pip install pdfminer.six` 후 다시 실행하세요."
        )
    try:
        text = extract_text(path)
    except Exception as error:                     # 손상된 PDF 등
        raise CoaError("PDF 를 읽지 못했습니다: %s" % error)
    return [line.strip() for line in text.split("\n") if line.strip()]


def parse_criteria(text):
    """'6.8 ~ 7.2' · '290 ~ 330 mOsm/kg' · '참고치' 를 (하한, 상한, 단위) 로."""
    if not text:
        return None, None, ""
    match = _RANGE.search(text)
    if not match:
        return None, None, ""
    low, high = float(match.group(1)), float(match.group(2))
    unit = text[match.end():].strip(" ()")
    return low, high, unit


def parse_result(text):
    """'7.00' · 'Av. 304mOsm/kg (302 ~ 305 mOsm/kg)' 에서 대표값을 뽑습니다.

    평균(Av.)이 적혀 있으면 그 값을, 아니면 처음 나오는 숫자를 씁니다.
    숫자가 없으면(성상 등) None 을 돌려줍니다.
    """
    if not text:
        return None
    average = re.search(r"(?:Av\.?|평균)\s*([\d.]+)", text, re.IGNORECASE)
    if average:
        return float(average.group(1))
    # 괄호 안의 범위는 참고값이므로 빼고 봅니다.
    outside = re.sub(r"\([^)]*\)", " ", text)
    match = _NUMBER.search(outside)
    return float(match.group(0)) if match else None


def _is_label(line, labels):
    return _squash(line) in labels


def parse_lines(lines):
    """성적서 한 장의 줄 목록을 dict 로 바꿉니다."""
    header = {}
    squashed = [_squash(line) for line in lines]

    for index, key in enumerate(squashed):
        field = HEADER_LABELS.get(key)
        if field and field not in header:
            for candidate in lines[index + 1:index + 2]:
                if _squash(candidate) not in HEADER_LABELS:
                    header[field] = candidate.strip()
    for field in ("mfg_date", "sampled_date", "tested_date"):
        if header.get(field):
            header[field] = re.sub(r"\s+", "", header[field]).replace("/", "-")

    start = None
    for index, key in enumerate(squashed):
        if key == _squash(TABLE_HEADER[0]):
            start = index + len(TABLE_HEADER) + 1     # 머리글 4줄(시험자/일자/판정 포함)
            break
    if start is None:
        raise CoaError("시험결과 표를 찾지 못했습니다.")

    tests = []
    current = None
    for line in lines[start:]:
        key = _squash(line)
        if any(key.startswith(_squash(end)) for end in TABLE_END):
            break
        if current is None:
            # 시험자·일자나 판정만 남은 줄은 항목 이름이 될 수 없습니다.
            # (변경관리로 시험이 생략되어 줄이 어긋난 경우를 걸러 냅니다.)
            if _TESTER_DATE.match(line.strip()) or line.strip() in VERDICTS:
                continue
            current = {"test_name": line.strip(), "criteria": [], "result": [], "verdict": ""}
            continue
        if line.strip() in VERDICTS:
            current["verdict"] = line.strip()
            tests.append(current)
            current = None
            continue
        if _TESTER_DATE.match(line.strip()):
            continue                                  # 시험자·시험일자 줄은 쓰지 않습니다
        if not current["criteria"]:
            current["criteria"].append(line.strip())
        else:
            current["result"].append(line.strip())

    for test in tests:
        test["criteria"] = " ".join(test["criteria"]).strip()
        test["result"] = " ".join(test["result"]).strip()
    return {"header": header, "tests": tests}


def to_rows(parsed, product_code=""):
    """`batches` 데이터셋과 같은 모양의 행 목록으로 바꿉니다."""
    header = parsed["header"]
    rows = []
    for test in parsed["tests"]:
        lsl, usl, unit = parse_criteria(test["criteria"])
        rows.append({
            "product_code": product_code,
            "product_name": header.get("product_name", ""),
            "batch_no": header.get("batch_no", ""),
            "mfg_date": header.get("mfg_date", ""),
            "stage": header.get("stage", ""),
            "test_name": test["test_name"],
            "value": parse_result(test["result"]),
            "unit": unit,
            "lsl": lsl,
            "usl": usl,
            "verdict": test["verdict"],
            "criteria_text": test["criteria"],
            "result_text": test["result"],
            "coa_no": header.get("coa_no", ""),
        })
    return rows


def read_coa(path, product_code=""):
    """PDF 성적서 한 건을 읽어 행 목록으로."""
    return to_rows(parse_lines(read_pdf_lines(path)), product_code)
