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


def read_pdf_boxes(path):
    """PDF 의 글자 상자를 (x0, y1, 글자) 목록으로 돌려줍니다 (페이지별).

    표를 좌표로 복원하기 위한 것입니다. 셀 안에서 줄이 바뀌는 서식은
    글자 순서만으로는 어느 칸의 내용인지 알 수 없습니다.
    """
    try:
        from pdfminer.high_level import extract_pages
        from pdfminer.layout import LTTextContainer
    except ImportError:
        raise CoaError(
            "PDF 성적서를 읽으려면 pdfminer.six 가 필요합니다. "
            "`pip install pdfminer.six` 후 다시 실행하세요."
        )
    try:
        pages = []
        for page in extract_pages(path):
            boxes = []
            for element in page:
                if not isinstance(element, LTTextContainer):
                    continue
                for line in element:
                    text = " ".join(line.get_text().split())
                    if text:
                        boxes.append({"x0": line.x0, "y1": line.y1, "text": text})
            boxes.sort(key=lambda box: (-box["y1"], box["x0"]))
            pages.append(boxes)
        return pages
    except CoaError:
        raise
    except Exception as error:
        raise CoaError("PDF 를 읽지 못했습니다: %s" % error)


def cluster_columns(boxes, gap=25.0):
    """왼쪽 좌표를 모아 칸(열) 경계를 찾습니다."""
    positions = sorted({round(box["x0"], 1) for box in boxes})
    if not positions:
        return []
    groups = [[positions[0]]]
    for value in positions[1:]:
        if value - groups[-1][-1] > gap:
            groups.append([value])
        else:
            groups[-1].append(value)
    return [(min(group), max(group)) for group in groups]


def _column_of(box, columns):
    for index, (low, high) in enumerate(columns):
        if low - 1 <= box["x0"] <= high + 1:
            return index
    # 어느 칸에도 딱 맞지 않으면 가장 가까운 칸으로 봅니다.
    return min(range(len(columns)),
               key=lambda i: abs(box["x0"] - columns[i][0])) if columns else 0


def parse_layout(boxes):
    """좌표가 있는 글자 상자에서 머리말과 시험결과 표를 읽어 냅니다."""
    squashed = [_squash(box["text"]) for box in boxes]

    header = {}
    for index, key in enumerate(squashed):
        field = HEADER_LABELS.get(key)
        if field and field not in header:
            label = boxes[index]
            same_row = [b for b, k in zip(boxes, squashed)
                        if k not in HEADER_LABELS
                        and abs(b["y1"] - label["y1"]) <= 3
                        and b["x0"] > label["x0"]]
            if same_row:
                # 같은 줄에서 라벨 바로 오른쪽에 있는 값이 그 라벨의 값입니다.
                header[field] = min(same_row, key=lambda b: b["x0"] - label["x0"])["text"]
    for field in ("mfg_date", "sampled_date", "tested_date"):
        if header.get(field):
            header[field] = re.sub(r"\s+", "", header[field]).replace("/", "-")

    start = next((i for i, key in enumerate(squashed)
                  if key == _squash(TABLE_HEADER[0])), None)
    if start is None:
        raise CoaError("시험결과 표를 찾지 못했습니다.")
    header_y = boxes[start]["y1"]

    end_y = 0.0
    for box, key in zip(boxes[start + 1:], squashed[start + 1:]):
        if box["y1"] < header_y - 5 and any(key.startswith(_squash(end)) for end in TABLE_END):
            end_y = box["y1"]
            break

    body = [box for box in boxes if end_y < box["y1"] < header_y - 5]
    if not body:
        raise CoaError("시험결과 표가 비어 있습니다.")

    columns = cluster_columns(body)
    if len(columns) < 3:
        raise CoaError("시험결과 표의 칸을 구분하지 못했습니다.")

    placed = [(box, _column_of(box, columns)) for box in body]

    def share(index, predicate):
        cells = [box["text"] for box, column in placed if column == index]
        return sum(1 for text in cells if predicate(text)) / len(cells) if cells else 0.0

    verdict_col = max(range(len(columns)), key=lambda i: share(i, lambda t: t in VERDICTS))
    tester_col = max(range(len(columns)),
                     key=lambda i: share(i, lambda t: bool(_TESTER_DATE.match(t))))
    if share(verdict_col, lambda t: t in VERDICTS) < 0.5:
        verdict_col = None
    if share(tester_col, lambda t: bool(_TESTER_DATE.match(t))) < 0.5:
        tester_col = None

    content = [i for i in range(len(columns)) if i not in (verdict_col, tester_col)]
    if len(content) < 2:
        raise CoaError("시험항목·기준 칸을 찾지 못했습니다.")
    name_col, criteria_col = content[0], content[1]
    result_col = content[2] if len(content) > 2 else None

    anchors = sorted((box for box, column in placed if column == name_col),
                     key=lambda box: -box["y1"])
    merged = []
    for box in anchors:                       # 항목 이름이 두 줄인 경우 하나로 봅니다.
        if merged and merged[-1]["y1"] - box["y1"] < 12:
            merged[-1] = {"x0": merged[-1]["x0"], "y1": merged[-1]["y1"],
                          "text": merged[-1]["text"] + " " + box["text"]}
        else:
            merged.append(dict(box))
    if not merged:
        raise CoaError("시험항목을 찾지 못했습니다.")

    tests = []
    for index, anchor in enumerate(merged):
        top = anchor["y1"] + 3
        bottom = merged[index + 1]["y1"] + 3 if index + 1 < len(merged) else end_y
        rows = [(box, column) for box, column in placed if bottom < box["y1"] <= top]

        def gather(column):
            picked = [box for box, col in rows if col == column and box is not anchor]
            picked.sort(key=lambda box: (-box["y1"], box["x0"]))
            return " ".join(box["text"] for box in picked).strip()

        verdict = ""
        if verdict_col is not None:
            for box, col in sorted(rows, key=lambda pair: -pair[0]["y1"]):
                if col == verdict_col and box["text"] in VERDICTS:
                    verdict = box["text"]
                    break
        tests.append({
            "test_name": anchor["text"],
            "criteria": gather(criteria_col),
            "result": gather(result_col) if result_col is not None else "",
            "verdict": verdict,
        })
    return {"header": header, "tests": tests}


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


# '7.00' · '0.52mL' · '304 mOsm/kg' · '0CFU/100mL' 처럼 숫자 하나로 이루어진 측정값.
# 단위 안의 숫자(100mL)는 허용하되, 한글이 섞이면('1회용') 측정값으로 보지 않습니다.
_MEASURED = re.compile(
    r"^-?\d+(?:\.\d+)?\s*(?:[A-Za-zμµ%℃·][A-Za-z0-9μµ%℃·/^\-]*)?\.?$")


def parse_result(text, numeric=True):
    """'7.00' · 'Av. 304mOsm/kg (302 ~ 305 mOsm/kg)' 에서 대표값을 뽑습니다.

    평균(Av.)이 적혀 있으면 그 값을 씁니다. 성상·이물검사처럼 문장으로 적힌
    결과에서 숫자를 억지로 뽑으면 안 되므로('1회용 용기' → 1), 측정값 모양일
    때만 숫자로 봅니다. numeric=False 면 아예 숫자를 뽑지 않습니다.
    """
    if not text or not numeric:
        return None
    average = re.search(r"(?:Av\.?|평균)\s*:?\s*([\d.]+)", text, re.IGNORECASE)
    if average:
        return float(average.group(1))
    # 괄호 안의 범위는 참고값이므로 빼고 봅니다.
    outside = re.sub(r"\([^)]*\)", " ", text).strip()
    first = outside.split()[0] if outside.split() else ""
    if _MEASURED.match(outside) or _MEASURED.match(first):
        match = _NUMBER.search(outside)
        return float(match.group(0)) if match else None
    return None


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
    """PDF 성적서 한 건을 읽어 행 목록으로.

    좌표로 표를 복원하는 방식을 먼저 쓰고, 그게 안 되면 글자 순서로 읽습니다.
    """
    pages = read_pdf_boxes(path)
    if pages and pages[0]:
        try:
            return to_rows(parse_layout(pages[0]), product_code)
        except CoaError:
            pass
    lines = [box["text"] for page in pages for box in page]
    if not lines:
        raise CoaError(
            "이 PDF 에는 글자 정보가 없습니다 (스캔한 이미지로 보입니다). "
            "원본 파일로 다시 내려받거나, 문자 인식(OCR)을 거쳐야 합니다."
        )
    return to_rows(parse_lines(lines), product_code)
