"""입력 파일을 읽어 PQR 데이터셋(집계 · 판정 결과)을 만듭니다.

`build()` 한 번이면 파일 적재 → 정규화 → 계산 → 판정까지 끝나고,
결과 dict 하나에 대시보드용 값과 보고서용 상세 값이 모두 들어갑니다.
계산에 쓴 파일의 이름 · 크기 · SHA-256 은 `sources` 에 남습니다 (감사추적용).
"""

import datetime as _dt
import hashlib
import json
import os
import re

from . import metrics, schema
from .tabular import TableError, read_table

_HERE = os.path.dirname(os.path.abspath(__file__))


def load_config(path=None):
    with open(path or os.path.join(_HERE, "data", "config.json"), encoding="utf-8") as handle:
        return json.load(handle)


def _sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


# 제품 공통 자료(설비 적격성, 위수탁 협약 등)를 두는 폴더 이름
COMMON_FOLDERS = ("공통", "_공통", "common", "shared", "전사", "site")

# '이력 없음으로 마감' 단추가 남기는 확인 기록 파일의 표식.
# 이름에 '일탈'·'불만' 같은 낱말이 들어가므로, 표 자료(대장)로 오인해 읽으면
# 있지도 않은 대장이 제출된 것처럼 보입니다. 표 인식에서는 반드시 건너뜁니다.
CLOSE_MARKER = "해당없음 확인"

DATA_SUFFIXES = (".csv", ".tsv", ".txt", ".xlsx", ".xlsm")


def _is_data_file(name):
    if name.startswith("~$") or name.startswith(".") or name.startswith("_읽어보기"):
        return False
    return name.lower().endswith(DATA_SUFFIXES)


def _folder_product_code(name):
    """폴더 이름에서 제품코드를 뽑습니다. 공통 폴더면 None."""
    cleaned = name.strip()
    if cleaned.lower() in [folder.lower() for folder in COMMON_FOLDERS]:
        return None
    # "HP-101" 또는 "HP-101 히알로타인 점안액" 처럼 앞머리에 코드가 오는 형태를 인정합니다.
    return cleaned.split()[0].split("_")[0].upper()


def discover(input_dir):
    """폴더 한 곳(제품 구분 없음)에서 데이터셋별 파일을 찾습니다. {dataset: [경로, ...]}"""
    found = {}
    for name in sorted(os.listdir(input_dir)):
        path = os.path.join(input_dir, name)
        if not os.path.isfile(path) or not _is_data_file(name):
            continue
        dataset = schema.detect_dataset(name)
        if dataset:
            found.setdefault(dataset, []).append(path)
    return found


def discover_tree(root, skip=None):
    """제품 폴더 구조를 훑습니다.

        입력루트/
          HP-101 히알로타인 점안액/   ← 담당자가 이 제품 자료를 올리는 곳
            배치성적서.xlsx
            일탈대장.xlsx
          HP-201 로수바틴 정/
            ...
          공통/                     ← 설비 적격성처럼 제품 공통 자료
            적격성.csv
          products_제품마스터.csv      ← 루트에 둔 파일은 제품 공통으로 봅니다

    반환값: (항목 목록, 인식 못 한 파일 목록)
    항목은 {"product_code", "dataset", "path"} 이며 공통 자료는 product_code 가 None 입니다.
    """
    entries = []
    unknown = []

    def scan(directory, product_code):
        for name in sorted(os.listdir(directory)):
            path = os.path.join(directory, name)
            if not os.path.isfile(path) or not _is_data_file(name):
                continue
            if skip and skip(name):       # 항 번호 근거 파일·마감 기록 — 대장이 아닙니다
                continue
            dataset = schema.detect_dataset(name)
            if dataset:
                entries.append({"product_code": product_code, "dataset": dataset, "path": path})
            else:
                unknown.append(path)

    scan(root, None)
    for name in sorted(os.listdir(root)):
        path = os.path.join(root, name)
        if os.path.isdir(path) and not name.startswith("."):
            scan(path, _folder_product_code(name))
    return entries, unknown


def has_product_folders(root):
    """제품 폴더 구조인지 판별합니다 (하위 폴더에 데이터 파일이 하나라도 있으면 참)."""
    for name in sorted(os.listdir(root)):
        path = os.path.join(root, name)
        if not os.path.isdir(path) or name.startswith("."):
            continue
        if any(_is_data_file(child) for child in os.listdir(path)):
            return True
    return False


def load(input_dir=None, files=None, config=None):
    """파일을 읽어 정규화합니다.

    input_dir 아래에 제품 폴더가 있으면 폴더 이름을 제품코드로 삼고, 없으면
    폴더 하나에 모든 제품 자료가 섞여 있는 것으로 봅니다.
    files 로 {dataset: [경로, ...]} 를 직접 지정할 수도 있습니다.

    항 번호로 시작하는 파일(회사 원본 양식·마감 기록)은 근거 자료입니다 —
    표준 대장 표가 아니므로 표로 읽지 않습니다. 낱말이 우연히 겹쳐 읽으려 들면
    수천 건의 적재 오류만 남고, 있지도 않은 대장이 제출된 것처럼 보입니다.

    반환값: (datasets, sources, issues, presence)
    presence 는 {"products": {제품코드: {데이터셋, ...}}, "common": {데이터셋, ...},
    "unknown": [인식 못 한 파일]} 로, "이 제품 폴더에 어떤 자료가 올라왔는가"를 담습니다.
    """
    config = config or load_config()
    matcher = item_matcher(config["items"])

    def is_evidence(name):
        return matcher(name) is not None or CLOSE_MARKER in name

    entries = []
    unknown = []
    if input_dir:
        if has_product_folders(input_dir):
            entries, unknown = discover_tree(input_dir, skip=is_evidence)
        else:
            for dataset, paths in discover(input_dir).items():
                entries.extend({"product_code": None, "dataset": dataset, "path": path}
                               for path in paths if not is_evidence(os.path.basename(path)))
            unknown = [os.path.join(input_dir, name) for name in sorted(os.listdir(input_dir))
                       if os.path.isfile(os.path.join(input_dir, name))
                       and _is_data_file(name) and schema.detect_dataset(name) is None
                       and not is_evidence(name)]
    for dataset, paths in (files or {}).items():
        entries.extend({"product_code": None, "dataset": dataset, "path": path}
                       for path in paths)

    datasets = {}
    sources = {}
    issues = []
    by_product = {}
    common = set()

    for entry in sorted(entries, key=lambda item: (item["dataset"], item["path"])):
        dataset = entry["dataset"]
        path = entry["path"]
        folder_code = entry["product_code"]
        try:
            raw = read_table(path)
        except TableError as error:
            issues.append({"source": os.path.basename(path), "row": 0, "field": "",
                           "level": "error", "message": str(error)})
            continue
        normalized, found_issues = schema.normalize(
            raw, dataset, os.path.basename(path), default_product_code=folder_code)
        datasets.setdefault(dataset, []).extend(normalized)
        issues.extend(found_issues)
        sources.setdefault(dataset, []).append({
            "file": os.path.basename(path),
            "path": os.path.abspath(path),
            "product": folder_code or "(공통)",
            "rows": len(normalized),
            "skipped": len(raw) - len(normalized),
            "sha256": _sha256(path),
            "size": os.path.getsize(path),
        })

        if folder_code:
            by_product.setdefault(folder_code, set()).add(dataset)
            continue
        # 제품 폴더가 아니라면 행에 적힌 제품코드로 제출 여부를 판단합니다.
        codes = {row.get("product_code") for row in normalized if row.get("product_code")}
        for code in codes:
            by_product.setdefault(code, set()).add(dataset)
        if not codes or any(not row.get("product_code") for row in normalized):
            common.add(dataset)

    for dataset in schema.DATASETS:
        datasets.setdefault(dataset, [])
    return datasets, sources, issues, {"products": by_product, "common": common,
                                       "unknown": unknown}


# --------------------------------------------------------------------------
# 분류 헬퍼
# --------------------------------------------------------------------------

def _matches(text, keywords):
    lowered = str(text or "").replace(" ", "").lower()
    return any(keyword.replace(" ", "").lower() in lowered for keyword in keywords)


FORM_FALLBACK = "기타제"


def form_group(text, config, name=""):
    """제형 표기를 config 의 5개 군(주사제·점안제·고형제·액제·기타제) 중 하나로 옮깁니다.

    위에서부터 처음 걸리는 군을 씁니다 — 점안제·주사제가 액제보다 앞에 있어야
    '점안액' · '주사액' 이 액제로 새지 않습니다. 제형 칸이 비어 있으면 제품명으로
    한 번 더 봅니다(마스터에 제형을 안 적는 담당자가 있습니다).
    """
    groups = config.get("dosage_forms") or {}
    for candidate in (text, name):
        if not candidate:
            continue
        for group, rule in groups.items():
            if group.startswith("_"):
                continue
            keywords = rule.get("keywords", []) if isinstance(rule, dict) else rule
            exclude = rule.get("exclude", []) if isinstance(rule, dict) else []
            if _matches(candidate, keywords) and not _matches(candidate, exclude):
                return group
    return FORM_FALLBACK if (text or name) else ""


class Classifier:
    """config 의 키워드로 구분값(일탈 · 변경 · 불만 …)을 판정합니다."""

    def __init__(self, config):
        self.keywords = config["type_keywords"]
        self.closed = config["closed_keywords"]
        self.fail = config["verdict_fail_keywords"]

    def is_kind(self, row, kind):
        return _matches(row.get("type"), self.keywords[kind])

    def is_closed(self, row):
        if row.get("closed_date"):
            return True
        return _matches(row.get("status"), self.closed)

    def capa_open(self, row):
        if not row.get("capa_no"):
            return False
        return not _matches(row.get("capa_status"), self.closed)

    def failed(self, row):
        return _matches(row.get("verdict"), self.fail)


# --------------------------------------------------------------------------
# 제품별 계산
# --------------------------------------------------------------------------

def _group(rows, key="product_code"):
    grouped = {}
    for row in rows:
        grouped.setdefault(row.get(key) or "", []).append(row)
    return grouped


def _test_summary(rows, config):
    """시험항목별 통계 · 공정능력 · 규격이탈 · 경향이탈."""
    sigma = config["thresholds"]["control_sigma"]
    by_test = {}
    for row in rows:
        by_test.setdefault(row.get("test_name") or "(미지정)", []).append(row)

    summary = []
    for test_name, test_rows in sorted(by_test.items()):
        values = [row.get("value") for row in test_rows]
        lsl = next((row.get("lsl") for row in test_rows if row.get("lsl") is not None), None)
        usl = next((row.get("usl") for row in test_rows if row.get("usl") is not None), None)
        cap = metrics.capability(
            values, lsl, usl,
            min_lots=config["thresholds"].get("cpk_min_lots", 1),
            two_sided_only=config["thresholds"].get("cpk_two_sided_only", False),
            threshold=config["thresholds"].get("cpk_sufficient", metrics.CPK_SUFFICIENT))
        limits, oot_flags = metrics.out_of_trend(values, sigma)
        oos_index = metrics.out_of_spec(values, lsl, usl)
        summary.append({
            "test_name": test_name,
            "unit": next((row.get("unit") for row in test_rows if row.get("unit")), ""),
            "lsl": lsl, "usl": usl,
            "n": cap["n"], "mean": cap["mean"], "sd": cap["sd"],
            "min": cap["min"], "max": cap["max"],
            "cp": cap["cp"], "cpk": cap["cpk"],
            "verdict": cap["verdict"], "reason": cap["reason"],
            "lcl": limits["lcl"], "ucl": limits["ucl"],
            "oos_count": len(oos_index),
            "oos_batches": [{"batch_no": test_rows[i].get("batch_no"),
                             "value": test_rows[i].get("value"),
                             "mfg_date": test_rows[i].get("mfg_date")} for i in oos_index],
            "oot_count": len(oot_flags),
            "oot_batches": [{"batch_no": test_rows[i].get("batch_no"),
                             "value": test_rows[i].get("value"),
                             "mfg_date": test_rows[i].get("mfg_date"),
                             "rule": rule}
                            for i, rule in sorted(oot_flags.items())],
            "has_spec": lsl is not None or usl is not None,
        })
    return summary


def _stability_summary(rows, config):
    horizon = config["thresholds"]["stability_horizon_months"]
    grouped = {}
    for row in rows:
        key = (row.get("condition") or "(미지정)", row.get("test_name") or "(미지정)")
        grouped.setdefault(key, []).append(row)

    summary = []
    for (condition, test_name), group in sorted(grouped.items()):
        points = sorted(group, key=lambda row: row.get("timepoint") or 0)
        timepoints = [row.get("timepoint") for row in points]
        values = [row.get("value") for row in points]
        lsl = next((row.get("lsl") for row in points if row.get("lsl") is not None), None)
        usl = next((row.get("usl") for row in points if row.get("usl") is not None), None)
        trend = metrics.stability_trend(timepoints, values, lsl, usl, horizon)
        oos_index = metrics.out_of_spec(values, lsl, usl)
        summary.append({
            "condition": condition, "test_name": test_name,
            "lsl": lsl, "usl": usl,
            "timepoints": timepoints, "values": values,
            "slope": trend["slope"], "r2": trend["r2"],
            "adverse": trend["adverse"],
            "months_to_limit": trend["months_to_limit"],
            "oos_count": len(oos_index),
            "batches": sorted({row.get("batch_no") for row in points if row.get("batch_no")}),
        })
    return summary


def _qualification_summary(rows, today, config):
    window = config["thresholds"]["qualification_due_days"]
    summary = []
    for row in rows:
        next_due = schema.parse_date(row.get("next_due"))
        days_left = (next_due - today).days if next_due else None
        if days_left is None:
            state = "일자 미기재"
        elif days_left < 0:
            state = "기한 초과"
        elif days_left <= window:
            state = "갱신 임박"
        else:
            state = "유효"
        summary.append({
            "asset": row.get("asset"), "type": row.get("type"),
            "last_qualified": row.get("last_qualified"), "next_due": row.get("next_due"),
            "status": row.get("status"), "days_left": days_left, "state": state,
        })
    return summary


# 파일 이름 맨 앞의 항 번호 — "9.2.1 조제 완료 후", "3. 허가증 ..." 의 숫자 부분
_ITEM_TOKEN = re.compile(r"^\s*(\d+(?:\.\d+)*)")
_ITEM_SEPARATORS = ". _-()[]·"


def item_matcher(items):
    """평가항목 번호로 파일 이름을 항목에 잇는 함수를 만듭니다.

    담당자가 폴더에 올리는 파일은 보고서 항 번호로 시작합니다 —
    "3. 허가증 ...pdf", "9.2.1 조제 완료 후" 폴더, "10.3, 10.4, 10.5 제조지원 ...xlsx".
    표 형식이 아니어서 적재하지 못하는 파일(PDF·스캔)도 '자료가 왔다'는 사실은
    수집 현황에 보여야 하므로, 이름만으로 항목을 알아봅니다.
    """
    ids = [row[0] for row in items]
    exact = set(ids)
    ranged = {}
    for item_id in ids:
        span = re.match(r"^(\d+(?:\.\d+)*)\.(\d+)-(\d+)$", item_id)
        if span:
            prefix, start, end = span.group(1), int(span.group(2)), int(span.group(3))
            for k in range(start, end + 1):
                ranged["%s.%d" % (prefix, k)] = item_id

    def match(name):
        found = _ITEM_TOKEN.match(name)
        if not found:
            return None
        token = found.group(1)
        rest = name[found.end():]
        # "3M필름" 같은 오인을 막습니다: 영숫자가 바로 붙으면 번호가 아닙니다.
        if rest and rest[0].isascii() and rest[0].isalnum():
            return None
        # 한 자리 번호("3", "13")는 "3분기"류 오인이 잦아 구분자를 요구합니다.
        if "." not in token and rest and rest[0] not in _ITEM_SEPARATORS:
            return None
        if token in exact:
            return token
        if token in ranged:
            return ranged[token]
        deeper = [item_id for item_id in exact if item_id.startswith(token + ".")]
        return deeper[0] if len(deeper) == 1 else None

    return match


def collect_item_files(input_dir, items):
    """제품 폴더에서 항 번호로 시작하는 파일·폴더를 찾습니다.

    반환값: {제품코드: {항목번호: [이름, ...]}}. 번호가 붙은 폴더는 안에 파일이
    하나라도 있어야 셉니다 — 빈 폴더는 자료가 아닙니다.
    """
    match = item_matcher(items)
    result = {}
    if not input_dir or not os.path.isdir(input_dir):
        return result
    for folder in sorted(os.listdir(input_dir)):
        folder_path = os.path.join(input_dir, folder)
        if not os.path.isdir(folder_path) or folder.startswith("."):
            continue
        code = _folder_product_code(folder)   # 공통 폴더는 여기서 None 이 됩니다
        if not code:
            continue
        found = item_files_in(folder_path, match)
        if found:
            result[code] = found
    return result


def zip_member_name(info):
    """압축 안 파일 이름 — Windows 가 만든 압축은 한글이 CP437 로 깨져 오므로 되돌립니다."""
    name = info.filename
    if not (info.flag_bits & 0x800):          # UTF-8 표시가 없으면 cp949 로 적힌 것
        try:
            name = name.encode("cp437").decode("cp949")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return name.replace("\\", "/").strip("/")


def item_files_in(folder_path, match, depth=2):
    """제품 폴더 하나에서 {항목번호: [이름, ...]} 을 모읍니다.

    담당자는 자료를 세 가지 모양으로 올립니다 — 항 번호 파일("3. 허가증.pdf"), 항 번호
    폴더("9.2.1 조제 완료 후/"), 그리고 **폴더째 묶은 압축**("디겐타 안연고 2026년 PQR
    필요 자료.zip"). 마지막 것은 이름에 번호가 없어 그냥 두면 자료가 0% 로 보입니다.
    번호 없는 압축과 번호 없는 중간 폴더는 지나쳐 그 안을 depth 단계까지 봅니다.
    번호가 붙은 폴더는 안에 파일이 하나라도 있어야 셉니다 — 빈 폴더는 자료가 아닙니다.
    """
    import zipfile
    found = {}
    seen = set()

    def add(item_id, label):
        if (item_id, label) in seen:
            return
        seen.add((item_id, label))
        found.setdefault(item_id, []).append(label)

    def scan_dir(path, prefix, left):
        for name in sorted(os.listdir(path)):
            if name.startswith(".") or name.startswith("~$"):
                continue
            full = os.path.join(path, name)
            label = prefix + name
            item_id = match(name)
            if os.path.isdir(full):
                if item_id:
                    if any(files for _, _, files in os.walk(full)):
                        add(item_id, label)
                elif left > 0:
                    scan_dir(full, label + "/", left - 1)
                continue
            if item_id:
                add(item_id, label)
            elif name.lower().endswith(".zip") and left > 0:
                scan_zip(full, label + "/", left - 1)

    def scan_zip(path, prefix, left):
        try:
            with zipfile.ZipFile(path) as archive:
                members = [zip_member_name(info) for info in archive.infolist()
                           if not info.is_dir()]
        except (zipfile.BadZipFile, OSError):
            return
        for member in members:
            parts = [part for part in member.split("/") if part]
            if not parts or parts[0] == "__MACOSX":
                continue
            # 번호 없는 폴더는 지나치고, 처음 만나는 번호 붙은 폴더·파일을 그 항으로 칩니다.
            for i, part in enumerate(parts[:left + 2]):
                item_id = match(part)
                if item_id:
                    add(item_id, prefix + "/".join(parts[:i + 1]))
                    break

    scan_dir(folder_path, "", depth)
    return found


FINAL_KEYWORDS = ("완성본", "제출")
FINAL_SUFFIXES = (".docx", ".doc", ".hwp", ".hwpx", ".pdf")
# 프로그램이 스스로 만든 초안을 적어 두는 파일입니다. 담당자가 실제로 작성한
# 제출본과 구분하려고 둡니다 — 이름만으로는 둘을 가릴 수 없습니다.
AUTO_DRAFT_MARKER = ".pqr_자동초안.json"


def mark_auto_draft(folder, path):
    """프로그램이 만든 초안임을 기록합니다 (파일 이름과 만든 시각)."""
    marker = os.path.join(folder, AUTO_DRAFT_MARKER)
    try:
        with open(marker, encoding="utf-8") as handle:
            drafts = json.load(handle)
    except Exception:
        drafts = {}
    if not isinstance(drafts, dict):
        drafts = {}
    try:
        drafts[os.path.basename(path)] = os.path.getmtime(path)
        with open(marker, "w", encoding="utf-8") as handle:
            json.dump(drafts, handle, ensure_ascii=False, indent=1)
    except OSError:
        pass
    return marker


def unmark_auto_draft(folder, path):
    """프로그램이 같은 이름으로 완성본(엔진)을 다시 쓰면 초안 표시를 지운다 — 그래야 완성본 단추가 선다."""
    marker = os.path.join(folder, AUTO_DRAFT_MARKER)
    try:
        with open(marker, encoding="utf-8") as handle:
            drafts = json.load(handle)
    except Exception:
        return
    if isinstance(drafts, dict) and drafts.pop(os.path.basename(path), None) is not None:
        try:
            with open(marker, "w", encoding="utf-8") as handle:
                json.dump(drafts, handle, ensure_ascii=False, indent=1)
        except OSError:
            pass


def _auto_drafts(folder):
    try:
        with open(os.path.join(folder, AUTO_DRAFT_MARKER), encoding="utf-8") as handle:
            drafts = json.load(handle)
    except Exception:
        return {}
    return drafts if isinstance(drafts, dict) else {}


# 프로그램이 만든 초안에는 이 문장이 들어 있습니다 (docx_report.DRAFT_NOTE 의 앞부분).
# 표시 파일이 없어도 문서를 열어 보면 초안인지 알 수 있습니다 — 담당자 폴더에는
# 표시가 생기기 전에 만든 초안이 이미 들어 있기 때문입니다.
AUTO_DRAFT_SIGNATURE = "프로그램이 자동 작성한 제출용 초안입니다"


def _looks_like_auto_draft(path):
    """문서 안에 초안 문구가 있으면 프로그램이 만든 초안으로 봅니다."""
    if not path.lower().endswith(".docx"):
        return False
    try:
        import zipfile
        with zipfile.ZipFile(path) as archive:
            body = archive.read("word/document.xml").decode("utf-8", "replace")
    except Exception:
        return False
    # Word 는 한 문장을 여러 run 으로 쪼개 두므로 태그를 걷어내고 봅니다.
    return AUTO_DRAFT_SIGNATURE in re.sub(r"<[^>]+>", "", body)


def _is_auto_draft(path, drafts):
    """기록된 초안이고 그 뒤로 손대지 않았으면 자동 초안으로 봅니다.

    담당자가 같은 이름으로 덮어썼다면 수정 시각이 달라지므로 '작성본' 으로 칩니다.
    기록이 없어도 문서 안의 초안 문구로 다시 확인합니다.
    """
    recorded = drafts.get(os.path.basename(path))
    if recorded is not None:
        try:
            if abs(os.path.getmtime(path) - float(recorded)) < 5:
                return True
        except (OSError, TypeError, ValueError):
            pass
    return _looks_like_auto_draft(path)


def find_final_report(folder, matcher=None, include_drafts=False):
    """제품 폴더에서 완성본(제출용) 보고서 파일을 찾습니다.

    파일 이름에 '완성본' 또는 '제출' 이 들어간 문서 파일(.docx/.doc/.hwp/.hwpx/.pdf)을
    완성본으로 봅니다. 항 번호로 시작하는 파일은 근거 자료이므로 제외합니다.

    담당자가 넣은 작성본이 있으면 그것을 먼저 씁니다 — 프로그램이 만든 초안은
    자리를 채워 두는 용도라, 완성본 단추는 실제 제출본을 열어야 합니다.
    둘 다 여럿이면 가장 최근에 고친 파일을 돌려줍니다.
    """
    if not folder or not os.path.isdir(folder):
        return None
    drafts = _auto_drafts(folder)
    authored, generated = [], []
    for name in sorted(os.listdir(folder)):
        path = os.path.join(folder, name)
        if not os.path.isfile(path) or name.startswith("~$") or name.startswith("."):
            continue
        if not name.lower().endswith(FINAL_SUFFIXES):
            continue
        if matcher and matcher(name):
            continue
        if any(keyword in name for keyword in FINAL_KEYWORDS):
            (generated if _is_auto_draft(path, drafts) else authored).append(path)
    # 프로그램이 만든 요약 초안은 '완성본' 이 아닙니다 — 담당자가 그 단추를 누르면
    # 결재본 양식의 제출용 보고서가 열려야 합니다. 초안만 있으면 없는 것으로 칩니다.
    candidates = authored or (generated if include_drafts else [])
    if not candidates:
        return None
    return max(candidates, key=os.path.getmtime)


def collect_final_reports(input_dir, items):
    """제품 폴더마다 완성본 보고서가 올라와 있는지 봅니다. {제품코드: 파일 이름}"""
    result = {}
    if not input_dir or not os.path.isdir(input_dir):
        return result
    matcher = item_matcher(items)
    for folder in sorted(os.listdir(input_dir)):
        folder_path = os.path.join(input_dir, folder)
        if not os.path.isdir(folder_path) or folder.startswith("."):
            continue
        code = _folder_product_code(folder)
        if not code:
            continue
        found = find_final_report(folder_path, matcher)
        if found:
            result[code] = os.path.basename(found)
    return result


def _checks(context, config, meta=None):
    """평가항목별 자료 상태를 config 의 item_rules 로 판정합니다.

    항목 목록이 회사 문서 번호(3 · 8.1.1 · 9.2.4 …)로 바뀔 수 있으므로 규칙을 코드에
    박지 않습니다. 자료가 아예 없으면 '미착수(n)', 있는데 미결 건이 남았으면 '진행 중(p)',
    다 확인됐으면 '완료(y)' 입니다. 건수 0 은 '해당 없음'으로 보아 완료로 칩니다.

    fields 를 적은 항목은 그 값들이 제품 마스터에 실제로 채워졌는지를 봅니다.
    '자료를 담을 파일이 있다' 와 '그 자료가 있다' 는 다릅니다 — 파일 존재만으로
    완료 처리하면 올리지도 않은 항목이 초록으로 보입니다.
    """
    has = context["has"]
    meta = meta or {}
    item_files = context.get("item_files") or {}
    rules = config.get("item_rules") or {}
    states = []
    for number, _label, _hint in config["items"]:
        rule = rules.get(number) or {}
        datasets = rule.get("datasets") or []
        fields = rule.get("fields") or []
        # 항 번호가 붙은 파일이 폴더에 있으면 그 항목의 자료는 온 것입니다.
        # 표로 못 읽는 파일(PDF·스캔)이어도 수집 자체는 됐다고 보여 줍니다.
        has_file = number in item_files
        if fields:
            filled = sum(1 for name in fields if str(meta.get(name) or "").strip())
            if has_file or filled == len(fields):
                states.append("y")
            else:
                states.append("p" if filled else "n")
            continue
        present = has_file or (any(has.get(name) for name in datasets)
                               if datasets else has.get(number, False))
        pending = any(context.get(counter) for counter in (rule.get("pending") or []))
        states.append("n" if not present else ("p" if pending else "y"))
    return states



def _reasons(context, checks):
    """지연 · 확인 필요 사유를 자동으로 붙입니다."""
    reasons = []
    if context["capa_open"]:
        reasons.append("일탈 CAPA 미완결 %d건" % context["capa_open"])
    if context["deviation_open"]:
        reasons.append("일탈 미종결 %d건" % context["deviation_open"])
    if context["oos_open"] or context["batch_fail"]:
        reasons.append("규격 부적합 조사 %d건" % (context["oos_open"] + context["batch_fail"]))
    if context["stability_adverse"]:
        reasons.append("안정성 부정적 경향 %d건" % context["stability_adverse"])
    if context["qualification_due"]:
        reasons.append("설비 적격성 갱신 %d건" % context["qualification_due"])
    if context["change_open"] or context["license_open"]:
        reasons.append("변경 미종결 %d건" % (context["change_open"] + context["license_open"]))
    if context["complaint_open"]:
        reasons.append("불만 · 회수 미종결 %d건" % context["complaint_open"])
    missing = checks.count("n")
    if missing:
        reasons.append("입력 자료 %d개 항목 미제출" % missing)
    return reasons


def _stage_index(name, stages):
    if not name:
        return None
    target = str(name).replace(" ", "")
    for index, stage in enumerate(stages):
        if stage.replace(" ", "") == target:
            return index
    for index, stage in enumerate(stages):
        if target in stage.replace(" ", "") or stage.replace(" ", "") in target:
            return index
    return None


def _month_key(value):
    date = schema.parse_date(value)
    return date.strftime("%y-%m") if date else None


def _month_labels(today, months=12):
    labels = []
    year, month = today.year, today.month
    for _ in range(months):
        labels.append("%02d-%02d" % (year % 100, month))
        month -= 1
        if month == 0:
            year, month = year - 1, 12
    labels.reverse()
    return labels


# 연간 계획서 비고에 적히는 구분 — 한 품목이 여러 PQR 건으로 갈리는 기준입니다.
SINGLE_USE_WORDS = ("1회용", "일회용")
US_EXPORT_WORD = "미국 수출용"


def _note_tokens(note):
    return [part.strip() for part in re.split(r"[/,]", note or "") if part.strip()]


def _plain_name(text):
    """제품명 비교용 — 공백과 대소문자 차이를 지웁니다."""
    return re.sub(r"\s+", "", (text or "")).strip().lower()


def _code_prefix(code):
    return re.split(r"[-_\s]", (code or "").strip().upper(), 1)[0]


def fold_placeholder_codes(master):
    """마스터 안에서 이름이 같은 두 코드 가운데 본보기 코드를 진짜 코드로 접습니다.

    담당자 입력 폴더에 제품 마스터가 둘 있었다 — 연간 계획서로 만든 330품목짜리 csv 와,
    처음 프로그램을 드릴 때 본보기로 넣어 둔 한 줄짜리 xlsx(PRD-001 레보클점안액). 둘 다
    '제품마스터' 라 함께 읽히니 레보클점안액이 QC1-5041 과 PRD-001 두 제품으로 보였고,
    마스터에 있는 코드라 reconcile_codes 도 손대지 않았다(2026-09, 두 번째 지적).

    같은 이름의 코드가 여럿이면, 마스터 대부분이 쓰는 앞머리(QC1)를 가진 코드를 진짜로 보고
    나머지(PRD)를 그리로 접는다. 앞머리가 같은 코드끼리(QC1-5041 · QC1-9999)는 고르지 않는다.
    (남은 마스터, {본보기 코드: 진짜 코드}, issues) 를 돌려줍니다.
    """
    if not master:
        return master, {}, []
    counts = {}
    for code in master:
        counts[_code_prefix(code)] = counts.get(_code_prefix(code), 0) + 1
    common = max(counts, key=counts.get)
    if counts[common] < 2:
        return master, {}, []
    by_name = {}
    for code, row in master.items():
        name = _plain_name(row.get("product_name"))
        if name:
            by_name.setdefault(name, []).append(code)
    kept, remap, issues = dict(master), {}, []
    for name, codes in by_name.items():
        real = [c for c in codes if _code_prefix(c) == common]
        others = [c for c in codes if _code_prefix(c) != common]
        if len(real) != 1 or not others:
            continue
        for wrong in others:
            kept.pop(wrong, None)
            remap[wrong] = real[0]
            issues.append({
                "source": "제품 마스터", "row": 0, "field": "product_code", "level": "warn",
                "message": "제품 마스터에 같은 제품명(%s)이 %s 와 %s 로 두 번 있어 %s 로 보았습니다 — "
                           "본보기 행이 든 옛 마스터 파일을 입력 폴더에서 지워 주세요."
                           % (master[wrong].get("product_name") or name, wrong, real[0], real[0])})
    return kept, remap, issues


def reconcile_codes(master, expanded, datasets):
    """제품 마스터에 없는 제품코드를 같은 제품명의 마스터 코드로 되돌립니다.

    자료 파일이나 폴더에 코드를 잘못 적으면(레보클점안액을 QC1-5041 대신 PRD-001 로 적는
    식) 그 코드가 마스터에 없는 새 제품처럼 잡혀, 대시보드에 자료 0% 짜리 행이 하나 더
    생기고 정작 올린 자료는 진짜 제품에 붙지 않습니다. 제품명이 마스터의 한 제품과 정확히
    같으면 코드 오기로 보고 마스터 코드로 돌립니다 — 제품 마스터가 코드의 근거입니다.

    이름이 마스터의 어느 제품과도 안 맞거나 여러 제품과 겹치면 손대지 않고 알리기만 합니다.
    (issues, {잘못된 코드: 마스터 코드}) 를 돌려줍니다 — 폴더 이름에서 온 코드로 모아 둔
    자료(항 번호 파일·최종 결재본)도 이 대응표로 같이 옮겨야 진짜 제품에 붙습니다.
    """
    known = set(expanded) | set(master)
    by_name = {}
    for source in (expanded, master):        # 마스터 이름이 나중에 덮어써 이깁니다
        for code, row in source.items():
            name = _plain_name(row.get("product_name"))
            if name:
                by_name.setdefault(name, set()).add(code)

    fixed, unknown = {}, {}
    for dataset, rows in datasets.items():
        for row in rows:
            code = (row.get("product_code") or "").strip()
            if not code or code in known:
                continue
            hits = by_name.get(_plain_name(row.get("product_name"))) or set()
            if len(hits) == 1:
                right = next(iter(hits))
                row["product_code"] = right
                key = (code, right, row.get("product_name") or "")
                fixed[key] = fixed.get(key, 0) + 1
            else:
                unknown.setdefault(code, row.get("product_name") or "")

    issues = []
    for (wrong, right, name), count in sorted(fixed.items()):
        issues.append({
            "source": "제품 마스터", "row": 0, "field": "product_code", "level": "warn",
            "message": "제품코드 %s 는 제품 마스터에 없어 같은 제품명(%s)의 %s 로 "
                       "보았습니다 — 자료 %d 행. 자료 파일이나 폴더 이름의 코드를 "
                       "고쳐 주세요." % (wrong, name, right, count)})
    for wrong, name in sorted(unknown.items()):
        issues.append({
            "source": "제품 마스터", "row": 0, "field": "product_code", "level": "warn",
            "message": "제품코드 %s(%s)는 제품 마스터에 없습니다 — 제품이 하나 더 있는 "
                       "것처럼 보입니다." % (wrong, name or "제품명 없음")})
    return issues, {wrong: right for wrong, right, _ in fixed}


def expand_product_variants(products_meta):
    """연간 계획서 비고에 따라 한 품목을 여러 PQR 건으로 나눕니다.

    - 비고가 '미국 수출용' 뿐이면 그 품목 자체가 미국 수출용입니다 — 제품명에만 표기합니다.
    - '1회용 / 다회용' 이면 기본 품명이 다회용이고, '(1회용)' 건을 하나 더 만듭니다.
    - '미국 수출용' 이 다른 구분과 함께 있으면 '(미국 수출용)' 건을 하나 더 만듭니다.
    마스터에 그 코드가 이미 있으면 손대지 않습니다 — 담당자가 적은 값이 먼저입니다.
    """
    out = dict(products_meta)
    for code, row in sorted(products_meta.items()):
        tokens = _note_tokens(row.get("note"))
        if not tokens:
            continue
        name = (row.get("product_name") or code).strip()
        us = US_EXPORT_WORD in tokens
        if us and set(tokens) == {US_EXPORT_WORD}:
            if "(미국 수출용)" not in name:
                out[code] = dict(row, product_name="%s (미국 수출용)" % name)
            continue
        extras = []
        if any(token in SINGLE_USE_WORDS for token in tokens):
            extras.append(("1회용", "%s (1회용)" % name))
        if us:
            extras.append(("미국수출용", "%s (미국 수출용)" % name))
        for suffix, label in extras:
            variant = "%s-%s" % (code, suffix)
            if variant in out:
                continue
            # 계획서가 그 건의 Lot 을 따로 적어 두지 않았으면 본 품목 합계가 붙는다 —
            # 화면에서 별표를 달 수 있게 표시해 둔다.
            out[variant] = dict(row, product_code=variant, product_name=label,
                                lots_shared=True)
    return out



def _cpk_trend(products, batches_by_product, today, months=12):
    """공정능력 부족(Cpk < 1) 제품의 월별 발생 추이 — 제형별로 나눠 셉니다.

    Cpk 는 한 해 전체 자료로 한 번 냅니다. 월별로 나누려면 '그 달에 제조된 배치를
    가진 제품 중 Cpk 부족 항목이 있는 제품이 몇 개인가' 로 봅니다. 그래야 어느 달의
    생산이 문제였는지 짚을 수 있습니다. 한 달에 여러 배치가 있어도 제품은 한 번만 셉니다.
    """
    labels = _month_labels(today, months)
    index = {label: position for position, label in enumerate(labels)}
    forms = []
    counted = {}
    for product in products:
        if not product.get("cpk_low"):
            continue
        form = product.get("form_group") or "기타제"
        if form not in counted:
            counted[form] = [set() for _ in labels]
            forms.append(form)
        for row in batches_by_product.get(product["code"], []):
            key = _month_key(row.get("mfg_date"))
            if key in index:
                counted[form][index[key]].add(product["code"])
    return {
        "months": labels,
        "series": [{"key": form, "data": [len(bucket) for bucket in counted[form]]}
                   for form in forms],
        "label": "Cpk 1 미만 제품",
    }


def _trend(deviations, changes, classifier, today, months=12):
    """최근 N개월 일탈 · OOS/OOT · 불만 발생 건수."""
    labels = []
    year, month = today.year, today.month
    for _ in range(months):
        labels.append("%02d-%02d" % (year % 100, month))
        month -= 1
        if month == 0:
            year, month = year - 1, 12
    labels.reverse()
    index = {label: position for position, label in enumerate(labels)}

    series = {"일탈": [0] * months, "OOS/OOT": [0] * months, "불만": [0] * months}
    for row in deviations:
        key = _month_key(row.get("opened_date"))
        if key not in index:
            continue
        if classifier.is_kind(row, "oos") or classifier.is_kind(row, "oot"):
            series["OOS/OOT"][index[key]] += 1
        elif classifier.is_kind(row, "deviation"):
            series["일탈"][index[key]] += 1
    for row in changes:
        key = _month_key(row.get("opened_date"))
        if key not in index:
            continue
        if (classifier.is_kind(row, "complaint") or classifier.is_kind(row, "recall")
                or classifier.is_kind(row, "return")):
            series["불만"][index[key]] += 1
    return {
        "months": labels,
        "series": [{"key": name, "data": series[name]} for name in ("일탈", "OOS/OOT", "불만")],
    }


def _leadtime(stagelog, config):
    """단계 진행 이력이 있으면 단계별 평균 소요일을 계산합니다."""
    stages = config["stages"]
    targets = config["leadtime_targets"]
    durations = {stage: [] for stage in stages}
    for _, rows in _group(stagelog).items():
        points = []
        for row in rows:
            date = schema.parse_date(row.get("entered_date"))
            index = _stage_index(row.get("stage"), stages)
            if date and index is not None:
                points.append((date, index))
        points.sort()
        for current, following in zip(points, points[1:]):
            days = (following[0] - current[0]).days
            if days >= 0:
                durations[stages[current[1]]].append(days)

    result = []
    for stage in stages[:-1]:
        samples = durations.get(stage) or []
        actual = round(sum(samples) / len(samples), 1) if samples else None
        result.append({
            "stage": stage,
            "actual": actual,
            "target": targets.get(stage),
            "samples": len(samples),
        })
    return result


def build(input_dir=None, files=None, today=None, config=None, period=None):
    """입력 → 계산 → PQR 데이터셋(dict). 이 결과가 대시보드와 보고서의 단일 원본입니다."""
    config = config or load_config()
    today = today or _dt.date.today()
    if isinstance(today, str):
        today = schema.parse_date(today)
    stages = config["stages"]
    classifier = Classifier(config)

    datasets, sources, issues, presence = load(input_dir=input_dir, files=files, config=config)
    submitted = presence["products"]
    common_datasets = presence["common"]

    # 항 번호가 붙은 파일·폴더 — 표로 못 읽는 자료도 수집 현황에는 보여야 합니다.
    item_files_by_product = collect_item_files(input_dir, config["items"]) if input_dir else {}
    final_reports = collect_final_reports(input_dir, config["items"]) if input_dir else {}
    matcher = item_matcher(config["items"])
    presence["unknown"] = [path for path in presence["unknown"]
                           if not matcher(os.path.basename(path))]

    master_meta = {row["product_code"]: row for row in datasets.get("products", [])}
    master_meta, folded, fold_issues = fold_placeholder_codes(master_meta)
    issues.extend(fold_issues)
    products_meta = expand_product_variants(master_meta)
    found_issues, remap = reconcile_codes(master_meta, products_meta, datasets)
    issues.extend(found_issues)
    for wrong, right in folded.items():          # 본보기 코드로 모아 둔 폴더 자료도 옮긴다
        remap.setdefault(wrong, right)
    for wrong, right in remap.items():
        # 폴더 이름이 코드였다면 그 폴더에 모아 둔 자료도 같이 옮깁니다.
        moved = item_files_by_product.pop(wrong, None)
        if moved:
            target = item_files_by_product.setdefault(right, {})
            for item_id, names in moved.items():
                target.setdefault(item_id, []).extend(
                    n for n in names if n not in target.get(item_id, []))
        report = final_reports.pop(wrong, None)
        if report and right not in final_reports:
            final_reports[right] = report
    batches = _group(datasets.get("batches", []))
    deviations = _group(datasets.get("deviations", []))
    changes = _group(datasets.get("changes", []))
    stability = _group(datasets.get("stability", []))

    qualification_rows = datasets.get("qualification", [])
    qualification = _group([row for row in qualification_rows if row.get("product_code")])
    shared_qualification = [row for row in qualification_rows if not row.get("product_code")]

    codes = set(products_meta)
    for group in (batches, deviations, changes, stability, qualification):
        codes.update(group)
    codes.discard("")

    period_from = period_to = None
    if period:
        period_from, period_to = period
    else:
        dates = [schema.parse_date(row.get("period_from")) for row in products_meta.values()]
        ends = [schema.parse_date(row.get("period_to")) for row in products_meta.values()]
        dates = [date for date in dates if date]
        ends = [date for date in ends if date]
        if dates:
            period_from = min(dates).isoformat()
        if ends:
            period_to = max(ends).isoformat()

    products = []
    quality = {}
    for code in sorted(codes):
        meta = products_meta.get(code, {})
        product_batches = batches.get(code, [])
        product_deviations = deviations.get(code, [])
        product_changes = changes.get(code, [])
        product_stability = stability.get(code, [])
        product_qualification = qualification.get(code, []) + shared_qualification

        material_flags = [
            classifier.is_kind(row, "material")
            or _matches(row.get("stage"), config["type_keywords"]["material"])
            for row in product_batches
        ]
        material_rows = [row for row, is_material in zip(product_batches, material_flags)
                         if is_material]
        batch_rows = [row for row, is_material in zip(product_batches, material_flags)
                      if not is_material]

        tests = _test_summary(batch_rows, config)
        stability_tests = _stability_summary(product_stability, config)
        qualification_state = _qualification_summary(product_qualification, today, config)

        deviation_rows = [row for row in product_deviations if classifier.is_kind(row, "deviation")]
        oos_rows = [row for row in product_deviations
                    if classifier.is_kind(row, "oos") or classifier.is_kind(row, "oot")]
        change_rows = [row for row in product_changes if classifier.is_kind(row, "change")
                       and not classifier.is_kind(row, "license_change")]
        license_rows = [row for row in product_changes if classifier.is_kind(row, "license_change")]
        commitment_rows = [row for row in product_changes if classifier.is_kind(row, "commitment")]
        complaint_rows = [row for row in product_changes
                          if classifier.is_kind(row, "complaint")
                          or classifier.is_kind(row, "recall")
                          or classifier.is_kind(row, "return")]

        product_has = {
            name: (name in submitted.get(code, set())) or (name in common_datasets)
            for name in schema.DATASETS
        }
        context = {
            "has": product_has,
            "item_files": item_files_by_product.get(code, {}),
            "material_rows": len(material_rows),
            "material_fail": sum(1 for row in material_rows if classifier.failed(row)),
            "batch_rows": len(batch_rows),
            "tests_without_spec": sum(1 for test in tests if not test["has_spec"]),
            "batch_fail": sum(1 for row in batch_rows if classifier.failed(row)),
            "oos_open": sum(1 for row in oos_rows if not classifier.is_closed(row)),
            "deviation_open": sum(1 for row in deviation_rows if not classifier.is_closed(row)),
            "capa_open": sum(1 for row in product_deviations if classifier.capa_open(row)),
            "change_open": sum(1 for row in change_rows if not classifier.is_closed(row)),
            "license_open": sum(1 for row in license_rows if not classifier.is_closed(row)),
            "commitment_open": sum(1 for row in commitment_rows if not classifier.is_closed(row)),
            "complaint_open": sum(1 for row in complaint_rows if not classifier.is_closed(row)),
            "stability_adverse": sum(1 for test in stability_tests if test["adverse"]),
            "stability_oos": sum(test["oos_count"] for test in stability_tests),
            "qualification_due": sum(1 for item in qualification_state
                                     if item["state"] in ("기한 초과", "갱신 임박")
                                     and not _matches(item["type"], config["type_keywords"]["contract"])),
            "contract_due": sum(1 for item in qualification_state
                                if _matches(item["type"], config["type_keywords"]["contract"])
                                and item["state"] in ("기한 초과", "갱신 임박")),
        }

        checks = _checks(context, config, meta)
        reasons = _reasons(context, checks)
        collected = checks.count("y")

        stage = _stage_index(meta.get("stage"), stages)
        if stage is None:
            stage = 0
        # 자료가 다 모이면 다음 단계(보고서 초안 작성)로 자동으로 넘어갑니다.
        if stage == 0 and checks and collected == len(checks) and len(stages) > 1:
            stage = 1

        due = meta.get("due")
        if not due and period_to:
            end = schema.parse_date(period_to)
            due = (end + _dt.timedelta(days=90)).isoformat() if end else None

        name = meta.get("product_name") or next(
            (row.get("product_name") for row in product_batches if row.get("product_name")), code)

        # 성적서 기준(시험결과가 규격을 벗어난 건)과 대장 기준(OOS/OOT 로 등록된 건)은
        # 같은 사건을 가리킬 수 있으므로 더하지 않고 따로 셉니다.
        oos_spec = sum(test["oos_count"] for test in tests)
        # 공정능력 부족(Cpk < 1) 항목 — 화면의 'Cpk 발생' 과 경향 그래프가 이 값을 씁니다.
        cpk_values = [test["cpk"] for test in tests if test.get("cpk") is not None]
        cpk_low = [test for test in tests if test.get("cpk") is not None
                   and test["cpk"] < config["thresholds"].get("cpk_sufficient",
                                                              metrics.CPK_SUFFICIENT)]
        owner = meta.get("owner") or (config.get("owners_by_form") or {}).get(
            form_group(meta.get("form", ""), config, name), "")
        products.append({
            "code": code,
            "name": name,
            "form": meta.get("form", ""),
            "form_group": form_group(meta.get("form", ""), config, name),
            "group": meta.get("group", ""),
            "item_files": item_files_by_product.get(code, {}),
            # 폴더에 완성본(제출용) 보고서가 올라오면 화면에 '완성본' 단추가 생깁니다.
            "final_report": final_reports.get(code, ""),
            # 3항(대상 제품)에 그대로 들어가는 허가 정보 — 제출용 보고서가 씁니다.
            "license": {name: meta.get(name, "") for name in
                        ("product_class", "license_no", "license_date",
                         "shelf_life", "storage")},
            # 담당자는 마스터에 적힌 값이 먼저고, 없으면 제형별 담당표에서 가져옵니다.
            "owner_source": "master" if meta.get("owner") else ("form" if owner else ""),
            "lots": meta.get("lots"),
            # 그 Lot 수가 품목 합계면 화면에서 별표를 답니다 — 계획서가 건별로 나누지
            # 않은 경우입니다(제품 마스터 비고의 '품목 합계' 또는 프로그램이 갈라낸 건).
            "lots_shared": bool(meta.get("lots_shared")
                                or "품목 합계" in (meta.get("note") or "")),
            "site": meta.get("site", ""),
            "owner": owner,
            "due": due,
            "stage": stage,
            "batches": len({row.get("batch_no") for row in batch_rows if row.get("batch_no")}),
            "dev": len(deviation_rows),
            "oos": len(oos_rows),
            "oos_spec": oos_spec,
            "cpk_low": len(cpk_low),
            "cpk_low_tests": [test["test_name"] for test in cpk_low],
            "cpk_min": min(cpk_values) if cpk_values else None,
            "chg": len(change_rows) + len(license_rows),
            "cmp": len(complaint_rows),
            "checks": checks,
            "collected": collected,
            "pct": int(round(collected / len(checks) * 100)),
            "reason": reasons[0] if reasons else "",
            "reasons": reasons,
            "submitted": sorted(submitted.get(code, set())),
            "missing_datasets": sorted(
                name for name, spec in schema.DATASETS.items()
                if name not in ("products", "stagelog")
                and not product_has[name]
            ),
        })

        quality[code] = metrics.round_all({
            "tests": tests,
            "stability": stability_tests,
            "qualification": qualification_state,
            "counts": {key: value for key, value in context.items()
                       if key not in ("has", "item_files")},
            "records": {
                "deviations": deviation_rows, "oos": oos_rows,
                "changes": change_rows + license_rows, "complaints": complaint_rows,
                "commitments": commitment_rows, "materials": material_rows,
                # 6항 제조내역은 배치 원본 행이 그대로 필요합니다.
                "batches": batch_rows,
            },
        })

    return {
        "generated_at": _dt.datetime.now().replace(microsecond=0).isoformat(),
        "today": today.isoformat(),
        "period": {"from": period_from, "to": period_to},
        "stages": stages,
        "items": config["items"],
        "dosage_forms": config.get("dosage_forms", {}),
        "products": products,
        "quality": quality,
        # 화면의 '품질 이슈 경향' 은 Cpk 1 미만 제품을 제형별로 봅니다.
        "trend": _cpk_trend(products, batches, today),
        "leadtime": _leadtime(datasets.get("stagelog", []), config),
        "sources": sources,
        "issues": issues,
        "unknown_files": [os.path.basename(path) for path in presence["unknown"]],
        "narrative": {},
    }
