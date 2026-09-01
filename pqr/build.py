"""입력 파일을 읽어 PQR 데이터셋(집계 · 판정 결과)을 만듭니다.

`build()` 한 번이면 파일 적재 → 정규화 → 계산 → 판정까지 끝나고,
결과 dict 하나에 대시보드용 값과 보고서용 상세 값이 모두 들어갑니다.
계산에 쓴 파일의 이름 · 크기 · SHA-256 은 `sources` 에 남습니다 (감사추적용).
"""

import datetime as _dt
import hashlib
import json
import os

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


def discover_tree(root):
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


def load(input_dir=None, files=None):
    """파일을 읽어 정규화합니다.

    input_dir 아래에 제품 폴더가 있으면 폴더 이름을 제품코드로 삼고, 없으면
    폴더 하나에 모든 제품 자료가 섞여 있는 것으로 봅니다.
    files 로 {dataset: [경로, ...]} 를 직접 지정할 수도 있습니다.

    반환값: (datasets, sources, issues, presence)
    presence 는 {"products": {제품코드: {데이터셋, ...}}, "common": {데이터셋, ...},
    "unknown": [인식 못 한 파일]} 로, "이 제품 폴더에 어떤 자료가 올라왔는가"를 담습니다.
    """
    entries = []
    unknown = []
    if input_dir:
        if has_product_folders(input_dir):
            entries, unknown = discover_tree(input_dir)
        else:
            for dataset, paths in discover(input_dir).items():
                entries.extend({"product_code": None, "dataset": dataset, "path": path}
                               for path in paths)
            unknown = [os.path.join(input_dir, name) for name in sorted(os.listdir(input_dir))
                       if os.path.isfile(os.path.join(input_dir, name))
                       and _is_data_file(name) and schema.detect_dataset(name) is None]
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


def _checks(context, config):
    """평가항목별 자료 상태를 config 의 item_rules 로 판정합니다.

    항목 목록이 회사 문서 번호(3 · 8.1.1 · 9.2.4 …)로 바뀔 수 있으므로 규칙을 코드에
    박지 않습니다. 자료가 아예 없으면 '미착수(n)', 있는데 미결 건이 남았으면 '진행 중(p)',
    다 확인됐으면 '완료(y)' 입니다. 건수 0 은 '해당 없음'으로 보아 완료로 칩니다.
    """
    has = context["has"]
    rules = config.get("item_rules") or {}
    states = []
    for number, _label, _hint in config["items"]:
        rule = rules.get(number) or {}
        datasets = rule.get("datasets") or []
        # datasets 를 안 적은 항목(허가증 등)은 파일을 올렸는지만 봅니다.
        present = any(has.get(name) for name in datasets) if datasets else has.get(number, False)
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

    datasets, sources, issues, presence = load(input_dir=input_dir, files=files)
    submitted = presence["products"]
    common_datasets = presence["common"]

    products_meta = {row["product_code"]: row for row in datasets.get("products", [])}
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

        checks = _checks(context, config)
        reasons = _reasons(context, checks)
        collected = checks.count("y")

        stage = _stage_index(meta.get("stage"), stages)
        if stage is None:
            stage = 2 if collected == len(checks) else 1

        due = meta.get("due")
        if not due and period_to:
            end = schema.parse_date(period_to)
            due = (end + _dt.timedelta(days=90)).isoformat() if end else None

        name = meta.get("product_name") or next(
            (row.get("product_name") for row in product_batches if row.get("product_name")), code)

        # 성적서 기준(시험결과가 규격을 벗어난 건)과 대장 기준(OOS/OOT 로 등록된 건)은
        # 같은 사건을 가리킬 수 있으므로 더하지 않고 따로 셉니다.
        oos_spec = sum(test["oos_count"] for test in tests)
        products.append({
            "code": code,
            "name": name,
            "form": meta.get("form", ""),
            "form_group": form_group(meta.get("form", ""), config, name),
            "group": meta.get("group", ""),
            "lots": meta.get("lots"),
            "site": meta.get("site", ""),
            "owner": meta.get("owner", ""),
            "due": due,
            "stage": stage,
            "batches": len({row.get("batch_no") for row in batch_rows if row.get("batch_no")}),
            "dev": len(deviation_rows),
            "oos": len(oos_rows),
            "oos_spec": oos_spec,
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
            "counts": {key: value for key, value in context.items() if key != "has"},
            "records": {
                "deviations": deviation_rows, "oos": oos_rows,
                "changes": change_rows + license_rows, "complaints": complaint_rows,
                "commitments": commitment_rows, "materials": material_rows,
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
        "trend": _trend(datasets.get("deviations", []), datasets.get("changes", []),
                        classifier, today),
        "leadtime": _leadtime(datasets.get("stagelog", []), config),
        "sources": sources,
        "issues": issues,
        "unknown_files": [os.path.basename(path) for path in presence["unknown"]],
        "narrative": {},
    }
