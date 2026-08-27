"""PQR 보고서 초안 생성 (Markdown).

`build()` 결과만 있으면 사람이 검토할 수 있는 형태로 출력합니다.
서술 문안(`narrate`)이 있으면 각 절에 함께 넣고, 없으면 집계 표만 나옵니다.
"""

import datetime as _dt

ITEM_SECTIONS = [
    ("a", "출발물질 · 포장재 검토", "materials"),
    ("b", "중요 공정관리 · 완제품 시험결과", "tests"),
    ("c", "규격 부적합 배치 및 조사", "oos"),
    ("d", "중대 일탈 · 부적합 및 CAPA", "deviations"),
    ("e", "공정 · 분석법 변경", "changes"),
    ("f", "허가사항 변경", "license"),
    ("g", "안정성 모니터링 및 부정적 경향", "stability"),
    ("h", "반품 · 불만 · 회수", "complaints"),
    ("i", "이전 개선조치의 적절성", "capa"),
    ("j", "허가 후속 확약사항", "commitments"),
    ("k", "관련 설비 적격성 상태", "qualification"),
    ("l", "위 · 수탁 기술협약 검토", "contract"),
]

STATE_LABEL = {"y": "수집 완료", "p": "진행 중", "n": "자료 미제출"}


def _fmt(value, digits=2):
    if value is None or value == "":
        return "—"
    if isinstance(value, float):
        return ("%%.%df" % digits) % value
    return str(value)


def _table(header, rows):
    if not rows:
        return "_해당 데이터가 없습니다._\n"
    lines = ["| " + " | ".join(header) + " |",
             "| " + " | ".join("---" for _ in header) + " |"]
    for row in rows:
        lines.append("| " + " | ".join(_fmt(cell) for cell in row) + " |")
    return "\n".join(lines) + "\n"


def _spec(lsl, usl):
    if lsl is None and usl is None:
        return "—"
    if lsl is None:
        return "≤ %s" % _fmt(usl)
    if usl is None:
        return "≥ %s" % _fmt(lsl)
    return "%s ~ %s" % (_fmt(lsl), _fmt(usl))


def _records_table(records, columns):
    header = [label for label, _ in columns]
    rows = [[record.get(field, "") for _, field in columns] for record in records]
    return _table(header, rows)


def product_report(data, code):
    """제품 한 건의 보고서 초안(Markdown 문자열)."""
    product = next((item for item in data["products"] if item["code"] == code), None)
    if product is None:
        raise KeyError("제품 코드를 찾을 수 없습니다: %s" % code)
    quality = data.get("quality", {}).get(code, {})
    narrative = data.get("narrative", {}).get(code, {})
    period = data.get("period", {})
    checks = dict(zip([row[0] for row in data["items"]], product["checks"]))
    records = quality.get("records", {})

    out = []
    add = out.append

    add("# 제품품질평가(PQR) 보고서 초안 — %s" % product["name"])
    add("")
    add("| 항목 | 내용 |")
    add("| --- | --- |")
    add("| 제품코드 | %s |" % product["code"])
    add("| 제형 · 제조소 | %s · %s |" % (product.get("form") or "—", product.get("site") or "—"))
    add("| 평가기간 | %s ~ %s |" % (period.get("from") or "—", period.get("to") or "—"))
    add("| 평가대상 배치 | %d 배치 |" % product["batches"])
    add("| 작성 기준일 | %s |" % data.get("today"))
    add("| 담당자 | %s |" % (product.get("owner") or "—"))
    add("| 자료 수집률 | %d / %d 항목 |" % (product["collected"], len(product["checks"])))
    add("")
    add("> 이 문서는 제출된 데이터로 **자동 생성한 초안**입니다. 서술 문안이 포함된 경우 "
        "생성형 AI가 작성한 초안이며, 발행 전 담당자 검토와 승인이 필요합니다.")
    add("")

    if narrative.get("overview"):
        add("## 평가 개요")
        add("")
        add(narrative["overview"])
        add("")

    for key, title, kind in ITEM_SECTIONS:
        add("## (%s) %s" % (key, title))
        add("")
        add("**자료 상태**: %s" % STATE_LABEL[checks.get(key, "n")])
        add("")
        add(_section_body(kind, quality, records, product))
        note = _narrative_for(kind, narrative)
        if note:
            add("**평가**: %s" % note)
            add("")

    add("## 종합 결론 (초안)")
    add("")
    if narrative.get("conclusion"):
        add(narrative["conclusion"])
    else:
        add("_서술 문안이 생성되지 않았습니다. `python -m pqr narrate` 로 초안을 만들거나 "
            "담당자가 직접 작성하십시오._")
    add("")

    if narrative.get("recommendations"):
        add("### 후속 조치 제안")
        add("")
        for line in narrative["recommendations"]:
            add("- %s" % line)
        add("")
    if narrative.get("open_questions"):
        add("### 자료 부족으로 판단 보류한 사항")
        add("")
        for line in narrative["open_questions"]:
            add("- %s" % line)
        add("")

    missing = [name for (key, name, _), state in zip(data["items"], product["checks"])
               if state == "n"]
    if missing:
        add("### 미제출 자료")
        add("")
        for name in missing:
            add("- %s" % name)
        add("")

    add(_sources_section(data))
    return "\n".join(out)


def _narrative_for(kind, narrative):
    mapping = {
        "tests": "capability_assessment",
        "oos": "capability_assessment",
        "deviations": "deviation_assessment",
        "capa": "deviation_assessment",
        "changes": "change_assessment",
        "license": "change_assessment",
        "complaints": "change_assessment",
        "stability": "stability_assessment",
        "qualification": "qualification_assessment",
        "contract": "qualification_assessment",
    }
    field = mapping.get(kind)
    if not field:
        return ""
    # 같은 문안을 여러 절에 반복해 넣지 않도록, 대표 절에서만 출력합니다.
    primary = {"capability_assessment": "tests", "deviation_assessment": "deviations",
               "change_assessment": "changes", "stability_assessment": "stability",
               "qualification_assessment": "qualification"}
    if primary.get(field) != kind:
        return ""
    return narrative.get(field, "")


def _section_body(kind, quality, records, product):
    if kind == "tests":
        rows = [[test["test_name"], test["n"], test["mean"], test["sd"],
                 _spec(test["lsl"], test["usl"]), test["cp"], test["cpk"],
                 test["verdict"], test["oos_count"], test["oot_count"]]
                for test in quality.get("tests", [])]
        return _table(["시험항목", "n", "평균", "표준편차", "규격", "Cp", "Cpk",
                       "공정능력", "규격이탈", "경향이탈"], rows)

    if kind == "oos":
        rows = []
        for test in quality.get("tests", []):
            for hit in test["oos_batches"]:
                rows.append([hit.get("batch_no"), test["test_name"], hit.get("value"),
                             _spec(test["lsl"], test["usl"]), "규격 이탈", "성적서"])
            for hit in test["oot_batches"]:
                rows.append([hit.get("batch_no"), test["test_name"], hit.get("value"),
                             "%s ~ %s (관리한계)" % (_fmt(test["lcl"]), _fmt(test["ucl"])),
                             "경향 이탈 (%s)" % hit.get("rule", "3σ"), "성적서"])
        for record in records.get("oos", []):
            rows.append([record.get("record_no"), record.get("title"), "—", "—",
                         record.get("status"), "일탈대장"])
        return _table(["배치 · 번호", "시험항목 · 제목", "결과값", "규격 · 관리한계",
                       "구분 · 상태", "출처"], rows)

    if kind == "deviations":
        return _records_table(records.get("deviations", []), [
            ("번호", "record_no"), ("등급", "severity"), ("발생일", "opened_date"),
            ("제목", "title"), ("상태", "status"), ("종결일", "closed_date"),
            ("CAPA", "capa_no"), ("CAPA 상태", "capa_status")])

    if kind == "capa":
        rows = [record for record in records.get("deviations", []) if record.get("capa_no")]
        return _records_table(rows, [
            ("CAPA", "capa_no"), ("관련 일탈", "record_no"), ("CAPA 상태", "capa_status"),
            ("종결일", "closed_date")])

    if kind == "changes":
        return _records_table(records.get("changes", []), [
            ("번호", "record_no"), ("구분", "type"), ("발생일", "opened_date"),
            ("제목", "title"), ("상태", "status"), ("완료일", "closed_date")])

    if kind == "license":
        rows = [record for record in records.get("changes", [])
                if "허가" in str(record.get("type", ""))]
        return _records_table(rows, [
            ("번호", "record_no"), ("발생일", "opened_date"), ("제목", "title"),
            ("상태", "status")])

    if kind == "complaints":
        return _records_table(records.get("complaints", []), [
            ("번호", "record_no"), ("구분", "type"), ("발생일", "opened_date"),
            ("제목", "title"), ("상태", "status")])

    if kind == "commitments":
        return _records_table(records.get("commitments", []), [
            ("번호", "record_no"), ("발생일", "opened_date"), ("제목", "title"),
            ("상태", "status")])

    if kind == "materials":
        # 배치를 전부 나열하면 보고서가 길어지기만 하므로 건수와 부적합만 씁니다.
        rows = records.get("materials", [])
        if not rows:
            return "_출발물질 · 포장재 자료가 제출되지 않았습니다._\n"
        failed = [row for row in rows if "부적합" in str(row.get("verdict", ""))]
        summary = _table(["검토 대상 행", "부적합", "대상 배치 수"],
                         [[len(rows), len(failed),
                           len({row.get("batch_no") for row in rows if row.get("batch_no")})]])
        if failed:
            summary += "\n부적합 내역:\n\n" + _table(
                ["배치", "제조일", "판정"],
                [[row.get("batch_no"), row.get("mfg_date"), row.get("verdict")]
                 for row in failed])
        return summary

    if kind == "stability":
        rows = [[row["condition"], row["test_name"], len(row["timepoints"]),
                 _spec(row["lsl"], row["usl"]), row["slope"], row["r2"],
                 "예" if row["adverse"] else "아니오", row["oos_count"]]
                for row in quality.get("stability", [])]
        return _table(["조건", "시험항목", "시점수", "규격", "기울기(월당)", "R²",
                       "부정적 경향", "규격이탈"], rows)

    if kind == "qualification":
        rows = [[row["asset"], row["type"], row["last_qualified"], row["next_due"], row["state"]]
                for row in quality.get("qualification", [])
                if "위수탁" not in str(row.get("type", ""))]
        return _table(["대상", "구분", "최종 적격성일", "차기 예정일", "상태"], rows)

    if kind == "contract":
        rows = [[row["asset"], row["type"], row["last_qualified"], row["next_due"], row["state"]]
                for row in quality.get("qualification", [])
                if "위수탁" in str(row.get("type", ""))]
        return _table(["수탁처 · 대상", "구분", "체결 · 갱신일", "차기 예정일", "상태"], rows)

    return "_해당 데이터가 없습니다._\n"


def _sources_section(data):
    lines = ["## 부록 — 데이터 출처", "",
             "아래 파일로부터 계산했습니다. SHA-256 은 원본 파일이 바뀌지 않았음을 확인하는 데 씁니다.", ""]
    rows = []
    for dataset, entries in sorted(data.get("sources", {}).items()):
        for entry in entries:
            rows.append([dataset, entry["file"], entry["rows"], entry["skipped"],
                         entry["sha256"][:16] + "…"])
    lines.append(_table(["데이터셋", "파일", "적재 행", "건너뛴 행", "SHA-256"], rows))
    errors = [issue for issue in data.get("issues", []) if issue["level"] == "error"]
    if errors:
        lines.append("### 적재 오류")
        lines.append("")
        lines.append(_table(["파일", "행", "항목", "내용"],
                            [[e["source"], e["row"], e["field"], e["message"]] for e in errors]))
    lines.append("")
    lines.append("_생성 시각: %s · 생성 도구: pqr_" % data.get("generated_at"))
    return "\n".join(lines)


def summary_report(data):
    """전체 제품 요약(Markdown)."""
    period = data.get("period", {})
    lines = ["# 제품품질평가(PQR) 진행 요약", "",
             "평가기간 %s ~ %s · 기준일 %s · 대상 %d품목" % (
                 period.get("from") or "—", period.get("to") or "—",
                 data.get("today"), len(data["products"])), ""]
    stages = data["stages"]
    rows = []
    for product in sorted(data["products"], key=lambda item: (item.get("due") or "9999")):
        rows.append([product["code"], product["name"], stages[product["stage"]],
                     product.get("due") or "—",
                     "%d/%d" % (product["collected"], len(product["checks"])),
                     product["dev"], product["oos"], product["chg"], product["cmp"],
                     product.get("reason") or "—"])
    lines.append(_table(["코드", "제품명", "단계", "마감일", "자료수집",
                         "일탈", "OOS/OOT", "변경", "불만", "확인 필요"], rows))
    lines.append("")
    lines.append(_sources_section(data))
    return "\n".join(lines)


def write_reports(data, out_dir, codes=None):
    """제품별 보고서와 요약을 파일로 씁니다. 만든 파일 경로 목록을 돌려줍니다."""
    import os
    os.makedirs(out_dir, exist_ok=True)
    written = []
    path = os.path.join(out_dir, "PQR_요약.md")
    with open(path, "w", encoding="utf-8") as handle:
        handle.write(summary_report(data))
    written.append(path)
    for product in data["products"]:
        if codes and product["code"] not in codes:
            continue
        path = os.path.join(out_dir, "PQR_%s.md" % product["code"])
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(product_report(data, product["code"]))
        written.append(path)
    return written
