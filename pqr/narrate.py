"""서술 문안 초안 작성 (Claude API).

집계 · 판정은 `build.py` 가 이미 끝냈습니다. 여기서는 그 결과를 근거로
"경향 평가", "결론 및 권고" 같은 **문장**의 초안만 만듭니다.

세 가지를 지킵니다.

1. **계산에 관여하지 않습니다.** 수치는 전부 build 결과에서 오고, 모델은 문장만 씁니다.
2. **보내는 내용을 먼저 볼 수 있습니다.** `--dry-run` 으로 전송될 페이로드를 그대로 출력합니다.
   (사외 전송 검토가 필요한 GMP 환경을 위한 장치입니다.)
3. **원자료를 보내지 않습니다.** 배치별 원본 값이 아니라 집계 · 판정 요약만 전송합니다.

`anthropic` 패키지가 있어야 하며(`pip install anthropic`), 인증은 SDK 기본 순서를 따릅니다
(ANTHROPIC_API_KEY 또는 `ant auth login` 프로필).
"""

import json

DEFAULT_MODEL = "claude-opus-5"

SYSTEM_PROMPT = """당신은 의약품 제조소 QA 부서의 제품품질평가(PQR) 보고서 작성을 돕습니다.

지켜야 할 원칙:
- 제공된 집계 결과에 있는 수치만 인용하십시오. 없는 수치를 만들어내지 마십시오.
- 데이터가 없는 항목은 "해당 자료가 제출되지 않았다"고 명확히 쓰고, 추정하지 마십시오.
- 규제 판단(적합/부적합 최종 결론)을 단정하지 말고, 검토자가 확인할 사항으로 제시하십시오.
- 한국어 공식 문서체(~함, ~임)로 간결하게 쓰십시오. 과장·홍보성 표현을 쓰지 마십시오.
- 이 문안은 사람이 검토·승인해야 하는 초안입니다."""

OUTPUT_SCHEMA = {
    "type": "object",
    "properties": {
        "overview": {"type": "string", "description": "평가 개요 3~4문장"},
        "capability_assessment": {"type": "string", "description": "공정능력·시험결과 경향 평가"},
        "deviation_assessment": {"type": "string", "description": "일탈·OOS·CAPA 평가"},
        "change_assessment": {"type": "string", "description": "변경·불만·회수 평가"},
        "stability_assessment": {"type": "string", "description": "안정성 경향 평가"},
        "qualification_assessment": {"type": "string", "description": "설비 적격성·위수탁 평가"},
        "conclusion": {"type": "string", "description": "종합 결론 초안"},
        "recommendations": {
            "type": "array", "items": {"type": "string"},
            "description": "후속 조치 제안 (검토자 확인 필요)",
        },
        "open_questions": {
            "type": "array", "items": {"type": "string"},
            "description": "자료가 부족해 판단할 수 없는 사항",
        },
    },
    "required": ["overview", "capability_assessment", "deviation_assessment",
                 "change_assessment", "stability_assessment", "qualification_assessment",
                 "conclusion", "recommendations", "open_questions"],
    "additionalProperties": False,
}


def build_payload(data, code):
    """모델에 보낼 요약을 만듭니다. 배치별 원본 값은 포함하지 않습니다."""
    product = next((item for item in data["products"] if item["code"] == code), None)
    if product is None:
        raise KeyError("제품 코드를 찾을 수 없습니다: %s" % code)
    quality = data.get("quality", {}).get(code, {})
    items = {key: name for key, name, _ in data["items"]}
    state_label = {"y": "수집 완료", "p": "진행 중", "n": "자료 미제출"}

    tests = [{
        "시험항목": test["test_name"], "n": test["n"], "평균": test["mean"],
        "표준편차": test["sd"], "규격": [test["lsl"], test["usl"]],
        "Cp": test["cp"], "Cpk": test["cpk"], "공정능력판정": test["verdict"],
        "규격이탈건수": test["oos_count"], "경향이탈건수": test["oot_count"],
    } for test in quality.get("tests", [])]

    stability = [{
        "조건": row["condition"], "시험항목": row["test_name"],
        "시점수": len(row["timepoints"]), "기울기(월당)": row["slope"],
        "결정계수": row["r2"], "부정적경향": row["adverse"],
        "규격도달까지(개월)": row["months_to_limit"], "규격이탈건수": row["oos_count"],
    } for row in quality.get("stability", [])]

    qualification = [{
        "대상": row["asset"], "구분": row["type"], "차기예정일": row["next_due"],
        "상태": row["state"],
    } for row in quality.get("qualification", [])]

    return {
        "제품": {"코드": product["code"], "제품명": product["name"],
               "제형": product.get("form"), "제조소": product.get("site")},
        "평가기간": data.get("period"),
        "기준일": data.get("today"),
        "배치": {"평가대상배치수": product["batches"]},
        "시험결과요약": tests,
        "이슈건수": _localize_counts(quality.get("counts", {})),
        "일탈": {"총건수": product["dev"], "OOS·OOT대장": product["oos"],
               "성적서기준규격이탈": product.get("oos_spec")},
        "변경·불만": {"변경건수": product["chg"], "불만·회수건수": product["cmp"]},
        "안정성요약": stability,
        "설비적격성": qualification,
        "평가항목_자료상태": {items[key]: state_label[state]
                       for key, state in zip([row[0] for row in data["items"]], product["checks"])},
    }


_COUNT_LABELS = {
    "material_rows": "출발물질·포장재 자료 행수", "material_fail": "출발물질 부적합",
    "batch_rows": "시험결과 행수", "tests_without_spec": "규격 미기재 시험항목",
    "batch_fail": "성적서 부적합 판정", "oos_open": "OOS 미종결",
    "deviation_open": "일탈 미종결", "capa_open": "CAPA 미완결",
    "change_open": "변경 미종결", "license_open": "허가변경 미종결",
    "commitment_open": "확약사항 미이행", "complaint_open": "불만·회수 미종결",
    "stability_adverse": "안정성 부정적 경향", "stability_oos": "안정성 규격이탈",
    "qualification_due": "설비 적격성 갱신 필요", "contract_due": "위수탁 협약 갱신 필요",
}


def _localize_counts(counts):
    return {_COUNT_LABELS.get(key, key): value for key, value in counts.items()}


def _user_message(payload):
    return (
        "다음은 한 제품의 제품품질평가(PQR) 집계 결과입니다. "
        "이 수치를 근거로 보고서 서술 문안 초안을 작성해 주십시오.\n\n"
        + json.dumps(payload, ensure_ascii=False, indent=2)
    )


def narrate_product(data, code, model=DEFAULT_MODEL, client=None, max_tokens=8000):
    """제품 하나의 서술 문안을 만듭니다. dict 를 돌려줍니다."""
    if client is None:
        try:
            import anthropic
        except ImportError:
            raise RuntimeError(
                "서술 문안 생성에는 anthropic 패키지가 필요합니다. `pip install anthropic` 후 다시 실행하세요. "
                "(집계·판정·보고서 생성은 이 패키지 없이도 동작합니다.)"
            )
        client = anthropic.Anthropic()
    payload = build_payload(data, code)
    response = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": _user_message(payload)}],
        output_config={"format": {"type": "json_schema", "schema": OUTPUT_SCHEMA}},
    )
    text = next(block.text for block in response.content if block.type == "text")
    result = json.loads(text)
    result["_meta"] = {
        "model": model,
        "generated_by": "claude-api",
        "review_required": True,
    }
    return result


def narrate(data, codes=None, model=DEFAULT_MODEL, client=None, dry_run=False, log=print):
    """여러 제품의 서술 문안을 만들어 data['narrative'] 에 채웁니다.

    dry_run 이면 API 를 호출하지 않고 전송될 내용만 돌려줍니다.
    """
    codes = codes or [product["code"] for product in data["products"]]
    if dry_run:
        preview = {code: build_payload(data, code) for code in codes}
        log("[dry-run] 전송될 내용 (%d개 제품) — 실제 호출은 하지 않았습니다." % len(codes))
        return preview

    for code in codes:
        log("서술 문안 생성 중: %s" % code)
        data.setdefault("narrative", {})[code] = narrate_product(
            data, code, model=model, client=client)
    return data["narrative"]
