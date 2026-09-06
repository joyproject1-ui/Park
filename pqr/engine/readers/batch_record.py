# -*- coding: utf-8 -*-
"""6항에 올린 공(빈) 기록서 — 제조기록서(MO)·충전기록서(FI)·포장기록서(PK) 워드 파일 읽기.

담당자 2026-09-06: "제조단위와 포장단위, 수율 기준, 주원료(원료코드가 R 로 시작하면 주원료 / 그 외는
부원료), 부원료, 포장자재 정보 등은 여기 업로드된 공 기록서를 참고하면 돼. 모든 PQR 작성할 때 참고해."

기록서는 제품마다 서식이 조금씩 달라, 표 안의 낱말(제조단위·포장단위·수율)과 코드 꼴(R…·E…·P…)로
찾는다. 값을 지어내지 않는다 — 못 찾은 것은 비워 두고 다른 자료(수율현황표·전년도 결재본)가 채운다.

read(path) → {"kind": "조제"|"충전"|"포장", "batch_size": "82,000 g", "pack_unit": "5g x 1Tube/Case",
              "yield_spec": "95.0% 이상", "materials": [{"code", "name", "spec", "group"}], "file": 이름}
merge(records) → {"batch_size", "pack_unit", "yield_specs": {공정: 기준}, "materials": {"주원료": […], "부원료": […], "포장자재": […]}}
"""
import os
import re

import docx
from docx.oxml.ns import qn

CODE = re.compile(r"^[A-Z]{1,3}\d{3,5}[A-Z]?$")
# 수율 기준: '95.0% 이상' · '96.0 ± 3.5%' · '91.0±4.0 %' · '88 ~ 100%'
SPEC = re.compile(r"(\d+(?:\.\d+)?\s*%\s*이상|\d+(?:\.\d+)?\s*±\s*\d+(?:\.\d+)?\s*%|\d+(?:\.\d+)?\s*[~～]\s*\d+(?:\.\d+)?\s*%)")
# 제조단위 값: '82,000 g' · '204,000g (40,800 Tube)' · '145 kg'
AMOUNT = re.compile(r"\d[\d,]*(?:\.\d+)?\s*(?:kg|g|mL|ml|L|Tube|tube|개|정|EA|ea)\b")
SIZE_WORDS = ("제조단위", "배치크기", "batchsize", "표준제조량", "제조규모", "제조량", "batch")
PACK_WORDS = ("포장단위", "포장형태", "포장규격", "포장사양")


def squeeze(text):
    return re.sub(r"[\s ]+", "", text or "")


def kind_of(name):
    """파일 이름·문서 제목으로 공정을 가른다 — 충전기록서 → 충전, 포장기록서 → 포장, 나머지 → 조제."""
    n = squeeze(name)
    if "충전" in n or re.search(r"\bFI-", name):
        return "충전"
    if "포장" in n or re.search(r"\bPK-", name):
        return "포장"
    return "조제"


def _cell_texts(row):
    seen, out = set(), []
    for tc in row._tr.findall(qn("w:tc")):
        if id(tc) in seen:
            continue
        seen.add(id(tc))
        out.append("\n".join("".join(t.text or "" for t in p.iter(qn("w:t"))) for p in tc.findall(qn("w:p"))).strip())
    return out


def _value_after(cells, i):
    """라벨 칸 뒤의 첫 값 — 라벨 칸 안에 ':' 로 붙어 있으면 그 뒤."""
    own = cells[i]
    if ":" in own or "：" in own:
        tail = re.split(r"[:：]", own, 1)[1].strip()
        if tail:
            return tail
    for j in range(i + 1, min(i + 4, len(cells))):
        if cells[j].strip():
            return cells[j].strip()
    return ""


def _group_of(code, kind):
    if code.startswith("R"):
        return "주원료"
    if code.startswith("P") or kind == "포장":
        return "포장자재"
    return "부원료"


def read(path):
    document = docx.Document(path)
    name = os.path.basename(path)
    out = {"kind": kind_of(name), "batch_size": "", "pack_unit": "", "yield_spec": "", "materials": [], "file": name}
    seen_codes = set()
    for table in document.tables:
        header = None
        for row in table.rows:
            cells = _cell_texts(row)
            keys = [squeeze(c).lower() for c in cells]
            # 머리행: 코드 열·이름 열·규격 열 자리
            if header is None and any(k and any(w in k for w in ("원료코드", "자재코드", "코드", "관리번호", "code")) for k in keys) \
                    and any(k and any(w in k for w in ("원료명", "자재명", "품명", "명칭", "원/자재명", "name")) for k in keys):
                header = {"code": next(i for i, k in enumerate(keys) if any(w in k for w in ("원료코드", "자재코드", "코드", "관리번호", "code"))),
                          "name": next(i for i, k in enumerate(keys) if any(w in k for w in ("원료명", "자재명", "품명", "명칭", "원/자재명", "name"))),
                          "spec": next((i for i, k in enumerate(keys) if "규격" in k or "기준" in k), None)}
                continue
            for i, k in enumerate(keys):
                if not k:
                    continue
                if not out["batch_size"] and any(w in k for w in SIZE_WORDS) and "포장" not in k:
                    m = AMOUNT.search(_value_after(cells, i))
                    if m:                                    # '204,000 g (40,800 Tube)' 에서 '204,000g' 만
                        out["batch_size"] = re.sub(r"\s+", "", m.group(0))
                if not out["pack_unit"] and any(w in k for w in PACK_WORDS):
                    v = _value_after(cells, i)
                    if v and re.search(r"\d", v):
                        out["pack_unit"] = v.split("\n")[0].strip()
                if not out["yield_spec"] and "수율" in k:
                    joined = " ".join(cells[i:i + 4])
                    m = SPEC.search(joined)
                    if m:
                        out["yield_spec"] = re.sub(r"\s+", " ", m.group(1)).strip()
            # 원/자재 줄: 코드 꼴 칸이 있으면
            code_i = next((i for i, c in enumerate(cells) if CODE.match(c.strip())), None)
            if code_i is None:
                continue
            code = cells[code_i].strip()
            if code in seen_codes:
                continue
            if header and header["code"] < len(cells) and CODE.match(cells[header["code"]].strip()):
                name_ = cells[header["name"]].strip() if header["name"] < len(cells) else ""
                spec = cells[header["spec"]].strip() if header["spec"] is not None and header["spec"] < len(cells) else ""
            else:
                # 머리행이 없으면 코드 오른쪽의 글자 칸을 이름으로
                name_ = next((c.strip() for c in cells[code_i + 1:code_i + 3]
                              if c.strip() and re.search(r"[가-힣A-Za-z]", c) and not CODE.match(c.strip())), "")
                spec = ""
            if not name_:
                continue
            seen_codes.add(code)
            out["materials"].append({"code": code, "name": name_.split("\n")[0].strip(), "spec": spec.split("\n")[0].strip(),
                                     "group": _group_of(code, out["kind"])})
    # 표 밖의 글줄('제조단위 : 82,000 g')
    if not out["batch_size"] or not out["pack_unit"]:
        for p in document.paragraphs:
            t = p.text.strip()
            k = squeeze(t).lower()
            if not out["batch_size"] and any(w in k for w in SIZE_WORDS) and "포장" not in k:
                m = AMOUNT.search(t)
                if m:
                    out["batch_size"] = re.sub(r"\s+", "", m.group(0))
            if not out["pack_unit"] and any(w in k for w in PACK_WORDS) and ":" in t:
                out["pack_unit"] = t.split(":", 1)[1].strip()
    return out


def merge(records):
    """기록서 셋을 하나로 — 제조단위는 제조기록서, 포장단위는 포장기록서 것을 먼저 쓴다."""
    out = {"batch_size": "", "pack_unit": "", "yield_specs": {}, "materials": {"주원료": [], "부원료": [], "포장자재": []},
           "files": [r["file"] for r in records]}
    order = {"조제": 0, "충전": 1, "포장": 2}
    for r in sorted(records, key=lambda r: order.get(r["kind"], 9)):
        if r["batch_size"] and not out["batch_size"]:
            out["batch_size"] = r["batch_size"]
        if r["yield_spec"]:
            out["yield_specs"].setdefault(r["kind"], r["yield_spec"])
    for r in sorted(records, key=lambda r: -order.get(r["kind"], 9)):
        if r["pack_unit"] and not out["pack_unit"]:
            out["pack_unit"] = r["pack_unit"]
    seen = set()
    for r in sorted(records, key=lambda r: order.get(r["kind"], 9)):
        for m in r["materials"]:
            if m["code"] in seen:
                continue
            seen.add(m["code"])
            out["materials"][m["group"]].append(m)
    return out
