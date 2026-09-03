# -*- coding: utf-8 -*-
"""연간 계획서(PQR Plan) 의 제형별 표를 읽어 제품 마스터를 맞춥니다.

계획서 6항 표는 한 품목이 여러 줄로 적히기도 합니다 — 줄마다 생산 Lot 수와 비고가 있고,
비고가 '1회용 / 다회용', '미국 수출용' 처럼 PQR 이 갈리는 기준입니다.

    2 | QC1 – 5002 | 누마렌점안액 | 16 | 1회용 / 다회용
      |            |             |  0 | 미국 수출용

이 표를 그대로 읽어 건마다 행을 만들면 대시보드의 생산 Lot 이 계획서와 같아집니다.
한 줄을 두 건이 나눠 쓰는 경우(1회용 / 다회용)는 계획서도 나누지 않으므로 같은 수를
넣고 비고에 '품목 합계' 를 남깁니다 — 화면이 그 표시를 보고 별표를 답니다.
"""
from __future__ import unicode_literals

import os
import re
import shutil
import tempfile
import zipfile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

SINGLE_USE_WORDS = ("1회용", "일회용")
US_EXPORT_WORD = "미국 수출용"
SHARED_NOTE = "품목 합계"
CODE_RE = re.compile(r"^([A-Z]{2,4}\d?)\s*[-–—]\s*(\w+)$")


def _text(el):
    return "".join(t.text or "" for t in el.iter(W + "t")).strip()


def normalize_code(text):
    """'QC1 – 5002' → 'QC1-5002'. 코드로 안 보이면 빈 문자열."""
    cleaned = " ".join((text or "").split())
    m = CODE_RE.match(cleaned)
    return "%s-%s" % (m.group(1), m.group(2)) if m else ""


def note_tokens(note):
    return [part.strip() for part in re.split(r"[/,]", note or "") if part.strip()]


def _lots(text):
    m = re.search(r"-?\d+", (text or "").replace(",", ""))
    return int(m.group()) if m else None


def read_plan(path, workdir=None):
    """연간 계획서 → {제품코드: [(생산Lot, 비고), ...]} (계획서에 적힌 줄 순서 그대로)."""
    made = None
    if path.lower().endswith(".doc"):
        from .engine import convert
        made = workdir or tempfile.mkdtemp(prefix="pqr-plan-")
        target = os.path.join(made, "plan.docx")
        convert.to_docx(path, target)
        path = target
    try:
        root = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    finally:
        if made and not workdir:
            shutil.rmtree(made, ignore_errors=True)

    found = {}
    for table in root.iter(W + "tbl"):
        rows = table.findall(W + "tr")
        head = [_text(c) for c in rows[0].findall(W + "tc")] if rows else []
        if not ("관리번호" in head and "제품명" in head and any("생산" in h for h in head)):
            continue
        code = ""
        for tr in rows[1:]:
            cells = [_text(c) for c in tr.findall(W + "tc")]
            if len(cells) < 5:
                continue
            here = normalize_code(cells[1])
            if here:                                   # 새 품목 줄
                code = here
            if not code:                               # 머리행이 이어지는 표 — 건너뛴다
                continue
            found.setdefault(code, []).append((_lots(cells[3]), cells[4]))
    return found


def entries(code, name, lines):
    """계획서 줄 → 그 품목의 PQR 건 [(제품코드, 제품명, 생산Lot, 비고)]."""
    us_lines = [ln for ln in lines if US_EXPORT_WORD in note_tokens(ln[1])]
    other = [ln for ln in lines if ln not in us_lines]
    out = []
    if us_lines and not other:                          # 품목 자체가 미국 수출용
        label = name if "(미국 수출용)" in name else "%s (미국 수출용)" % name
        return [(code, label, us_lines[0][0], us_lines[0][1])]
    for lots, note in other:
        out.append((code, name, lots, note))
        if any(token in SINGLE_USE_WORDS for token in note_tokens(note)):
            out.append(("%s-1회용" % code, "%s (1회용)" % name, lots,
                        "1회용 · %s" % SHARED_NOTE))
    for lots, note in us_lines:
        out.append(("%s-미국수출용" % code, "%s (미국 수출용)" % name, lots, note))
    return out


def apply_to_master(plan, rows, name_of=None):
    """제품 마스터 행 목록에 계획서 줄을 반영한 새 행 목록을 돌려줍니다.

    rows 는 {열이름: 값} 사전들이고 '제품코드'·'제품명'·'생산수량'·'비고' 열을 씁니다.
    계획서에 없는 품목은 그대로 둡니다. 갈라진 건은 본 품목의 다른 열을 물려받습니다.
    """
    out, seen = [], set()
    # 갈라낸 행이 이미 들어 있는 마스터에 다시 돌려도 늘어나지 않게, 본 품목이 만들어
    # 낼 코드를 미리 세어 둡니다.
    made_by_plan = set()
    for row in rows:
        lines = plan.get((row.get("제품코드") or "").strip())
        if lines:
            made_by_plan.update(e[0] for e in entries(
                (row.get("제품코드") or "").strip(), (row.get("제품명") or "").strip(), lines))
    for row in rows:
        code = (row.get("제품코드") or "").strip()
        lines = plan.get(code)
        if not lines:
            if code not in seen and code not in made_by_plan:
                out.append(row)
                seen.add(code)
            continue
        name = (name_of(code) if name_of else None) or (row.get("제품명") or "").strip()
        for new_code, label, lots, note in entries(code, name, lines):
            if new_code in seen:
                continue
            made = dict(row, 제품코드=new_code, 제품명=label, 비고=note)
            if lots is not None:
                made["생산수량"] = lots
            out.append(made)
            seen.add(new_code)
    return out
