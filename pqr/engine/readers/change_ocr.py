# -*- coding: utf-8 -*-
"""글자 없는 스캔 변경요청서를 OCR 한 글에서 필요한 값만 뽑는다.

담당자 2026-09-07: "OCR 로 변환해서 읽으면 안 되는 거야?" — 된다. 다만 한글 인식 모델이
있어야 한다(handwriting.korean_engine). 판독 글은 줄이 끊기고 글자가 틀리기도 하므로,
'변경명'·'변경 사유' 같은 이름표를 찾아 그 뒤나 다음 줄을 값으로 삼는 느슨한 방식으로 읽는다.
값을 지어내지 않는다 — 못 찾으면 빈 값으로 두고, 그 항은 '확인 필요' 로 남는다.
"""
import re

DOC_NO = re.compile(r"(CC-\d{6}-\d{2})")
# 이름표 → 결과 열쇠. OCR 이 빈칸을 흘리므로 빈칸은 무시하고 견준다.
LABELS = (("변경명", "title"), ("변경사유", "reason"), ("변경내용", "description"),
          ("관련제품", "products"), ("승인일", "approved"))
TEAM = re.compile(r"^(QA|QC|생산|포장|물류|기술지원|품질보증|품질관리|품질개선|자재|구매)\S*")
NOISE = re.compile(r"^[\s\W_]*$")


def _squeeze(text):
    return re.sub(r"\s+", "", text or "")


def _value_after(lines, i, label):
    """이름표 뒤에 붙은 글, 없으면 다음 줄들 가운데 첫 쓸 만한 글."""
    tail = re.sub(r"^.*?%s\s*[:：)]?" % re.escape(label), "", lines[i]).strip()
    if tail and not NOISE.match(tail):
        return tail
    for k in range(i + 1, min(i + 3, len(lines))):
        text = lines[k].strip()
        if text and not NOISE.match(text) and not any(_squeeze(lb) in _squeeze(text) for lb, _ in LABELS):
            return text
    return ""


def parse(lines):
    """OCR 한 줄 목록 → readers.change.read_change 와 같은 꼴."""
    out = {"doc_no": "", "title": "", "description": "", "reason": "", "products": "",
           "approved": "", "attachments": "", "target_date": "", "all_dates": [], "actions": []}
    for line in lines:
        m = DOC_NO.search(line)
        if m:
            out["doc_no"] = m.group(1)
            break
    for i, line in enumerate(lines):
        flat = _squeeze(line)
        for label, key in LABELS:
            if out[key] or label not in flat:
                continue
            out[key] = _value_after(lines, i, label)
    seen = set()
    for line in lines:
        m = TEAM.match(line.strip())
        if not m:
            continue
        act = line.strip()[m.end():].strip(" :：-")
        if act and not NOISE.match(act) and (m.group(1), act) not in seen:
            seen.add((m.group(1), act))
            out["actions"].append((m.group(1), act))
    out["all_dates"] = sorted(set(re.findall(r"\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}", " ".join(lines))))
    return out


def read(path, log=None, pages=3):
    """스캔 변경요청서를 이 PC 의 OCR 로 읽는다 — 한글 모델이 없으면 RuntimeError."""
    from .. import handwriting
    say = log or (lambda *a: None)
    engine = handwriting.korean_engine()
    if engine is None:
        raise RuntimeError("이 PC 의 OCR 에 한글 인식 모델이 없습니다 — 인터넷에서 한 번 받아야 합니다")
    lines = []
    n = min(pages, handwriting.page_count(path) or 1)
    for page_no in range(1, n + 1):
        lines += handwriting.page_text(path, page_no, engine)
    got = parse(lines)
    say("    [12] OCR(한글)로 읽음: %s (조치 %d건)" % (got["title"] or "제목 못 읽음", len(got["actions"])))
    return got
