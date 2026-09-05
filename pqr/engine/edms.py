# -*- coding: utf-8 -*-
"""EDMS 결재본 서식(E-HLF-32) 찾기 — 모든 PQR 은 이 서식을 바탕으로 쓴다.

담당자 지시(2026-09): "앞으로 작성하는 모든 PQR 은 이 양식을 참고해서 작성한다."
EDMS 서식은 바닥글에 'EHLF-32/Rev.000' 이 찍혀 있고, 표지·결재표·개정 내역이 없다
(결재는 EDMS 에서 이뤄진다). 전년도 결재본(HLF-QC-126-01 양식)은 값과 문안의 근거로만
쓰고, 서식은 언제나 EDMS 것을 쓴다.

서식 파일은 제품 폴더(또는 평가항목 16 폴더, 입력 폴더의 '공통') 에 .docx 로 둔다.
바닥글로 알아본다 — 이름은 자유지만, 화면의 평가항목 (v) 'PQR 작성 공양식' 으로 올린
파일(항 번호 0)이 있으면 그것을 먼저 쓴다(담당자 지시 2026-09).
"""
import os
import re
import zipfile
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FORM_MARK = re.compile(r"E-?HLF-?32", re.I)
OUTPUT_WORDS = ("완성본", "제출")          # 프로그램이나 담당자가 만든 결과물은 서식이 아니다
COMMON_FOLDERS = ("공통", "_공통", "common", "shared")


def footer_text(path):
    """docx 의 머리글·바닥글 글자를 전부 이어 돌려준다. docx 가 아니면 빈 문자열."""
    try:
        with zipfile.ZipFile(path) as z:
            parts = [n for n in z.namelist()
                     if n.startswith("word/") and os.path.basename(n).startswith(("footer", "header"))
                     and n.endswith(".xml")]
            out = []
            for n in parts:
                root = ET.fromstring(z.read(n))
                out.append("".join(t.text or "" for t in root.iter(W + "t")))
            return "\n".join(out)
    except (zipfile.BadZipFile, OSError, ET.ParseError):
        return ""


def is_edms_form(path):
    """바닥글에 EHLF-32 가 있으면 EDMS 결재본 서식(또는 그 서식으로 쓴 문서)이다."""
    if not path or not str(path).lower().endswith(".docx"):
        return False
    return bool(FORM_MARK.search(footer_text(path)))


FORM_WORDS = ("공양식", "빈양식", "빈 양식", "양식")      # 평가항목 (v) 'PQR 작성 공양식' 으로 올린 파일
FORM_ITEM = re.compile(r"^\s*0[.\s_\-]")                 # 항 번호 0 — 화면에서 올리면 '0 PQR 작성 공양식 - …'
PREVIOUS_ITEM = re.compile(r"^\s*16[.\s_\-]")           # 항 번호 16 — 전년도 결재본은 서식이 아니다
CODE = re.compile(r"\bQC\d-\d{4}\b", re.I)


def is_named_form(path):
    """담당자가 '공양식' 이라고 올린 파일인가 — 항 번호 0 이거나 이름에 양식이 든 .docx."""
    name = os.path.basename(str(path or ""))
    return name.lower().endswith(".docx") and (bool(FORM_ITEM.match(name)) or
                                                 any(w in name for w in FORM_WORDS))


def mentions_other_code(path, code):
    """서식 안에 이 제품이 아닌 제품코드(QC1-xxxx)가 적혀 있으면 그 코드를 돌려준다.

    퀴노비드 양식을 디겐타 폴더에 공양식으로 올리는 실수를 막는다 — 양식은 제품마다 항·표가
    달라 남의 것으로 만들면 표가 통째로 남의 것이 된다.
    """
    mine = (code or "").strip().upper()
    try:
        with zipfile.ZipFile(path) as z:
            text = " ".join(z.read(n).decode("utf-8", "replace")
                            for n in z.namelist() if n.startswith("word/") and n.endswith(".xml"))
    except (zipfile.BadZipFile, OSError):
        return None
    text = re.sub(r"<[^>]+>", "", text)
    for found in CODE.findall(text):
        if found.upper() != mine:
            return found.upper()
    return None


def _candidates(folder):
    if not folder or not os.path.isdir(folder):
        return
    for name in sorted(os.listdir(folder)):
        path = os.path.join(folder, name)
        if name.startswith(("~$", ".")) or not name.lower().endswith(".docx"):
            continue
        if any(w in name for w in OUTPUT_WORDS):
            continue
        if PREVIOUS_ITEM.match(name):                 # '16. 전년도 PQR….docx' 는 결재본이지 서식이 아니다
            continue
        if os.path.isfile(path):
            yield path


def find_form(folder, depth=2):
    """제품 폴더 → 그 안의 하위 폴더(평가항목 16 등) → 입력 폴더의 '공통' 순으로 EDMS 서식을 찾는다.

    같은 곳에 여럿이면 가장 최근에 고친 파일. 없으면 None.
    """
    if not folder or not os.path.isdir(folder):
        return None
    places = [folder]
    from .collect import is_output_dir
    for name in sorted(os.listdir(folder)):
        sub = os.path.join(folder, name)
        if os.path.isdir(sub) and not name.startswith(".") and not is_output_dir(name):
            places.append(sub)
            if depth > 1:
                places.extend(os.path.join(sub, n) for n in sorted(os.listdir(sub))
                              if os.path.isdir(os.path.join(sub, n)))
    parent = os.path.dirname(os.path.abspath(folder))
    for common in COMMON_FOLDERS:
        places.append(os.path.join(parent, common))
    for place in places:
        found = [p for p in _candidates(place) if is_edms_form(p)]
        if found:
            # 담당자가 평가항목 (v) 'PQR 작성 공양식' 으로 올린 파일이 먼저다 (담당자 지시 2026-09:
            # "앞으로 PQR 보고서 작성할 때 이 양식을 사용해서 작성"). 여럿이면 가장 최근 것.
            named = [p for p in found if is_named_form(p)]
            return max(named or found, key=os.path.getmtime)
    return shipped_form()


SHIPPED_FORM = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "edms_form.docx")


def shipped_form():
    """프로그램에 든 EDMS 결재본 서식(E-HLF-32) 빈 서식 — 목차·머리글·바닥글과 항별 빈 표까지.

    담당자가 서식을 따로 두지 않아도 '보고서 작성' 이 EDMS 양식으로 나오게 한다. 폴더에 회사
    서식(.docx)이 있으면 그것이 먼저다 — 서식이 개정되면 공통 폴더에 새 것을 두면 된다.

    빈 표까지 들어 있어야 전년도 결재본을 못 읽는 때(옛 .doc 를 바꿀 길이 없는 PC)에도 이
    서식만으로 결재본 양식의 보고서를 만들 수 있다.
    """
    return SHIPPED_FORM if os.path.isfile(SHIPPED_FORM) and is_edms_form(SHIPPED_FORM) else None


def is_shipped(path):
    return bool(path) and os.path.abspath(path) == os.path.abspath(SHIPPED_FORM)


def choose_base(form, previous):
    """(바탕 문서, 설명). 서식이 있으면 서식, 없으면 전년도 결재본, 둘 다 없으면 (None, 이유)."""
    if form:
        return form, "EDMS 서식"
    if previous:
        return previous, "전년도 결재본(서식 없음)"
    return None, "EDMS 결재본 서식(E-HLF-32)도 전년도 결재본(평가항목 16)도 제품 폴더에 없습니다."
