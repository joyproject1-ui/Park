# -*- coding: utf-8 -*-
"""EDMS 결재본 서식(E-HLF-32) 찾기 — 모든 PQR 은 이 서식을 바탕으로 쓴다.

담당자 지시(2026-09): "앞으로 작성하는 모든 PQR 은 이 양식을 참고해서 작성한다."
EDMS 서식은 바닥글에 'EHLF-32/Rev.000' 이 찍혀 있고, 표지·결재표·개정 내역이 없다
(결재는 EDMS 에서 이뤄진다). 전년도 결재본(HLF-QC-126-01 양식)은 값과 문안의 근거로만
쓰고, 서식은 언제나 EDMS 것을 쓴다.

서식 파일은 제품 폴더(또는 평가항목 16 폴더, 입력 폴더의 '공통') 에 .docx 로 둔다.
파일 이름은 보지 않고 바닥글로 알아본다 — 담당자가 이름을 어떻게 붙이든 상관없다.
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


def _candidates(folder):
    if not folder or not os.path.isdir(folder):
        return
    for name in sorted(os.listdir(folder)):
        path = os.path.join(folder, name)
        if name.startswith(("~$", ".")) or not name.lower().endswith(".docx"):
            continue
        if any(w in name for w in OUTPUT_WORDS):
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
    for name in sorted(os.listdir(folder)):
        sub = os.path.join(folder, name)
        if os.path.isdir(sub) and not name.startswith("."):
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
            return max(found, key=os.path.getmtime)
    return None


def choose_base(form, previous):
    """(바탕 문서, 설명). 서식이 있으면 서식, 없으면 전년도 결재본, 둘 다 없으면 (None, 이유)."""
    if form:
        return form, "EDMS 서식"
    if previous:
        return previous, "전년도 결재본(서식 없음)"
    return None, "EDMS 결재본 서식(E-HLF-32)도 전년도 결재본(평가항목 16)도 제품 폴더에 없습니다."
