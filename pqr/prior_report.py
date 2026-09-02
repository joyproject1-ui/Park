# -*- coding: utf-8 -*-
"""전년도 PQR 을 기준 본으로 삼아 새 연도 제출용 보고서를 만듭니다.

담당자가 실제로 하는 일이 그렇습니다 — 새 서식에 처음부터 쓰지 않고, 전년도
작성본을 열어 첨부 자료 값을 올해 것으로 갈아 끼웁니다. 프로그램도 같은 방식을
따릅니다. 표 안의 시험값은 원본을 봐야 하므로 담당자가 채우고, 프로그램은
서식을 그대로 물려주고 연도 표기만 새로 맞춥니다.

표준 라이브러리만 씁니다 — 담당자 PC 에 추가 설치를 요구하지 않습니다.
"""

import io
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PREVIOUS_ITEM = "0"                      # 평가항목 (0) 전년도 PQR word & excel
DOC_SUFFIXES = (".docx", ".zip")   # 담당자가 압축으로 올리는 일이 잦습니다
_OUTPUT_KEYWORDS = ("완성본", "제출")     # 우리가 만든 결과물은 기준 본이 아닙니다


def find_previous_report(folder, matcher=None):
    """제품 폴더에서 전년도 PQR 워드 파일을 찾습니다 (평가항목 0 번 파일)."""
    if not folder or not os.path.isdir(folder):
        return None
    found = []
    for name in sorted(os.listdir(folder)):
        path = os.path.join(folder, name)
        if not os.path.isfile(path) or name.startswith("~$") or name.startswith("."):
            continue
        if not name.lower().endswith(DOC_SUFFIXES):
            continue
        if any(keyword in name for keyword in _OUTPUT_KEYWORDS):
            continue
        item = matcher(name) if matcher else None
        if item == PREVIOUS_ITEM or (matcher is None and name.strip().startswith("0")):
            found.append(path)
    if not found:
        return None
    return max(found, key=os.path.getmtime)


def read_document(path):
    """기준 본 워드의 내용을 돌려줍니다 — 압축(.zip)으로 올렸으면 안에서 꺼냅니다.

    돌려주는 값: (docx 바이트, 표시용 이름)
    """
    if path.lower().endswith(".zip"):
        with zipfile.ZipFile(path) as archive:
            names = [name for name in archive.namelist()
                     if name.lower().endswith(".docx")
                     and not name.startswith("__MACOSX")
                     and not os.path.basename(name).startswith("~$")]
            if not names:
                return None, None
            # 여러 개면 가장 큰 것을 봅니다 — 보고서 본문이 부속 문서보다 큽니다.
            pick = max(names, key=lambda name: archive.getinfo(name).file_size)
            return archive.read(pick), "%s (%s)" % (os.path.basename(path),
                                                    os.path.basename(pick))
    with open(path, "rb") as handle:
        return handle.read(), os.path.basename(path)


def _paragraph_text(paragraph):
    return "".join(node.text or "" for node in paragraph.iter(W + "t"))


def _set_paragraph_text(paragraph, text):
    """문단 글자를 바꿉니다 — 첫 조각에 넣고 나머지는 비웁니다(서식은 그대로)."""
    nodes = list(paragraph.iter(W + "t"))
    if not nodes:
        return False
    nodes[0].text = text
    for node in nodes[1:]:
        node.text = ""
    return True


def _shift_years(text, mapping):
    """한 번에 바꿉니다 — 차례로 바꾸면 2024→2025→2026 처럼 두 번 밀립니다."""
    if not mapping:
        return text
    pattern = re.compile("|".join(re.escape(key) for key in
                                  sorted(mapping, key=len, reverse=True)))
    return pattern.sub(lambda found: mapping[found.group(0)], text)


def _year_mapping(previous_year, target_year):
    """연도 표기를 바꿀 짝을 만듭니다 — 네 자리 연도와 'PQR26' 같은 두 자리."""
    mapping = {}
    for offset in (0, 1):                      # 평가 연도와 차년도(작성 연도)
        old = previous_year + offset
        new = target_year + offset
        mapping[str(old)] = str(new)
        mapping["PQR%02d" % (old % 100)] = "PQR%02d" % (new % 100)
    return mapping


def rewrite_years(document_xml, previous_year, target_year):
    """본문 문단(표 밖)과 머리글의 연도 표기만 새 연도로 바꿉니다.

    표 안의 값은 시험 결과·날짜라 프로그램이 손대면 안 됩니다 — 담당자가
    원본을 보고 채웁니다. 연도가 적힌 문단(4항 일정 등)만 맞춰 둡니다.
    """
    ET.register_namespace("w", W[1:-1])
    root = ET.fromstring(document_xml)
    body = root.find(W + "body")
    mapping = _year_mapping(previous_year, target_year)
    changed = 0
    if body is not None:
        for child in list(body):
            if child.tag != W + "p":           # 표(w:tbl)는 건드리지 않습니다
                continue
            text = _paragraph_text(child)
            if not text.strip():
                continue
            shifted = _shift_years(text, mapping)
            if shifted != text and _set_paragraph_text(child, shifted):
                changed += 1
    return ET.tostring(root, encoding="unicode"), changed


def rewrite_header_years(header_xml, previous_year, target_year):
    """머리글의 제품명 줄('PQR26 제품명')처럼 연도가 붙은 이름표를 맞춥니다."""
    ET.register_namespace("w", W[1:-1])
    root = ET.fromstring(header_xml)
    mapping = _year_mapping(previous_year, target_year)
    changed = 0
    for paragraph in root.iter(W + "p"):
        text = _paragraph_text(paragraph)
        if not text.strip():
            continue
        shifted = _shift_years(text, mapping)
        if shifted != text and _set_paragraph_text(paragraph, shifted):
            changed += 1
    return ET.tostring(root, encoding="unicode"), changed


def guess_previous_year(document_xml):
    """기준 본이 어느 해 보고서인지 봅니다 — 'PQR26' 표기를 먼저 봅니다."""
    text = re.sub(r"<[^>]+>", "", document_xml)
    tag = re.search(r"PQR(\d{2})", text)
    if tag:
        return 2000 + int(tag.group(1))
    years = [int(value) for value in re.findall(r"\b(20\d{2})\b", text)]
    return max(years) if years else None


def write_from_previous(previous_path, target_path, target_year):
    """전년도 PQR 을 복제해 새 연도 보고서 파일을 만듭니다.

    돌려주는 값: {"previous": 기준 본 이름, "previous_year": 연도, "changed": 바뀐 줄 수}
    """
    payload, label = read_document(previous_path)
    if payload is None:
        return None
    with zipfile.ZipFile(io.BytesIO(payload)) as archive:
        infos = archive.infolist()
        payloads = {info.filename: archive.read(info.filename) for info in infos}
    document_xml = payloads.get("word/document.xml", b"").decode("utf-8", "replace")
    previous_year = guess_previous_year(document_xml)

    changed = 0
    if previous_year is not None and target_year is not None:
        new_document, count = rewrite_years(document_xml, previous_year, target_year)
        payloads["word/document.xml"] = new_document.encode("utf-8")
        changed += count
        for name in list(payloads):
            if re.match(r"word/header\d*\.xml$", name):
                header, count = rewrite_header_years(
                    payloads[name].decode("utf-8", "replace"), previous_year, target_year)
                payloads[name] = header.encode("utf-8")
                changed += count
    with zipfile.ZipFile(target_path, "w", zipfile.ZIP_DEFLATED) as target:
        for info in infos:
            target.writestr(info, payloads[info.filename])
    return {"previous": label, "previous_year": previous_year, "changed": changed}


ATTACHMENT_SUFFIXES = (".xlsx", ".xlsm", ".xls")


def copy_attachments(previous_path, folder, previous_year, target_year):
    """기준 본 압축에 들어 있던 첨부 엑셀을 새 연도 이름으로 함께 꺼내 둡니다.

    보고서만 새로 만들고 첨부 엑셀(경향분석 Sheet 등)이 없으면 담당자가 다시
    찾아 헤맵니다. 전년도 것을 옆에 놓아 주면 값만 갈아 끼우면 됩니다.
    """
    if not previous_path.lower().endswith(".zip"):
        return []
    written = []
    mapping = (_year_mapping(previous_year, target_year)
               if previous_year and target_year else {})
    with zipfile.ZipFile(previous_path) as archive:
        for name in archive.namelist():
            if not name.lower().endswith(ATTACHMENT_SUFFIXES):
                continue
            if name.startswith("__MACOSX") or os.path.basename(name).startswith("~$"):
                continue
            base = os.path.basename(name)
            target = os.path.join(folder, _shift_years(base, mapping) if mapping else base)
            with open(target, "wb") as handle:
                handle.write(archive.read(name))
            written.append(os.path.basename(target))
    return written
