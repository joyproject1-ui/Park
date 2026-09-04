# -*- coding: utf-8 -*-
"""내보내기 전 마지막 관문 — Word 가 '일부 콘텐츠를 읽을 수 없습니다' 로 복구를 묻는 문제를 찾는다.

이런 문제는 담당자 PC 에서 Word 로 열어야만 드러난다. LibreOffice 변환도, 개발 쪽 XSD 검증도
`word/document.xml` 만 보느라 놓쳤다(2026-09-04 디겐타안연고에서 두 번 생김):

  1) 스타일 이름이 겹침 — 서식의 `a`(이름 Normal)를 이미 `Normal` 이 있는 문서에 보탬
  2) `settings.xml` 의 자식 차례가 스키마와 다름 — `updateFields` 를 맨 끝에 붙임

그래서 스키마 파일 없이도 담당자 PC 에서 도는 검사를 둔다. 스키마가 정한 자식 차례
(`ooxml_order._SEQ`)와 패키지 짜임새(부품 형식 등록·관계 대상)를 본다.
"""
import posixpath
import zipfile
from xml.etree import ElementTree as ET

from . import ooxml_order
from .rehouse import check_styles

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
W = "{%s}" % W_NS
CT = "{%s}" % CT_NS


def _local(tag):
    return tag.split("}")[-1]


def bad_order(root):
    """스키마가 차례를 정한 요소 안에서 자식이 뒤집힌 곳을 찾는다. [(요소, 앞선 자식, 뒤진 자식)]"""
    out = []
    for el in root.iter():
        seq = ooxml_order._SEQ.get(_local(el.tag))
        if not seq:
            continue
        rank = {n: i for i, n in enumerate(seq)}
        last, last_name = -1, None
        for child in el:
            if not isinstance(child.tag, str):
                continue
            r = rank.get(_local(child.tag))
            if r is None:                      # 스키마에 없는 확장 요소는 차례를 따지지 않는다
                continue
            if r < last:
                out.append((_local(el.tag), last_name, _local(child.tag)))
            else:
                last, last_name = r, _local(child.tag)
    return out


def check_docx(path):
    """문제 설명 목록(없으면 빈 목록). 담당자 PC 에서도 도는 가벼운 검사다."""
    problems = []
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if not n.endswith("/")]
        if "word/document.xml" not in names:
            return ["워드 문서가 아닙니다 (word/document.xml 없음)"]

        # 1) 부품 형식 등록
        ct = ET.fromstring(z.read("[Content_Types].xml"))
        defaults = {d.get("Extension", "").lower() for d in ct if d.tag == CT + "Default"}
        overrides = {o.get("PartName") for o in ct if o.tag == CT + "Override"}
        for n in names:
            if n == "[Content_Types].xml" or posixpath.basename(n).startswith("."):
                continue
            ext = posixpath.splitext(n)[1].lstrip(".").lower()
            if "/" + n not in overrides and ext not in defaults:
                problems.append("부품 형식이 등록되지 않았습니다: %s" % n)
        for part in overrides:
            if part.lstrip("/") not in names:
                problems.append("[Content_Types].xml 이 없는 부품을 가리킵니다: %s" % part)

        # 2) 관계가 가리키는 부품
        for rels in [n for n in names if n.endswith(".rels")]:
            base = posixpath.dirname(posixpath.dirname(rels))
            seen = set()
            for rel in ET.fromstring(z.read(rels)):
                rid = rel.get("Id")
                if rid in seen:
                    problems.append("관계 Id 가 겹칩니다: %s %s" % (rels, rid))
                seen.add(rid)
                if rel.get("TargetMode") == "External":
                    continue
                target = posixpath.normpath(posixpath.join(base, rel.get("Target", "")))
                if target not in names:
                    problems.append("관계가 없는 부품을 가리킵니다: %s %s → %s" % (rels, rid, target))

        # 3) 스타일
        if "word/styles.xml" in names:
            problems.extend(check_styles(z.read("word/styles.xml")))

        # 4) 스키마가 정한 자식 차례 (settings.xml 의 updateFields 자리 같은 것)
        for n in names:
            if not n.startswith("word/") or not n.endswith(".xml"):
                continue
            try:
                root = ET.fromstring(z.read(n))
            except ET.ParseError as error:
                problems.append("XML 을 읽지 못했습니다: %s (%s)" % (n, error))
                continue
            for owner, before, after in bad_order(root)[:5]:
                problems.append("%s 의 <%s> 안에서 차례가 뒤집혔습니다: %s 뒤에 %s"
                                % (n, owner, before, after))
    return problems
