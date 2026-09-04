# -*- coding: utf-8 -*-
"""채운 보고서를 EDMS 결재본 서식(E-HLF-32) 껍데기에 옮겨 담는다.

담당자가 EDMS 에서 하는 일과 같다 — EDMS 서식을 열고, 전년도 결재본으로 만든 본문을 그대로
옮긴 뒤 값을 갈아 끼운다. 엔진은 전년도 결재본을 바탕으로 본문을 채우므로(제품마다 다른 항·표
구조와 문안이 거기 있다), 채운 문서의 **본문 3항 이후**는 그대로 두고

  - 표지·'검토 및 승인' 결재표·개정 내역(EDMS 에는 없다)을 버리고
  - EDMS 서식의 목차(제목 문단 + 목차 표)를 맨 앞에 넣고
  - 머리글·바닥글(로고, 문서번호·Rev. No.·Page, EHLF-32/Rev.000)과 쪽 설정을 EDMS 것으로 바꾼다.

패키지는 채운 문서 쪽을 쓴다(그 문서의 스타일·번호 매기기·그림이 본문에 필요하다). EDMS 에서
가져오는 부품(머리글·바닥글·로고·목차가 쓰는 스타일)만 새 이름으로 붙여 넣는다.
"""
import copy
import os
import posixpath
import re
import zipfile

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
REL_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
W = "{%s}" % W_NS
R = "{%s}" % R_NS
FIRST_HEADING = re.compile(r"^\s*1\.\s*목\s*적")


class RehouseError(Exception):
    pass


def _read(z):
    return {n: z.read(n) for n in z.namelist()}


def _text(el):
    return "".join(t.text or "" for t in el.iter(W + "t"))


def _body_items(root):
    body = root.find(W + "body")
    return body, [c for c in body if c.tag != W + "sectPr"]


def _first_heading_index(items):
    for i, el in enumerate(items):
        if el.tag == W + "p" and FIRST_HEADING.match(_text(el)):
            return i
    return None


def _rels(parts, name):
    raw = parts.get(name)
    if raw is None:
        root = etree.Element("{%s}Relationships" % PKG_NS, nsmap={None: PKG_NS})
    else:
        root = etree.fromstring(raw)
    return root


def _next_rid(rels_root):
    used = {r.get("Id") for r in rels_root}
    n = 1
    while "rId%d" % n in used:
        n += 1
    return "rId%d" % n


def _unique_name(parts, base):
    stem, ext = os.path.splitext(base)
    cand, k = base, 1
    while cand in parts:
        k += 1
        cand = "%s_%d%s" % (stem, k, ext)
    return cand


def _copy_part_with_deps(src_parts, src_name, dst_parts, prefix="edms_"):
    """src 패키지의 부품(머리글 등)과 그 부품이 참조하는 부품(로고 그림)을 dst 에 새 이름으로 넣는다.
    돌려주는 값: dst 안의 새 이름."""
    folder = posixpath.dirname(src_name)
    base = posixpath.basename(src_name)
    new_name = _unique_name(dst_parts, posixpath.join(folder, prefix + base))
    data = src_parts[src_name]
    rels_name = posixpath.join(folder, "_rels", base + ".rels")
    if rels_name in src_parts:
        rels = etree.fromstring(src_parts[rels_name])
        for rel in rels:
            if rel.get("TargetMode") == "External":
                continue
            target = posixpath.normpath(posixpath.join(folder, rel.get("Target")))
            if target in src_parts:
                sub_new = _copy_part_with_deps(src_parts, target, dst_parts, prefix)
                rel.set("Target", posixpath.relpath(sub_new, folder))
        new_rels_name = posixpath.join(folder, "_rels", posixpath.basename(new_name) + ".rels")
        dst_parts[new_rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    dst_parts[new_name] = data
    return new_name


def _ensure_content_type(dst_parts, part_name, src_types):
    """[Content_Types].xml 에 새 부품의 형식을 등록한다(원본 패키지의 등록을 따라)."""
    ct = etree.fromstring(dst_parts["[Content_Types].xml"])
    src_ct = etree.fromstring(src_types)
    ext = posixpath.splitext(part_name)[1].lstrip(".").lower()
    defaults = {d.get("Extension").lower(): d.get("ContentType") for d in ct if d.tag == "{%s}Default" % CT_NS}
    overrides = {o.get("PartName") for o in ct if o.tag == "{%s}Override" % CT_NS}
    if "/" + part_name in overrides:
        return
    # 원본에서 그 부품의 형식을 찾는다 — Override 우선, 없으면 Default
    src_over = {o.get("PartName"): o.get("ContentType") for o in src_ct if o.tag == "{%s}Override" % CT_NS}
    src_def = {d.get("Extension").lower(): d.get("ContentType") for d in src_ct if d.tag == "{%s}Default" % CT_NS}
    ctype = None
    for name, t in src_over.items():
        if posixpath.basename(name) == posixpath.basename(part_name).replace("edms_", "", 1):
            ctype = t
            break
    if ctype:
        etree.SubElement(ct, "{%s}Override" % CT_NS, PartName="/" + part_name, ContentType=ctype)
    elif ext not in defaults and ext in src_def:
        etree.SubElement(ct, "{%s}Default" % CT_NS, Extension=ext, ContentType=src_def[ext])
    dst_parts["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)


def _style_key(el):
    """(종류, 이름) — Word 는 이 짝이 겹치면 문서를 손상으로 보고 복구를 묻는다."""
    name = el.find(W + "name")
    return (el.get(W + "type") or "paragraph", (name.get(W + "val") if name is not None else "") or "")


def _merge_styles(dst_styles, src_styles, wanted):
    """서식 쪽 스타일을 채운 문서에 보탠다. (styles.xml, {서식 styleId: 쓸 styleId}, 보탠 수)

    이름이 같은 스타일이 이미 있으면 **보태지 않고 그 스타일로 연결한다.** 한글 Word 서식의
    'a'(이름 Normal)를 그냥 넣으면 문서에 이름이 Normal 인 스타일이 둘이 되어, Word 가
    "일부 콘텐츠를 읽을 수 없습니다" 로 복구를 묻는다(2026-09-04 담당자 PC 에서 실제로 생김).
    styleId 만 겹치고 이름이 다르면 새 이름(edms_…)으로 넣는다.
    """
    dst = etree.fromstring(dst_styles)
    src = etree.fromstring(src_styles)
    dst_styles_el = dst.findall(W + "style")
    have_ids = {s.get(W + "styleId") for s in dst_styles_el}
    by_key = {_style_key(s): s.get(W + "styleId") for s in dst_styles_el}
    by_id = {s.get(W + "styleId"): s for s in src.findall(W + "style")}
    remap, brought = {}, []
    queue = [w for w in wanted if w]
    while queue:
        sid = queue.pop()
        if sid in remap or sid not in by_id:
            continue
        st = copy.deepcopy(by_id[sid])
        key = _style_key(st)
        same = by_key.get(key)
        if same is not None:                      # 이름이 같은 스타일이 이미 있다 — 그것을 쓴다
            remap[sid] = same
            continue
        new_id = sid if sid not in have_ids else "edms_%s" % sid
        st.set(W + "styleId", new_id)
        st.attrib.pop(W + "default", None)         # 기본 스타일은 종류마다 하나뿐이어야 한다
        for tag in ("basedOn", "next", "link"):
            ref = st.find(W + tag)
            if ref is not None:
                queue.append(ref.get(W + "val"))
        dst.append(st)
        have_ids.add(new_id)
        by_key[key] = new_id
        remap[sid] = new_id
        brought.append(st)
    for st in brought:                             # 보탠 스타일 안의 basedOn·next·link 도 맞춘다
        for tag in ("basedOn", "next", "link"):
            ref = st.find(W + tag)
            if ref is None:
                continue
            val = ref.get(W + "val")
            if val in remap:
                ref.set(W + "val", remap[val])
            elif val not in have_ids:              # 서식에도 없는 스타일을 가리키면 지운다
                st.remove(ref)
    return etree.tostring(dst, xml_declaration=True, encoding="UTF-8", standalone=True), remap, len(brought)


def _apply_style_remap(elements, remap):
    """복사해 온 문단·표의 스타일 참조를 실제로 쓸 styleId 로 바꾼다."""
    n = 0
    for el in elements:
        for tag in ("pStyle", "rStyle", "tblStyle"):
            for ref in el.iter(W + tag):
                val = ref.get(W + "val")
                if val in remap and remap[val] != val:
                    ref.set(W + "val", remap[val])
                    n += 1
    return n


def check_styles(styles_xml):
    """Word 가 복구를 묻게 만드는 스타일 문제를 찾는다. 문제 설명 목록(없으면 빈 목록)."""
    root = etree.fromstring(styles_xml) if isinstance(styles_xml, bytes) else styles_xml
    out, seen_key, seen_id, defaults = [], {}, set(), {}
    for st in root.findall(W + "style"):
        sid, key = st.get(W + "styleId"), _style_key(st)
        if sid in seen_id:
            out.append("styleId 가 겹칩니다: %s" % sid)
        seen_id.add(sid)
        if key in seen_key:
            out.append("이름이 겹치는 스타일: %s(%s) — %s, %s" % (key[1], key[0], seen_key[key], sid))
        seen_key[key] = sid
        if st.get(W + "default") == "1":
            if key[0] in defaults:
                out.append("기본 스타일이 둘: %s — %s, %s" % (key[0], defaults[key[0]], sid))
            defaults[key[0]] = sid
    return out


def _styles_used(elements):
    out = set()
    for el in elements:
        for tag in ("pStyle", "rStyle", "tblStyle"):
            for s in el.iter(W + tag):
                out.add(s.get(W + "val"))
    return out


def _renumber_bookmarks(body):
    counter, pending, names = 0, {}, {}
    for node in body.iter():
        if node.tag == W + "bookmarkStart":
            counter += 1
            pending.setdefault(node.get(W + "id"), []).append(str(counter))
            node.set(W + "id", str(counter))
            name = node.get(W + "name") or ""
            names[name] = names.get(name, 0) + 1
            if names[name] > 1:
                node.set(W + "name", "%s_%d" % (name, names[name]))
        elif node.tag == W + "bookmarkEnd":
            queue = pending.get(node.get(W + "id"))
            if queue:
                node.set(W + "id", queue.pop(0))


def rehouse(filled_path, form_path, out_path):
    """filled(전년도 결재본을 바탕으로 채운 .docx)의 본문을 form(EDMS 서식 .docx) 껍데기에 담아
    out 에 저장한다. {'앞부분 삭제': n, '목차 요소': n, '스타일 추가': n, '머리글·바닥글': n} 을 돌려준다."""
    with zipfile.ZipFile(filled_path) as z:
        parts = _read(z)
    with zipfile.ZipFile(form_path) as z:
        form = _read(z)
    if "word/document.xml" not in parts or "word/document.xml" not in form:
        raise RehouseError("워드 문서가 아닙니다")

    doc = etree.fromstring(parts["word/document.xml"])
    fdoc = etree.fromstring(form["word/document.xml"])
    body, items = _body_items(doc)
    fbody, fitems = _body_items(fdoc)

    # ---- 1) 채운 문서의 앞부분(표지·결재표·개정 내역·옛 목차)을 버린다 ----
    start = _first_heading_index(items)
    if start is None:
        raise RehouseError("본문에서 '1. 목적' 제목을 찾지 못했습니다")
    for el in items[:start]:
        body.remove(el)
    dropped = start

    # ---- 2) EDMS 서식의 목차(제목 문단·목차 표·그 사이 문단)를 맨 앞에 넣는다 ----
    fstart = _first_heading_index(fitems)
    if fstart is None:
        raise RehouseError("EDMS 서식에서 '1. 목적' 제목을 찾지 못했습니다")
    head = [copy.deepcopy(el) for el in fitems[:fstart]]
    # 목차 표 뒤에 남는 빈 문단 뭉치는 하나만 두고, 마지막 문단은 '앞에서 쪽 나눔' 대신 1항 제목에 건다
    kept = []
    for el in head:
        if el.tag == W + "p" and not _text(el).strip() and kept and kept[-1].tag == W + "p" and not _text(kept[-1]).strip():
            continue
        kept.append(el)
    first_heading = body[0]
    for el in reversed(kept):
        body.insert(0, el)
    ppr = first_heading.find(W + "pPr")
    if ppr is None:
        ppr = etree.SubElement(first_heading, W + "pPr")
        first_heading.insert(0, ppr)
    if ppr.find(W + "pageBreakBefore") is None:
        pbb = etree.Element(W + "pageBreakBefore")
        ppr.insert(0, pbb)
    # 서식 쪽 요소가 쓰는 스타일을 채운 문서에 보탠다
    wanted = _styles_used(kept)
    parts["word/styles.xml"], remap, added = _merge_styles(
        parts["word/styles.xml"], form["word/styles.xml"], wanted)
    _apply_style_remap(kept, remap)
    bad = check_styles(parts["word/styles.xml"])
    if bad:
        raise RehouseError("스타일이 어긋나 Word 가 열지 못합니다: %s" % "; ".join(bad))

    # ---- 3) 머리글·바닥글·쪽 설정을 EDMS 것으로 ----
    sect = body.find(W + "sectPr")
    fsect = fbody.find(W + "sectPr")
    if sect is None or fsect is None:
        raise RehouseError("구역 설정(sectPr)이 없습니다")
    rels = _rels(parts, "word/_rels/document.xml.rels")
    frels = {r.get("Id"): r for r in _rels(form, "word/_rels/document.xml.rels")}
    for ref in list(sect):
        if ref.tag in (W + "headerReference", W + "footerReference"):
            sect.remove(ref)
    swapped = 0
    for ref in fsect.findall(W + "headerReference") + fsect.findall(W + "footerReference"):
        rid = ref.get(R + "id")
        frel = frels.get(rid)
        if frel is None:
            continue
        target = posixpath.normpath(posixpath.join("word", frel.get("Target")))
        new_name = _copy_part_with_deps(form, target, parts)
        _ensure_content_type(parts, new_name, form["[Content_Types].xml"])
        for dep in [n for n in parts if n.startswith("word/media/edms_")]:
            _ensure_content_type(parts, dep, form["[Content_Types].xml"])
        new_rid = _next_rid(rels)
        etree.SubElement(rels, "{%s}Relationship" % PKG_NS, Id=new_rid, Type=frel.get("Type"),
                         Target=posixpath.relpath(new_name, "word"))
        new_ref = copy.deepcopy(ref)
        new_ref.set(R + "id", new_rid)
        sect.insert(0, new_ref)
        swapped += 1
    for tag in ("pgSz", "pgMar", "titlePg", "cols", "docGrid"):
        mine = sect.find(W + tag)
        theirs = fsect.find(W + tag)
        if mine is not None:
            sect.remove(mine)
        if theirs is not None:
            sect.append(copy.deepcopy(theirs))
    # sectPr 자식 순서: headerReference·footerReference 가 먼저, 그다음 type·pgSz·pgMar·…·cols·docGrid
    order = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
             "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
             "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"]
    rank = {n: i for i, n in enumerate(order)}
    kids = list(sect)
    for k in kids:
        sect.remove(k)
    for k in sorted(kids, key=lambda e: rank.get(e.tag.replace(W, ""), 99)):
        sect.append(k)

    _renumber_bookmarks(body)
    parts["word/document.xml"] = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    parts["word/_rels/document.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml 이 맨 앞이어야 하는 판독기가 있다
        z.writestr("[Content_Types].xml", parts["[Content_Types].xml"])
        for name, data in parts.items():
            if name != "[Content_Types].xml":
                z.writestr(name, data)
    return {"앞부분 삭제": dropped, "목차 요소": len(kept), "스타일 추가": added, "머리글·바닥글": swapped}
