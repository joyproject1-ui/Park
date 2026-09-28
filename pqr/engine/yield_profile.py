# -*- coding: utf-8 -*-
"""수율 기록서 '서식 프로필' — 어느 쪽에 수율이 있는지를 **수탁사 × 기록서 서식 × 판(Rev.)**
마다 한 번만 적어 두고, 같은 서식을 쓰는 제품이 모두 물려받는다.

담당자 2026-09-27: "쪽수를 제품마다 받지 말고 수탁사 × 기록서 서식마다 한 번 받아 프로필로
저장 … 버전마다 관리하는 것으로 할께."

제품이 300품목이 넘어 제품마다 쪽수를 적게 하면 아무도 쓰지 않는다. 서식은 수탁사마다
정해져 있으므로, 한 수탁사의 정산기록서 쪽수를 한 번 적어 두면 그 수탁사 제품은 그대로 쓴다.
기록서 서식번호나 판(Rev.)이 바뀌면 **새 프로필**을 만들고 옛 것은 남겨 둔다 — 지난해 PQR 을
다시 만들 때 그때 쓰던 쪽수가 필요하다.

파일은 입력 폴더의 '공통' 폴더에 둔다(설정 파일). 사람이 열어 고칠 수 있는 JSON 이다.
"""
import datetime as _dt
import json
import os
import re

FILE_NAME = "수율-기록서-서식.json"
COMMON_FOLDERS = ("공통", "_공통", "common", "shared", "전사", "site")

HELP = ("수율이 적힌 기록서 쪽수를 수탁사·기록서 서식·판(Rev.)마다 적어 두는 파일입니다. "
        "'쪽' 은 공정 이름과 쪽 번호이고, 한 공정이 여러 쪽에 걸치면 [7, 8] 처럼 적습니다. "
        "쪽 번호는 PDF 를 열었을 때의 물리 쪽수입니다(기록서에 인쇄된 쪽수가 아닙니다). "
        "서식번호나 Rev. 가 바뀌면 새 프로필을 만드세요 — 옛 프로필은 지우지 말고 두세요.")


def folder(input_dir, create=False):
    """설정 파일을 둘 '공통' 폴더."""
    for name in COMMON_FOLDERS:
        path = os.path.join(input_dir, name)
        if os.path.isdir(path):
            return path
    path = os.path.join(input_dir, COMMON_FOLDERS[0])
    if create:
        os.makedirs(path, exist_ok=True)
    return path


def path_of(input_dir, create=False):
    return os.path.join(folder(input_dir, create), FILE_NAME)


def empty():
    return {"_설명": HELP, "profiles": [], "products": {}}


def load(input_dir):
    """프로필 파일 — 없으면 빈 꼴. 깨진 파일은 빈 꼴로 보되 지우지 않는다."""
    path = path_of(input_dir)
    if not os.path.isfile(path):
        return empty()
    try:
        with open(path, encoding="utf-8-sig") as handle:
            got = json.load(handle)
    except (OSError, ValueError):
        return empty()
    if not isinstance(got, dict):
        return empty()
    got.setdefault("_설명", HELP)
    got["profiles"] = [p for p in (got.get("profiles") or []) if isinstance(p, dict)]
    got["products"] = got.get("products") if isinstance(got.get("products"), dict) else {}
    return got


def save(input_dir, data):
    path = path_of(input_dir, create=True)
    body = dict(data or empty())
    body["_설명"] = HELP
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(body, handle, ensure_ascii=False, indent=2)
    return path


def slug(text):
    return re.sub(r"[^0-9A-Za-z가-힣._-]+", "-", str(text or "").strip()).strip("-")


def make_id(maker, record, form_no="", rev=""):
    """프로필 이름 — 수탁사·기록서·서식번호·Rev. 를 이어 붙인다. 판이 다르면 다른 프로필이다."""
    parts = [slug(maker), slug(record), slug(form_no), slug(rev)]
    return "-".join(p for p in parts if p) or "프로필"


def _norm(text):
    return re.sub(r"[\s.\-_]+", "", str(text or "")).lower()


def find(data, profile_id):
    for one in data.get("profiles") or []:
        if str(one.get("id") or "") == str(profile_id or ""):
            return one
    return None


def for_product(data, code):
    """그 제품에 배정된 프로필 — 없으면 None."""
    return find(data, (data.get("products") or {}).get(str(code or "").strip().upper()))


def assign(data, code, profile_id):
    data.setdefault("products", {})[str(code or "").strip().upper()] = str(profile_id or "")
    return data


def put(data, profile):
    """프로필을 넣거나 같은 이름이면 갈아 끼운다."""
    profile = dict(profile or {})
    profile.setdefault("id", make_id(profile.get("maker"), profile.get("record"),
                                     profile.get("form_no"), profile.get("rev")))
    profile.setdefault("checked", _dt.date.today().isoformat())
    rows = [p for p in (data.get("profiles") or []) if p.get("id") != profile["id"]]
    rows.append(profile)
    data["profiles"] = rows
    return profile


def bump(profile, rev, pages=None):
    """판(Rev.)이 바뀌었을 때 — 옛 프로필을 본떠 새 판 프로필을 만든다. 옛 것은 그대로 둔다."""
    new = dict(profile or {})
    new["rev"] = str(rev or "")
    new["id"] = make_id(new.get("maker"), new.get("record"), new.get("form_no"), new["rev"])
    new["checked"] = _dt.date.today().isoformat()
    if pages is not None:
        new["pages"] = pages
    new["from"] = str((profile or {}).get("id") or "")
    return new


def rev_changed(profile, form_no="", rev=""):
    """기록서에서 읽은 서식번호·판이 프로필과 다른가 — 다르면 쪽수를 다시 확인해야 한다."""
    if not profile:
        return False
    for key, seen in (("form_no", form_no), ("rev", rev)):
        mine, theirs = _norm(profile.get(key)), _norm(seen)
        if mine and theirs and mine != theirs:
            return True
    return False


PAGE_WORD = re.compile(r"(?:^|[,;/])\s*(?:([^,;/:]+?)\s*[:=]\s*)?([0-9,\s\-~]+)(?=$|[,;/])")


def parse_pages(text):
    """사람이 적은 쪽수를 {공정: [쪽]} 로. 공정 이름이 없으면 {"": [쪽]}.

    받아들이는 꼴: "조제 7, 충전 12, 포장 18-19" · "조제: 7 / 충전: 12" · "7,12,18-19" · "7-9"
    """
    out = {}
    text = str(text or "").replace("쪽", " ").replace("p.", " ").replace("P.", " ")
    for chunk in re.split(r"[,;/\n]+", text):
        chunk = chunk.strip()
        if not chunk:
            continue
        m = re.match(r"^(.*?)[\s:=]*([0-9][0-9\s\-~]*)$", chunk)
        if not m:
            continue
        name = " ".join((m.group(1) or "").split())
        pages = []
        for part in re.split(r"[\s]+", m.group(2).strip()):
            part = part.strip()
            if not part:
                continue
            span = re.match(r"^(\d+)\s*[-~]\s*(\d+)$", part)
            if span:
                a, b = int(span.group(1)), int(span.group(2))
                pages.extend(range(min(a, b), max(a, b) + 1))
            elif part.isdigit():
                pages.append(int(part))
        if not pages:
            continue
        out.setdefault(name, [])
        for p in pages:
            if p not in out[name]:
                out[name].append(p)
    return out


def page_list(profile):
    """프로필의 모든 쪽 번호 — 겹치지 않게, 차례대로."""
    seen = []
    for pages in (profile or {}).get("pages", {}).values():
        for p in pages if isinstance(pages, (list, tuple)) else [pages]:
            try:
                p = int(p)
            except (TypeError, ValueError):
                continue
            if p not in seen:
                seen.append(p)
    return sorted(seen)


# 파일 이름에 적은 쪽 번호 — '… 기록서 (수율 6페이지).pdf' · '수율 6, 8쪽' · '수율 6-7 p' · 'yield p6'.
# 담당자 2026-09-28: "제조기록서 스캔 파일 이름명 - 수율 페이지를 기재하면 알아서 수율 엑셀파일로
# 만들어주는게 좋겠어" · "해당 파일을 끌어오면 그런식으로 엑셀파일로 만들어주는거지".
NAME_PAGES = re.compile(
    r"(?:수율|yield)\s*[:：]?\s*(?:p\.?\s*|page\s*)?([0-9](?:[0-9,\s\-~]*[0-9])?)\s*(?:페이지|쪽|p\b|page|pages)?",
    re.I)


def pages_from_name(name):
    """파일 이름에서 수율 쪽 번호를 뽑는다 — 없으면 빈 목록."""
    stem = os.path.splitext(os.path.basename(str(name or "")))[0]
    found = NAME_PAGES.search(stem)
    if not found:
        return []
    return page_list({"pages": parse_pages(found.group(1))})
