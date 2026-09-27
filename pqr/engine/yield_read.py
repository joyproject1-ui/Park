# -*- coding: utf-8 -*-
"""'수율 판독' 단추 — 기록서 스캔에서 **수율이 적힌 쪽만** 읽어 7항 수율현황표를 만든다.

담당자 2026-09-27: "우리 팀은 직접 작성하는 것이 편하다는 의견이고 다른 팀은 기록서 스캔본에서
수율이 기재된 페이지를 직접 AI 가 읽어서 정리해줬으면 좋겠다는 의견이야."

그래서 두 길을 다 둔다.
  · 담당자가 만든 '7. 수율현황표 - 개인.xlsx' 가 있으면 보고서는 늘 **그것**을 쓴다.
  · 없으면 이 판독이 만든 '7. 수율현황표 - 판독.xlsx' 를 쓴다.
사람 값이 언제나 기계 값을 이긴다 — 판독은 원본을 고치지 않고 새 파일을 만든다.

쪽은 제품마다 받지 않는다. **수탁사 × 기록서 서식 × 판(Rev.)** 프로필(yield_profile)이 정해
주므로, 같은 서식을 쓰는 제품은 쪽수를 다시 적지 않아도 된다. 스캔 740쪽을 다 훑으면 약 119만
토큰이지만 지정한 쪽만 읽으면 그 1/20 이다.

판독값은 그대로 결재에 올라가지 않는다 — 근거 쪽수를 비고에 남기고, 흐린 칸·기준 밖 값은
'확인 필요' 로 적어 둔다.
"""
import datetime as _dt
import json
import os
import re

from . import collect as collect_mod
from . import yield_profile as profile_mod

READING_NAME = "PQR 수율 판독.json"
SHEET_NAME = "7. 수율현황표 - 판독.xlsx"
OWN_SHEET = re.compile(r"^7[.\s].*수율현황표(?!.*판독)", re.I)
LOG_NAME = "PQR 수율 판독 기록.txt"
DPI = 200


def own_sheet(folder):
    """담당자가 직접 만든 수율현황표 — 있으면 판독값보다 이것이 먼저다."""
    for path in collect_mod.discover(folder).get("7", []):
        name = os.path.basename(path)
        if name.lower().endswith((".xlsx", ".xls")) and OWN_SHEET.match(name):
            return path
    return None


def scans(folder):
    """수율을 읽을 기록서 스캔 — 7항에 없으면 6항(제조기록서)에서 찾는다."""
    got = collect_mod.discover(folder)
    out = []
    for item in ("7", "6"):
        for path in got.get(item, []):
            if path.lower().endswith(".pdf"):
                out.append(path)
        if out:
            break
    return collect_mod.same_files_once(out)


def apply_names(values, rules):
    """프로필의 공정 이름 대응표를 먹인다 — '캡슐충전' → '충전'."""
    table = (rules or {}).get("names") or {}
    if not table:
        return dict(values or {})
    out = {}
    for name, value in (values or {}).items():
        out[str(table.get(name, name))] = value
    return out


def _spec_range(text):
    """'99.5±0.45%' · '98.0~102.0' · '98.0 이상' → (하한, 상한). 못 읽으면 (None, None)."""
    text = str(text or "").replace("%", "").replace(" ", "")
    m = re.match(r"^([\d.]+)[±\+\-/]{1,2}([\d.]+)$", text)
    if m:
        try:
            mid, half = float(m.group(1)), float(m.group(2))
            return mid - half, mid + half
        except ValueError:
            return None, None
    m = re.match(r"^([\d.]+)[~∼\-]([\d.]+)$", text)
    if m:
        try:
            return float(m.group(1)), float(m.group(2))
        except ValueError:
            return None, None
    m = re.match(r"^([\d.]+)이상$", text)
    if m:
        try:
            return float(m.group(1)), None
        except ValueError:
            return None, None
    return None, None


def check(reading, known_lots=()):
    """판독값을 검산해 경고 목록을 만든다 — 값을 고치지는 않는다."""
    notes = []
    known = {str(l).strip().upper() for l in known_lots or []}
    for one in reading.get("lots") or []:
        if known and one["lot"] not in known:
            notes.append("%s — 6항 제조내역에 없는 제조번호입니다. 제조번호를 확인하세요"
                         % one["lot"])
        for name in one.get("unsure") or []:
            notes.append("%s %s — 흐려서 읽지 못했습니다. 기록서를 보고 직접 채우세요"
                         % (one["lot"], name))
        for name, value in (one.get("values") or {}).items():
            lo, hi = _spec_range((reading.get("specs") or {}).get(name))
            if lo is not None and value < lo or (hi is not None and value > hi):
                notes.append("%s %s %.2f%% — 수율 기준(%s)을 벗어납니다. 일탈 여부를 확인하세요"
                             % (one["lot"], name, value,
                                (reading.get("specs") or {}).get(name) or "?"))
    if not (reading.get("lots") or []):
        notes.append("읽어 낸 제조번호가 없습니다 — 프로필의 쪽수가 맞는지 확인하세요")
    return notes


def process_order(reading):
    """표에 넣을 공정 차례 — 먼저 나온 차례를 지킨다."""
    order = []
    for one in reading.get("lots") or []:
        for name in (one.get("values") or {}):
            if name not in order:
                order.append(name)
    for name in reading.get("specs") or {}:
        if name not in order:
            order.append(name)
    return order


def write_sheet(folder, reading, notes=(), path=None):
    """'7. 수율현황표 - 판독.xlsx' — 담당자가 쓰던 수율현황표와 같은 모양으로 쓴다.

    보고서 쪽(readers.yield_sheet)은 자리를 정해 두지 않고 표를 보고 찾으므로, 이 모양이면
    '보고서 작성' 이 그대로 읽는다.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, Side

    order = process_order(reading)
    book = Workbook()
    ws = book.active
    ws.title = "Sheet1"
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    mid = Alignment(horizontal="center", vertical="center", wrap_text=True)
    last = 2 + len(order)                       # A 연번 · B 라벨 · 공정들 · 비고
    ws.cell(row=2, column=1, value="연번")
    ws.cell(row=2, column=2, value="중요공정 별 수율 현황(%)")
    ws.cell(row=2, column=last + 1, value="비고")
    ws.cell(row=3, column=2, value="공정")
    ws.cell(row=5, column=2, value="기준")
    ws.cell(row=6, column=2, value="Lot No.")
    for i, name in enumerate(order):
        ws.cell(row=3, column=3 + i, value=name)
        ws.cell(row=5, column=3 + i, value=(reading.get("specs") or {}).get(name) or "확인 필요")
    row = 7
    for n, one in enumerate(reading.get("lots") or [], start=1):
        ws.cell(row=row, column=1, value=n)
        ws.cell(row=row, column=2, value=one["lot"])
        pages = one.get("page") or {}
        for i, name in enumerate(order):
            value = (one.get("values") or {}).get(name)
            ws.cell(row=row, column=3 + i,
                    value=value if value is not None else "확인 필요")
        seen = sorted({p for p in pages.values()})
        ws.cell(row=row, column=last + 1,
                value=("기록서 %s쪽 판독" % ", ".join(str(p) for p in seen)) if seen else "")
        row += 1
    for r in range(2, row):
        for c in range(1, last + 2):
            cell = ws.cell(row=r, column=c)
            cell.border = box
            cell.alignment = mid
    for c in range(1, last + 2):
        ws.cell(row=2, column=c).font = Font(bold=True)
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 14
    for i in range(len(order)):
        ws.column_dimensions[ws.cell(row=3, column=3 + i).column_letter].width = 13
    ws.cell(row=3, column=last + 1)
    ws.column_dimensions[ws.cell(row=2, column=last + 1).column_letter].width = 24
    # 아래에 판독 안내와 검산 결과를 적어 둔다 — 결재 전에 사람이 봐야 하는 자리다
    row += 1
    ws.cell(row=row, column=1, value="※ 이 표는 기록서 스캔을 판독해 만든 **초안**입니다. "
                                     "기록서와 대조해 확인한 뒤 쓰세요.")
    for note in notes or []:
        row += 1
        ws.cell(row=row, column=1, value="· " + note)
    target = path or os.path.join(folder, SHEET_NAME)
    book.save(target)
    return target


def save_reading(folder, reading, notes=(), extra=None):
    """판독 결과를 제품 폴더에 남긴다 — 다시 판독하지 않아도 값을 볼 수 있게."""
    body = dict(extra or {})
    body.update({"read_at": _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 "lots": reading.get("lots") or [], "specs": reading.get("specs") or {},
                 "notes": list(notes or [])})
    path = os.path.join(folder, READING_NAME)
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(body, handle, ensure_ascii=False, indent=2)
    return path


def load_reading(folder):
    path = os.path.join(folder, READING_NAME)
    if not os.path.isfile(path):
        return None
    try:
        with open(path, encoding="utf-8-sig") as handle:
            return json.load(handle)
    except (OSError, ValueError):
        return None


def save_log(folder, lines, result=None):
    """무엇을 했는지 제품 폴더에 남긴다 — 화면 알림은 사라져도 이 파일은 남는다."""
    if not folder or not os.path.isdir(folder):
        return ""
    got = dict(result or {})
    head = ["PQR 수율 판독 기록", _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ""]
    if got.get("ok"):
        head.append("결과: 판독했습니다 — %s Lot · %s" % (got.get("lots"), got.get("how") or ""))
        head.append("만든 파일: %s" % os.path.basename(str(got.get("sheet") or "")))
    elif got:
        head.append("결과: 판독하지 못했습니다")
        head.append("까닭: %s" % (got.get("why") or "알 수 없음"))
    for note in got.get("notes") or []:
        head.append("  · " + note)
    path = os.path.join(folder, LOG_NAME)
    try:
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("\n".join(head + [""] + ["진행 기록:"] + list(lines or [])) + "\n")
    except OSError:
        return ""
    return path


def _page_images(path, pages, work, log=None):
    """지정한 쪽만 그림으로 뽑는다 — 이름 끝에 쪽 번호를 남겨 판독기가 쪽을 알 수 있게."""
    from . import handwriting
    say = log or (lambda *a: None)
    made = []
    for page_no in pages:
        try:
            image = handwriting.render(path, DPI, page_no)
        except Exception as error:
            say("  %d쪽을 그림으로 뽑지 못했습니다: %s" % (page_no, error))
            continue
        out = os.path.join(work, "수율-%d쪽.png" % page_no)
        image.save(out, format="PNG", optimize=True)
        made.append(out)
    return made


def make_reading(input_dir, folder, code="", log=None, known_lots=(), allow_pc=False):
    """제품 폴더의 기록서 스캔에서 수율을 읽어 판독 파일과 수율현황표 초안을 만든다.

    돌려주는 값: {"ok", "how", "lots", "sheet", "reading", "notes"} 또는
    {"ok": False, "why", "need": "profile"|"reader"}
    """
    import shutil
    import tempfile
    from . import stability_read

    say = log or (lambda *a: None)
    folder = os.path.abspath(folder)
    mine = own_sheet(folder)
    if mine:
        say("담당자가 만든 수율현황표가 있습니다(%s) — 보고서는 그것을 씁니다. "
            "판독본은 참고용으로만 만듭니다." % os.path.basename(mine))
    data = profile_mod.load(input_dir)
    prof = profile_mod.for_product(data, code)
    if not prof:
        return {"ok": False, "need": "profile",
                "why": "이 제품에 기록서 서식 프로필이 배정되어 있지 않습니다 — '수율현황표 분석' "
                       "화면에서 수탁사·기록서 서식을 고르거나 새로 만들어 쪽수를 한 번 적어 주세요."}
    pages = profile_mod.page_list(prof)
    if not pages:
        return {"ok": False, "need": "profile",
                "why": "프로필 '%s' 에 쪽수가 없습니다 — 수율이 적힌 쪽 번호를 적어 주세요."
                       % prof.get("id", "")}
    paths = scans(folder)
    if not paths:
        return {"ok": False, "why": "7항(또는 6항)에 기록서 스캔 PDF 가 없습니다 — 먼저 올려 주세요."}
    kind, label = stability_read.how(folder)
    if kind not in ("api", "cli"):
        return {"ok": False, "need": "reader",
                "why": "이 PC 에서 쓸 수 있는 판독기가 없습니다 — Claude Code 를 깔거나 "
                       "ANTHROPIC_API_KEY 를 두세요. (이 PC 판독기는 표 판독에 쓰지 않습니다.)"}
    rules = prof.get("rules") or {}
    rule_text = str(rules.get("note") or "")
    say("7항 수율 판독: %s · 쪽 %s · %s"
        % (os.path.basename(paths[0]), ", ".join(str(p) for p in pages), label))
    if kind == "api":
        from . import vision_claude
        got = vision_claude.read_yield(paths[0], pages, rule_text, say)
    else:
        from . import claude_cli
        work = tempfile.mkdtemp(prefix="pqr-yield-")
        try:
            images = _page_images(paths[0], pages, work, say)
            if not images:
                return {"ok": False, "why": "지정한 쪽을 그림으로 뽑지 못했습니다 — 쪽 번호가 "
                                            "이 PDF 의 쪽수를 넘지 않는지 확인하세요."}
            got = claude_cli.read_yield(images, rule_text, folder, say)
        finally:
            shutil.rmtree(work, ignore_errors=True)
    for one in got.get("lots") or []:
        one["values"] = apply_names(one.get("values"), rules)
        one["production"] = apply_names(one.get("production"), rules)
        one["page"] = apply_names(one.get("page"), rules)
    got["specs"] = apply_names(got.get("specs"), rules)
    notes = check(got, known_lots)
    for note in notes:
        say("  확인 필요: " + note)
    sheet = write_sheet(folder, got, notes)
    reading = save_reading(folder, got, notes,
                           {"profile": prof.get("id", ""), "pages": pages,
                            "source": os.path.basename(paths[0]), "how": label})
    say("7항 수율 판독 끝: %d Lot → %s" % (len(got.get("lots") or []), os.path.basename(sheet)))
    return {"ok": True, "how": label, "lots": len(got.get("lots") or []),
            "sheet": sheet, "reading": reading, "notes": notes,
            "own_sheet": os.path.basename(mine) if mine else ""}
