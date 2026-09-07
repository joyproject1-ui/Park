# -*- coding: utf-8 -*-
"""'안정성 판독' 단추 — 13 폴더의 손글씨 시험일지를 읽어 판독 파일(json)을 만든다.

담당자 2026-09-07: "여기에 13. 안정성시험일지 판독.json 을 만드는 단추를 만들어서 누르고,
완성되면 보고서 작성 단추를 누르게 하면 안 될까?"

'보고서 작성' 은 몇십 초면 끝나야 하는데 손글씨 판독은 길게는 한 시간이 걸린다. 그래서 판독을
따로 떼어 단추 하나로 돌리고, 끝난 뒤 보고서를 만들게 한다. 만들어 둔 판독 파일에는
covers_all 을 적어 두므로 '보고서 작성' 은 시험일지를 다시 읽지 않는다.

차례: ① API 키(Claude) → ② 이 PC 의 Claude Code. 이 PC 판독기(RapidOCR)는 한 장에 1~3분이라
(20장에 71분, 담당자 2026-09-07: "시간이 너무 걸려서 PC 로 안 읽기로 한 거 아냐?") 저 혼자 시작하지
않는다 — 담당자가 "그래도 이 PC 로 읽겠다" 고 한 번 더 고를 때(allow_pc)만 돌린다. 고르지 않으면
Claude 대화에 올릴 판독 요청 묶음(zip)을 만들어 준다.

값을 지어내지 않는다 — 애매한 칸은 unsure 에 남아 보고서에서 노랑·주황으로 표시된다.
"""
import os

from . import collect as collect_mod
from .pdftext import is_scanned


def scans(folder):
    """13 폴더의 손글씨 시험일지(스캔 PDF) — 내용이 같은 파일은 하나만."""
    got = collect_mod.discover(folder).get("13", [])
    return collect_mod.same_files_once([p for p in got if p.lower().endswith(".pdf") and is_scanned(p)])


def specs_of(folder, log=None):
    """완제 성적서(9.2.4)에서 성분 이름과 규격 — 판독한 함량에 성분 이름을 붙이는 데 쓴다."""
    from .readers import coa as coa_reader
    out = {}
    for path in collect_mod.discover(folder).get("9.2.4", []):
        if not path.lower().endswith(".pdf"):
            continue
        try:
            rec = coa_reader.read_fp(path)
        except Exception:
            continue
        for a in rec.get("assays") or []:
            try:
                if a.get("part") and a["part"] not in out:
                    out[a["part"]] = (float(a["lo"]), float(a["hi"]))
            except (TypeError, ValueError, KeyError):
                continue
    return out


def how(folder=None):
    """지금 이 PC 에서 쓸 수 있는 판독 길 — ("api"|"cli"|"pc"|"", 설명). 빠른 길이 먼저."""
    try:
        from . import vision as vision_mod
        if vision_mod.available():
            return "api", "Claude (API 키)"
    except Exception:
        pass
    try:
        from . import claude_cli
        if claude_cli.available():
            return "cli", "이 PC 의 Claude Code"
    except Exception:
        pass
    try:
        from . import handwriting
        if handwriting.available():
            return "pc", "이 PC 의 판독기 (한 장에 몇 분)"
    except Exception:
        pass
    return "", "이 PC 에는 판독기가 없습니다"


MINUTES_PER_PAGE = 3          # 이 PC 판독기(RapidOCR) 실측: 20장에 71분


def make_reading(folder, product="", log=None, allow_pc=False):
    """13 폴더를 읽어 판독 파일을 만든다.

    돌려주는 값: {"ok", "how", "lots", "unsure", "path", "pages"} 또는
    {"ok": False, "need": "pc"|"", "pages", "minutes", "zip": [묶음 이름], "why"}
    """
    say = log or (lambda *a: None)
    from . import handwriting
    folder = os.path.abspath(folder)
    paths = scans(folder)
    if not paths:
        return {"ok": False, "why": "13 폴더에 손글씨 시험일지(스캔 PDF)가 없습니다 — 먼저 (r) 13. 안정성 시험에 올려 주세요.",
                "pages": 0}
    kind, label = how(folder)
    if kind == "pc" and not allow_pc:
        # 이 PC 판독기는 느리다 — 담당자가 한 번 더 고를 때만 돌린다. 그동안 쓸 묶음을 만들어 둔다.
        from . import handreq
        made = handreq.make_request(folder, paths, product or "", say)
        return {"ok": False, "need": "pc", "pages": len(paths),
                "minutes": len(paths) * MINUTES_PER_PAGE, "zip": [name for name, _ in made],
                "why": "이 PC 에는 Claude(API 키)도, Claude Code 도 없습니다. 이 PC 판독기로 읽으면 "
                       "시험일지 %d장에 %d분쯤 걸립니다." % (len(paths), len(paths) * MINUTES_PER_PAGE)}
    if not kind:
        from . import handreq
        made = handreq.make_request(folder, paths, product or "", say)
        return {"ok": False, "need": "", "pages": len(paths), "zip": [name for name, _ in made],
                "why": "이 PC 에서 쓸 수 있는 판독기가 없습니다 — 이 PC 에 Claude Code 를 깔거나, "
                       "만들어 둔 판독 요청 묶음을 Claude 대화에 올려 주세요."}
    say("13항 안정성 판독: 시험일지 %d장 — %s" % (len(paths), label))
    specs = specs_of(folder, log) or None
    if kind == "api":
        from . import vision_claude
        logs = vision_claude.read_logs(paths, specs, say)
    elif kind == "cli":
        from . import claude_cli
        logs = claude_cli.read_logs(paths, specs, say, folder)
    else:
        logs = handwriting.read_folder(paths, specs, say)
    if not logs:
        return {"ok": False, "pages": len(paths), "how": label,
                "why": "시험일지에서 읽어 낸 것이 없습니다 — 원본이 너무 흐리거나 표가 다른 꼴일 수 있습니다."}
    # 이미 있던 판독 파일(담당자가 고쳐 둔 값)이 있으면 그 값이 우선이다
    old = handwriting.load_cache(folder) if hasattr(handwriting, "load_cache") else []
    merged = list(old or [])
    handwriting.merge_logs(merged, logs, say)
    merged.sort(key=lambda r: (r.get("year") or "", r.get("lot") or ""))
    path = handwriting.save_cache(folder, merged, say, covers_all=True)
    unsure = sum(len(p.get("unsure") or []) for one in merged for p in (one.get("points") or []))
    say("13항 안정성 판독 끝: %d Lot · 애매한 칸 %d개 → %s" % (len(merged), unsure, os.path.basename(str(path or ""))))
    return {"ok": True, "how": label, "lots": len(merged), "unsure": unsure,
            "pages": len(paths), "path": path or ""}
