# -*- coding: utf-8 -*-
"""이 PC 에 깔린 Claude Code 로 13항 손글씨 시험일지를 읽는다 — API 키 없이, 구독 그대로.

담당자 2026-09-07: "매번 명령하라는 거야? 이것을 자동화하면 안 돼?"
Claude Code 는 `claude -p` 로 부를 수 있어서, 프로그램이 대신 물어보고 답만 받아 온다.
담당자는 '보고서 작성' 만 누르면 된다 — 창을 열거나 물어볼 일이 없다.

읽기(Read)만 허용하고 명령 실행 도구는 빼고 부른다(--restricted). 시험일지 파일은
이 PC 안에서만 열리고, 결과는 PC 판독·API 판독과 똑같은 꼴이라 뒤 과정이 모두 같다.
"""
import json
import os
import re
import shutil
import subprocess

TIMEOUT = 60 * 30                 # 넉넉하게 — 스무 장이면 몇 분
CHUNK = 6                         # 한 번에 물어볼 시험일지 수 (답이 너무 길어지지 않게)

PROMPT = """다음 안정성 시험일지(손글씨 스캔 PDF)를 모두 읽고 JSON 만 출력하세요.

읽을 파일 (%(count)d개):
%(files)s

각 파일의 모든 쪽을 읽습니다. 한 쪽이 한 제조번호(Lot)입니다.

출력 형식 — 이 꼴의 JSON 객체 하나만, 설명이나 코드 표시 없이:
{"logs": [
  {
    "lot": "OEX101",
    "year": "2024",
    "kind": "장기",
    "market": "내수",
    "pack": "5g tube/갑",
    "store": "25±2°C,\\n60±5%%RH",
    "mfg": "2024.01.10",
    "expiry": "2027.01.09",
    "source": "읽은 파일 이름.pdf",
    "points": [
      {"period": "12M", "done": "2025.03.10", "assays": {"오플록사신": 101.2}, "unsure": []}
    ]
  }
]}

규칙
· lot: 제조번호. year: 제조 연도 네 자리.
· kind: 표의 시험구분대로 "장기" 또는 "시판후". 없으면 시점이 3M·6M·9M 이면 "장기",
  12M·24M·36M 만 있으면 "시판후".
· market: 제품명이나 파일 이름에 '수출' 이 있으면 "수출", 아니면 "내수".
· period: Initial(초기)·3M·6M·9M·12M·18M·24M·36M. 시험하지 않은(사선·빈) 시점은 넣지 않습니다.
· done: 그 시점의 결재(확인자·팀장) 일자. 없으면 시험일자. 형식 YYYY.MM.DD.
· assays: 성분 이름과 숫자. 성분이 둘이면 둘 다. 성분 이름을 모르면 "함량".
%(parts)s
· unsure: 읽기 애매한 것의 이름을 넣습니다 — 완료 일자가 애매하면 "done", 함량이 애매하면 그 성분 이름.
· 값을 지어내지 마세요. 읽지 못한 시점은 빼거나 unsure 에 적습니다.
"""


HINT_FILE = "CLAUDE 경로.txt"          # 프로그램 폴더에 두면 그 경로를 쓴다


def _program_dir():
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _hint():
    """담당자가 알려 준 claude 경로 — 환경변수 PQR_CLAUDE_EXE 또는 프로그램 폴더의 'CLAUDE 경로.txt'."""
    got = (os.environ.get("PQR_CLAUDE_EXE") or "").strip().strip('"')
    if got:
        return got
    path = os.path.join(_program_dir(), HINT_FILE)
    try:
        with open(path, encoding="utf-8-sig") as handle:
            for line in handle:
                line = line.strip().strip('"')
                if line and not line.startswith("#"):
                    return line
    except OSError:
        pass
    return ""


def places():
    """claude 를 찾아볼 자리 — 어디를 뒤졌는지 진단 화면에 그대로 보여 준다.

    Windows 는 설치 방법마다 자리가 다르고(네이티브 설치·npm 전역·Claude 앱), 새로 깔았다면
    이미 떠 있는 창의 PATH 에는 아직 없다 (담당자 2026-09-07: "이 PC 로 Claude Code 로 작업하고
    있는데 없다는 게 무슨 말이냐" — 프로그램은 PATH 만 보고 있었다).
    """
    out = []
    for base in (os.path.expanduser("~/.local/bin"),
                 os.path.expandvars(r"%USERPROFILE%\.local\bin"),
                 os.path.expandvars(r"%APPDATA%\npm"),
                 os.path.expandvars(r"%LOCALAPPDATA%\Programs\claude"),
                 os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WindowsApps"),
                 os.path.expanduser("~/.claude/local"),
                 os.path.expanduser("~/.bun/bin"), "/usr/local/bin", "/opt/homebrew/bin"):
        if not base or "%" in base:
            continue
        for name in ("claude.exe", "claude.cmd", "claude.bat", "claude"):
            out.append(os.path.join(base, name))
    return out


def _exe():
    hint = _hint()
    if hint and os.path.isfile(hint):
        return hint
    for name in ("claude", "claude.cmd", "claude.exe", "claude.bat"):
        got = shutil.which(name)
        if got:
            return got
    for guess in places():
        if os.path.isfile(guess):
            return guess
    return ""


def available():
    """이 PC 에 Claude Code 가 깔려 있는가."""
    return bool(_exe())


def _ask(exe, prompt, folder, log=None):
    """claude -p 로 한 번 물어보고 답 글을 돌려준다.

    물음은 표준 입력으로 넣는다 — Windows 에서 claude 는 .cmd 라 명령줄로 넘기면 cmd.exe 가
    큰따옴표·%·& 를 제 나름대로 해석해 물음이 깨진다.
    """
    cmd = [exe, "-p", "--output-format", "json", "--restricted",
           "--allowedTools", "Read", "Glob", "--add-dir", os.path.abspath(folder)]
    got = subprocess.run(cmd, cwd=folder, input=prompt.encode("utf-8"),
                         capture_output=True, timeout=TIMEOUT)
    out = (got.stdout or b"").decode("utf-8", "replace")
    err = (got.stderr or b"").decode("utf-8", "replace").strip()
    if got.returncode != 0:
        raise RuntimeError("claude -p 가 %d 로 끝났습니다 — %s" % (got.returncode, err[:300] or out[:300]))
    try:                                        # --output-format json 은 겉봉투를 씌워 준다
        envelope = json.loads(out)
    except ValueError:
        return out
    if isinstance(envelope, dict):
        if envelope.get("is_error") or envelope.get("subtype") not in (None, "success"):
            raise RuntimeError(str(envelope.get("result") or envelope.get("subtype") or envelope)[:300])
        return str(envelope.get("result") or "")
    return out


def _json_object(text):
    """답 글에서 JSON 객체만 꺼낸다 — 코드 표시(```)나 앞뒤 설명이 붙어 있어도."""
    body = re.sub(r"^```(?:json)?|```$", "", (text or "").strip(), flags=re.M).strip()
    start = body.find("{")
    if start < 0:
        raise ValueError("답에 JSON 이 없습니다: %s" % (text or "")[:200])
    depth, in_str, esc = 0, False, False
    for i in range(start, len(body)):
        ch = body[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return json.loads(body[start:i + 1])
    raise ValueError("JSON 이 끝나지 않았습니다: %s" % body[start:start + 200])


def _clean(logs, paths, specs=None):
    """받은 기록을 판독 파일 꼴로 다듬는다 — 모르는 칸은 버리고, 숫자는 숫자로."""
    from . import handwriting, vision_claude
    names = {os.path.basename(p) for p in paths}
    out = []
    for one in logs or []:
        if not isinstance(one, dict):
            continue
        lot = re.sub(r"\s+", "", str(one.get("lot") or "")).upper()
        if not lot:
            continue
        points = []
        for p in one.get("points") or []:
            if not isinstance(p, dict):
                continue
            period = vision_claude._period(p.get("period") or p.get("label"))
            if not period:
                continue
            assays = {}
            for name, value in (p.get("assays") or {}).items():
                text = re.sub(r"[^\d.]", "", str(value))
                if not text:
                    continue
                try:
                    assays[vision_claude._part(name, specs)] = float(text)
                except ValueError:
                    continue
            done = vision_claude._date(p.get("done"))
            unsure = [str(u) for u in (p.get("unsure") or []) if str(u).strip()]
            if not done and "done" not in unsure:
                unsure.append("done")
            if not assays and not done:
                continue
            points.append({"period": period, "done": done, "assays": assays,
                           "unsure": sorted(set(unsure))})
        if not points:
            continue
        points.sort(key=lambda p: handwriting.period_order(p["period"]))
        source = str(one.get("source") or "")
        if source not in names:                      # 파일 이름을 다르게 적어 왔으면 하나만 읽은 경우로 본다
            source = os.path.basename(paths[0]) if len(paths) == 1 else source
        market = one.get("market") if one.get("market") in ("내수", "수출") else ""
        kind = one.get("kind") if one.get("kind") in ("장기", "시판후") else "장기"
        out.append({"lot": lot, "year": str(one.get("year") or "")[:4],
                    "pack": str(one.get("pack") or "").strip(),
                    "store": str(one.get("store") or "").strip(),
                    "kind": kind, "kind_sure": bool(one.get("kind")),
                    "market": market, "market_hint": bool(market),
                    "mfg": vision_claude._date(one.get("mfg")),
                    "expiry": vision_claude._date(one.get("expiry")),
                    "why": "", "points": points, "notes": [], "source": source})
    return out


def read_logs(paths, specs=None, log=None, folder=None):
    """시험일지 paths 를 이 PC 의 Claude Code 로 읽어 판독 파일 꼴의 목록을 돌려준다."""
    from . import handwriting
    say = log or (lambda *a: None)
    exe = _exe()
    if not exe:
        raise RuntimeError("이 PC 에 Claude Code 가 없습니다")
    where = folder or os.path.dirname(os.path.abspath(paths[0]))
    parts = ""
    if specs:
        parts = "· 이 제품의 성분 이름은 %s 입니다 — 성분 이름을 이 가운데 하나로 맞춰 주세요.\n" % ", ".join(specs)
    merged = []
    for i in range(0, len(paths), CHUNK):
        group = paths[i:i + CHUNK]
        say("    Claude Code 판독 %d~%d/%d장" % (i + 1, i + len(group), len(paths)))
        prompt = PROMPT % {"count": len(group), "parts": parts,
                           "files": "\n".join(os.path.abspath(p) for p in group)}
        text = _ask(exe, prompt, where, log)
        got = _clean(_json_object(text).get("logs"), group, specs)
        for one in got:
            say("      %s: %s·%s 시점 %d개 (애매 %d칸)"
                % (one["lot"], one["kind"], one["market"] or "구분 없음", len(one["points"]),
                   sum(len(p["unsure"]) for p in one["points"])))
        handwriting.merge_logs(merged, got, log)
    merged.sort(key=lambda r: (r.get("year") or "", r.get("lot") or ""))
    return merged
