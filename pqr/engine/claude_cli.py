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
CHUNK = 2                         # 한 번에 물어볼 시험일지 수 (답이 너무 길어지지 않게)
WORKERS = 4                       # 동시에 띄울 Claude Code 수 — 한 번에 한 묶음씩 물으면 너무 느리다
                                  # (담당자 2026-09-07: "PQR 하나 만드는 데 너무 오래 걸리네")

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


def command(exe=None):
    """실제로 실행할 명령 — Windows 의 claude.CMD 는 cmd.exe 를 한 겹 더 거친다.

    그 겹을 지나면 표준 입력이 claude 에 닿지 않아 물음이 빈 채로 돌았다(담당자 PC 2026-09-07:
    input_tokens 0, duration_api_ms 0). 옆에 있는 node 와 cli.js 를 바로 부르면 그 겹이 없어진다.
    """
    exe = exe or _exe()
    if not exe:
        return []
    if exe.lower().endswith((".cmd", ".bat")):
        base = os.path.dirname(exe)
        for js in (os.path.join(base, "node_modules", "@anthropic-ai", "claude-code", "cli.js"),
                   os.path.join(base, "..", "lib", "node_modules", "@anthropic-ai", "claude-code", "cli.js")):
            js = os.path.normpath(js)
            node = shutil.which("node") or shutil.which("node.exe")
            if node and os.path.isfile(js):
                return [node, js]
    return [exe]


def ask_once(prompt, folder=None, timeout=180, extra=()):
    """claude 에 한 번 물어보고 (성공 여부, 답 또는 까닭) 을 돌려준다.

    Windows 의 `claude.CMD` 를 그냥 부르면 cmd.exe 를 한 겹 더 지나면서 표준 입력이 끊겨,
    물음이 빈 채로 돌아 아무것도 하지 않고 끝났다 (담당자 PC 2026-09-07: `input_tokens 0 ·
    duration_api_ms 0`). 그래서

      ① node 와 cli.js 를 바로 부를 수 있으면 그 길로, 물음은 **명령 인자**로 넘긴다
         (셸을 거치지 않으므로 따옴표·줄바꿈·%가 그대로 간다),
      ② 그 길이 없어 .cmd 를 거쳐야 하면 물음을 **임시 파일**로 만들어 표준 입력에 물려 준다
         (파이프는 그 겹을 지나며 끊긴다).
    """
    import tempfile
    cmd = command()
    if not cmd:
        return False, "이 PC 에 Claude Code 가 없습니다"
    through_cmd = cmd[0].lower().endswith((".cmd", ".bat"))
    args = list(cmd) + ["-p"] + list(extra) + ["--output-format", "json"]
    if not through_cmd:
        args.append(prompt)
    work = tempfile.mkdtemp(prefix="pqr-claude-")
    try:
        stdin = None
        if through_cmd:
            path = os.path.join(work, "ask.txt")
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(prompt)
            stdin = open(path, "rb")
        try:
            run = subprocess.run(args, cwd=folder or work, stdin=stdin,
                                 stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=timeout)
        except (OSError, subprocess.TimeoutExpired) as error:
            return False, str(error)
        finally:
            if stdin is not None:
                stdin.close()
    finally:
        shutil.rmtree(work, ignore_errors=True)
    out = (run.stdout or b"").decode("utf-8", "replace")
    try:                                       # --output-format json 은 겉봉투를 씌워 준다
        envelope = json.loads(out)
    except ValueError:
        envelope = None
    if isinstance(envelope, dict):
        text = str(envelope.get("result") or "")
        bad = envelope.get("is_error") or envelope.get("subtype") not in (None, "success")
        if run.returncode != 0 or bad:
            return False, (text or str(envelope.get("subtype") or ""))[:400] or "까닭을 알 수 없습니다"
        if not text.strip():
            # 답이 비었다 = 물음이 닿지 않았다는 뜻이다. 무엇으로 불렀는지 함께 알린다.
            return False, ("물음이 claude 에 닿지 않았습니다(답이 비어 있음) — %s"
                           % os.path.basename(cmd[0]))
        return True, text
    if run.returncode != 0:
        return False, " ".join(out.split())[:400] or "까닭을 알 수 없습니다"
    return True, out


def _ask(exe, prompt, folder, log=None, paths=()):
    r"""읽기(Read)만 허용해 한 번 물어본다 — 시험일지 판독용.

    읽을 파일이 있는 폴더를 모두 --add-dir 로 준다. 담당자 PC 2026-09-07: 시험일지가
    압축 안에 있어 임시 폴더(%TEMP%\pqr-engine-…)에 풀렸는데 제품 폴더만 열어 주는 바람에
    Claude Code 가 "The four PDFs are not readable in this session" 이라며 빈손으로 돌아왔다.
    """
    dirs, seen = [], set()
    for one in [folder] + [os.path.dirname(os.path.abspath(p)) for p in paths]:
        if not one:
            continue
        full = os.path.abspath(one)
        key = os.path.normcase(full)
        if key in seen:
            continue
        seen.add(key)
        dirs.append(full)
    extra = ["--restricted", "--allowedTools", "Read", "Glob"]
    for one in dirs:
        extra += ["--add-dir", one]
    ok, text = ask_once(prompt, folder, TIMEOUT, extra=extra)
    if not ok:
        raise RuntimeError("claude -p 로 읽지 못했습니다 — %s" % text)
    return text


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


def read_logs(paths, specs=None, log=None, folder=None, workers=WORKERS, chunk=CHUNK):
    """시험일지 paths 를 이 PC 의 Claude Code 로 읽어 판독 파일 꼴의 목록을 돌려준다.

    묶음을 **여러 개 동시에** 물어본다 — 한 묶음씩 차례로 물으면 넉 장에 5분이 넘었다
    (담당자 2026-09-07: "PQR 하나 만드는 데 너무 오래 걸리네"). 한 묶음이 실패해도 나머지는
    살린다 — 다 실패했을 때만 멈춘다.
    """
    from concurrent.futures import ThreadPoolExecutor
    from . import handwriting
    say = log or (lambda *a: None)
    exe = _exe()
    if not exe:
        raise RuntimeError("이 PC 에 Claude Code 가 없습니다")
    where = folder or os.path.dirname(os.path.abspath(paths[0]))
    parts = ""
    if specs:
        parts = "· 이 제품의 성분 이름은 %s 입니다 — 성분 이름을 이 가운데 하나로 맞춰 주세요.\n" % ", ".join(specs)
    groups = [paths[i:i + max(1, chunk)] for i in range(0, len(paths), max(1, chunk))]
    at_once = max(1, min(workers, len(groups)))
    say("    Claude Code 판독: %d장 (%d장씩 %d묶음을 동시에)" % (len(paths), max(1, chunk), at_once))
    done = [0]

    def read_group(job):
        gi, group = job
        prompt = PROMPT % {"count": len(group), "parts": parts,
                           "files": "\n".join(os.path.abspath(p) for p in group)}
        try:
            got = _clean(_json_object(_ask(exe, prompt, where, None, group)).get("logs"), group, specs)
        except Exception as error:
            return gi, [], error
        return gi, got, None

    results, trouble = {}, []
    with ThreadPoolExecutor(max_workers=at_once) as pool:
        for gi, got, error in pool.map(read_group, list(enumerate(groups))):
            done[0] += len(groups[gi])
            say("    Claude Code 판독 %d/%d장" % (done[0], len(paths)))
            if error is not None:
                trouble.append(error)
                say("      읽지 못한 묶음: %s — %s"
                    % (", ".join(os.path.basename(p) for p in groups[gi]), error))
                continue
            results[gi] = got
            for one in got:
                say("      %s: %s·%s 시점 %d개 (애매 %d칸)"
                    % (one["lot"], one["kind"], one["market"] or "구분 없음", len(one["points"]),
                       sum(len(p["unsure"]) for p in one["points"])))
    if trouble and not results:
        raise RuntimeError("claude -p 로 읽지 못했습니다 — %s" % trouble[0])
    merged = []
    for gi in sorted(results):                        # 물어본 차례대로 합친다 — 결과가 늘 같게
        handwriting.merge_logs(merged, results[gi], log)
    merged.sort(key=lambda r: (r.get("year") or "", r.get("lot") or ""))
    return merged

CHANGE_PROMPT = """다음 변경요청서(스캔 PDF)를 읽고 JSON 만 출력하세요.

읽을 파일:
%(file)s

규칙
· 보이는 대로만 적습니다 — 안 보이면 빈 값으로 두고 지어내지 않습니다.
· "actions" 는 '변경 실행 계획' 표의 부서별 조치사항입니다. 부서 이름과 할 일을 짝으로 적습니다.
· "products" 는 '관련 제품' 에 적힌 제품 이름입니다.

{"doc_no": "CC-240723-08",
 "title": "변경명 한 줄",
 "description": "변경 내용 (여러 줄이면 줄바꿈으로)",
 "reason": "변경 사유",
 "products": "관련 제품",
 "approved": "2024.08.01",
 "actions": [["부서", "조치사항"], ["부서", "조치사항"]]}
"""


def read_change(path, folder=None, log=None):
    """스캔 변경요청서를 이 PC 의 Claude Code 로 읽는다 — readers.change.read_change 와 같은 꼴.

    담당자 2026-09-07: "변경요청서 읽으면 돼 PDF 라서 못 읽는 거야?" — PDF 라서가 아니라
    글자가 없는 스캔본이라 글자 판독기로는 한 자도 안 나온다. 시험일지와 같은 길로 읽는다.
    """
    say = log or (lambda *a: None)
    if not _exe():
        raise RuntimeError("이 PC 에 Claude Code 가 없습니다")
    where = folder or os.path.dirname(os.path.abspath(path))
    got = _json_object(_ask(_exe(), CHANGE_PROMPT % {"file": os.path.abspath(path)},
                            where, None, [path]))
    acts = []
    for one in got.get("actions") or []:
        if isinstance(one, (list, tuple)) and len(one) >= 2:
            acts.append((str(one[0] or "").strip(), str(one[1] or "").strip()))
        elif isinstance(one, dict):
            acts.append((str(one.get("team") or "").strip(), str(one.get("action") or "").strip()))
    out = {"doc_no": str(got.get("doc_no") or "").strip(),
           "title": str(got.get("title") or "").strip(),
           "description": str(got.get("description") or "").strip(),
           "reason": str(got.get("reason") or "").strip(),
           "products": str(got.get("products") or "").strip(),
           "approved": str(got.get("approved") or "").strip(),
           "attachments": "", "target_date": "", "all_dates": [],
           "actions": [(t, a) for t, a in acts if a]}
    say("    [12] %s — Claude Code 로 읽음: %s (조치 %d건)"
        % (os.path.basename(path), out["title"] or "제목 못 읽음", len(out["actions"])))
    return out
