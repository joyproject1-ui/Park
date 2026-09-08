"""pqr 명령줄 도구.

    python -m pqr demo                       예시 입력 파일 만들기
    python -m pqr check  --in <입력폴더>       파일 인식 · 열 매핑 점검 (계산 전)
    python -m pqr build  --in <입력폴더>       집계 · 판정 → 대시보드 데이터 + 보고서
    python -m pqr launch                     폴더 준비 + 브라우저 열기 + 실행 (가장 간단)
    python -m pqr serve  --in <입력폴더>       대시보드만 띄우기
    python -m pqr narrate --data <pqr.json>   서술 문안 초안 (Claude API)
    python -m pqr report  --data <pqr.json>   보고서만 다시 생성
"""

import argparse
import io
import json
import os
import sys

from . import build as build_module
from . import report as report_module
from . import schema
from .sample import write_samples

DEFAULT_OUT = "out"


def _print(*parts):
    print(*parts)


def _load_data(path):
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


def _save_json(path, payload):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)


def _write_dashboard(data, out_dir):
    """대시보드가 읽는 형태로 내보냅니다.

    data.js 는 file:// 로 열어도 동작하도록 전역 변수에 담고,
    data.json 은 사내 웹서버에 올려 쓸 때를 위한 것입니다.
    """
    os.makedirs(out_dir, exist_ok=True)
    payload = {key: data[key] for key in
               ("generated_at", "today", "period", "stages", "items", "products",
                "trend", "leadtime", "sources", "narrative")}
    payload["issue_count"] = len([i for i in data.get("issues", []) if i["level"] == "error"])
    json_path = os.path.join(out_dir, "data.json")
    with open(json_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    js_path = os.path.join(out_dir, "data.js")
    with open(js_path, "w", encoding="utf-8") as handle:
        handle.write("window.PQR_DATA = ")
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write(";\n")
    return [json_path, js_path]


# --------------------------------------------------------------------------


def cmd_demo(args):
    files = write_samples(args.out, layout=args.layout)
    _print("예시 입력 파일 %d개를 만들었습니다: %s" % (len(files), args.out))
    for path in files:
        _print("  -", os.path.relpath(path, args.out))
    _print("")
    _print("다음 단계:  python -m pqr build --in %s" % args.out)
    return 0


# (데이터셋, 설명, 표 머리글, 권장 파일명, 내용)
UPLOAD_GUIDE = [
    ("batches", "시험성적서 · 공정관리", "시험성적서", "시험성적서.xlsx", "배치별 완제품·공정관리 시험결과"),
    ("deviations", "일탈 · OOS/OOT · CAPA", "일탈·CAPA", "일탈대장.xlsx", "일탈, 규격부적합, CAPA 이력"),
    ("changes", "변경 · 불만 · 회수", "변경·불만", "변경불만대장.xlsx", "변경관리, 허가변경, 불만·회수·반품"),
    ("stability", "안정성 모니터링", "안정성", "안정성.xlsx", "장기·가속 안정성 시험결과"),
]


def _width(text):
    """한글은 두 칸으로 세어 표를 맞춥니다."""
    import unicodedata
    return sum(2 if unicodedata.east_asian_width(char) in "WF" else 1 for char in str(text))


def _pad(text, size):
    return str(text) + " " * max(0, size - _width(text))
COMMON_GUIDE = [
    ("products", "제품 마스터", "products_제품마스터.csv", "제품코드·담당자·평가기간·마감일"),
    ("qualification", "설비 적격성 · 위수탁", "qualification_적격성.csv", "HVAC·용수·압축가스·위수탁 협약"),
    ("stagelog", "단계 진행 이력 (선택)", "stagelog_단계이력.csv", "리드타임 분석용"),
]

FOLDER_README = """이 폴더는 [{code}] {name} 의 제품품질평가(PQR) 자료를 올리는 곳입니다.

■ 가장 간단한 방법 — 파일 이름을 보고서 항 번호로 시작하세요.
  예)  3. 허가증.pdf · 6. 제조내역.pdf · 7. 수율현황표.xlsx
       8.1.1 원료 공급업체 List.xlsx · 9.2.1 조제 완료 후 (폴더도 됩니다)
       10.3, 10.4, 10.5 제조지원 설비.xlsx · 12. 변경관리 CC-XXXXXX.pdf · 13. 안정성 시험
  번호가 붙어 있으면 PDF·스캔이라도 '자료 수집' 화면에 녹색으로 표시됩니다.
  번호가 붙은 폴더는 안에 파일이 있어야 인정됩니다.

■ 집계·판정까지 하려면 아래 표 파일(xlsx·csv)도 함께 올리세요.
  파일 이름에 아래 낱말이 들어 있으면 자동으로 읽습니다.

{files}
제품코드 열은 넣지 않아도 됩니다 — 이 폴더 이름에서 자동으로 채워집니다.
열 이름은 한국어·영어 모두 인식합니다 (예: 배치번호 / 제조번호 / Lot).

■ '보고서 작성' 은 16. 전년도 PQR(워드 — 압축이어도, 자료 전체를 zip 하나로 올려도 됩니다)을
  바탕으로 채워 EDMS 결재본 양식(E-HLF-32)으로 만듭니다. 서식은 프로그램에 들어 있어 따로 둘
  것이 없습니다. 서식이 개정되면 새 서식(.docx)을 '공통' 폴더에 두면 그것을 씁니다.

올린 뒤 담당자가 아래를 실행하면 대시보드와 보고서 초안이 만들어집니다.

    python -m pqr check --in <입력폴더>     자료가 제대로 인식되는지 확인
    python -m pqr build --in <입력폴더>     집계 · 판정 · 보고서 생성
"""


def cmd_init(args):
    """담당자가 자료를 올릴 제품 폴더를 만듭니다."""
    codes = []
    if args.master:
        from .tabular import read_table
        rows = read_table(args.master)
        normalized, _ = schema.normalize(rows, "products", os.path.basename(args.master))
        codes = [(row["product_code"], row.get("product_name", "")) for row in normalized]
    for value in args.product or []:
        name = ""
        if "=" in value:
            value, name = value.split("=", 1)
        codes.append((value.strip().upper(), name.strip()))
    if not codes:
        _print("만들 제품이 없습니다. --master 로 제품 마스터 파일을 주거나 "
               "--product HP-101=제품명 형태로 지정하세요.")
        return 2

    os.makedirs(args.out, exist_ok=True)
    guide_lines = "\n".join(
        "  - %s %s (%s)" % (_pad(filename, 24), label, hint)
        for _, label, _, filename, hint in UPLOAD_GUIDE)
    made = 0
    for code, name in codes:
        folder = os.path.join(args.out, ("%s %s" % (code, name)).strip())
        os.makedirs(folder, exist_ok=True)
        readme = os.path.join(folder, "_읽어보기.txt")
        if not os.path.exists(readme):
            with open(readme, "w", encoding="utf-8") as handle:
                handle.write(FOLDER_README.format(code=code, name=name or code,
                                                  files=guide_lines + "\n"))
        made += 1
    common = os.path.join(args.out, "공통")
    os.makedirs(common, exist_ok=True)
    _print("제품 폴더 %d개를 만들었습니다: %s" % (made, args.out))
    _print("  공통 자료(설비 적격성 등)는 '공통' 폴더에 넣으세요.")
    _print("")
    _print("각 폴더에 올릴 파일:")
    _print(guide_lines)
    return 0


def cmd_check(args):
    if not os.path.isdir(args.input):
        _print("입력 폴더를 찾을 수 없습니다: %s" % args.input)
        return 2

    tree_mode = build_module.has_product_folders(args.input)
    _print("입력 폴더: %s  (%s)"
           % (args.input, "제품 폴더 방식" if tree_mode else "단일 폴더 방식"))
    datasets, sources, issues, presence = build_module.load(input_dir=args.input)

    submitted = presence["products"]
    common = presence["common"]
    _print("")
    _print("제출 현황  (O 제출 · - 없음 · C 공통 자료로 대체)")
    _print("")
    _print("  " + _pad("제품", 16) + "".join(_pad(short, 14) for _, _, short, _, _ in UPLOAD_GUIDE))
    codes = sorted(submitted) or ["(제품 폴더 없음)"]
    incomplete = 0
    for code in codes:
        marks = []
        for dataset, _, _, _, _ in UPLOAD_GUIDE:
            if dataset in submitted.get(code, set()):
                marks.append("O")
            elif dataset in common:
                marks.append("C")
            else:
                marks.append("-")
        if "-" in marks:
            incomplete += 1
        _print("  " + _pad(code, 16) + "".join(_pad(mark, 14) for mark in marks))

    _print("")
    _print("공통 자료")
    for dataset, label, filename, _ in COMMON_GUIDE:
        state = "O" if datasets.get(dataset) else "-"
        _print("  %s %s %s" % (state, _pad(label, 24), filename))

    if presence["unknown"]:
        _print("")
        _print("종류를 알 수 없어 건너뛴 파일 — 파일 이름에 시험성적서 · 일탈 · 변경 · 안정성 ·"
               " 적격성 같은 낱말을 넣어 주세요:")
        for path in presence["unknown"]:
            _print("  -", os.path.relpath(path, args.input))

    _print("")
    for dataset, entries in sorted(sources.items()):
        for entry in entries:
            _print("  %s %s %s %4d행 적재 · %d행 건너뜀"
                   % (_pad(dataset, 14), _pad(entry["product"], 12), _pad(entry["file"], 24),
                      entry["rows"], entry["skipped"]))

    errors = [issue for issue in issues if issue["level"] == "error"]
    warnings = [issue for issue in issues if issue["level"] == "warning"]
    _print("")
    _print("오류 %d건 · 경고 %d건 · 자료 미제출 제품 %d품목"
           % (len(errors), len(warnings), incomplete))
    for issue in (errors + warnings)[:20]:
        _print("  [%s] %s %s행 %s — %s"
               % (issue["level"], issue["source"], issue["row"], issue["field"], issue["message"]))
    if len(errors) + len(warnings) > 20:
        _print("  ... 외 %d건" % (len(errors) + len(warnings) - 20))
    return 1 if errors else 0


def cmd_build(args):
    if not os.path.isdir(args.input):
        _print("입력 폴더를 찾을 수 없습니다: %s" % args.input)
        return 2
    period = None
    if args.period_from and args.period_to:
        period = (args.period_from, args.period_to)
    data = build_module.build(input_dir=args.input, today=args.today, period=period)

    if args.narrate:
        from . import narrate as narrate_module
        try:
            narrate_module.narrate(data, model=args.model, log=_print)
        except RuntimeError as error:
            _print("서술 문안을 건너뜁니다: %s" % error)

    data_path = os.path.join(args.out, "pqr.json")
    _save_json(data_path, data)
    dashboard_files = _write_dashboard(data, os.path.join(args.out, "dashboard"))
    reports = report_module.write_reports(data, os.path.join(args.out, "reports"))

    errors = [issue for issue in data["issues"] if issue["level"] == "error"]
    _print("제품 %d품목 · 오류 %d건" % (len(data["products"]), len(errors)))
    _print("")
    _print("  데이터셋   %s" % data_path)
    for path in dashboard_files:
        _print("  대시보드   %s" % path)
    _print("  보고서     %s 외 %d건" % (reports[0], len(reports) - 1))
    _print("")
    _print("대시보드에서 보려면 data.js 를 docs/pqr/ 에 두고 index.html 을 여세요:")
    _print("  cp %s docs/pqr/data.js" % dashboard_files[1])
    return 1 if errors else 0


DEFAULT_INPUT = "입력폴더"


def _free_port(host, preferred, attempts=10):
    """원하는 포트가 이미 쓰이면 다음 번호로 넘어갑니다."""
    import socket
    for offset in range(attempts):
        port = preferred + offset
        probe = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            probe.bind((host, port))
            return port
        except OSError:
            continue
        finally:
            probe.close()
    return preferred


# 더블클릭 실행에서 입력 폴더를 이 이름들로 찾습니다 (앞이 우선).
INPUT_FOLDER_NAMES = ("PQR_입력폴더", "입력폴더", "PQR입력폴더")


def _looks_like_input_dir(path):
    """제품 폴더('QC1-1022 …' 처럼 코드로 시작)나 제품 마스터가 있으면 입력 폴더로 봅니다."""
    try:
        names = os.listdir(path)
    except OSError:
        return False
    if any("제품마스터" in name or "products" in name.lower() for name in names):
        return True
    from . import build as build_module
    return any(os.path.isdir(os.path.join(path, name))
               and build_module._folder_product_code(name) for name in names)


def find_input_dir():
    """담당자가 어디에 뒀든 입력 폴더를 찾아봅니다.

    cmd 에 경로를 입력하는 일이 없도록, 더블클릭 실행이 프로그램 폴더 주변과
    바탕화면·문서·다운로드에서 이름이 맞는 폴더를 차례로 살핍니다.
    """
    program_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    home = os.path.expanduser("~")
    bases = [os.getcwd(), program_dir, os.path.dirname(program_dir),
             os.path.join(home, "Desktop"), os.path.join(home, "바탕 화면"),
             os.path.join(home, "Documents"), os.path.join(home, "문서"),
             os.path.join(home, "Downloads"), os.path.join(home, "다운로드")]
    seen = set()
    for base in bases:
        base = os.path.abspath(base)
        if base in seen or not os.path.isdir(base):
            continue
        seen.add(base)
        for name in INPUT_FOLDER_NAMES:
            candidate = os.path.join(base, name)
            if os.path.isdir(candidate) and _looks_like_input_dir(candidate):
                return candidate
        # 한 단계 아래도 봅니다 — 예: D:\PQR\PQR_입력폴더
        try:
            children = sorted(os.listdir(base))
        except OSError:
            continue
        for child in children:
            child_path = os.path.join(base, child)
            if not os.path.isdir(child_path) or child.startswith("."):
                continue
            for name in INPUT_FOLDER_NAMES:
                candidate = os.path.join(child_path, name)
                if os.path.isdir(candidate) and _looks_like_input_dir(candidate):
                    return candidate
    return None


UPDATE_URL = ("https://github.com/joyproject1-ui/Park/archive/refs/heads/"
              "claude/pqr-dashboard-uo3dno.zip")

# 프로그램 파일만 바꿉니다. 담당자가 모아 둔 자료와 만들어 둔 보고서는 그대로 둡니다.
UPDATE_KEEP = ("PQR_입력폴더", "입력폴더", "PQR입력폴더", "out", ".git")


def _program_root():
    """프로그램 폴더(= pqr 패키지의 부모)."""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def cmd_doctor(args):
    """옛 워드(.doc) 변환·엑셀 채우기가 이 PC 에서 되는지 알아봅니다."""
    from .doctor import report
    for line in report():
        _print(line)
    return 0


def cmd_install_claude(args):
    """이 PC 에 Claude Code(claude 명령)를 깔아 13항 손글씨를 이 PC 안에서 읽게 합니다.

    담당자 2026-09-07: "설치하려면 어떻게 해야 돼?" — 명령을 외우지 않아도 되게 두 번 클릭
    파일(PQR-Claude설치.bat)이 이것을 부른다. npm 이 없으면 어디서 무엇을 받아야 하는지 알려 준다.
    """
    import shutil
    import subprocess
    from .engine import claude_cli

    _print("Claude Code 설치 도우미")
    _print("=" * 58)
    already = claude_cli._exe()
    if already:
        _print("  이미 깔려 있습니다: %s" % already)
        _remember_claude(already)
        return _check_login(already)
    npm = shutil.which("npm") or shutil.which("npm.cmd")
    if not npm:
        _print("  [X] npm(Node.js)이 없습니다 — 먼저 Node.js 를 까셔야 합니다.")
        _print("      1) https://nodejs.org 에서 LTS 를 받아 설치하세요 (다음·다음·완료).")
        _print("      2) 설치가 끝나면 이 창을 닫고 PQR-Claude설치.bat 을 다시 두 번 누르세요.")
        _print("         (새로 깐 프로그램은 새 창부터 잡힙니다)")
        return 1
    _print("  npm: %s" % npm)
    _print("  Claude Code 를 내려받아 깝니다 — 몇 분 걸립니다. 창을 닫지 마세요.")
    try:
        run = subprocess.run([npm, "install", "-g", "@anthropic-ai/claude-code"],
                             stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=60 * 20)
    except (OSError, subprocess.TimeoutExpired) as error:
        _print("  [X] 설치하지 못했습니다 — %s" % error)
        return 1
    text = (run.stdout or b"").decode("utf-8", "replace")
    for line in text.splitlines()[-8:]:
        _print("    %s" % line)
    if run.returncode != 0:
        _print("  [X] 설치가 끝나지 못했습니다. 위 글을 그대로 알려 주시면 그 길을 뚫겠습니다.")
        _print("      (회사 망에서 막히면 사내 프록시 설정이 필요할 수 있습니다)")
        return 1
    got = claude_cli._exe()
    if not got:
        _print("  설치는 됐는데 이 창에서는 아직 claude 를 찾지 못합니다 — 창을 닫았다 다시 여세요.")
        _print("  그래도 못 찾으면 명령 프롬프트에서  where claude  를 쳐서 나온 경로 한 줄을")
        _print("  프로그램 폴더의 '%s' 에 적어 두세요." % claude_cli.HINT_FILE)
        return 1
    _print("  [O] 깔렸습니다: %s" % got)
    _remember_claude(got)
    return _check_login(got)


def cmd_login_claude(args):
    """이 PC 의 Claude Code 에 로그인합니다 — 창이 닫히지 않게 이 명령이 붙들고 있습니다.

    담당자 2026-09-07: "이 창에서 claude 를 치면 창이 닫아지는데 어떻게 해?" — 설치 파일은
    끝나면서 창을 닫는다. 그래서 로그인만 하는 두 번 클릭 파일(PQR-Claude로그인.bat)을 두고,
    이 명령이 claude 를 대화형으로 띄운 뒤 정말 되는지까지 확인한다.
    """
    import subprocess
    from .engine import claude_cli
    exe = claude_cli._exe()
    _print("Claude Code 로그인")
    _print("=" * 58)
    if not exe:
        _print("  이 PC 에서 claude 를 찾지 못했습니다 — 먼저 PQR-Claude설치.bat 을 두 번 누르세요.")
        return 1
    _print("  claude: %s" % exe)
    _print("")
    _print("  잠시 뒤 Claude 화면이 이 창에 열립니다. 차례대로 이렇게 하세요:")
    _print("   1) 'Choose the text style…' (글자 스타일) — 그냥 Enter (나중에 /theme 로 바꿉니다)")
    _print("   2) 'Select login method' — 'Claude account with subscription'(구독 계정)을 고르고 Enter")
    _print("      · 'Anthropic Console (API key)' 는 고르지 마세요 — 유료 API 결제 쪽입니다")
    _print("   3) 브라우저가 열리면 늘 쓰시는 계정으로 로그인하고 허용을 누르세요")
    _print("      · 코드를 붙여넣으라고 하면, 이 창에 마우스 오른쪽 클릭(또는 Ctrl+V)으로 붙여넣고 Enter")
    _print("   4) 'Do you trust the files in this folder?' 가 나오면 Yes")
    _print("   5) 입력줄 '>' 이 보이면 로그인 끝 — /exit 를 치고 Enter")
    _print("")
    try:
        subprocess.call(claude_cli.command(exe))      # 대화형 — 이 창을 그대로 물려 준다
    except OSError as error:
        _print("  claude 를 띄우지 못했습니다 — %s" % error)
        return 1
    _print("")
    return _check_login(exe)


def _remember_claude(path):
    """찾은 경로를 프로그램 폴더에 적어 둔다 — 다음 실행에서 헤매지 않게."""
    from .engine import claude_cli
    try:
        with open(os.path.join(_program_root(), claude_cli.HINT_FILE), "w", encoding="utf-8") as handle:
            handle.write(path + "\n")
    except OSError:
        pass


def _check_login(exe):
    """정말로 쓸 수 있는지 한마디 물어본다 — 깔려 있어도 로그인 전이면 판독이 안 된다.

    담당자 2026-09-07: "방법 A 로 완료된 거야?" — 깔렸다는 말만으로는 알 수 없어서, 프로그램이
    실제로 한 번 불러 보고 되는지/무엇이 모자란지 알려 준다.
    """
    from .engine import claude_cli
    _print("")
    _print("  정말로 쓸 수 있는지 한 번 불러 봅니다 (10~60초)…")
    ok, text = claude_cli.ask_once("1+1은? 숫자만 답하세요.", timeout=180)
    if ok:
        _print("  [O] 됩니다 — 로그인까지 끝났습니다. (답: %s)" % " ".join(str(text).split())[:40])
        _print("")
        _print("  이제 대시보드에서 '안정성 판독 🔍' 를 누르면 손글씨 시험일지를 Claude 가")
        _print("  이 PC 안에서 읽습니다 (한 장 1~2분, 사외로 나가지 않습니다).")
        _print("  판독이 끝나면 '보고서 작성' 을 누르세요 — 13항이 채워집니다.")
        return 0
    _print("  [X] 아직 쓸 수 없습니다 — %s" % (" ".join(str(text).split())[:300] or "까닭을 알 수 없습니다"))
    _print("")
    _print("  대개 로그인 전이라 그렇습니다 — 프로그램 폴더의 'PQR-Claude로그인.bat' 을 두 번 누르세요.")
    _print("  (그 창은 닫히지 않습니다. 브라우저가 열리면 늘 쓰시는 계정으로 로그인하고,")
    _print("   끝나면 그 창에 /exit 를 치고 Enter 하면 됩니다.)")
    return 2                     # 2 = 깔려 있는데 로그인 전 — 배치 파일이 로그인 창을 띄운다


def cmd_update(args):
    """최신 버전을 내려받아 프로그램 파일을 바꿉니다.

    담당자가 매번 브라우저로 ZIP 을 받아 폴더를 통째로 바꾸는 일을 없애기 위한
    명령입니다. 입력 폴더와 out 폴더는 손대지 않습니다 — 자료가 사라지면 안 됩니다.
    """
    import shutil
    import tempfile
    import urllib.request
    import zipfile

    root = os.path.abspath(getattr(args, "target", None) or _program_root())
    _print("")
    _print("  PQR 프로그램 업데이트")
    _print("  " + "-" * 56)
    _print("  프로그램 폴더: %s" % root)
    _print("  내려받는 중… (사내망에서는 시간이 걸릴 수 있습니다)")

    workspace = tempfile.mkdtemp(prefix="pqr_update_")
    archive_path = os.path.join(workspace, "update.zip")
    try:
        try:
            with urllib.request.urlopen(args.url, timeout=120) as response:
                payload = response.read()
        except Exception as error:
            _print("")
            _print("  [문제] 내려받지 못했습니다: %s" % error)
            _print("")
            _print("  인터넷이 막혀 있으면 브라우저로 아래 주소를 열어 ZIP 을 받은 뒤,")
            _print("  압축을 풀어 이 폴더의 파일을 덮어써 주세요.")
            _print("  %s" % args.url)
            return 2
        with open(archive_path, "wb") as handle:
            handle.write(payload)

        with zipfile.ZipFile(archive_path) as archive:
            archive.extractall(workspace)
        tops = [name for name in os.listdir(workspace)
                if os.path.isdir(os.path.join(workspace, name))]
        if not tops:
            _print("  [문제] 내려받은 파일에서 프로그램을 찾지 못했습니다.")
            return 2
        source = os.path.join(workspace, tops[0])

        changed = _copy_program(source, root)
        _print("")
        _print("  %d개 파일을 새 것으로 바꿨습니다." % changed)
        _print("  입력 폴더와 out 폴더는 그대로 두었습니다.")
        _print("")
        _print("  화면을 다시 띄우세요:  PQR-대시보드-실행.bat")
        return 0
    finally:
        shutil.rmtree(workspace, ignore_errors=True)


def _copy_program(source, target):
    """새 파일을 덮어씁니다. 지우지는 않습니다 — 담당자 자료가 섞여 있을 수 있습니다."""
    import shutil
    changed = 0
    for base, dirs, files in os.walk(source):
        relative = os.path.relpath(base, source)
        relative = "" if relative == "." else relative
        head = relative.split(os.sep)[0] if relative else ""
        if head in UPDATE_KEEP:
            dirs[:] = []
            continue
        dirs[:] = [name for name in dirs if name not in UPDATE_KEEP]
        destination = os.path.join(target, relative) if relative else target
        if not os.path.isdir(destination):
            os.makedirs(destination)
        for name in files:
            src, dst = os.path.join(base, name), os.path.join(destination, name)
            if _same_file(src, dst):
                continue                      # 같은 내용이면 건드리지 않는다 — 실행 중인 .bat 를
            shutil.copy2(src, dst)            # 덮어쓰면 cmd 가 다음 줄을 엉뚱한 자리에서 읽는다
            changed += 1
    return changed


def _same_file(a, b):
    try:
        if not os.path.isfile(b) or os.path.getsize(a) != os.path.getsize(b):
            return False
        with open(a, "rb") as fa, open(b, "rb") as fb:
            return fa.read() == fb.read()
    except OSError:
        return False


def cmd_launch(args):
    """더블클릭 실행용 — 폴더 준비 · 브라우저 열기 · 서버 실행을 한 번에.

    실행 파일(.bat)이 한글을 다루지 않도록, 안내 문구는 모두 여기서 출력합니다.
    """
    import threading
    import webbrowser
    from . import server as server_module

    input_dir = args.input
    found = None
    if not input_dir:
        found = find_input_dir()
        input_dir = found or DEFAULT_INPUT
    common = os.path.join(input_dir, "공통")
    first_time = not os.path.isdir(input_dir)
    os.makedirs(common, exist_ok=True)

    _print("")
    _print("  PQR 대시보드")
    _print("  " + "-" * 56)
    if found:
        _print("  입력 폴더를 찾았습니다: %s" % found)
    if first_time:
        _print("  입력 폴더를 만들었습니다: %s" % os.path.abspath(input_dir))
        _print("")
        _print("  [먼저 할 일] 제품 마스터 파일을 아래 폴더에 넣어 주세요.")
        _print("               %s" % os.path.abspath(common))
        _print("               넣은 뒤 화면 오른쪽 위의 ↻ 단추를 누르면 반영됩니다.")

    port = _free_port(args.host, args.port)
    try:
        httpd = server_module.serve(input_dir, host=args.host, port=port,
                                    out_dir=args.out, today=args.today, log=_print)
    except OSError as error:
        _print("")
        _print("  [문제] 대시보드를 시작하지 못했습니다: %s" % error)
        return 2

    url = "http://%s:%d" % (args.host, httpd.server_port)
    if not args.no_open:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        _print("")
        _print("  종료합니다.")
    finally:
        httpd.server_close()
    return 0


def cmd_serve(args):
    from . import server as server_module
    if not os.path.isdir(args.input):
        _print("입력 폴더를 찾을 수 없습니다: %s" % args.input)
        return 2
    httpd = server_module.serve(args.input, host=args.host, port=args.port,
                                out_dir=args.out, today=args.today, log=_print)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        _print("")
        _print("종료합니다.")
    finally:
        httpd.server_close()
    return 0


def cmd_narrate(args):
    from . import narrate as narrate_module
    data = _load_data(args.data)
    codes = args.product or None
    try:
        result = narrate_module.narrate(data, codes=codes, model=args.model,
                                        dry_run=args.dry_run, log=_print)
    except RuntimeError as error:
        _print(str(error))
        return 2
    if args.dry_run:
        _print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    _save_json(args.data, data)
    _write_dashboard(data, os.path.join(os.path.dirname(args.data), "dashboard"))
    report_module.write_reports(data, os.path.join(os.path.dirname(args.data), "reports"))
    _print("서술 문안 %d품목을 저장하고 보고서를 다시 만들었습니다." % len(result))
    return 0


def cmd_plan(args):
    """연간 계획서의 제형별 표대로 제품 마스터의 건·생산 Lot 을 맞춥니다."""
    import csv
    from . import plan as plan_module

    lines = plan_module.read_plan(args.doc)
    with io.open(args.master, encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        rows, fields = list(reader), list(reader.fieldnames or [])
    if "비고" not in fields:
        fields.append("비고")
    made = plan_module.apply_to_master(lines, rows)
    target = args.out or args.master
    with io.open(target, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for row in made:
            writer.writerow({key: row.get(key, "") for key in fields})
    added = len(made) - len(rows)
    _print("계획서 %d품목을 읽어 마스터를 %d행 → %d행으로 맞췄습니다 (+%d)."
           % (len(lines), len(rows), len(made), added))
    _print("  저장: %s" % target)
    return 0


def cmd_report(args):
    data = _load_data(args.data)
    out_dir = args.out or os.path.join(os.path.dirname(args.data), "reports")
    written = report_module.write_reports(data, out_dir, codes=args.product or None)
    _print("보고서 %d건을 만들었습니다: %s" % (len(written), out_dir))
    return 0


def cmd_write(args):
    """대시보드 없이 한 제품의 제출용 보고서를 만든다 — '보고서 작성' 단추와 같은 길.

    PC 에서 Claude(Cowork·Claude Code)가 제품 폴더를 직접 보며 작성할 때 쓴다: 판독 파일을
    폴더에 놓고 이 명령을 돌리고, 나온 문의 목록을 읽고, 고치고, 다시 돌린다 — 담당자 2026-09-08:
    "너가 직접 Cowork 가 하는 것처럼 대시보드 폴더에 접근해서 PQR 파일을 직접 읽고 작성해 줘."
    """
    from . import build as build_module, docx_report, server as server_module
    from .engine import writer as engine_writer
    folder = os.path.abspath(args.folder)
    if not os.path.isdir(folder):
        _print("제품 폴더를 찾을 수 없습니다: %s" % folder)
        return 2
    input_dir = os.path.dirname(folder)
    code = build_module._folder_product_code(os.path.basename(folder))
    workspace = server_module.Workspace(input_dir, out_dir=args.out or os.path.join(input_dir, "..", "out"),
                                        today=args.today)
    data = workspace.data                       # Workspace 가 만들며 이미 집계했다
    product = next((item for item in data.get("products") or [] if item["code"] == code), None)
    if product is None:
        # 제품 마스터가 없어도 폴더 이름('QC1-5087 아이퓨어점안액')만으로 만든다 — 제형은 이름으로 짐작
        name = " ".join(os.path.basename(folder).split()[1:]) or code
        group = "점안제" if "점안" in name else "연고제" if "연고" in name else ""
        product = {"code": code, "name": name, "group": group}
        _print("제품 마스터에 %s 가 없어 폴더 이름으로 만듭니다: %s (%s)" % (code, name, group or "제형 모름"))
    period = dict(data.get("period") or {})
    if not (period.get("from") and period.get("to")):
        # 제품 마스터가 없으면 평가 기간도 없다 — PQR 은 지난 한 해를 평가한다
        import datetime as _dt
        year = (schema.parse_date(args.today) if args.today else _dt.date.today()).year - 1
        period = {"from": "%d-01-01" % year, "to": "%d-12-31" % year}
    target = os.path.join(server_module._made_dir(folder), docx_report.report_filename(product, period))
    _print("[%s] %s — %s ~ %s" % (product["code"], product["name"], period.get("from"), period.get("to")))
    steps = []

    def say(msg):
        steps.append(msg)
        _print("  " + msg)

    try:
        result = engine_writer.write_report(folder, product, period, target, today=data.get("today"),
                                            log=say, vision=server_module._vision_hook())
    except Exception as error:
        import traceback
        trace = traceback.format_exc()
        server_module.write_work_log(folder, product, steps, str(error), trace)
        server_module.write_failure_note(server_module._made_dir(folder), product, str(error), trace, steps)
        _print("작성 실패: %s" % error)
        _print(trace)
        return 1
    issues = list(result.get("issues") or [])
    from .engine import review as review_module
    issues += review_module.review(folder, product, target, issues, log=say, period=period)   # Claude Code 가 있으면 검토
    server_module.write_issue_list(folder, product, issues)
    server_module.write_work_log(folder, product, steps)
    if result.get("blank_sections"):
        build_module.mark_auto_draft(folder, target)
    else:
        build_module.unmark_auto_draft(folder, target)
    _print("")
    _print("보고서: %s" % result.get("path"))
    _print("문의 목록 %d건 (%s)" % (len(issues), os.path.join(folder, server_module.ISSUE_LIST_NAME % product["code"])))
    for i, (item, name, why) in enumerate(issues, 1):
        _print("  %2d. [%s] %s — %s" % (i, item, name, why))
    return 0


def build_parser():
    parser = argparse.ArgumentParser(
        prog="pqr", description="제품품질평가(PQR) 자동 집계 · 보고서 생성 도구")
    subparsers = parser.add_subparsers(dest="command")

    demo = subparsers.add_parser("demo", help="예시 입력 파일 만들기")
    demo.add_argument("-o", "--out", default="sample_input", help="저장 폴더 (기본 sample_input)")
    demo.add_argument("--layout", choices=["tree", "flat"], default="tree",
                      help="tree: 제품 폴더 방식(기본) · flat: 한 폴더에 모아두기")
    demo.set_defaults(func=cmd_demo)

    init = subparsers.add_parser("init", help="담당자가 자료를 올릴 제품 폴더 만들기")
    init.add_argument("-o", "--out", required=True, help="제품 폴더를 만들 위치")
    init.add_argument("--master", help="제품 마스터 파일 (제품코드·제품명)")
    init.add_argument("-p", "--product", action="append",
                      help="제품 직접 지정 (예: --product HP-101=히알로타인 점안액)")
    init.set_defaults(func=cmd_init)

    check = subparsers.add_parser("check", help="파일 인식 · 열 매핑 점검")
    check.add_argument("-i", "--in", dest="input", required=True, help="입력 폴더")
    check.set_defaults(func=cmd_check)

    build_cmd = subparsers.add_parser("build", help="집계 · 판정 → 대시보드 데이터 + 보고서")
    build_cmd.add_argument("-i", "--in", dest="input", required=True, help="입력 폴더")
    build_cmd.add_argument("-o", "--out", default=DEFAULT_OUT, help="출력 폴더 (기본 out)")
    build_cmd.add_argument("--today", help="기준일 (기본 오늘, 예: 2026-08-27)")
    build_cmd.add_argument("--period-from", help="평가기간 시작 (제품 마스터가 없을 때)")
    build_cmd.add_argument("--period-to", help="평가기간 종료")
    build_cmd.add_argument("--narrate", action="store_true",
                           help="Claude API 로 서술 문안 초안까지 생성")
    build_cmd.add_argument("--model", default="claude-opus-5", help="서술 문안 생성 모델")
    build_cmd.set_defaults(func=cmd_build)

    launch = subparsers.add_parser(
        "launch", help="폴더 준비 · 브라우저 열기 · 서버 실행을 한 번에 (더블클릭 실행용)")
    launch.add_argument("-i", "--in", dest="input", help="입력 폴더 (기본 '입력폴더')")
    launch.add_argument("-o", "--out", default=DEFAULT_OUT, help="보고서 출력 폴더")
    launch.add_argument("--host", default="127.0.0.1")
    launch.add_argument("--port", type=int, default=8787,
                        help="이미 쓰이고 있으면 다음 번호를 씁니다")
    launch.add_argument("--today", help="기준일 (기본 오늘)")
    launch.add_argument("--no-open", action="store_true", help="브라우저를 열지 않음")
    launch.set_defaults(func=cmd_launch)

    serve_cmd = subparsers.add_parser(
        "serve", help="대시보드를 띄우고 화면에서 자료를 올릴 수 있게 함")
    serve_cmd.add_argument("-i", "--in", dest="input", required=True, help="입력 폴더")
    serve_cmd.add_argument("-o", "--out", default=DEFAULT_OUT, help="보고서 출력 폴더 (기본 out)")
    serve_cmd.add_argument("--host", default="127.0.0.1",
                           help="기본 127.0.0.1 (이 PC에서만 접속). 사내에 열려면 0.0.0.0")
    serve_cmd.add_argument("--port", type=int, default=8787)
    serve_cmd.add_argument("--today", help="기준일 (기본 오늘)")
    serve_cmd.set_defaults(func=cmd_serve)

    narrate_cmd = subparsers.add_parser("narrate", help="서술 문안 초안 (Claude API)")
    narrate_cmd.add_argument("-d", "--data", required=True, help="build 가 만든 pqr.json")
    narrate_cmd.add_argument("-p", "--product", action="append", help="제품 코드 (반복 가능)")
    narrate_cmd.add_argument("--model", default="claude-opus-5")
    narrate_cmd.add_argument("--dry-run", action="store_true",
                             help="전송될 내용만 출력하고 API 를 호출하지 않음")
    narrate_cmd.set_defaults(func=cmd_narrate)

    doctor_cmd = subparsers.add_parser(
        "doctor", help="이 PC 에서 무엇이 되고 무엇이 막혔는지 알아봅니다")
    doctor_cmd.set_defaults(func=cmd_doctor)

    claude_cmd = subparsers.add_parser(
        "install-claude", help="이 PC 에 Claude Code 를 깔아 13항 손글씨를 여기서 읽게 합니다")
    claude_cmd.set_defaults(func=cmd_install_claude)

    login_cmd = subparsers.add_parser(
        "login-claude", help="이 PC 의 Claude Code 에 로그인합니다 (창이 닫히지 않습니다)")
    login_cmd.set_defaults(func=cmd_login_claude)

    update_cmd = subparsers.add_parser(
        "update", help="프로그램을 최신 버전으로 바꿉니다 (입력 폴더는 건드리지 않습니다)")
    update_cmd.add_argument("--url", default=UPDATE_URL, help="내려받을 ZIP 주소")
    update_cmd.add_argument("--dir", dest="target", help="바꿀 프로그램 폴더 (기본: 지금 이 폴더)")
    update_cmd.set_defaults(func=cmd_update)

    plan_cmd = subparsers.add_parser(
        "plan", help="연간 계획서로 제품 마스터의 생산 Lot·구분 맞추기")
    plan_cmd.add_argument("--doc", required=True, help="연간 계획서 (.doc/.docx)")
    plan_cmd.add_argument("--master", required=True, help="제품 마스터 (.csv)")
    plan_cmd.add_argument("-o", "--out", help="저장할 파일 (없으면 --master 를 덮어씁니다)")
    plan_cmd.set_defaults(func=cmd_plan)

    write_cmd = subparsers.add_parser(
        "write", help="한 제품의 제출용 보고서를 만든다 (대시보드의 '보고서 작성' 과 같음)")
    write_cmd.add_argument("folder", help="제품 폴더 (예: PQR_입력폴더\\QC1-5087 아이퓨어점안액)")
    write_cmd.add_argument("-o", "--out", help="집계 결과를 둘 폴더 (기본: 입력 폴더 옆 out)")
    write_cmd.add_argument("--today", help="작성 일자 YYYY-MM-DD (기본: 오늘)")
    write_cmd.set_defaults(func=cmd_write)

    report_cmd = subparsers.add_parser("report", help="보고서만 다시 생성")
    report_cmd.add_argument("-d", "--data", required=True, help="build 가 만든 pqr.json")
    report_cmd.add_argument("-p", "--product", action="append", help="제품 코드 (반복 가능)")
    report_cmd.add_argument("-o", "--out", help="출력 폴더")
    report_cmd.set_defaults(func=cmd_report)
    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    if not getattr(args, "func", None):
        parser.print_help()
        return 0
    try:
        return args.func(args)
    except BrokenPipeError:
        # `... | head` 처럼 받는 쪽이 먼저 닫힌 경우 — 조용히 끝냅니다.
        os.dup2(os.open(os.devnull, os.O_WRONLY), sys.stdout.fileno())
        return 0
    except KeyboardInterrupt:
        _print("")
        _print("중단했습니다.")
        return 130


if __name__ == "__main__":
    sys.exit(main())
