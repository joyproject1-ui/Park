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


def cmd_launch(args):
    """더블클릭 실행용 — 폴더 준비 · 브라우저 열기 · 서버 실행을 한 번에.

    실행 파일(.bat)이 한글을 다루지 않도록, 안내 문구는 모두 여기서 출력합니다.
    """
    import threading
    import webbrowser
    from . import server as server_module

    input_dir = args.input or DEFAULT_INPUT
    common = os.path.join(input_dir, "공통")
    first_time = not os.path.isdir(input_dir)
    os.makedirs(common, exist_ok=True)

    _print("")
    _print("  PQR 대시보드")
    _print("  " + "-" * 56)
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


def cmd_report(args):
    data = _load_data(args.data)
    out_dir = args.out or os.path.join(os.path.dirname(args.data), "reports")
    written = report_module.write_reports(data, out_dir, codes=args.product or None)
    _print("보고서 %d건을 만들었습니다: %s" % (len(written), out_dir))
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
