"""담당자용 로컬 서버 (`python -m pqr serve`).

대시보드를 브라우저로 열고, 평가항목 칸을 눌러 그 자리에서 자료를 올릴 수 있게 합니다.
올린 파일은 해당 **제품 폴더**에 저장되고 곧바로 다시 집계되어 화면에 반영됩니다.

표준 라이브러리(http.server)만 사용하며, 기본적으로 이 PC(127.0.0.1)에서만 접속됩니다.
"""

import json
import os
import posixpath
import re
import subprocess
import sys
import shutil
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from . import build as build_module
from . import report as report_module

MAX_UPLOAD = 64 * 1024 * 1024          # 파일 하나당 64 MB
ALLOWED_SUFFIXES = (".csv", ".tsv", ".txt", ".xlsx", ".xlsm")
# 평가항목 근거 자료는 회사 원본 그대로 들어옵니다 — 성적서 PDF, ERP 가 뽑은 구형 .xls,
# 한글 문서 등. 표로 읽지 않고 '자료가 왔다' 는 근거로 두므로 형식을 넓게 받습니다.
ALLOWED_ITEM_SUFFIXES = ALLOWED_SUFFIXES + (
    ".pdf", ".xls", ".doc", ".docx", ".hwp", ".hwpx", ".ppt", ".pptx",
    ".png", ".jpg", ".jpeg", ".zip")
_UNSAFE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


class UploadError(Exception):
    """사용자에게 그대로 보여줄 수 있는 업로드 오류."""


def safe_filename(name, allowed=None):
    """경로 요소를 제거하고 확장자를 확인합니다."""
    allowed = allowed or ALLOWED_SUFFIXES
    base = os.path.basename(str(name or "").replace("\\", "/")).strip()
    base = _UNSAFE.sub("_", base).lstrip(".")
    if not base:
        raise UploadError("파일 이름이 비어 있습니다.")
    if not base.lower().endswith(allowed):
        raise UploadError("올릴 수 있는 형식은 %s 입니다." % ", ".join(allowed))
    return base


def parse_multipart(content_type, body):
    """multipart/form-data 를 {필드명: 값 또는 (파일명, 바이트)} 로 바꿉니다.

    한국어 파일 이름은 헤더에 UTF-8 바이트로 그대로 실려 오므로, email 모듈에
    맡기지 않고 경계선으로 직접 잘라 UTF-8 로 읽습니다.
    """
    match = re.search(r'boundary="?([^";]+)"?', content_type or "")
    if not match:
        raise UploadError("form-data 형식이 아닙니다.")
    boundary = b"--" + match.group(1).encode("utf-8")

    fields = {}
    for chunk in body.split(boundary)[1:]:
        if chunk.startswith(b"\r\n"):
            chunk = chunk[2:]
        if chunk.endswith(b"\r\n"):
            chunk = chunk[:-2]
        if not chunk or chunk.startswith(b"--"):
            continue                      # 마지막 경계선
        head, separator, payload = chunk.partition(b"\r\n\r\n")
        if not separator:
            continue
        headers = head.decode("utf-8", "replace")
        name = re.search(r'name="([^"]*)"', headers)
        if not name:
            continue
        filename = re.search(r'filename="([^"]*)"', headers)
        if filename:
            fields[name.group(1)] = (filename.group(1), payload)
        else:
            fields[name.group(1)] = payload.decode("utf-8", "replace").strip()
    return fields


def open_in_file_manager(path):
    """폴더를 운영체제 파일 관리자로 엽니다 (안 되면 조용히 False).

    브라우저는 로컬 폴더를 열 수 없으므로, 담당자 PC 에서 도는 이 서버가 대신 엽니다.
    """
    try:
        if sys.platform == "win32":
            os.startfile(path)                                  # noqa: S606 — 로컬 도구
            return True
        command = ["open", path] if sys.platform == "darwin" else ["xdg-open", path]
        subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        return False


class Workspace(object):
    """입력 폴더 하나를 다루는 작업 공간. 적재 결과를 캐시로 들고 있습니다."""

    def __init__(self, input_dir, out_dir="out", today=None, config=None):
        self.input_dir = os.path.abspath(input_dir)
        self.out_dir = os.path.abspath(out_dir)
        self.today = today
        self.config = config or build_module.load_config()
        self.lock = threading.Lock()
        self.data = None
        self.rebuild()

    # ---------------- 집계 ----------------

    def rebuild(self):
        with self.lock:
            self.data = build_module.build(input_dir=self.input_dir, today=self.today,
                                           config=self.config)
            return self.data

    def data_js(self):
        payload = self.dashboard_payload()
        return "window.PQR_DATA = %s;\n" % json.dumps(payload, ensure_ascii=False)

    def dashboard_payload(self):
        data = self.data
        payload = {key: data[key] for key in
                   ("generated_at", "today", "period", "stages", "items", "products",
                    "trend", "leadtime", "sources", "narrative")}
        payload["issue_count"] = len([i for i in data.get("issues", []) if i["level"] == "error"])
        payload["upload"] = {
            "enabled": True,
            "input_dir": self.input_dir,
            "item_datasets": self.config["item_datasets"],
            "datasets": self.config["dataset_files"],
        }
        return payload

    # ---------------- 업로드 ----------------

    def product_folder(self, code):
        """제품 폴더를 찾고, 없으면 만듭니다."""
        code = (code or "").strip().upper()
        if not code or _UNSAFE.search(code) or code in (".", ".."):
            raise UploadError("제품 코드가 올바르지 않습니다.")
        for name in sorted(os.listdir(self.input_dir)):
            path = os.path.join(self.input_dir, name)
            if os.path.isdir(path) and build_module._folder_product_code(name) == code:
                return path
        product = next((item for item in self.data["products"] if item["code"] == code), None)
        label = ("%s %s" % (code, product["name"] if product else "")).strip()
        path = os.path.join(self.input_dir, _UNSAFE.sub("_", label))
        os.makedirs(path, exist_ok=True)
        return path

    def common_folder(self):
        for name in build_module.COMMON_FOLDERS:
            path = os.path.join(self.input_dir, name)
            if os.path.isdir(path):
                return path
        path = os.path.join(self.input_dir, build_module.COMMON_FOLDERS[0])
        os.makedirs(path, exist_ok=True)
        return path

    def target_path(self, dataset, code, filename):
        """저장 위치를 정합니다. 공통 자료는 '공통' 폴더로 갑니다."""
        spec = self.config["dataset_files"].get(dataset)
        if spec is None:
            raise UploadError("알 수 없는 자료 종류입니다: %s" % dataset)
        folder = self.common_folder() if spec["scope"] == "common" else self.product_folder(code)
        base = safe_filename(filename)
        stem, suffix = os.path.splitext(base)
        # 파일 이름으로 종류를 알아보므로, 인식용 낱말이 없으면 앞에 붙여 줍니다.
        from .schema import detect_dataset
        if detect_dataset(base) != dataset:
            stem = "%s_%s" % (spec["filename"], stem)
        return os.path.join(folder, stem + suffix)

    def inspect(self, temp_path, dataset, code):
        """저장 전에 파일을 읽어 봅니다. (정규화된 행, 문제 목록) 을 돌려줍니다."""
        from . import schema
        from .tabular import TableError, read_table
        try:
            raw = read_table(temp_path)
        except TableError as error:
            raise UploadError(str(error))
        if not raw:
            raise UploadError("빈 표입니다. 첫 줄에 열 이름이 있는지 확인하세요.")
        folder_code = code if self.config["dataset_files"][dataset]["scope"] == "product" else None
        return schema.normalize(raw, dataset, os.path.basename(temp_path), folder_code) + (raw,)

    def guess_dataset(self, raw, code):
        """열 구성을 보고 어느 자료에 가장 잘 맞는지 추정합니다.

        적재되는 행 수가 같으면(여러 종류의 필수 열을 모두 갖춘 경우), 그 종류의
        표준 필드를 더 많이 채우는 쪽을 고릅니다. 예를 들어 '조건'·'시점' 열이 있으면
        시험성적서보다 안정성 자료로 봅니다.
        """
        from . import schema
        best, best_score = None, (0, 0)
        for name, spec in self.config["dataset_files"].items():
            folder_code = code if spec["scope"] == "product" else None
            rows, _ = schema.normalize(list(raw), name, "", folder_code)
            if not rows:
                continue
            fields = set()
            for row in rows:
                fields.update(key for key, value in row.items() if value not in (None, ""))
            score = (len(rows), len(fields))
            if score > best_score:
                best, best_score = name, score
        return best

    def save_item_file(self, code, item_id, filename, payload):
        """평가항목 자료를 제품 폴더에 '항 번호로 시작하는 이름' 으로 저장합니다.

        담당자가 올리는 것은 표준 대장이 아니라 회사 원본(공급업체 List · 성적서 PDF …)
        입니다. 이것을 대장으로 읽으려 들면 맞지도 않는 자료 종류가 붙고 저장 위치도
        엉뚱해집니다 — '원료 공급업체 List' 를 '설비 적격성' 으로 저장하던 문제입니다.
        파일 이름 앞의 항 번호가 곧 인식 규칙이므로 그대로 따릅니다.
        """
        if len(payload) > MAX_UPLOAD:
            raise UploadError("파일이 너무 큽니다 (최대 %d MB)." % (MAX_UPLOAD // 1024 // 1024))
        if not payload:
            raise UploadError("빈 파일입니다.")
        labels = {row[0]: row[1] for row in self.data["items"]}
        if item_id not in labels:
            raise UploadError("평가항목을 찾지 못했습니다: %s" % item_id)
        folder = self.product_folder(code)
        base = safe_filename(filename, ALLOWED_ITEM_SUFFIXES) or "자료"
        # 회사 원본은 이미 항 번호로 시작하는 일이 많습니다. 그럴 때는 이름을 그대로 둡니다 —
        # 앞에 번호를 또 붙이면 '8.1.1 … - 8.1.1 …' 처럼 됩니다.
        matcher = build_module.item_matcher(self.data["items"])
        if matcher(base) == item_id:
            target = os.path.join(folder, base)
        else:
            stem, suffix = os.path.splitext(base)
            label = " ".join(_UNSAFE.sub(" ", labels[item_id]).split())
            target = os.path.join(folder, "%s %s - %s%s" % (item_id, label, stem, suffix))
        with self.lock:
            with open(target, "wb") as handle:
                handle.write(payload)
        self.rebuild()
        return {"saved": os.path.relpath(target, self.input_dir),
                "folder": os.path.abspath(folder),
                "name": os.path.basename(target)}

    def save_upload(self, dataset, code, filename, payload, replace=True):
        """검증한 뒤에만 제품(또는 공통) 폴더에 저장하고 다시 집계합니다.

        맞지 않는 파일이 폴더에 남지 않도록, 임시 파일로 먼저 읽어 보고
        한 행도 적재되지 않으면 저장하지 않습니다.
        """
        if len(payload) > MAX_UPLOAD:
            raise UploadError("파일이 너무 큽니다 (최대 %d MB)." % (MAX_UPLOAD // 1024 // 1024))
        if not payload:
            raise UploadError("빈 파일입니다.")
        path = self.target_path(dataset, code, filename)

        temp_dir = tempfile.mkdtemp(prefix="pqr-upload-")
        temp_path = os.path.join(temp_dir, os.path.basename(path))
        try:
            with open(temp_path, "wb") as handle:
                handle.write(payload)
            rows, issues, raw = self.inspect(temp_path, dataset, code)
            if not rows:
                label = self.config["dataset_files"][dataset]["label"]
                guess = self.guess_dataset(raw, code)
                detail = "'%s' 자료로 읽었지만 한 행도 적재되지 않았습니다." % label
                if guess and guess != dataset:
                    detail += " 이 파일은 '%s' 자료로 보입니다 — 자료 종류를 바꿔 다시 올려 보세요." \
                              % self.config["dataset_files"][guess]["label"]
                else:
                    missing = [issue["field"] for issue in issues if issue["level"] == "error"]
                    if missing:
                        detail += " 필수 항목이 비어 있습니다: %s" % ", ".join(sorted(set(missing))[:4])
                raise UploadError(detail)

            with self.lock:
                backup = None
                if os.path.exists(path):
                    if replace:
                        backup = path + ".bak"
                        shutil.copy2(path, backup)
                    else:
                        stem, suffix = os.path.splitext(path)
                        index = 2
                        while os.path.exists("%s(%d)%s" % (stem, index, suffix)):
                            index += 1
                        path = "%s(%d)%s" % (stem, index, suffix)
                shutil.move(temp_path, path)
            try:
                data = self.rebuild()
            except Exception:                   # 집계가 깨지면 이전 상태로 되돌립니다.
                if backup:
                    shutil.move(backup, path)
                elif os.path.exists(path):
                    os.remove(path)
                self.rebuild()
                raise
            if backup and os.path.exists(backup):
                os.remove(backup)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

        # 다른 제품 폴더에 같은 이름의 파일이 있을 수 있으므로 전체 경로로 찾습니다.
        name = os.path.basename(path)
        target = os.path.abspath(path)
        loaded = [entry for entries in data["sources"].values() for entry in entries
                  if entry.get("path") == target]
        return {
            "saved": os.path.relpath(path, self.input_dir),
            "dataset": dataset,
            "rows": sum(entry["rows"] for entry in loaded),
            "skipped": sum(entry["skipped"] for entry in loaded),
            "errors": [i for i in data["issues"] if i["source"] == name and i["level"] == "error"][:10],
            "warnings": [i for i in data["issues"]
                         if i["source"] == name and i["level"] == "warning"][:10],
        }

    def write_outputs(self):
        os.makedirs(self.out_dir, exist_ok=True)
        with open(os.path.join(self.out_dir, "pqr.json"), "w", encoding="utf-8") as handle:
            json.dump(self.data, handle, ensure_ascii=False, indent=2)
        return report_module.write_reports(self.data, os.path.join(self.out_dir, "reports"))


class Handler(BaseHTTPRequestHandler):
    workspace = None
    dashboard_dir = None
    server_version = "pqr"

    def log_message(self, fmt, *args):           # 접속 로그를 간결하게
        print("  %s %s" % (self.command, self.path.split("?")[0]))

    # ---------------- 응답 도우미 ----------------

    def _send(self, code, body, content_type="application/json; charset=utf-8"):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code, payload):
        self._send(code, json.dumps(payload, ensure_ascii=False))

    # ---------------- 라우팅 ----------------

    def do_GET(self):
        path = posixpath.normpath(self.path.split("?")[0])
        if path in ("/", "/index.html"):
            return self._serve_file("index.html", "text/html; charset=utf-8")
        if path == "/data.js":
            return self._send(200, self.workspace.data_js(),
                              "application/javascript; charset=utf-8")
        if path == "/api/data":
            return self._json(200, self.workspace.dashboard_payload())
        if path == "/api/health":
            return self._json(200, {"ok": True, "input_dir": self.workspace.input_dir})
        if path == "/favicon.ico":
            return self._send(404, b"", "text/plain")
        return self._serve_file(path.lstrip("/"), None)

    def _serve_file(self, relative, content_type):
        base = self.dashboard_dir
        target = os.path.normpath(os.path.join(base, relative))
        if not target.startswith(base) or not os.path.isfile(target):
            return self._send(404, "찾을 수 없습니다", "text/plain; charset=utf-8")
        guess = content_type or {
            ".html": "text/html; charset=utf-8", ".js": "application/javascript; charset=utf-8",
            ".css": "text/css; charset=utf-8", ".json": "application/json; charset=utf-8",
        }.get(os.path.splitext(target)[1], "application/octet-stream")
        with open(target, "rb") as handle:
            self._send(200, handle.read(), guess)

    def do_POST(self):
        path = posixpath.normpath(self.path.split("?")[0])
        try:
            if path == "/api/upload":
                return self._json(200, self._handle_upload())
            if path == "/api/rebuild":
                self.workspace.rebuild()
                return self._json(200, {"ok": True, "data": self.workspace.dashboard_payload()})
            if path == "/api/export":
                written = self.workspace.write_outputs()
                return self._json(200, {"ok": True, "files": [os.path.basename(p) for p in written],
                                        "out_dir": self.workspace.out_dir})
            if path == "/api/open":
                return self._json(200, self._handle_open())
            if path == "/api/close":
                return self._json(200, self._handle_close())
            if path == "/api/report":
                return self._json(200, self._handle_report())
            if path == "/api/final":
                return self._json(200, self._handle_final())
        except UploadError as error:
            return self._json(400, {"ok": False, "error": str(error)})
        except Exception as error:               # 서버가 죽지 않도록 오류를 그대로 알려 줍니다.
            return self._json(500, {"ok": False, "error": "%s: %s"
                                    % (type(error).__name__, error)})
        return self._send(404, "찾을 수 없습니다", "text/plain; charset=utf-8")

    def _read_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0 or length > 65536:
            raise UploadError("요청 내용이 없습니다.")
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except ValueError:
            raise UploadError("요청을 해석하지 못했습니다.")

    def _handle_open(self):
        """제품 폴더를 만들고(없으면) 파일 관리자로 엽니다 — 담당자가 바로 파일을 끌어다 놓게."""
        body = self._read_json()
        folder = self.workspace.product_folder(str(body.get("product") or "").strip())
        opened = open_in_file_manager(folder)
        return {"ok": True, "path": folder, "opened": opened,
                "hint": "" if opened else "폴더를 자동으로 열지 못했습니다. 위 경로를 파일 탐색기에 붙여넣으세요."}

    def _handle_close(self):
        """평가항목을 '확인 결과 이력 없음' 으로 마감합니다.

        그냥 초록으로 칠하지 않고 제품 폴더에 확인 기록 파일을 남깁니다 —
        공란과 '확인했더니 없음' 은 다른 것이고, 근거는 폴더에 남아야 합니다.
        파일 이름이 항 번호로 시작하므로 기존 인식 규칙이 그대로 녹색불을 켭니다.
        잘못 눌렀다면 폴더에서 그 파일을 지우면 됩니다.
        """
        import datetime
        body = self._read_json()
        code = str(body.get("product") or "").strip()
        item_id = str(body.get("item") or "").strip()
        labels = {row[0]: row[1] for row in self.workspace.data["items"]}
        if item_id not in labels:
            raise UploadError("평가항목을 찾지 못했습니다: %s" % item_id)
        product = next((item for item in self.workspace.data["products"]
                        if item["code"] == code), None)
        if product is None:
            raise UploadError("제품을 찾지 못했습니다: %s" % code)
        folder = self.workspace.product_folder(code)
        label = _UNSAFE.sub(" ", labels[item_id])
        label = " ".join(label.split())
        filename = "%s %s - %s.txt" % (item_id, label, build_module.CLOSE_MARKER)
        target = os.path.join(folder, _UNSAFE.sub("_", filename))
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        note = str(body.get("note") or "").strip()
        lines = [
            "평가항목 마감 기록",
            "==================",
            "제품      : [%s] %s" % (code, product["name"]),
            "평가항목  : (%s) %s" % (item_id, labels[item_id]),
            "확인 결과 : 평가 년도 내 해당 이력 없음",
            "확인 일시 : %s" % stamp,
        ]
        if note:
            lines.append("비고      : %s" % note)
        lines += [
            "확인자    : (서명 또는 이름 기입)",
            "",
            "이 파일은 대시보드의 '이력 없음으로 마감' 단추가 만들었습니다.",
            "마감을 취소하려면 이 파일을 지우고 화면의 새로고침(↻)을 누르세요.",
        ]
        with open(target, "w", encoding="utf-8") as handle:
            handle.write("\n".join(lines) + "\n")
        self.workspace.rebuild()
        return {"ok": True, "saved": os.path.relpath(target, self.workspace.input_dir),
                "data": self.workspace.dashboard_payload()}

    def _handle_report(self):
        """한 제품의 보고서 초안을 만듭니다. 자료 수집이 덜 됐으면 만들지 않고 알려 줍니다."""
        from . import report as report_module
        body = self._read_json()
        code = str(body.get("product") or "").strip()
        product = next((item for item in self.workspace.data["products"]
                        if item["code"] == code), None)
        if product is None:
            raise UploadError("제품을 찾지 못했습니다: %s" % code)
        if product.get("pct", 0) < 100 and not body.get("force"):
            missing = [item_id for item_id, state
                       in zip((row[0] for row in self.workspace.data["items"]), product["checks"])
                       if state == "n"]
            return {"ok": False, "error": "자료 수집이 %d%% 입니다. 미착수 항목: %s"
                    % (product.get("pct", 0), ", ".join(missing) or "-")}
        from . import docx_report
        out_dir = os.path.join(self.workspace.out_dir, "reports")
        written = report_module.write_reports(self.workspace.data, out_dir, codes=[code])
        # 제출용 보고서(.docx)는 제품 폴더에 둡니다 — 근거 자료 옆에 있어야 검토가 쉽고,
        # 이름에 '제출용' 이 들어가 있어 다음 새로고침에서 '완성본' 단추가 생깁니다.
        folder = self.workspace.product_folder(code)
        final = docx_report.write_docx(self.workspace.data, code, folder,
                                       config=self.workspace.config)
        self.workspace.rebuild()
        # '어디 있지?' 를 없앱니다 — 제출용 보고서가 있는 제품 폴더를 바로 열어 줍니다.
        opened = open_in_file_manager(folder)
        return {"ok": True, "files": [os.path.abspath(p) for p in written],
                "final": os.path.abspath(final), "final_name": os.path.basename(final),
                "folder": os.path.abspath(folder),
                "out_dir": os.path.abspath(out_dir), "opened": opened,
                "data": self.workspace.dashboard_payload()}

    def _handle_final(self):
        """제품 폴더의 완성본(제출용) 보고서를 찾아 엽니다.

        완성본은 프로그램이 만들지 않습니다 — 실제 자료로 작성한 제출용 문서를
        담당자가 제품 폴더에 넣으면, 파일 이름의 '완성본'/'제출' 로 알아봅니다.
        """
        body = self._read_json()
        code = str(body.get("product") or "").strip()
        product = next((item for item in self.workspace.data["products"]
                        if item["code"] == code), None)
        if product is None:
            raise UploadError("제품을 찾지 못했습니다: %s" % code)
        folder = self.workspace.product_folder(code)
        matcher = build_module.item_matcher(self.workspace.data["items"])
        # 화면의 목록이 오래됐을 수 있으니 폴더를 지금 다시 봅니다.
        target = build_module.find_final_report(folder, matcher)
        if not target:
            return {"ok": False, "error": "완성본 파일이 없습니다. 파일 이름에 '완성본' 또는 "
                    "'제출' 이 들어간 문서(.docx 등)를 제품 폴더에 넣으세요."}
        opened = open_in_file_manager(target)
        return {"ok": True, "path": os.path.abspath(target),
                "name": os.path.basename(target), "opened": opened,
                "hint": "" if opened else "파일을 자동으로 열지 못했습니다. 위 경로를 파일 탐색기에 붙여넣으세요."}

    def _handle_upload(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            raise UploadError("업로드된 내용이 없습니다.")
        if length > MAX_UPLOAD + 65536:
            raise UploadError("파일이 너무 큽니다 (최대 %d MB)." % (MAX_UPLOAD // 1024 // 1024))
        body = self.rfile.read(length)
        fields = parse_multipart(self.headers.get("Content-Type", ""), body)
        upload = fields.get("file")
        if not isinstance(upload, tuple):
            raise UploadError("파일을 선택하세요.")
        item_id = (fields.get("item") or "").strip()
        if item_id:
            result = self.workspace.save_item_file(
                code=fields.get("product") or "", item_id=item_id,
                filename=upload[0], payload=upload[1])
            result["ok"] = True
            result["data"] = self.workspace.dashboard_payload()
            return result
        result = self.workspace.save_upload(
            dataset=fields.get("dataset") or "",
            code=fields.get("product") or "",
            filename=upload[0],
            payload=upload[1],
            replace=fields.get("replace", "1") != "0",
        )
        result["ok"] = True
        result["data"] = self.workspace.dashboard_payload()
        return result


def serve(input_dir, host="127.0.0.1", port=8787, out_dir="out", today=None,
          dashboard_dir=None, log=print):
    """서버를 띄웁니다. 테스트에서는 port=0 으로 부르고 server.server_port 를 읽습니다."""
    dashboard_dir = os.path.abspath(
        dashboard_dir or os.path.join(os.path.dirname(os.path.dirname(
            os.path.abspath(__file__))), "docs", "pqr"))
    if not os.path.isfile(os.path.join(dashboard_dir, "index.html")):
        raise UploadError("대시보드를 찾을 수 없습니다: %s" % dashboard_dir)

    handler = type("BoundHandler", (Handler,), {
        "workspace": Workspace(input_dir, out_dir=out_dir, today=today),
        "dashboard_dir": dashboard_dir,
    })
    httpd = ThreadingHTTPServer((host, port), handler)
    log("입력 폴더: %s" % handler.workspace.input_dir)
    log("제품 %d품목을 불러왔습니다." % len(handler.workspace.data["products"]))
    log("")
    log("  브라우저에서 열기:  http://%s:%d" % (host, httpd.server_port))
    log("  평가항목 칸을 누르면 그 자리에서 자료를 올릴 수 있습니다. (Ctrl+C 로 종료)")
    return httpd
