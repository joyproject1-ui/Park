"""담당자용 로컬 서버 (`python -m pqr serve`).

대시보드를 브라우저로 열고, 평가항목 칸을 눌러 그 자리에서 자료를 올릴 수 있게 합니다.
올린 파일은 해당 **제품 폴더**에 저장되고 곧바로 다시 집계되어 화면에 반영됩니다.

표준 라이브러리(http.server)만 사용하며, 기본적으로 이 PC(127.0.0.1)에서만 접속됩니다.
"""

import json
import os
import posixpath
import re
import shutil
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from . import build as build_module
from . import report as report_module

MAX_UPLOAD = 64 * 1024 * 1024          # 파일 하나당 64 MB
ALLOWED_SUFFIXES = (".csv", ".tsv", ".txt", ".xlsx", ".xlsm")
_UNSAFE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


class UploadError(Exception):
    """사용자에게 그대로 보여줄 수 있는 업로드 오류."""


def safe_filename(name):
    """경로 요소를 제거하고 확장자를 확인합니다."""
    base = os.path.basename(str(name or "").replace("\\", "/")).strip()
    base = _UNSAFE.sub("_", base).lstrip(".")
    if not base:
        raise UploadError("파일 이름이 비어 있습니다.")
    if not base.lower().endswith(ALLOWED_SUFFIXES):
        raise UploadError("올릴 수 있는 형식은 %s 입니다." % ", ".join(ALLOWED_SUFFIXES))
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
        except UploadError as error:
            return self._json(400, {"ok": False, "error": str(error)})
        except Exception as error:               # 서버가 죽지 않도록 오류를 그대로 알려 줍니다.
            return self._json(500, {"ok": False, "error": "%s: %s"
                                    % (type(error).__name__, error)})
        return self._send(404, "찾을 수 없습니다", "text/plain; charset=utf-8")

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
