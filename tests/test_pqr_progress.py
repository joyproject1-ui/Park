"""보고서 작성 중 진행 조회(/api/progress) — 담당자가 '꽤 오래 작성 중' 이라며 멈춘 줄 알았다(2026-09)."""
import json
import os
import sys
import threading
import time
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pqr import server as sm            # noqa: E402
from pqr.engine import handwriting      # noqa: E402


def test_progress_endpoint_reports_running_and_idle(tmp_path):
    inp = tmp_path / "in"
    inp.mkdir()
    httpd = sm.serve(str(inp), port=0, out_dir=str(tmp_path / "out"), log=lambda *a: None)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    try:
        base = "http://127.0.0.1:%d" % httpd.server_port
        ws = httpd.RequestHandlerClass.workspace
        with urllib.request.urlopen(base + "/api/progress?product=X", timeout=5) as r:
            idle = json.loads(r.read().decode())
        assert idle == {"ok": True, "running": False}
        steps = ["  하나", "  시험일지 판독 중 2/4: a.pdf"]
        ws.progress["X"] = {"steps": steps, "started": time.time() - 65, "running": True}
        with urllib.request.urlopen(base + "/api/progress?product=X", timeout=5) as r:
            busy = json.loads(r.read().decode())
        assert busy["running"] is True
        assert busy["last"] == "시험일지 판독 중 2/4: a.pdf"
        assert busy["elapsed"] >= 65 and busy["count"] == 2
        assert busy["recent"] == ["하나", "시험일지 판독 중 2/4: a.pdf"]
    finally:
        httpd.shutdown()


def test_read_folder_logs_each_page(monkeypatch):
    lines = []
    monkeypatch.setattr(handwriting, "read_log", lambda p, specs, log, page_no=0: {"lot": os.path.basename(p), "year": "2025", "points": []})
    handwriting.read_folder(["/x/b.pdf", "/x/a.pdf"], None, lines.append)
    assert [l.strip() for l in lines] == ["시험일지 판독 중 1/2: b.pdf", "시험일지 판독 중 2/2: a.pdf"]


def test_dashboard_polls_progress_while_writing():
    html = open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                             "docs", "pqr", "index.html"), encoding="utf-8").read()
    assert "api/progress?product=" in html
    assert "watchProgress(code, button)" in html
    assert "clearInterval(watcher)" in html
