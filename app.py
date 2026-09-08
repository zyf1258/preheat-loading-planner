"""独立预热炉排炉项目的本地 HTTP 服务。"""
from __future__ import annotations

import cgi
import json
import os
import sys
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse

import legacy_adapter


ROOT = Path(__file__).resolve().parent
STATIC_ROOT = ROOT / "static"
STATE_ROOT = Path(os.environ.get("PREHEAT_PLANNER_DATA_ROOT", ROOT / "data")).resolve()

# Keep the standalone planner's remembered billet lengths separate from the
# continuous dashboard and from any older standalone planner installation.
os.environ.setdefault("PREHEAT_DATA_ROOT", str(STATE_ROOT))


def _body(handler: SimpleHTTPRequestHandler) -> dict:
    length = int(handler.headers.get("Content-Length", "0") or 0)
    if not length:
        return {}
    return json.loads(handler.rfile.read(length).decode("utf-8"))


class Handler(SimpleHTTPRequestHandler):
    server_version = "PreheatFurnacePlannerStandalone/1.0"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(STATIC_ROOT), **kwargs)

    def log_message(self, fmt, *args):
        sys.stdout.write("[%s] %s\n" % (self.log_date_time_string(), fmt % args))

    def _json(self, data, status=200):
        raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _xlsx(self, payload: bytes):
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", 'attachment; filename="furnace-loading-map.xlsx"')
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def end_headers(self):
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "SAMEORIGIN")
        super().end_headers()

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/":
            self.path = "/loading-plan.html"
            return super().do_GET()
        if path == "/api/v1/health":
            return self._json({"ok": True, "version": "standalone-loading-planner", "data_root": str(STATE_ROOT)})
        if path == "/api/legacy/billet-lengths":
            return self._json({"lengths": legacy_adapter.read_lengths()})
        if path == "/legacy-planner.html":
            return self._html(legacy_adapter.render_legacy_page())
        return super().do_GET()

    def _html(self, html: str):
        raw = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _legacy_post(self, path):
        if path == "/api/legacy/billet-lengths":
            body = _body(self)
            return self._json(legacy_adapter.remember_length(body.get("code"), body.get("length_mm")))
        if path == "/api/legacy/billet-lengths/read":
            return self._json({"lengths": legacy_adapter.read_lengths()})
        if path == "/api/legacy/loading-plans":
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers.get("Content-Type", "")},
            )
            plan = form["plan_file"] if "plan_file" in form else None
            billet = form["billet_file"] if "billet_file" in form else None
            plan_bytes = plan.file.read() if plan is not None and getattr(plan, "filename", "") else b""
            billet_bytes = billet.file.read() if billet is not None and getattr(billet, "filename", "") else None
            return self._json(legacy_adapter.generate_loading_plan(plan_bytes, billet_bytes))
        if path == "/api/legacy/loading-plans/import":
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers.get("Content-Type", "")},
            )
            loading_map = form["loading_map_file"] if "loading_map_file" in form else None
            map_bytes = (
                loading_map.file.read()
                if loading_map is not None and getattr(loading_map, "filename", "")
                else b""
            )
            return self._json(legacy_adapter.import_loading_map(map_bytes))
        if path in {"/api/legacy/loading-plans/move", "/api/legacy/loading-plans/batch-move"}:
            body = _body(self)
            if path.endswith("/move"):
                result = legacy_adapter.move_item(body)
            else:
                result = legacy_adapter.batch_move(body)
            return self._json(result)
        if path == "/api/legacy/loading-plans/export":
            return self._xlsx(legacy_adapter.export_loading_map(_body(self)["result"]))
        return self._json({"ok": False, "message": "旧排程接口不存在"}, 404)

    def do_POST(self):
        try:
            path = urlparse(self.path).path
            if path.startswith("/api/legacy/"):
                return self._legacy_post(path)
            return self._json({"ok": False, "message": "独立排炉接口不存在"}, 404)
        except (ValueError, KeyError, json.JSONDecodeError) as exc:
            return self._json({"ok": False, "message": str(exc)}, 422)
        except Exception as exc:
            return self._json({"ok": False, "message": f"系统处理失败：{type(exc).__name__} - {exc}"}, 500)


def create_server(host="127.0.0.1", port=8770):
    return HTTPServer((host, port), Handler)


if __name__ == "__main__":
    port = int(os.environ.get("PREHEAT_PLANNER_PORT", "8770"))
    server = create_server(port=port)
    print(f"STANDALONE_PLANNER_READY:http://127.0.0.1:{server.server_port}", flush=True)
    server.serve_forever()
