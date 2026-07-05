#!/usr/bin/env python3
"""Local HTTP server for the PPTX lint/fix studio.

Zero third-party deps (stdlib ``http.server`` only). Run:

    python3 studio/server.py            # http://127.0.0.1:8765
    PORT=9000 python3 studio/server.py

Routes
  GET  /                         -> static/index.html
  GET  /<asset>                  -> static/<asset>
  POST /api/upload               raw octet-stream PPTX body, X-Filename header
                                 -> session + grouped lint findings
  POST /api/apply                {session_id, selections:[{slide_index,check}]}
                                 -> applied/skipped FixAction summary
  GET  /api/download?session_id= -> fixed.pptx attachment
  POST /api/render               {session_id, which:'source'|'fixed'}
                                 -> on-demand slide PNG urls
  GET  /api/render/<sid>/<which>/<slide-NN.png> -> PNG bytes
"""

from __future__ import annotations

import json
import os
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs, unquote

import engine

STUDIO_DIR = Path(__file__).resolve().parent
STATIC_DIR = STUDIO_DIR / "static"

MAX_UPLOAD = 200 * 1024 * 1024  # 200 MB

CONTENT_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".png": "image/png",
    ".svg": "image/svg+xml",
    ".json": "application/json; charset=utf-8",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
}


class Handler(BaseHTTPRequestHandler):
    server_version = "PptxStudio/0.1"

    # ---- helpers ---------------------------------------------------------

    def _send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_error_json(self, message, status=400):
        self._send_json({"error": message}, status=status)

    def _send_bytes(self, data: bytes, content_type: str, status=200, extra=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _read_body(self) -> bytes:
        length = int(self.headers.get("Content-Length", 0))
        if length > MAX_UPLOAD:
            raise ValueError("upload too large")
        return self.rfile.read(length) if length else b""

    def log_message(self, fmt, *args):  # quieter console
        return

    # ---- routing ---------------------------------------------------------

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            if path == "/" or path == "":
                return self._serve_static("index.html")
            if path == "/api/download":
                return self._download(parse_qs(parsed.query))
            if path.startswith("/api/render/"):
                return self._render_png(path)
            if path.startswith("/api/"):
                return self._send_error_json("unknown endpoint", 404)
            return self._serve_static(path.lstrip("/"))
        except Exception as exc:  # noqa: BLE001
            traceback.print_exc()
            return self._send_error_json(str(exc), 500)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            if path == "/api/upload":
                return self._upload()
            if path == "/api/apply":
                return self._apply()
            if path == "/api/render":
                return self._render()
            return self._send_error_json("unknown endpoint", 404)
        except FileNotFoundError as exc:
            return self._send_error_json(str(exc), 404)
        except ValueError as exc:
            return self._send_error_json(str(exc), 400)
        except Exception as exc:  # noqa: BLE001
            traceback.print_exc()
            return self._send_error_json(str(exc), 500)

    # ---- static ----------------------------------------------------------

    def _serve_static(self, rel: str):
        rel = unquote(rel)
        target = (STATIC_DIR / rel).resolve()
        if STATIC_DIR not in target.parents and target != STATIC_DIR:
            return self._send_error_json("forbidden", 403)
        if not target.is_file():
            return self._send_error_json("not found", 404)
        ctype = CONTENT_TYPES.get(target.suffix, "application/octet-stream")
        self._send_bytes(target.read_bytes(), ctype)

    # ---- api -------------------------------------------------------------

    def _upload(self):
        data = self._read_body()
        if not data:
            return self._send_error_json("empty upload", 400)
        filename = self.headers.get("X-Filename", "deck.pptx")
        filename = unquote(filename)
        result = engine.create_session(filename, data)
        self._send_json(result)

    def _apply(self):
        payload = json.loads(self._read_body() or b"{}")
        session_id = payload.get("session_id")
        selections = payload.get("selections", [])
        if not session_id:
            return self._send_error_json("session_id required", 400)
        summary = engine.apply_session(session_id, selections)
        self._send_json(summary)

    def _download(self, query):
        session_id = (query.get("session_id") or [None])[0]
        if not session_id:
            return self._send_error_json("session_id required", 400)
        path = engine.fixed_path(session_id)
        if not path.exists():
            return self._send_error_json("no fixed.pptx; apply first", 404)
        disp = 'attachment; filename="fixed.pptx"'
        self._send_bytes(
            path.read_bytes(),
            CONTENT_TYPES[".pptx"],
            extra={"Content-Disposition": disp},
        )

    def _render(self):
        payload = json.loads(self._read_body() or b"{}")
        session_id = payload.get("session_id")
        which = payload.get("which", "source")
        if not session_id:
            return self._send_error_json("session_id required", 400)
        result = engine.render_session(session_id, which)
        self._send_json(result)

    def _render_png(self, path: str):
        # /api/render/<sid>/<which>/<slide-NN.png>
        parts = path[len("/api/render/"):].split("/")
        if len(parts) != 3:
            return self._send_error_json("bad render path", 400)
        sid, which, name = parts
        png = engine.render_png_path(sid, which, unquote(name))
        if not png.is_file():
            return self._send_error_json("png not found", 404)
        self._send_bytes(png.read_bytes(), CONTENT_TYPES[".png"])


def main():
    engine.SESSIONS_DIR.mkdir(parents=True, exist_ok=True)
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8765"))
    httpd = ThreadingHTTPServer((host, port), Handler)
    print(f"PPTX lint/fix studio  ->  http://{host}:{port}")
    print("Ctrl-C to stop")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nbye")


if __name__ == "__main__":
    main()
