"""Local HTTP mirror for the HTML wheel UI (OBS Browser Source or any browser)."""

from __future__ import annotations

import json
import mimetypes
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Callable
from urllib.parse import parse_qs, unquote, urlparse

SnapshotFn = Callable[[], dict[str, Any]]
ImageFn = Callable[[str], tuple[bytes, str] | None]


def _make_handler(
    html_path: Path,
    get_snapshot: SnapshotFn,
    get_image: ImageFn,
) -> type[BaseHTTPRequestHandler]:
    class WheelHtmlHandler(BaseHTTPRequestHandler):
        protocol_version = "HTTP/1.1"

        def log_message(self, *_args: object) -> None:
            return

        def _send(self, code: int, body: bytes, content_type: str) -> None:
            self.send_response(code)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self) -> None:
            parsed = urlparse(self.path)
            path = parsed.path or "/"
            if path in ("/", "/index.html"):
                try:
                    data = html_path.read_bytes()
                except OSError:
                    self.send_error(500, "Could not read wheel_spin.html")
                    return
                self._send(200, data, "text/html; charset=utf-8")
                return
            if path == "/api/wheel.json":
                try:
                    snap = get_snapshot()
                    raw = json.dumps(snap, separators=(",", ":")).encode("utf-8")
                except Exception:
                    raw = b'{"mode":"banner","title":"Snapshot error","subtitle":""}'
                self._send(200, raw, "application/json; charset=utf-8")
                return
            if path == "/wheel/img":
                qs = parse_qs(parsed.query or "")
                sku = (qs.get("sku") or [""])[0]
                sku = unquote(sku)
                got = get_image(sku)
                if not got:
                    self.send_error(404)
                    return
                blob, ctype = got
                self._send(200, blob, ctype)
                return
            self.send_error(404)

    WheelHtmlHandler.__module__ = __name__
    return WheelHtmlHandler


def start_wheel_html_server(
    html_path: Path,
    get_snapshot: SnapshotFn,
    get_image: ImageFn,
) -> tuple[ThreadingHTTPServer | None, int]:
    """Bind loopback on an ephemeral port; serve HTML, JSON snapshot, and proxied cell images."""
    if not html_path.is_file():
        return None, 0
    handler = _make_handler(html_path, get_snapshot, get_image)
    httpd = ThreadingHTTPServer(("127.0.0.1", 0), handler)
    port = int(httpd.server_address[1])
    threading.Thread(target=httpd.serve_forever, name="WheelHtmlServer", daemon=True).start()
    return httpd, port


def guess_mime(path: Path) -> str:
    mt, _enc = mimetypes.guess_type(str(path))
    return mt or "application/octet-stream"
