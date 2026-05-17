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

# Fixed loopback port so OBS Browser Source URL stays the same every run (see Readme).
WHEEL_HTML_SERVER_HOST = "127.0.0.1"
WHEEL_HTML_SERVER_PORT = 8765


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
    *,
    port: int | None = None,
    allow_ephemeral_fallback: bool = True,
) -> tuple[ThreadingHTTPServer | None, int, str | None]:
    """
    Serve HTML, JSON snapshot, and proxied cell images on loopback.

    Returns ``(server, port, warning)``. Tries ``port`` (default ``WHEEL_HTML_SERVER_PORT``) first.
    If that address is in use and ``allow_ephemeral_fallback``, binds an ephemeral port instead.
    """
    if not html_path.is_file():
        return None, 0, None
    handler = _make_handler(html_path, get_snapshot, get_image)
    preferred = WHEEL_HTML_SERVER_PORT if port is None else int(port)
    warning: str | None = None
    bind_ports: list[int] = [preferred]
    if allow_ephemeral_fallback and preferred != 0:
        bind_ports.append(0)
    last_err: OSError | None = None
    for bind_port in bind_ports:
        try:
            httpd = ThreadingHTTPServer((WHEEL_HTML_SERVER_HOST, bind_port), handler)
            actual = int(httpd.server_address[1])
            if bind_port == 0 and preferred != 0:
                warning = (
                    f"Port {preferred} is already in use on {WHEEL_HTML_SERVER_HOST}; "
                    f"using http://{WHEEL_HTML_SERVER_HOST}:{actual}/ instead. "
                    "Close the other program (or another Energy Break) to use the fixed OBS URL."
                )
            threading.Thread(
                target=httpd.serve_forever, name="WheelHtmlServer", daemon=True
            ).start()
            return httpd, actual, warning
        except OSError as e:
            last_err = e
            continue
    if last_err is not None:
        return None, 0, str(last_err)
    return None, 0, None


def guess_mime(path: Path) -> str:
    mt, _enc = mimetypes.guess_type(str(path))
    return mt or "application/octet-stream"
