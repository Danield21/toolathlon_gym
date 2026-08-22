#!/usr/bin/env python3
"""Login-node HTTPS reverse proxy for official DeepSeek.

Compute nodes have no external network. They speak plain HTTP OpenAI
(/v1/chat/completions, /v1/models) to this process; we CONNECT through the
login-node HTTP proxy and forward to:

    https://api.deepseek.com/v1/...

Usage:
    python3 https_deepseek_relay.py [LISTEN_PORT] [HTTP_PROXY_URL]

Defaults: 19325  http://127.0.0.1:7893
"""

from __future__ import annotations

import json
import os
import socket
import ssl
import sys
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

DEFAULT_LISTEN_PORT = 19325
DEFAULT_PROXY = "http://127.0.0.1:7893"
UPSTREAM_HOST = os.environ.get("DEEPSEEK_UPSTREAM_HOST", "api.deepseek.com")
UPSTREAM_PORT = int(os.environ.get("DEEPSEEK_UPSTREAM_PORT", "443"))
# GLM-style gateways (Ark / llmapi.blsc.cn) reject kimi's 131072 max_tokens.
MAX_OUTPUT_TOKENS = int(os.environ.get("MAX_OUTPUT_TOKENS", "0") or "0")
HOP_BY_HOP = {
    "connection",
    "keep-alive",
    "proxy-authenticate",
    "proxy-authorization",
    "te",
    "trailers",
    "upgrade",
    "host",
}


def _log(msg: str) -> None:
    sys.stderr.write(f"[ds-relay] {msg}\n")
    sys.stderr.flush()


def _port_in_use(port: int) -> bool:
    """True only if something is actually accepting on the port."""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(0.5)
    try:
        s.connect(("127.0.0.1", port))
    except OSError:
        return False
    finally:
        s.close()
    return True


def _read_until_double_crlf(sock: socket.socket, leftover: bytes = b"") -> tuple[bytes, bytes]:
    buf = leftover
    while b"\r\n\r\n" not in buf:
        chunk = sock.recv(65536)
        if not chunk:
            break
        buf += chunk
        if len(buf) > 1024 * 1024:
            raise OSError("HTTP header too large")
    header, _, rest = buf.partition(b"\r\n\r\n")
    return header, rest


def _connect_upstream(proxy_host: str, proxy_port: int) -> ssl.SSLSocket:
    raw = socket.create_connection((proxy_host, proxy_port), timeout=30)
    raw.settimeout(30)
    req = (
        f"CONNECT {UPSTREAM_HOST}:{UPSTREAM_PORT} HTTP/1.1\r\n"
        f"Host: {UPSTREAM_HOST}:{UPSTREAM_PORT}\r\n"
        f"Proxy-Connection: keep-alive\r\n"
        f"\r\n"
    ).encode("ascii")
    raw.sendall(req)
    header, leftover = _read_until_double_crlf(raw)
    status = header.split(b"\r\n", 1)[0].decode("latin1", "replace")
    if b" 200" not in header.split(b"\r\n", 1)[0]:
        raw.close()
        raise OSError(f"proxy CONNECT failed: {status!r}")
    if leftover:
        pass
    raw.settimeout(7200)
    ctx = ssl.create_default_context()
    return ctx.wrap_socket(raw, server_hostname=UPSTREAM_HOST)


def _clamp_output_tokens(body: bytes) -> bytes:
    """Cap max_tokens when the upstream GLM ceiling is below kimi's default."""
    if not body or MAX_OUTPUT_TOKENS <= 0:
        return body
    try:
        payload = json.loads(body)
    except (json.JSONDecodeError, UnicodeDecodeError):
        return body
    if not isinstance(payload, dict):
        return body
    changed = []
    for key in ("max_tokens", "max_completion_tokens"):
        raw = payload.get(key)
        if not isinstance(raw, (int, float)) or isinstance(raw, bool):
            continue
        value = int(raw)
        if value > MAX_OUTPUT_TOKENS:
            payload[key] = MAX_OUTPUT_TOKENS
            changed.append(f"{key}:{value}->{MAX_OUTPUT_TOKENS}")
    if not changed:
        return body
    _log("clamp " + " ".join(changed))
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")


def _forward(
    handler: BaseHTTPRequestHandler,
    proxy_host: str,
    proxy_port: int,
    method: str,
    upstream_path: str,
    body: bytes,
) -> None:
    auth = handler.headers.get("Authorization") or ""
    ctype = handler.headers.get("Content-Type") or "application/json"
    accept = handler.headers.get("Accept") or "*/*"
    request = (
        f"{method} {upstream_path} HTTP/1.1\r\n"
        f"Host: {UPSTREAM_HOST}\r\n"
        f"Authorization: {auth}\r\n"
        f"Content-Type: {ctype}\r\n"
        f"Content-Length: {len(body)}\r\n"
        f"Accept: {accept}\r\n"
        f"Connection: close\r\n"
        f"\r\n"
    ).encode("latin1") + body

    upstream = None
    t0 = time.time()
    try:
        upstream = _connect_upstream(proxy_host, proxy_port)
        upstream.sendall(request)
        header_bytes, rest = _read_until_double_crlf(upstream)
        if not header_bytes:
            raise OSError("empty upstream response")
        lines = header_bytes.split(b"\r\n")
        status_line = lines[0].decode("latin1", "replace")
        parts = status_line.split(" ", 2)
        code = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 502
        handler.send_response(code)
        for raw in lines[1:]:
            if not raw or b":" not in raw:
                continue
            k, _, v = raw.partition(b":")
            name = k.decode("latin1").strip()
            if name.lower() in HOP_BY_HOP:
                continue
            handler.send_header(name, v.decode("latin1").strip())
        handler.send_header("Connection", "close")
        handler.end_headers()
        if rest:
            handler.wfile.write(rest)
        while True:
            chunk = upstream.recv(65536)
            if not chunk:
                break
            handler.wfile.write(chunk)
        try:
            handler.wfile.flush()
        except OSError:
            pass
        _log(
            f"{handler.command} {handler.path} -> {code} "
            f"bytes_in={len(body)} {time.time() - t0:.2f}s"
        )
    except Exception as exc:
        _log(f"forward error {handler.path}: {exc}")
        try:
            msg = json.dumps({"error": {"message": f"relay: {exc}", "type": "relay_error"}}).encode()
            handler.send_response(502)
            handler.send_header("Content-Type", "application/json")
            handler.send_header("Content-Length", str(len(msg)))
            handler.send_header("Connection", "close")
            handler.end_headers()
            handler.wfile.write(msg)
        except Exception:
            pass
    finally:
        if upstream is not None:
            try:
                upstream.close()
            except OSError:
                pass


def make_handler(proxy_host: str, proxy_port: int):
    class Handler(BaseHTTPRequestHandler):
        protocol_version = "HTTP/1.1"

        def log_message(self, fmt: str, *args) -> None:
            _log("%s - %s" % (self.address_string(), fmt % args))

        def _send_health(self) -> None:
            payload = {
                "object": "list",
                "data": [{"id": "deepseek-v4-flash", "object": "model", "owned_by": "deepseek"}],
            }
            body = json.dumps(payload).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Connection", "close")
            self.end_headers()
            if self.command != "HEAD":
                self.wfile.write(body)

        def do_GET(self) -> None:  # noqa: N802
            path = self.path.split("?", 1)[0]
            if path in ("/healthz", "/"):
                self._send_health()
                return
            if path == "/v1/models":
                _forward(self, proxy_host, proxy_port, "GET", "/v1/models", b"")
                return
            self.send_error(404, "not proxied")

        def do_HEAD(self) -> None:  # noqa: N802
            self.do_GET()

        def do_POST(self) -> None:  # noqa: N802
            path = self.path.split("?", 1)[0]
            if path.rstrip("/") == "/v1/chat/completions":
                length = int(self.headers.get("Content-Length") or "0")
                body = self.rfile.read(length) if length else b""
                _forward(
                    self,
                    proxy_host,
                    proxy_port,
                    "POST",
                    "/v1/chat/completions",
                    _clamp_output_tokens(body),
                )
                return
            self.send_error(404, "not proxied")

    return Handler


def main() -> int:
    listen_port = int(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_LISTEN_PORT
    proxy_url = sys.argv[2] if len(sys.argv) > 2 else os.environ.get("HTTPS_RELAY_PROXY", DEFAULT_PROXY)
    parsed = urlparse(proxy_url)
    proxy_host = parsed.hostname or "127.0.0.1"
    proxy_port = parsed.port or 7893

    if _port_in_use(listen_port):
        _log(f"port {listen_port} already in use, assuming healthy")
        return 0

    handler = make_handler(proxy_host, proxy_port)

    class RelayHTTPServer(ThreadingHTTPServer):
        def __init__(self, *args, **kwargs):
            self.request_queue_size = int(os.environ.get("RELAY_BACKLOG", "256"))
            super().__init__(*args, **kwargs)

    httpd = RelayHTTPServer(("0.0.0.0", listen_port), handler)
    _log(
        f"0.0.0.0:{listen_port}/v1/* -> "
        f"https://{UPSTREAM_HOST} via {proxy_host}:{proxy_port}"
    )
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
