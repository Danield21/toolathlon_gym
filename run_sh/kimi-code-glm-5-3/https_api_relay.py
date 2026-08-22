#!/usr/bin/env python3
"""Login-node HTTPS reverse proxy for Volcengine Ark.

Compute nodes have no external network. They speak plain HTTP OpenAI
(/v1/chat/completions) to this process; we CONNECT through the login-node
HTTP proxy and forward to:

    https://ark.cn-beijing.volces.com/api/plan/v3/chat/completions

Usage:
    python3 https_api_relay.py [LISTEN_PORT] [HTTP_PROXY_URL]

Defaults: 19320  http://127.0.0.1:7893
"""

from __future__ import annotations

import json
import os
import socket
import ssl
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

DEFAULT_LISTEN_PORT = 19320
DEFAULT_PROXY = "http://127.0.0.1:7893"
UPSTREAM_HOST = "ark.cn-beijing.volces.com"
UPSTREAM_PORT = 443
UPSTREAM_CHAT_PATH = "/api/plan/v3/chat/completions"
# Ark GLM output ceiling is 128000 decimal, not 128*1024. kimi-code emits
# max_tokens=131072 when KIMI_MAX_CONTEXT=262144; Ark returns 400.
ARK_MAX_OUTPUT_TOKENS = 128000
HOP_BY_HOP = {
    "connection",
    "keep-alive",
    "proxy-authenticate",
    "proxy-authorization",
    "te",
    "trailers",
    "transfer-encoding",
    "upgrade",
    "host",
}


def _log(msg: str) -> None:
    sys.stderr.write(f"[glm-relay] {msg}\n")
    sys.stderr.flush()


def _port_in_use(port: int) -> bool:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(0.5)
    try:
        s.bind(("0.0.0.0", port))
    except OSError:
        return True
    finally:
        s.close()
    return False


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


def _clamp_ark_token_fields(body: bytes) -> bytes:
    """Cap OpenAI output-token fields to Ark's GLM ceiling."""
    if not body:
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
        if value > ARK_MAX_OUTPUT_TOKENS:
            payload[key] = ARK_MAX_OUTPUT_TOKENS
            changed.append(f"{key}:{value}->{ARK_MAX_OUTPUT_TOKENS}")
    if not changed:
        return body
    _log("clamp " + " ".join(changed))
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")


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
    raw.settimeout(600)
    ctx = ssl.create_default_context()
    return ctx.wrap_socket(raw, server_hostname=UPSTREAM_HOST)


def _forward_chat(handler: BaseHTTPRequestHandler, proxy_host: str, proxy_port: int) -> None:
    length = int(handler.headers.get("Content-Length") or "0")
    body = handler.rfile.read(length) if length else b""
    body = _clamp_ark_token_fields(body)
    auth = handler.headers.get("Authorization") or ""
    ctype = handler.headers.get("Content-Type") or "application/json"
    request = (
        f"POST {UPSTREAM_CHAT_PATH} HTTP/1.1\r\n"
        f"Host: {UPSTREAM_HOST}\r\n"
        f"Authorization: {auth}\r\n"
        f"Content-Type: {ctype}\r\n"
        f"Content-Length: {len(body)}\r\n"
        f"Accept: {handler.headers.get('Accept') or '*/*'}\r\n"
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

        def _send_models(self) -> None:
            payload = {
                "object": "list",
                "data": [{"id": "glm-5-3-260801", "object": "model", "owned_by": "ark"}],
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
            if path in ("/v1/models", "/healthz", "/"):
                self._send_models()
                return
            self.send_error(404, "not proxied")

        def do_HEAD(self) -> None:  # noqa: N802
            self.do_GET()

        def do_POST(self) -> None:  # noqa: N802
            path = self.path.split("?", 1)[0]
            if path.rstrip("/") == "/v1/chat/completions":
                _forward_chat(self, proxy_host, proxy_port)
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
    httpd = ThreadingHTTPServer(("0.0.0.0", listen_port), handler)
    _log(
        f"0.0.0.0:{listen_port}/v1/chat/completions -> "
        f"https://{UPSTREAM_HOST}{UPSTREAM_CHAT_PATH} via {proxy_host}:{proxy_port}"
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
