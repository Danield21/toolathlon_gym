#!/usr/bin/env python3
"""Local authenticated gateway. Credentials stay outside task containers."""
import hmac
import json
import os
from pathlib import Path
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.error import HTTPError
from urllib.request import Request, urlopen
from urllib.parse import urlsplit

from schema import break_json_schema_cycles, _strip_remaining_refs


def rewrite(body, settings):
    payload = json.loads(body)
    if not isinstance(payload, dict):
        return body
    cap = int(settings.get("GATEWAY_MAX_OUTPUT_TOKENS", "0"))
    for key in ("max_tokens", "max_completion_tokens"):
        value = payload.get(key)
        if cap and isinstance(value, (int, float)) and not isinstance(value, bool) and value > cap:
            payload[key] = cap
    if settings.get("KIMI_PROVIDER_TYPE") != "anthropic":
        effort = settings.get("KIMI_MODEL_THINKING_EFFORT", "")
        if effort:
            payload["reasoning_effort"] = effort
        if settings.get("GATEWAY_ENABLE_THINKING") == "1":
            payload["enable_thinking"] = True
        if settings.get("GATEWAY_BREAK_SCHEMA_CYCLES") == "1":
            for key in ("tools", "functions"):
                for tool in payload.get(key, []) or []:
                    fn = tool.get("function", tool)
                    if isinstance(fn.get("parameters"), dict):
                        fn["parameters"] = _strip_remaining_refs(break_json_schema_cycles(fn["parameters"]))
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode()


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *args):
        pass

    def reply_error(self, status):
        body = json.dumps({"error": {"type": "relay_error", "message": f"API request failed (HTTP {status})"}}).encode()
        self.send_response_only(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Connection", "close")
        self.end_headers()
        self.wfile.write(body)
        self.close_connection = True

    def read_body(self):
        if "chunked" not in self.headers.get("Transfer-Encoding", "").lower():
            length = int(self.headers.get("Content-Length", "0"))
            if not 0 <= length <= 64 * 1024 * 1024:
                raise ValueError("Invalid body length")
            body = self.rfile.read(length)
            if len(body) != length:
                raise ValueError("Incomplete body")
            return body
        chunks, total = [], 0
        while True:
            size = int(self.rfile.readline(8192).split(b";", 1)[0].strip(), 16)
            total += size
            if size < 0 or total > 64 * 1024 * 1024:
                raise ValueError("Invalid chunk size")
            if size == 0:
                while True:
                    trailer = self.rfile.readline(8192)
                    if trailer in (b"\r\n", b"\n", b""):
                        return b"".join(chunks)
            chunk = self.rfile.read(size)
            if len(chunk) != size or self.rfile.read(2) != b"\r\n":
                raise ValueError("Incomplete chunk")
            chunks.append(chunk)

    def handle_request(self):
        settings = self.server.settings
        incoming = self.headers.get("Authorization", "").removeprefix("Bearer ") or self.headers.get("x-api-key", "")
        if not hmac.compare_digest(incoming, settings["GATEWAY_TOKEN"]):
            self.reply_error(401)
            return
        routes = ("/v1/chat/completions", "/v1/messages", "/v1/messages/count_tokens", "/v1/models")
        if self.path not in routes:
            self.reply_error(404)
            return
        try:
            body = self.read_body() if self.command == "POST" else None
            if body and self.path in ("/v1/chat/completions", "/v1/messages"):
                body = rewrite(body, settings)
        except (ValueError, TypeError, AttributeError):
            self.reply_error(400)
            return
        base = settings["API_BASE_URL"].rstrip("/")
        # A non-empty base path is an explicit API prefix (including non-/v1 APIs).
        url = base + (self.path[3:] if urlsplit(base).path else self.path)
        headers = {"Content-Type": "application/json", "Accept-Encoding": "identity"}
        if settings.get("KIMI_PROVIDER_TYPE") == "anthropic":
            headers["x-api-key"] = settings["API_KEY"]
            for key in ("anthropic-version", "anthropic-beta"):
                if self.headers.get(key):
                    headers[key] = self.headers[key]
            headers.setdefault("anthropic-version", "2023-06-01")
        else:
            headers["Authorization"] = "Bearer " + settings["API_KEY"]
        started = False
        try:
            with urlopen(Request(url, data=body, headers=headers, method=self.command), timeout=300) as response:
                self.send_response_only(response.status)
                self.send_header("Content-Type", response.headers.get("Content-Type", "application/json"))
                self.send_header("Connection", "close")
                self.end_headers()
                started = True
                while chunk := response.read1(65536):
                    self.wfile.write(chunk)
                    self.wfile.flush()
        except HTTPError as exc:
            # Upstream error bodies can contain account IDs, endpoints or credentials.
            self.reply_error(exc.code)
            exc.close()
        except Exception:
            if not started:
                self.reply_error(502)
        finally:
            self.close_connection = True

    do_POST = do_GET = handle_request


def main():
    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    server.settings = dict(os.environ)
    ready = Path(sys.argv[1])
    temporary = ready.with_suffix(".tmp")
    temporary.write_text(f"http://127.0.0.1:{server.server_port}")
    temporary.replace(ready)
    server.serve_forever()


if __name__ == "__main__":
    main()
