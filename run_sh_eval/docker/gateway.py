#!/usr/bin/env python3
"""Container transport for the selected run_sh_eval API gateway implementation."""
import importlib.util
import hmac
from http.server import ThreadingHTTPServer
import json
import os
from pathlib import Path
import sys
import threading
from urllib.error import HTTPError
from urllib.parse import urlsplit
from urllib.request import Request, urlopen


def allowed_route(target):
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc or parsed.fragment:
        return None
    routes = {'/v1/chat/completions', '/v1/messages', '/v1/messages/count_tokens', '/v1/models'}
    if parsed.path not in routes:
        return None
    if parsed.query and not (parsed.query == 'beta=true' and parsed.path.startswith('/v1/messages')):
        return None
    return parsed.path


def handler_for(module):
    """Retain preset schema/body handling, including Kimi's beta Messages query."""
    class Handler(module.Handler):
        def handle_request(self):
            settings = self.server.settings
            if settings.get('MSE_SMOKE') == '1':
                if self.path == '/v1/models':
                    data = json.dumps({'object': 'list', 'data': [], 'mse_smoke': True}).encode()
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json')
                    self.send_header('Content-Length', str(len(data)))
                    self.end_headers()
                    self.wfile.write(data)
                else:
                    self.reply_error(404)
                return
            incoming = self.headers.get('Authorization', '').removeprefix('Bearer ') or self.headers.get('x-api-key', '')
            if not hmac.compare_digest(incoming, settings['GATEWAY_TOKEN']):
                self.reply_error(401)
                return
            route = allowed_route(self.path)
            if route is None:
                self.reply_error(404)
                return
            try:
                body = self.read_body() if self.command == 'POST' else None
                if body and route in ('/v1/chat/completions', '/v1/messages'):
                    body = module.rewrite(body, settings)
            except (ValueError, TypeError, AttributeError):
                self.reply_error(400)
                return
            base = settings['API_BASE_URL'].rstrip('/')
            # Keep the beta query on the fixed upstream; never accept arbitrary URLs.
            url = base + (self.path[3:] if urlsplit(base).path else self.path)
            headers = {'Content-Type': 'application/json', 'Accept-Encoding': 'identity'}
            if settings.get('KIMI_PROVIDER_TYPE') == 'anthropic':
                headers['x-api-key'] = settings['API_KEY']
                for key in ('anthropic-version', 'anthropic-beta'):
                    if self.headers.get(key):
                        headers[key] = self.headers[key]
                headers.setdefault('anthropic-version', '2023-06-01')
            else:
                headers['Authorization'] = 'Bearer ' + settings['API_KEY']
            started = False
            try:
                with urlopen(Request(url, data=body, headers=headers, method=self.command), timeout=300) as response:
                    self.send_response_only(response.status)
                    self.send_header('Content-Type', response.headers.get('Content-Type', 'application/json'))
                    self.send_header('Connection', 'close')
                    self.end_headers()
                    started = True
                    while chunk := response.read1(65536):
                        self.wfile.write(chunk)
                        self.wfile.flush()
            except HTTPError as exc:
                self.reply_error(exc.code)
                exc.close()
            except Exception:
                if not started:
                    self.reply_error(502)
            finally:
                self.close_connection = True

        do_GET = do_POST = handle_request

    return Handler


def main():
    preset = Path('/workspace/run_sh_eval') / os.environ['MSE_PRESET']
    sys.path.insert(0, str(preset))
    spec = importlib.util.spec_from_file_location('preset_gateway', preset / 'api_gateway.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    Handler = handler_for(module)
    server = ThreadingHTTPServer(('0.0.0.0', 8080), Handler)
    server.settings = dict(os.environ)
    if os.environ.get('MSE_SMOKE') == '1':
        # A known listening port provides a negative firewall control.
        canary = ThreadingHTTPServer(('0.0.0.0', 8082), Handler)
        canary.settings = server.settings
        threading.Thread(target=canary.serve_forever, daemon=True).start()
    print('MS-Eval API gateway ready', flush=True)
    server.serve_forever()


if __name__ == '__main__':
    main()
