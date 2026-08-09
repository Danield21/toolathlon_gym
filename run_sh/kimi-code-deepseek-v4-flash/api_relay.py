#!/usr/bin/env python3
"""Tiny bidirectional TCP relay for the eval API.

Compute nodes on this cluster have no external network, but the login node
does. This relay listens on 0.0.0.0:<port> and forwards every byte to the
upstream API server, so compute nodes can call the API via the login node's
internal IP.

Run:
    python3 api_relay.py [LISTEN_PORT] [UPSTREAM_HOST] [UPSTREAM_PORT]

Defaults: 19317  104.168.43.47  8317

Idempotent: if the port is already taken by a healthy relay, exits 0.
"""

import os
import socket
import sys
import threading
import time

DEFAULT_LISTEN_PORT = 19317
DEFAULT_UPSTREAM_HOST = "104.168.43.47"
DEFAULT_UPSTREAM_PORT = 8317


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


def _forward(src: socket.socket, dst: socket.socket) -> None:
    try:
        while True:
            data = src.recv(65536)
            if not data:
                break
            dst.sendall(data)
    except OSError:
        pass
    finally:
        for s in (src, dst):
            try:
                s.close()
            except OSError:
                pass


def _handle(client: socket.socket, upstream_addr: tuple) -> None:
    try:
        upstream = socket.create_connection(upstream_addr, timeout=10)
    except OSError as exc:
        sys.stderr.write(f"[relay] upstream {upstream_addr} connect failed: {exc}\n")
        try:
            client.close()
        except OSError:
            pass
        return
    threading.Thread(target=_forward, args=(client, upstream), daemon=True).start()
    threading.Thread(target=_forward, args=(upstream, client), daemon=True).start()


def main() -> int:
    listen_port = int(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_LISTEN_PORT
    upstream_host = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_UPSTREAM_HOST
    upstream_port = int(sys.argv[3]) if len(sys.argv) > 3 else DEFAULT_UPSTREAM_PORT
    upstream_addr = (upstream_host, upstream_port)

    if _port_in_use(listen_port):
        sys.stderr.write(f"[relay] port {listen_port} already in use, assuming healthy\n")
        return 0

    srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    srv.bind(("0.0.0.0", listen_port))
    srv.listen(256)
    sys.stderr.write(f"[relay] 0.0.0.0:{listen_port} -> {upstream_host}:{upstream_port}\n")
    sys.stderr.flush()

    while True:
        try:
            client, _ = srv.accept()
        except OSError:
            time.sleep(0.1)
            continue
        threading.Thread(target=_handle, args=(client, upstream_addr), daemon=True).start()


if __name__ == "__main__":
    raise SystemExit(main())
