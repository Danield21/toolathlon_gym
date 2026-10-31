#!/usr/bin/env python3
"""Exercise the actual DB, preprocessing, MCP transports and network boundary.

No model API calls, agent completion claims or evaluation scores are generated.
"""
import asyncio
from datetime import timedelta
import json
import os
from pathlib import Path
import socket
import struct
import subprocess
import sys
import urllib.request
from urllib.parse import urlsplit


def expect_blocked(host, port):
    try:
        with socket.create_connection((host, port), timeout=1):
            raise AssertionError(f'Unexpected outbound connection to {host}:{port}')
    except OSError:
        pass


def check_network():
    with urllib.request.urlopen(os.environ['MODEL_API_URL'] + '/v1/models', timeout=5) as response:
        assert json.load(response)['mse_smoke'], 'Local API gateway unreachable'
    expect_blocked(urlsplit(os.environ['MODEL_API_URL']).hostname, 8082)
    expect_blocked('1.1.1.1', 443)
    # Check the container-network host gateway as well as public egress.
    for line in Path('/proc/net/route').read_text().splitlines()[1:]:
        fields = line.split()
        if fields[1] == '00000000' and fields[2] != '00000000':
            host = socket.inet_ntoa(struct.pack('<I', int(fields[2], 16)))
            expect_blocked(host, 22)
    resolvers = [line.split()[1] for line in Path('/etc/resolv.conf').read_text().splitlines()
                 if line.startswith('nameserver ') and ':' not in line]
    packet = b'\x4d\x53\x01\x00\x00\x01\x00\x00\x00\x00\x00\x00\x07example\x03com\x00\x00\x01\x00\x01'
    for host in resolvers:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.settimeout(1)
            try:
                sock.sendto(packet, (host, 53))
                sock.recv(4096)
                raise AssertionError(f'External DNS is reachable through {host}')
            except OSError:
                pass


async def check_mcps(servers):
    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client
    counts = {}
    for name, cfg in servers.items():
        if cfg.get('type', 'stdio') != 'stdio':
            raise RuntimeError(f'Unexpected non-local MCP transport: {name}')
        params = StdioServerParameters(command=cfg['command'], args=cfg.get('args', []),
                                       env={**os.environ, **cfg.get('env', {})}, cwd=cfg.get('cwd'))
        async with stdio_client(params) as (read, write):
            async with ClientSession(read, write, read_timeout_seconds=timedelta(seconds=60)) as session:
                await session.initialize()
                result = await session.list_tools()
                assert result.tools, f'No tools from MCP server {name}'
                counts[name] = len(result.tools)
                if name == 'playwright_with_chunk':
                    result = await session.call_tool('browser_navigate', {'url': 'about:blank'})
                    assert not result.isError, f'Browser MCP could not launch Chromium: {result.content}'
    return counts


def main(harness, task):
    import psycopg2
    with psycopg2.connect('') as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT ready FROM mse_runtime.ready')
            assert cur.fetchone() == (True,)
            cur.execute('SHOW timezone')
            assert cur.fetchone() == ('UTC',)
    check_network()
    config = harness.TaskConfig.build(task, agent_short_name='smoke',
                                       global_task_config={'dump_path': './dumps/',
                                                           'max_steps_under_single_turn_mode': 1},
                                       single_turn_mode=True, cn_mode=False)
    workspace = harness.setup_workspace(config)
    assert harness.run_preprocess(config, True), 'Task preprocessing failed'
    mocks = harness._relocate_mock_servers(config)
    try:
        task_root = Path(config.task_root)
        home = task_root / '.kimi_home'
        harness.write_kimi_home(str(home), config, workspace, str(task_root / '.claim_done'), 1, rewrite={})
        mcp_file = home / 'mcp.json'
        assert not harness._check_mcp_servers_health(str(mcp_file)), 'MCP entrypoint check failed'
        counts_file = task_root / 'smoke-mcp-tools.json'
        mounted = harness._mask_blackbox()
        try:
            assert not list(Path('/workspace/tasks').iterdir()), 'Task sources are still visible'
            agent_env = {key: value for key, value in os.environ.items() if not key.upper().startswith('PG')}
            subprocess.run(['setpriv', '--bounding-set=-all', '--inh-caps=-all', '--ambient-caps=-all',
                            '--no-new-privs', '/opt/venv/bin/python', __file__,
                            str(mcp_file), str(counts_file)], check=True, env=agent_env, timeout=240)
        finally:
            harness._unmask_blackbox(mounted)
        counts = json.loads(counts_file.read_text())
        report = {'task': task, 'database': 'ok', 'preprocess': 'ok', 'mcp_tools': counts,
                  'network_isolation': 'ok', 'blackbox_mounts': 'ok', 'model_api_calls': 0}
        Path('/workspace/dumps/smoke.json').write_text(json.dumps(report, indent=2) + '\n')
        print(json.dumps(report, indent=2), flush=True)
    finally:
        harness._shutdown_mock_servers(mocks)


if __name__ == '__main__':
    # Run MCP startup under the same masks and capability boundary as Kimi.
    assert not list(Path('/workspace/tasks').iterdir()), 'Task sources are visible'
    assert 'CapBnd:\t0000000000000000' in Path('/proc/self/status').read_text()
    servers = json.loads(Path(sys.argv[1]).read_text())['mcpServers']
    counts = asyncio.run(asyncio.wait_for(check_mcps(servers), timeout=180))
    Path(sys.argv[2]).write_text(json.dumps(counts, indent=2) + '\n')
