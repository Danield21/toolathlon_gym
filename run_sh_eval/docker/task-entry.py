#!/usr/bin/env python3
"""Run the unchanged gym harness, enforcing its existing blackbox boundary."""
import os
from pathlib import Path
import subprocess
import sys
import time
import urllib.request

sys.path.insert(0, '/workspace')
from kimi_harness import kimi_main as harness

# Kimi 0.34 reports HTTP 4xx/5xx as provider.api_error. With no model progress
# this must be infrastructure failure, never a scored model failure.
harness._PROVIDER_ERR_PATTERNS += ('provider.api_error',)


def strict_mask():
    expected = [str(Path('/workspace') / p) for p in harness._BLACKBOX_DIRS + harness._BLACKBOX_FILES
                if (Path('/workspace') / p).exists()]
    expected += [p for p in harness._BLACKBOX_ABS_FILES if Path(p).is_file()]
    mounted = original_mask()
    missing = sorted(set(expected) - set(mounted))
    if missing:
        harness._unmask_blackbox(mounted)
        raise RuntimeError('Blackbox mounts failed; refusing evaluation: ' + ', '.join(missing))
    return mounted


original_mask = harness._mask_blackbox
harness._mask_blackbox = strict_mask


def start_woocommerce(task):
    import json
    config = json.loads((Path('/workspace/tasks/finalpool') / task / 'task_config.json').read_text())
    if 'woocommerce' not in str(config):
        return None
    log = open('/tmp/wc-rest.log', 'w')
    proc = subprocess.Popen(['node', '/opt/local_servers/woocommerce-mcp/dist/services/pg-rest-server.js'],
                            stdout=log, stderr=subprocess.STDOUT)
    log.close()
    for _ in range(60):
        if proc.poll() is not None:
            raise RuntimeError('WooCommerce mock failed; see wc-rest.log')
        try:
            with urllib.request.urlopen('http://127.0.0.1:8081/health', timeout=1) as response:
                if response.status == 200:
                    return proc
        except OSError:
            time.sleep(0.25)
    proc.terminate()
    raise RuntimeError('WooCommerce mock did not become ready')


def main():
    os.chdir('/workspace')
    task = os.environ['MSE_TASK']
    preset = os.environ['MSE_PRESET']
    service = start_woocommerce(task)
    try:
        if os.environ.get('MSE_SMOKE') == '1':
            import smoke
            smoke.main(harness, task)
        else:
            # Verify mount capability before any paid model requests.
            masked = strict_mask()
            harness._unmask_blackbox(masked)
            sys.argv = ['kimi_main.py', '--eval_config', f'/workspace/run_sh_eval/{preset}/eval_config.json',
                        '--task_dir', task, '--max_steps', os.environ['MAX_STEPS'], '--debug']
            harness.main()
    finally:
        if service is not None:
            service.terminate()
            service.wait(timeout=10)


if __name__ == '__main__':
    main()
