#!/usr/bin/env python3
"""Docker/Podman backend for the existing run_sh_eval presets (Python 3.10+)."""
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import csv
import importlib.util
import json
import os
from pathlib import Path
import re
import secrets
import signal
import subprocess
import sys
import tempfile
import threading
import time
from urllib.parse import urlsplit

GYM = Path(__file__).resolve().parents[2]
ROOT = GYM.parent
GOOD = {'success', 'case_failed'}
AGENT_KEYS = {
    'MODEL_NAME', 'MAX_STEPS', 'KIMI_MAX_CONTEXT', 'KIMI_TASK_TIMEOUT_S',
    'KIMI_PROVIDER_TYPE', 'KIMI_PROVIDER_NAME', 'KIMI_MODEL_ALIAS',
    'KIMI_MODEL_TEMPERATURE', 'KIMI_MODEL_TOP_P', 'KIMI_MODEL_THINKING_EFFORT',
    'KIMI_CODE_SUBAGENT', 'KIMI_PLAN_FIRST',
}


def require(condition, message):
    if not condition:
        raise ValueError(message)


def load_preset(preset, single_agent=False):
    require(bool(re.fullmatch(r'[A-Za-z0-9_.-]+/(sa|spec|gen|oracle|plan1st)', preset)), 'Invalid preset')
    directory = GYM / 'run_sh_eval' / preset
    require((directory / 'launch.py').is_file(), f'Unknown preset: {preset}')
    spec = importlib.util.spec_from_file_location('mse_preset', directory / 'launch.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    # Reuse validation and all sampling defaults; Enroot preflight is not called.
    before = dict(os.environ)
    try:
        os.environ['PROJECT_ROOT'] = str(GYM)
        os.environ['ORACLE_TASK_DIR'] = str(ROOT / 'data/oracle166')
        env, tasks, plan = module.prepare()
    finally:
        os.environ.clear()
        os.environ.update(before)
    if single_agent or plan['variant'] == 'sa':
        env['KIMI_CODE_SUBAGENT'] = 'none'
        env.pop('KIMI_PLAN_FIRST', None)
        plan.update(subagent_roster='none', plan_first=False, paradigm='SA')
    else:
        plan['paradigm'] = 'MAD-SAE'
    plan['backend'] = 'containers'
    return env, tasks, plan


def classify(directory, rc, smoke=False):
    for marker, status in [('.preprocess_failed', 'preprocess_failed'),
                           ('.provider_invalid', 'provider_invalid'), ('.infra_failed', 'infra_failed')]:
        if any(directory.rglob(marker)):
            return status
    if smoke:
        return 'success' if rc == 0 and (directory / 'smoke.json').exists() else 'infra_failed'
    # A crashed process with no evaluation must never enter the model denominator.
    if rc not in (0, 1) or not any(directory.rglob('eval_res.json')):
        return 'infra_failed'
    return 'success' if rc == 0 else 'case_failed'


class Engine:
    def __init__(self, executable):
        require(executable in ('docker', 'podman'), 'Engine must be docker or podman')
        self.executable = executable

    def run(self, *args, check=True, timeout=120, **kwargs):
        result = subprocess.run([self.executable, *map(str, args)], capture_output=True,
                                text=True, timeout=timeout, **kwargs)
        if check and result.returncode:
            # Secret values never appear in argv; only the private env-file path does.
            raise RuntimeError(f'{self.executable} {args[0]} failed: {result.stderr.strip()[-3000:]}')
        return result

    def json(self, *args):
        return json.loads(self.run(*args).stdout)

    def create(self, args, env):
        for value in env.values():
            require('\n' not in str(value) and '\r' not in str(value), 'Environment values cannot contain newlines')
        with tempfile.NamedTemporaryFile(mode='w', prefix='mse-env-') as file:
            file.write(''.join(f'{key}={value}\n' for key, value in env.items()))
            file.flush()
            return self.run('create', '--env-file', file.name, *args).stdout.strip()


class Runner:
    def __init__(self, args, env, plan, output):
        self.args, self.env, self.plan, self.output = args, env, plan, output
        self.engine = Engine(args.engine)
        self.stop = threading.Event()
        self.lock = threading.Lock()
        self.resources = []
        self.egress = 'mse-' + secrets.token_hex(8) + '-egress'

    def prepare_network(self):
        # Own the outbound bridge as well: do not depend on the host's docker0.
        self.engine.run('network', 'create', '--label', 'org.ms-eval.run=' + self.output.name, self.egress)
        self.register('network', self.egress)

    def register(self, kind, name):
        with self.lock:
            self.resources.append((kind, name))

    def remove(self, kind, name):
        command = ('rm', '-f', '-v', name) if kind == 'container' else ('network', 'rm', name)
        try:
            result = self.engine.run(*command, check=False, timeout=30)
        except (OSError, subprocess.SubprocessError) as exc:
            print(f'[cleanup] could not remove {kind} {name}: {exc}', file=sys.stderr)
            return
        if result.returncode:
            print(f'[cleanup] could not remove {kind} {name}: {result.stderr.strip()}', file=sys.stderr)
        else:
            with self.lock:
                if (kind, name) in self.resources:
                    self.resources.remove((kind, name))

    def cleanup(self):
        with self.lock:
            resources = list(reversed(self.resources))
        for kind, name in resources:
            self.remove(kind, name)

    def wait_ready(self, container, command, timeout=180):
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline and not self.stop.is_set():
            result = self.engine.run('exec', container, *command, check=False, timeout=10)
            if result.returncode == 0:
                return
            state = self.engine.json('inspect', container)[0]['State']
            if not state.get('Running'):
                raise RuntimeError(f'{container} exited during startup')
            self.stop.wait(0.5)
        raise RuntimeError(f'{container} readiness timed out or run interrupted')

    def run_task(self, task, slot):
        if self.stop.is_set():
            return {'task': task, 'status': 'interrupted', 'exit_code': 130, 'output_dir': '', 'pg_port': 5432, 'duration_s': 0}
        start = time.monotonic()
        name = 'mse-' + secrets.token_hex(8)
        net, gateway, database, agent = (name + suffix for suffix in ('-net', '-api', '-db', '-agent'))
        out = self.output / task / (self.output.name + f'_slot{slot}')
        out.mkdir(parents=True)
        resources = []
        rc, status = 2, 'infra_failed'

        def track(kind, value):
            self.register(kind, value)
            resources.append((kind, value))

        try:
            print(f'[start] {task}', flush=True)
            self.engine.run('network', 'create', '--internal', '--label', 'org.ms-eval.run=' + self.output.name, net)
            track('network', net)
            token = secrets.token_hex(32)
            settings = {k: v for k, v in self.env.items() if k.startswith('GATEWAY_')}
            settings.update({k: self.env[k] for k in ('KIMI_PROVIDER_TYPE', 'KIMI_MODEL_THINKING_EFFORT')})
            settings.update(API_KEY=os.environ.get('API_KEY', ''), API_BASE_URL=os.environ.get('API_BASE_URL', ''),
                            GATEWAY_TOKEN=token, MSE_PRESET=self.args.preset, MSE_SMOKE=str(int(self.args.smoke)))
            # The gateway is the only container with outbound access and real credentials.
            for key in ('HTTPS_PROXY', 'HTTP_PROXY', 'NO_PROXY', 'https_proxy', 'http_proxy', 'no_proxy'):
                if key in os.environ:
                    settings[key] = os.environ[key]
            self.engine.create(['--name', gateway, '--label', 'org.ms-eval.run=' + self.output.name,
                                '--network', self.egress, '--read-only', '--cap-drop', 'ALL',
                                '--security-opt', 'no-new-privileges', '--user', '65534:65534',
                                '--tmpfs', '/tmp:rw,nosuid,nodev,size=32m', self.args.image,
                                '/opt/venv/bin/python', '/opt/ms-eval/gateway.py'], settings)
            track('container', gateway)
            self.engine.run('network', 'connect', net, gateway)
            self.engine.run('start', gateway)
            ip = self.engine.json('inspect', gateway)[0]['NetworkSettings']['Networks'][net]['IPAddress']
            require(bool(re.fullmatch(r'\d+\.\d+\.\d+\.\d+', ip)), 'Internal network needs IPv4')
            self.wait_ready(gateway, ['/opt/venv/bin/python', '-c',
                                     "import socket; socket.create_connection(('127.0.0.1',8080),2).close()"])
            if self.args.smoke:
                self.wait_ready(gateway, ['/opt/venv/bin/python', '-c',
                                         "import socket; socket.create_connection(('127.0.0.1',8082),2).close()"])
            self.engine.create(['--name', database, '--label', 'org.ms-eval.run=' + self.output.name,
                                '--network', net, '--cap-add', 'NET_ADMIN',
                                '--shm-size', '256m', self.args.db_image],
                               {'POSTGRES_USER': 'eigent', 'POSTGRES_PASSWORD': 'camel', 'PGPASSWORD': 'camel',
                                'POSTGRES_DB': 'toolathlon_gym', 'POSTGRES_INITDB_ARGS': '--encoding=UTF8 --locale=C',
                                'MSE_GATEWAY_IPV4': ip})
            track('container', database)
            self.engine.run('start', database)
            # TCP readiness avoids the temporary socket-only server during initdb.
            self.wait_ready(database, ['psql', '-h', '127.0.0.1', '-U', 'eigent', '-d', 'toolathlon_gym',
                                       '-v', 'ON_ERROR_STOP=1', '-tAc', 'SELECT ready FROM mse_runtime.ready'])
            agent_env = {k: v for k, v in self.env.items() if k in AGENT_KEYS}
            agent_env.update(PGHOST='127.0.0.1', PG_HOST='127.0.0.1', PGPORT='5432', PG_PORT='5432',
                             PGUSER='eigent', PG_USER='eigent', PGPASSWORD='camel', PG_PASSWORD='camel',
                             PGDATABASE='toolathlon_gym', PG_DATABASE='toolathlon_gym', PGTZ='UTC',
                             MODEL_API_URL=f'http://{ip}:8080', MODEL_API_KEY=token,
                             MSE_TASK=task, MSE_PRESET=self.args.preset, MSE_SMOKE=str(int(self.args.smoke)),
                             MCP_STDIO_TIMEOUT_MIN='90', WC_REST_PORT='8081',
                             NO_PROXY='*', no_proxy='*')
            if self.plan['variant'] == 'oracle':
                agent_env['KIMI_TASK_MD_DIR'] = '/workspace/tasks/oracle166'
            security = ['--cap-add', 'SYS_ADMIN', '--cap-drop', 'NET_RAW', '--security-opt', 'no-new-privileges']
            if self.args.apparmor:
                security += ['--security-opt', 'apparmor=unconfined']
            self.engine.create(['--name', agent, '--label', 'org.ms-eval.run=' + self.output.name,
                                '--network', 'container:' + database,
                                '--shm-size', self.args.shm_size, *security,
                                self.args.image, '/opt/venv/bin/python', '/opt/ms-eval/task-entry.py'], agent_env)
            track('container', agent)
            self.engine.run('start', agent)
            deadline = time.monotonic() + int(self.env['KIMI_TASK_TIMEOUT_S']) + 1200
            while not self.stop.is_set():
                state = self.engine.json('inspect', agent)[0]['State']
                if not state.get('Running'):
                    rc = int(state['ExitCode'])
                    break
                if time.monotonic() > deadline:
                    rc = 124
                    self.engine.run('stop', '--time', '10', agent, check=False)
                    break
                self.stop.wait(1)
            else:
                rc = 130
                self.engine.run('stop', '--time', '10', agent, check=False)
            # docker cp makes local files owned by the invoking user, also with remote contexts.
            self.engine.run('cp', agent + ':/workspace/dumps/.', str(out), timeout=300)
            status = classify(out, rc, self.args.smoke)
            if rc == 130:
                status = 'interrupted'
        except Exception as exc:
            (out / '.infra_failed').write_text(str(exc) + '\n')
            print(f'[error] {task}: {exc}', file=sys.stderr, flush=True)
        finally:
            for container, filename in [(agent, 'run.log'), (database, 'postgres.log'), (gateway, 'gateway.log')]:
                if ('container', container) in resources:
                    try:
                        result = self.engine.run('logs', container, check=False, timeout=30)
                        (out / filename).write_text(result.stdout + result.stderr)
                    except Exception as exc:
                        print(f'[logs] {container}: {exc}', file=sys.stderr)
            for kind, resource in reversed(resources):
                self.remove(kind, resource)
        row = {'task': task, 'status': status, 'exit_code': rc, 'output_dir': str(out),
               'pg_port': 5432, 'duration_s': int(time.monotonic() - start)}
        print(f'[done] {task}: {status} ({row["duration_s"]}s)', flush=True)
        return row


def preflight(args, env):
    engine = Engine(args.engine)
    engine.run('info')
    for image in (args.image, args.db_image):
        engine.run('image', 'inspect', image)
    reference = json.loads((ROOT / 'docker/locks/reference.json').read_text())
    db_meta = engine.json('image', 'inspect', args.db_image)[0]
    require(db_meta['Config'].get('Labels', {}).get('org.ms-eval.database-seed-sha256') ==
            reference['database_seed_sha256'], 'Database fixture differs from reference; rebuild with docker/build.sh')
    info = engine.json('info', '--format', '{{json .}}')
    args.apparmor = any('apparmor' in str(x).lower() for x in info.get('SecurityOptions', []))
    if args.engine == 'podman':
        host = info.get('host', {})
        args.apparmor = host.get('security', {}).get('apparmorEnabled', False)
        mappings = host.get('idMappings', {})
        if host.get('security', {}).get('rootless'):
            require(sum(x.get('size', 0) for x in mappings.get('uidmap', [])) > 1,
                    'Rootless Podman requires subuid/subgid ranges; a single-UID mapping cannot run PostgreSQL')
    metadata = engine.run('run', '--rm', '--network', 'none', args.image,
                          'cat', '/opt/ms-eval/source-commits.json').stdout
    commits = json.loads(metadata)
    for label, directory in [('gym', GYM), ('kimi', ROOT / 'kimi-code')]:
        actual = subprocess.check_output(['git', '-C', str(directory), 'rev-parse', 'HEAD'], text=True).strip()
        require(commits[label] == actual, f'{label} commit differs from image; rebuild with docker/build.sh')
        dirty = subprocess.check_output(['git', '-C', str(directory), 'diff', 'HEAD', '--name-only'], text=True)
        require(not dirty.strip(), f'{label} has tracked local edits; commit and rebuild before evaluation')
    if not args.smoke:
        require(bool(os.environ.get('API_KEY')), 'Set API_KEY in the preset config.env or --config file')
        url = urlsplit(os.environ.get('API_BASE_URL', ''))
        require(url.scheme in ('http', 'https') and url.hostname and not url.username and not url.password
                and not url.query and not url.fragment, 'API_BASE_URL must be an HTTP(S) base URL without credentials')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('preset', help='e.g. claude-sonnet-5/spec; reads toolathlon_gym/run_sh_eval/<preset>')
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument('--dry-run', action='store_true', help='inspect settings, no engine or API required')
    modes.add_argument('--check', action='store_true', help='check images and configuration, no API calls')
    modes.add_argument('--smoke', action='store_true', help='offline infrastructure test with real fixtures and MCPs')
    parser.add_argument('--task', action='append', help='run a manifest subset; omitted = all 166 cases')
    parser.add_argument('--single-agent', action='store_true', help='SA baseline, same model preset with no sub-agents')
    parser.add_argument('--engine', choices=['docker', 'podman'], default=os.environ.get('CONTAINER_ENGINE', 'docker'))
    parser.add_argument('--image', default=os.environ.get('MSE_IMAGE', 'ms-eval:local'))
    parser.add_argument('--db-image', default=os.environ.get('MSE_DB_IMAGE', 'ms-eval-db:local'))
    parser.add_argument('--shm-size', default='2g')
    parser.add_argument('--output', type=Path, default=os.environ.get('OUTPUT_ROOT'),
                        help='override the results parent (default: results/<model>/<variant>)')
    args = parser.parse_args()
    env, all_tasks, plan = load_preset(args.preset, args.single_agent)
    model, variant = args.preset.split('/')
    if args.single_agent:
        variant = 'sa'
    args.output = (args.output if args.output is not None else ROOT / 'results' / model / variant).resolve()
    plan['output_root'] = str(args.output)
    tasks = args.task or (['canvas-course-comparison'] if args.smoke else all_tasks)
    require(len(tasks) == len(set(tasks)) and set(tasks) <= set(all_tasks), 'Tasks must be unique members of the 166-case manifest')
    plan.update(selected_cases=len(tasks), smoke=args.smoke, full_suite=tasks == all_tasks)
    print(json.dumps(plan, indent=2), flush=True)
    if args.dry_run:
        return 0
    preflight(args, env)
    if args.check:
        print('Preflight passed; no API requests made.')
        return 0
    output = args.output.resolve() / (time.strftime('%Y%m%d-%H%M%S') + '-' + secrets.token_hex(3))
    output.mkdir(parents=True)
    (output / 'run.json').write_text(json.dumps(plan, indent=2) + '\n')
    (output / 'tasks.txt').write_text('\n'.join(tasks) + '\n')
    runner = Runner(args, env, plan, output)
    previous = {sig: signal.signal(sig, lambda *_: runner.stop.set()) for sig in (signal.SIGINT, signal.SIGTERM)}
    rows = []
    try:
        runner.prepare_network()
        with (output / f'summary_parallel_{output.name}.csv').open('w', newline='') as handle:
            writer = csv.DictWriter(handle, fieldnames=['task', 'status', 'exit_code', 'output_dir', 'pg_port', 'duration_s'])
            writer.writeheader()
            handle.flush()
            with ThreadPoolExecutor(max_workers=int(env['MAX_CONCURRENT'])) as pool:
                futures = [pool.submit(runner.run_task, task, i) for i, task in enumerate(tasks)]
                for future in as_completed(futures):
                    row = future.result()
                    rows.append(row)
                    writer.writerow(row)
                    handle.flush()
    finally:
        runner.cleanup()
        for sig, handler in previous.items():
            signal.signal(sig, handler)
    if not args.smoke:
        audit = GYM / 'scripts/audit_html_gen.py'
        result = subprocess.run([sys.executable, str(audit), str(output)], check=False)
        if result.returncode:
            print('Audit HTML generation failed; raw artifacts retained.', file=sys.stderr)
    print(f'Results: {output}', flush=True)
    return 130 if runner.stop.is_set() else (0 if all(row['status'] in GOOD for row in rows) else 2)


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except (ValueError, OSError, subprocess.SubprocessError, RuntimeError) as exc:
        print(f'Error: {exc}', file=sys.stderr)
        raise SystemExit(2)
