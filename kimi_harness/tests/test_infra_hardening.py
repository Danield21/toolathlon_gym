import os
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from kimi_harness import kimi_main, local_tools_server, mcp_json_gen


class KimiInfraHardeningTests(unittest.TestCase):
    def test_yahoo_finance_mcp_runs_from_its_server_directory(self):
        project_root = Path(__file__).resolve().parents[2]
        local_servers = project_root / "local_servers"

        servers = mcp_json_gen.build_mcp_servers(
            ["yahoo-finance"],
            agent_workspace="/tmp/agent-workspace",
            config_dir=str(project_root / "configs/mcp_servers"),
            local_servers_path=str(local_servers),
        )

        yahoo = servers["yahoo-finance"]
        self.assertEqual(yahoo["command"], "uv")
        self.assertEqual(yahoo["args"], ["run", "server.py"])
        self.assertEqual(
            yahoo["cwd"],
            str(local_servers / "yahoo-finance-mcp"),
        )

    def test_agent_image_build_locks_and_syncs_yahoo_finance_project(self):
        project_root = Path(__file__).resolve().parents[2]
        script = project_root / "scripts" / "enroot_build_agent.sh"
        text = script.read_text(encoding="utf-8")

        self.assertIn("YAHOO_FINANCE_MCP=/opt/local_servers/yahoo-finance-mcp", text)
        self.assertIn("uv lock && uv sync", text)
        self.assertIn("$YAHOO_FINANCE_MCP/.venv/bin/python", text)
        self.assertIn("Yahoo Finance MCP local shim imports", text)

    def test_python_execute_uses_direct_python_without_project_discovery(self):
        with tempfile.TemporaryDirectory() as workspace:
            local_tools_server.WORKSPACE = workspace
            completed = subprocess.CompletedProcess(
                args=[],
                returncode=0,
                stdout="ok\n",
                stderr="",
            )

            with mock.patch.dict(os.environ, {"PYTHON_EXECUTE_BIN": "/opt/venv/bin/python3"}):
                with mock.patch("kimi_harness.local_tools_server.subprocess.run", return_value=completed) as run:
                    output = local_tools_server._python_execute("print('ok')", filename="probe.py")

            self.assertIn("ok", output)
            args, kwargs = run.call_args
            self.assertEqual(args[0][0], "/opt/venv/bin/python3")
            self.assertTrue(args[0][1].endswith("/.python_tmp/probe.py"))
            self.assertFalse(kwargs.get("shell", False))

    def test_python_execute_strips_credentials_and_pg_keys_from_child_env(self):
        # Audit §A.3: agent-authored Python must not see MODEL_API_KEY or any
        # PG* credential, even though the kimi-code CLI parent needs them.
        with tempfile.TemporaryDirectory() as workspace:
            local_tools_server.WORKSPACE = workspace
            completed = subprocess.CompletedProcess(
                args=[], returncode=0, stdout="", stderr="",
            )
            secrets_env = {
                "MODEL_API_KEY": "sk-super-secret",
                "MODEL_API_URL": "https://model.example.com",
                "PGPASSWORD": "camel",
                "PGUSER": "eigent",
                "PGHOST": "127.0.0.1",
                "PGPORT": "5432",
                "PGDATABASE": "toolathlon_gym",
                "PG_HOST": "127.0.0.1",
                "PG_PASSWORD": "camel",
                "PATH": "/usr/bin:/bin",
                "KIMI_DISABLE_BOUNDARY": "",
            }
            with mock.patch.dict(os.environ, secrets_env, clear=False):
                with mock.patch("kimi_harness.local_tools_server.subprocess.run", return_value=completed) as run:
                    local_tools_server._python_execute("pass", filename="probe.py")
            child_env = run.call_args.kwargs["env"]
            # Secrets are gone …
            for secret in ("MODEL_API_KEY", "MODEL_API_URL",
                           "PGPASSWORD", "PGUSER", "PGHOST", "PGPORT", "PGDATABASE",
                           "PG_HOST", "PG_PASSWORD"):
                self.assertNotIn(secret, child_env, f"{secret} leaked into python_execute child env")
            # … but benign vars survive.
            self.assertIn("PATH", child_env)

    def test_python_execute_keeps_secrets_when_boundary_disabled(self):
        # The KIMI_DISABLE_BOUNDARY escape hatch must leave the env intact so
        # local debugging with real credentials still works.
        with tempfile.TemporaryDirectory() as workspace:
            local_tools_server.WORKSPACE = workspace
            completed = subprocess.CompletedProcess(args=[], returncode=0, stdout="", stderr="")
            with mock.patch.dict(os.environ, {
                "MODEL_API_KEY": "sk-super-secret",
                "PGPASSWORD": "camel",
                "KIMI_DISABLE_BOUNDARY": "1",
            }, clear=False):
                with mock.patch("kimi_harness.local_tools_server.subprocess.run", return_value=completed) as run:
                    local_tools_server._python_execute("pass", filename="probe.py")
            child_env = run.call_args.kwargs["env"]
            self.assertEqual(child_env.get("MODEL_API_KEY"), "sk-super-secret")
            self.assertEqual(child_env.get("PGPASSWORD"), "camel")

    def test_launch_kimi_strips_pg_credentials_from_agent_environment(self):
        # Audit §A.3 / security-boundary: the kimi-code CLI process must not
        # carry backing-DB credentials. MCP servers get them via mcp.json.
        captured = {}

        class _FakeProc:
            def __init__(self):
                self.returncode = None
            def poll(self):
                return self.returncode
            def terminate(self):
                self.returncode = 0
            def wait(self, timeout=None):
                return 0

        def fake_popen(*args, **kwargs):
            captured["env"] = dict(kwargs.get("env", {}))
            return _FakeProc()

        secrets_env = {
            "PGPASSWORD": "camel",
            "PGUSER": "eigent",
            "PGHOST": "127.0.0.1",
            "PGPORT": "5432",
            "PGDATABASE": "toolathlon_gym",
            "PG_HOST": "127.0.0.1",
            "PG_PASSWORD": "camel",
            "MODEL_API_KEY": "sk-model",
            "PATH": "/usr/bin:/bin",
        }
        with mock.patch.dict(os.environ, secrets_env, clear=False):
            with mock.patch("kimi_harness.kimi_main.subprocess.Popen", side_effect=fake_popen), \
                 mock.patch("kimi_harness.kimi_main.time.sleep", return_value=None), \
                 mock.patch("kimi_harness.kimi_main.os.path.exists", return_value=True):
                rc, done = kimi_main.launch_kimi(
                    task_str="t", agentfile="/tmp/a.md", home="/tmp/h",
                    stream_path="/tmp/s", workspace="/tmp/w", marker="/tmp/m",
                    timeout_s=1, debug=False,
                )

        agent_env = captured["env"]
        for pg_key in ("PGPASSWORD", "PGUSER", "PGHOST", "PGPORT", "PGDATABASE",
                       "PG_HOST", "PG_PASSWORD"):
            self.assertNotIn(pg_key, agent_env, f"{pg_key} leaked into kimi agent env")
        # MODEL_API_KEY stays — the CLI needs it to call the model. It is
        # stripped again at the python_execute tool boundary.
        self.assertIn("MODEL_API_KEY", agent_env)

    def test_main_agentfile_overrides_defaults_and_blocks_builtin_shell(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "agent_main.md")

            with mock.patch.dict(os.environ, {"KIMI_SUBAGENTS": ""}):
                kimi_main.write_agentfile(path, "System prompt")

            text = Path(path).read_text(encoding="utf-8")
            self.assertIn("override: true\n", text)
            self.assertIn("disallowedTools:\n", text)
            self.assertIn("  - Bash\n", text)
            self.assertIn("  - Shell\n", text)

    def test_generated_explore_profile_is_valid_and_overrides_builtin_tools(self):
        text = kimi_main.render_subagent_profile(
            "explore",
            tools_list=["mcp__filesystem__read_*"],
            disallowed=[kimi_main.CLAIM_TOOL, *kimi_main.BUILTIN_DISALLOWED_TOOLS],
        )

        self.assertIn("name: explore\n", text)
        self.assertIn("description: ", text)
        self.assertIn("whenToUse: ", text)
        self.assertIn("override: true\n", text)
        self.assertIn("  - mcp__filesystem__read_*\n", text)
        self.assertIn("disallowedTools:\n", text)
        self.assertIn("  - Bash\n", text)
        self.assertIn("  - WebSearch\n", text)
        self.assertIn("  - FetchURL\n", text)
        self.assertIn("  - Read\n", text)
        self.assertIn("  - Fetch\n", text)

    def test_compute_blocked_read_prefixes_blocks_sibling_kimi_home(self):
        # P0-1 boundary hardening: .kimi_home is a sibling of the workspace
        # under task_root and holds session internals (wire.jsonl, MCP config,
        # per-agent tool-result offloads). It must be added to the read
        # block-list at server startup. See dev_docs/2026-08-13-c2-tz-fix-design.md §1.
        with tempfile.TemporaryDirectory() as task_root:
            workspace = os.path.join(task_root, "workspace")
            kimi_home = os.path.join(task_root, ".kimi_home")
            os.makedirs(workspace)
            os.makedirs(kimi_home)

            original = list(local_tools_server._BLOCKED_READ_PREFIXES)
            try:
                local_tools_server._BLOCKED_READ_PREFIXES = list(original)
                local_tools_server._compute_blocked_read_prefixes(workspace)
                blocked = local_tools_server._BLOCKED_READ_PREFIXES
                self.assertIn(os.path.realpath(kimi_home), blocked)
                # The workspace itself is never blocked.
                self.assertNotIn(os.path.realpath(workspace), blocked)
            finally:
                local_tools_server._BLOCKED_READ_PREFIXES = original

    def test_compute_blocked_read_prefixes_picks_up_kimi_code_home_env(self):
        with tempfile.TemporaryDirectory() as workspace:
            custom_home = tempfile.mkdtemp(prefix="custom_kimi_home_")
            original = list(local_tools_server._BLOCKED_READ_PREFIXES)
            try:
                local_tools_server._BLOCKED_READ_PREFIXES = list(original)
                with mock.patch.dict(os.environ, {"KIMI_CODE_HOME": custom_home}):
                    local_tools_server._compute_blocked_read_prefixes(workspace)
                self.assertIn(os.path.realpath(custom_home),
                              local_tools_server._BLOCKED_READ_PREFIXES)
            finally:
                local_tools_server._BLOCKED_READ_PREFIXES = original
                import shutil
                shutil.rmtree(custom_home, ignore_errors=True)

    def test_read_guard_blocks_relative_path_into_kimi_home(self):        # The injected read-guard must resolve symlinks/`..` so an agent cannot
        # escape the block with `../.kimi_home/...`. This runs the actual guard
        # preamble in a real subprocess.
        with tempfile.TemporaryDirectory() as task_root:
            workspace = os.path.join(task_root, "workspace")
            kimi_home = os.path.join(task_root, ".kimi_home")
            os.makedirs(workspace)
            os.makedirs(kimi_home, exist_ok=True)
            secret = os.path.join(kimi_home, "wire.jsonl")
            with open(secret, "w", encoding="utf-8") as f:
                f.write("SECRET_TRANSCRIPT")

            original = list(local_tools_server._BLOCKED_READ_PREFIXES)
            local_tools_server._BLOCKED_READ_PREFIXES = list(original)
            local_tools_server._compute_blocked_read_prefixes(workspace)
            try:
                guard = local_tools_server._READ_GUARD.format(
                    blocked=list(local_tools_server._BLOCKED_READ_PREFIXES),
                    msg=local_tools_server._BLOCKED_READ_MSG)
                # Read via a relative `..` escape from the workspace cwd.
                code = guard + (
                    "import os\n"
                    "try:\n"
                    "    open('../.kimi_home/wire.jsonl').read()\n"
                    "    print('LEAK')\n"
                    "except PermissionError as e:\n"
                    "    print('BLOCKED')\n"
                )
                proc = subprocess.run(
                    [local_tools_server.sys.executable, "-c", code],
                    cwd=workspace, capture_output=True, text=True, timeout=20,
                )
                self.assertIn("BLOCKED", proc.stdout)
                self.assertNotIn("LEAK", proc.stdout)
            finally:
                local_tools_server._BLOCKED_READ_PREFIXES = original

    def test_mcp_health_check_detects_missing_uv_venv(self):
        # P0-2: the google_sheet MCP was silently dropped because its uv .venv
        # was missing at runtime. The pre-flight health check must flag a
        # uv-launched server whose project dir has no .venv/bin/python, so the
        # run is classified as infra_failed (auto-rerun) instead of a model fail.
        with tempfile.TemporaryDirectory() as tmp:
            # A uv project dir WITHOUT a .venv.
            proj = os.path.join(tmp, "mcp-fake")
            os.makedirs(proj)
            # mcp.json declaring a uv server rooted at `proj`.
            mcp_json = os.path.join(tmp, "mcp.json")
            with open(mcp_json, "w", encoding="utf-8") as f:
                json = __import__("json")
                json.dump({"mcpServers": {
                    "fake_sheet": {
                        "command": "uv", "args": ["run", "server.py"],
                        "env": {}, "cwd": proj,
                    }}}, f)
            failures = kimi_main._check_mcp_servers_health(mcp_json)
            self.assertTrue(any("no .venv/bin/python" in m for m in failures),
                            f"expected missing-venv failure, got: {failures}")

    def test_mcp_health_check_detects_missing_command_binary(self):
        with tempfile.TemporaryDirectory() as tmp:
            mcp_json = os.path.join(tmp, "mcp.json")
            with open(mcp_json, "w", encoding="utf-8") as f:
                json = __import__("json")
                json.dump({"mcpServers": {
                    "ghost": {
                        "command": "/opt/does/not/exist/bin/ghost",
                        "args": [], "env": {}, "cwd": tmp,
                    }}}, f)
            failures = kimi_main._check_mcp_servers_health(mcp_json)
            self.assertTrue(any("not found" in m or "not executable" in m
                                for m in failures),
                            f"expected missing-binary failure, got: {failures}")

    def test_mcp_health_check_passes_for_valid_uv_server(self):
        with tempfile.TemporaryDirectory() as tmp:
            proj = os.path.join(tmp, "mcp-ok")
            venv_bin = os.path.join(proj, ".venv", "bin")
            os.makedirs(venv_bin)
            # Touch a fake python binary so the check sees it.
            ok_python = os.path.join(venv_bin, "python")
            with open(ok_python, "w", encoding="utf-8") as f:
                f.write("#!/bin/sh\n")
            os.chmod(ok_python, 0o755)
            mcp_json = os.path.join(tmp, "mcp.json")
            with open(mcp_json, "w", encoding="utf-8") as f:
                json = __import__("json")
                json.dump({"mcpServers": {
                    "ok_sheet": {
                        "command": "uv", "args": ["run", "server.py"],
                        "env": {}, "cwd": proj,
                    }}}, f)
            failures = kimi_main._check_mcp_servers_health(mcp_json)
            # Only the uv command itself might be unresolved on the host; the
            # .venv check must NOT fire for a valid venv.
            self.assertFalse(any("no .venv/bin/python" in m for m in failures),
                             f"unexpected missing-venv failure: {failures}")

    def test_mcp_health_check_uses_directory_arg_not_cwd_for_uv_venv(self):
        # Regression: excel/google_sheet pass the uv project via `--directory`
        # while `cwd` is the agent workspace. The health check MUST resolve the
        # .venv from `--directory`, NOT from `cwd`, or every excel/google_sheet
        # task is misclassified infra_failed. See C.3 rerun 2026-08-13.
        with tempfile.TemporaryDirectory() as tmp:
            proj = os.path.join(tmp, "excel-mcp-server")
            workspace = os.path.join(tmp, "workspace")
            os.makedirs(proj)
            os.makedirs(workspace)
            # .venv lives in the PROJECT dir (--directory), not in workspace.
            venv_bin = os.path.join(proj, ".venv", "bin")
            os.makedirs(venv_bin)
            ok_python = os.path.join(venv_bin, "python")
            with open(ok_python, "w", encoding="utf-8") as f:
                f.write("#!/bin/sh\n")
            os.chmod(ok_python, 0o755)
            mcp_json = os.path.join(tmp, "mcp.json")
            with open(mcp_json, "w", encoding="utf-8") as f:
                json = __import__("json")
                json.dump({"mcpServers": {
                    "excel": {
                        "command": "uv",
                        "args": ["--directory", proj, "run", "excel-mcp-server", "stdio"],
                        "env": {}, "cwd": workspace,
                    }}}, f)
            failures = kimi_main._check_mcp_servers_health(mcp_json)
            self.assertFalse(any("no .venv/bin/python" in m for m in failures),
                             f"expected --directory venv to pass, got: {failures}")

    def test_mcp_health_check_fails_when_directory_venv_missing(self):
        # Counterpart: if the --directory project has no .venv, the check must
        # still fire (cwd having or not having a .venv is irrelevant).
        with tempfile.TemporaryDirectory() as tmp:
            proj = os.path.join(tmp, "mcp-google-sheets")  # no .venv
            workspace = os.path.join(tmp, "workspace")
            os.makedirs(proj)
            os.makedirs(workspace)
            mcp_json = os.path.join(tmp, "mcp.json")
            with open(mcp_json, "w", encoding="utf-8") as f:
                json = __import__("json")
                json.dump({"mcpServers": {
                    "google_sheet": {
                        "command": "uv",
                        "args": ["--directory", proj, "run", "mcp-google-sheets"],
                        "env": {}, "cwd": workspace,
                    }}}, f)
            failures = kimi_main._check_mcp_servers_health(mcp_json)
            self.assertTrue(any("no .venv/bin/python" in m for m in failures),
                            f"expected missing-venv failure from --directory, got: {failures}")


if __name__ == "__main__":
    unittest.main()
