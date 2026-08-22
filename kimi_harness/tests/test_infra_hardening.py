import os
import subprocess
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import yaml

from kimi_harness import kimi_main, local_tools_server, mcp_json_gen


class KimiInfraHardeningTests(unittest.TestCase):
    def test_provider_wire_and_base_url_normalization(self):
        with mock.patch.dict(
            os.environ,
            {"MODEL_API_URL": "https://gateway.example/v1/"},
            clear=True,
        ):
            self.assertEqual(
                kimi_main.resolve_model_provider_config(),
                ("openai", "https://gateway.example/v1"),
            )

        with mock.patch.dict(
            os.environ,
            {
                "KIMI_PROVIDER_TYPE": "anthropic",
                "MODEL_API_URL": "https://gateway.example/v1/",
            },
            clear=True,
        ):
            self.assertEqual(
                kimi_main.resolve_model_provider_config(),
                ("anthropic", "https://gateway.example"),
            )

        with mock.patch.dict(
            os.environ,
            {
                "KIMI_PROVIDER_TYPE": "unsupported-wire",
                "MODEL_API_URL": "https://gateway.example",
            },
            clear=True,
        ):
            with self.assertRaisesRegex(ValueError, "unsupported KIMI_PROVIDER_TYPE"):
                kimi_main.resolve_model_provider_config()

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

    def test_readonly_allowlist_grants_misnamed_read_tools_and_blocks_writes(self):
        servers = [
            "canvas", "woocommerce", "notion", "youtube", "howtocook",
            "rail_12306", "scholarly", "playwright_with_chunk", "snowflake",
            "filesystem", "emails", "excel",
        ]
        allow = kimi_main.build_readonly_tool_patterns(servers)
        deny = kimi_main.READONLY_TOOL_DENY_PATTERNS
        granted = lambda name: kimi_main.mcp_tool_granted_to_readonly_agent(
            name, allow, deny
        )

        reads = [
            "mcp__canvas__canvas_get_course_grades",
            "mcp__canvas__canvas_list_courses",
            "mcp__canvas__canvas_health_check",
            "mcp__woocommerce__woo_products_list",
            "mcp__woocommerce__woo_products_get",
            "mcp__woocommerce__woo_reports_sales",
            "mcp__woocommerce__woo_system_status",
            "mcp__notion__API-get-user",
            "mcp__notion__API-get-users",
            "mcp__notion__API-retrieve-a-page",
            "mcp__notion__API-post-search",
            "mcp__notion__API-post-database-query",
            "mcp__youtube__videos_getVideo",
            "mcp__youtube__channels_listVideos",
            "mcp__youtube__playlists_searchPlaylists",
            "mcp__howtocook__mcp_howtocook_getAllRecipes",
            "mcp__howtocook__mcp_howtocook_getRecipeById",
            "mcp__howtocook__mcp_howtocook_whatToEat",
            "mcp__howtocook__mcp_howtocook_recommendMeals",
            "mcp__rail_12306__get-tickets",
            "mcp__rail_12306__get-station-code-by-names",
            "mcp__scholarly__search-arxiv",
            "mcp__scholarly__search-google-scholar",
            "mcp__playwright_with_chunk__browser_snapshot",
            "mcp__playwright_with_chunk__browser_snapshot_search",
            "mcp__playwright_with_chunk__browser_navigate",
            "mcp__snowflake__read_query",
            "mcp__snowflake__list_tables",
            "mcp__filesystem__read_text_file",
            "mcp__filesystem__directory_tree",
            "mcp__emails__get_emails",
            "mcp__excel__read_data_from_excel",
        ]
        writes = [
            "mcp__canvas__canvas_create_course",
            "mcp__canvas__canvas_update_assignment",
            "mcp__canvas__canvas_delete_quiz",
            "mcp__canvas__canvas_submit_grade",
            "mcp__woocommerce__woo_products_create",
            "mcp__woocommerce__woo_orders_update",
            "mcp__woocommerce__woo_products_delete",
            "mcp__woocommerce__woo_system_tools_run",
            "mcp__notion__API-post-page",
            "mcp__notion__API-patch-page",
            "mcp__notion__API-create-a-database",
            "mcp__notion__API-delete-a-block",
            "mcp__notion__API-update-a-page",
            "mcp__playwright_with_chunk__browser_click",
            "mcp__playwright_with_chunk__browser_type",
            "mcp__playwright_with_chunk__browser_select_option",
            "mcp__snowflake__write_query",
            "mcp__filesystem__write_file",
            "mcp__excel__write_data_to_excel",
            "mcp__emails__send_email",
        ]
        for name in reads:
            self.assertTrue(granted(name), f"expected read tool granted: {name}")
        for name in writes:
            self.assertFalse(granted(name), f"expected write tool denied: {name}")

    def test_explore_profile_denies_write_query_glob(self):
        patterns = kimi_main.build_readonly_tool_patterns(["canvas", "notion"])
        text = kimi_main.render_subagent_profile(
            "explore",
            tools_list=patterns,
            disallowed=[
                kimi_main.CLAIM_TOOL,
                *kimi_main.BUILTIN_DISALLOWED_TOOLS,
                *kimi_main.READONLY_TOOL_DENY_PATTERNS,
            ],
        )
        self.assertIn("mcp__canvas__*get*\n", text)
        self.assertIn("mcp__notion__*retrieve*\n", text)
        self.assertIn("mcp__*__write_*\n", text)

    def test_orchestration_rules_do_not_claim_shared_full_toolset(self):
        inherit = (
            "By default, sub-agents inherit the same task-scoped tools and "
            "workspace permissions as the parent agent"
        )
        override = (
            "if a sub-agent YAML configuration defines customized tool "
            "permissions or workspace directory access, the customized "
            "configuration overrides the inherited defaults"
        )
        for text in (kimi_main.ORCHESTRATION_RULES, kimi_main.ORCHESTRATION_RULES_LEGACY):
            self.assertNotIn("Sub-agents share the same task-approved tools", text)
            self.assertIn(inherit, text)
            self.assertIn(override, text)
        self.assertIn("read-only", kimi_main.ORCHESTRATION_RULES)

    def test_crosscut_profile_sources_are_complete_leaf_contracts(self):
        required = {
            "name", "description", "whenToUse", "override", "tools",
            "disallowedTools", "subagents",
        }
        for name in ("evidence-integrator", "deliverable-auditor"):
            path = Path(kimi_main.HARNESS_DIR) / "assets" / "subagents" / f"{name}.md"
            raw = path.read_text(encoding="utf-8")
            metadata = yaml.safe_load(raw.split("---", 2)[1])
            self.assertEqual(required, set(metadata), name)
            self.assertEqual([], metadata["subagents"], name)
            self.assertEqual(
                {"Agent", "AgentSwarm", kimi_main.CLAIM_TOOL},
                set(metadata["disallowedTools"]),
                name,
            )

    def test_fine_grained_tool_ceilings_follow_role_boundaries(self):
        all_servers = [
            "arxiv-latex", "arxiv_local", "canvas", "emails", "excel",
            "fetch", "filesystem", "google_calendar", "google_forms",
            "google_sheet", "howtocook", "local", "notion", "pdf-tools",
            "playwright_with_chunk", "pptx", "rail_12306", "scholarly",
            "snowflake", "terminal", "woocommerce", "word",
            "yahoo-finance", "youtube", "youtube-transcript",
        ]

        def granted(name):
            return set(kimi_main.profile_tools_for_task(name, all_servers))

        enterprise = granted("enterprise-data-analyst")
        self.assertIn("mcp__canvas__canvas_list_courses", enterprise)
        self.assertIn("mcp__canvas__canvas_get_syllabus", enterprise)
        self.assertIn("mcp__woocommerce__woo_products_reviews_list", enterprise)
        self.assertIn("mcp__woocommerce__woo_reports_sales", enterprise)

        academic = granted("academic-literature-researcher")
        self.assertFalse(any("__youtube__" in tool for tool in academic))

        web = granted("web-domain-researcher")
        self.assertIn("mcp__youtube__videos_getVideo", web)

        workspace = granted("workspace-data-engineer")
        workspace_servers = {tool.split("__", 2)[1] for tool in workspace}
        self.assertLessEqual(
            workspace_servers,
            {"excel", "filesystem", "local", "pdf-tools", "terminal", "word"},
        )

        office = granted("office-report-builder")
        self.assertIn("mcp__pptx__create_presentation", office)
        self.assertIn("mcp__pptx__extract_presentation_text", office)
        self.assertFalse(
            any("__snowflake__" in tool or "__woocommerce__" in tool for tool in office)
        )

        external = granted("external-workflow-operator")
        self.assertIn("mcp__emails__search_emails", external)
        self.assertIn("mcp__notion__API-post-search", external)
        self.assertIn("mcp__notion__API-post-page", external)
        self.assertFalse(
            any("__excel__" in tool or "__snowflake__" in tool for tool in external)
        )

        auditor = granted("deliverable-auditor")
        self.assertIn("mcp__emails__read_email", auditor)
        self.assertIn("mcp__pptx__extract_presentation_text", auditor)
        self.assertNotIn("mcp__emails__send_email", auditor)
        self.assertNotIn("mcp__pptx__save_presentation", auditor)

    def test_fine_grained_profiles_share_versioned_handoff_contracts(self):
        profile_dir = Path(kimi_main.HARNESS_DIR) / "assets" / "subagents"

        for name in (
            "academic-literature-researcher", "web-domain-researcher",
            "enterprise-data-analyst", "financial-market-analyst",
        ):
            text = (profile_dir / f"{name}.md").read_text(encoding="utf-8")
            self.assertIn("EvidencePacket v1", text, name)
            for field in ("scope", "records", "coverage", "missing", "conflicts", "verification"):
                self.assertIn(f"`{field}`", text, f"{name}: {field}")

        contracts = {
            "workspace-data-engineer": "WorkspaceArtifactPacket v1",
            "office-report-builder": "DeliverableReceipt v1",
            "external-workflow-operator": "DeliverableReceipt v1",
            "evidence-integrator": "CanonicalEvidence v1",
            "deliverable-auditor": "AuditReport v1",
        }
        for name, contract in contracts.items():
            text = (profile_dir / f"{name}.md").read_text(encoding="utf-8")
            self.assertIn(contract, text, name)

        self.assertIn("EvidencePacket v1", kimi_main.ORCHESTRATION_RULES)
        self.assertIn("named workspace files", kimi_main.ORCHESTRATION_RULES)
        self.assertIn("path, size, and digest", kimi_main.ORCHESTRATION_RULES)
        self.assertNotIn("digest when available", kimi_main.ORCHESTRATION_RULES)
        self.assertIn("must stop before integration", kimi_main.ORCHESTRATION_RULES)

    def test_ten_agent_prompt_has_one_dedicated_example_per_role(self):
        task_config = SimpleNamespace(
            agent_workspace="/tmp/toolathlon-example-contract",
            system_prompts=SimpleNamespace(agent="Complete the task."),
        )
        expected_roles = (
            "plan",
            "academic-literature-researcher",
            "web-domain-researcher",
            "enterprise-data-analyst",
            "financial-market-analyst",
            "workspace-data-engineer",
            "office-report-builder",
            "external-workflow-operator",
            "evidence-integrator",
            "deliverable-auditor",
        )

        with mock.patch.dict(os.environ, {"KIMI_SUBAGENTS": "ten"}, clear=False):
            prompt = kimi_main.render_system_prompt(task_config)

        examples = prompt.split("<example>")[1:]
        self.assertEqual(len(expected_roles), len(examples))
        for role in expected_roles:
            marker = f"**Primary sub-agent:** `{role}`"
            self.assertEqual(
                1,
                sum(marker in example for example in examples),
                f"expected one dedicated example for {role}",
            )

    def test_three_agent_prompt_includes_plan_example_without_plan_first(self):
        task_config = SimpleNamespace(
            agent_workspace="/tmp/toolathlon-three-example",
            system_prompts=SimpleNamespace(agent="Complete the task."),
        )
        env = {"KIMI_SUBAGENTS": "three"}
        env.pop("KIMI_PLAN_FIRST", None)
        with mock.patch.dict(os.environ, env, clear=False):
            os.environ.pop("KIMI_PLAN_FIRST", None)
            prompt = kimi_main.render_system_prompt(task_config)

        examples = prompt.split("<example>")[1:]
        self.assertEqual(4, len(examples))
        self.assertIn("`plan`", examples[0])
        self.assertIn("Lumenport", examples[0])
        self.assertIn("`explore`", "".join(examples))
        self.assertIn("`coder`", "".join(examples))
        self.assertNotIn("Plan-First Protocol", prompt)
        self.assertNotIn("academic-literature-researcher", "".join(examples))
        self.assertNotIn("evidence-integrator", "".join(examples))
        self.assertNotIn("deliverable-auditor", "".join(examples))

    def test_pre_completion_audit_is_a_fail_closed_gate(self):
        rules = kimi_main.ORCHESTRATION_RULES
        self.assertIn("AuditReport v1", rules)
        self.assertIn("every criterion is PASS", rules)
        self.assertIn("FAIL or UNKNOWN", rules)
        self.assertIn("must not call `mcp__local__claim_done`", rules)

        auditor = (
            Path(kimi_main.HARNESS_DIR)
            / "assets" / "subagents" / "deliverable-auditor.md"
        ).read_text(encoding="utf-8")
        self.assertIn("overall_verdict", auditor)
        self.assertIn("only when every criterion is PASS", auditor)

    def test_orchestration_splits_cross_profile_work_instead_of_broadening_roles(self):
        rules = kimi_main.ORCHESTRATION_RULES
        self.assertIn("producer and consumer phases", rules)
        self.assertIn("Never broaden a profile's tool ceiling", rules)
        self.assertIn("EvidencePacket v1", rules)
        self.assertIn("DeliverableReceipt v1", rules)


if __name__ == "__main__":
    unittest.main()
