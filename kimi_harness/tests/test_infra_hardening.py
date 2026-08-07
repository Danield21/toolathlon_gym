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


if __name__ == "__main__":
    unittest.main()
