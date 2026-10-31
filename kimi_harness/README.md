# Kimi evaluation harness

This directory connects the experiment launchers in [run_sh_eval](../run_sh_eval) to Kimi Code and the Toolathlon task environment. `kimi_main.py` prepares the task workspace, renders prompts and agent profiles, configures MCP servers, runs the agent, and records trajectories and evaluation results.

## Runtime files

| File | Purpose |
| --- | --- |
| `kimi_main.py` | Task lifecycle, prompt assembly, tool permissions, Kimi configuration, and evaluation |
| `mcp_json_gen.py` | Resolve task MCP configurations into Kimi's `mcp.json` |
| `local_tools_server.py` | Serve task-local tools, including Python execution and completion signaling |
| `traj_convert.py` | Convert the raw Kimi stream into the trajectory structure consumed by the harness |
| `assets/subagents_gen/` | Gen profiles: `coder`, `explore`, and `plan` |
| `assets/subagents_spec/` | Spec profiles: `plan` and six domain specialists |
| `assets/sections/` | Agent catalogs, delegation examples, and plan-first instructions for the two rosters |
| `tests/test_infra_hardening.py` | Offline checks for tool boundaries, configuration, infrastructure handling, and prompt stability |

All four Python runtime modules are used by the evaluation path. The two `plan.md` files belong to different rosters and have distinct content; both are required.

## Experiment variants

| Variant | Roster selection | Prompt assets |
| --- | --- | --- |
| `spec` | `KIMI_CODE_SUBAGENT=spec` | `subagents_spec/` and `*_spec.md` sections |
| `gen` | `KIMI_CODE_SUBAGENT=gen` | `subagents_gen/` and `*_gen.md` sections |
| `oracle` | `KIMI_CODE_SUBAGENT=gen` | Same three-agent assets, with the task brief selected by the launcher |
| `plan1st` | `KIMI_CODE_SUBAGENT=gen` and `KIMI_PLAN_FIRST=1` | Three-agent assets plus `plan_first_gen.md` |
| `sa` | `KIMI_CODE_SUBAGENT=none` | No sub-agent profiles or delegation sections |

`gen` names the general-purpose roster (`coder`, `explore`, `plan`); `spec` names the specialized roster (`plan` plus six domain specialists). `oracle` and `plan1st` use gen with different task guidance or planning instructions.

The standard launchers in [run_sh_eval](../run_sh_eval/README.md) set these variables for each variant. Custom experiments outside the launchers can set `KIMI_CODE_SUBAGENT` and `KIMI_PLAN_FIRST` directly.

## Custom prompt files

Custom integrations can use `KIMI_EXAMPLES_FILE`, `KIMI_COORDINATION_FILE`, and `KIMI_PLAN_FIRST_FILE` to select section files. Paths are relative to `assets/sections/`, or absolute within the runtime environment. Coordination instructions are omitted when no file is selected. Set `KIMI_EXAMPLES_FILE=/dev/null` to omit examples; `KIMI_PLAN_FIRST_FILE` takes effect only when plan-first is enabled and sub-agents are active. Standard launchers enforce the published prompt settings.
