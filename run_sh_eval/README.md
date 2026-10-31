# Evaluation Launchers

English | [简体中文](README.zh-CN.md)

Each `<model>/<variant>/` directory launches the fixed 166-task MS-Eval suite for one configuration. Results go to `results/<model>/<variant>/` in the MS-Eval root.

## Quick Start

Build the images first ([Docker](../../docker/README.md) or [Enroot](../../enroot/README.md)). Then, from the MS-Eval root:

```bash
cd toolathlon_gym/run_sh_eval/kimi-k3
cp spec/config.env.example spec/config.env   # set API_BASE_URL and API_KEY
chmod 600 spec/config.env
cp -p spec/config.env sa/config.env          # the SA baseline uses the same settings

bash spec/run.sh --backend docker --dry-run                         # check settings, no API calls
bash spec/run.sh --backend docker --task canvas-course-comparison   # try one task
bash spec/run.sh --backend docker                                   # MAD-SAE, all 166 tasks
bash sa/run.sh --backend docker                                     # SA baseline, all 166 tasks
```

Next, [compute the metrics](../../eval/README.md). On a cluster, run `source .local/enroot/env.sh` in the MS-Eval root and use `--backend enroot` instead; each Enroot invocation runs all 166 tasks.

## Presets

| Model | Variants |
| --- | --- |
| `claude-sonnet-5`, `glm-5.3`, `glm-5.3-flash`, `gpt-5.6-sol`, `kimi-k3`, `minimax-m3`, `qwen-3.8-max` | `sa`, `spec` |
| `deepseek-v4-flash-0731` | `sa`, `spec`, `gen`, `oracle`, `plan1st` |

| Variant | Paradigm | Sub-agents | Notes |
| --- | --- | --- | --- |
| `sa` | SA | none | Baseline for the other variants |
| `spec` | MAD-SAE | `plan` + 6 domain specialists | |
| `gen` | MAD-SAE | `coder`, `explore`, `plan` | |
| `oracle` | MAD-SAE | same as `gen` | Task briefs include the Delegation Prior ([`data/oracle166`](../../data/oracle166/)) |
| `plan1st` | MAD-SAE | same as `gen` | The main agent gets a plan from `plan` before acting |

Each launcher fixes the sub-agents and prompts of its variant; `config.env` cannot change them. Agent profiles and prompts are described in [kimi_harness](../kimi_harness/README.md).

## Configuration

Each preset reads its own `config.env` (Git-ignored), copied from `config.env.example`. The defaults are the settings used in the paper; each preset's README lists them.

| Key | Description |
| --- | --- |
| `API_BASE_URL`, `API_KEY` | API base URL (without `/chat/completions` or `/messages`) and key. Required. |
| `API_FORMAT` | `openai` (OpenAI-compatible) or `anthropic` (Messages API) |
| `MODEL_NAME` | API model ID; change it only if your endpoint uses another ID |
| `TEMPERATURE`, `TOP_P`, `REASONING_EFFORT`, ... | Request settings; empty means the model default |
| `MAX_STEPS`, `TASK_TIMEOUT_S`, `MAX_CONTEXT` | Per-task limits: 100 steps, 7200 s, 262,144 tokens |
| `MAX_CONCURRENT` | Tasks run in parallel; start with `1` |
| `OUTPUT_ROOT`, `HTTPS_PROXY` | Optional: results parent directory; proxy for model API requests |

The commented runtime paths at the end of the file are Enroot overrides; Docker does not use them.

## Options

| Option | Description |
| --- | --- |
| `--backend docker` / `enroot` | Runtime; default `enroot` |
| `--config FILE` | Use another config file |
| `--task NAME` | Run one task; repeat for several (Docker only) |
| `--output DIR` | Results parent directory (Docker only; `OUTPUT_ROOT` works for both backends) |
| `--dry-run` | Print the settings and exit |
| `--check` | Check the runtime and configuration without calling the API |
| `--smoke` | Offline check of the database, MCP servers and network; no API calls (Docker only) |

`bash run.sh --backend docker --help` lists all options.

## Output

```text
results/<model>/<variant>/<run-id>/
├── run.json, tasks.txt              # settings and task list
├── summary_parallel_<run-id>.csv    # execution status per task
├── audit_index.html                 # overview of all tasks
└── <task>/<run-id>_slotN/           # logs, agent session, audit.html
```

Each invocation creates a new `<run-id>` and keeps earlier runs. To retry tasks, run them again with `--task`; the metric scripts use the latest run of each task. To check that a run covered all 166 tasks, run `python3 report.py <run-dir>` in the preset directory.
