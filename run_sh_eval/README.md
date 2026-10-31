# Evaluation launchers

[中文说明](README.zh-CN.md)

Choose a model/variant directory to run the fixed 166-task suite. Each launcher explicitly selects its roster and prompts.

- **Docker:** build the MS-Eval images, then follow [Running evaluations](#running-evaluations). Use the model's `run.sh --backend docker` entrypoint.
- **Cluster:** prepare the image and runtime files with the [Enroot build guide](../../enroot/README.md). `run.sh` keeps the existing cluster backend by default; `--backend enroot` selects it explicitly.

Start evaluations with `<model_name>/<execution_mode>/run.sh`. It loads that preset's `config.env` (or `--config FILE`) and selects the requested backend. `_common/` and `docker/` contain internal support code.

## Directory layout

```text
run_sh_eval/
├── README.md
├── README.zh-CN.md
├── _common/
│   └── launch.sh               # Internal configuration loading and backend selection
├── docker/                    # Docker runner, API gateway and container entrypoint
├── claude-sonnet-5/
│   ├── sa/                    # Single-Agent baseline
│   └── spec/                  # Specialist sub-Agents
├── deepseek-v4-flash-0731/
│   ├── sa/                    # Single-Agent baseline
│   ├── spec/                  # Specialist sub-Agents
│   ├── gen/                   # Generalist sub-Agents
│   ├── oracle/                # Generalist sub-Agents + oracle task briefs
│   └── plan1st/               # Generalist sub-Agents + plan-first
├── glm-5.3/
│   ├── sa/                    # Single-Agent baseline
│   └── spec/                  # Specialist sub-Agents
├── glm-5.3-flash/
│   ├── sa/                    # Single-Agent baseline
│   └── spec/                  # Specialist sub-Agents
├── gpt-5.6-sol/
│   ├── sa/                    # Single-Agent baseline
│   └── spec/                  # Specialist sub-Agents
├── kimi-k3/
│   ├── sa/                    # Single-Agent baseline
│   └── spec/                  # Specialist sub-Agents
├── minimax-m3/
│   ├── sa/                    # Single-Agent baseline
│   └── spec/                  # Specialist sub-Agents
└── qwen-3.8-max/
    ├── sa/                    # Single-Agent baseline
    └── spec/                  # Specialist sub-Agents
```

Each `<model_name>/<execution_mode>/` directory contains:

| File | Purpose |
| --- | --- |
| `run.sh` | Evaluation entrypoint for this model and variant |
| `config.env.example` | API and runtime configuration template; copy to `config.env` |
| `model.json`, `eval_config.json` | Model defaults, experiment variant and evaluation settings |
| `tasks.txt` | Fixed 166-task manifest |
| `launch.py`, `runner.sh` | Preset preparation and cluster backend |
| `api_gateway.py`, `schema.py` | Model API adaptation |
| `report.py` | Check completion of a full evaluation run |
| `README.md` | Instructions for this preset |

Keep the full directory structure when running a preset. Local `config.env` files and run results are Git-ignored.

## Evaluation modes

| Variant | Roster | Prompt sections | Plan-first | Environment variables |
| --- | --- | --- | --- | --- |
| `sa` | `none`: single Agent | No delegation sections | Off | `KIMI_CODE_SUBAGENT=none` |
| `spec` | `spec`: plan and six specialists | `subagent_types_spec.md`, `examples_spec.md` | Off | `KIMI_CODE_SUBAGENT=spec` |
| `gen` | `gen`: coder, explore, plan | `subagent_types_gen.md`, `examples_gen.md` | Off | `KIMI_CODE_SUBAGENT=gen` |
| `oracle` | `gen`, with oracle task briefs | Same as gen | Off | `KIMI_CODE_SUBAGENT=gen` |
| `plan1st` | `gen` | Same as gen, plus `plan_first_gen.md` | On | `KIMI_CODE_SUBAGENT=gen`<br>`KIMI_PLAN_FIRST=1` |

The launcher sets these variables automatically and unsets `KIMI_PLAN_FIRST` for every variant except `plan1st`.

All model directories provide `sa` and `spec`, each with its own `config.env.example`, `model.json`, `tasks.txt`, and `run.sh`. SA uses the same model, sampling defaults and task set as that model's spec preset, with delegation and plan-first disabled. The [DeepSeek directory](deepseek-v4-flash-0731/) additionally provides `gen`, `oracle`, and `plan1st`. Profiles and prompt sections live under [kimi_harness/assets](../kimi_harness/assets/).

`run.sh` loads the local `config.env`, then `launch.py` applies the selected experiment's fixed roster and prompt settings. Setting `KIMI_CODE_SUBAGENT` or `KIMI_PLAN_FIRST` in `config.env` does not override the published preset. Use `./run.sh --dry-run` inside an experiment directory to inspect the selected settings without calling an API. To run SA, enter the model's `sa/` directory and use `bash run.sh --backend docker` or `bash run.sh --backend enroot`. Configure SA and spec separately with matching model and sampling settings for comparison.

For runtime module responsibilities and the asset layout, see the [harness README](../kimi_harness/README.md).

## Running evaluations

First build the image using the [Docker](../../docker/README.md) or [Enroot](../../enroot/README.md) guide. Replace `<model_name>` and `<execution_mode>` with the corresponding directory names, then run from the MS-Eval repository root:

```bash
# execution_mode = sa, spec, gen, oracle, plan1st
cd toolathlon_gym/run_sh_eval/<model_name>/<execution_mode>
cp config.env.example config.env
# Edit config.env: set API_BASE_URL and API_KEY; confirm MODEL_NAME and API_FORMAT.
bash run.sh --backend docker --task canvas-course-comparison
bash run.sh --backend docker
```

For Enroot, first run `source .local/enroot/env.sh` from the repository root (adjust for a custom build directory), then use `bash run.sh --backend enroot` in place of the launch commands above. This backend runs all 166 tasks per invocation.

Each preset has its own `config.env`; keep model and sampling settings identical when comparing SA and spec. Start with `MAX_CONCURRENT=1`. Set evaluation limits such as `MAX_STEPS` and `TASK_TIMEOUT_S` in the same file. Optional runtime path overrides may remain commented out for Docker.

Results are grouped by model preset directory name and evaluation variant under `results/<model_name>/<execution_mode>/<run-id>/` at the MS-Eval root, for example `results/deepseek-v4-flash-0731/spec/<run-id>/`. Each run gets a unique `<run-id>` to avoid overwriting earlier results. Per-task logs, trajectories and `eval_res.json` contain the evaluation results; the batch directory also contains a summary CSV. A successful runner exit means the evaluation completed; inspect `eval_res.json` for task scores.

## run.sh Arguments

Run `bash run.sh [arguments]` from the model's execution mode directory. Common arguments:

| Argument | Purpose |
| --- | --- |
| `--backend docker` / `--backend enroot` | Select the backend; defaults to Enroot |
| `--config FILE` | Select a configuration file; defaults to `run_sh_eval/<model_name>/<execution_mode>/config.env` |
| `--task NAME` | Select a task; repeat for a subset, or omit to run all 166 tasks (Docker backend only) |
| `--output DIR` | Set the results parent directory (Docker only); precedence: `--output` > `OUTPUT_ROOT` in `config.env` > default directory |
| `--dry-run` | Show preset settings without starting evaluation or calling the model API |
| `--check` | Check the runtime environment and configuration without calling the model API |
| `--help` | Show all arguments supported by the selected backend |
