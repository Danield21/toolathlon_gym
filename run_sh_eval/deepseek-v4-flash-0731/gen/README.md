# deepseek-v4-flash-0731 · gen

English | [简体中文](README.zh-CN.md)

MAD-SAE with generalist sub-agents: `coder`, `explore` and `plan`. Its single-agent baseline is [`../sa`](../sa/README.md).

## Run

```bash
cp config.env.example config.env && chmod 600 config.env   # set API_BASE_URL and API_KEY
bash run.sh --backend docker --dry-run                     # check settings, no API calls
bash run.sh --backend docker                               # all 166 tasks
```

Results go to `results/deepseek-v4-flash-0731/gen/` in the MS-Eval root. For Enroot, options and output files, see the [run guide](../../README.md).

## Defaults

These are the settings used in the paper; keep them to reproduce the results. API model ID: `deepseek-v4-flash` ([docs](https://api-docs.deepseek.com/quick_start/pricing)); `API_FORMAT=openai`.

| `config.env` | Default | Meaning |
| --- | --- | --- |
| `TEMPERATURE`, `TOP_P` | model default, model default | Sampling |
| `REASONING_EFFORT` | high | Reasoning effort |
| `ENABLE_THINKING` | 0 | `1` adds `enable_thinking: true` to requests |
| `BREAK_SCHEMA_CYCLES` | 0 | `1` removes recursive references from tool schemas |
| `MAX_OUTPUT_TOKENS` | 0 (no cap) | Output token cap per request |
| `MAX_CONTEXT`, `MAX_STEPS`, `TASK_TIMEOUT_S` | 262144, 100, 7200 | Per-task limits: tokens, steps, seconds |
| `MAX_CONCURRENT` | 1 | Tasks run in parallel |

**Version note:** according to DeepSeek's documentation, the `deepseek-v4-flash` API ID now serves DeepSeek-V4.1-Flash. `0731` names the snapshot used in the paper; to reproduce it, use an endpoint that still serves that snapshot and set its model ID as `MODEL_NAME`.
