# minimax-m3 · sa

English | [简体中文](README.zh-CN.md)

Single-agent (SA) baseline: the same model and settings as [`../spec`](../spec/README.md), without sub-agents. Every MAD-SAE variant of this model is compared with it.

## Run

```bash
cp -p ../spec/config.env config.env      # same settings as spec, or fill in config.env.example
bash run.sh --backend docker --dry-run   # check settings, no API calls
bash run.sh --backend docker             # all 166 tasks
```

Results go to `results/minimax-m3/sa/` in the MS-Eval root. For Enroot, options and output files, see the [run guide](../../README.md).

## Defaults

These are the settings used in the paper; keep them to reproduce the results. API model ID: `MiniMax-M3` ([docs](https://platform.minimax.io/docs/guides/text-generation)); `API_FORMAT=openai`.

| `config.env` | Default | Meaning |
| --- | --- | --- |
| `TEMPERATURE`, `TOP_P` | 1, 0.95 | Sampling |
| `REASONING_EFFORT` | adaptive | Reasoning effort |
| `ENABLE_THINKING` | 0 | `1` adds `enable_thinking: true` to requests |
| `BREAK_SCHEMA_CYCLES` | 0 | `1` removes recursive references from tool schemas |
| `MAX_OUTPUT_TOKENS` | 0 (no cap) | Output token cap per request |
| `MAX_CONTEXT`, `MAX_STEPS`, `TASK_TIMEOUT_S` | 262144, 100, 7200 | Per-task limits: tokens, steps, seconds |
| `MAX_CONCURRENT` | 1 | Tasks run in parallel |
