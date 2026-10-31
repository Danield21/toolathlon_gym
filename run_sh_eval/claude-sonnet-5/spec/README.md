# claude-sonnet-5 · spec

English | [简体中文](README.zh-CN.md)

MAD-SAE with specialist sub-agents: `plan` plus six domain specialists. Its single-agent baseline is [`../sa`](../sa/README.md).

## Run

```bash
cp config.env.example config.env && chmod 600 config.env   # set API_BASE_URL and API_KEY
bash run.sh --backend docker --dry-run                     # check settings, no API calls
bash run.sh --backend docker                               # all 166 tasks
```

Results go to `results/claude-sonnet-5/spec/` in the MS-Eval root. For Enroot, options and output files, see the [run guide](../../README.md).

## Defaults

These are the settings used in the paper; keep them to reproduce the results. API model ID: `claude-sonnet-5` ([docs](https://platform.claude.com/docs/en/about-claude/model-deprecations)); `API_FORMAT=anthropic`. `anthropic` selects the native Messages API (`x-api-key` authentication); use `openai` for an OpenAI-compatible endpoint.

| `config.env` | Default | Meaning |
| --- | --- | --- |
| `TEMPERATURE`, `TOP_P` | 1, model default | Sampling |
| `REASONING_EFFORT` | model default | Reasoning effort |
| `ENABLE_THINKING` | 0 | `1` adds `enable_thinking: true` to requests |
| `BREAK_SCHEMA_CYCLES` | 0 | `1` removes recursive references from tool schemas |
| `MAX_OUTPUT_TOKENS` | 128000 | Output token cap per request |
| `MAX_CONTEXT`, `MAX_STEPS`, `TASK_TIMEOUT_S` | 262144, 100, 7200 | Per-task limits: tokens, steps, seconds |
| `MAX_CONCURRENT` | 1 | Tasks run in parallel |
