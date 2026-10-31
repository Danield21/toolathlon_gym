# claude-sonnet-5 · sa

English | [简体中文](README.zh-CN.md)

Single-agent (SA) baseline for `claude-sonnet-5`: the same model and settings as [`spec`](../spec/README.md), without sub-agents. Keep this `config.env` identical to the one in `../spec`.

```bash
cp config.env.example config.env && chmod 600 config.env   # set API_BASE_URL and API_KEY
bash run.sh --backend docker --dry-run                       # print settings, no API calls
bash run.sh --backend docker                                 # all 166 tasks
```

Results go to `results/claude-sonnet-5/sa/` in the MS-Eval root. Enroot, options and output files: [run guide](../../README.md).

## Defaults

API model ID `claude-sonnet-5` ([docs](https://platform.claude.com/docs/en/about-claude/model-deprecations)), `API_FORMAT=anthropic`. `anthropic` selects the native Messages API (`x-api-key` authentication); use `openai` if your endpoint is OpenAI-compatible only.

| `config.env` | Default | Meaning |
| --- | --- | --- |
| `TEMPERATURE`, `TOP_P` | 1, model default | Sampling |
| `REASONING_EFFORT` | model default | Reasoning effort |
| `ENABLE_THINKING` | 0 | Add `enable_thinking: true` to requests |
| `BREAK_SCHEMA_CYCLES` | 0 | Remove recursive references from tool schemas |
| `MAX_OUTPUT_TOKENS` | 128000 | Output token cap per request |
| `MAX_CONTEXT`, `MAX_STEPS`, `TASK_TIMEOUT_S` | 262144, 100, 7200 | Limits per task: tokens, steps, seconds |
| `MAX_CONCURRENT` | 1 | Tasks run in parallel |
