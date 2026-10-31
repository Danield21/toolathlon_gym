# gpt-5.6-sol · spec

English | [简体中文](README.zh-CN.md)

MAD-SAE run of `gpt-5.6-sol` with specialist sub-agents: `plan` plus six domain specialists. The SA baseline is in [`../sa`](../sa/README.md).

```bash
cp config.env.example config.env && chmod 600 config.env   # set API_BASE_URL and API_KEY
bash run.sh --backend docker --dry-run                       # print settings, no API calls
bash run.sh --backend docker                                 # all 166 tasks
```

Results go to `results/gpt-5.6-sol/spec/` in the MS-Eval root. Enroot, options and output files: [run guide](../../README.md).

## Defaults

API model ID `gpt-5.6-sol` ([docs](https://developers.openai.com/api/docs/models/gpt-5.6-sol)), `API_FORMAT=openai`.

| `config.env` | Default | Meaning |
| --- | --- | --- |
| `TEMPERATURE`, `TOP_P` | model default, model default | Sampling |
| `REASONING_EFFORT` | medium | Reasoning effort |
| `ENABLE_THINKING` | 0 | Add `enable_thinking: true` to requests |
| `BREAK_SCHEMA_CYCLES` | 1 | Remove recursive references from tool schemas |
| `MAX_OUTPUT_TOKENS` | 0 (none) | Output token cap per request |
| `MAX_CONTEXT`, `MAX_STEPS`, `TASK_TIMEOUT_S` | 262144, 100, 7200 | Limits per task: tokens, steps, seconds |
| `MAX_CONCURRENT` | 1 | Tasks run in parallel |
