# deepseek-v4-flash-0731 · gen

[English](README.md) | 简体中文

`deepseek-v4-flash-0731` 的 MAD-SAE 评测，使用通用子 Agent：`coder`、`explore`、`plan`。SA 基线见 [`../sa`](../sa/README.zh-CN.md)。

```bash
cp config.env.example config.env && chmod 600 config.env   # 填写 API_BASE_URL 和 API_KEY
bash run.sh --backend docker --dry-run                       # 打印配置，不调用 API
bash run.sh --backend docker                                 # 运行全部 166 个任务
```

结果写入 MS-Eval 根目录下的 `results/deepseek-v4-flash-0731/gen/`。Enroot、参数及输出文件见[运行指南](../../README.zh-CN.md)。

## 默认配置

API 模型 ID 为 `deepseek-v4-flash`（[官方文档](https://api-docs.deepseek.com/quick_start/pricing)），`API_FORMAT=openai`。

| `config.env` | 默认值 | 含义 |
| --- | --- | --- |
| `TEMPERATURE`、`TOP_P` | 模型默认、模型默认 | 采样参数 |
| `REASONING_EFFORT` | high | 推理强度 |
| `ENABLE_THINKING` | 0 | 在请求中加入 `enable_thinking: true` |
| `BREAK_SCHEMA_CYCLES` | 0 | 去除工具 schema 中的循环引用 |
| `MAX_OUTPUT_TOKENS` | 0（不限制） | 单次请求的输出 token 上限 |
| `MAX_CONTEXT`、`MAX_STEPS`、`TASK_TIMEOUT_S` | 262144、100、7200 | 单任务上限：token、步数、秒 |
| `MAX_CONCURRENT` | 1 | 并行任务数 |

**版本说明：** 据 DeepSeek 官方文档，`deepseek-v4-flash` 这一 API ID 现已指向 DeepSeek-V4.1-Flash。目录名中的 `0731` 指论文使用的快照；如需复现，请使用仍提供该快照的接口，并将其模型 ID 填入 `MODEL_NAME`。
