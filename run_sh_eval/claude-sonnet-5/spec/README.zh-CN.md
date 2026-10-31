# claude-sonnet-5 · spec

[English](README.md) | 简体中文

MAD-SAE 模式，使用专家子 Agent：`plan` 加 6 个领域专家。对应的单 Agent 基线为 [`../sa`](../sa/README.zh-CN.md)。

## 运行

```bash
cp config.env.example config.env && chmod 600 config.env   # 填写 API_BASE_URL 和 API_KEY
bash run.sh --backend docker --dry-run                     # 检查配置，不调用 API
bash run.sh --backend docker                               # 运行全部 166 个任务
```

结果写入 MS-Eval 根目录下的 `results/claude-sonnet-5/spec/`。Enroot、参数及输出文件见[运行指南](../../README.zh-CN.md)。

## 默认配置

以下为论文使用的配置，复现结果时请保持不变。API 模型 ID：`claude-sonnet-5`（[官方文档](https://platform.claude.com/docs/en/about-claude/model-deprecations)）；`API_FORMAT=anthropic`。`anthropic` 表示原生 Messages 接口（`x-api-key` 鉴权）；若接口只支持 OpenAI 兼容协议，改为 `openai`。

| `config.env` | 默认值 | 含义 |
| --- | --- | --- |
| `TEMPERATURE`、`TOP_P` | 1、模型默认 | 采样参数 |
| `REASONING_EFFORT` | 模型默认 | 推理强度 |
| `ENABLE_THINKING` | 0 | 为 `1` 时在请求中加入 `enable_thinking: true` |
| `BREAK_SCHEMA_CYCLES` | 0 | 为 `1` 时去除工具 schema 中的循环引用 |
| `MAX_OUTPUT_TOKENS` | 128000 | 单次请求的输出 token 上限 |
| `MAX_CONTEXT`、`MAX_STEPS`、`TASK_TIMEOUT_S` | 262144、100、7200 | 单任务上限：token、步数、秒 |
| `MAX_CONCURRENT` | 1 | 并行任务数 |
