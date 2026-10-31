# minimax-m3 · sa

[English](README.md) | 简体中文

单 Agent（SA）基线：模型与配置同 [`../spec`](../spec/README.zh-CN.md)，但不使用子 Agent。该模型的各个 MAD-SAE 模式都与它对比。

## 运行

```bash
cp -p ../spec/config.env config.env      # 与 spec 配置相同；也可由 config.env.example 填写
bash run.sh --backend docker --dry-run   # 检查配置，不调用 API
bash run.sh --backend docker             # 运行全部 166 个任务
```

结果写入 MS-Eval 根目录下的 `results/minimax-m3/sa/`。Enroot、参数及输出文件见[运行指南](../../README.zh-CN.md)。

## 默认配置

以下为论文使用的配置，复现结果时请保持不变。API 模型 ID：`MiniMax-M3`（[官方文档](https://platform.minimax.io/docs/guides/text-generation)）；`API_FORMAT=openai`。

| `config.env` | 默认值 | 含义 |
| --- | --- | --- |
| `TEMPERATURE`、`TOP_P` | 1、0.95 | 采样参数 |
| `REASONING_EFFORT` | adaptive | 推理强度 |
| `ENABLE_THINKING` | 0 | 为 `1` 时在请求中加入 `enable_thinking: true` |
| `BREAK_SCHEMA_CYCLES` | 0 | 为 `1` 时去除工具 schema 中的循环引用 |
| `MAX_OUTPUT_TOKENS` | 0（不限制） | 单次请求的输出 token 上限 |
| `MAX_CONTEXT`、`MAX_STEPS`、`TASK_TIMEOUT_S` | 262144、100、7200 | 单任务上限：token、步数、秒 |
| `MAX_CONCURRENT` | 1 | 并行任务数 |
