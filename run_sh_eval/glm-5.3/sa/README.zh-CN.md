# glm-5.3 · sa

[English](README.md) | 简体中文

`glm-5.3` 的单 Agent（SA）基线：模型与配置同 [`spec`](../spec/README.zh-CN.md)，但不使用子 Agent。请保持本目录的 `config.env` 与 `../spec` 中的一致。

```bash
cp config.env.example config.env && chmod 600 config.env   # 填写 API_BASE_URL 和 API_KEY
bash run.sh --backend docker --dry-run                       # 打印配置，不调用 API
bash run.sh --backend docker                                 # 运行全部 166 个任务
```

结果写入 MS-Eval 根目录下的 `results/glm-5.3/sa/`。Enroot、参数及输出文件见[运行指南](../../README.zh-CN.md)。

## 默认配置

API 模型 ID 为 `glm-5.3`（[官方文档](https://docs.bigmodel.cn/cn/guide/models/text/glm-5.3)），`API_FORMAT=openai`。

| `config.env` | 默认值 | 含义 |
| --- | --- | --- |
| `TEMPERATURE`、`TOP_P` | 1、0.95 | 采样参数 |
| `REASONING_EFFORT` | max | 推理强度 |
| `ENABLE_THINKING` | 0 | 在请求中加入 `enable_thinking: true` |
| `BREAK_SCHEMA_CYCLES` | 1 | 去除工具 schema 中的循环引用 |
| `MAX_OUTPUT_TOKENS` | 0（不限制） | 单次请求的输出 token 上限 |
| `MAX_CONTEXT`、`MAX_STEPS`、`TASK_TIMEOUT_S` | 262144、100、7200 | 单任务上限：token、步数、秒 |
| `MAX_CONCURRENT` | 1 | 并行任务数 |
