# MS-Eval · kimi-k3 / sa

单 Agent （SA）基线，运行本目录 `tasks.txt` 中固定的 **166 个任务**。
模型、API 格式、采样默认值、任务描述和评分配置与同模型的 [spec](../spec/README.md) 一致。
启动器固定 `KIMI_CODE_SUBAGENT=none`，关闭子 Agent 委派和 plan-first，不加载子 Agent 提示词及示例。

## 配置模型 API

在本目录执行：

```bash
cp config.env.example config.env
chmod 600 config.env
```

在 `config.env` 中填写 `API_BASE_URL` 和 `API_KEY`，按服务商要求确认 `MODEL_NAME`。
本目录有独立的配置文件，不会自动读取 `../spec/config.env`。
对比 SA 与 spec 时，两边应使用相同的模型、采样参数、步数及超时设置。
更改 `KIMI_CODE_SUBAGENT` 或 `KIMI_PLAN_FIRST` 不会覆盖固定的 SA 模式。

## Docker 评测

镜像准备、运行选项和结果说明见[公共评测说明](../../README.zh-CN.md#运行评测)。
在本目录执行，无需再加 `--single-agent`：

```bash
bash run.sh --backend docker --dry-run
bash run.sh --backend docker --task canvas-course-comparison
bash run.sh --backend docker
```

## 集群评测

先按 [Enroot 构建指南](../../../../enroot/README.zh-CN.md)准备镜像，并加载生成的 `env.sh` 后：

```bash
bash run.sh --backend enroot --dry-run
bash run.sh --backend enroot --check
bash run.sh --backend enroot
```

`--dry-run` 和 `--check` 不调用模型 API；省略 `--backend` 时使用集群后端。
