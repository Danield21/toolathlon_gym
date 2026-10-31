# 评测启动脚本

[English](README.md) | 简体中文

每个 `<model>/<variant>/` 目录对应一组配置，运行固定的 166 个 MS-Eval 任务，结果写入 MS-Eval 根目录下的 `results/<model>/<variant>/`。

## 快速开始

先构建镜像（[Docker](../../docker/README.zh-CN.md) 或 [Enroot](../../enroot/README.zh-CN.md)），然后在 MS-Eval 根目录执行：

```bash
cd toolathlon_gym/run_sh_eval/kimi-k3
cp spec/config.env.example spec/config.env   # 填写 API_BASE_URL 和 API_KEY
chmod 600 spec/config.env
cp -p spec/config.env sa/config.env          # SA 基线使用相同配置

bash spec/run.sh --backend docker --dry-run                         # 检查配置，不调用 API
bash spec/run.sh --backend docker --task canvas-course-comparison   # 先试跑一个任务
bash spec/run.sh --backend docker                                   # MAD-SAE，全部 166 个任务
bash sa/run.sh --backend docker                                     # SA 基线，全部 166 个任务
```

随后[计算指标](../../eval/README.zh-CN.md)。在集群上，先在 MS-Eval 根目录执行 `source .local/enroot/env.sh`，再改用 `--backend enroot`；Enroot 每次运行全部 166 个任务。

## 预设

| 模型 | 模式 |
| --- | --- |
| `claude-sonnet-5`、`glm-5.3`、`glm-5.3-flash`、`gpt-5.6-sol`、`kimi-k3`、`minimax-m3`、`qwen-3.8-max` | `sa`、`spec` |
| `deepseek-v4-flash-0731` | `sa`、`spec`、`gen`、`oracle`、`plan1st` |

| 模式 | 范式 | 子 Agent | 说明 |
| --- | --- | --- | --- |
| `sa` | SA | 无 | 其他模式的对照基线 |
| `spec` | MAD-SAE | `plan` + 6 个领域专家 | |
| `gen` | MAD-SAE | `coder`、`explore`、`plan` | |
| `oracle` | MAD-SAE | 同 `gen` | 任务说明附带委派先验（[`data/oracle166`](../../data/oracle166/)） |
| `plan1st` | MAD-SAE | 同 `gen` | 主 Agent 先向 `plan` 获取计划再行动 |

各模式的子 Agent 组合和提示词由启动脚本固定，`config.env` 无法修改。Agent 配置与提示词说明见 [kimi_harness](../kimi_harness/README.md)。

## 配置

每个预设读取各自的 `config.env`（已被 Git 忽略），由 `config.env.example` 复制而来。默认值即论文使用的配置，各预设的 README 中列有详细取值。

| 配置项 | 说明 |
| --- | --- |
| `API_BASE_URL`、`API_KEY` | API 根地址（不含 `/chat/completions` 或 `/messages`）及密钥，必填 |
| `API_FORMAT` | `openai`（OpenAI 兼容接口）或 `anthropic`（Messages 接口） |
| `MODEL_NAME` | API 模型 ID，仅当接口使用其他 ID 时修改 |
| `TEMPERATURE`、`TOP_P`、`REASONING_EFFORT` 等 | 请求参数；留空表示使用模型默认值 |
| `MAX_STEPS`、`TASK_TIMEOUT_S`、`MAX_CONTEXT` | 单任务上限：100 步、7200 秒、262,144 token |
| `MAX_CONCURRENT` | 并行任务数，建议从 `1` 开始 |
| `OUTPUT_ROOT`、`HTTPS_PROXY` | 可选：结果父目录；访问模型 API 时使用的代理 |

文件末尾注释掉的运行路径是 Enroot 的覆盖项，Docker 不使用。

## 参数

| 参数 | 说明 |
| --- | --- |
| `--backend docker` / `enroot` | 运行后端，默认 `enroot` |
| `--config FILE` | 使用其他配置文件 |
| `--task NAME` | 运行指定任务，可重复（仅 Docker） |
| `--output DIR` | 结果父目录（仅 Docker；`OUTPUT_ROOT` 对两种后端均有效） |
| `--dry-run` | 打印配置后退出 |
| `--check` | 检查运行环境和配置，不调用 API |
| `--smoke` | 离线检查数据库、MCP 服务和网络，不调用 API（仅 Docker） |

`bash run.sh --backend docker --help` 可查看全部参数。

## 输出

```text
results/<model>/<variant>/<run-id>/
├── run.json, tasks.txt              # 配置与任务清单
├── summary_parallel_<run-id>.csv    # 各任务的执行状态
├── audit_index.html                 # 全部任务概览
└── <task>/<run-id>_slotN/           # 日志、Agent 会话、audit.html
```

每次运行都会新建 `<run-id>`，不会覆盖已有结果。需要重跑部分任务时，用 `--task` 再运行一次即可，指标脚本以每个任务最近一次的运行为准。在预设目录中执行 `python3 report.py <run-dir>`，可检查一次运行是否覆盖了全部 166 个任务。
