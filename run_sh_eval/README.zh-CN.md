# 评测启动说明

[English](README.md)

选择模型及评测模式对应的目录，运行固定的 166 个任务。各启动器会明确指定该模式的子 Agent 组合和提示词。

- **Docker：** 构建 MS-Eval 镜像后，按[运行评测](#运行评测)操作，使用模型目录中的 `run.sh --backend docker` 入口。
- **集群：** 按 [Enroot 构建指南](../../enroot/README.zh-CN.md)准备镜像和运行文件。`run.sh` 默认使用集群后端，也可通过 `--backend enroot` 显式指定。

统一从 `<model_name>/<execution_mode>/run.sh` 启动评测。脚本会读取该预设的 `config.env`（或 `--config FILE` 指定的文件），并选择相应后端。`_common/` 和 `docker/` 存放内部支持代码。

## 目录结构

```text
run_sh_eval/
├── README.md
├── README.zh-CN.md
├── _common/
│   └── launch.sh               # 内部配置加载和后端选择逻辑
├── docker/                    # Docker 运行器、API 网关和容器入口
├── claude-sonnet-5/
│   ├── sa/                    # 单 Agent 基线
│   └── spec/                  # 专家子 Agent 模式
├── deepseek-v4-flash-0731/
│   ├── sa/                    # 单 Agent 基线
│   ├── spec/                  # 专家子 Agent 模式
│   ├── gen/                   # 通用子 Agent 模式
│   ├── oracle/                # 通用子 Agent + oracle 任务说明
│   └── plan1st/               # 通用子 Agent + 先规划
├── glm-5.3/
│   ├── sa/                    # 单 Agent 基线
│   └── spec/                  # 专家子 Agent 模式
├── glm-5.3-flash/
│   ├── sa/                    # 单 Agent 基线
│   └── spec/                  # 专家子 Agent 模式
├── gpt-5.6-sol/
│   ├── sa/                    # 单 Agent 基线
│   └── spec/                  # 专家子 Agent 模式
├── kimi-k3/
│   ├── sa/                    # 单 Agent 基线
│   └── spec/                  # 专家子 Agent 模式
├── minimax-m3/
│   ├── sa/                    # 单 Agent 基线
│   └── spec/                  # 专家子 Agent 模式
└── qwen-3.8-max/
    ├── sa/                    # 单 Agent 基线
    └── spec/                  # 专家子 Agent 模式
```

每个 `<model_name>/<execution_mode>/` 目录包含：

| 文件 | 用途 |
| --- | --- |
| `run.sh` | 当前模型及模式的评测入口 |
| `config.env.example` | API 和运行配置模板，复制为 `config.env` 后编辑 |
| `model.json`、`eval_config.json` | 模型默认参数、实验模式及评测配置 |
| `tasks.txt` | 固定的 166 个任务清单 |
| `launch.py`、`runner.sh` | 预设配置准备及集群后端 |
| `api_gateway.py`、`schema.py` | 模型 API 适配 |
| `report.py` | 检查完整评测的完成情况 |
| `README.md` | 当前预设的使用说明 |

运行时请保留完整目录结构。本地 `config.env` 和运行结果已被 Git 忽略。

## 评测模式

| 模式 | 子 Agent 组合 | 提示词片段 | 先规划 | 环境变量配置 |
| --- | --- | --- | --- | --- |
| `sa` | `none`：单 Agent | 不包含委派相关片段 | 关闭 | `KIMI_CODE_SUBAGENT=none` |
| `spec` | `spec`：plan 和六个领域专家 | `subagent_types_spec.md`、`examples_spec.md` | 关闭 | `KIMI_CODE_SUBAGENT=spec` |
| `gen` | `gen`：coder、explore、plan | `subagent_types_gen.md`、`examples_gen.md` | 关闭 | `KIMI_CODE_SUBAGENT=gen` |
| `oracle` | `gen`，使用 oracle 任务说明 | 与 gen 相同 | 关闭 | `KIMI_CODE_SUBAGENT=gen` |
| `plan1st` | `gen` | 在 gen 基础上加入 `plan_first_gen.md` | 开启 | `KIMI_CODE_SUBAGENT=gen`<br>`KIMI_PLAN_FIRST=1` |

表中环境变量由启动器自动设置；除 `plan1st` 外，均清除 `KIMI_PLAN_FIRST`，关闭先规划。

所有模型目录均提供 `sa` 和 `spec`，各自有独立的 `config.env.example`、`model.json`、`tasks.txt` 和 `run.sh`。SA 的模型、采样默认值和任务集合与同模型的 spec 预设一致，关闭子 Agent 委派和先规划。[DeepSeek 目录](deepseek-v4-flash-0731/) 还提供 `gen`、`oracle` 和 `plan1st`。子 Agent 定义与提示词片段位于 [kimi_harness/assets](../kimi_harness/assets/)。

`run.sh` 读取本目录的 `config.env` 后，由 `launch.py` 应用该实验固定的子 Agent 组合与提示词设置。在 `config.env` 中设置 `KIMI_CODE_SUBAGENT` 或 `KIMI_PLAN_FIRST` 不会覆盖预设。可在实验目录内运行 `./run.sh --dry-run` 查看配置，不调用模型 API。运行 SA 时，进入对应模型的 `sa/` 目录，执行 `bash run.sh --backend docker` 或 `bash run.sh --backend enroot`。对比 SA 与 spec 时，分别配置两者，并保持模型和采样设置一致。

运行模块职责和提示词文件布局见 [harness README](../kimi_harness/README.md)。

## 运行评测

先按 [Docker](../../docker/README.zh-CN.md) 或 [Enroot](../../enroot/README.zh-CN.md) 指南构建镜像。将 `<model_name>` 和 `<execution_mode>` 替换为对应目录名，从 MS-Eval 仓库根目录执行：

```bash
# execution_mode = sa, spec, gen, oracle, plan1st
cd toolathlon_gym/run_sh_eval/<model_name>/<execution_mode>
cp config.env.example config.env
# 编辑 config.env：填写 API_BASE_URL 和 API_KEY，确认 MODEL_NAME 和 API_FORMAT。
bash run.sh --backend docker --task canvas-course-comparison
bash run.sh --backend docker
```

使用 Enroot 时，先从仓库根目录执行 `source .local/enroot/env.sh`（自定义构建目录时替换路径），将上述启动命令改为 `bash run.sh --backend enroot`。该后端每次运行全部 166 题。

每个预设使用独立的 `config.env`；对比 SA 与 spec 时，请保持模型和采样设置一致。建议从 `MAX_CONCURRENT=1` 开始，并在同一文件中设置 `MAX_STEPS`、`TASK_TIMEOUT_S` 等评测限制。使用 Docker 时，可选的运行环境路径覆盖配置保持注释即可。

评测结果默认按模型预设目录名和评测模式保存在 MS-Eval 根目录的 `results/<model_name>/<execution_mode>/<run-id>/` 下，例如 `results/deepseek-v4-flash-0731/spec/<run-id>/`。每次运行使用独立的 `<run-id>`，避免覆盖已有结果。结果包含各任务的日志、执行轨迹和 `eval_res.json`，批次目录另有汇总 CSV。运行器正常退出表示评测流程已完成，任务得分请查看 `eval_res.json`。

## run.sh 的 Arguments

在模型对应的评测模式目录中，通过 `bash run.sh [参数]` 启动。常用参数如下：

| 参数 | 用途 |
| --- | --- |
| `--backend docker` / `--backend enroot` | 选择运行后端；省略时默认使用 Enroot |
| `--config FILE` | 指定配置文件；默认读取 `run_sh_eval/<model_name>/<execution_mode>/config.env` |
| `--task NAME` | 指定任务，可重复传入；省略时运行全部 166 题（仅 Docker 后端） |
| `--output DIR` | 指定结果父目录（仅 Docker）；优先级：`--output` > `config.env` 中的 `OUTPUT_ROOT` > 默认目录 |
| `--dry-run` | 查看预设配置，不启动评测或调用模型 API |
| `--check` | 检查运行环境和配置，不调用模型 API |
| `--help` | 查看所选后端支持的全部参数 |
