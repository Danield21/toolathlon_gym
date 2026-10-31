# MS-Eval · deepseek-v4-flash-0731 / spec

单 Agent 对照实验见同模型的 [sa 目录](../sa/README.md)。

固定评测 **166 个 case**，清单见本目录 `tasks.txt`。`KIMI_CODE_SUBAGENT=spec`：plan + 6 个领域专家。任务描述读取 `tasks/finalpool/<case>/docs/task.md`。

## 角色与提示词

- **角色：** 启动器固定 `KIMI_CODE_SUBAGENT=spec`，加载 [subagents_spec](../../../kimi_harness/assets/subagents_spec/) 中的定义。
- **提示词：** 使用 `subagent_types_spec.md` 和 `examples_spec.md`；plan-first 关闭。
- **覆盖规则：** `run.sh` 清除继承的 roster 和 prompt 变量，再应用本目录的固定设置；在 `config.env` 中填写 `KIMI_CODE_SUBAGENT`、`KIMI_PLAN_FIRST` 或 `KIMI_*_FILE` 不会覆盖这些设置。默认不插入额外 coordination 指令。

模式与环境变量的对应关系见[评测模式](../../README.zh-CN.md#评测模式)，无需手动配置这些开关。

## Docker 评测

使用已构建的 MS-Eval 镜像时，在本目录执行 `bash run.sh --backend docker`。
API 配置、运行选项和结果说明见[公共评测说明](../../README.zh-CN.md#运行评测)。

## 集群启动

先按 [Enroot 构建指南](../../../../enroot/README.zh-CN.md)准备镜像，并加载生成的 `env.sh`。

```bash
./run.sh --dry-run                 # 校验 166 个任务与模式，无需 key，不调用 API
cp config.env.example config.env
chmod 600 config.env
# 编辑 config.env，填写 API_BASE_URL、API_KEY，并确认 MODEL_NAME 与接口支持的名称一致
./run.sh --check                   # 检查本机依赖与配置，不调用 API
./run.sh                           # 前台运行全部 166 个 case
```

`API_BASE_URL` 填 API base（可以是根地址、`/v1` 或接口要求的其他前缀，不含 `/chat/completions` 或 `/messages`），
`API_KEY` 填自己的凭据。对于不鉴权的本地服务，也需填非空占位值。
默认 `API_FORMAT=openai` 表示 OpenAI-compatible wire format；原生 Messages 接口使用 `anthropic`。
后者通过 harness 配置 thinking，不注入 OpenAI 专属参数。
脚本只加载本目录的 `config.env`（优先于同名环境变量）；没有该文件时可直接 export 配置。
不会读取旧 `run_sh`、其他模型目录或全局 API 配置。

## 评测设置

默认 API 模型 ID 为 `deepseek-v4-flash`（[官方文档](https://api-docs.deepseek.com/quick_start/pricing)）。
`model.json` 固定本目录的模式及默认设置；`config.env` 可调整模型接口别名和采样参数。

**版本说明：** 官方文档说明 `deepseek-v4-flash` 已转到 DeepSeek-V4.1-Flash。
目录名中的 `0731` 是原实验标识，这个 API ID 不能锁定该历史快照。
如需复现 0731，请使用保留该快照的接口，并在 `config.env` 中填写它要求的模型 ID。
这里保留旧模型 ID，不自动改用其他模型名称。

| 设置 | 默认值 |
|---|---|
| temperature | 模型默认 |
| top_p | 模型默认 |
| reasoning_effort | high |
| enable_thinking 注入 | 0 |
| 工具 schema 循环引用展开 | 0 |
| 请求中的输出 token 上限 | 0（`0` 关闭限制；不会补写缺失字段） |
| context / steps / 每 case 超时 | 262144 / 100 / 7200 秒 |
| 并发 case 数 | 1，通过 `MAX_CONCURRENT` 调整 |

所有模式使用同一份固定的 166-case 任务清单，完整性校验由启动器内部维护。
`finalpool` 还含有其他任务，脚本只运行此清单；修改、缺失或重复清单会报错。
Oracle 会进一步检查 166 份 brief 与清单严格一致。非 Oracle 模式不依赖 oracle 数据目录。
