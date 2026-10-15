# MS-Eval · kimi-k3 / spec

固定评测 **166 个 case**，清单见本目录 `tasks.txt`。`KIMI_CODE_SUBAGENT=7`：plan + 6 个领域专家。任务描述读取 `tasks/finalpool/<case>/docs/task.md`。

## 启动

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

默认 API 模型 ID 为 `kimi-k3`（[官方文档](https://platform.moonshot.ai/docs/guide/kimi-k3-quickstart)）。
`model.json` 固定本目录的模式及默认设置；`config.env` 可调整模型接口别名和采样参数。

| 设置 | 默认值 |
|---|---|
| temperature | 1 |
| top_p | 0.95 |
| reasoning_effort | max |
| enable_thinking 注入 | 0 |
| 工具 schema 循环引用展开 | 1 |
| 请求中的输出 token 上限 | 0（`0` 关闭限制；不会补写缺失字段） |
| context / steps / 每 case 超时 | 262144 / 100 / 7200 秒 |
| 并发 case 数 | 1，通过 `MAX_CONCURRENT` 调整 |

所有模式使用同一份固定的 166-case 任务清单，完整性校验由启动器内部维护。
`finalpool` 还含有其他任务，脚本只运行此清单；修改、缺失或重复清单会报错。
Oracle 会进一步检查 166 份 brief 与清单严格一致。非 Oracle 模式不依赖 oracle 数据目录。
模式开关由启动器固定，清除外部遗留的 Kimi roster、prompt 和采样变量。

## 运行依赖和隔离

启动文件全部在本目录：`run.sh` → `launch.py` → `runner.sh`，
`api_gateway.py` + `schema.py` 处理接口适配，`report.py` 检查完成情况。
可把整个目录复制到其他位置，再设置 `PROJECT_ROOT` 和 `RUNTIME_ROOT`；没有跨模型软链接。

评测数据和重型运行环境不复制进启动目录，需要先准备：

- 完整的 Toolathlon 项目：`tasks/finalpool`、`db/init.sql.gz`、`kimi_harness`、`utils`、MCP 服务及其构建产物。
- 已构建的 Toolathlon Enroot 镜像（包含 Python/MCP 依赖、Node、工具和 `/workspace/main.py`），配置 `AGENT_SQSH` 或 `AGENT_TEMPLATE`。
- Kimi CLI 发行目录 `KIMI_DIST`，包含 `dist/main.mjs`；必须支持现有 harness 的 3/7 子代理模式。
- 若镜像内的 uv Python symlink 依赖外部缓存，准备 `RUNTIME_ROOT/uv_python_cache`。
- Linux 普通用户，Bash、Python ≥3.10、PostgreSQL 15 的命令行工具、Enroot、squashfs-tools、rsync、curl、util-linux 和常规 GNU 工具可从 `PATH` 找到。
- 足够的 `/dev/shm` 空间：每个并发 case 需要独立 rootfs、数据库和工作目录；运行器会估算容量，默认预留每 case 8 GiB。

默认项目根从本目录位置推导；默认 runtime 为项目旁的 `toolathlon_gym_eval_dockers`。
Oracle 默认数据路径为项目旁的 `metrics_eval_bw/oracle166`。所有路径均可在配置中覆盖。
每个 case 创建独立 PostgreSQL 数据与 socket、独立可写 Enroot rootfs，完成后清理本次运行的资源。
固定 mock 服务端口通过本机共享锁避免冲突。不要以 root 启动 PostgreSQL。
若 Enroot 为用户目录安装，可设置 `ENROOT_INSTALL_ROOT`（包含 `bin/`、`lib/`、`etc/`）。
启动器会识别 `enroot-wrapper` 符号链接的实际目录，避免包装脚本循环调用；`--check` 会带超时检查 Enroot 版本。

API gateway 与评测进程运行在同一台机器，监听 loopback，使用每次运行随机生成的内部 token。
真实 key 和上游 URL 只交给 gateway；agent 及其配置只看到本机 gateway。
`config.env` 和 `results/` 已忽略，启动器不打印真实 key 或上游 URL。
如在 Slurm 上运行，在已申请的计算节点执行 `./run.sh`；该节点必须能直连配置的 API，或配置可达的 `HTTPS_PROXY`。
本目录不绑定集群分区、节点、内部网络地址，也不自动启动远程中继。

## 输出和完成判断

结果位于 `results/<run-id>/`（或 `OUTPUT_ROOT/<run-id>/`）：

- `run.json`：不含凭据/endpoint 的实际评测设置；`tasks.txt`：该次 166 个任务清单。
- `summary_parallel_<run-id>.csv`：逐 case 状态、退出码、时长和输出路径。
- `<case>/<run-id>_slot<N>/`：`run.log`、数据库检查日志、harness 轨迹与 evaluation 产物。

```bash
python3 report.py results/<run-id>
```

报告会检查 166 个 case 是否各有一条结果，指出缺失、重复、基础设施失败和未调度任务。
`success` / `case_failed` 是 harness 执行状态，不等同于最终评分；分数以每个 case 的 evaluation 产物为准。
基础设施失败、结果丢失或 Slurm 时间不足导致未完成时，运行器/报告返回非零；错误答案属于有效评测结果。
`--check` 是本地预检，不验证 API 可用性；没有偷偷发送试探请求，也不会代替实际的完整评测。
