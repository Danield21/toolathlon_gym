# 踩点率 v2 定义与实现（P0-P4）

> 版本：2.3.0 · 2026-08-30
> 实现：`scripts/caidian_v2.py` + `scripts/gen_caidian_v2.py`
> Legacy 口径保留在 `backups/caidian_legacy_20260830T190032+0800/`，v2 不覆盖其脚本或报告。

## 1. P0：先固定证据语义

1. 仅当 Agent/AgentSwarm 的 tool result 证实 harness 已接受请求时才计数。缺失 result、参数/schema 拒绝、ToolError 不进入实际 wave。子代理接受后因 maxSteps 等失败，仍是真实委派。
2. `resume=<agent_id>` 链回原 agent/node/wave，不另建新 wave；找不到原 agent 的 resume 记为 orphan。
3. 没有 resume ID 但语义上是 continue/retry/repair 的新委派仍是 fresh wave，另外标记 repair，不把它隐藏掉。
4. 轨迹严格根据 `SOURCE_MAP.tsv:target_run` 定位；一题必须恰好有一个非空壳 `raw_stream.jsonl`。不再按 mtime 选 slot。
5. 计划侧 prompt 优先取该条轨迹的 `traj_log.json:config.task_str`；旧 harness 若将 traj_log 覆写为 evaluator 结果，则从 `run.log` 的 Kimi 启动命令提取。两者皆不存在时才 fallback 到当前参考 prompt，并单列数量。
6. 条件 wave 使用 eligible/N/A 语义。只有经轨迹证据确认的 runtime gate 才可从 Recall 分母排除；每个排除都有版本化 override、原因和证据。
7. runtime-dependent fan-out 不得由已接受的 agent 数量、Agent/AgentSwarm 参数或委派文本反推。五个此类 Wave 由第一次委派之前的成对结构化 Canvas tool results 执行确定性派生（课程数、分页数、每四页分块、有 quiz 课程数）；第六个两分片任务的 exact=2 直接由 prompt 明说的 `two shards` 得到。override 同时绑定该 `raw_stream.jsonl` 的 SHA-256，证据行必须早于首个 Agent 调用；多模板 Wave 的 page/teacher multiplicity 也必须分别等于结构化派生的具名组件，不只校验两者之和。
8. JSONL 中只白名单保留已知 schema 日志行；任意其他非 JSON 行、非 object 记录、非 object/损坏的 tool arguments、重复 Agent 请求/result ID 或 provider termination 使该题 fail closed，不进入任何指标的分子或分母。

## 2. P1：统一 IR

计划节点与实际节点共用下列字段：

- `wave_id`, `ordinal`, `dependency`
- `agent_type`: explore / coder / plan / unspecified
- `actions`: read / compute / write 的多标签集合
- `domains`: Canvas、WooCommerce、arXiv、workspace 等真实数据域
- `write_roles`: excel / word / ppt / pdf / notion / email / calendar / gsheet / gform / json 等多标签
- `targets`: 文件名、收件人、显式标题
- `object_ids`: course_id、course_code、paper_id、row/page range、AgentSwarm item
- `shard_key`: course / paper / page / teacher / supplier / destination
- `fanout`: exact 或 `[min,max]` 区间
- `condition`: always / conditional + eligible + evidence

节点是多动作的：例如“读取分页、聚合、写中间 JSON”是 `{read,compute,write}`，不再被压成单一 read 或 write。实际 Wave 也完整保留 `dependency/fanout/condition`：同一 assistant response 内为并行，后续 response 中的 fresh Wave 记录已发生的串行依赖。计划依赖只从 `After/Once/Following/...` 等明示转换提取；`After Wave 1` 保留精确目标 `wave_1`，不再简化成“前一 Wave”。
分片只表示“一个 agent 被分配的单位”：分页参数 `page=1`、交付物中的论文/行数和共享 manifest 不会自动变成分片键。范围编号如 `1-2`/`3-6` 会保留为 2/4 个节点实例，混合 Coder/Explore Wave 不再被压成单一角色。

## 3. P2：粗判 v2（有序 Wave 语义兼容）

1. 将已验证 eligible 的 GT Waves 与 accepted fresh Waves 做保序对齐；可跳过，不可交叉或换序。对齐目标按 `细判 contract hits > 粗判语义对数 > 相似度` 字典序最大化，避免宽松对齐抢走本应严格命中的 Wave。
2. 粗命中必须是 `semantic_compatible`：至少存在节点 1:1 语义匹配，template coverage ≥ 0.50、planned coverage ≥ 0.34、action coverage ≥ 0.34，加权相似度 ≥ 0.48。相似度同时考虑 action、domain、write role、fan-out 比例和节点覆盖，不会因一个泛化 read 节点就命中。
3. 粗判允许 fan-out、依赖、个别 action 或具体对象存在偏差；这些偏差会使细判失败。多出的 accepted fresh Wave 仍进 Precision 分母，Rejected call 不进分母。

```text
Coarse Recall = semantic-compatible planned Waves / eligible planned Waves
Coarse Precision = semantic-compatible actual Waves / accepted fresh Waves
```

## 4. P3：细判 v2（粗命中的严格子集）

1. 细命中必须首先是粗命中的同一个保序 Wave 对。
2. 该 Wave 还必须满足：全部计划 action、write 动作双向一致、全部显式 write roles、精确 dependency 目标、exact/range fan-out，以及全部计划节点的严格 1:1 覆盖。
3. 节点同时比较 action、agent type、domain、write role、target、object ID 和 shard key；显式角色、写动作或 object ID 冲突直接 veto。每个计划/实际实例最多匹配一次。
4. AgentSwarm `items` 逐项展开；`exactly 22` 展开为 22 个实例。runtime-dependent fan-out 只由委派前 manifest 固定。任一计划节点缺失、错误分片或依赖偏差都不算细命中；多出的实际 Wave 降低 Precision，不撤销已对齐 Wave 的 Recall 命中。

```text
Fine Recall = strict-contract planned Waves / eligible planned Waves
Fine Precision = strict-contract actual Waves / accepted fresh Waves
```

`Fine hits ⊆ Coarse hits`，两者使用相同 Wave 分母，因此 Recall 和 Precision 均保证 `Fine ≤ Coarse`。

## 5. P4：只保留粗/细两类踩点率

正式报告只展示：

- 粗踩点率（Recall / Precision）
- 细踩点率（Recall / Precision）

节点局部覆盖仅保留在 alignment 证据中，不汇总成第三类正式指标。Resume/repair/reject 是证据诊断，不是第三类踩点率。Legacy 宽/严数字只用于回放对照，不与 v2.3 混算。

## 6. 产物

- `caidian_v2_report.md`
- `caidian_v2_summary.json`
- `caidian_v2_tasks.csv`
- `caidian_v2_detail.json`
- `caidian_v2_source_manifest.tsv`
- `SHA256SUMS`
- `runtime_prompts/<task>.task.md`

其中 source manifest 含每条 raw stream 的精确路径、大小、SHA-256，以及 runtime/reference prompt 哈希与来源。输出先写入同文件系统的 staging 目录，完整生成后才通过 rename 发布；替换旧 v2 产物时保留上一个完整目录，不会得到新旧文件混合的报告。
