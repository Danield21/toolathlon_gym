# 踩点率 v2 定义与实现（P0-P4）

> 版本：2.2.5 · 2026-08-30
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

## 3. P2：粗判 v2（有序 wave contract）

1. 将已验证 eligible 的 GT waves 与 accepted fresh waves 做保序对齐；跳过 wave 可以，交叉/换序不可以。对齐目标按 `contract hits > 语义诊断对数 > 相似度` 字典序最大化，避免后续多余 Wave 抢走本应命中的对齐。
2. 对齐候选同时比较 action set、节点 1:1 粗语义覆盖、domain、write role 和 fan-out 比例。“任意 read 节点相同”不再能占领整个 wave。
3. 对齐后仍需满足全部计划 action、write 动作的双向一致性、全部显式 write roles、显式 dependency 与 exact/range fan-out，才算 wave hit。dependency 通过已选 alignment 比较精确目标，而不只比较“是否非 main”。实际多出 write、漏 compute、漏任一计划写入角色，都只保留语义对齐以便诊断，不算 contract hit。
4. 多出的 accepted fresh wave 是 unplanned wave，进 Precision 分母。Rejected call 不进分母。

```text
Planned Wave Recall = contract-hit eligible GT waves / eligible GT waves
Fresh Wave Precision = contract-hit actual fresh waves / accepted actual fresh waves
```

## 4. P3：细判 v2（wave 内节点 1:1）

1. 只在 P2 已保序对齐的 wave 对内做节点匹配。
2. 节点同时比较 action、agent type、domain、write role、target、object ID 和 shard key；显式 agent type 不一致、写动作不一致、具体写角色冲突或显式 object ID 冲突会直接 veto。`other-write` 只表示未解析出具体类型，不伪装成具体角色。
3. 每个计划实例与实际实例最多匹配一次。AgentSwarm `items` 逐项展开；计划头明确 `exactly 22`即使只写一个同模板 owner，也展开为 22 个实例。runtime-dependent fan-out 只使用委派前 manifest 证据固定计划分母，绝不读实际 fresh nodes 的数量。若模型把同一计划 Wave 拆成多个 fresh Wave，节点仍只能在已对齐的 Wave 内命中，不会跨 Wave 偷配。
4. 不再将同一 domain 的多个 read owner 全局去重为 1 点。

```text
Fine Node Recall = matched eligible planned node instances / eligible planned node instances
Fine Node Precision = matched accepted fresh node instances / accepted fresh node instances
```

## 5. P4：四组指标与开销分离

正式报告不再用一个“踩点率”混合表达，而是并列：

- Planned Wave Recall
- Fresh Wave Precision
- Fine-grained Node Recall
- Fine-grained Node Precision
- Resume/Repair Overhead（accepted resume、linked/orphan resume、repair fresh wave/node、rejected call）

Legacy 宽/严判数字仅在报告的“新旧口径对照”中显示，不与 v2 分子分母混算。

## 6. 产物

- `caidian_v2_report.md`
- `caidian_v2_summary.json`
- `caidian_v2_tasks.csv`
- `caidian_v2_detail.json`
- `caidian_v2_source_manifest.tsv`
- `SHA256SUMS`
- `runtime_prompts/<task>.task.md`

其中 source manifest 含每条 raw stream 的精确路径、大小、SHA-256，以及 runtime/reference prompt 哈希与来源。输出先写入同文件系统的 staging 目录，完整生成后才通过 rename 发布；替换旧 v2 产物时保留上一个完整目录，不会得到新旧文件混合的报告。
