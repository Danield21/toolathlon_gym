# deepseek-v4-pro evaluation

## Concurrency on login01

Measured on `login01`: 96 CPUs, load ≈ 9–10, ~166 GiB mem available, dumps/rootfs on NFS `/lintaoLab2`.

| Setting | Value | Why |
|---------|-------|-----|
| **Recommended** | **3** | Matches your 3 tasks; polite to shared login node |
| Soft max in script | 5 | Beyond this, parallel `cp -a` of 3.8G rootfs hammers NFS |
| Avoid | ≥8 on login | Prefer a compute node for larger sweeps |

Safe parallelism = **one PostgreSQL per task** (unique port + pgdata) + **one enroot rootfs per task**. Shared-PG parallel is unsafe (email/notion/gcal wipe).

## Sampling

`config.env` sets greedy decoding:

```bash
MODEL_GREEDY=1
MODEL_TEMPERATURE=0
MODEL_TOP_P=1
```

## Run (safe parallel)

```bash
source /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers/env.sh
conda activate toolathlon_gym

bash /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/deepseek-v4-pro/run_eval_parallel.sh
```

Override concurrency:

```bash
MAX_CONCURRENT=3 bash run_sh/deepseek-v4-pro/run_eval_parallel.sh
```

## Run (sequential, shared PG)

```bash
bash run_sh/deepseek-v4-pro/run_eval.sh
```

## Outputs

```text
dumps/deepseek-v4-pro/
  <task>/<runid>_slotN/run.log
  <task>/<runid>_slotN/...
  summary_parallel_<ts>.csv
```

Edit API settings in `config.env` (gitignored).
