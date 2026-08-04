# Qwen3.6-35B-A3B Toolathlon-GYM eval

## Layout

| Script | Role |
|--------|------|
| `config.env` | Model API + SGLang deploy knobs |
| `sbatch_sglang.sh` | Slurm job: hold gnho019 8×H800 + start SGLang |
| `start_sglang.sh` | Launch/reuse SGLang (offline-safe) |
| `stop_sglang.sh` | Kill SGLang on the node |
| `run_eval_parallel.sh` | Parallel isolated PG+enroot eval (all `tasks/finalpool`) |
| `run_orchestrator.sh` | Wait for SGLang healthy → full eval → case study |
| `EVAL_STABILITY_SPEC.md` | Stability fixes and staged acceptance gates before the 503-case run |

Official Qwen3.6 docs require **sglang ≥ 0.5.10**. Serving env:

`SGLANG_CONDA_ENV=/storage/lintaoLab/bowending/miniconda3/envs/bowen_sglang_qwen36`

(`bowen_verl2` stays at sglang 0.5.2 / torch 2.8 for training.)

## Network + driver constraints

- **login01** has outbound net (proxy `127.0.0.1:7890`). **gnho019 has no outbound net.**
- gnho019 driver is **555 / CUDA 12.5** → use **PyTorch cu128** wheels. Default `sglang[all]` often pulls `torch+cu130`, which fails with `driver too old`.

Install / upgrade packages **only on login01**. `start_sglang.sh` never runs `pip` on the compute node.

```bash
# on login01
export http_proxy=http://127.0.0.1:7890 https_proxy=http://127.0.0.1:7890
ENV=/storage/lintaoLab/bowending/miniconda3/envs/bowen_sglang_qwen36
$ENV/bin/python -m pip install --force-reinstall 'torch==2.11.0' torchvision torchaudio \
  --index-url https://download.pytorch.org/whl/cu128
printf '%s\n' 'torch==2.11.0+cu128' 'torchaudio==2.11.0+cu128' > /tmp/sglang_cu128_constraints.txt
$ENV/bin/python -m pip install -U 'sglang[all]>=0.5.10' \
  -c /tmp/sglang_cu128_constraints.txt \
  -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn \
  --extra-index-url https://download.pytorch.org/whl/cu128
```

## Deploy on gnho019

```bash
mkdir -p run_sh/qwen3.6-35B-A3B/logs
sbatch run_sh/qwen3.6-35B-A3B/sbatch_sglang.sh
# wait until (from login01):
curl -sS http://gnho019:30002/v1/models
```

## Eval (login01 tmux)

```bash
source toolathlon_gym_eval_dockers/env.sh
conda activate toolathlon_gym
bash run_sh/qwen3.6-35B-A3B/run_orchestrator.sh
```

Sampling is greedy (`MODEL_GREEDY=1`, `temperature=0`, `top_p=1`).

### Enroot isolation guarantees

`run_eval_parallel.sh` gives every task all of the following:

- a leased, unique host PostgreSQL TCP port;
- a unique node-local `PGDATA` directory;
- a unique PostgreSQL Unix-socket directory;
- an independent Enroot rootfs/workspace;
- schema and `data_directory` identity checks before MCP startup;
- task-local cleanup on success, failure, interrupt, or parent shutdown.

Port leases are atomic across separate evaluator processes, not just workers
inside one invocation. Both `PGPORT` and the equivalent `PG_PORT` variables are
passed into the Enroot agent so preprocess, MCP servers, and evaluation use the
same task database.

Run the fast PostgreSQL-only concurrency regression test with:

```bash
MAX_CONCURRENT=2 PG_TEST_ONLY=1 \
  bash run_sh/qwen3.6-35B-A3B/run_eval_parallel.sh \
  12306-beijing-shanghai-trip-notion-gcal-word \
  12306-canvas-fieldtrip-gcal-word-email
```

Run a short full-stack smoke with:

```bash
MAX_CONCURRENT=2 MAX_STEPS=3 \
  bash run_sh/qwen3.6-35B-A3B/run_eval_parallel.sh \
  12306-beijing-shanghai-trip-notion-gcal-word \
  12306-canvas-fieldtrip-gcal-word-email
```

The PostgreSQL runtime defaults to `/tmp/toolathlon_pg_<uid>` so database
initialization does not pay the NFS write penalty. Enroot rootfs copies still
live under `ENROOT_DATA_PATH`; on this cluster those NFS copies are the main
startup cost.

## Outputs

```text
dumps/qwen3.6-35B-A3B/
  <task>/<runid>_slotN/
  summary_parallel_<ts>.csv
  CASE_STUDY_latest.md
```
