#!/usr/bin/env bash
set -euo pipefail
export http_proxy=http://127.0.0.1:7890 https_proxy=http://127.0.0.1:7890
ENV=/storage/lintaoLab/bowending/miniconda3/envs/bowen_sglang_qwen36
LOG=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/qwen3.6-35B-A3B/logs/finish_cu128.log
mkdir -p "$(dirname "$LOG")"
exec > >(tee -a "$LOG") 2>&1
echo "[$(date)] finish start"
"$ENV/bin/python" -c 'import torch; print("torch", torch.__version__, torch.version.cuda)'
# no-deps: keep already-installed CUDA 12.8 torch stack
"$ENV/bin/python" -m pip install --force-reinstall --no-deps 'sglang==0.5.16' \
  -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
"$ENV/bin/python" -m pip install --force-reinstall --no-deps \
  'torch==2.11.0+cu128' 'torchaudio==2.11.0+cu128' 'torchvision==0.26.0+cu128' \
  --index-url https://download.pytorch.org/whl/cu128
"$ENV/bin/python" -c 'import sglang,torch; print("OK", sglang.__version__, torch.__version__, torch.version.cuda)'
echo "[$(date)] finish done"
