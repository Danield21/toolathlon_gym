#!/usr/bin/env bash
# Install torch 2.11.0+cu128 + sglang on login01 (gnho019 has no net; driver 555 needs cu128 not cu130)
set -euo pipefail
export http_proxy=http://127.0.0.1:7890 https_proxy=http://127.0.0.1:7890
export HTTP_PROXY=$http_proxy HTTPS_PROXY=$https_proxy
ENV=/storage/lintaoLab/bowending/miniconda3/envs/bowen_sglang_qwen36
LOG=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/qwen3.6-35B-A3B/logs/fix_torch_cu128_v2.log
mkdir -p "$(dirname "$LOG")"
exec > >(tee -a "$LOG") 2>&1
echo "[$(date)] start torch2.11+cu128"
"${ENV}/bin/python" -m pip install --force-reinstall \
  'torch==2.11.0' 'torchvision' 'torchaudio' \
  --index-url https://download.pytorch.org/whl/cu128
printf '%s\n' 'torch==2.11.0+cu128' 'torchaudio==2.11.0+cu128' > /tmp/sglang_cu128_constraints.txt
"${ENV}/bin/python" -m pip install -U 'sglang[all]>=0.5.10' \
  -c /tmp/sglang_cu128_constraints.txt \
  -i https://pypi.tuna.tsinghua.edu.cn/simple \
  --trusted-host pypi.tuna.tsinghua.edu.cn \
  --extra-index-url https://download.pytorch.org/whl/cu128
"${ENV}/bin/python" -m pip install --force-reinstall \
  'torch==2.11.0' 'torchaudio' 'torchvision' \
  --index-url https://download.pytorch.org/whl/cu128
"${ENV}/bin/python" -c 'import sglang,torch; print("OK", sglang.__version__, torch.__version__, torch.version.cuda)'
echo "[$(date)] done"
