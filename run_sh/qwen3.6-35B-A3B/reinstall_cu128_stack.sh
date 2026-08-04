#!/usr/bin/env bash
# login01 only — build a CUDA-12.8-only stack for gnho019 (driver 555).
set -euo pipefail
export http_proxy=http://127.0.0.1:7890 https_proxy=http://127.0.0.1:7890
ENV=/storage/lintaoLab/bowending/miniconda3/envs/bowen_sglang_qwen36
LOG=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/qwen3.6-35B-A3B/logs/reinstall_cu128_stack.log
mkdir -p "$(dirname "$LOG")"
exec > >(tee -a "$LOG") 2>&1
echo "[$(date)] reinstall cu128-only stack"

# Drop CUDA-13 runtime packages that poison LD_LIBRARY_PATH / Triton
"$ENV/bin/python" -m pip uninstall -y \
  nvidia-cuda-runtime-cu13 nvidia-cublas-cu13 nvidia-cudnn-cu13 \
  nvidia-cufft-cu13 nvidia-curand-cu13 nvidia-cusolver-cu13 nvidia-cusparse-cu13 \
  nvidia-nvjitlink-cu13 nvidia-nvtx-cu13 nvidia-nccl-cu13 nvidia-nvshmem-cu13 \
  nvidia-cusparselt-cu13 cuda-toolkit 2>/dev/null || true
rm -rf "$ENV/lib/python3.12/site-packages/nvidia/cu13" || true

# Torch cu128
"$ENV/bin/python" -m pip install --force-reinstall \
  'torch==2.9.1' 'torchvision==0.24.1' 'torchaudio==2.9.1' \
  --index-url https://download.pytorch.org/whl/cu128

# SGLang 0.5.10.post1 (Qwen3.6 min) + matching cu128 kernel
"$ENV/bin/python" -m pip install --force-reinstall --no-deps 'sglang==0.5.10.post1' \
  -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn

"$ENV/bin/python" -m pip install --force-reinstall --no-deps \
  'sgl-kernel==0.3.14.post1+cu128' \
  -f https://docs.sglang.io/whl/cu128/sgl-kernel/ || \
"$ENV/bin/python" -m pip install --force-reinstall --no-deps \
  'sglang-kernel==0.4.1' \
  -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn

# Core deps commonly needed (best-effort)
"$ENV/bin/python" -m pip install -U \
  'transformers>=4.57.0' 'fastapi' 'uvicorn' 'orjson' 'pydantic' 'numpy' \
  'sentencepiece' 'protobuf' 'tiktoken' 'einops' 'modelscope' \
  -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn \
  || true

# Ensure torch stayed cu128
"$ENV/bin/python" -m pip install --force-reinstall --no-deps \
  'torch==2.9.1+cu128' 'torchaudio==2.9.1+cu128' 'torchvision==0.24.1+cu128' \
  --index-url https://download.pytorch.org/whl/cu128

"$ENV/bin/python" -c 'import sglang,torch; print("OK", sglang.__version__, torch.__version__, torch.version.cuda)'
ldd "$ENV"/lib/python3.12/site-packages/sgl_kernel/sm90/common_ops.abi3.so 2>&1 | grep -E 'nvrtc|cudart' | head -10 || true
echo "[$(date)] done"
