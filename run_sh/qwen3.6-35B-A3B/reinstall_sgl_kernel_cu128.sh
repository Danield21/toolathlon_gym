#!/usr/bin/env bash
# On login01: replace CUDA13-linked sglang-kernel with cu128 build if available,
# else keep 0.4.5 and rely on LD_LIBRARY_PATH + --disable-cuda-graph.
set -euo pipefail
export http_proxy=http://127.0.0.1:7890 https_proxy=http://127.0.0.1:7890
ENV=/storage/lintaoLab/bowending/miniconda3/envs/bowen_sglang_qwen36
LOG=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/qwen3.6-35B-A3B/logs/reinstall_sgl_kernel_cu128.log
exec > >(tee -a "$LOG") 2>&1
echo "[$(date)] try install sgl-kernel cu128 latest from docs.sglang.io"
# Latest cu128 wheel on the index (may be older than 0.4.5)
"$ENV/bin/python" -m pip install --force-reinstall --no-deps \
  'sgl-kernel==0.3.14.post1+cu128' \
  -f https://docs.sglang.io/whl/cu128/sgl-kernel/ \
  2>&1 || echo "cu128 0.3.14.post1 install failed; keep existing kernel"
"$ENV/bin/python" -m pip show sglang-kernel sgl-kernel 2>&1 | head -20 || true
ldd "$ENV"/lib/python3.12/site-packages/sgl_kernel/sm90/common_ops.abi3.so 2>&1 | grep -E 'nvrtc|cudart' | head -10 || true
echo "[$(date)] done"
