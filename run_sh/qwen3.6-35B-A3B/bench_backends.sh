#!/usr/bin/env bash
# A/B backend smoke+throughput on the allocated GPU node (run via srun --overlap).
# Usage:
#   srun --jobid=JOB --overlap -n1 -c16 --mem=64G bash bench_backends.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck disable=SC1091
source "${SCRIPT_DIR}/config.env"

ENV="${SGLANG_CONDA_ENV}"
PY="${ENV}/bin/python"
PORT="${SGLANG_PORT:-30002}"
MODEL_PATH="${MODEL_PATH}"
NAME="${MODEL_NAME}"
LOG_DIR="${SCRIPT_DIR}/logs/backend_ab"
mkdir -p "$LOG_DIR" /storage/lintaoLab/bowending/tmp
export TMPDIR=/storage/lintaoLab/bowending/tmp
export CUDA_VISIBLE_DEVICES="${SGLANG_CUDA_VISIBLE_DEVICES:-0,1,2,3,4,5,6,7}"
export SGLANG_ENABLE_JIT_DEEPGEMM=0
export NVCC_PREPEND_FLAGS="${NVCC_PREPEND_FLAGS:-} -std=c++17"

_SP="${ENV}/lib/python3.12/site-packages"
export LD_LIBRARY_PATH=""
for d in \
  "${_SP}/nvidia/cuda_nvrtc/lib" \
  "${_SP}/nvidia/cuda_runtime/lib" \
  "${_SP}/torch/lib"
do
  [[ -d "$d" ]] && LD_LIBRARY_PATH="${d}${LD_LIBRARY_PATH:+:$LD_LIBRARY_PATH}"
done
export LD_LIBRARY_PATH
export PATH="${ENV}/bin:${PATH}"

kill_port() {
  local pids
  pids=$(ss -tlnp 2>/dev/null | awk -v p=":${PORT}" '$4 ~ p"$" {print}' | sed -n 's/.*pid=\([0-9]*\).*/\1/p' | sort -u)
  if [[ -z "${pids}" ]]; then
    # fallback: sglang python
    pids=$(pgrep -f "sglang.launch_server.*--port ${PORT}" || true)
  fi
  if [[ -n "${pids}" ]]; then
    echo "[kill] pids: ${pids}"
    # shellcheck disable=SC2086
    kill ${pids} 2>/dev/null || true
    sleep 5
    # shellcheck disable=SC2086
    kill -9 ${pids} 2>/dev/null || true
    sleep 3
  fi
  # also kill child multiprocess
  pkill -f "sglang.launch_server.*--port ${PORT}" 2>/dev/null || true
  sleep 2
}

wait_ready() {
  local tag="$1" max="${2:-120}"
  local i code
  for i in $(seq 1 "$max"); do
    code=$(curl -sS --noproxy '*' -o /dev/null -w '%{http_code}' --connect-timeout 2 \
      "http://127.0.0.1:${PORT}/v1/models" 2>/dev/null || echo 000)
    if [[ "$code" == "200" ]]; then
      echo "[${tag}] READY try=${i}"
      return 0
    fi
    if ! pgrep -f "sglang.launch_server.*--port ${PORT}" >/dev/null; then
      echo "[${tag}] DEAD before ready (try=${i})"
      return 1
    fi
    sleep 5
  done
  echo "[${tag}] TIMEOUT waiting ready"
  return 1
}

bench_once() {
  local tag="$1" max_tokens="${2:-256}"
  local out t0 t1 ms
  t0=$(date +%s%N)
  out=$(curl -sS --noproxy '*' --max-time 300 "http://127.0.0.1:${PORT}/v1/chat/completions" \
    -H 'Content-Type: application/json' \
    -d "{
      \"model\": \"${NAME}\",
      \"messages\": [{\"role\":\"user\",\"content\":\"Write a detailed step-by-step plan to organize a 3-day Beijing trip. Use bullet points. Do not use tools.\"}],
      \"temperature\": 0,
      \"max_tokens\": ${max_tokens},
      \"chat_template_kwargs\": {\"enable_thinking\": false}
    }" 2>&1) || true
  t1=$(date +%s%N)
  ms=$(( (t1 - t0) / 1000000 ))
  echo "$out" >"${LOG_DIR}/${tag}_resp.json"
  python3 - "$out" "$ms" "$tag" <<'PY'
import json,sys,re
raw, ms, tag = sys.argv[1], int(sys.argv[2]), sys.argv[3]
try:
    d=json.loads(raw)
except Exception as e:
    print(f"[{tag}] FAIL parse ms={ms} err={e} raw={raw[:300]!r}")
    sys.exit(0)
if "error" in d:
    print(f"[{tag}] FAIL api_error ms={ms} {d['error']}")
    sys.exit(0)
u=d.get("usage") or {}
ct=int(u.get("completion_tokens") or 0)
pt=int(u.get("prompt_tokens") or 0)
choice=(d.get("choices") or [{}])[0]
msg=(choice.get("message") or {})
content=(msg.get("content") or "") or ""
reason=(msg.get("reasoning_content") or "") or ""
fr=choice.get("finish_reason")
tps = (ct / (ms/1000.0)) if ms>0 and ct>0 else 0.0
print(f"[{tag}] OK ms={ms} prompt={pt} completion={ct} tps={tps:.2f} finish={fr} content_len={len(content)} reason_len={len(reason)}")
print(f"[{tag}] content_head={content[:120]!r}")
PY
}

start_server() {
  local tag="$1"
  shift
  local extra=("$@")
  local log="${LOG_DIR}/${tag}.log"
  echo "===== START ${tag} extra=${extra[*]:-(none)} ====="
  kill_port
  # shellcheck disable=SC2086
  nohup "$PY" -m sglang.launch_server \
    --model-path "${MODEL_PATH}" \
    --served-model-name "${NAME}" \
    --host 127.0.0.1 \
    --port "${PORT}" \
    --tp-size "${SGLANG_TP_SIZE:-8}" \
    --mem-fraction-static "${SGLANG_MEM_FRACTION:-0.80}" \
    --context-length "${SGLANG_CONTEXT_LENGTH:-131072}" \
    --trust-remote-code \
    --reasoning-parser qwen3 \
    --tool-call-parser qwen3_coder \
    --allow-auto-truncate \
    ${extra[@]+"${extra[@]}"} \
    >"$log" 2>&1 &
  echo $! >"${LOG_DIR}/${tag}.pid"
  if ! wait_ready "$tag" 150; then
    echo "[${tag}] FAIL ready — tail log:"
    tail -60 "$log" || true
    rg -n 'Error|Exception|Traceback|CUDA|Illegal|insufficient|failed' "$log" | tail -40 || true
    return 1
  fi
  # warmup request (short)
  bench_once "${tag}_warmup" 32 || true
  # measured
  bench_once "${tag}_gen256" 256
  bench_once "${tag}_gen512" 512
  return 0
}

echo "[host] $(hostname) $(date)"
nvidia-smi -L | head -8
"$PY" -c 'import torch,sglang; print("torch",torch.__version__,"cuda",torch.version.cuda,"sglang",sglang.__version__)'

# A: current production flags
start_server A_triton_nocudagraph \
  --disable-cuda-graph --attention-backend triton --moe-runner-backend triton --skip-server-warmup \
  || true

# B: triton but allow cuda graph
start_server B_triton_cudagraph \
  --attention-backend triton --moe-runner-backend triton --skip-server-warmup \
  || true

# C: default backends, no cuda-graph disable (sglang auto)
start_server C_default_backends \
  --skip-server-warmup \
  || true

# D: default backends + explicit disable cuda graph (isolate backend vs graph)
start_server D_default_nocudagraph \
  --disable-cuda-graph --skip-server-warmup \
  || true

echo "===== DONE $(date) ====="
kill_port
echo "[note] server stopped after A/B; resubmit sbatch to restore serving."
