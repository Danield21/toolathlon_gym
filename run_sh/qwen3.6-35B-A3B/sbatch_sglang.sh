#!/usr/bin/env bash
#SBATCH -J qwen36-sglang
#SBATCH -A lintao
#SBATCH -p linlab
#SBATCH -w gnho019
#SBATCH --nodes=1
#SBATCH --gres=gpu:8
#SBATCH -c 64
#SBATCH --mem=800G
#SBATCH --time=7-00:00:00
#SBATCH -o /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/qwen3.6-35B-A3B/logs/sbatch_%j.out
#SBATCH -e /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/qwen3.6-35B-A3B/logs/sbatch_%j.err

set -euo pipefail
SCRIPT_DIR=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/qwen3.6-35B-A3B
mkdir -p "${SCRIPT_DIR}/logs"
cd "$SCRIPT_DIR"

echo "[sbatch] host=$(hostname) job=${SLURM_JOB_ID:-na} $(date)"
nvidia-smi -L || true

bash "${SCRIPT_DIR}/start_sglang.sh"

# Keep allocation alive for the SGLang process
# shellcheck disable=SC1091
source "${SCRIPT_DIR}/config.env"
PID_FILE="${SCRIPT_DIR}/logs/sglang_${SGLANG_PORT}.pid"
if [[ ! -f "$PID_FILE" ]]; then
  echo "[sbatch] missing pid file" >&2
  exit 1
fi
PID=$(cat "$PID_FILE")
echo "[sbatch] watching sglang pid=$PID"
while kill -0 "$PID" 2>/dev/null; do
  sleep 60
done
echo "[sbatch] sglang exited $(date)"
exit 1
