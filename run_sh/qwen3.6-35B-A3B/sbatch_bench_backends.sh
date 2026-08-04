#!/usr/bin/env bash
#SBATCH -J qwen36-bench
#SBATCH -A lintao
#SBATCH -p linlab
#SBATCH -w gnho019
#SBATCH --nodes=1
#SBATCH --gres=gpu:8
#SBATCH -c 64
#SBATCH --mem=800G
#SBATCH --time=0-04:00:00
#SBATCH -o /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/qwen3.6-35B-A3B/logs/bench_%j.out
#SBATCH -e /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/qwen3.6-35B-A3B/logs/bench_%j.err

set -euo pipefail
SCRIPT_DIR=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/qwen3.6-35B-A3B
cd "$SCRIPT_DIR"
bash "${SCRIPT_DIR}/bench_backends.sh"
