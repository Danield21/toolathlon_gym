#!/bin/bash
# Slurm job spec: build the toolathlon-pack image ON A COMPUTE NODE.
#
# Rationale (2026-08-15 case-study): login nodes carry a 12GB cgroup cap that
# repeatedly OOM-killed tsc / unsquashfs / mksquashfs mid-build, shipping broken
# images (google-forms-mcp missing build/index.js -> c5 batch infra_failed).
# Compute nodes have ~1TB RAM but no direct egress; the build script routes all
# downloads through the login node's mihomo proxy at 192.168.180.240:7893
# (verified CONNECT 200 to registry.npmjs.org / github.com from gnho019).
#
# Usage: sbatch scripts/enroot_build_agent.sbatch
# Log:   toolathlon_gym_eval_dockers/logs/build_agent_compute.log
#SBATCH --job-name=pack-build
#SBATCH --partition=linlab
#SBATCH --nodes=1
#SBATCH --ntasks=1
#SBATCH --cpus-per-task=32
#SBATCH --mem=200G
#SBATCH --time=04:00:00
#SBATCH --output=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers/logs/build_agent_compute.log
#SBATCH --error=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers/logs/build_agent_compute.log

set -euo pipefail

# Fail fast if the proxy is unreachable from this node — a silent proxy outage
# would drop the build into "direct domestic mirrors" mode and die at apt.
timeout 10 curl -sI --proxy http://192.168.180.240:7893 https://registry.npmjs.org >/dev/null \
  || { echo "[FATAL] cluster proxy unreachable from $(hostname) — build aborted"; exit 1; }
echo "[build] proxy OK from $(hostname), SLURM_JOB_ID=${SLURM_JOB_ID:-none}"

export CLUSTER_PROXY=http://192.168.180.240:7893
export http_proxy="$CLUSTER_PROXY" https_proxy="$CLUSTER_PROXY"
export HTTP_PROXY="$CLUSTER_PROXY" HTTPS_PROXY="$CLUSTER_PROXY"
export no_proxy="127.0.0.1,localhost" NO_PROXY="127.0.0.1,localhost"

bash /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/scripts/enroot_build_agent.sh
