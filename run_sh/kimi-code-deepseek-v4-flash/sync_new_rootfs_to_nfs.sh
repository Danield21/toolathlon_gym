#!/bin/bash
# Sync the freshly-built shm rootfs to the NFS snapshot that slurm compute
# nodes read via AGENT_TEMPLATE.
#
# Why: enroot_build_agent.sh writes the new rootfs to /dev/shm/enroot_data/
# toolathlon-pack (tmpfs on the login node). Slurm compute nodes cannot see
# the login node's /dev/shm, so run_on_slurm.sh points AGENT_TEMPLATE at an
# NFS directory (toolathlon-pack-rootfs). This script refreshes that snapshot.
#
# Run AFTER a successful image build, BEFORE launching the slurm eval.
#
# Usage:
#   bash run_sh/kimi-code-deepseek-v4-flash/sync_new_rootfs_to_nfs.sh
set -euo pipefail

SRC="/dev/shm/enroot_data/toolathlon-pack"
DST="/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers/toolathlon-pack-rootfs"

if [[ ! -d "$SRC" ]]; then
  echo "[sync] ERROR: source rootfs $SRC does not exist." >&2
  echo "[sync]        Run enroot_build_agent.sh first." >&2
  exit 1
fi

echo "[sync] source:      $SRC ($(du -sh "$SRC" | cut -f1))"
echo "[sync] destination: $DST ($(du -sh "$DST" 2>/dev/null | cut -f1 || echo 'n/a'))"
echo "[sync] rsyncing (this may take a few minutes on NFS)..."

# Use --delete so stale files from old builds are removed; exclude enroot
# runtime metadata that is host/instance specific.
rsync -a --delete \
  --exclude '.enroot' \
  "$SRC/" "$DST/"

echo "[sync] done. Verifying key fixes are present in NFS snapshot:"

verify() {
  local label="$1" path="$2" pattern="$3"
  if [[ -e "$DST/$path" ]] && grep -qE "$pattern" "$DST/$path" 2>/dev/null; then
    echo "  [OK]   $label"
  else
    echo "  [MISS] $label  ($DST/$path)"
  fi
}

verify "canvas UTC task.md" \
  "workspace/tasks/finalpool/canvas-announcement-summary/docs/task.md" "UTC"
verify "sf-customer fixed date task.md" \
  "workspace/tasks/finalpool/sf-customer-health-dashboard/docs/task.md" "2026-03-08"
verify "ppt-snowflake cancelled clause task.md" \
  "workspace/tasks/finalpool/ppt-snowflake-executive/docs/task.md" "[Cc]ancelled"
verify "woo orderby alias fix" \
  "opt/local_servers/woocommerce-mcp/dist/services/pg-rest-router.js" "date_created"
verify "gform checkbox tool" \
  "opt/local_servers/google-forms-mcp/dist/index.js" "add_checkbox_question|addCheckboxQuestion"
verify "snowflake fold fix" \
  "opt/local_servers/mcp-snowflake-server/dist/mcp_snowflake_server/db_client.py" "_snowflake_fold" \
  2>/dev/null || verify "snowflake fold fix (src)" \
  "opt/local_servers/mcp-snowflake-server/src/mcp_snowflake_server/db_client.py" "_snowflake_fold"

echo "[sync] NFS snapshot refreshed. Compute nodes will see the new image."
