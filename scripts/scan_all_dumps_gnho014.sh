#!/bin/bash
# Full scan of all dump trees for mock-port overflow / MCP Invalid URL.
# Runs on a compute node (gnho014) where CPU is plentiful; paths identical.
set -euo pipefail
cd /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym
OUT=/lintaoLab2/bowending/project_agent_swarm_benchmark/dev_docs/mcp_invalid_url_scan_20260825
mkdir -p "$OUT"

ROOTS=(
  dumps/glm-5.3
  dumps/deepseek-v4-flash-0731
  dumps/kimi-k3
  dumps/kimi-code_Qwen3.8-Max-sa167
  dumps/minimax-m3
)
for r in "${ROOTS[@]}"; do
  tag=$(echo "$r" | sed 's|dumps/||; s|/|_|g')
  echo "=== scanning $r ==="
  python3 scripts/scan_mcp_invalid_url.py "$r" --out "$OUT/${tag}.tsv" > "$OUT/${tag}.summary.txt" 2>&1 || echo "  rc=$?"
  head -1 "$OUT/${tag}.summary.txt"
done
echo "ALL DONE -> $OUT"
