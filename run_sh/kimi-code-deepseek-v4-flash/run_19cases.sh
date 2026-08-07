#!/bin/bash
# Launch the 19-case benchmark batch for deepseek-v4-flash (high concurrency).
#
# Run from toolathlon_gym/:
#   nohup bash run_sh/kimi-code-deepseek-v4-flash/run_19cases.sh \
#     > run_sh/kimi-code-deepseek-v4-flash/run_19cases.log 2>&1 &
#
# Monitor:
#   tail -f run_sh/kimi-code-deepseek-v4-flash/run_19cases.log
#   tail -f dumps/kimi-code_deepseek-v4-flash/summary_parallel_*.csv

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
cd "$PROJECT_ROOT"

mapfile -t TASKS < "$SCRIPT_DIR/tasks_19.txt"
FILTERED=()
for t in "${TASKS[@]}"; do
  t="${t//$'\r'/}"
  t="${t#"${t%%[![:space:]]*}"}"
  t="${t%"${t##*[![:space:]]}"}"
  [[ -n "$t" && "$t" != \#* ]] && FILTERED+=("$t")
done

echo "[run_19cases] model=deepseek-v4-flash tasks=${#FILTERED[@]} MAX_CONCURRENT=${MAX_CONCURRENT:-6}"
exec bash "$SCRIPT_DIR/run_eval_parallel.sh" "${FILTERED[@]}"
