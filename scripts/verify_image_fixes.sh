#!/bin/bash
# Verify the rebuilt Enroot agent image contains all audit-fix dependencies.
# Run after scripts/rebuild_image_after_fixes.sh.
#
# Usage:
#   bash scripts/verify_image_fixes.sh

set -euo pipefail

RUNTIME_ROOT="${TOOLATHLON_EVAL_DOCKER_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers}"
INST="verify-image-$$"
FAIL=0

cleanup() {
  enroot remove -f "$INST" >/dev/null 2>&1 || true
}
trap cleanup EXIT

echo "[verify] starting ephemeral container from toolathlon-pack.sqsh ..."
enroot start -r "$TOOLATHLON_AGENT_NAME" "$INST" /bin/bash -c '
  set -e
  echo "=== PDF tools ==="
  command -v pdftotext && command -v pdfinfo && command -v qpdf || { echo "MISSING pdf tools"; exit 1; }
  echo "=== Sandbox ==="
  command -v bwrap || { echo "MISSING bwrap"; exit 1; }
  ldconfig -p | grep -q libseccomp || { echo "MISSING libseccomp"; exit 1; }
  echo "=== Fonts ==="
  fc-list 2>/dev/null | grep -qi "noto.*cjk\|liberation" || { echo "MISSING fonts"; exit 1; }
  echo "=== Playwright Chromium ==="
  ls /root/.cache/ms-playwright/chromium-*/chrome-linux/chrome >/dev/null 2>&1 \
    || ls /root/.cache/ms-playwright/chromium-*/chrome-linux64/chrome >/dev/null 2>&1 \
    || { echo "MISSING chromium"; exit 1; }
  echo "=== MCP uv venvs (runtime-required) ==="
  [[ -x /opt/local_servers/cli-mcp-server/.venv/bin/python ]] || { echo "MISSING cli-mcp-server .venv"; exit 1; }
  echo "cli-mcp-server .venv OK"
  if [[ -x /opt/local_servers/mcp-youtube-transcript/.venv/bin/python ]]; then
    echo "mcp-youtube-transcript .venv OK"
  else
    echo "WARN: mcp-youtube-transcript .venv missing (optional for C.1, required for full benchmark)"
  fi
  echo "=== Calendar MCP timezone fix ==="
  grep -q "toUtcIso" /opt/local_servers/Calendar-Autoauth-MCP-Server/build/index.js || { echo "MISSING Calendar tz fix (toUtcIso)"; exit 1; }
  echo "=== Kimi dist (runtime-staged, just confirm node + path) ==="
  node --version
  echo "ALL CHECKS PASSED"
' && echo "[verify] PASS" || { echo "[verify] FAIL"; FAIL=1; }

exit "$FAIL"
