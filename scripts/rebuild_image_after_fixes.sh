#!/bin/bash
# Rebuild the Enroot agent image after the 2026-08-12 audit fixes.
#
# This applies the system-package layer that CANNOT be hot-staged at runtime:
#   - poppler-utils / qpdf           (PDF parsing fallback)
#   - bubblewrap / libseccomp2       (sandbox capability)
#   - fonts-liberation / fonts-noto-cjk (browser + PDF rendering)
#   - playwright chromium            (browser MCP tasks)
#
# Everything else (Kimi dist, notion cli.mjs, notion-openapi.json, canvas
# build, terminal.yaml, kimi_harness, tasks/) is already hot-staged per-worker
# by run_eval_parallel.sh, so this rebuild only needs to land the OS layer.
#
# PREREQUISITE: proxy must be reachable for apt/npm/pip:
#   export http_proxy=http://127.0.0.1:7891 https_proxy=http://127.0.0.1:7891
#
# Usage:
#   bash scripts/rebuild_image_after_fixes.sh
#
# After the build, verify with:
#   bash scripts/verify_image_fixes.sh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
RUNTIME_ROOT="${TOOLATHLON_EVAL_DOCKER_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers}"

# Proxy sanity check
if ! curl -fsS -m 5 -x "${http_proxy:-http://127.0.0.1:7891}" \
     https://registry.npmjs.org/ >/dev/null 2>&1; then
  echo "[error] proxy ${http_proxy:-http://127.0.0.1:7891} is not reachable." >&2
  echo "        Set http_proxy/https_proxy to a working proxy before rebuilding." >&2
  exit 1
fi

echo "[rebuild] proxy OK, launching enroot_build_agent.sh ..."
bash "${PROJECT_ROOT}/scripts/enroot_build_agent.sh"

echo "[rebuild] build complete. New image:"
ls -la "${RUNTIME_ROOT}/images/toolathlon-pack.sqsh"
echo "[rebuild] sha256:"
sha256sum "${RUNTIME_ROOT}/images/toolathlon-pack.sqsh"
