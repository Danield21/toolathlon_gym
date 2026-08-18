#!/bin/bash
# Final post-fix: repair the 2 bad Word/PPT .venv symlinks in the c2s/postfix-v3
# image, then re-export. This is the ONLY remaining defect (full audit passed
# all other checks — see case-study 2026-08-12).
#
# CRITICAL: this script verifies everything with `enroot start` (real container
# execution, not chroot which needs root) BEFORE exporting, so we never ship a
# broken image. If verification fails, the script exits WITHOUT overwriting the
# sqsh.
#
# Usage (from login node):
#   srun -p linlab -N1 -n1 -c8 --mem=32G bash scripts/enroot_postfix_final.sh
#
set -euo pipefail

# Compute nodes inherit a bogus http_proxy=127.0.0.1:7890; clear it.
unset http_proxy https_proxy HTTP_PROXY HTTPS_PROXY ALL_PROXY all_proxy

RUNTIME_ROOT="/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers"
SQSH_IN="${RUNTIME_ROOT}/images/toolathlon-pack.sqsh"
SQSH_OUT="${RUNTIME_ROOT}/images/toolathlon-pack.sqsh"
SQSH_TMP="${RUNTIME_ROOT}/images/toolathlon-pack.sqsh.final"

# ── enroot on /dev/shm (proven pattern from run_on_slurm.sh) ──────────────────
ENROOT_SRC="/storage/lintaoLab/bowending/.local/enroot"
ENROOT_LOCAL="/dev/shm/enroot_install_final"
mkdir -p "$ENROOT_LOCAL"
rsync -a "$ENROOT_SRC/" "$ENROOT_LOCAL/" 2>/dev/null || cp -a "$ENROOT_SRC/." "$ENROOT_LOCAL/"
export ENROOT_LIBRARY_PATH="${ENROOT_LOCAL}/lib"
export ENROOT_SYSCONF_PATH="${ENROOT_LOCAL}/etc"
ENROOT_BIN="${ENROOT_LOCAL}/bin/enroot"
export PATH="${ENROOT_LOCAL}/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin"

export ENROOT_DATA_PATH="/dev/shm/enroot_final_data"
export ENROOT_TEMP_PATH="/dev/shm/enroot_final_tmp"
export ENROOT_RUNTIME_PATH="/dev/shm/enroot_final_runtime"
mkdir -p "$ENROOT_DATA_PATH" "$ENROOT_TEMP_PATH" "$ENROOT_RUNTIME_PATH"

log() { echo "[final][$(date +%H:%M:%S)] $*"; }

cleanup() {
  log "Cleaning up rootfs..."
  "$ENROOT_BIN" remove -f toolathlon-final 2>/dev/null || true
}
trap cleanup EXIT

log "=== STEP 1: unsquashfs the current image ==="
log "Input: $SQSH_IN ($(du -h "$SQSH_IN" | cut -f1))"
"$ENROOT_BIN" create -n toolathlon-final "$SQSH_IN"
ROOTFS="${ENROOT_DATA_PATH}/toolathlon-final"
log "Rootfs ready: $(du -sh "$ROOTFS" | cut -f1)"

log "=== STEP 2: repair Word/PPT .venv symlinks ==="
# The postfix-v3 build created these .venvs with UV_PYTHON_INSTALL_DIR pointing
# at the host rootfs path (/dev/shm/enroot_postfix_data/...), baking that
# absolute path into the symlinks. Rewrite them to the container path.
CORRECT_TGT="/opt/uv_python_cache/cpython-3.13.13-linux-x86_64-gnu/bin/python3.13"
_fixed=0
for srv in Office-Word-MCP-Server Office-PowerPoint-MCP-Server; do
  py="$ROOTFS/opt/local_servers/$srv/.venv/bin/python"
  if [[ -L "$py" ]]; then
    cur=$(readlink "$py")
    if [[ "$cur" != "$CORRECT_TGT" ]]; then
      ln -sf "$CORRECT_TGT" "$py"
      log "  fixed $srv: $cur -> $CORRECT_TGT"
      _fixed=$((_fixed+1))
    else
      log "  $srv already correct"
    fi
  elif [[ -e "$py" ]]; then
    log "  $srv python is a real file (not symlink), leaving as-is"
  else
    log "  [WARN] $srv python MISSING entirely!"
  fi
done
log "Fixed $_fixed symlinks"

log "=== STEP 3: verify ALL .venv symlinks point to /opt/uv_python_cache ==="
_bad=0
for srv in arxiv-mcp-server arxiv-latex-mcp emails-mcp mcp-snowflake-server \
           mcp-scholarly Office-Word-MCP-Server Office-PowerPoint-MCP-Server \
           excel-mcp-server pdf-tools-mcp mcp-youtube-transcript \
           mcp-google-sheets cli-mcp-server yahoo-finance-mcp; do
  py="$ROOTFS/opt/local_servers/$srv/.venv/bin/python"
  if [[ -L "$py" ]]; then
    tgt=$(readlink "$py")
    case "$tgt" in
      /opt/uv_python_cache/*|/usr/bin/python*) log "  [ok] $srv";;
      *) log "  [BAD] $srv -> $tgt"; _bad=$((_bad+1));;
    esac
  elif [[ -e "$py" ]]; then
    log "  [ok] $srv (real file)"
  else
    log "  [MISSING] $srv"; _bad=$((_bad+1))
  fi
done
if [[ $_bad -gt 0 ]]; then
  log "[FATAL] $_bad .venv symlinks still bad — aborting (NOT overwriting sqsh)"
  exit 1
fi
log "All 13 .venv symlinks verified"

log "=== STEP 4: functional test via enroot start (real container exec) ==="
# This is the definitive test: actually RUN the python inside the container.
"$ENROOT_BIN" start --root --rw toolathlon-final bash -c '
set -e
echo "[test] cli-mcp-server python:"
/opt/local_servers/cli-mcp-server/.venv/bin/python -c "import sys; print(sys.version)"
echo "[test] Office-Word python:"
/opt/local_servers/Office-Word-MCP-Server/.venv/bin/python -c "import sys; print(sys.version)"
echo "[test] Office-PowerPoint python:"
/opt/local_servers/Office-PowerPoint-MCP-Server/.venv/bin/python -c "import sys; print(sys.version)"
echo "[test] arxiv-mcp-server python:"
/opt/local_servers/arxiv-mcp-server/.venv/bin/python -c "import sys; print(sys.version)"
echo "[test] chromium binary exists:"
ls -la /root/.cache/ms-playwright/chromium-1234/chrome-linux64/chrome | head -1
echo "[test] Calendar MCP toUtcIso present:"
grep -c toUtcIso /opt/local_servers/Calendar-Autoauth-MCP-Server/build/index.js
echo "[test] ALL FUNCTIONAL TESTS PASSED"
' || {
  log "[FATAL] functional test FAILED — aborting (NOT overwriting sqsh)"
  exit 1
}
log "Functional tests passed"

log "=== STEP 5: export fixed image ==="
log "Exporting to $SQSH_TMP first (don't clobber the input until success)..."
rm -f "$SQSH_TMP"
export ENROOT_MAX_PROCESSORS=4
"$ENROOT_BIN" export -o "$SQSH_TMP" toolathlon-final
log "Exported: $(ls -lh "$SQSH_TMP" | awk '{print $5}')"

log "=== STEP 6: validate the new sqsh is a valid squashfs ==="
unsquashfs -s "$SQSH_TMP" >/dev/null 2>&1 || {
  log "[FATAL] new sqsh is invalid — keeping old image"
  rm -f "$SQSH_TMP"
  exit 1
}
log "New sqsh validated"

log "=== STEP 7: swap in the new image ==="
mv "$SQSH_TMP" "$SQSH_OUT"
log "Done. Final image: $SQSH_OUT ($(ls -lh "$SQSH_OUT" | awk '{print $5}'))"
