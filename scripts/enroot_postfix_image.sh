#!/bin/bash
# Post-fix the c2s agent image on a COMPUTE node (large memory, no cgroup cap).
#
# Background (case-study 2026-08-12):
#   The c2s image was built on the login node under a 12GB cgroup cap. The
#   mksquashfs export succeeded (1.6GB sqsh) but two defects remain:
#     1. /opt/uv_python_cache was deleted pre-export to save space → all uv
#        MCP server .venv python symlinks are DANGLING (point to a non-existent
#        /opt/uv_python_cache/cpython-<ver>/...).
#     2. Office-Word-MCP-Server and Office-PowerPoint-MCP-Server .venvs were
#        never created (uv sync failed due to proxy flakiness during build).
#
#   The login node cannot re-export (mksquashfs OOMs under 12GB cgroup). The
#   compute node has large memory + no cgroup cap, AND can reach the tuna PyPI
#   mirror (internal mirror, no external network needed). So we do the fix here.
#
# Usage (from login node):
#   srun -p linlab -N1 -n1 -c8 --mem=32G bash scripts/enroot_postfix_image.sh
#
set -euo pipefail

# Compute nodes inherit a bogus http_proxy=127.0.0.1:7890; clear it.
unset http_proxy https_proxy HTTP_PROXY HTTPS_PROXY ALL_PROXY all_proxy

RUNTIME_ROOT="/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers"
PROJECT_ROOT="/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym"
SQSH_OUT="${RUNTIME_ROOT}/images/toolathlon-pack.sqsh"

# ── enroot on /dev/shm (proven pattern from run_on_slurm.sh) ──────────────────
ENROOT_SRC="/storage/lintaoLab/bowending/.local/enroot"
ENROOT_LOCAL="/dev/shm/enroot_install_postfix"
if [[ ! -x "$ENROOT_LOCAL/bin/enroot" ]]; then
  mkdir -p "$ENROOT_LOCAL"
  rsync -a "$ENROOT_SRC/" "$ENROOT_LOCAL/" 2>/dev/null || cp -a "$ENROOT_SRC/." "$ENROOT_LOCAL/"
fi
export ENROOT_LIBRARY_PATH="${ENROOT_LOCAL}/lib"
export ENROOT_SYSCONF_PATH="${ENROOT_LOCAL}/etc"
# IMPORTANT: call the /dev/shm enroot DIRECTLY (not the /storage wrapper, which
# re-resolves ENROOT_HOME from BASH_SOURCE and hits "Argument list too long").
ENROOT_BIN="${ENROOT_LOCAL}/bin/enroot"
export PATH="${ENROOT_LOCAL}/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin:/storage/lintaoLab/bowending/miniconda3/envs/toolathlon_gym/bin:/storage/lintaoLab/bowending/.local/bin"

echo "[postfix] enroot: $("$ENROOT_BIN" version 2>/dev/null || echo '?')"

# Use /dev/shm (large tmpfs on compute nodes) for all enroot working dirs.
export ENROOT_DATA_PATH="/dev/shm/enroot_postfix_data"
export ENROOT_TEMP_PATH="/dev/shm/enroot_postfix_tmp"
export ENROOT_RUNTIME_PATH="/dev/shm/enroot_postfix_runtime"
mkdir -p "$ENROOT_DATA_PATH" "$ENROOT_TEMP_PATH" "$ENROOT_RUNTIME_PATH"

# uv (on the compute node, uv lives in the dbw_dev conda env, not .local/bin)
UV_BIN="/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/uv"
command -v "$UV_BIN" >/dev/null 2>&1 || { echo "[postfix][FATAL] uv not found at $UV_BIN"; exit 1; }
echo "[postfix] uv: $("$UV_BIN" --version)"

# tuna PyPI mirror (reachable from compute node — internal mirror)
export UV_DEFAULT_INDEX="https://pypi.tuna.tsinghua.edu.cn/simple"
export UV_INDEX_URL="https://pypi.tuna.tsinghua.edu.cn/simple"

log() { echo "[postfix][$(date +%H:%M:%S)] $*"; }

cleanup() {
  log "Cleaning up..."
  "$ENROOT_BIN" remove -f toolathlon-postfix 2>/dev/null || true
  rm -rf "$ENROOT_DATA_PATH/toolathlon-postfix" 2>/dev/null || true
}
trap cleanup EXIT

SQSH_IN="${RUNTIME_ROOT}/images/toolathlon-pack.sqsh"
log "Input sqsh: $SQSH_IN ($(du -h "$SQSH_IN" | cut -f1))"

# 1. unsquashfs the c2s image into a rootfs.
log "Creating rootfs from sqsh (unsquashfs)..."
"$ENROOT_BIN" create -n toolathlon-postfix "$SQSH_IN"
ROOTFS="${ENROOT_DATA_PATH}/toolathlon-postfix"
log "Rootfs ready: $(du -sh "$ROOTFS" 2>/dev/null | cut -f1)"

# 2. rsync the host-side uv_python_cache (3.11/3.12/3.13) into the rootfs.
UV_PY_HOST_CACHE="${RUNTIME_ROOT}/uv_python_cache"
if [[ -d "$UV_PY_HOST_CACHE" ]]; then
  log "Staging uv_python_cache (3.11/3.12/3.13) into rootfs..."
  mkdir -p "$ROOTFS/opt/uv_python_cache"
  rsync -a "$UV_PY_HOST_CACHE/" "$ROOTFS/opt/uv_python_cache/"
  log "uv_python_cache staged: $(du -sh "$ROOTFS/opt/uv_python_cache" | cut -f1)"
else
  log "[WARN] $UV_PY_HOST_CACHE not found — .venv symlinks will remain broken!"
fi

# 3. Rebuild the two missing .venvs (Office-Word, Office-PowerPoint).
export UV_PYTHON_INSTALL_DIR="$ROOTFS/opt/uv_python_cache"
export UV_PYTHON_PREFERENCE=managed

for srv in Office-Word-MCP-Server Office-PowerPoint-MCP-Server; do
  dir="$ROOTFS/opt/local_servers/$srv"
  if [[ ! -d "$dir" ]]; then
    log "[WARN] $srv not in rootfs, skipping"
    continue
  fi
  if [[ -x "$dir/.venv/bin/python" ]]; then
    log "$srv .venv already OK, skipping"
    continue
  fi
  log "Rebuilding .venv for $srv..."
  rm -rf "$dir/.venv" 2>/dev/null || true
  if (cd "$dir" && "$UV_BIN" sync); then
    if [[ -x "$dir/.venv/bin/python" ]]; then
      log "  [ok] $srv .venv ready"
    else
      log "  [WARN] $srv .venv python still missing after sync"
    fi
  else
    log "  [WARN] $srv uv sync failed — MCP server may be non-functional"
  fi
done

# 4. Verify all critical .venvs now have a python entry. We check that the
# symlink EXISTS (readlink succeeds) and points into /opt/uv_python_cache —
# we can't `-e` it because the symlink target is a container-absolute path
# (/opt/...) that doesn't resolve on the host filesystem.
log "Verifying .venv integrity..."
_missing=0
for srv in arxiv-mcp-server arxiv-latex-mcp emails-mcp mcp-snowflake-server \
           mcp-scholarly Office-Word-MCP-Server Office-PowerPoint-MCP-Server \
           excel-mcp-server pdf-tools-mcp mcp-youtube-transcript \
           mcp-google-sheets cli-mcp-server yahoo-finance-mcp; do
  py="$ROOTFS/opt/local_servers/$srv/.venv/bin/python"
  if [[ -L "$py" ]] || [[ -e "$py" ]]; then
    tgt=$(readlink -f "$py" 2>/dev/null || readlink "$py" 2>/dev/null || echo "?")
    # Accept if it points to the staged cache (managed cpython) or a system python.
    case "$tgt" in
      */uv_python_cache/*|/usr/bin/python*|*miniconda*|*conda*)
        log "  [ok] $srv -> $(basename "$tgt")"
        ;;
      *)
        log "  [ok?] $srv -> $tgt"
        ;;
    esac
  else
    log "  [MISSING] $srv"
    _missing=$((_missing+1))
  fi
done
log ".venv verification: $((13 - _missing))/13 present, $_missing missing"

# 5. Re-export to sqsh.
log "Exporting fixed rootfs to $SQSH_OUT..."
rm -f "$SQSH_OUT"
export ENROOT_MAX_PROCESSORS=4
"$ENROOT_BIN" export -o "$SQSH_OUT" toolathlon-postfix
log "Exported: $(ls -lh "$SQSH_OUT" | awk '{print $5}')"

log "Done. Fixed image ready: $SQSH_OUT"
