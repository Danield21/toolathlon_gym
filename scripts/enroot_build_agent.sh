#!/bin/bash
# Build toolathlon-pack.sqsh from ubuntu:22.04 via enroot (replaces: docker build -t toolathlon-pack:latest .)
#
# Usage:
#   bash scripts/enroot_build_agent.sh
#
# Requires proxy for apt/npm/pip on this cluster:
#   export http_proxy=${CLUSTER_PROXY} https_proxy=${CLUSTER_PROXY}

set -euo pipefail

# Cluster proxy: on the login node mihomo listens on 127.0.0.1:7893; on compute
# nodes (no local proxy, no direct egress) the same mihomo is reachable at the
# login node's internal IP. Building on a COMPUTE node via sbatch is the
# preferred path: login nodes carry a 12GB cgroup cap that OOM-kills tsc /
# mksquashfs mid-build (case-study 2026-08-15: google-forms-mcp tsc killed ->
# broken image -> c5 batch infra_failed).
export CLUSTER_PROXY="${CLUSTER_PROXY:-http://192.168.180.240:7893}"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
RUNTIME_ROOT="${TOOLATHLON_EVAL_DOCKER_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers}"

# shellcheck disable=SC1091
source "$RUNTIME_ROOT/env.sh"

# Build on tmpfs (/dev/shm) NOT NFS. /storage/lintaoLab is NFS and enroot's
# unsquashfs/library reads intermittently hang on NFS read(), deadlocking even
# `enroot version`. /dev/shm is 126G RAM-backed on this node — more than enough
# for a ~6GB build rootfs, and extraction is near-instant vs 8+min on NFS.
# This mirrors what run_eval_parallel.sh does for the runtime path.
#
# IMPORTANT (case-study 2026-08-12): env.sh above sets ENROOT_DATA_PATH to the
# NFS path via ${ENROOT_DATA_PATH:-NFS}. Because it runs BEFORE this block, a
# mere ${ENROOT_DATA_PATH:-/dev/shm/...} override here is a no-op (the var is
# already non-empty), and the build silently writes the rootfs to NFS. NFS
# unlink races then leave .nfs* ghost files that block `rm -rf` cleanup,
# looping the build at the pre-build cleanup step forever. Fix: force-override
# unconditionally (drop the :- default) so the build rootfs ALWAYS lives in tmpfs.
export ENROOT_DATA_PATH="/dev/shm/enroot_build_data"
export ENROOT_TEMP_PATH="/dev/shm/enroot_build_tmp"
export ENROOT_RUNTIME_PATH="/dev/shm/enroot_build_runtime"
export ENROOT_CACHE_PATH="/dev/shm/enroot_build_cache"
mkdir -p "$ENROOT_DATA_PATH" "$ENROOT_TEMP_PATH" "$ENROOT_RUNTIME_PATH" "$ENROOT_CACHE_PATH"

BUILD_NAME="toolathlon-pack-build"
PROVISION_SCRIPT="${RUNTIME_ROOT}/tmp/provision_agent.sh"
BUILD_LOG="${RUNTIME_ROOT}/logs/build_agent.log"

die() { echo "[error] $*" >&2; exit 1; }
log() { echo "[$(date +%H:%M:%S)] $*"; }

# Mirror the enroot install into /dev/shm to avoid NFS hangs on the wrapper
# script + libraries (same technique run_on_slurm.sh uses on compute nodes).
_ENROOT_SRC="/storage/lintaoLab/bowending/.local/enroot"
_ENROOT_LOCAL="/dev/shm/enroot_install_build"
if [[ ! -x "$_ENROOT_LOCAL/bin/enroot" ]]; then
  mkdir -p "$_ENROOT_LOCAL"
  rsync -a "$_ENROOT_SRC/" "$_ENROOT_LOCAL/" 2>/dev/null || cp -a "$_ENROOT_SRC/." "$_ENROOT_LOCAL/"
fi
export ENROOT_LIBRARY_PATH="${_ENROOT_LOCAL}/lib"
export ENROOT_SYSCONF_PATH="${_ENROOT_LOCAL}/etc"
export PATH="${_ENROOT_LOCAL}/bin:${PATH}"
log "enroot on tmpfs: $_ENROOT_LOCAL"

# Proxy is OPTIONAL on this cluster: domestic mirrors (tuna/aliyun) and
# npm/pypi registries are directly reachable without a proxy. If the proxy
# env var is unset OR points at a dead port, switch to "direct domestic"
# mode: clear the proxy vars and let the provision script re-point apt/npm/pip
# at Tsinghua mirrors. Set USE_PROXY=1 to force the proxy path.
_probe_proxy() {
  local p="${http_proxy:-${https_proxy:-}}"
  [[ -z "$p" ]] && return 1
  # Strip scheme, then split host:port (works for 127.0.0.1:7893 and
  # 192.168.180.240:7893 alike — the login-node hardcode made compute-node
  # builds silently fall into no-proxy mode).
  local hp="${p#*://}"
  local host="${hp%%:*}"
  local port="${hp##*:}"; port="${port%%/*}"
  [[ -z "$host" ]] && host=127.0.0.1
  timeout 3 bash -c "echo > /dev/tcp/${host}/${port}" 2>/dev/null
}
if [[ "${USE_PROXY:-0}" != "1" ]] && ! _probe_proxy; then
  log "No working proxy detected — using direct domestic-mirror mode (tuna/aliyun)."
  unset http_proxy https_proxy HTTP_PROXY HTTPS_PROXY ALL_PROXY all_proxy
  export USE_DIRECT_MIRRORS=1
else
  export http_proxy="${http_proxy:-${CLUSTER_PROXY}}"
  export https_proxy="${https_proxy:-${CLUSTER_PROXY}}"
  export HTTP_PROXY="$http_proxy" HTTPS_PROXY="$https_proxy"
  export no_proxy="${no_proxy:-127.0.0.1,localhost}"
  export NO_PROXY="$no_proxy"
fi

[[ -f "$TOOLATHLON_UBUNTU_SQSH" ]] || die "Missing $TOOLATHLON_UBUNTU_SQSH — import ubuntu:22.04 first"
[[ -d "$PROJECT_ROOT/local_servers" ]] || die "Missing $PROJECT_ROOT/local_servers"

mkdir -p "$RUNTIME_ROOT/tmp" "$RUNTIME_ROOT/logs" "$RUNTIME_ROOT/images"

# ---------------------------------------------------------------------------
# Provision script executed INSIDE the enroot container
# ---------------------------------------------------------------------------
cat >"$PROVISION_SCRIPT" <<'PROVISION'
#!/bin/bash
set -euo pipefail
export DEBIAN_FRONTEND=noninteractive

# The build host exports CLUSTER_PROXY (login-node mihomo address) and passes
# it via `enroot start -e`; default here as a safety net so the `env ...=` call
# sites below never hit set -u on an unbound variable (case-study 2026-08-15:
# the first compute-node build died at line 80 exactly this way).
export CLUSTER_PROXY="${CLUSTER_PROXY:-http://192.168.180.240:7893}"

# Honor proxy for apt
if [[ -n "${http_proxy:-}" ]]; then
  cat >/etc/apt/apt.conf.d/95proxy <<EOF
Acquire::http::Proxy "${http_proxy}";
Acquire::https::Proxy "${https_proxy:-$http_proxy}";
EOF
fi

# Direct domestic-mirror mode: the cluster can reach Tsinghua/Aliyun mirrors and
# npm/pypi directly, but the local mihomo proxy (7891/7892) is often dead.
# Re-point apt to Tsinghua, and set npm/pip/uv to domestic registries.
#
# IMPORTANT: the base ubuntu rootfs has NO ca-certificates, so apt over HTTPS
# fails ("No system certificates available"). Use HTTP for apt until
# ca-certificates is installed, then we can use HTTPS elsewhere (npm/pip carry
# their own CA bundles).
if [[ "${USE_DIRECT_MIRRORS:-0}" == "1" ]]; then
  echo "=== switching apt to Tsinghua mirror (direct, no proxy, HTTP for bootstrap) ==="
  if grep -q "Ubuntu 22.04" /etc/os-release 2>/dev/null; then
    sed -i 's|http://archive.ubuntu.com/ubuntu|http://mirrors.tuna.tsinghua.edu.cn/ubuntu|g; s|http://[a-z]*.archive.ubuntu.com/ubuntu|http://mirrors.tuna.tsinghua.edu.cn/ubuntu|g; s|http://security.ubuntu.com/ubuntu|http://mirrors.tuna.tsinghua.edu.cn/ubuntu|g; s|https://mirrors.tuna.tsinghua.edu.cn/ubuntu|http://mirrors.tuna.tsinghua.edu.cn/ubuntu|g' /etc/apt/sources.list 2>/dev/null || true
  fi
  # npm / pip / uv domestic registries (npm/pip ship their own CAs)
  mkdir -p /root/.npm /root/.config/pip /root/.cache/uv
  cat >/root/.npmrc <<'NPMEOF'
registry=https://registry.npmmirror.com/
strict-ssl=false
NPMEOF
  cat >/root/.config/pip/pip.conf <<'PIPEOF'
[global]
index-url=https://pypi.tuna.tsinghua.edu.cn/simple
trusted-host=pypi.tuna.tsinghua.edu.cn
PIPEOF
  export UV_DEFAULT_INDEX="https://pypi.tuna.tsinghua.edu.cn/simple"
  export UV_INDEX_URL="https://pypi.tuna.tsinghua.edu.cn/simple"
  # Make sure no stale proxy is inherited
  unset http_proxy https_proxy HTTP_PROXY HTTPS_PROXY || true
  # apt: allow HTTP (no cert needed for bootstrap) and disable certificate checks
  echo 'Acquire::https::Verify-Peer "false";' > /etc/apt/apt.conf.d/99no-verify
fi

echo "=== [1/6] apt base packages ==="
apt-get update
apt-get install -y \
  curl wget git ca-certificates gnupg \
  python3 python3-pip rsync postgresql-client \
  libnss3 libnspr4 libatk1.0-0 libatk-bridge2.0-0 \
  libcups2 libdrm2 libdbus-1-3 libatspi2.0-0 \
  libx11-6 libxcomposite1 libxdamage1 libxext6 \
  libxfixes3 libxrandr2 libgbm1 libxcb1 \
  libxkbcommon0 libpango-1.0-0 libcairo2 libasound2 \
  squashfs-tools \
  file \
  poppler-utils qpdf \
  bubblewrap libseccomp2 libseccomp-dev \
  fonts-liberation fonts-noto-cjk

echo "=== [2/6] uv ==="
# astral.sh (the install.sh landing page) IS reachable directly from this
# cluster, but the installer script then downloads the uv binary from
# GitHub releases, which is NOT reachable directly. A `curl --proxy` on the
# install.sh fetch does NOT propagate to the installer's internal download, so
# the previous "direct then proxy" logic hung forever at "downloading uv ..."
# (the direct curl succeeded, `| sh` then blocked on GitHub, and the pipe's
# --max-time only bounded curl, not sh).
#
# Fix (case-study 2026-08-12): fetch install.sh (direct, or via proxy on retry),
# then RUN it with the cluster proxy exported as http_proxy/https_proxy so the
# installer's internal GitHub-releases download also goes through 7893. Wrap the
# whole pipe in `timeout` so a residual hang fails fast and retries instead of
# blocking the build indefinitely.
UV_INSTALL_URL="https://astral.sh/uv/install.sh"
_uv_done=0
for _i in 1 2 3; do
  # Attempt 1: direct fetch of install.sh; attempts 2-3: fetch via proxy.
  _proxy_flag=""
  [[ $_i -ge 2 ]] && _proxy_flag="--proxy ${CLUSTER_PROXY}"
  if curl -LsSf --connect-timeout 10 --max-time 60 $_proxy_flag "$UV_INSTALL_URL" \
    | env http_proxy=${CLUSTER_PROXY} https_proxy=${CLUSTER_PROXY} \
      HTTP_PROXY=${CLUSTER_PROXY} HTTPS_PROXY=${CLUSTER_PROXY} \
      timeout 240 sh; then
    _uv_done=1; break
  fi
  echo "    [retry] uv attempt $_i failed, sleeping 3s..." >&2
  sleep 3
done
[[ "$_uv_done" == "1" ]] || { echo "[error] uv install failed after retries" >&2; exit 1; }
# The installer writes to /root/.local/bin; export BEFORE the existence check
# so `command -v uv` can resolve it.
export PATH="/root/.local/bin:$PATH"
command -v uv >/dev/null 2>&1 || { echo "[error] uv binary missing after install" >&2; exit 1; }

# Stage pre-downloaded managed cpython interpreters (3.11/3.12/3.13) so that
# `uv sync` for MCP servers (arxiv=3.11, emails=3.12, cli-mcp-server=3.13,
# youtube=3.13, pdf-tools=3.12, google-sheets=3.11, scholarly=3.11,
# yahoo-finance=3.11) does NOT need to download them from GitHub at build time.
# The cluster proxy (7893) is flaky and intermittently refuses connections,
# which made cpython downloads fail mid-build and ship a broken image with
# missing .venvs (case-study 2026-08-12, c2n/c2o builds: cli-mcp-server .venv
# missing → FATAL). With these interpreters staged, uv finds them locally and
# never touches the network for Python itself.
#
# The cache is populated on the HOST (outside the container) by:
#   UV_PYTHON_INSTALL_DIR=<this dir> uv python install 3.11 3.12 3.13
# and copied into the rootfs by the build script (OUTSIDE the provision
# heredoc, where host paths are valid). Inside the provision heredoc we only
# reference the container-absolute path /opt/uv_python_cache.
UV_PY_CONTAINER_CACHE="/opt/uv_python_cache"
export UV_PYTHON_INSTALL_DIR="$UV_PY_CONTAINER_CACHE"
export UV_PYTHON_PREFERENCE=managed

echo "=== [3/6] Node.js 22 ==="
# NodeSource is unusable here: its setup_22.x script reports "Exit Code: 0" even
# when the signing-key import fails (curl SSL error), the nodesource apt list is
# then absent, and `apt-get install nodejs` silently resolves to Ubuntu's
# nodejs-12 (no npm binary). The npm loop below then fails 3x on
# "npm: command not found" and every downstream npm build dies. Instead install
# the official node 22 tarball from npmmirror (directly reachable, no proxy) and
# hard-assert the result: node >= 22 AND a working npm, or the build aborts.
NODE_MAJOR=22
NODE_MIRROR="https://cdn.npmmirror.com/binaries/node"
# SHASUMS256.txt line: "<sha256>  node-v22.20.0-linux-x64.tar.gz". The download
# path is /node/<vX.Y.Z>/<filename> — derive the version dir from the filename
# (node-v22.20.0-linux-x64.tar.gz → v22.20.0), don't use the filename as the dir.
_node_file="$(curl -fsSL --connect-timeout 10 --max-time 30 \
  "${NODE_MIRROR}/latest-v${NODE_MAJOR}.x/SHASUMS256.txt" \
  | awk '/linux-x64\.tar\.gz$/ {print $2; exit}')"
if [[ -z "$_node_file" ]]; then
  echo "[error] cannot resolve latest node${NODE_MAJOR} from npmmirror" >&2
  exit 1
fi
_node_dir="${_node_file#node-v}"   # 22.20.0-linux-x64.tar.gz
_node_dir="v${_node_dir%%-*}"      # v22.20.0
_node_tgz="/tmp/${_node_file}"
_node_ok=0
for _i in 1 2 3; do
  if curl -fL --connect-timeout 10 --max-time 600 -o "$_node_tgz" \
    "${NODE_MIRROR}/${_node_dir}/${_node_file}"; then
    _node_ok=1; break
  fi
  echo "    [retry] node tarball attempt $_i failed, sleeping 5s..." >&2
  sleep 5
done
[[ "$_node_ok" == "1" ]] || { echo "[error] node tarball download failed" >&2; exit 1; }
# .tar.gz (not .tar.xz): the base rootfs may lack xz-utils; gzip is always present.
tar -xzf "$_node_tgz" -C /usr/local --strip-components=1
rm -f "$_node_tgz"
# Fail-fast gates: without these the build ships a node-12/no-npm image and all
# downstream npm builds fail silently (only [WARN]).
if ! command -v npm >/dev/null 2>&1; then
  echo "[error] npm not on PATH after node install — aborting" >&2
  exit 1
fi
if [[ "$(node -v | sed 's/^v//; s/\..*//')" -lt "$NODE_MAJOR" ]]; then
  echo "[error] node $(node -v) < ${NODE_MAJOR}.x — stale node shadowing /usr/local, aborting" >&2
  exit 1
fi

echo "=== [4/6] Python venv + deps ==="
uv venv /opt/venv
# NOTE: keep the package list contiguous (no inline `#` comments between the
# backslash-continued lines — a `\` followed by `# comment` breaks the
# continuation and the next package string becomes a standalone command).
# camel-ai 0.2.x imports `from mcp.server import FastMCP`, but mcp>=2.0 removed
# it, so pin mcp to 1.x (last compatible: 1.29.0) to avoid ImportError at eval
# time. Order matters: list the mcp pin FIRST so pip's resolver honours it for
# camel-ai's transitive mcp dependency.
_uv_pkgs=(
  "mcp>=1.3.0,<2.0.0" \
  "camel-ai" \
  "anthropic" \
  "psycopg2-binary" \
  "openpyxl" \
  "python-docx" \
  "python-pptx" \
  "pypdf" \
  "PyPDF2" \
  "pdfplumber" \
  "termcolor" \
  "aiofiles" \
  "psutil" \
  "addict" \
  "arxiv" \
  "bibtexparser" \
  "canvasapi" \
  "prompt_toolkit" \
  "playwright"
)
for _i in 1 2 3 4 5; do
  if uv pip install --python /opt/venv/bin/python "${_uv_pkgs[@]}"; then
    break
  fi
  echo "    [retry] uv pip install attempt $_i failed, sleeping 5s..." >&2
  sleep 5
done
export PATH="/opt/venv/bin:$PATH"
export VIRTUAL_ENV="/opt/venv"
# Chromium / Playwright deps MUST install successfully — the image ships to
# compute nodes with no external network, so a missing binary cannot be
# recovered at runtime (case-study 2026-08-12: arxiv-survey Chromium missing
# → browser_install EAI_AGAIN → false failures). Remove the `|| true` so a
# failed install fails the build loudly instead of shipping a broken image.
# `playwright install chromium` fetches the browser from cdn.playwright.dev,
# which is NOT reachable directly on this cluster. Run it WITH the cluster
# proxy exported and wrap in `timeout` so a hang fails fast and retries
# (case-study 2026-08-12: c2g build hung at "0% of 184.3 MiB" for 6+ min
# because playwright install ignored http_proxy).
#
# c2l update: even WITH the proxy, cdn.playwright.dev is flaky (repeated
# "fetch failed: path/host undefined" then retry, never completing). The fix
# is to point Playwright at the npmmirror CDN (cdn.npmmirror.com), which IS
# directly reachable on this cluster (Tengine, HTTP/2 200). Playwright honors
# PLAYWRIGHT_DOWNLOAD_HOST for CFT browser zips.
_pw_ok=0
if [[ -d /opt/playwright_cache/chromium-1234 ]]; then
  # Pre-staged cache (populated on the host, outside the container): playwright's
  # downloader cannot resume — every `timeout` kill restarts the 115MB headless-
  # shell zip from 0% (case-study 2026-08-14: 3 retries × 600s all died at ~80%,
  # ~250KB/s). Stage the complete cache instead; `playwright install` then only
  # verifies markers and is a no-op.
  echo "    [staged] using pre-staged playwright cache from /opt/playwright_cache"
  mkdir -p /root/.cache
  cp -a /opt/playwright_cache /root/.cache/ms-playwright
else
  for _i in 1 2 3; do
    if env http_proxy=${CLUSTER_PROXY} https_proxy=${CLUSTER_PROXY} \
          HTTP_PROXY=${CLUSTER_PROXY} HTTPS_PROXY=${CLUSTER_PROXY} \
          PLAYWRIGHT_DOWNLOAD_HOST=https://cdn.npmmirror.com/binaries/playwright \
          timeout 900 playwright install chromium; then
      _pw_ok=1; break
    fi
    echo "    [retry] playwright install chromium attempt $_i failed, sleeping 5s..." >&2
    sleep 5
  done
  [[ "$_pw_ok" == "1" ]] || { echo "[error] playwright install chromium failed after retries" >&2; exit 1; }
fi
playwright install-deps chromium
# Hard gate: the chromium binary must actually exist after install, or the
# image is unusable for any playwright task.
#
# Playwright has shipped two on-disk layouts:
#  - legacy:  chromium-<rev>/chrome-linux/chrome
#  - CFT:     chromium-<rev>/chrome-linux64/chrome   (Chrome for Testing, used
#             by playwright >= 1.49 / "playwright chromium vNNNN" log lines)
# Accept either; the gate's job is "is there a chrome binary at all", not which
# subdir layout this exact playwright version happens to use. A wrong glob here
# fails a successful build (case-study 2026-08-12: c2h shipped Chromium fine but
# the chrome-linux-only glob triggered a false FATAL).
if ! ls /root/.cache/ms-playwright/chromium-*/chrome-linux/chrome >/dev/null 2>&1 \
  && ! ls /root/.cache/ms-playwright/chromium-*/chrome-linux64/chrome >/dev/null 2>&1; then
  echo "    [FATAL] Playwright Chromium binary missing after install — refusing to ship a broken image." >&2
  exit 1
fi

echo "=== [5/6] local_servers (from /src) ==="
# Never ship host-side .venvs into the container: a .venv materialized on the
# host carries host-absolute interpreter paths in entry scripts / .pth files
# (the exact c3b contamination vector). `uv sync` below rebuilds every .venv
# INSIDE the container with container-correct paths; rsync --delete plus this
# --exclude guarantees no host venv can leak through.
rsync -a --delete --exclude '.venv' --exclude 'node_modules' \
  /src/local_servers/ /opt/local_servers/

# Purge stale build artifacts BEFORE building. rsync copies host-side bin/dist/build
# which may be outdated (e.g. notion-mcp-server shipped a cli.mjs built before its
# PgHttpClient patch landed → container ran the old bundle → hit live Notion API
# → 401 "API token is invalid"). Removing them forces `npm run build` to regenerate
# from current source; if build then fails, no stale artifact is left behind.
for dir in /opt/local_servers/*/; do
  [[ -f "$dir/package.json" ]] || continue
  rm -rf "$dir/bin" "$dir/dist" "$dir/build" 2>/dev/null || true
done

for dir in \
  /opt/local_servers/Calendar-Autoauth-MCP-Server \
  /opt/local_servers/google-forms-mcp \
  /opt/local_servers/youtube-mcp-server \
  /opt/local_servers/filesystem \
  /opt/local_servers/HowToCook-mcp \
  /opt/local_servers/servers \
  /opt/local_servers/mcp-canvas-lms \
  /opt/local_servers/notion-mcp-server \
  /opt/local_servers/mcp-npx-fetch \
  /opt/local_servers/playwright-mcp \
  /opt/local_servers/woocommerce-mcp \
  /opt/local_servers/12306-mcp; do
  if [[ -f "$dir/package.json" ]]; then
    echo "=== npm: $dir ==="
    _npm_ok=0
    # tsc occasionally gets OOM-killed on this memory-capped login node
    # (case-study 2026-08-15: google-forms-mcp tsc "Killed" mid-build, only a
    # [WARN] shipped, and the c5 batch lost every google_forms task to a
    # 7-second infra_failed). Retry the BUILD (not just install) so an
    # OOM-killed tsc gets another chance instead of shipping broken artifacts.
    _build_ok=0
    for _i in 1 2 3 4 5; do
      if (cd "$dir" && npm install --ignore-scripts); then
        _npm_ok=1
        break
      fi
      echo "    [retry] npm install attempt $_i failed for $dir, sleeping 5s..." >&2
      sleep 5
    done
    for _i in 1 2 3; do
      if [[ $_npm_ok -eq 1 ]] && (cd "$dir" && npm run build); then
        _build_ok=1
        break
      fi
      echo "    [retry] npm build attempt $_i failed for $dir, sleeping 5s..." >&2
      sleep 5
    done
    if [[ $_build_ok -eq 1 ]]; then
      echo "    [ok] built $dir"
    else
      # Previously this was `(npm run build 2>/dev/null || true) || true`, which
      # silently swallowed failures and produced incomplete dist/ output (e.g.
      # woocommerce-mcp shipped without dist/services/* -> broken MCP). Warn loudly.
      echo "    [WARN] BUILD FAILED for $dir — server may be non-functional at runtime" >&2
    fi
  fi
done

# Sanity check: woocommerce-mcp must ship a complete dist (services + pg backend).
if [[ ! -f /opt/local_servers/woocommerce-mcp/dist/services/pg-rest-server.js ]]; then
  echo "    [WARN] woocommerce-mcp/dist/services/pg-rest-server.js missing — 8081 PG backend will not start" >&2
fi

# Hard gate: every npm server's runtime entry must exist after the build loop.
# A missing entry script fails kimi_main's MCP health check at eval time and
# kills every task that grants that server (c5 case-study: google-forms-mcp
# shipped without build/index.js -> canvas-course-feedback infra_failed in 7s).
NPM_REQUIRED_ENTRIES=(
  "Calendar-Autoauth-MCP-Server:build/index.js"
  "google-forms-mcp:build/index.js"
  "youtube-mcp-server:dist/index.js"
  "mcp-canvas-lms:build/index.js"
  "12306-mcp:build/index.js"
  "filesystem:dist/index.js"
  "HowToCook-mcp:build/index.js"
  "servers:src/memory/dist/index.js"
  "mcp-npx-fetch:dist/index.js"
  "notion-mcp-server:bin/cli.mjs"
  "woocommerce-mcp:dist/index.js"
)
for _e in "${NPM_REQUIRED_ENTRIES[@]}"; do
  _proj="${_e%%:*}"; _entry="${_e##*:}"
  # Entries are "proj:relpath-from-/opt/local_servers/proj". Some entries live
  # in nested workspaces (e.g. servers/src/memory) so _entry may itself
  # contain path segments — the check below joins them correctly regardless.
  if [[ ! -f "/opt/local_servers/$_proj/$_entry" ]]; then
    echo "    [FATAL] /opt/local_servers/$_proj/$_entry missing after build — refusing to ship a broken image." >&2
    exit 1
  fi
  echo "    [ok] $_proj/$_entry present"
done

# Sanity check: notion-mcp-server must ship a cli.mjs that contains the PgHttpClient
# patch. Without it the server falls back to the live Notion API and returns 401
# "API token is invalid" (the ntn-placeholder token never reaches a real account).
NOTION_CLI=/opt/local_servers/notion-mcp-server/bin/cli.mjs
if [[ -f "$NOTION_CLI" ]]; then
  if ! grep -q "PgHttpClient" "$NOTION_CLI" 2>/dev/null; then
    echo "    [WARN] notion-mcp-server/cli.mjs is missing PgHttpClient — server will hit live Notion API (401)" >&2
  fi
else
  echo "    [WARN] notion-mcp-server/bin/cli.mjs missing — notion MCP will not start" >&2
fi

# yahoo-finance-mcp requires Python >=3.11 (.python-version = 3.11) but the
# ubuntu:22.04 base image only ships python3.10. uv can download a managed
# cpython-3.11 from python-build-standalone (GitHub releases), but that host
# is NOT reachable directly on this cluster — it hangs forever at
# "Downloading cpython-3.11.15 (29.5MiB)".
#
# Fix (case-study 2026-08-12, c2i/c2j builds):
#  - Do NOT set UV_PYTHON_PREFERENCE=only-system globally (c2j tried this; it
#    made uv refuse to download 3.11 at all → "No interpreter found").
#  - Instead, run uv with the cluster proxy exported so the managed-cpython
#    download goes through 7893, and wrap in `timeout` so a hang fails fast.
#  - Re-enable managed downloads just for this block by unsetting
#    UV_PYTHON_PREFERENCE (in case it was inherited) and pointing uv at the proxy.
export UV_PYTHON_INSTALL_DIR="${UV_PYTHON_INSTALL_DIR:-/root/.local/share/uv/python}"

YAHOO_FINANCE_MCP=/opt/local_servers/yahoo-finance-mcp
if [[ -f "$YAHOO_FINANCE_MCP/pyproject.toml" ]]; then
  echo "=== uv lock/sync: $YAHOO_FINANCE_MCP (python 3.11 via managed download) ==="
  _yf_ok=0
  for _i in 1 2 3; do
    if (cd "$YAHOO_FINANCE_MCP" \
        && env http_proxy=${CLUSTER_PROXY} https_proxy=${CLUSTER_PROXY} \
               HTTP_PROXY=${CLUSTER_PROXY} HTTPS_PROXY=${CLUSTER_PROXY} \
               UV_PYTHON_PREFERENCE=managed \
               timeout 600 uv lock \
        && env http_proxy=${CLUSTER_PROXY} https_proxy=${CLUSTER_PROXY} \
               HTTP_PROXY=${CLUSTER_PROXY} HTTPS_PROXY=${CLUSTER_PROXY} \
               UV_PYTHON_PREFERENCE=managed \
               timeout 600 uv sync); then
      _yf_ok=1; break
    fi
    echo "    [retry] yahoo-finance uv attempt $_i failed, sleeping 5s..." >&2
    sleep 5
  done
  if [[ "$_yf_ok" != "1" ]]; then
    echo "    [WARN] yahoo-finance-mcp uv sync failed — yfinance MCP may be non-functional at runtime" >&2
  fi
  if [[ -x "$YAHOO_FINANCE_MCP/.venv/bin/python" ]]; then
    "$YAHOO_FINANCE_MCP/.venv/bin/python" <<'PY'
import pandas  # noqa: F401
import psycopg2  # noqa: F401
import yfinance

assert yfinance.__version__, yfinance.__file__
print("Yahoo Finance MCP local shim imports")
PY
  fi
fi

# uv-based MCP servers: fail-fast `uv lock && uv sync` with a smoke check.
# Previously this loop used `(cd "$dir" && uv sync) || true`, which silently
# swallowed build failures (e.g. missing deps, lock drift, network hiccups)
# and shipped a broken/empty `.venv`. At runtime `uv run` then tried to
# rebuild from source inside the no-egress eval container, blew past the MCP
# startupTimeoutMs, and kimi dropped the server entirely — which is exactly
# the 20260807-224427 batch failure (arxiv_local / excel / emails / snowflake
# all "Timed out after 90000ms"). mcp_json_gen.py now pins VIRTUAL_ENV +
# UV_FROZEN + UV_OFFLINE so `uv run` reuses this .venv deterministically;
# this block is what makes that .venv actually exist and be complete.
#
# Fail-fast policy (case-study 2026-08-12, P1 #11): a missing .venv cannot be
# recovered at runtime (compute nodes have no external network), so a build
# that ships without it produces silent false-failures. cli-mcp-server is
# REQUIRED (46 of the 91 §C.1 cases depend on the terminal MCP); a missing
# cli-mcp-server .venv now fails the build. mcp-google-sheets is REQUIRED too:
# the §C.2 rerun failures (canvas-course-comparison / arxiv-latex-review-notion-word,
# case-study 2026-08-13) traced to its .venv being silently missing, which made
# the `google_sheet` MCP fail to register at runtime ("Tool NOT FOUND"). A
# missing mcp-google-sheets .venv now fails the build.
#
# P0 hardening (case-study 2026-08-14, clean-and-rerun-v2 §5 P0): ALL uv-based
# MCP servers are now REQUIRED. The c3b batch lost 22 cases because excel /
# terminal / word / pptx .venvs shipped broken (host-absolute paths inside venv
# entry scripts and .pth files) and the build only fail-fasted on 2 of 12.
# Every server below must have a complete .venv or the build dies. This also
# guards against the manual-host-edit contamination vector: the smoke test and
# the host-path grep below make a polluted venv impossible to ship.
UV_REQUIRED_DIRS=(
  cli-mcp-server
  mcp-google-sheets
  excel-mcp-server
  Office-Word-MCP-Server
  Office-PowerPoint-MCP-Server
  emails-mcp
  arxiv-mcp-server
  arxiv-latex-mcp
  mcp-scholarly
  mcp-snowflake-server
  pdf-tools-mcp
  mcp-youtube-transcript
)
for dir in \
  /opt/local_servers/arxiv-mcp-server \
  /opt/local_servers/arxiv-latex-mcp \
  /opt/local_servers/emails-mcp \
  /opt/local_servers/mcp-snowflake-server \
  /opt/local_servers/mcp-scholarly \
  /opt/local_servers/Office-Word-MCP-Server \
  /opt/local_servers/Office-PowerPoint-MCP-Server \
  /opt/local_servers/excel-mcp-server \
  /opt/local_servers/pdf-tools-mcp \
  /opt/local_servers/mcp-youtube-transcript \
  /opt/local_servers/mcp-google-sheets \
  /opt/local_servers/cli-mcp-server; do
  if [[ -f "$dir/pyproject.toml" ]]; then
    echo "=== uv lock/sync: $dir ==="
    _uv_ok=0
    # Many MCP servers pin python 3.11/3.12/3.13 via .python-version, but the
    # base image only ships 3.10. uv will try to download managed cpython from
    # python-build-standalone (GitHub), which is unreachable directly here, so
    # run uv with the cluster proxy exported + `timeout` so a hang fails fast
    # (case-study 2026-08-12, c2i/c2j). UV_PYTHON_PREFERENCE=managed
    # lets uv fetch the pinned interpreter via the proxy.
    if (cd "$dir" \
        && env http_proxy=${CLUSTER_PROXY} https_proxy=${CLUSTER_PROXY} \
               HTTP_PROXY=${CLUSTER_PROXY} HTTPS_PROXY=${CLUSTER_PROXY} \
               UV_PYTHON_PREFERENCE=managed \
               timeout 600 uv lock \
        && env http_proxy=${CLUSTER_PROXY} https_proxy=${CLUSTER_PROXY} \
               HTTP_PROXY=${CLUSTER_PROXY} HTTPS_PROXY=${CLUSTER_PROXY} \
               UV_PYTHON_PREFERENCE=managed \
               timeout 600 uv sync); then
      echo "    [ok] synced $dir"
      _uv_ok=1
    else
      echo "    [WARN] UV SYNC FAILED for $dir — MCP server will be non-functional at runtime" >&2
    fi
    _base="$(basename "$dir")"
    if [[ -x "$dir/.venv/bin/python" ]]; then
      echo "    [ok] $dir/.venv ready"
    else
      echo "    [WARN] $dir/.venv/bin/python missing — uv run will fail to start this server" >&2
      # Fail-fast for REQUIRED servers (cli-mcp-server underpins 46 §C.1 cases).
      for _req in "${UV_REQUIRED_DIRS[@]}"; do
        if [[ "$_base" == "$_req" ]]; then
          echo "    [FATAL] required uv server $_req has no .venv — refusing to ship a broken image." >&2
          exit 1
        fi
      done
    fi
  fi
done

# ---------------------------------------------------------------------------
# P0 hardening (case-study 2026-08-14): venv integrity gates.
#
# The c3b batch lost 22 cases to venvs that EXISTED but were broken: entry
# scripts (dotenv/httpx/mcp/uvicorn/...) carried HOST-absolute interpreter
# paths because someone ran `uv sync` on the host copy of the rootfs. A
# missing .venv is caught above; these gates catch the poisoned variant.
#
# Gate 1 — import smoke test: every REQUIRED server's own entry point must
#          import under its venv python (uv run executes the same code).
# Gate 2 — host-path grep: no file inside any uv .venv may reference the
#          host prefixes that only exist outside the container. Inside the
#          build container these strings are canaries: their presence means
#          the venv was materialized on the host, not here.
# ---------------------------------------------------------------------------
UV_ENTRY_MODULES=(
  "excel-mcp-server:excel_mcp"
  "Office-Word-MCP-Server:word_document_server"
  "Office-PowerPoint-MCP-Server:ppt_mcp_server"
  "cli-mcp-server:cli_mcp_server"
  "emails-mcp:emails_mcp"
  "arxiv-mcp-server:arxiv_mcp_server"
  "arxiv-latex-mcp:arxiv_to_prompt"
  "mcp-scholarly:mcp_scholarly"
  "mcp-snowflake-server:mcp_snowflake_server"
  "pdf-tools-mcp:pdf_tools_mcp"
  "mcp-youtube-transcript:mcp_youtube_transcript"
  "mcp-google-sheets:mcp_google_sheets"
)
echo "=== venv integrity gates (P0) ==="
for _entry in "${UV_ENTRY_MODULES[@]}"; do
  _proj="${_entry%%:*}"
  _mod="${_entry##*:}"
  _dir="/opt/local_servers/$_proj"
  [[ -f "$_dir/pyproject.toml" ]] || continue
  _py="$_dir/.venv/bin/python"
  if [[ ! -x "$_py" ]]; then
    echo "    [FATAL] $_proj: $_py missing — refusing to ship." >&2
    exit 1
  fi
  # Gate 1: import the server's own top-level module under its venv python.
  # cli_mcp_server instantiates its CommandExecutor at import time and needs
  # an existing ALLOWED_DIR; any writable dir works for a smoke import.
  if ! ALLOWED_DIR=/tmp "$_py" -c "import importlib; importlib.import_module('$_mod')"; then
    echo "    [FATAL] $_proj: cannot import '$_mod' under its venv python — refusing to ship." >&2
    exit 1
  fi
  echo "    [ok] $_proj: import smoke passed ($_mod)"
  # Gate 2: entry scripts must not carry host prefixes in shebang/exec lines.
  for _s in "$_dir/.venv/bin"/*; do
    [[ -f "$_s" && ! -L "$_s" ]] || continue
    if head -c 4096 "$_s" | grep -qE '/lintaoLab2/|/storage/lintaoLab/'; then
      echo "    [FATAL] $_proj: host path inside $(basename "$_s") — venv was built on the host. Refusing to ship." >&2
      exit 1
    fi
  done
done
# Gate 2b: whole-venv host-prefix grep (.pth / direct_url.json / RECORD carry
# absolute paths and poison `uv run` at runtime even when entry scripts look fine).
_HOST_PATH_HITS="$(grep -rIl -E '/lintaoLab2/|/storage/lintaoLab/' /opt/local_servers/*/.venv/ 2>/dev/null || true)"
if [[ -n "$_HOST_PATH_HITS" ]]; then
  echo "    [FATAL] host-absolute paths leaked into venv files:" >&2
  echo "$_HOST_PATH_HITS" | head -20 >&2
  exit 1
fi
echo "    [ok] no host paths under any .venv"

echo "=== [6/6] copy project into /workspace ==="
mkdir -p /workspace
rsync -a --delete \
  --exclude dumps --exclude benchmark_logs --exclude .git \
  --exclude '__pycache__' --exclude '*.pyc' \
  /src/ /workspace/

# Persist PATH for later enroot starts
cat >/etc/profile.d/toolathlon.sh <<'EOF'
export PATH="/opt/venv/bin:/root/.local/bin:$PATH"
export VIRTUAL_ENV="/opt/venv"
export LOCAL_SERVERS_PATH=/opt/local_servers
export PYTHON_BIN=/opt/venv/bin/python3
EOF

echo "=== provision complete ==="
PROVISION
chmod +x "$PROVISION_SCRIPT"

# ---------------------------------------------------------------------------
# Create / refresh build rootfs from ubuntu
# ---------------------------------------------------------------------------
# Build-rootfs cleanup (#6): never leave the build instance behind, even on failure.
cleanup_build() {
  if [[ -d "${ENROOT_DATA_PATH}/${BUILD_NAME}" ]]; then
    log "Cleaning up build rootfs $BUILD_NAME ..."
    enroot remove -f "$BUILD_NAME" 2>/dev/null || rm -rf "${ENROOT_DATA_PATH}/${BUILD_NAME}" 2>/dev/null || true
  fi
}
trap cleanup_build EXIT

log "Preparing build rootfs: $BUILD_NAME"
if [[ -d "${ENROOT_DATA_PATH}/${BUILD_NAME}" ]]; then
  log "Removing existing $BUILD_NAME ..."
  enroot remove -f "$BUILD_NAME" 2>/dev/null || rm -rf "${ENROOT_DATA_PATH}/${BUILD_NAME}"
fi
enroot create -n "$BUILD_NAME" "$TOOLATHLON_UBUNTU_SQSH"

# Avoid enroot -m bind mounts on this NFS cluster (often fail with
# "No such file or directory" / "No such device"). Copy inputs into rootfs instead.
log "Staging project + provision script into rootfs (no bind mounts) ..."
ROOTFS="${ENROOT_DATA_PATH}/${BUILD_NAME}"
# NOTE: enroot typically mounts a tmpfs over /tmp, so do NOT place scripts under /tmp.
mkdir -p "$ROOTFS/src" "$ROOTFS/opt"
rsync -a --delete \
  --exclude dumps --exclude benchmark_logs --exclude .git \
  --exclude '__pycache__' --exclude '*.pyc' \
  --exclude toolathlon_gym_eval_dockers \
  "${PROJECT_ROOT}/" "$ROOTFS/src/"
cp -f "$PROVISION_SCRIPT" "$ROOTFS/opt/provision_agent.sh"
chmod +x "$ROOTFS/opt/provision_agent.sh"

# Stage pre-downloaded managed cpython interpreters (3.11/3.12/3.13) into the
# rootfs BEFORE provisioning, so `uv sync` inside the container finds them at
# /opt/uv_python_cache and never needs GitHub (proxy is flaky — see comment in
# the provision heredoc). This runs on the HOST where RUNTIME_ROOT is valid.
UV_PY_HOST_CACHE="${RUNTIME_ROOT}/uv_python_cache"
if [[ -d "$UV_PY_HOST_CACHE" ]]; then
  log "Staging cpython cache (3.11/3.12/3.13) into rootfs..."
  mkdir -p "$ROOTFS/opt/uv_python_cache"
  rsync -a "$UV_PY_HOST_CACHE/" "$ROOTFS/opt/uv_python_cache/"
else
  log "WARN: $UV_PY_HOST_CACHE not found — uv will attempt GitHub download at sync time"
fi

# Stage the playwright browser cache the same way. Populated on the host from a
# prior good rootfs (chromium-1234 + headless shell + ffmpeg, INSTALLATION_COMPLETE
# markers included). `playwright install` inside the container then verifies and
# no-ops instead of re-downloading 300MB from a flaky CDN.
PW_HOST_CACHE="${RUNTIME_ROOT}/playwright_cache"
if [[ -d "$PW_HOST_CACHE/chromium-1234" ]]; then
  log "Staging playwright browser cache into rootfs..."
  mkdir -p "$ROOTFS/opt/playwright_cache"
  rsync -a "$PW_HOST_CACHE/" "$ROOTFS/opt/playwright_cache/"
else
  log "WARN: $PW_HOST_CACHE incomplete — playwright will download inside the container"
fi

log "Provisioning inside enroot (log: $BUILD_LOG) ..."
log "This can take a long time (apt + npm + uv). Keep proxy on."

enroot start -r -w \
  -e "CLUSTER_PROXY=${CLUSTER_PROXY}" \
  -e "http_proxy=${http_proxy:-}" \
  -e "https_proxy=${https_proxy:-}" \
  -e "HTTP_PROXY=${HTTP_PROXY:-}" \
  -e "HTTPS_PROXY=${HTTPS_PROXY:-}" \
  -e "no_proxy=${no_proxy:-}" \
  -e "NO_PROXY=${NO_PROXY:-}" \
  -e "USE_DIRECT_MIRRORS=${USE_DIRECT_MIRRORS:-0}" \
  -e "DEBIAN_FRONTEND=noninteractive" \
  "$BUILD_NAME" \
  bash /opt/provision_agent.sh \
  2>&1 | tee "$BUILD_LOG"

# Drop staged source copy to shrink the exported image. The .venvs created by
# uv symlink to the managed interpreters under /opt/uv_python_cache, so we MUST
# KEEP that cache or the .venv python symlinks break (case-study 2026-08-12:
# c2s build shipped 1.6GB sqsh but every .venv pointed at a non-existent
# interpreter because the cache was deleted to save space — runtime MCP servers
# all failed to start). Only remove /src (the staged project copy; workspace
# already has its own).
rm -rf "$ROOTFS/src" || true

# Show rootfs size before the (memory-constrained) squashfs compress.
log "Rootfs size before export: $(du -sh "$ROOTFS" 2>/dev/null | cut -f1)"

log "Exporting $TOOLATHLON_AGENT_SQSH ..."
rm -f "$TOOLATHLON_AGENT_SQSH"
# mksquashfs thread/memory policy. On the LOGIN node the user cgroup has a
# 12GB cap: with default threads mksquashfs OOM-killed mid-compress and shipped
# a truncated .sqsh (case-study 2026-08-12). On COMPUTE nodes there is no such
# cap, so use full parallelism and normal compression — faster AND smaller.
if [[ -n "${SLURM_JOB_ID:-}" ]]; then
  log "Running under Slurm (job ${SLURM_JOB_ID}) — full-parallel mksquashfs"
  unset ENROOT_MAX_PROCESSORS 2>/dev/null || true
  unset ENROOT_SQUASH_OPTIONS 2>/dev/null || true
else
  log "Running on login node — single-thread low-mem mksquashfs (12GB cgroup cap)"
  export ENROOT_MAX_PROCESSORS=1
  export ENROOT_SQUASH_OPTIONS="-noI -no-xattrs -mem 256M"
fi

# Export to tmpfs first, then mv to NFS. Streaming mksquashfs output directly
# to NFS risks a mid-write hang (the same NFS stalls that deadlocked
# unsquashfs reads in earlier builds) — and the truncated-sqsh case-study
# showed that a partial export can look like a finished image. mv is atomic
# on the same filesystem, so either the FULL image lands or nothing does.
SQSH_TMP="${ENROOT_DATA_PATH}/toolathlon-pack.sqsh.partial"
rm -f "$SQSH_TMP"
enroot export -o "$SQSH_TMP" "$BUILD_NAME"
[[ -s "$SQSH_TMP" ]] || die "enroot export produced empty $SQSH_TMP"
mv "$SQSH_TMP" "$TOOLATHLON_AGENT_SQSH"

# The runtime rootfs is NOT a build artifact — it's a convenience for host-side
# `enroot start` debugging. It costs a full ~6.9G unsquash on the build node's
# /dev/shm. Skip by default; set CREATE_RUNTIME_ROOTFS=1 if you actually need it.
if [[ "${CREATE_RUNTIME_ROOTFS:-0}" == "1" ]]; then
  log "Creating runtime rootfs: toolathlon-pack"
  if [[ -d "${ENROOT_DATA_PATH}/toolathlon-pack" ]]; then
    enroot remove -f toolathlon-pack 2>/dev/null || rm -rf "${ENROOT_DATA_PATH}/toolathlon-pack"
  fi
  enroot create -n toolathlon-pack "$TOOLATHLON_AGENT_SQSH"
else
  log "Skipping runtime rootfs creation (set CREATE_RUNTIME_ROOTFS=1 to enable)"
fi

ls -lh "$TOOLATHLON_AGENT_SQSH"
log "Done. Agent image ready: $TOOLATHLON_AGENT_SQSH"
