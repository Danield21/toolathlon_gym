#!/bin/bash
# Build toolathlon-pack.sqsh from ubuntu:22.04 via enroot (replaces: docker build -t toolathlon-pack:latest .)
#
# Usage:
#   bash scripts/enroot_build_agent.sh
#
# Requires proxy for apt/npm/pip on this cluster:
#   export http_proxy=http://127.0.0.1:7891 https_proxy=http://127.0.0.1:7891

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
RUNTIME_ROOT="${TOOLATHLON_EVAL_DOCKER_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers}"

# shellcheck disable=SC1091
source "$RUNTIME_ROOT/env.sh"

BUILD_NAME="toolathlon-pack-build"
PROVISION_SCRIPT="${RUNTIME_ROOT}/tmp/provision_agent.sh"
BUILD_LOG="${RUNTIME_ROOT}/logs/build_agent.log"

export http_proxy="${http_proxy:-http://127.0.0.1:7891}"
export https_proxy="${https_proxy:-http://127.0.0.1:7891}"
export HTTP_PROXY="$http_proxy" HTTPS_PROXY="$https_proxy"
export no_proxy="${no_proxy:-127.0.0.1,localhost}"
export NO_PROXY="$no_proxy"

die() { echo "[error] $*" >&2; exit 1; }
log() { echo "[$(date +%H:%M:%S)] $*"; }

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

# Honor proxy for apt
if [[ -n "${http_proxy:-}" ]]; then
  cat >/etc/apt/apt.conf.d/95proxy <<EOF
Acquire::http::Proxy "${http_proxy}";
Acquire::https::Proxy "${https_proxy:-$http_proxy}";
EOF
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
  squashfs-tools

echo "=== [2/6] uv ==="
# curl over the flaky proxy can SSL-error once and abort the whole build.
# Retry the uv installer a few times before giving up.
for _i in 1 2 3 4 5; do
  if curl -LsSf https://astral.sh/uv/install.sh | sh; then
    break
  fi
  echo "    [retry] uv install attempt $_i failed, sleeping 5s..." >&2
  sleep 5
done
# The installer writes to /root/.local/bin; export BEFORE the existence check
# so `command -v uv` can resolve it.
export PATH="/root/.local/bin:$PATH"
command -v uv >/dev/null 2>&1 || { echo "[error] uv install failed after retries" >&2; exit 1; }

echo "=== [3/6] Node.js 22 ==="
# Retry the NodeSource setup script (curl over the proxy is flaky).
for _i in 1 2 3 4 5; do
  if curl -fsSL https://deb.nodesource.com/setup_22.x | bash -; then
    break
  fi
  echo "    [retry] nodesource setup attempt $_i failed, sleeping 5s..." >&2
  sleep 5
done
apt-get install -y nodejs
# Retry global npm upgrade (also network-sensitive).
for _i in 1 2 3; do npm install -g npm@latest && break; sleep 5; done

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
  "termcolor" \
  "aiofiles" \
  "psutil" \
  "addict" \
  "arxiv" \
  "bibtexparser" \
  "canvasapi" \
  "prompt_toolkit"
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
playwright install chromium || true
playwright install-deps chromium || true

echo "=== [5/6] local_servers (from /src) ==="
rsync -a --delete /src/local_servers/ /opt/local_servers/

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
  /opt/local_servers/mcp-google-sheets \
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
    for _i in 1 2 3 4 5; do
      if (cd "$dir" && npm install --ignore-scripts); then
        _npm_ok=1
        break
      fi
      echo "    [retry] npm install attempt $_i failed for $dir, sleeping 5s..." >&2
      sleep 5
    done
    if [[ $_npm_ok -eq 1 ]] && (cd "$dir" && npm run build); then
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

YAHOO_FINANCE_MCP=/opt/local_servers/yahoo-finance-mcp
if [[ -f "$YAHOO_FINANCE_MCP/pyproject.toml" ]]; then
  echo "=== uv lock/sync: $YAHOO_FINANCE_MCP ==="
  (cd "$YAHOO_FINANCE_MCP" && uv lock && uv sync)
  "$YAHOO_FINANCE_MCP/.venv/bin/python" <<'PY'
import pandas  # noqa: F401
import psycopg2  # noqa: F401
import yfinance

assert yfinance.__version__, yfinance.__file__
print("Yahoo Finance MCP local shim imports")
PY
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
  /opt/local_servers/cli-mcp-server; do
  if [[ -f "$dir/pyproject.toml" ]]; then
    echo "=== uv lock/sync: $dir ==="
    if (cd "$dir" && uv lock && uv sync); then
      echo "    [ok] synced $dir"
    else
      echo "    [WARN] UV SYNC FAILED for $dir — MCP server will be non-functional at runtime" >&2
    fi
    if [[ -x "$dir/.venv/bin/python" ]]; then
      echo "    [ok] $dir/.venv ready"
    else
      echo "    [WARN] $dir/.venv/bin/python missing — uv run will fail to start this server" >&2
    fi
  fi
done

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

log "Provisioning inside enroot (log: $BUILD_LOG) ..."
log "This can take a long time (apt + npm + uv). Keep proxy on."

enroot start -r -w \
  -e "http_proxy=${http_proxy}" \
  -e "https_proxy=${https_proxy}" \
  -e "HTTP_PROXY=${HTTP_PROXY}" \
  -e "HTTPS_PROXY=${HTTPS_PROXY}" \
  -e "no_proxy=${no_proxy}" \
  -e "NO_PROXY=${NO_PROXY}" \
  -e "DEBIAN_FRONTEND=noninteractive" \
  "$BUILD_NAME" \
  bash /opt/provision_agent.sh \
  2>&1 | tee "$BUILD_LOG"

# Drop staged source copy to shrink exported image a bit (workspace already has a copy)
rm -rf "$ROOTFS/src" || true

log "Exporting $TOOLATHLON_AGENT_SQSH ..."
rm -f "$TOOLATHLON_AGENT_SQSH"
enroot export -o "$TOOLATHLON_AGENT_SQSH" "$BUILD_NAME"

log "Creating runtime rootfs: toolathlon-pack"
if [[ -d "${ENROOT_DATA_PATH}/toolathlon-pack" ]]; then
  enroot remove -f toolathlon-pack 2>/dev/null || rm -rf "${ENROOT_DATA_PATH}/toolathlon-pack"
fi
enroot create -n toolathlon-pack "$TOOLATHLON_AGENT_SQSH"

ls -lh "$TOOLATHLON_AGENT_SQSH"
log "Done. Agent image ready: $TOOLATHLON_AGENT_SQSH"
