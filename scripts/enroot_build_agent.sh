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
curl -LsSf https://astral.sh/uv/install.sh | sh
export PATH="/root/.local/bin:$PATH"

echo "=== [3/6] Node.js 22 ==="
curl -fsSL https://deb.nodesource.com/setup_22.x | bash -
apt-get install -y nodejs
npm install -g npm@latest

echo "=== [4/6] Python venv + deps ==="
uv venv /opt/venv
uv pip install --python /opt/venv/bin/python \
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
export PATH="/opt/venv/bin:$PATH"
export VIRTUAL_ENV="/opt/venv"
playwright install chromium || true
playwright install-deps chromium || true

echo "=== [5/6] local_servers (from /src) ==="
rsync -a --delete /src/local_servers/ /opt/local_servers/
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
    if (cd "$dir" && npm install --ignore-scripts && npm run build); then
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

YAHOO_FINANCE_MCP=/opt/local_servers/yahoo-finance-mcp
if [[ -f "$YAHOO_FINANCE_MCP/pyproject.toml" ]]; then
  echo "=== uv lock/sync: $YAHOO_FINANCE_MCP ==="
  (cd "$YAHOO_FINANCE_MCP" && uv lock && uv sync)
  "$YAHOO_FINANCE_MCP/.venv/bin/python" <<'PY'
import pandas  # noqa: F401
import psycopg2  # noqa: F401
import yfinance

assert yfinance.__file__.endswith("yfinance.py"), yfinance.__file__
print("Yahoo Finance MCP local shim imports")
PY
fi

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
    echo "=== uv sync: $dir ==="
    (cd "$dir" && uv sync) || true
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
