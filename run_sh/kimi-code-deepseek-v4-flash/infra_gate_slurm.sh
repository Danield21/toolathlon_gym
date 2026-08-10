#!/bin/bash
# Stage-1 infra gate, run on a slurm compute node (no model calls).
#
#   (0) start PostgreSQL on the compute-node HOST (conda initdb; the
#       toolathlon-pack image only ships PG client binaries)
#   (1) MCP server gate   -> test_mcp_servers.py --list-tools (24 servers)
#   (2) preprocess gate   -> preprocess/main.py for all 160 T1-T4 tasks
#
# Why slurm: the login node has a 12G cgroup cap; enroot create of the 6.4G
# rootfs + 160 preprocess runs OOM there. Compute nodes (2TB) don't.
#
# The inner payload is written to a temp script and executed via `srun bash
# <file>` (NOT `bash -c "$VAR"`), so compute-side $VARS are NOT prematurely
# expanded by the login-node shell. Task list + per-task logs are exchanged
# with the container through a bind-mounted dir.
#
# Usage:
#   bash run_sh/kimi-code-deepseek-v4-flash/infra_gate_slurm.sh
#   SLURM_TIME=04:00:00 bash .../infra_gate_slurm.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
GYM="$(cd "$SCRIPT_DIR/../.." && pwd)"

SLURM_MEM="${SLURM_MEM:-64G}"
SLURM_CPUS="${SLURM_CPUS:-32}"
SLURM_TIME="${SLURM_TIME:-04:00:00}"
SLURM_PARTITION="${SLURM_PARTITION:-linlab}"

TASK_LIST="$SCRIPT_DIR/tasks_T1T4_160.txt"
[ -f "$TASK_LIST" ] || { echo "missing $TASK_LIST"; exit 1; }
NTASKS=$(grep -c . "$TASK_LIST")

echo "[infra-gate] partition=$SLURM_PARTITION mem=$SLURM_MEM cpus=$SLURM_CPUS time=$SLURM_TIME tasks=$NTASKS"

# Write the compute-node payload to a NFS-shared temp file (compute nodes can't
# see the login node's /tmp). Single-quoted heredoc => $VARS stay literal and
# are expanded ON the compute node.
PAYLOAD_DIR="$GYM/dumps/infra_gate/.payloads"
mkdir -p "$PAYLOAD_DIR"
PAYLOAD="$(mktemp "$PAYLOAD_DIR/payload_XXXXXXXX.sh")"
cat > "$PAYLOAD" <<'PAYLOAD_EOF'
set -uo pipefail
GYM=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym
cd "$GYM"
STAMP=$(date +%Y%m%d-%H%M%S)
LOGDIR=$GYM/dumps/infra_gate/$STAMP
mkdir -p "$LOGDIR/pp"

# ── proxy hygiene (compute nodes inherit bogus 127.0.0.1:7890) ─────────────
unset http_proxy https_proxy HTTP_PROXY HTTPS_PROXY ALL_PROXY all_proxy || true
export no_proxy="127.0.0.1,localhost,::1,192.168.180.240,172.16.0.0/12,10.0.0.0/8"

# ── enroot: mirror to node-local /dev/shm (avoid NFS dithering) ────────────
ENROOT_SRC="/storage/lintaoLab/bowending/.local/enroot"
ENROOT_LOCAL="/dev/shm/enroot_install"
if [[ ! -x "$ENROOT_LOCAL/bin/enroot" ]]; then
  mkdir -p "$ENROOT_LOCAL"
  rsync -a "$ENROOT_SRC/" "$ENROOT_LOCAL/" 2>/dev/null || cp -a "$ENROOT_SRC/." "$ENROOT_LOCAL/"
fi
CONDA_BIN=/storage/lintaoLab/bowending/miniconda3/envs/toolathlon_gym/bin
export ENROOT_LIBRARY_PATH="${ENROOT_LOCAL}/lib"
export ENROOT_SYSCONF_PATH="${ENROOT_LOCAL}/etc"
export PATH="${ENROOT_LOCAL}/bin:$CONDA_BIN:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin"
export ENROOT_DATA_PATH="/dev/shm/enroot_data"
export ENROOT_TEMP_PATH="/dev/shm/enroot_tmp"
export ENROOT_RUNTIME_PATH="/dev/shm/enroot_runtime"
export ENROOT_CACHE_PATH="/dev/shm/enroot_cache"
mkdir -p "$ENROOT_DATA_PATH" "$ENROOT_TEMP_PATH" "$ENROOT_RUNTIME_PATH" "$ENROOT_CACHE_PATH"

SQSH=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers/images/toolathlon-pack.sqsh
INST=infra-gate-$STAMP

# task list + per-task logs exchanged with the container via a bind-mounted dir
GATE_XCHG=/dev/shm/gate_xchg_$STAMP
mkdir -p "$GATE_XCHG/pp"
cp "$TASK_LIST" "$GATE_XCHG/tasks.txt"

# ── (0) PostgreSQL on the compute-node HOST ────────────────────────────────
# The login node exports polluted PG* vars (PGPORT=0000 etc.) which
# `srun --export=ALL` would carry here and break initdb's internal bootstrap.
# Clear them first, then set exactly what we need.
unset PGPORT PGHOST PGUSER PGPASSWORD PGDATABASE PG_USER PG_PASSWORD PG_HOST PG_PORT PG_DATABASE PGDATA PGSOCKET 2>/dev/null || true
PGUSER=eigent; PGPASSWORD=camel; PGDATABASE=toolathlon_gym
PGDATA=/dev/shm/infra_pg_$STAMP/data
PGSOCK=/dev/shm/infra_pg_$STAMP/sock
PORT=25444
PGLOG="$GATE_XCHG/pg.log"
mkdir -p "$PGDATA" "$PGSOCK"; chmod 700 "$PGDATA" "$PGSOCK"

echo "[inner] (0) start host PostgreSQL (conda initdb) ..."
_pw=$(mktemp); printf '%s\n' "$PGPASSWORD" > "$_pw"
initdb -D "$PGDATA" -U "$PGUSER" --pwfile="$_pw" --auth-local=scram-sha-256 \
  --auth-host=md5 --encoding=UTF8 --locale=C >/dev/null
rm -f "$_pw"
{ echo "listen_addresses = '127.0.0.1'"; echo "port = 25444";
  echo "max_connections = 60"; echo "shared_buffers = 256MB";
  echo "unix_socket_directories = '$PGSOCK'"; } >> "$PGDATA/postgresql.conf"
echo "host all all 127.0.0.1/32 md5" >> "$PGDATA/pg_hba.conf"
pg_ctl -D "$PGDATA" -l "$PGLOG" -o "-k $PGSOCK -p 25444" start >/dev/null 2>&1
for i in $(seq 1 60); do
  env -u PGDATABASE pg_isready -h "$PGSOCK" -p 25444 -U "$PGUSER" -d postgres >/dev/null 2>&1 && break
  sleep 1
done
env -u PGDATABASE pg_isready -h "$PGSOCK" -p 25444 -U "$PGUSER" -d postgres >/dev/null 2>&1 || { echo "[inner] FATAL pg not ready"; tail -15 "$PGLOG"; exit 1; }
PGPASSWORD="$PGPASSWORD" env -u PGDATABASE psql -h "$PGSOCK" -p 25444 -U "$PGUSER" -d postgres -v ON_ERROR_STOP=1 \
  -c "ALTER USER $PGUSER WITH PASSWORD '$PGPASSWORD';" >/dev/null
PGPASSWORD="$PGPASSWORD" env -u PGDATABASE createdb -h "$PGSOCK" -p 25444 -U "$PGUSER" "$PGDATABASE" 2>/dev/null || true
if ! gunzip -c "$GYM/db/init.sql.gz" | PGPASSWORD="$PGPASSWORD" env -u PGDATABASE psql -h "$PGSOCK" -p 25444 -U "$PGUSER" -d "$PGDATABASE" -v ON_ERROR_STOP=1 >"$GATE_XCHG/init_load.log" 2>&1; then
  echo "[inner] FATAL: init.sql.gz load failed"; tail -15 "$GATE_XCHG/init_load.log"; exit 1
fi
PGPASSWORD="$PGPASSWORD" psql -h 127.0.0.1 -p 25444 -U "$PGUSER" -d "$PGDATABASE" -v ON_ERROR_STOP=1 >/dev/null 2>&1 <<SQL
ALTER TABLE email.sent_log DROP CONSTRAINT IF EXISTS sent_log_message_id_fkey;
ALTER TABLE email.sent_log ADD CONSTRAINT sent_log_message_id_fkey
  FOREIGN KEY (message_id) REFERENCES email.messages(id) ON DELETE CASCADE;
SQL
echo "[inner] PostgreSQL ready on host (schemas loaded)"

# default-port symlink so container tools using port 5432 also resolve
ln -sfn ".s.PGSQL.25444" "${PGSOCK}/.s.PGSQL.5432"
ENV_ARGS=(
  -e PGHOST=/run/toolathlon_pg -e PGPORT=25444 -e PGUSER=eigent -e PGPASSWORD=camel -e PGDATABASE=toolathlon_gym
  -e PG_HOST=/run/toolathlon_pg -e PG_USER=eigent -e PG_PASSWORD=camel -e PG_PORT=25444 -e PG_DATABASE=toolathlon_gym
  # Mirror the real agent runtime: MCP servers live in /opt/local_servers and
  # the global venv is /opt/venv. Without these, test_mcp_servers.py probes
  # /workspace/local_servers (which has no .venv) and every uv/node MCP fails
  # for the wrong reason.
  -e LOCAL_SERVERS_PATH=/opt/local_servers
  -e VIRTUAL_ENV=/opt/venv
  -e PATH=/opt/venv/bin:/root/.local/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin
)

echo "[inner] host=$(hostname) enroot=$(enroot version 2>&1 | head -1)"
echo "[inner] /dev/shm: $(df -h /dev/shm | tail -1 | awk '{print $2" tot, "$4" free"}')"

echo "[inner] enroot create $INST from $(basename $SQSH) ..."
if ! enroot create -n "$INST" "$SQSH" >"$LOGDIR/enroot_create.log" 2>&1; then
  echo "[inner] FATAL: enroot create failed"; tail -20 "$LOGDIR/enroot_create.log"; exit 1
fi
echo "[inner] create OK"

# ── MCP + preprocess gates inside ONE container invocation ─────────────────
enroot start -r --rw -m "$PGSOCK:/run/toolathlon_pg" -m "$GATE_XCHG:/gate_xchg" "${ENV_ARGS[@]}" "$INST" bash -lc '
set -uo pipefail
PY=/opt/venv/bin/python3
SRC=/workspace
cd "$SRC"
XC=/gate_xchg

MCPS="12306 arxiv-latex-mcp arxiv_local canvas emails excel filesystem google_calendar google_forms google_sheet howtocook memory notion npx-fetch pdf-tools playwright_with_chunk pptx scholarly_search snowflake terminal woocommerce word yahoo-finance youtube_transcript youtube"

echo "[gate] (1) MCP list-tools gate"
$PY "$SRC/test_mcp_servers.py" --list-tools $MCPS >"$XC/mcp_gate.log" 2>&1
MCP_RC=$?
echo "[gate] MCP list-tools rc=$MCP_RC"; tail -30 "$XC/mcp_gate.log"

echo "[gate] (2) preprocess gate"
PASSN=0; FAILN=0
printf "task\trc\n" > "$XC/pp_summary.tsv"
# preprocess scripts expect --agent_workspace (agent runtime always provides
# it). Create a scratch dir per task so file-creating preprocessors work.
AW_ROOT=/dev/shm/agent_ws_gate
mkdir -p "$AW_ROOT"
while read -r task; do
  [ -z "$task" ] && continue
  T="$SRC/tasks/finalpool/$task"
  if [ ! -f "$T/preprocess/main.py" ]; then
    printf "%s\tNOPRE\n" "$task" >> "$XC/pp_summary.tsv"; PASSN=$((PASSN+1)); continue
  fi
  AW="$AW_ROOT/$task"; rm -rf "$AW"; mkdir -p "$AW"
  # seed initial_workspace if present (agent runtime copies it first)
  if [ -d "$T/initial_workspace" ]; then cp -a "$T/initial_workspace/." "$AW/" 2>/dev/null || true; fi
  timeout 300 $PY "$T/preprocess/main.py" --agent_workspace "$AW" >"$XC/pp/$task.log" 2>&1
  rc=$?
  printf "%s\t%s\n" "$task" "$rc" >> "$XC/pp_summary.tsv"
  if [ $rc -eq 0 ]; then PASSN=$((PASSN+1)); else FAILN=$((FAILN+1)); echo "  [pp-FAIL rc=$rc] $task"; fi
done < "$XC/tasks.txt"

echo ""
echo "================ INFRA GATE SUMMARY ================"
echo "MCP list-tools rc=$MCP_RC"
echo "preprocess: PASS=$PASSN FAIL=$FAILN"
echo "===================================================="
[ $MCP_RC -eq 0 ] && [ $FAILN -eq 0 ]
' >"$LOGDIR/gate.log" 2>&1
GATE_RC=$?

# ── teardown host PG + collect artifacts ───────────────────────────────────
pg_ctl -D "$PGDATA" -m fast stop >/dev/null 2>&1 || true
cp "$GATE_XCHG/pp_summary.tsv" "$LOGDIR/" 2>/dev/null || true
cp "$GATE_XCHG/mcp_gate.log" "$LOGDIR/" 2>/dev/null || true
cp "$GATE_XCHG/pg.log" "$LOGDIR/" 2>/dev/null || true
cp -r "$GATE_XCHG/pp/." "$LOGDIR/pp/" 2>/dev/null || true
rm -rf "$GATE_XCHG" /dev/shm/infra_pg_$STAMP
enroot remove -f "$INST" >/dev/null 2>&1 || true

echo "[inner] gate rc=$GATE_RC  (log: $LOGDIR/gate.log)"
tail -30 "$LOGDIR/gate.log"
exit $GATE_RC
PAYLOAD_EOF

# inject the host-side task-list path (single-line value, safe to substitute
# into the payload file via sed — no heredoc re-expansion because we edit the
# file, not a shell variable passed through bash -c)
sed -i "s|__TASK_LIST_PATH__|$TASK_LIST|g" "$PAYLOAD" 2>/dev/null || true
# TASK_LIST is referenced as $TASK_LIST inside payload; export it for the
# compute node instead (cleaner than sed). srun propagates exported vars.
export TASK_LIST SLURM_MEM SLURM_CPUS SLURM_TIME SLURM_PARTITION

echo "[infra-gate] dispatching payload $PAYLOAD to slurm..."
srun -p "$SLURM_PARTITION" \
     -N1 -n1 -c"$SLURM_CPUS" \
     --mem="$SLURM_MEM" --time="$SLURM_TIME" \
     --job-name="infra-gate" \
     --export=ALL \
     bash "$PAYLOAD"
RC=$?
rm -f "$PAYLOAD"
exit $RC
