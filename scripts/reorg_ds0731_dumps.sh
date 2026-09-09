#!/bin/bash
# Rehome DeepSeek-V4-Flash-0731 dumps into mode x annotation-split folders.
# Run on a compute node (gnho019 / gnho014). Same-NFS rsync + rename.
set -euo pipefail

DUMPS="/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps"
DEST="${DUMPS}/deepseek-v4-flash-0731"
SPLIT="/lintaoLab2/bowending/project_agent_swarm_benchmark/dev_docs/split_167cases"
LOG="${DEST}/reorg_$(date +%Y%m%d-%H%M%S).log"
mkdir -p "$DEST"
exec > >(tee -a "$LOG") 2>&1

echo "=== reorg start $(date) host=$(hostname) nproc=$(nproc) ==="

declare -A SRC=(
  [single_agent]="${DUMPS}/kimi-code_deepseek-v4-flash-linslab-single-agent-full_20260817-002741"
  [auto_subagent]="${DUMPS}/kimi-code_deepseek-v4-flash-subagent"
  [plan1st_subagent]="${DUMPS}/kimi-code_deepseek-v4-flash-plan-first"
)

for mode in single_agent auto_subagent plan1st_subagent fg_subagent oracle_subagent; do
  mkdir -p \
    "${DEST}/${mode}/no_subagent" \
    "${DEST}/${mode}/need_subagent/need_subagent_parallel" \
    "${DEST}/${mode}/need_subagent/need_subagent_no_parallel" \
    "${DEST}/${mode}/_unlisted"
done

for mode in "${!SRC[@]}"; do
  s="${SRC[$mode]}"
  if [[ ! -d "$s" ]]; then
    echo "FATAL: source missing: $s" >&2
    exit 1
  fi
  echo "[src] $mode <- $s"
done

# Parallel whole-file rsync of each mode tree into DEST/<mode>/_incoming
# so we never mix split folders with still-arriving files.
rsync_one() {
  local mode="$1" src="$2"
  local incoming="${DEST}/${mode}/_incoming"
  mkdir -p "$incoming"
  echo "[rsync] START $mode $(date)"
  rsync -aW --no-compress --human-readable --info=stats2 \
    "${src}/" "${incoming}/"
  echo "[rsync] DONE  $mode $(date)"
}

export -f rsync_one
export DEST
# bash parallel: three modes
pids=()
for mode in single_agent auto_subagent plan1st_subagent; do
  rsync_one "$mode" "${SRC[$mode]}" &
  pids+=($!)
done
fail=0
for pid in "${pids[@]}"; do
  if ! wait "$pid"; then
    echo "FATAL: rsync pid $pid failed" >&2
    fail=1
  fi
done
if [[ "$fail" != 0 ]]; then
  exit 1
fi

echo "=== rsync complete, splitting task dirs $(date) ==="

python3 - "$DEST" "$SPLIT" <<'PY'
import os, sys, shutil
from pathlib import Path

dest = Path(sys.argv[1])
split = Path(sys.argv[2])

def load(name):
    rows = [ln.strip() for ln in (split / name).read_text().splitlines() if ln.strip()]
    return rows, set(rows)

nos_l, nos = load("no_subagent.txt")
par_l, par = load("need_subagent_parallel.txt")
ser_l, ser = load("need_subagent_no_parallel.txt")
assert len(nos) == 33, len(nos)
assert len(par) == 131, len(par)
assert len(ser) == 3, len(ser)
assert not (nos & par or nos & ser or par & ser)

modes = ["single_agent", "auto_subagent", "plan1st_subagent"]
meta_dirs = {
    "no_subagent", "need_subagent", "_unlisted", "_incoming",
}

def bucket(name: str):
    if name in nos:
        return "no_subagent"
    if name in par:
        return "need_subagent/need_subagent_parallel"
    if name in ser:
        return "need_subagent/need_subagent_no_parallel"
    return "_unlisted"

for mode in modes:
    incoming = dest / mode / "_incoming"
    if not incoming.is_dir():
        raise SystemExit(f"missing incoming {incoming}")
    counts = {"no_subagent": 0, "need_subagent/need_subagent_parallel": 0,
              "need_subagent/need_subagent_no_parallel": 0, "_unlisted": 0}
    for child in sorted(incoming.iterdir()):
        if child.is_dir() and child.name not in meta_dirs:
            rel = bucket(child.name)
            target = dest / mode / rel / child.name
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.exists():
                raise SystemExit(f"target exists: {target}")
            shutil.move(str(child), str(target))
            counts[rel] += 1
        else:
            # root files / leftover dirs stay at mode root
            target = dest / mode / child.name
            if target.exists():
                raise SystemExit(f"target exists: {target}")
            shutil.move(str(child), str(target))
    leftover = [p.name for p in incoming.iterdir()]
    if leftover:
        raise SystemExit(f"{mode} incoming not empty: {leftover}")
    incoming.rmdir()
    print(f"[split] {mode} {counts}")

print("[split] done")
PY

echo "=== verify then remove sources $(date) ==="

python3 - "$DEST" <<'PY'
import os, sys
from pathlib import Path
dest = Path(sys.argv[1])
ok = True
for mode in ["single_agent", "auto_subagent", "plan1st_subagent"]:
    root = dest / mode
    n_nos = len([p for p in (root / "no_subagent").iterdir() if p.is_dir()]) if (root / "no_subagent").is_dir() else 0
    n_par = len([p for p in (root / "need_subagent/need_subagent_parallel").iterdir() if p.is_dir()]) if (root / "need_subagent/need_subagent_parallel").is_dir() else 0
    n_ser = len([p for p in (root / "need_subagent/need_subagent_no_parallel").iterdir() if p.is_dir()]) if (root / "need_subagent/need_subagent_no_parallel").is_dir() else 0
    n_unl = len([p for p in (root / "_unlisted").iterdir() if p.is_dir()]) if (root / "_unlisted").is_dir() else 0
    print(f"[verify] {mode}: no_subagent={n_nos} parallel={n_par} serial={n_ser} unlisted={n_unl} total_listed={n_nos+n_par+n_ser}")
    if n_nos + n_par + n_ser == 0:
        ok = False
        print(f"[verify] FAIL {mode} has no listed tasks")
if not ok:
    raise SystemExit(2)
print("[verify] ok")
PY

for mode in single_agent auto_subagent plan1st_subagent; do
  src="${SRC[$mode]}"
  echo "[rm] $src"
  rm -rf --one-file-system "$src"
done

echo "=== dest tree (depth 3) ==="
find "$DEST" -mindepth 1 -maxdepth 3 \( -type d -o -name '*.csv' -o -name '*.log' \) | sort | head -n 200
echo "=== reorg done $(date) ==="
