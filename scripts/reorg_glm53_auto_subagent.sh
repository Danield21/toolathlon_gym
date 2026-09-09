#!/bin/bash
# Move glm-5-3 par47 dump into dumps/glm-5.3/auto-subagent with the 167-case split.
set -euo pipefail

DUMPS="/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps"
DEST="${DUMPS}/glm-5.3"
MODE="auto-subagent"
SPLIT="/lintaoLab2/bowending/project_agent_swarm_benchmark/dev_docs/split_167cases"
SRC="${DUMPS}/kimi-code_glm-5-3-par47-noeffort"
LOG="${DEST}/reorg_auto-subagent_$(date +%Y%m%d-%H%M%S).log"

mkdir -p \
  "${DEST}/${MODE}/no_subagent" \
  "${DEST}/${MODE}/need_subagent/need_subagent_parallel" \
  "${DEST}/${MODE}/need_subagent/need_subagent_no_parallel" \
  "${DEST}/${MODE}/_unlisted"
exec > >(tee -a "$LOG") 2>&1

echo "=== reorg start $(date) host=$(hostname) nproc=$(nproc) ==="
if [[ ! -d "$SRC" ]]; then
  echo "FATAL: source missing: $SRC" >&2
  exit 1
fi
echo "[src] $SRC"

incoming="${DEST}/${MODE}/_incoming"
mkdir -p "$incoming"
echo "[rsync] START $(date)"
rsync -aW --no-compress --human-readable --info=stats2 "${SRC}/" "${incoming}/"
echo "[rsync] DONE  $(date)"

echo "=== split $(date) ==="
python3 - "$DEST" "$MODE" "$SPLIT" <<'PY'
from pathlib import Path
import shutil
import sys

root = Path(sys.argv[1]) / sys.argv[2]
split = Path(sys.argv[3])
incoming = root / "_incoming"

def load(name):
    return {ln.strip() for ln in (split / name).read_text().splitlines() if ln.strip()}

nos, par, ser = load("no_subagent.txt"), load("need_subagent_parallel.txt"), load("need_subagent_no_parallel.txt")
assert len(nos) == 33 and len(par) == 131 and len(ser) == 3
assert not (nos & par or nos & ser or par & ser)
skip = {"no_subagent", "need_subagent", "_unlisted", "_incoming"}

def bucket(name: str) -> str:
    if name in nos:
        return "no_subagent"
    if name in par:
        return "need_subagent/need_subagent_parallel"
    if name in ser:
        return "need_subagent/need_subagent_no_parallel"
    return "_unlisted"

counts = {
    "no_subagent": 0,
    "need_subagent/need_subagent_parallel": 0,
    "need_subagent/need_subagent_no_parallel": 0,
    "_unlisted": 0,
}
for child in sorted(incoming.iterdir()):
    if child.is_dir() and child.name not in skip:
        rel = bucket(child.name)
        target = root / rel / child.name
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists():
            raise SystemExit(f"target exists: {target}")
        shutil.move(str(child), str(target))
        counts[rel] += 1
    else:
        target = root / child.name
        if target.exists():
            raise SystemExit(f"target exists: {target}")
        shutil.move(str(child), str(target))

leftover = [p.name for p in incoming.iterdir()]
if leftover:
    raise SystemExit(f"incoming not empty: {leftover}")
incoming.rmdir()
listed = counts["no_subagent"] + counts["need_subagent/need_subagent_parallel"] + counts["need_subagent/need_subagent_no_parallel"]
print(f"[split] {counts} listed={listed}")
PY

echo "=== verify $(date) ==="
python3 - "$DEST" "$MODE" <<'PY'
from pathlib import Path
import sys
root = Path(sys.argv[1]) / sys.argv[2]

def ndirs(p):
    return len([x for x in p.iterdir() if x.is_dir()]) if p.is_dir() else 0

n_nos = ndirs(root / "no_subagent")
n_par = ndirs(root / "need_subagent/need_subagent_parallel")
n_ser = ndirs(root / "need_subagent/need_subagent_no_parallel")
n_unl = ndirs(root / "_unlisted")
print(f"[verify] no_subagent={n_nos} parallel={n_par} serial={n_ser} unlisted={n_unl} listed={n_nos+n_par+n_ser}")
if n_nos + n_par + n_ser == 0:
    raise SystemExit("no listed tasks")
print("[verify] ok")
PY

echo "[rm] $SRC"
rm -rf --one-file-system "$SRC"

echo "=== dest tree (depth 3) ==="
find "${DEST}/${MODE}" -mindepth 1 -maxdepth 3 \( -type d -o -name '*.csv' -o -name '*.log' \) | sort | head -n 120
echo "=== reorg done $(date) ==="
