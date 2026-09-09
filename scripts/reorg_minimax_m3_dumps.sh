#!/bin/bash
# Rehome MiniMax-M3 dumps into dumps/minimax-m3/{single_agent,auto_subagent}
# with the 167-case split. auto_subagent prefers three-apifail, then par47.
set -euo pipefail

DUMPS="/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps"
DEST="${DUMPS}/minimax-m3"
SPLIT="/lintaoLab2/bowending/project_agent_swarm_benchmark/dev_docs/split_167cases"
SRC_SA="${DUMPS}/kimi-code_MiniMax-M3-sa167"
SRC_AF="${DUMPS}/kimi-code_MiniMax-M3-three-apifail"
SRC_PAR="${DUMPS}/kimi-code_MiniMax-M3-par47-noeffort"
LOG="${DEST}/reorg_$(date +%Y%m%d-%H%M%S).log"

mkdir -p \
  "${DEST}/single_agent/no_subagent" \
  "${DEST}/single_agent/need_subagent/need_subagent_parallel" \
  "${DEST}/single_agent/need_subagent/need_subagent_no_parallel" \
  "${DEST}/single_agent/_unlisted" \
  "${DEST}/auto_subagent/no_subagent" \
  "${DEST}/auto_subagent/need_subagent/need_subagent_parallel" \
  "${DEST}/auto_subagent/need_subagent/need_subagent_no_parallel" \
  "${DEST}/auto_subagent/_unlisted"
exec > >(tee -a "$LOG") 2>&1

echo "=== reorg start $(date) host=$(hostname) nproc=$(nproc) ==="
for s in "$SRC_SA" "$SRC_AF" "$SRC_PAR"; do
  if [[ ! -d "$s" ]]; then
    echo "FATAL: source missing: $s" >&2
    exit 1
  fi
  echo "[src] $s"
done

rsync_tree() {
  local tag="$1" src="$2" dst="$3"
  mkdir -p "$dst"
  echo "[rsync] START ${tag} -> ${dst} $(date)"
  rsync -aW --no-compress --human-readable --info=stats2 "${src}/" "${dst}/"
  echo "[rsync] DONE  ${tag} $(date)"
}

rsync_tree sa167 "$SRC_SA" "${DEST}/single_agent/_incoming" &
pid_sa=$!
rsync_tree three_apifail "$SRC_AF" "${DEST}/auto_subagent/_incoming" &
pid_af=$!
wait "$pid_sa"
wait "$pid_af"

echo "=== fill auto_subagent gaps from par47 $(date) ==="
python3 - "$DEST" "$SRC_PAR" <<'PY'
from pathlib import Path
import shutil
import subprocess
import sys

dest = Path(sys.argv[1])
par = Path(sys.argv[2])
incoming = dest / "auto_subagent" / "_incoming"
skip = {
    "no_subagent", "need_subagent", "_unlisted", "_incoming",
}
have = {p.name for p in incoming.iterdir() if p.is_dir() and p.name not in skip}
added = []
for child in sorted(par.iterdir()):
    if not child.is_dir() or child.name in skip:
        continue
    if child.name in have:
        continue
    target = incoming / child.name
    print(f"[gap] rsync par47 {child.name}")
    subprocess.check_call(
        ["rsync", "-aW", "--no-compress", f"{child}/", f"{target}/"]
    )
    added.append(child.name)
# keep par47 root files if names don't collide
for child in sorted(par.iterdir()):
    if child.is_dir():
        continue
    target = incoming / child.name
    if target.exists():
        renamed = incoming / f"par47__{child.name}"
        shutil.copy2(child, renamed)
        print(f"[gap] keep root {child.name} -> {renamed.name}")
    else:
        shutil.copy2(child, target)
        print(f"[gap] keep root {child.name}")
print(f"[gap] added_from_par47={len(added)} already_in_apifail={len(have)}")
if added:
    print("[gap] tasks:", ", ".join(added))
PY

echo "=== split $(date) ==="
python3 - "$DEST" "$SPLIT" <<'PY'
from pathlib import Path
import shutil
import sys

dest = Path(sys.argv[1])
split = Path(sys.argv[2])

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

for mode in ("single_agent", "auto_subagent"):
    incoming = dest / mode / "_incoming"
    root = dest / mode
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
        raise SystemExit(f"{mode} incoming not empty: {leftover}")
    incoming.rmdir()
    listed = (
        counts["no_subagent"]
        + counts["need_subagent/need_subagent_parallel"]
        + counts["need_subagent/need_subagent_no_parallel"]
    )
    print(f"[split] {mode} {counts} listed={listed}")
PY

echo "=== verify $(date) ==="
python3 - "$DEST" <<'PY'
from pathlib import Path
import sys
dest = Path(sys.argv[1])

def ndirs(p):
    return len([x for x in p.iterdir() if x.is_dir()]) if p.is_dir() else 0

ok = True
for mode in ("single_agent", "auto_subagent"):
    root = dest / mode
    n_nos = ndirs(root / "no_subagent")
    n_par = ndirs(root / "need_subagent/need_subagent_parallel")
    n_ser = ndirs(root / "need_subagent/need_subagent_no_parallel")
    n_unl = ndirs(root / "_unlisted")
    print(f"[verify] {mode}: no_subagent={n_nos} parallel={n_par} serial={n_ser} unlisted={n_unl} listed={n_nos+n_par+n_ser}")
    if n_nos + n_par + n_ser == 0:
        ok = False
if not ok:
    raise SystemExit("no listed tasks")
print("[verify] ok")
PY

echo "[rm] $SRC_SA"
rm -rf --one-file-system "$SRC_SA"
echo "[rm] $SRC_AF"
rm -rf --one-file-system "$SRC_AF"
echo "[rm] $SRC_PAR"
rm -rf --one-file-system "$SRC_PAR"

echo "=== dest tree (depth 3, first 80) ==="
find "$DEST" -mindepth 1 -maxdepth 3 \( -type d -o -name '*.csv' -o -name '*.log' \) | sort | head -n 80
echo "=== reorg done $(date) ==="
