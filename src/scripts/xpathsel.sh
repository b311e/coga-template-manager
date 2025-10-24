#!/usr/bin/env bash
set -euo pipefail

xml=${1:?usage: xpathsel.sh <xml-file> [xpath ...]}
shift || true
ns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"

run() {
  local xp=$1
  echo "=== $xp ==="
  xmlstarlet sel -N w="$ns" -t -c "$xp" "$xml" || echo "not found"
}

if [ "$#" -gt 0 ]; then
  for xp in "$@"; do run "$xp"; done
else
  while IFS= read -r xp; do
    [[ -z "$xp" || "$xp" =~ ^# ]] && continue
    run "$xp"
  done
fi
