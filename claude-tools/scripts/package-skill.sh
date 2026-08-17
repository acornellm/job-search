#!/usr/bin/env bash
# Zips a skill folder for upload to the Claude desktop app or the Skills API.
# Usage: scripts/package-skill.sh plugins/writing/skills/resume-cover-letter
set -euo pipefail

if [ $# -ne 1 ]; then
  echo "Usage: $0 <path-to-skill-folder>" >&2
  exit 1
fi

SKILL_DIR="$1"
SKILL_DIR="${SKILL_DIR%/}"

if [ ! -f "$SKILL_DIR/SKILL.md" ]; then
  echo "No SKILL.md found at $SKILL_DIR/SKILL.md" >&2
  exit 1
fi

SKILL_NAME="$(basename "$SKILL_DIR")"
REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
DIST_DIR="$REPO_ROOT/dist"
OUT_ZIP="$DIST_DIR/$SKILL_NAME.zip"

mkdir -p "$DIST_DIR"
rm -f "$OUT_ZIP"

# Zip so SKILL.md sits at the archive root (required by both the desktop
# app and the Skills API) rather than nested under the skill folder name.
(cd "$SKILL_DIR" && zip -r -X "$OUT_ZIP" . -x '.*')

echo "Wrote $OUT_ZIP"
