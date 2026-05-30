#!/usr/bin/env bash
# Nightly SQLite backup → private GitHub repo.
# Run from Render Cron Job (free tier) once per day.
#
# Required env vars:
#   DATABASE_PATH       — path to live SQLite file (set by Render)
#   GH_BACKUP_REPO      — e.g. "yogendradas017/tms-backups" (private)
#   GH_BACKUP_TOKEN     — fine-grained PAT with contents:write on that repo
#   GH_BACKUP_BRANCH    — default "main"
#
# What it does:
#   1. sqlite3 .backup to a tmp file (online-safe even with WAL writers)
#   2. gzip
#   3. PUT to the repo at backups/YYYY-MM-DD.db.gz
#
# Restore: download the .db.gz, gunzip, replace DATABASE_PATH file.

set -euo pipefail

DB_SRC="${DATABASE_PATH:-data/training.db}"
REPO="${GH_BACKUP_REPO:?GH_BACKUP_REPO required}"
TOKEN="${GH_BACKUP_TOKEN:?GH_BACKUP_TOKEN required}"
BRANCH="${GH_BACKUP_BRANCH:-main}"

STAMP=$(date -u +%Y-%m-%d)
TMP=$(mktemp -d)
SNAP="$TMP/snap.db"
GZ="$TMP/$STAMP.db.gz"

# 1. Online-safe snapshot (works while gunicorn holds connection)
sqlite3 "$DB_SRC" ".backup '$SNAP'"

# 2. Compress
gzip -9 -c "$SNAP" > "$GZ"

# 3. PUT via GitHub Contents API (base64-encoded)
B64=$(base64 -w 0 "$GZ")
PATH_IN_REPO="backups/$STAMP.db.gz"

# Check if file already exists (need SHA for update)
SHA=$(curl -s -H "Authorization: token $TOKEN" \
  "https://api.github.com/repos/$REPO/contents/$PATH_IN_REPO?ref=$BRANCH" \
  | grep -oP '"sha":\s*"\K[^"]+' | head -n1 || true)

PAYLOAD=$(cat <<EOF
{"message":"backup $STAMP","branch":"$BRANCH","content":"$B64"$( [ -n "$SHA" ] && echo ",\"sha\":\"$SHA\"" )}
EOF
)

curl -s -X PUT \
  -H "Authorization: token $TOKEN" \
  -H "Accept: application/vnd.github+json" \
  "https://api.github.com/repos/$REPO/contents/$PATH_IN_REPO" \
  -d "$PAYLOAD" | grep -E '"(name|message)"' | head -5

rm -rf "$TMP"
echo "Backup $STAMP uploaded to $REPO:$PATH_IN_REPO"
