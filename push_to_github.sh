#!/bin/bash
# Push minhuizhou repo to GitHub (xionghul/minhuizhou)
set -euo pipefail

REPO_URL="https://github.com/xionghul/minhuizhou.git"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

if git remote get-url origin >/dev/null 2>&1; then
  :
else
  TOKEN=""
  if [ -d /workspace/.git ]; then
    TOKEN=$(git -C /workspace remote get-url origin | sed -n 's/.*x-access-token:\([^@]*\)@.*/\1/p')
  fi
  if [ -n "$TOKEN" ]; then
    git remote add origin "https://x-access-token:${TOKEN}@github.com/xionghul/minhuizhou.git"
  else
    git remote add origin "$REPO_URL"
  fi
fi

echo "Checking remote repository..."
if git ls-remote origin >/dev/null 2>&1; then
  git push -u origin main
  echo "Done: $REPO_URL"
else
  cat <<EOF
Remote repository not found or not accessible.

Please create an empty repository on GitHub first:
  1. Open https://github.com/new
  2. Repository name: minhuizhou
  3. Do NOT initialize with README
  4. Run this script again: ./push_to_github.sh

After creation, grant Cursor GitHub App access to the new repo
if you want Cloud Agents to work on it.
EOF
  exit 1
fi
