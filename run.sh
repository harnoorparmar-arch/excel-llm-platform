#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"

# Restrict reload to app code only; --reload-exclude avoids churn from .venv,
# node_modules, etc. (requires watchfiles; uvicorn[standard] includes it)
exec uvicorn api.main:app \
  --reload \
  --reload-dir "$ROOT/api" \
  --reload-dir "$ROOT/parser" \
  --reload-dir "$ROOT/frontend" \
  --reload-dir "$ROOT/storage" \
  --reload-exclude ".venv" \
  --reload-exclude "**/.venv/**" \
  --reload-exclude "node_modules" \
  --reload-exclude "**/node_modules/**" \
  --reload-exclude "__pycache__" \
  --reload-exclude "**/__pycache__/**" \
  --reload-exclude "test_data" \
  --reload-exclude "**/test_data/**" \
  --reload-exclude "*.db" \
  --host 0.0.0.0 \
  --port 8000 \
  "$@"
