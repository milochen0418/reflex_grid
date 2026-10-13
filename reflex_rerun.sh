#!/bin/bash

echo "Checking for processes on ports 3000 and 8000..."

# Find PIDs using ports 3000 and 8000
PIDS=$(lsof -ti:3000,8000)

if [ -n "$PIDS" ]; then
  echo "Killing processes: $PIDS"
  kill -9 $PIDS
  echo "Processes killed."
else
  echo "No processes found on ports 3000 or 8000."
fi

# Test artifacts and docs must not trigger a backend hot reload (it would reset sessions mid-test).
export REFLEX_HOT_RELOAD_EXCLUDE_PATHS="${REFLEX_HOT_RELOAD_EXCLUDE_PATHS:-testcases:tests:docs:scripts}"

echo "Starting Reflex app with args: $@"

# If this script is executed via `poetry run`, avoid nesting Poetry.
if [ -n "${POETRY_ACTIVE:-}" ]; then
  reflex run "$@"
else
  poetry run reflex run "$@"
fi
