#!/usr/bin/env bash
set -euo pipefail

SUITE_NAME="${1:-}"
if [[ -z "$SUITE_NAME" ]]; then
  echo "Usage: $0 <suite_name>" >&2
  echo "Example: $0 smoke_home" >&2
  exit 2
fi

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SUITE_DIR="$ROOT_DIR/testcases/$SUITE_NAME"
TEST_SCRIPT="$SUITE_DIR/run_test.py"
OUTPUT_DIR="$SUITE_DIR/output"

if [[ ! -f "$TEST_SCRIPT" ]]; then
  echo "Test suite not found: $TEST_SCRIPT" >&2
  exit 2
fi

mkdir -p "$OUTPUT_DIR"

# The suite's server gets ports of its own, never 3000/8000: those are for the app you run yourself
# (reflex_rerun.sh) and for other local projects, and this runner frees its ports before and after.
TEST_FRONTEND_PORT="${TEST_FRONTEND_PORT:-3611}"
TEST_BACKEND_PORT="${TEST_BACKEND_PORT:-8611}"
# Prefer IPv4 loopback to avoid IPv6/localhost quirks.
BASE_URL="${BASE_URL:-http://127.0.0.1:$TEST_FRONTEND_PORT}"
# The page's websocket goes to the test backend (rxconfig.py reads API_URL).
API_URL="http://localhost:$TEST_BACKEND_PORT"
export BASE_URL API_URL
export OUTPUT_DIR

REFLEX_LOG="$OUTPUT_DIR/reflex.log"
TEST_LOG="$OUTPUT_DIR/test.log"

# The workbooks a suite creates go to its own output folder (fresh each run), not to yours.
if [[ -z "${REFLEX_GRID_HOME:-}" ]]; then
  export REFLEX_GRID_HOME="$OUTPUT_DIR/data"
  rm -rf "$REFLEX_GRID_HOME"
fi

run_in_poetry_env() {
  if [[ -n "${POETRY_ACTIVE:-}" ]]; then
    "$@"
  else
    poetry run "$@"
  fi
}

free_test_ports() {
  local pids
  pids="$(lsof -ti:"$TEST_FRONTEND_PORT","$TEST_BACKEND_PORT" 2>/dev/null || true)"
  if [[ -n "$pids" ]]; then
    kill -9 $pids >/dev/null 2>&1 || true
  fi
}

cleanup() {
  set +e
  if [[ -n "${SERVER_PID:-}" ]]; then
    kill "$SERVER_PID" >/dev/null 2>&1 || true
    wait "$SERVER_PID" >/dev/null 2>&1 || true
  fi
  # Also ensure the test ports are freed.
  free_test_ports
}
trap cleanup EXIT

echo "[runner] suite=$SUITE_NAME"
echo "[runner] base_url=$BASE_URL (backend port $TEST_BACKEND_PORT)"
echo "[runner] output_dir=$OUTPUT_DIR"

# Free the test ports first, so the readiness check below can't see a server that is about to be
# replaced.
free_test_ports
for i in $(seq 1 20); do
  if ! lsof -ti:"$TEST_FRONTEND_PORT","$TEST_BACKEND_PORT" >/dev/null 2>&1; then break; fi
  sleep 0.5
done

echo "[runner] starting server on ports $TEST_FRONTEND_PORT/$TEST_BACKEND_PORT ..."
(
  cd "$ROOT_DIR"
  # rxconfig.py keeps test artifacts and docs from hot-reloading it (that would reset sessions mid-test).
  run_in_poetry_env reflex run --frontend-port "$TEST_FRONTEND_PORT" --backend-port "$TEST_BACKEND_PORT"
) >"$REFLEX_LOG" 2>&1 &
SERVER_PID=$!

echo "[runner] waiting for server to be reachable..."
READY=0
for i in $(seq 1 240); do
  code="$(curl -s -m 5 -o /dev/null -w "%{http_code}" "$BASE_URL" || true)"
  if [[ "$code" =~ ^2|^3 ]]; then
    READY=1
    break
  fi
  sleep 1
  # Surface a tiny bit of log if it looks stuck.
  if [[ $i -eq 30 || $i -eq 90 || $i -eq 150 || $i -eq 210 ]]; then
    echo "[runner] still waiting (attempt $i/240), last_http=$code" >&2
  fi
  if ! kill -0 "$SERVER_PID" >/dev/null 2>&1; then
    echo "[runner] server process exited early; see $REFLEX_LOG" >&2
    exit 1
  fi
done

if [[ "$READY" -ne 1 ]]; then
  echo "[runner] server not ready after timeout; see $REFLEX_LOG" >&2
  exit 1
fi

echo "[runner] running test: $TEST_SCRIPT"
set +e
run_in_poetry_env python "$TEST_SCRIPT" >"$TEST_LOG" 2>&1
TEST_EXIT=$?
set -e

if [[ $TEST_EXIT -ne 0 ]]; then
  echo "[runner] FAIL (exit=$TEST_EXIT). Artifacts in: $OUTPUT_DIR" >&2
  echo "[runner] tail test.log:" >&2
  tail -n 80 "$TEST_LOG" >&2 || true
  exit "$TEST_EXIT"
fi

echo "[runner] PASS. Artifacts in: $OUTPUT_DIR"
