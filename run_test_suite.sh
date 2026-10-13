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

# Prefer IPv4 loopback to avoid IPv6/localhost quirks.
BASE_URL="${BASE_URL:-http://127.0.0.1:3000}"
export BASE_URL
export OUTPUT_DIR

REFLEX_LOG="$OUTPUT_DIR/reflex.log"
TEST_LOG="$OUTPUT_DIR/test.log"

run_in_poetry_env() {
  if [[ -n "${POETRY_ACTIVE:-}" ]]; then
    "$@"
  else
    poetry run "$@"
  fi
}

cleanup() {
  set +e
  if [[ -n "${SERVER_PID:-}" ]]; then
    kill "$SERVER_PID" >/dev/null 2>&1 || true
    wait "$SERVER_PID" >/dev/null 2>&1 || true
  fi
  # Also ensure ports are freed.
  lsof -ti:3000,8000 2>/dev/null | xargs -r kill -9 >/dev/null 2>&1 || true
}
trap cleanup EXIT

echo "[runner] suite=$SUITE_NAME"
echo "[runner] base_url=$BASE_URL"
echo "[runner] output_dir=$OUTPUT_DIR"

# Free the ports first, so the readiness check below can't see a server that is about to be
# replaced.
lsof -ti:3000,8000 2>/dev/null | xargs -r kill -9 >/dev/null 2>&1 || true
for i in $(seq 1 20); do
  if ! lsof -ti:3000,8000 >/dev/null 2>&1; then break; fi
  sleep 0.5
done

echo "[runner] starting server via ./reflex_rerun.sh ..."
(
  cd "$ROOT_DIR"
  # Ensure we don't double-wrap poetry: reflex_rerun.sh already handles POETRY_ACTIVE.
  ./reflex_rerun.sh
) >"$REFLEX_LOG" 2>&1 &
SERVER_PID=$!

echo "[runner] waiting for server to be reachable..."
READY=0
for i in $(seq 1 240); do
  code="$(curl -s -o /dev/null -w "%{http_code}" "$BASE_URL" || true)"
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
