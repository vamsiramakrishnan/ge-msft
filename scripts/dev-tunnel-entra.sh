#!/usr/bin/env bash
# Restart the web-shell dev server, create a fresh Cloudflare quick tunnel,
# regenerate development manifests, and sync the Entra SPA redirect URI.

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
WEB_SHELL_ENV="${GE_WEB_SHELL_ENV:-${ROOT_DIR}/packages/web-shell/.env}"
STATE_DIR="${GE_DEV_STATE_DIR:-${ROOT_DIR}/.ge-dev}"
LOG_DIR="${STATE_DIR}/logs"
VITE_PID_FILE="${STATE_DIR}/vite.pid"
CLOUDFLARED_PID_FILE="${STATE_DIR}/cloudflared.pid"
VITE_LOG="${LOG_DIR}/vite.log"
CLOUDFLARED_LOG="${LOG_DIR}/cloudflared.log"

usage() {
  cat <<'EOF'
Usage:
  scripts/dev-tunnel-entra.sh [options]

Options:
  --port PORT              Dev server port. Default: GE_DEV_PORT from env/.env, else 13000.
  --target URL             Local tunnel target. Default: detected http(s)://localhost:PORT.
  --skip-entra             Do not patch the Entra app registration redirect URI.
  --skip-manifests         Do not regenerate/validate/package development manifests.
  --keep-stale-redirects   Keep existing trycloudflare auth redirects instead of replacing them.
  -h, --help               Show this help.

Environment:
  GE_WEB_SHELL_ENV         .env file to update. Default: packages/web-shell/.env.
  CLOUDFLARED_BIN          cloudflared binary. Default: /tmp/cloudflared, then PATH.
  AZ_BIN                   Azure CLI binary. Default: ./bin/az, .venv-az/bin/az, then PATH az.
  GE_DEV_STATE_DIR         PID/log directory. Default: .ge-dev.

What it does:
  1. Stops the prior pid-file managed Vite and cloudflared processes.
  2. Starts Vite on the requested port.
  3. Starts a Cloudflare quick tunnel and captures the generated origin.
  4. Writes GE_DEV_WEB_ORIGIN/GE_DEV_WEB_DOMAIN/GE_DEV_PORT to packages/web-shell/.env.
  5. Restarts Vite so Vite allowedHosts includes the new tunnel host.
  6. Regenerates dev manifests/packages.
  7. Replaces stale trycloudflare auth redirects in Entra with the new redirect URI.
EOF
}

read_env_file_value() {
  local key="$1"
  [[ -f "${WEB_SHELL_ENV}" ]] || return 0
  awk -F= -v key="$key" '
    $1 == key || $1 == "export " key {
      value=$0
      sub("^[^=]*=", "", value)
      gsub(/^["'\'' ]+|["'\'' ]+$/, "", value)
      print value
    }
  ' "${WEB_SHELL_ENV}" | tail -n 1
}

env_or_file() {
  local key="$1"
  local fallback="${2:-}"
  local value="${!key:-}"
  if [[ -z "${value}" ]]; then
    value="$(read_env_file_value "${key}")"
  fi
  if [[ -z "${value}" ]]; then
    value="${fallback}"
  fi
  printf '%s' "${value}"
}

PORT="$(env_or_file GE_DEV_PORT 13000)"
LOCAL_TARGET=""
SKIP_ENTRA=0
SKIP_MANIFESTS=0
KEEP_STALE_REDIRECTS=0

while [[ $# -gt 0 ]]; do
  case "$1" in
    --port)
      PORT="${2:-}"
      shift 2
      ;;
    --target)
      LOCAL_TARGET="${2:-}"
      shift 2
      ;;
    --skip-entra)
      SKIP_ENTRA=1
      shift
      ;;
    --skip-manifests)
      SKIP_MANIFESTS=1
      shift
      ;;
    --keep-stale-redirects)
      KEEP_STALE_REDIRECTS=1
      shift
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown option: $1" >&2
      usage >&2
      exit 2
      ;;
  esac
done

mkdir -p "${LOG_DIR}"

section() {
  printf '\n==> %s\n' "$1"
}

stop_pid_file() {
  local label="$1"
  local pid_file="$2"
  [[ -f "${pid_file}" ]] || return 0
  local pid
  pid="$(cat "${pid_file}" 2>/dev/null || true)"
  if [[ -n "${pid}" ]] && kill -0 "${pid}" 2>/dev/null; then
    echo "Stopping ${label} pid ${pid}"
    kill "${pid}" 2>/dev/null || true
    for _ in 1 2 3 4 5; do
      kill -0 "${pid}" 2>/dev/null || break
      sleep 1
    done
    if kill -0 "${pid}" 2>/dev/null; then
      kill -9 "${pid}" 2>/dev/null || true
    fi
  fi
  rm -f "${pid_file}"
}

find_cloudflared() {
  if [[ -n "${CLOUDFLARED_BIN:-}" && -x "${CLOUDFLARED_BIN}" ]]; then
    printf '%s\n' "${CLOUDFLARED_BIN}"
  elif [[ -x /tmp/cloudflared ]]; then
    printf '%s\n' /tmp/cloudflared
  elif command -v cloudflared >/dev/null 2>&1; then
    command -v cloudflared
  else
    echo "cloudflared not found. Set CLOUDFLARED_BIN or install cloudflared." >&2
    exit 1
  fi
}

start_vite() {
  section "Start Vite dev server"
  : > "${VITE_LOG}"
  (
    cd "${ROOT_DIR}/packages/web-shell"
    GE_DEV_PORT="${PORT}" bun run dev -- --host 0.0.0.0 --port "${PORT}"
  ) >"${VITE_LOG}" 2>&1 &
  echo "$!" > "${VITE_PID_FILE}"
  echo "Vite pid: $(cat "${VITE_PID_FILE}")"
  echo "Vite log: ${VITE_LOG}"
}

detect_local_target() {
  if [[ -n "${LOCAL_TARGET}" ]]; then
    printf '%s\n' "${LOCAL_TARGET}"
    return 0
  fi

  local candidates=()
  if [[ "${GOOGLE_CLOUD_WORKSTATIONS:-}" == "true" ]]; then
    candidates+=("http://localhost:${PORT}" "https://localhost:${PORT}")
  else
    candidates+=("https://localhost:${PORT}" "http://localhost:${PORT}")
  fi

  local url
  for url in "${candidates[@]}"; do
    if curl -kfsS "${url}/taskpane.html" >/dev/null 2>&1; then
      printf '%s\n' "${url}"
      return 0
    fi
  done

  echo "Vite did not become reachable on ${candidates[*]}." >&2
  tail -80 "${VITE_LOG}" >&2 || true
  exit 1
}

wait_for_vite() {
  section "Wait for Vite"
  local target
  for _ in $(seq 1 45); do
    target="$(detect_local_target 2>/dev/null || true)"
    if [[ -n "${target}" ]]; then
      printf '%s\n' "${target}"
      return 0
    fi
    sleep 1
  done
  echo "Timed out waiting for Vite." >&2
  tail -120 "${VITE_LOG}" >&2 || true
  exit 1
}

start_cloudflared() {
  local target="$1"
  local cloudflared_bin="$2"
  section "Start Cloudflare quick tunnel"
  : > "${CLOUDFLARED_LOG}"
  local args=(tunnel --url "${target}")
  if [[ "${target}" == https://* ]]; then
    args+=(--no-tls-verify)
  fi
  "${cloudflared_bin}" "${args[@]}" >"${CLOUDFLARED_LOG}" 2>&1 &
  echo "$!" > "${CLOUDFLARED_PID_FILE}"
  echo "cloudflared pid: $(cat "${CLOUDFLARED_PID_FILE}")"
  echo "cloudflared log: ${CLOUDFLARED_LOG}"
}

wait_for_tunnel_origin() {
  section "Wait for Cloudflare origin"
  local origin=""
  for _ in $(seq 1 60); do
    origin="$(grep -Eo 'https://[A-Za-z0-9-]+\.trycloudflare\.com' "${CLOUDFLARED_LOG}" | tail -n 1 || true)"
    if [[ -n "${origin}" ]]; then
      printf '%s\n' "${origin}"
      return 0
    fi
    sleep 1
  done
  echo "Timed out waiting for Cloudflare tunnel URL." >&2
  tail -120 "${CLOUDFLARED_LOG}" >&2 || true
  exit 1
}

write_web_env() {
  local origin="$1"
  local domain="${origin#https://}"
  section "Update web-shell env"
  node - "${WEB_SHELL_ENV}" "${origin}" "${domain}" "${PORT}" <<'NODE'
const fs = require("fs");
const [path, origin, domain, port] = process.argv.slice(2);
const updates = new Map([
  ["GE_DEV_WEB_ORIGIN", origin],
  ["GE_DEV_WEB_DOMAIN", domain],
  ["GE_DEV_PORT", port],
]);
let lines = [];
if (fs.existsSync(path)) {
  lines = fs.readFileSync(path, "utf8").split(/\r?\n/);
}
const seen = new Set();
lines = lines.map((line) => {
  const match = line.match(/^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=/);
  if (!match || !updates.has(match[1])) return line;
  seen.add(match[1]);
  return `${match[1]}=${updates.get(match[1])}`;
});
for (const [key, value] of updates) {
  if (!seen.has(key)) lines.push(`${key}=${value}`);
}
fs.writeFileSync(path, lines.join("\n").replace(/\n*$/, "\n"));
NODE
  echo "Wrote ${WEB_SHELL_ENV}"
}

regenerate_manifests() {
  if [[ "${SKIP_MANIFESTS}" == "1" ]]; then
    section "Skip manifest generation"
    return 0
  fi

  section "Regenerate development manifests"
  (
    cd "${ROOT_DIR}"
    GE_DEV_PORT="${PORT}" bun run manifests:generate -- --profile development
    GE_DEV_PORT="${PORT}" bun run manifests:validate -- --profile development
    GE_DEV_PORT="${PORT}" bun run package:dev
  )
}

sync_entra_redirect() {
  if [[ "${SKIP_ENTRA}" == "1" ]]; then
    section "Skip Entra redirect sync"
    return 0
  fi

  local origin="$1"
  section "Sync Entra SPA redirect"
  local args=(tools/release/sync-entra-spa.mjs --profile development --origin "${origin}")
  if [[ "${KEEP_STALE_REDIRECTS}" == "1" ]]; then
    args+=(--keep-stale-redirects)
  fi
  (
    cd "${ROOT_DIR}"
    bun "${args[@]}"
  )
}

section "Stop prior managed processes"
stop_pid_file "cloudflared" "${CLOUDFLARED_PID_FILE}"
stop_pid_file "Vite" "${VITE_PID_FILE}"

CLOUDFLARED_BIN_RESOLVED="$(find_cloudflared)"

start_vite
LOCAL_TARGET="$(wait_for_vite | tail -n 1)"
echo "Local target: ${LOCAL_TARGET}"

start_cloudflared "${LOCAL_TARGET}" "${CLOUDFLARED_BIN_RESOLVED}"
ORIGIN="$(wait_for_tunnel_origin | tail -n 1)"
DOMAIN="${ORIGIN#https://}"
echo "Cloudflare origin: ${ORIGIN}"

write_web_env "${ORIGIN}"

section "Restart Vite with new allowed host"
stop_pid_file "Vite" "${VITE_PID_FILE}"
start_vite
LOCAL_TARGET="$(wait_for_vite | tail -n 1)"
echo "Local target: ${LOCAL_TARGET}"

regenerate_manifests
sync_entra_redirect "${ORIGIN}"

section "Done"
cat <<EOF
Dev origin:
  ${ORIGIN}

Redirect URI registered:
  ${ORIGIN}/auth-redirect.html

Logs:
  ${VITE_LOG}
  ${CLOUDFLARED_LOG}

PIDs:
  Vite:        $(cat "${VITE_PID_FILE}")
  cloudflared: $(cat "${CLOUDFLARED_PID_FILE}")

Manifest package:
  dist/package/development/xml/
EOF
