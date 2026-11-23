#!/usr/bin/env bash
# Start/stop the preview gallery server detached, so an SSH disconnect or a
# closed terminal does NOT kill it.
#
# Usage:
#   tools/serve-preview.sh            # start on 127.0.0.1:8000 (PORT= to change)
#   tools/serve-preview.sh --status
#   tools/serve-preview.sh --stop
#   tools/serve-preview.sh --foreground   # run in this terminal instead
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
outdir="${PREVIEW_DIR:-$root/preview}"
port="${PORT:-8000}"
pidfile="$outdir/.server.pid"
logfile="$outdir/server.log"
url="http://127.0.0.1:$port/"

is_up() { curl -fsS -o /dev/null --max-time 1 "$url" 2>/dev/null; }
pid_alive() { [[ -f "$pidfile" ]] && kill -0 "$(cat "$pidfile")" 2>/dev/null; }

tunnel_hint() {
  local user host
  user="$(whoami)"
  host="$(hostname -f 2>/dev/null || hostname)"
  echo
  echo "On YOUR LAPTOP (not on this server):"
  echo "  ssh -N -L $port:127.0.0.1:$port $user@$host"
  echo "Then open: $url"
}

case "${1:-start}" in
  --status)
    if is_up; then
      echo "running: $url (log: $logfile)"
    elif pid_alive; then
      echo "process $(cat "$pidfile") exists but $url is not answering — see $logfile"
      exit 1
    else
      echo "not running"
      exit 1
    fi
    ;;
  --stop)
    if pid_alive; then
      kill "$(cat "$pidfile")" && echo "stopped (pid $(cat "$pidfile"))"
      rm -f "$pidfile"
    else
      echo "not running"
    fi
    ;;
  --foreground)
    [[ -f "$outdir/index.html" ]] || { echo "No $outdir/index.html — run tools/preview.sh first" >&2; exit 1; }
    echo "Serving $outdir on $url (Ctrl+C to stop)"
    tunnel_hint
    exec python3 -m http.server "$port" --bind 127.0.0.1 --directory "$outdir"
    ;;
  start|"")
    [[ -f "$outdir/index.html" ]] || { echo "No $outdir/index.html — run tools/preview.sh first" >&2; exit 1; }
    if is_up; then
      echo "already running: $url"
      tunnel_hint
      exit 0
    fi
    if pid_alive; then
      echo "stale process $(cat "$pidfile") — stopping it first"
      kill "$(cat "$pidfile")" 2>/dev/null || true
      rm -f "$pidfile"
    fi
    # setsid + nohup: new session, no controlling terminal, immune to SIGHUP.
    setsid nohup python3 -m http.server "$port" --bind 127.0.0.1 --directory "$outdir" \
      >"$logfile" 2>&1 &
    echo $! >"$pidfile"
    for _ in 1 2 3 4 5 6 7 8 9 10; do is_up && break; sleep 0.3; done
    if is_up; then
      echo "started detached: $url (pid $(cat "$pidfile"), log: $logfile)"
      tunnel_hint
    else
      echo "failed to start — see $logfile" >&2
      tail -5 "$logfile" >&2 || true
      exit 1
    fi
    ;;
  *)
    echo "usage: $0 [--status|--stop|--foreground]" >&2
    exit 2
    ;;
esac
