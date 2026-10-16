#!/bin/bash
token=$(printf "protocol=https\nhost=github.com\n" | git credential fill 2>/dev/null | grep "^password=" | cut -d= -f2)
for i in $(seq 1 40); do
  st=$(curl -s --max-time 20 -H "Authorization: token $token" \
    "https://api.github.com/repos/wskmp88/Invoice-printing-assistant/actions/runs?per_page=1" \
    | python -X utf8 -c "import json,sys; d=json.load(sys.stdin); r=d['workflow_runs'][0]; print(r['status'], r.get('conclusion') or '-')" 2>/dev/null)
  echo "poll $i: $st ($(date +%H:%M:%S))"
  case "$st" in
    completed*) echo "BUILD_DONE: $st"; exit 0;;
  esac
  sleep 40
done
echo "WATCH_TIMEOUT"
