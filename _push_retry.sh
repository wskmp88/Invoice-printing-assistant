#!/bin/bash
for i in $(seq 1 60); do
  if git -c http.proxy= -c https.proxy= push origin main 2>/dev/null; then
    echo "PUSH_OK attempt=$i $(date +%H:%M:%S)"
    exit 0
  fi
  echo "retry $i failed $(date +%H:%M:%S)"
  sleep 120
done
echo "PUSH_GIVEUP"
exit 1
