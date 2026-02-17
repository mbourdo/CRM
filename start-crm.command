#!/bin/zsh
set -e

cd "$(dirname "$0")"

PORT=3000
if lsof -iTCP:$PORT -sTCP:LISTEN -t >/dev/null 2>&1; then
  PORT=4173
fi

echo "Starting CRM on http://localhost:$PORT"
echo "Data file: $(pwd)/data/leads.json"
echo "Keep this window open while using the CRM."
echo ""

PORT="$PORT" node server.js &
SERVER_PID=$!

sleep 1
open "http://localhost:$PORT"

wait $SERVER_PID
