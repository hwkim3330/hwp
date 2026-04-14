#!/usr/bin/env bash
set -euo pipefail

if ! command -v docker >/dev/null 2>&1; then
  echo "docker not found"
  exit 1
fi

docker run -d \
  --name hwp-onlyoffice \
  -p 8080:80 \
  -e JWT_ENABLED=false \
  onlyoffice/documentserver

echo "ONLYOFFICE Docs started at http://127.0.0.1:8080"
