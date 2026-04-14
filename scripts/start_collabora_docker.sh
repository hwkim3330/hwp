#!/usr/bin/env bash
set -euo pipefail

if ! command -v docker >/dev/null 2>&1; then
  echo "docker not found"
  exit 1
fi

docker run -d \
  --name hwp-collabora \
  -p 9980:9980 \
  -e aliasgroup1=http://127.0.0.1:8765 \
  collabora/code

echo "Collabora started at http://127.0.0.1:9980"
