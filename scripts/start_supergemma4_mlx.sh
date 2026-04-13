#!/bin/zsh
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
VENV_DIR="${MLX_EXPERIMENTAL_VENV:-$ROOT_DIR/.venv-mlx}"
SERVER_BIN="${MLX_EXPERIMENTAL_SERVER:-$VENV_DIR/bin/mlx_lm.server}"
MODEL_ID="${MLX_EXPERIMENTAL_MODEL:-Jiunsong/supergemma4-26b-uncensored-mlx-4bit-v2}"
HOST="${MLX_EXPERIMENTAL_HOST:-127.0.0.1}"
PORT="${MLX_EXPERIMENTAL_PORT:-8081}"

if [[ ! -x "$SERVER_BIN" ]]; then
  echo "mlx_lm.server not found: $SERVER_BIN" >&2
  echo "Create the MLX venv first or set MLX_EXPERIMENTAL_SERVER." >&2
  exit 1
fi

exec "$SERVER_BIN" \
  --model "$MODEL_ID" \
  --host "$HOST" \
  --port "$PORT" \
  --use-default-chat-template
