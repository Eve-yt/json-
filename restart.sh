#!/bin/zsh

set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"

"$PROJECT_DIR/stop.sh" || true
"$PROJECT_DIR/start.sh"
