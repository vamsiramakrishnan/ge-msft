#!/usr/bin/env bash
# Compatibility entrypoint; bundle.py owns inventory, normalized ZIP metadata and validation.
set -euo pipefail
exec python3 "$(dirname "$0")/bundle.py" "${1:-m365-surface-commander}"
