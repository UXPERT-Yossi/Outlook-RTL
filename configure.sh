#!/usr/bin/env bash
# Bakes your GitHub Pages URL into manifest.xml (replaces every __BASEURL__ placeholder).
#
# Usage:
#   ./configure.sh <github-username> [repo-name]
#
# Example:
#   ./configure.sh yossi outlook-rtl-patch
#   -> base URL becomes https://yossi.github.io/outlook-rtl-patch
set -euo pipefail

USER="${1:-}"
REPO="${2:-outlook-rtl-patch}"

if [ -z "$USER" ]; then
  echo "Usage: ./configure.sh <github-username> [repo-name]" >&2
  exit 1
fi

BASEURL="https://${USER}.github.io/${REPO}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
MANIFEST="${SCRIPT_DIR}/manifest.xml"

if grep -q "__BASEURL__" "$MANIFEST"; then
  # macOS/BSD sed in-place.
  sed -i '' "s|__BASEURL__|${BASEURL}|g" "$MANIFEST"
  echo "Set base URL to: ${BASEURL}"
  echo "manifest.xml is ready to deploy/sideload."
else
  echo "No __BASEURL__ placeholders found — manifest.xml may already be configured." >&2
  echo "Current bt:Url entries:" >&2
  grep -o 'DefaultValue="https://[^"]*"' "$MANIFEST" | head -3 >&2
fi
