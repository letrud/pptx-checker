#!/usr/bin/env bash
# Validate an Office file with the bundled Open XML SDK validator.
#   validate.sh <file> [--json] [--format Office2016]
# Exit: 0 valid · 1 errors found · 2 usage · 3 no validator binary for this platform
set -euo pipefail

here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
bin_dir="${here}/../bin"

case "$(uname -s)-$(uname -m)" in
  Linux-x86_64)        exe="ooxml-validate-linux-x64" ;;
  Linux-aarch64)       exe="ooxml-validate-linux-arm64" ;;
  Darwin-arm64)        exe="ooxml-validate-osx-arm64" ;;
  MINGW*|MSYS*|CYGWIN*) exe="ooxml-validate-win-x64.exe" ;;
  *)                   exe="" ;;
esac

if [[ -z "$exe" || ! -f "${bin_dir}/${exe}" ]]; then
  echo "No validator binary for $(uname -s)-$(uname -m) in ${bin_dir}." >&2
  echo "Download it from the latest release of github.com/letrud/pptx-checker into that folder." >&2
  exit 3
fi

chmod +x "${bin_dir}/${exe}" 2>/dev/null || true
exec "${bin_dir}/${exe}" "$@"
