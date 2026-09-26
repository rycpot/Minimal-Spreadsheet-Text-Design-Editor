#!/usr/bin/env bash
# Builds the Chrome Web Store upload: dist/minimal-editor-v<version>.zip,
# with manifest.json at the zip's root. Repo-only files (docs, git and CI
# config, this script) are left out. Run from anywhere: scripts/package.sh
set -euo pipefail

cd "$(dirname "$0")/.."

version=$(python3 -c 'import json; print(json.load(open("manifest.json"))["version"])')
out="dist/minimal-editor-v${version}.zip"

rm -rf dist
mkdir -p dist
zip -r -X -q "$out" . \
  -x '.git/*' '.github/*' 'scripts/*' 'dist/*' \
     'README.md' 'SETUP.md' '.gitignore' '*.DS_Store'

echo "$out"
