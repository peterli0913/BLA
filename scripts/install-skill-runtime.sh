#!/usr/bin/env bash
# Install local runtimes used by presentation, PDF, and spreadsheet skills.
# Safe to re-run. Does not commit secrets. Does not vendor Anthropic proprietary skills.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
RUNTIME_DIR="${SKILL_RUNTIME_DIR:-$ROOT/.skill-runtime}"
export PATH="${HOME}/.local/bin:${PATH}"

python3 -m pip install --user -q "markitdown[pptx,docx,xlsx,pdf]" python-pptx pillow pypdf openpyxl

mkdir -p "$RUNTIME_DIR"
if [ ! -f "$RUNTIME_DIR/package.json" ]; then
  printf '%s\n' '{"name":"bla-skill-runtime","private":true}' > "$RUNTIME_DIR/package.json"
fi
if [ ! -d "$RUNTIME_DIR/node_modules/pptxgenjs" ]; then
  npm install --prefix "$RUNTIME_DIR" pptxgenjs
fi

echo "Skill runtime ready:"
python3 -c "import markitdown, pptx, PIL, pypdf, openpyxl; print(' python ok')"
node -e "require(process.argv[1])" "$RUNTIME_DIR/node_modules/pptxgenjs"
echo " node pptxgenjs ok"
echo " NODE_PATH hint: $RUNTIME_DIR/node_modules"
