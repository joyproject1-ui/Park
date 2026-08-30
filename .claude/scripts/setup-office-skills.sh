#!/usr/bin/env bash
# Install the runtime dependencies the docx / pptx / xlsx skills expect.
#
# The skills themselves ship with Claude Code, but they assume a runtime that
# already has python-docx, python-pptx, openpyxl, pandas, markitdown, the npm
# `docx` / `pptxgenjs` packages, a full LibreOffice, pandoc and Poppler. A fresh
# remote container only has libreoffice-core, so office file generation fails
# until these are installed. Every step is idempotent and skipped when already
# satisfied, so re-running costs a few seconds.

set -uo pipefail

log() { printf '[office-skills] %s\n' "$*"; }

# --- apt: LibreOffice filters, pandoc, Poppler -------------------------------
apt_missing=()
for pkg in libreoffice-calc libreoffice-writer libreoffice-impress pandoc poppler-utils; do
  dpkg -s "$pkg" >/dev/null 2>&1 || apt_missing+=("$pkg")
done
if ((${#apt_missing[@]})); then
  log "apt-get install: ${apt_missing[*]}"
  DEBIAN_FRONTEND=noninteractive apt-get update -qq >/dev/null 2>&1
  DEBIAN_FRONTEND=noninteractive apt-get install -y -qq "${apt_missing[@]}" >/dev/null 2>&1 \
    || log "WARNING: apt-get install failed; soffice/pandoc conversions may not work"
fi

# --- pip ---------------------------------------------------------------------
# Import names differ from distribution names, so probe the import.
pip_missing=()
probe() { python3 -c "import $1" >/dev/null 2>&1; }
probe docx     || pip_missing+=(python-docx)
probe pptx     || pip_missing+=(python-pptx)
probe openpyxl || pip_missing+=(openpyxl)
probe pandas   || pip_missing+=(pandas)
probe PIL      || pip_missing+=(Pillow)
probe defusedxml || pip_missing+=(defusedxml)
probe lxml     || pip_missing+=(lxml)
command -v markitdown >/dev/null 2>&1 || pip_missing+=("markitdown[all]")
if ((${#pip_missing[@]})); then
  log "pip install: ${pip_missing[*]}"
  pip install --quiet --disable-pip-version-check "${pip_missing[@]}" >/dev/null 2>&1 \
    || log "WARNING: pip install failed"
fi

# --- npm ---------------------------------------------------------------------
if command -v npm >/dev/null 2>&1; then
  npm_root="$(npm root -g 2>/dev/null)"
  npm_missing=()
  for mod in docx pptxgenjs react react-dom react-icons sharp; do
    [[ -d "$npm_root/$mod" ]] || npm_missing+=("$mod")
  done
  if ((${#npm_missing[@]})); then
    log "npm install -g: ${npm_missing[*]}"
    npm install -g --silent "${npm_missing[@]}" >/dev/null 2>&1 \
      || log "WARNING: npm install failed"
  fi
  # Skills write throwaway scripts in arbitrary directories and `require('docx')`
  # directly. Node resolves bare specifiers by walking parent directories looking
  # for node_modules, so a /node_modules symlink makes the global installs
  # resolvable from anywhere — NODE_PATH would not survive, since the tool shells
  # are non-interactive and never source ~/.bashrc.
  if [[ -n "$npm_root" && ! -e /node_modules ]]; then
    ln -sfn "$npm_root" /node_modules 2>/dev/null && log "linked /node_modules -> $npm_root"
  fi
fi

log "ready: Word (.docx), PowerPoint (.pptx) and Excel (.xlsx) generation"
