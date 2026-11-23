#!/usr/bin/env bash
# Render a .pptx to PDF + per-slide PNGs, build an index.html gallery,
# and optionally serve it on localhost for viewing through an SSH tunnel.
#
# Usage:
#   tools/preview.sh [deck.pptx] [outdir] [--serve]
#
# Defaults: newest *.pptx in the project root, outdir "preview", port 8000 (PORT=).
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
serve=0
args=()
for a in "$@"; do
  if [[ "$a" == "--serve" ]]; then serve=1; else args+=("$a"); fi
done

deck="${args[0]:-}"
outdir="${args[1]:-$root/preview}"
port="${PORT:-8000}"

if [[ -z "$deck" ]]; then
  deck="$(find "$root" -maxdepth 2 -name '*.pptx' -not -path '*/preview/*' -printf '%T@ %p\n' 2>/dev/null | sort -nr | head -1 | cut -d' ' -f2-)"
fi
[[ -n "$deck" && -f "$deck" ]] || { echo "No .pptx found. Pass one: tools/preview.sh path/to/deck.pptx" >&2; exit 1; }

deck="$(cd "$(dirname "$deck")" && pwd)/$(basename "$deck")"
name="$(basename "${deck%.*}")"
mkdir -p "$outdir"
outdir="$(cd "$outdir" && pwd)"

echo "Deck:  $deck"
echo "Out:   $outdir"

# 1. PPTX -> PDF (LibreOffice keeps fonts/layout close enough for review)
soffice --headless --convert-to pdf --outdir "$outdir" "$deck" >/dev/null 2>&1 ||
  { echo "LibreOffice conversion failed" >&2; exit 1; }
pdf="$outdir/$name.pdf"

# 2. PDF -> PNG per slide
rm -f "$outdir"/slide-*.png
pdftoppm -png -r 120 "$pdf" "$outdir/slide" >/dev/null 2>&1

# 3. index.html gallery
{
  printf '%s\n' '<!doctype html><meta charset="utf-8">' \
    "<title>$name — preview</title>" \
    '<style>
      body{margin:0;padding:32px;background:#16181d;color:#e7e9ee;font:15px/1.5 system-ui,sans-serif}
      header{display:flex;gap:16px;align-items:baseline;margin-bottom:24px;flex-wrap:wrap}
      h1{font-size:20px;margin:0;font-weight:600}
      a{color:#7aa2ff}
      .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(420px,1fr));gap:24px}
      figure{margin:0}
      img{width:100%;height:auto;display:block;border-radius:8px;box-shadow:0 8px 24px #0007}
      figcaption{margin-top:8px;color:#9aa0ac;font-size:13px}
    </style>' \
    '<header>'"<h1>$name</h1><a href=\"$name.pdf\">Open PDF</a><span>rendered with LibreOffice — fonts may differ slightly from PowerPoint</span></header>" \
    '<div class="grid">'
  for png in "$outdir"/slide-*.png; do
    [[ -e "$png" ]] || continue
    base="$(basename "$png")"
    n="${base#slide-}"; n="${n%.png}"
    printf '  <figure><img src="%s" alt="slide %s" loading="lazy"><figcaption>Slide %s</figcaption></figure>\n' "$base" "$n" "$n"
  done
  printf '%s\n' '</div>'
} > "$outdir/index.html"

slides=$(find "$outdir" -maxdepth 1 -name 'slide-*.png' | wc -l)
echo "Rendered $slides slide(s) + $name.pdf"
echo
echo "On your laptop, tunnel the port and open the gallery:"
echo "  ssh -N -L $port:127.0.0.1:$port $(whoami)@$(hostname -f 2>/dev/null || hostname)"
echo "  http://localhost:$port/"
echo "Or pull the files: scp $(whoami)@$(hostname):$pdf ."

if [[ "$serve" == 1 ]]; then
  echo
  # Detached on purpose: a dropped SSH connection must not take the server with it.
  PREVIEW_DIR="$outdir" PORT="$port" bash "$root/tools/serve-preview.sh"
fi
