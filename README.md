# presentations

Private repo for building and versioning slide decks.

## Layout

```text
decks/            # .pptx sources (and .potx templates) — committed
tools/
  preview.sh      # .pptx -> PDF + per-slide PNGs + preview/index.html
  serve-preview.sh# detached local web server for the gallery
preview/          # generated renders (git-ignored)
```

## Build a deck

Decks are generated with [`pptxgenjs`](https://gitbrent.github.io/PptxGenJS/) via the `pptx` skill; the source script for each deck lives next to its output:

```text
decks/<name>/deck.js      # generator
decks/<name>/deck.pptx    # output (committed)
decks/<name>/deck.pdf     # export (committed) — GitHub renders PDFs in the browser
```

```bash
cd decks/<name>
node deck.js
```

## Preview a deck

```bash
tools/preview.sh decks/<name>/deck.pptx     # render only
tools/preview.sh decks/<name>/deck.pptx --serve
```

`--serve` starts a detached server on `127.0.0.1:8000`, so it survives an SSH
disconnect. Manage it with:

```bash
tools/serve-preview.sh --status
tools/serve-preview.sh --stop
```

To view it from another machine, tunnel the port from your laptop (not from the
server):

```bash
ssh -N -L 8000:127.0.0.1:8000 <user>@<host>
# then open http://localhost:8000/
```

Or pull the rendered PDF:

```bash
scp <user>@<host>:~/Projects/Pfe_prez/preview/<name>.pdf .
```

LibreOffice renders the previews; fonts can differ slightly from PowerPoint, so
open the `.pptx` itself for the final check.

## Local toolchain (this machine)

`node`, `npm`, `gh` live under `~/.local` (shims in `~/.local/bin`); the Python
deps (`python-pptx`, `lxml`, `defusedxml`) are installed to the user site. The
`pptx` skill is at `~/.config/opencode/skills/pptx`.
