# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Goal

A public web platform for **mibluemedical.com** (GitHub: `mjjuarezm`) that lets doctors look up which active pharmaceutical ingredients are in the formulary protocol. Hosted on GitHub Pages at `https://mjjuarezm.github.io/<repo-name>/`.

## How Updates Work

1. Edit `Protocolo.xlsx` locally.
2. `git add Protocolo.xlsx && git commit -m "..." && git push`
3. GitHub Actions automatically runs `python generate.py` → redeploys `index.html`.

No manual HTML editing needed.

## Key Files

| File | Purpose |
|------|---------|
| `Protocolo.xlsx` | Source of truth — all ingredient and product data |
| `generate.py` | Reads Excel → produces `index.html` |
| `index.html` | Generated output; **never edit by hand** — overwritten on every deploy |
| `protocolo_farmacia.html` | Original reference design; not deployed |
| `requirements.txt` | Python dependency: `openpyxl` |
| `.github/workflows/deploy.yml` | CI/CD: runs generate.py and deploys to GitHub Pages on push to `main` |

## Commands

```bash
# Install dependencies (Python 3.8+ required)
pip install -r requirements.txt

# Regenerate index.html locally to preview changes
python generate.py

# Open locally in browser (Windows)
start index.html
```

## Excel Schema

`generate.py` reads two sheets from `Protocolo.xlsx`:

**Sheet: `Protocolo`** — columns A–N, data from row 2
| Col | Field | Used |
|-----|-------|------|
| C | Nombre (product name) | yes |
| D | Ingrediente Activo | yes |
| E | Concentración | yes |
| J | Tipo de presentación | yes (combined with K) |
| K | Unidad de medida | yes (combined with J) |
| M | Es protocolo? | filter — only rows with value `1` |

**Sheet: `Novedades`** — data from row 4, columns B–O
| Col | Field | Used |
|-----|-------|------|
| D | Nombre del producto | yes |
| E | Ingrediente Activo | yes |
| F | Concentración | yes (pill badge) |
| K | Tipo de presentación | yes (pill badge) |
| L | Unidad de medida | yes (pill badge) |

Column mappings are defined as constants at the top of `generate.py`. If columns shift in the Excel, update those constants.

## GitHub Pages Setup (one-time)

After pushing to GitHub for the first time:
1. Go to the repo → **Settings** → **Pages**
2. Under "Source", select **GitHub Actions**
3. The next push to `main` will deploy automatically.
