# logfall

Logical Fallacies repository and GitHub Pages site generator.

## Current Structure

- Published site files live at the repository root for GitHub Pages.
- Clean source data lives in [data/fallacies.json](/Users/philstilwell/Documents/Codex/2026-05-09/install-ghostscript/logfall-repo/data/fallacies.json).
- Build scripts live in [scripts](/Users/philstilwell/Documents/Codex/2026-05-09/install-ghostscript/logfall-repo/scripts).
- Visual source assets live in [site](/Users/philstilwell/Documents/Codex/2026-05-09/install-ghostscript/logfall-repo/site).

## Source Rules

- Content source: `ROOT` tab from the legacy workbook.
- Inclusion rule: only fallacies that also appeared on the legacy WordPress site are included.
- Logo retained from the original LogFall assets.

## Build

1. Refresh the cleaned data:
   - `python3 scripts/extract_root.py`
2. Rebuild the site and workbook:
   - `node scripts/build.mjs`

## Outputs

- Site home: `index.html`
- Workbook: `logfall-root-edition.xlsx`
