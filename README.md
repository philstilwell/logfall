# LogFall

Logical fallacies reference, teaching site, and GitHub Pages generator.

## Current Structure

- Published site files live at the repository root for GitHub Pages.
- Clean source data lives in [data/fallacies.json](data/fallacies.json).
- Build scripts live in [scripts](scripts).
- Visual source assets live in [site](site).

## Project Handoff

- [PROJECT_BRIEF.md](PROJECT_BRIEF.md) summarizes the project state and major sections.
- [WORKFLOW.md](WORKFLOW.md) records common edit, rebuild, and verification tasks.
- [ASSET_INVENTORY.md](ASSET_INVENTORY.md) identifies tracked public assets and local-only artifacts.

## Build

Rebuild the site and workbook:

```bash
node scripts/build.mjs
```

## Outputs

- Site home: `index.html`
- Workbook: `logfall-root-edition.xlsx`
