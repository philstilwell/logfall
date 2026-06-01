# LogFall Project Brief

LogFall is a static GitHub Pages site at `https://logfall.com/` for teaching logical fallacies. It combines a fallacy reference, generated fallacy pages, companion illustrations, assessment tools, theory articles, AI prompts, an interactive map, and the weekly Fallacy Detective headline-analysis feature.

## Current State

- Repository: `philstilwell/logfall`
- Local path: `/Users/philstilwell/Documents/Codex/2026-05-09/install-ghostscript/logfall-repo`
- Published branch: `main`
- Last verified baseline before this handoff: `86261f97`
- Fallacy count: `144`
- Fallacy poster count: `144`
- Assessment dialogue image count: `40`
- Site output is committed at the repository root for GitHub Pages.

## Primary Build Command

```bash
node scripts/build.mjs
```

This rebuilds generated HTML, the interactive map payload, sitemap data, and `logfall-root-edition.xlsx`.

## Data Sources

- `data/fallacies.json`: generated canonical fallacy corpus used by the site.
- `data/editorial_overrides.json`: family/category and editorial overrides.
- `data/supplemental_fallacies.json`: manually added fallacies, including Motte and Bailey.
- `data/case_study_library.json`: source-backed case-study material.
- `data/rationality_enrichment.json`: rationality-lab and practice data.
- `data/poster_captions.json`: concrete descriptions for fallacy poster images.
- `data/byteseismic_crossrefs.json`: Byteseismic related-reading links.
- `data/slugfester_crossrefs.json`: Slugfester reference links.

## Major Public Sections

- `/fallacies/`: searchable fallacy index.
- `/categories/`: diagnostic category pages.
- `/families/`: mutually exclusive family pages.
- `/map/`: interactive scatter plot using teaching gauges.
- `/features/`: Fallacy Detective weekly headline analysis.
- `/check-yourself/`: mixed fallacy identification quiz.
- `/assessment/`: dialogue-based formal/informal fallacy assessment.
- `/prompts/`: reusable AI prompts.
- `/theory/`: theory and pedagogy articles.
- `/about/`: author and project background.

## Important Working Rules

- Generated pages should be rebuilt with `node scripts/build.mjs` after generator or data changes.
- Manual edits should be made in source data or generator code, not directly in generated HTML, unless the change is deliberately temporary.
- `output/imagegen/`, `tmp/`, Python caches, and `fallacy_rationality_tool.zip` are local artifacts and are intentionally ignored.
- The committed PDF at `output/pdf/logfall-dialogue-assessment.pdf` is a public artifact and should remain tracked.

