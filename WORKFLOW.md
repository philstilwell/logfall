# LogFall Workflow

## Standard Edit Flow

1. Inspect the relevant source file in `data/`, `scripts/`, or `site/`.
2. Make source-level changes.
3. Rebuild:

```bash
node scripts/build.mjs
```

4. Spot-check generated HTML with `rg` or a browser.
5. Commit and push to `main`.

## Common Tasks

### Add or Revise a Fallacy

- Add new manual entries in `data/supplemental_fallacies.json`.
- Adjust families, categories, aliases, and editorial data in `data/editorial_overrides.json`.
- Add rationality/practice content in `data/rationality_enrichment.json` when needed.
- Add or update the poster caption in `data/poster_captions.json`.
- Rebuild with `node scripts/build.mjs`.

### Update Case Studies

- Prefer source-backed entries in `data/case_study_library.json`.
- Keep case studies directly connected to the fallacy being taught.
- Rebuild with `node scripts/build.mjs`.

### Update Poster Captions

- Edit `data/poster_captions.json`.
- Captions should identify concrete visual elements, name what is being inferred, and explain the contrast between what is visible and what is not justified.
- Rebuild with `node scripts/build.mjs`.

### Update Fallacy Detective

- Feature article definitions live in `scripts/build.mjs`.
- Public pages live under `/features/` after rebuild.
- Weekly headline scouting is supported by:

```bash
python3 scripts/headline_scout.py --days 7 --limit 18 --out reports/headline-scout/latest.md --json-out reports/headline-scout/latest.json
```

- The public exercise should include self-knowledge and self-control: readers should notice when political sympathy or hostility makes them eager or reluctant to diagnose a fallacy.

### Regenerate Assessment Images or PDF

- Dialogue image rendering: `scripts/render_dialogue_assessment_images.py`
- PDF rendering: `scripts/build_dialogue_assessment_pdf.py`
- Keep final public dialogue panels in `assets/assessment-dialogues/`.
- Keep the public PDF at `output/pdf/logfall-dialogue-assessment.pdf`.
- Keep raw experiments in `output/imagegen/`; they are ignored.

## Verification Shortlist

Use these checks after broad generator changes:

```bash
node scripts/build.mjs
rg -n "Fallacy Detective|Interactive Logical Fallacy Map|Assessment:" index.html features/index.html map/index.html assessment/index.html
git status --short
```

For fallacy-page changes, spot-check:

```bash
rg -n "False dilemma|Teaching gauges|What this image shows" fallacies/false-dilemma/index.html
```

