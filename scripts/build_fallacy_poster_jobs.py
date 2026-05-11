#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_DATA_PATH = PROJECT_ROOT / "data" / "fallacies.json"
DEFAULT_ASSETS_DIR = PROJECT_ROOT / "assets"
DEFAULT_OUTPUT_PATH = PROJECT_ROOT / "tmp" / "imagegen" / "fallacy_poster_jobs.jsonl"

STYLE = (
    "retro mid-century modern editorial cartoon illustration in a 1950s-1960s educational-poster style; "
    "simplified geometric characters; angular faces; exaggerated noses; expressive theatrical gestures; "
    "thick dark clean outlines with slight hand-drawn irregularity; muted browns, creams, ochres, burnt reds, "
    "black, and restrained cool blue accents; subtle paper grain; distressed screen-print texture; "
    "polished vintage didactic tone; clear symbolic staging"
)

COMPOSITION = (
    "vertical poster composition, one dominant central scene, bold readable silhouettes, minimal background clutter, "
    "foreground action with one or two secondary cues that clarify the reasoning mistake"
)

LIGHTING = "vintage poster contrast with warm highlights, deep shadows, and slightly noir dramatic focus"

PALETTE = "cream, deep brown, ochre gold, burnt red, black, muted cool blue accents"

MATERIALS = "paper grain, stipple, distressed ink, lightly worn poster surface, screen-print texture"

CONSTRAINTS = (
    "make the mistaken reasoning immediately understandable to a first-time viewer; depict the actual fallacy rather than just the topic; "
    "2 to 4 main figures maximum; no watermark; no logos; no photorealism; no anime; no modern UI; no decorative clutter; "
    "do not include the fallacy title in the image; avoid in-image text unless absolutely necessary for clarity, and if text appears keep it very short and perfectly legible; "
    "maintain strong style consistency across the full series"
)

NEGATIVE = (
    "photorealism, anime, superhero comic style, glossy 3D rendering, generic clip art, meme style, cluttered collage, "
    "stock-photo look, empty background, unrelated symbolism, illegible text, malformed hands"
)


def build_prompt(record: dict) -> str:
    return (
        f"Create a companion poster image for the logical fallacy '{record['name']}'. "
        f"The image must clearly depict this fallacy: {record['definition']} "
        f"Use a single concrete symbolic scene that makes the mistaken reasoning obvious at a glance. "
        f"If helpful, draw from this example scenario: {record['example']} "
        f"The viewer should understand what is being confused, distorted, ignored, or overclaimed from the visual staging alone."
    )


def build_job(record: dict, output_format: str) -> dict:
    return {
        "prompt": build_prompt(record),
        "use_case": "illustration-story",
        "style": STYLE,
        "composition": COMPOSITION,
        "lighting": LIGHTING,
        "palette": PALETTE,
        "materials": MATERIALS,
        "constraints": CONSTRAINTS,
        "negative": NEGATIVE,
        "size": "1024x1536",
        "quality": "medium",
        "output_format": output_format,
        "out": f"fallacy-{record['slug']}-poster.{output_format}",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Build JSONL jobs for LogFall companion poster generation.")
    parser.add_argument("--data", default=str(DEFAULT_DATA_PATH))
    parser.add_argument("--assets-dir", default=str(DEFAULT_ASSETS_DIR))
    parser.add_argument("--out", default=str(DEFAULT_OUTPUT_PATH))
    parser.add_argument("--output-format", default="webp", choices=["png", "webp", "jpeg"])
    parser.add_argument("--missing-only", action="store_true")
    args = parser.parse_args()

    data_path = Path(args.data)
    assets_dir = Path(args.assets_dir)
    output_path = Path(args.out)

    payload = json.loads(data_path.read_text())
    records = payload["records"]

    jobs = []
    for record in records:
        expected_asset = assets_dir / f"fallacy-{record['slug']}-poster.{args.output_format}"
        if args.missing_only and expected_asset.exists():
            continue
        jobs.append(build_job(record, args.output_format))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("".join(json.dumps(job, ensure_ascii=False) + "\n" for job in jobs))

    print(
        json.dumps(
            {
                "output": str(output_path),
                "jobCount": len(jobs),
                "missingOnly": args.missing_only,
                "outputFormat": args.output_format,
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
