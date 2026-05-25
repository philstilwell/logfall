#!/usr/bin/env python3

from __future__ import annotations

import argparse
import base64
import json
import time
from pathlib import Path

from openai import OpenAI


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_DATA_PATH = PROJECT_ROOT / "data" / "fallacies.json"
DEFAULT_ASSETS_DIR = PROJECT_ROOT / "assets"
DEFAULT_OUTPUT_PATH = PROJECT_ROOT / "data" / "poster_captions.json"

PROMPT_TEMPLATE = """You are writing a replacement for the section titled "What this image shows" beneath a logical-fallacy teaching poster.

Write exactly 4 sentences in this pattern:
1. Sentence 1: name the main visible elements in the image and say what immediate situation or impression they create.
2. Sentence 2: explain what further claim is being inferred from those elements, and be specific if the inferred thing is a god, spirit, fate, hidden power, conspiracy, causal force, guilt, inevitability, or some other unseen conclusion.
3. Sentence 3: begin with "The important contrast is between" and state the difference between what is visible and what is inferred.
4. Sentence 4: begin with "We can see" and explain why the visible scene does not by itself prove the inferred conclusion.

Rules:
- Use plain, precise language.
- Refer to concrete visible elements, not vague abstractions.
- Mention the most striking visual cues if they are present: symbols, props, facial expressions, split scenes, scales, doors, masks, crowds, arrows, or theatrical gestures.
- Do not quote or rely on any in-image text or labels; describe objects, symbols, and gestures instead.
- Do not invent props or labels that are not clearly visible.
- Do not say "the image works," "this poster," or other self-referential phrases.
- Do not name the fallacy.
- Keep the total length between 95 and 140 words.

Fallacy name: {name}
Definition: {definition}
Example: {example}
"""


def resolve_poster_path(assets_dir: Path, slug: str) -> Path | None:
    stem = f"fallacy-{slug}-poster"
    for extension in ("webp", "png", "jpeg", "jpg"):
        candidate = assets_dir / f"{stem}.{extension}"
        if candidate.exists():
            return candidate
    return None


def encode_image(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("ascii")


def build_caption(client: OpenAI, model: str, record: dict, poster_path: Path) -> str:
    prompt = PROMPT_TEMPLATE.format(
        name=record["name"],
        definition=record["definition"],
        example=record["example"],
    )
    mime = "image/webp" if poster_path.suffix.lower() == ".webp" else f"image/{poster_path.suffix.lower().lstrip('.')}"
    encoded = encode_image(poster_path)
    response = client.responses.create(
        model=model,
        input=[
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": prompt},
                    {"type": "input_image", "image_url": f"data:{mime};base64,{encoded}"},
                ],
            }
        ],
    )
    return " ".join(response.output_text.strip().split())


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate image-aware LogFall poster captions.")
    parser.add_argument("--data", default=str(DEFAULT_DATA_PATH))
    parser.add_argument("--assets-dir", default=str(DEFAULT_ASSETS_DIR))
    parser.add_argument("--out", default=str(DEFAULT_OUTPUT_PATH))
    parser.add_argument("--model", default="gpt-4.1")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--pause", type=float, default=0.0)
    args = parser.parse_args()

    data_path = Path(args.data)
    assets_dir = Path(args.assets_dir)
    output_path = Path(args.out)

    payload = json.loads(data_path.read_text())
    existing_payload = {"captions": {}}
    if output_path.exists():
        existing_payload = json.loads(output_path.read_text())
        if "captions" not in existing_payload:
            existing_payload = {"captions": existing_payload}

    captions = dict(existing_payload.get("captions", {}))
    client = OpenAI()

    processed = 0
    generated = 0
    skipped = 0

    for record in payload["records"]:
        slug = record["slug"]
        processed += 1
        if not args.overwrite and slug in captions:
            skipped += 1
            continue

        poster_path = resolve_poster_path(assets_dir, slug)
        if not poster_path:
            skipped += 1
            continue

        captions[slug] = build_caption(client, args.model, record, poster_path)
        generated += 1

        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(
            json.dumps(
                {
                    "model": args.model,
                    "captions": captions,
                },
                ensure_ascii=False,
                indent=2,
            )
            + "\n"
        )

        print(f"[{generated}] {record['name']}")

        if args.limit and generated >= args.limit:
            break
        if args.pause:
            time.sleep(args.pause)

    print(
        json.dumps(
            {
                "processed": processed,
                "generated": generated,
                "skipped": skipped,
                "output": str(output_path),
                "model": args.model,
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
