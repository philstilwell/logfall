#!/usr/bin/env python3
from __future__ import annotations

import json
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
ASSESSMENT_HTML = ROOT / "assessment" / "index.html"
OUT_PATH = ROOT / "tmp" / "imagegen" / "dialogue_assessment_scene_jobs.jsonl"


def load_bank() -> list[dict]:
    html = ASSESSMENT_HTML.read_text(encoding="utf-8")
    match = re.search(
        r'<script id="dialogue-assessment-bank" type="application/json">(.*?)</script>',
        html,
        re.S,
    )
    if not match:
        raise RuntimeError("Could not find dialogue assessment bank in assessment/index.html")
    bank = json.loads(match.group(1))
    if len(bank) != 40:
        raise RuntimeError(f"Expected 40 dialogue assessment items, found {len(bank)}")
    return bank


def tone_guidance(answer_key: str) -> tuple[str, str]:
    if answer_key.startswith("left-"):
        return (
            "left speaker should look more assertive, overconfident, or rhetorically aggressive",
            "right speaker should look more analytical, skeptical, or quietly corrective",
        )
    if answer_key.startswith("right-"):
        return (
            "left speaker should look more analytical, skeptical, or quietly corrective",
            "right speaker should look more assertive, overconfident, or rhetorically aggressive",
        )
    return (
        "left speaker should look thoughtful, careful, and fair-minded",
        "right speaker should look thoughtful, careful, and fair-minded",
    )


def build_prompt(item: dict) -> str:
    left_tone, right_tone = tone_guidance(item["answerKey"])
    lines = [turn["text"] for turn in item["turns"]]
    context = " | ".join(lines)
    fallacy_hint = item.get("fallacyName") or "no fallacy"
    return (
        "Create a grayscale mid-century editorial cartoon assessment illustration on a pure white background. "
        "Two unique humorous cartoon speakers must be fully visible from head to shoes: one pushed as close as possible to the far left page edge and one pushed as close as possible to the far right page edge, while still remaining fully visible and not cut off. "
        "Make both speakers large and tall enough that the top and bottom white space is minimized, but never crop any part of them. "
        "Leave a clean rectangular center corridor for six future dialogue boxes in two aligned columns. "
        "Do not draw any speech bubbles, text, letters, numbers, captions, labels, signs, or tail shapes. "
        "No background clutter. Only subtle ground shadows if needed. "
        "The linework should feel like a crisp grayscale editorial cartoon meant for economical printing. "
        f"{left_tone}. {right_tone}. "
        "Any gestures must stay outside the central dialogue corridor. "
        "Make the speakers visually fit the subject matter and emotional dynamic of this exchange, but do not print the words anywhere in the image. "
        f"Context for the exchange: {context} "
        f"Diagnostic context: {fallacy_hint}. "
        "The middle must remain clean and spacious enough for six same-sized rounded dialogue rectangles arranged in two aligned columns, with the opening speaker's column slightly higher than the reply column."
    )


def main() -> int:
    bank = load_bank()
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with OUT_PATH.open("w", encoding="utf-8") as fh:
        for item in bank:
            job = {
                "prompt": build_prompt(item),
                "use_case": "print-friendly assessment illustration",
                "style": "grayscale mid-century editorial cartoon, witty but classroom-serious",
                "composition": "wide landscape scene, speakers pushed toward outer edges, vertically dense full-body composition, broad empty central space, no speech bubbles or text",
                "constraints": "pure white background; grayscale only; no printed text; no captions; no tails; no cropping; keep speakers close to outer margins; minimize top and bottom whitespace; leave center open for six future dialogue boxes",
                "size": "1536x1024",
                "quality": "high",
                "output_format": "png",
                "out": f"{item['id']}-scene.png",
            }
            fh.write(json.dumps(job, ensure_ascii=False) + "\n")
    print(OUT_PATH)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
