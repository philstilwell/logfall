#!/usr/bin/env python3
from __future__ import annotations

import json
import re
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
ASSESSMENT_HTML = ROOT / "assessment" / "index.html"
SCENE_DIR = ROOT / "output" / "imagegen" / "dialogue-assessment-scenes"
ASSET_DIR = ROOT / "assets" / "assessment-dialogues"
FONT_PATH = Path("/System/Library/Fonts/Supplemental/Comic Sans MS.ttf")

# Alternating staggered slots modeled on the approved pilot style.
BUBBLE_SLOTS = [
    (250, 74, 468, 142),
    (692, 171, 468, 142),
    (288, 336, 430, 142),
    (724, 433, 430, 142),
    (330, 598, 430, 142),
    (756, 695, 430, 142),
]


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


def fit_font(draw: ImageDraw.ImageDraw, text: str, max_width: int, max_height: int) -> tuple[ImageFont.FreeTypeFont, tuple[int, int, int, int]]:
    for size in range(31, 18, -1):
        font = ImageFont.truetype(str(FONT_PATH), size=size)
        bbox = draw.multiline_textbbox((0, 0), text, font=font, spacing=4, align="center")
        width = bbox[2] - bbox[0]
        height = bbox[3] - bbox[1]
        if width <= max_width and height <= max_height:
            return font, bbox
    font = ImageFont.truetype(str(FONT_PATH), size=18)
    bbox = draw.multiline_textbbox((0, 0), text, font=font, spacing=4, align="center")
    return font, bbox


def rounded_shadow_box(base: Image.Image, box: tuple[int, int, int, int]) -> None:
    x, y, w, h = box
    shadow = Image.new("RGBA", base.size, (255, 255, 255, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle(
        (x + 5, y + 7, x + w + 5, y + h + 7),
        radius=24,
        fill=(0, 0, 0, 22),
    )
    base.alpha_composite(shadow)


def draw_bubble(base: Image.Image, box: tuple[int, int, int, int], text: str) -> None:
    rounded_shadow_box(base, box)
    draw = ImageDraw.Draw(base)
    x, y, w, h = box
    outline = (82, 82, 82, 255)
    fill = (253, 253, 253, 255)
    draw.rounded_rectangle((x, y, x + w, y + h), radius=24, fill=fill, outline=outline, width=3)
    font, bbox = fit_font(draw, text, w - 42, h - 30)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    tx = x + (w - tw) / 2
    ty = y + (h - th) / 2 - 2
    draw.multiline_text((tx, ty), text, font=font, fill=(26, 26, 26, 255), spacing=4, align="center")


def main() -> int:
    bank = load_bank()
    if not FONT_PATH.exists():
        raise RuntimeError(f"Missing font at {FONT_PATH}")
    ASSET_DIR.mkdir(parents=True, exist_ok=True)

    for item in bank:
        scene_path = SCENE_DIR / f"{item['id']}-scene.png"
        if not scene_path.exists():
            raise RuntimeError(f"Missing generated scene: {scene_path}")
        image = Image.open(scene_path).convert("RGBA")
        for turn, slot in zip(item["turns"], BUBBLE_SLOTS):
            draw_bubble(image, slot, turn["text"])
        out_path = ASSET_DIR / f"{item['id']}.png"
        image.convert("RGB").save(out_path, optimize=True)
        print(out_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
