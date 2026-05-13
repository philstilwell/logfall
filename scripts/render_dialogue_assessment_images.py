#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
ASSESSMENT_HTML = ROOT / "assessment" / "index.html"
SCENE_DIR = ROOT / "output" / "imagegen" / "dialogue-assessment-scenes"
ASSET_DIR = ROOT / "assets" / "assessment-dialogues"
CROP_HEIGHT = 940

FONT_CANDIDATES = [
    Path("/System/Library/Fonts/MarkerFelt.ttc"),
    Path("/System/Library/Fonts/Supplemental/Bradley Hand Bold.ttf"),
    Path("/System/Library/Fonts/Noteworthy.ttc"),
    Path("/System/Library/Fonts/Supplemental/ChalkboardSE.ttc"),
    Path("/System/Library/Fonts/Supplemental/Comic Sans MS.ttf"),
]


@dataclass(frozen=True)
class BubbleSlot:
    x: int
    y: int
    width: int
    height: int


# Two rigorously aligned columns of uniform boxes.
# The opening speaker's column is raised so the conversation order is obvious.
LEFT_X = 330
RIGHT_X = 786
BUBBLE_WIDTH = 420
BUBBLE_HEIGHT = 164
LEAD_YS = [42, 278, 514]
FOLLOW_YS = [84, 320, 556]

PADDING_X = 28
PADDING_Y = 22
TEXT_SPACING = 6
TEXT_FILL = (27, 27, 27, 255)
BOX_FILL = (252, 252, 252, 255)
BOX_OUTLINE = (84, 84, 84, 255)
SHADOW_FILL = (0, 0, 0, 18)


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


def resolve_font_path() -> Path:
    for candidate in FONT_CANDIDATES:
        if candidate.exists():
            return candidate
    raise RuntimeError("No suitable handwritten font found on this system.")


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont, max_width: int) -> list[str]:
    words = text.split()
    if not words:
        return [""]

    lines: list[str] = []
    current = words[0]
    for word in words[1:]:
        trial = f"{current} {word}"
        trial_bbox = draw.textbbox((0, 0), trial, font=font)
        if (trial_bbox[2] - trial_bbox[0]) <= max_width:
            current = trial
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def fit_wrapped_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    slot: BubbleSlot,
    font_path: Path,
) -> tuple[ImageFont.FreeTypeFont, list[str], tuple[int, int]]:
    max_width = slot.width - (PADDING_X * 2)
    max_height = slot.height - (PADDING_Y * 2)

    for size in range(31, 19, -1):
        font = ImageFont.truetype(str(font_path), size=size)
        lines = wrap_text(draw, text, font, max_width)
        paragraph = "\n".join(lines)
        bbox = draw.multiline_textbbox(
            (0, 0),
            paragraph,
            font=font,
            spacing=TEXT_SPACING,
            align="center",
        )
        width = bbox[2] - bbox[0]
        height = bbox[3] - bbox[1]
        if width <= max_width and height <= max_height and len(lines) <= 4:
            return font, lines, (width, height)

    font = ImageFont.truetype(str(font_path), size=19)
    lines = wrap_text(draw, text, font, max_width)
    paragraph = "\n".join(lines)
    bbox = draw.multiline_textbbox(
        (0, 0),
        paragraph,
        font=font,
        spacing=TEXT_SPACING,
        align="center",
    )
    width = bbox[2] - bbox[0]
    height = bbox[3] - bbox[1]
    return font, lines, (width, height)


def rounded_shadow_box(base: Image.Image, slot: BubbleSlot) -> None:
    shadow = Image.new("RGBA", base.size, (255, 255, 255, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle(
        (slot.x + 6, slot.y + 8, slot.x + slot.width + 6, slot.y + slot.height + 8),
        radius=26,
        fill=SHADOW_FILL,
    )
    base.alpha_composite(shadow)


def draw_bubble(base: Image.Image, slot: BubbleSlot, text: str, font_path: Path) -> None:
    rounded_shadow_box(base, slot)
    draw = ImageDraw.Draw(base)
    draw.rounded_rectangle(
        (slot.x, slot.y, slot.x + slot.width, slot.y + slot.height),
        radius=26,
        fill=BOX_FILL,
        outline=BOX_OUTLINE,
        width=3,
    )

    font, lines, (tw, th) = fit_wrapped_text(draw, text, slot, font_path)
    paragraph = "\n".join(lines)
    tx = slot.x + (slot.width - tw) / 2
    ty = slot.y + (slot.height - th) / 2 - 2
    draw.multiline_text(
        (tx, ty),
        paragraph,
        font=font,
        fill=TEXT_FILL,
        spacing=TEXT_SPACING,
        align="center",
    )


def slots_for_item(item: dict) -> list[BubbleSlot]:
    first_side = item["turns"][0]["side"]
    if first_side == "left":
        left_ys, right_ys = LEAD_YS, FOLLOW_YS
    else:
        left_ys, right_ys = FOLLOW_YS, LEAD_YS

    left_slots = [BubbleSlot(LEFT_X, y, BUBBLE_WIDTH, BUBBLE_HEIGHT) for y in left_ys]
    right_slots = [BubbleSlot(RIGHT_X, y, BUBBLE_WIDTH, BUBBLE_HEIGHT) for y in right_ys]

    ordered: list[BubbleSlot] = []
    left_index = 0
    right_index = 0
    for turn in item["turns"]:
        if turn["side"] == "left":
            ordered.append(left_slots[left_index])
            left_index += 1
        else:
            ordered.append(right_slots[right_index])
            right_index += 1
    return ordered


def render_item(item: dict, out_dir: Path, font_path: Path, scene_dir: Path) -> Path:
    scene_path = scene_dir / f"{item['id']}-scene.png"
    if not scene_path.exists():
        raise RuntimeError(f"Missing generated scene: {scene_path}")

    image = Image.open(scene_path).convert("RGBA")
    for turn, slot in zip(item["turns"], slots_for_item(item)):
        draw_bubble(image, slot, turn["text"], font_path)

    if image.height > CROP_HEIGHT:
        top = max(0, (image.height - CROP_HEIGHT) // 2)
        image = image.crop((0, top, image.width, top + CROP_HEIGHT))

    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{item['id']}.png"
    image.convert("RGB").save(out_path, optimize=True)
    return out_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ids", nargs="*", help="Optional subset of assessment item ids to render.")
    parser.add_argument("--out-dir", default=str(ASSET_DIR), help="Directory to write rendered panels to.")
    parser.add_argument("--font-path", help="Optional explicit font path to use for dialogue rendering.")
    parser.add_argument("--scene-dir", default=str(SCENE_DIR), help="Directory containing <id>-scene.png files.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    font_path = Path(args.font_path) if args.font_path else resolve_font_path()
    if not font_path.exists():
        raise RuntimeError(f"Missing font at {font_path}")
    bank = load_bank()
    selected_ids = set(args.ids or [])
    if selected_ids:
        bank = [item for item in bank if item["id"] in selected_ids]
        missing = selected_ids - {item["id"] for item in bank}
        if missing:
            raise RuntimeError(f"Unknown assessment ids: {', '.join(sorted(missing))}")

    out_dir = Path(args.out_dir)
    scene_dir = Path(args.scene_dir)
    for item in bank:
        out_path = render_item(item, out_dir, font_path, scene_dir)
        print(out_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
