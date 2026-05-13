#!/usr/bin/env python3
from __future__ import annotations

import json
import math
import re
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph


ROOT = Path(__file__).resolve().parents[1]
ASSESSMENT_HTML = ROOT / "assessment" / "index.html"
OUTPUT_DIR = ROOT / "output" / "pdf"
OUTPUT_PATH = OUTPUT_DIR / "logfall-dialogue-assessment.pdf"
ASSET_DIR = ROOT / "assets" / "assessment-dialogues"

PAGE_WIDTH, PAGE_HEIGHT = landscape(letter)
MARGIN_X = 0.45 * inch
MARGIN_TOP = 0.42 * inch
MARGIN_BOTTOM = 0.42 * inch
COLUMN_GAP = 0.32 * inch
CONTENT_WIDTH = PAGE_WIDTH - (2 * MARGIN_X)
COLUMN_WIDTH = (CONTENT_WIDTH - COLUMN_GAP) / 2
IMAGE_RATIO = 940 / 1536  # height / width
TOTAL_PAGES = 40
ITEMS_PER_PAGE = 2

CHOICE_LABELS = {
    "left-formal": ("Left", "Formal"),
    "left-informal": ("Left", "Informal"),
    "none": ("No speaker", "None"),
    "right-informal": ("Right", "Informal"),
    "right-formal": ("Right", "Formal"),
}


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


def styles() -> dict[str, ParagraphStyle]:
    sample = getSampleStyleSheet()
    return {
        "part": ParagraphStyle(
            "part",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=9.2,
            leading=11,
            textColor=colors.HexColor("#145a8d"),
        ),
        "title": ParagraphStyle(
            "title",
            parent=sample["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=20,
            textColor=colors.HexColor("#111111"),
        ),
        "subtitle": ParagraphStyle(
            "subtitle",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=12.5,
            textColor=colors.HexColor("#4b5563"),
        ),
        "item_label": ParagraphStyle(
            "item_label",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#145a8d"),
        ),
        "item_title": ParagraphStyle(
            "item_title",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=12.2,
            leading=14.2,
            textColor=colors.HexColor("#111111"),
        ),
        "small": ParagraphStyle(
            "small",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=8.3,
            leading=10.4,
            textColor=colors.HexColor("#4b5563"),
        ),
        "answer": ParagraphStyle(
            "answer",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=9.7,
            leading=12,
            textColor=colors.white,
        ),
        "body": ParagraphStyle(
            "body",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=9.4,
            leading=12,
            textColor=colors.HexColor("#1f2937"),
        ),
        "link": ParagraphStyle(
            "link",
            parent=sample["BodyText"],
            fontName="Helvetica",
            fontSize=8.6,
            leading=10.8,
            textColor=colors.HexColor("#145a8d"),
        ),
        "choice_top": ParagraphStyle(
            "choice_top",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=6.6,
            leading=8.2,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#64748b"),
        ),
        "choice_bottom": ParagraphStyle(
            "choice_bottom",
            parent=sample["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8.2,
            leading=9.8,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#111827"),
        ),
    }


def draw_paragraph(pdf: canvas.Canvas, text: str, style: ParagraphStyle, x: float, y_top: float, width: float) -> float:
    para = Paragraph(text, style)
    _, height = para.wrap(width, PAGE_HEIGHT)
    para.drawOn(pdf, x, y_top - height)
    return y_top - height


def draw_header(pdf: canvas.Canvas, page_num: int, part_label: str, title: str, subtitle: str, styleset: dict[str, ParagraphStyle]) -> float:
    y = PAGE_HEIGHT - MARGIN_TOP
    pdf.setStrokeColor(colors.HexColor("#d8dee8"))
    pdf.line(MARGIN_X, y + 6, PAGE_WIDTH - MARGIN_X, y + 6)
    y = draw_paragraph(pdf, part_label, styleset["part"], MARGIN_X, y, CONTENT_WIDTH)
    y -= 4
    y = draw_paragraph(pdf, title, styleset["title"], MARGIN_X, y, CONTENT_WIDTH)
    y -= 2
    y = draw_paragraph(pdf, subtitle, styleset["subtitle"], MARGIN_X, y, CONTENT_WIDTH)
    pdf.setFillColor(colors.HexColor("#64748b"))
    pdf.setFont("Helvetica-Bold", 8.2)
    pdf.drawRightString(PAGE_WIDTH - MARGIN_X, PAGE_HEIGHT - MARGIN_TOP + 2, f"Page {page_num} of {TOTAL_PAGES}")
    return y - 10


def draw_choice_row(pdf: canvas.Canvas, x: float, y_top: float, width: float, styleset: dict[str, ParagraphStyle], *, show_answer: str | None = None) -> float:
    gap = 6
    box_w = (width - (gap * 4)) / 5
    box_h = 34
    order = ["left-formal", "left-informal", "none", "right-informal", "right-formal"]
    cursor_x = x
    for key in order:
        top, bottom = CHOICE_LABELS[key]
        selected = show_answer == key
        pdf.setFillColor(colors.HexColor("#e9f7ef") if selected else colors.white)
        pdf.setStrokeColor(colors.HexColor("#22a06b") if selected else colors.HexColor("#d5dbe4"))
        pdf.roundRect(cursor_x, y_top - box_h, box_w, box_h, 10, fill=1, stroke=1)
        top_p = Paragraph(top, styleset["choice_top"])
        bottom_p = Paragraph(bottom, styleset["choice_bottom"])
        _, top_h = top_p.wrap(box_w - 10, PAGE_HEIGHT)
        _, bottom_h = bottom_p.wrap(box_w - 10, PAGE_HEIGHT)
        top_p.drawOn(pdf, cursor_x + 5, y_top - 8 - top_h)
        bottom_p.drawOn(pdf, cursor_x + 5, y_top - 12 - top_h - bottom_h)
        if not selected:
            pdf.setStrokeColor(colors.HexColor("#64748b"))
            pdf.circle(cursor_x + 10, y_top - 10, 4, stroke=1, fill=0)
        cursor_x += box_w + gap
    return y_top - box_h


def image_height(width: float) -> float:
    return width * IMAGE_RATIO


def item_image_path(item: dict) -> Path:
    return ASSET_DIR / f"{item['id']}.png"


def draw_test_item(pdf: canvas.Canvas, item: dict, item_number: int, x: float, y_top: float, styleset: dict[str, ParagraphStyle]) -> None:
    y = y_top
    y = draw_paragraph(pdf, f"Assessment Item {item_number} of 40", styleset["item_label"], x, y, COLUMN_WIDTH)
    y -= 2
    y = draw_paragraph(
        pdf,
        "Choose one: Left Formal, Left Informal, No Fallacy, Right Informal, or Right Formal.",
        styleset["item_title"],
        x,
        y,
        COLUMN_WIDTH,
    )
    y -= 8

    img_path = item_image_path(item)
    img_h = image_height(COLUMN_WIDTH)
    pdf.setStrokeColor(colors.HexColor("#d5dbe4"))
    pdf.setFillColor(colors.white)
    pdf.roundRect(x, y - img_h, COLUMN_WIDTH, img_h, 16, fill=1, stroke=1)
    pdf.drawImage(str(img_path), x, y - img_h, width=COLUMN_WIDTH, height=img_h, preserveAspectRatio=True, mask="auto")
    y -= img_h + 10

    y = draw_paragraph(pdf, "Mark one choice:", styleset["small"], x, y, COLUMN_WIDTH)
    y -= 8
    draw_choice_row(pdf, x, y, COLUMN_WIDTH, styleset)


def draw_answer_item(pdf: canvas.Canvas, item: dict, item_number: int, x: float, y_top: float, styleset: dict[str, ParagraphStyle]) -> None:
    y = y_top
    y = draw_paragraph(pdf, f"Answer Key Item {item_number} of 40", styleset["item_label"], x, y, COLUMN_WIDTH)
    y -= 2
    title = item.get("fallacyName") or "No fallacy"
    y = draw_paragraph(pdf, title, styleset["item_title"], x, y, COLUMN_WIDTH)
    y -= 8

    img_path = item_image_path(item)
    img_h = image_height(COLUMN_WIDTH)
    pdf.setStrokeColor(colors.HexColor("#d5dbe4"))
    pdf.setFillColor(colors.white)
    pdf.roundRect(x, y - img_h, COLUMN_WIDTH, img_h, 16, fill=1, stroke=1)
    pdf.drawImage(str(img_path), x, y - img_h, width=COLUMN_WIDTH, height=img_h, preserveAspectRatio=True, mask="auto")
    y -= img_h + 10

    answer_label = (
        "Correct answer: None"
        if item["answerKey"] == "none"
        else f"Correct answer: {item['answerKey'].replace('-', ' ').title()}"
    )
    bar_h = 24
    pdf.setFillColor(colors.HexColor("#0f172a"))
    pdf.roundRect(x, y - bar_h, COLUMN_WIDTH, bar_h, 10, fill=1, stroke=0)
    draw_paragraph(pdf, answer_label, styleset["answer"], x + 10, y - 4, COLUMN_WIDTH - 20)
    y -= bar_h + 10

    y = draw_paragraph(pdf, item["explanation"], styleset["body"], x, y, COLUMN_WIDTH)
    y -= 8
    if item.get("fallacyUrl"):
        draw_paragraph(pdf, f"LogFall reference: {item['fallacyUrl']}", styleset["link"], x, y, COLUMN_WIDTH)
    else:
        draw_paragraph(pdf, "This is a control item with no fallacy.", styleset["small"], x, y, COLUMN_WIDTH)


def build_pdf(bank: list[dict]) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    styleset = styles()
    pdf = canvas.Canvas(str(OUTPUT_PATH), pagesize=landscape(letter))
    pdf.setTitle("LogFall Dialogue Assessment")
    pdf.setAuthor("Phil Stilwell")
    pdf.setSubject("40-item dialogue assessment with illustrated items and answer commentary")

    test_pages = math.ceil(len(bank) / ITEMS_PER_PAGE)

    # Test half
    page_num = 1
    for page_index in range(test_pages):
        y_top = draw_header(
            pdf,
            page_num,
            "Part I",
            "Illustrated dialogue assessment",
            "Choose whether the fallacy is on the left, on the right, or nowhere, and whether it is formal or informal.",
            styleset,
        )
        first_index = page_index * 2
        left_item = bank[first_index]
        draw_test_item(pdf, left_item, first_index + 1, MARGIN_X, y_top, styleset)
        if first_index + 1 < len(bank):
            right_item = bank[first_index + 1]
            draw_test_item(pdf, right_item, first_index + 2, MARGIN_X + COLUMN_WIDTH + COLUMN_GAP, y_top, styleset)
        pdf.showPage()
        page_num += 1

    # Answer half
    for page_index in range(test_pages):
        y_top = draw_header(
            pdf,
            page_num,
            "Part II",
            "Answers and commentary",
            "Each illustrated item is shown again with the correct diagnosis and a short explanation of why it fits.",
            styleset,
        )
        first_index = page_index * 2
        left_item = bank[first_index]
        draw_answer_item(pdf, left_item, first_index + 1, MARGIN_X, y_top, styleset)
        if first_index + 1 < len(bank):
            right_item = bank[first_index + 1]
            draw_answer_item(pdf, right_item, first_index + 2, MARGIN_X + COLUMN_WIDTH + COLUMN_GAP, y_top, styleset)
        pdf.showPage()
        page_num += 1

    pdf.save()


def main() -> int:
    bank = load_bank()
    missing = [item["id"] for item in bank if not item_image_path(item).exists()]
    if missing:
        raise RuntimeError(f"Missing assessment dialogue images for: {', '.join(missing)}")
    build_pdf(bank)
    print(OUTPUT_PATH)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
