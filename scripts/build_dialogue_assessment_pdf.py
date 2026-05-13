#!/usr/bin/env python3
from __future__ import annotations

import json
import re
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph


ROOT = Path(__file__).resolve().parents[1]
ASSESSMENT_HTML = ROOT / "assessment" / "index.html"
OUTPUT_DIR = ROOT / "output" / "pdf"
OUTPUT_PATH = OUTPUT_DIR / "logfall-dialogue-assessment.pdf"

PAGE_WIDTH, PAGE_HEIGHT = letter
MARGIN_X = 0.58 * inch
MARGIN_TOP = 0.62 * inch
MARGIN_BOTTOM = 0.58 * inch
CONTENT_WIDTH = PAGE_WIDTH - (2 * MARGIN_X)
LEFT_X = MARGIN_X
RIGHT_X = PAGE_WIDTH / 2 + 0.12 * inch
BUBBLE_WIDTH = (CONTENT_WIDTH / 2) - 0.12 * inch
TOTAL_PAGES = 80

CHOICE_LABELS = {
    "left-formal": ("Left", "Formal"),
    "left-informal": ("Left", "Informal"),
    "none": ("No speaker", "None"),
    "right-informal": ("Right", "Informal"),
    "right-formal": ("Right", "Formal"),
}

LINK_BASE = "https://logfall.com/fallacies/"


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


def make_styles():
    styles = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "title",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=19,
            leading=22,
            textColor=colors.HexColor("#111111"),
            spaceAfter=4,
        ),
        "subtitle": ParagraphStyle(
            "subtitle",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10.5,
            leading=13.5,
            textColor=colors.HexColor("#4b5563"),
        ),
        "part": ParagraphStyle(
            "part",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#145a8d"),
        ),
        "turn_label": ParagraphStyle(
            "turn_label",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=10,
            textColor=colors.HexColor("#475569"),
        ),
        "turn_text": ParagraphStyle(
            "turn_text",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9.8,
            leading=13.5,
            textColor=colors.HexColor("#111827"),
        ),
        "prompt": ParagraphStyle(
            "prompt",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=12,
            textColor=colors.HexColor("#111111"),
        ),
        "choice_top": ParagraphStyle(
            "choice_top",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.6,
            leading=9,
            textColor=colors.HexColor("#64748b"),
            alignment=TA_CENTER,
        ),
        "choice_bottom": ParagraphStyle(
            "choice_bottom",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=9.4,
            leading=11,
            textColor=colors.HexColor("#111827"),
            alignment=TA_CENTER,
        ),
        "answer_bar": ParagraphStyle(
            "answer_bar",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=10.2,
            leading=13,
            textColor=colors.white,
        ),
        "body": ParagraphStyle(
            "body",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=13.5,
            textColor=colors.HexColor("#1f2937"),
        ),
        "body_bold": ParagraphStyle(
            "body_bold",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=13.5,
            textColor=colors.HexColor("#111111"),
        ),
        "small": ParagraphStyle(
            "small",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8.7,
            leading=11.5,
            textColor=colors.HexColor("#4b5563"),
        ),
        "link": ParagraphStyle(
            "link",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9.2,
            leading=12,
            textColor=colors.HexColor("#145a8d"),
        ),
    }


def draw_paragraph(pdf: canvas.Canvas, text: str, style: ParagraphStyle, x: float, y_top: float, width: float) -> float:
    para = Paragraph(text, style)
    _, height = para.wrap(width, PAGE_HEIGHT)
    para.drawOn(pdf, x, y_top - height)
    return y_top - height


def draw_page_header(pdf: canvas.Canvas, page_num: int, title: str, subtitle: str, part_label: str, styles: dict) -> float:
    y = PAGE_HEIGHT - MARGIN_TOP
    pdf.setStrokeColor(colors.HexColor("#cbd5e1"))
    pdf.line(MARGIN_X, y + 7, PAGE_WIDTH - MARGIN_X, y + 7)
    y = draw_paragraph(pdf, part_label, styles["part"], MARGIN_X, y, CONTENT_WIDTH)
    y -= 6
    y = draw_paragraph(pdf, title, styles["title"], MARGIN_X, y, CONTENT_WIDTH)
    y -= 2
    y = draw_paragraph(pdf, subtitle, styles["subtitle"], MARGIN_X, y, CONTENT_WIDTH)

    page_label = f"Page {page_num} of {TOTAL_PAGES}"
    pdf.setFillColor(colors.HexColor("#64748b"))
    pdf.setFont("Helvetica-Bold", 8.5)
    pdf.drawRightString(PAGE_WIDTH - MARGIN_X, PAGE_HEIGHT - MARGIN_TOP + 2, page_label)
    return y - 14


def draw_turn(pdf: canvas.Canvas, turn: dict, turn_number: int, y_top: float, styles: dict) -> float:
    x = LEFT_X if turn["side"] == "left" else RIGHT_X
    label = f"{'Left' if turn['side'] == 'left' else 'Right'} {turn_number}"
    fill = colors.HexColor("#eef6ff") if turn["side"] == "left" else colors.HexColor("#fff4ef")
    stroke = colors.HexColor("#bfdcf3") if turn["side"] == "left" else colors.HexColor("#f0c9b9")

    label_text = Paragraph(label, styles["turn_label"])
    _, label_h = label_text.wrap(BUBBLE_WIDTH - 20, PAGE_HEIGHT)
    turn_text = Paragraph(turn["text"], styles["turn_text"])
    _, text_h = turn_text.wrap(BUBBLE_WIDTH - 20, PAGE_HEIGHT)
    bubble_h = label_h + text_h + 20

    pdf.setFillColor(fill)
    pdf.setStrokeColor(stroke)
    pdf.roundRect(x, y_top - bubble_h, BUBBLE_WIDTH, bubble_h, 12, stroke=1, fill=1)
    label_text.drawOn(pdf, x + 10, y_top - 10 - label_h)
    turn_text.drawOn(pdf, x + 10, y_top - 16 - label_h - text_h)
    return y_top - bubble_h - 8


def draw_dialogue(pdf: canvas.Canvas, item: dict, y_top: float, styles: dict) -> float:
    current_y = y_top
    left_count = 0
    right_count = 0
    for turn in item["turns"]:
      if turn["side"] == "left":
          left_count += 1
          turn_number = left_count
      else:
          right_count += 1
          turn_number = right_count
      current_y = draw_turn(pdf, turn, turn_number, current_y, styles)
    return current_y


def draw_choices_row(pdf: canvas.Canvas, y_top: float, styles: dict) -> float:
    pdf.setFont("Helvetica-Bold", 10, leading=None)
    pdf.setFillColor(colors.HexColor("#111111"))
    pdf.drawString(MARGIN_X, y_top, "Mark one choice:")
    y = y_top - 10
    gap = 8
    box_w = (CONTENT_WIDTH - (gap * 4)) / 5
    box_h = 42
    x = MARGIN_X
    for key in ["left-formal", "left-informal", "none", "right-informal", "right-formal"]:
        side, kind = CHOICE_LABELS[key]
        pdf.setFillColor(colors.white)
        pdf.setStrokeColor(colors.HexColor("#cbd5e1"))
        pdf.roundRect(x, y - box_h, box_w, box_h, 10, stroke=1, fill=1)
        pdf.circle(x + 12, y - (box_h / 2), 5, stroke=1, fill=0)
        top = Paragraph(side, styles["choice_top"])
        bottom = Paragraph(kind, styles["choice_bottom"])
        _, top_h = top.wrap(box_w - 28, 40)
        _, bottom_h = bottom.wrap(box_w - 28, 40)
        top.drawOn(pdf, x + 20, y - 11 - top_h)
        bottom.drawOn(pdf, x + 20, y - 15 - top_h - bottom_h)
        x += box_w + gap
    return y - box_h - 8


def draw_answer_summary(pdf: canvas.Canvas, item: dict, y_top: float, styles: dict) -> float:
    if item["answerKey"] == "none":
        summary = "Correct diagnosis: None"
    else:
        choice = " ".join(CHOICE_LABELS[item["answerKey"]])
        summary = f"Correct diagnosis: {choice} - {item['fallacyName']}"

    bar_h = 26
    pdf.setFillColor(colors.HexColor("#1e293b"))
    pdf.setStrokeColor(colors.HexColor("#1e293b"))
    pdf.roundRect(MARGIN_X, y_top - bar_h, CONTENT_WIDTH, bar_h, 12, stroke=0, fill=1)
    draw_paragraph(pdf, summary, styles["answer_bar"], MARGIN_X + 12, y_top - 5, CONTENT_WIDTH - 24)
    return y_top - bar_h - 12


def draw_footer(pdf: canvas.Canvas):
    footer_y = MARGIN_BOTTOM - 6
    pdf.setStrokeColor(colors.HexColor("#e2e8f0"))
    pdf.line(MARGIN_X, footer_y + 12, PAGE_WIDTH - MARGIN_X, footer_y + 12)
    pdf.setFont("Helvetica", 8.3)
    pdf.setFillColor(colors.HexColor("#64748b"))
    pdf.drawString(MARGIN_X, footer_y, "LogFall printable assessment")
    pdf.drawRightString(PAGE_WIDTH - MARGIN_X, footer_y, "Copyright © Phil Stilwell")


def render_question_page(pdf: canvas.Canvas, item: dict, index: int, styles: dict):
    y = draw_page_header(
        pdf,
        index + 1,
        f"Assessment Item {index + 1} of 40",
        "Read the full exchange, then decide whether the fallacy is on the left, on the right, or nowhere.",
        "Part I - Dialogue assessment items",
        styles,
    )
    y = draw_dialogue(pdf, item, y, styles)
    y -= 4
    y = draw_choices_row(pdf, y, styles)
    draw_paragraph(
        pdf,
        "Reminder: there is at most one fallacy in the exchange, and some items contain none at all.",
        styles["small"],
        MARGIN_X,
        y,
        CONTENT_WIDTH,
    )
    draw_footer(pdf)
    pdf.showPage()


def render_answer_page(pdf: canvas.Canvas, item: dict, index: int, styles: dict):
    y = draw_page_header(
        pdf,
        40 + index + 1,
        f"Answer Guide {index + 1} of 40",
        "Use the diagnosis and commentary below to see exactly why this dialogue counts as left formal, left informal, none, right informal, or right formal.",
        "Part II - Correct answers and commentary",
        styles,
    )
    y = draw_answer_summary(pdf, item, y, styles)
    y = draw_dialogue(pdf, item, y, styles)
    y -= 4
    y = draw_paragraph(pdf, "<b>Commentary</b>", styles["body_bold"], MARGIN_X, y, CONTENT_WIDTH)
    y -= 4
    y = draw_paragraph(pdf, item["explanation"], styles["body"], MARGIN_X, y, CONTENT_WIDTH)
    y -= 10
    if item.get("fallacySlug"):
        link = f'{LINK_BASE}{item["fallacySlug"]}/'
        draw_paragraph(
            pdf,
            f'<b>LogFall reference</b>: <link href="{link}" color="#145a8d">{link}</link>',
            styles["link"],
            MARGIN_X,
            y,
            CONTENT_WIDTH,
        )
    else:
        draw_paragraph(
            pdf,
            "<b>LogFall reference</b>: No fallacy page is linked here because this dialogue is a control item with no fallacy.",
            styles["small"],
            MARGIN_X,
            y,
            CONTENT_WIDTH,
        )
    draw_footer(pdf)
    pdf.showPage()


def build_pdf():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    bank = load_bank()
    styles = make_styles()
    pdf = canvas.Canvas(str(OUTPUT_PATH), pagesize=letter)
    pdf.setTitle("LogFall Dialogue Assessment")
    pdf.setAuthor("Phil Stilwell")
    pdf.setSubject("40-item dialogue assessment with answers and commentary")
    pdf.setCreator("LogFall build_dialogue_assessment_pdf.py")

    for index, item in enumerate(bank):
        render_question_page(pdf, item, index, styles)
    for index, item in enumerate(bank):
        render_answer_page(pdf, item, index, styles)

    pdf.save()
    print(OUTPUT_PATH)


if __name__ == "__main__":
    build_pdf()
