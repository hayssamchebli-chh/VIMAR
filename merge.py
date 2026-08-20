"""Build the merged pack: a clickable table of contents, then one cover page
plus datasheet per item.

Mirrors the layout used by the Lightning datasheet pack so the two tools
produce packs that look alike.
"""

import io
import os
import re
import time

from pypdf import PdfReader, PdfWriter
from pypdf.annotations import Link
from pypdf.generic import Fit
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas as pdf_canvas

HERE = os.path.dirname(os.path.abspath(__file__))
COVER_TEMPLATE_PATH = os.path.join(HERE, "item_type_template.pdf")
TOC_LOGO_PATH = os.path.join(HERE, "toc_logo.png")

COVER_TEXT_COLOR = "#1F4EA1"
COVER_TEXT_X = 42
COVER_TEXT_TOP_OFFSET = 170
COVER_TEXT_MAX_WIDTH = 340
COVER_TEXT_FONT = "Helvetica-Bold"
COVER_TEXT_FONT_SIZE = 34
COVER_TEXT_MIN_FONT_SIZE = 18

TOC_ACCENT_COLOR = "#1F4EA1"
TOC_TITLE_COLOR = "#102033"
TOC_DOTS_COLOR = "#9AA7B5"
TOC_MARGIN_X = 48
TOC_ENTRY_SPACING = 28
TOC_ENTRIES_FIRST_PAGE = 18
TOC_ENTRIES_LATER_PAGES = 22


def to_pdf_text(value: str) -> str:
    """Make text safe for the standard PDF fonts."""
    text = str(value or "")
    for source, target in (
        ("丨", " | "), ("｜", " | "), ("·", "-"),
        ("–", "-"), ("—", "-"), ("×", "x"),
        ("’", "'"), ("“", '"'), ("”", '"'),
    ):
        text = text.replace(source, target)
    text = text.encode("latin-1", "ignore").decode("latin-1")
    return re.sub(r"\s+", " ", text).strip()


def load_cover_template() -> bytes | None:
    try:
        with open(COVER_TEMPLATE_PATH, "rb") as f:
            return f.read()
    except Exception:
        return None


def build_type_overlay(type_text: str, page_width: float, page_height: float) -> bytes:
    """Draw the item type in the blank space under the logo."""
    buffer = io.BytesIO()
    overlay = pdf_canvas.Canvas(buffer, pagesize=(page_width, page_height))
    overlay.setFillColor(HexColor(COVER_TEXT_COLOR))

    def wrap(font_size: int) -> list[str]:
        lines, current = [], ""
        for word in to_pdf_text(type_text).split():
            candidate = f"{current} {word}".strip()
            if overlay.stringWidth(candidate, COVER_TEXT_FONT, font_size) <= COVER_TEXT_MAX_WIDTH:
                current = candidate
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)
        return lines

    font_size = COVER_TEXT_FONT_SIZE
    lines = wrap(font_size)
    while font_size > COVER_TEXT_MIN_FONT_SIZE and len(lines) > 3:
        font_size -= 2
        lines = wrap(font_size)

    overlay.setFont(COVER_TEXT_FONT, font_size)
    y = page_height - COVER_TEXT_TOP_OFFSET
    for line in lines:
        overlay.drawString(COVER_TEXT_X, y, line)
        y -= font_size * 1.3

    overlay.save()
    return buffer.getvalue()


def build_cover_page(template_bytes: bytes, type_text: str):
    """The template page, with the item's type written on it when given."""
    page = PdfReader(io.BytesIO(template_bytes)).pages[0]

    if type_text:
        overlay = build_type_overlay(
            type_text, float(page.mediabox.width), float(page.mediabox.height)
        )
        page.merge_page(PdfReader(io.BytesIO(overlay)).pages[0])

    return page


def toc_pages_needed(entry_count: int) -> int:
    if entry_count <= TOC_ENTRIES_FIRST_PAGE:
        return 1
    remaining = entry_count - TOC_ENTRIES_FIRST_PAGE
    return 1 + -(-remaining // TOC_ENTRIES_LATER_PAGES)


def build_toc_pdf(entries: list[dict], page_width: float, page_height: float):
    """Draw the contents pages; returns (pdf_bytes, clickable boxes)."""
    buffer = io.BytesIO()
    toc = pdf_canvas.Canvas(buffer, pagesize=(page_width, page_height))

    accent = HexColor(TOC_ACCENT_COLOR)
    title_color = HexColor(TOC_TITLE_COLOR)
    dots_color = HexColor(TOC_DOTS_COLOR)

    number_x = TOC_MARGIN_X
    title_x = TOC_MARGIN_X + 34
    page_num_right = page_width - TOC_MARGIN_X
    max_title_width = page_num_right - title_x - 60

    def truncate(text: str, size: float) -> str:
        if toc.stringWidth(text, "Helvetica-Bold", size) <= max_title_width:
            return text
        while text and toc.stringWidth(text + "...", "Helvetica-Bold", size) > max_title_width:
            text = text[:-1]
        return text.rstrip() + "..."

    def first_header() -> float:
        y_top = page_height - 52
        try:
            from reportlab.lib.utils import ImageReader

            logo = ImageReader(TOC_LOGO_PATH)
            logo_w, logo_h = logo.getSize()
            draw_h = 26
            toc.drawImage(logo, TOC_MARGIN_X, y_top - draw_h,
                          width=logo_w * draw_h / logo_h, height=draw_h, mask="auto")
        except Exception:
            pass

        title_y = y_top - 64
        toc.setFillColor(title_color)
        toc.setFont("Helvetica-Bold", 27)
        toc.drawString(TOC_MARGIN_X, title_y, "Table of Contents")
        toc.setFillColor(accent)
        toc.rect(TOC_MARGIN_X, title_y - 14, 64, 4, stroke=0, fill=1)
        return title_y - 52

    def later_header() -> float:
        toc.setFillColor(dots_color)
        toc.setFont("Helvetica", 11)
        toc.drawString(TOC_MARGIN_X, page_height - 56, "Table of Contents (continued)")
        toc.setFillColor(accent)
        toc.rect(TOC_MARGIN_X, page_height - 64, 42, 2.6, stroke=0, fill=1)
        return page_height - 100

    boxes = []
    toc_page_index = 0
    y = first_header()
    capacity = TOC_ENTRIES_FIRST_PAGE
    drawn = 0

    for position, entry in enumerate(entries, start=1):
        if drawn >= capacity:
            toc.showPage()
            toc_page_index += 1
            y = later_header()
            capacity = TOC_ENTRIES_LATER_PAGES
            drawn = 0

        title = truncate(to_pdf_text(entry["title"]), 12.5)
        page_label = str(entry["target_page"] + 1)

        toc.setFillColor(accent)
        toc.setFont("Helvetica-Bold", 10.5)
        toc.drawString(number_x, y, f"{position:02d}")

        toc.setFillColor(title_color)
        toc.setFont("Helvetica-Bold", 12.5)
        toc.drawString(title_x, y, title)

        toc.setFont("Helvetica-Bold", 11.5)
        toc.setFillColor(accent)
        toc.drawRightString(page_num_right, y, page_label)

        title_end = title_x + toc.stringWidth(title, "Helvetica-Bold", 12.5) + 8
        num_start = page_num_right - toc.stringWidth(page_label, "Helvetica-Bold", 11.5) - 8
        if num_start > title_end + 12:
            toc.setFillColor(dots_color)
            toc.setFont("Helvetica", 10)
            x = title_end
            step = toc.stringWidth(".", "Helvetica", 10) + 3.2
            while x < num_start:
                toc.drawString(x, y + 1, ".")
                x += step

        boxes.append({
            "page": toc_page_index,
            "rect": (TOC_MARGIN_X - 6, y - 8, page_num_right + 6, y + 14),
            "target": entry["target_page"],
        })

        y -= TOC_ENTRY_SPACING
        drawn += 1

    toc.save()
    return buffer.getvalue(), boxes


def merge_pack(items: list[dict], with_covers: bool = True, with_toc: bool = True) -> bytes:
    """Merge downloaded datasheets into one pack.

    items: [{"code", "type", "title", "content"}] in pack order.
    """
    prepared = []
    for item in items:
        if not item.get("content"):
            continue
        prepared.append((item, PdfReader(io.BytesIO(item["content"]))))

    if not prepared:
        return b""

    template_bytes = load_cover_template() if with_covers else None
    cover_pages = 1 if template_bytes else 0

    if template_bytes:
        template_page = PdfReader(io.BytesIO(template_bytes)).pages[0]
        page_width = float(template_page.mediabox.width)
        page_height = float(template_page.mediabox.height)
    else:
        first = prepared[0][1].pages[0]
        page_width = float(first.mediabox.width)
        page_height = float(first.mediabox.height)

    toc_page_count = toc_pages_needed(len(prepared)) if with_toc else 0

    entries = []
    cursor = toc_page_count
    for item, reader in prepared:
        title = (item.get("type") or "").strip() or item.get("title") or item["code"]
        entries.append({"title": title, "target_page": cursor})
        cursor += cover_pages + len(reader.pages)

    writer = PdfWriter()
    boxes = []

    if with_toc:
        toc_bytes, boxes = build_toc_pdf(entries, page_width, page_height)
        for page in PdfReader(io.BytesIO(toc_bytes)).pages:
            writer.add_page(page)

    for item, reader in prepared:
        if template_bytes:
            writer.add_page(build_cover_page(template_bytes, (item.get("type") or "").strip()))
        for page in reader.pages:
            writer.add_page(page)

    for box in boxes:
        writer.add_annotation(
            page_number=box["page"],
            annotation=Link(rect=box["rect"], target_page_index=box["target"],
                            fit=Fit(fit_type="/Fit")),
        )

    for entry in entries:
        writer.add_outline_item(entry["title"], entry["target_page"])

    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()
