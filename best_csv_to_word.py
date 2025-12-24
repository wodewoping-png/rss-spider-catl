# -*- coding: utf-8 -*-
"""
Convert a CSV (title, link, published, source, pub_date, doi, abstract, ...)
into a nicely formatted Word document (.docx).

Enhancements in this version:
1) Convert "digit-only lines" inside abstract into subscript in Word.
   Example:
       with Zn
           2
           ⁺-imidazole ...
   -> Zn₂⁺-imidazole ...

2) Remove notes for must_have_abstract=False (actually: never output must_have_abstract at all unless it's TRUE-ish)
"""

import csv
import re
from pathlib import Path
from datetime import datetime

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


INPUT_CSV = "news_with_abstract_2025-12-22.csv"
OUTPUT_DOCX = "news_with_abstract.docx"
TITLE = "Tech Tracking Digest"


def strip_leading_abstract(text: str) -> str:
    """
    Remove leading 'Abstract' (case-insensitive) from the beginning of abstract.
    Examples removed:
      Abstract
      Abstract:
      ABSTRACT –
      Abstract—
      Abstract.
    """
    if not text:
        return text

    s = text.lstrip()

    # Regex: start of string + "abstract" + optional punctuation
    s = re.sub(
        r'^(abstract)\s*[:.\-–—]*\s*',
        '',
        s,
        flags=re.IGNORECASE
    )
    return s

# ----------------------------
# Helpers: hyperlink + styling
# ----------------------------
def add_hyperlink(paragraph, url: str, text: str, color_hex="1155CC", underline=True):
    """Add a clickable hyperlink to a paragraph in python-docx."""
    if not url:
        paragraph.add_run(text)
        return

    part = paragraph.part
    r_id = part.relate_to(
        url,
        reltype="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")

    c = OxmlElement("w:color")
    c.set(qn("w:val"), color_hex)
    r_pr.append(c)

    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single" if underline else "none")
    r_pr.append(u)

    new_run.append(r_pr)

    text_elem = OxmlElement("w:t")
    text_elem.text = text
    new_run.append(text_elem)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def set_doc_default_style(doc: Document, font_name="Calibri", font_size_pt=11):
    style = doc.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(font_size_pt)
    # Chinese font fallback
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")


def add_divider_line(doc: Document):
    p = doc.add_paragraph()
    p_format = p.paragraph_format
    p_format.space_before = Pt(6)
    p_format.space_after = Pt(6)
    run = p.add_run("—" * 70)
    run.font.size = Pt(9)


def safe_get(row: dict, key: str) -> str:
    return (row.get(key) or "").strip()


def doi_to_url(doi: str) -> str:
    doi = (doi or "").strip()
    if not doi:
        return ""
    if doi.lower().startswith("http"):
        return doi
    return f"https://doi.org/{doi}"


def normalize_datetime_str(s: str) -> str:
    s = (s or "").strip()
    return s


def is_truthy_flag(s: str) -> bool:
    """Interpret CSV bool-like strings."""
    s = (s or "").strip().lower()
    return s in {"true", "1", "yes", "y", "t"}


# ----------------------------
# Abstract processing
# ----------------------------
def abstract_to_runs(abstract: str):
    """
    Convert abstract text into a list of (text, is_subscript) chunks.

    Rule: Any line that becomes digits-only after strip() will be subscript.
    Everything else becomes normal text. We also collapse line breaks into spaces
    (to avoid weird forced line wrapping from RSS HTML extraction).
    """
    if not abstract:
        return [("(empty)", False)]

    # 🔹 删除开头的 "Abstract"
    abstract = strip_leading_abstract(abstract)

    text = abstract.replace("\r\n", "\n").replace("\r", "\n")

    # Split by lines and classify
    lines = text.split("\n")

    chunks = []
    pending_space = False

    for line in lines:
        s = line.strip()
        if s == "":
            # Treat blank lines as a space boundary, but don't spam spaces
            pending_space = True
            continue

        # If we previously saw a blank line, ensure a space separation
        if pending_space and chunks:
            chunks.append((" ", False))
            pending_space = False

        # If this line is purely digits: subscript it
        if s.isdigit():
            # Usually this digit should attach to previous token (e.g., "Zn" + "2")
            # So we do NOT force leading space here; just add directly.
            chunks.append((s, True))
        else:
            # For normal text lines: if previous chunk is normal text and doesn't end with a space,
            # add a space before appending, to collapse line breaks into readable text.
            if chunks:
                prev_text, prev_sub = chunks[-1]
                if not prev_sub and not prev_text.endswith(" "):
                    # If previous ended with hyphen or slash etc., you might NOT want a space,
                    # but for safety: only skip space if prev ends with "--/" (common joiners).
                    if prev_text and prev_text[-1] not in {"-", "-", "/", "−"}:
                        chunks.append((" ", False))
            chunks.append((s, False))

    # Post-clean: remove redundant multiple spaces
    cleaned = []
    for t, sub in chunks:
        if not cleaned:
            cleaned.append((t, sub))
            continue
        if t == " " and cleaned[-1][0] == " ":
            continue
        cleaned.append((t, sub))

    return cleaned


def add_abstract_with_subscripts(paragraph, abstract: str):
    """Write abstract into a paragraph, converting digit-only lines into subscript runs."""
    chunks = abstract_to_runs(abstract)
    for t, is_sub in chunks:
        run = paragraph.add_run(t)
        if is_sub:
            run.font.subscript = True


# ----------------------------
# Main conversion
# ----------------------------
def csv_to_word(
    input_csv: str,
    output_docx: str,
    report_title: str = "Literature Digest",
    encoding: str = "utf-8-sig",
):
    input_path = Path(input_csv)
    if not input_path.exists():
        raise FileNotFoundError(f"CSV not found: {input_path.resolve()}")

    doc = Document()
    set_doc_default_style(doc, font_name="Calibri", font_size_pt=11)

    # Page setup
    section = doc.sections[0]
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)

    # Title
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title_p.add_run(report_title)
    r.bold = True
    r.font.size = Pt(18)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_after = Pt(14)
    sub_p.add_run(datetime.now().strftime("%Y-%m-%d")).italic = True

    # Read CSV
    with open(input_path, "r", encoding=encoding, newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    if not rows:
        doc.add_paragraph("No records found in CSV.")
        doc.save(output_docx)
        return

    for idx, row in enumerate(rows, start=1):
        title = safe_get(row, "title")
        link = safe_get(row, "link")
        source = safe_get(row, "source")
        pub_date = safe_get(row, "pub_date") or safe_get(row, "published")
        doi = safe_get(row, "doi")
        abstract = safe_get(row, "abstract")
        abstract_source = safe_get(row, "abstract_source")
        must_have_abstract = safe_get(row, "must_have_abstract")

        # Record header: [#] Title (hyperlink)
        p = doc.add_paragraph()
        p_format = p.paragraph_format
        p_format.space_before = Pt(8)
        p_format.space_after = Pt(4)
        p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

        num_run = p.add_run(f"[{idx}] ")
        num_run.bold = True

        if link:
            add_hyperlink(p, link, title or "(no title)")
        else:
            t_run = p.add_run(title or "(no title)")
            t_run.bold = True

        # Meta info
        meta = doc.add_paragraph()
        meta_format = meta.paragraph_format
        meta_format.space_before = Pt(0)
        meta_format.space_after = Pt(6)
        meta_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

        meta_run = meta.add_run("Source: ")
        meta_run.bold = True
        meta.add_run(source or "-")

        meta.add_run("    ")

        meta_run = meta.add_run("Pub date: ")
        meta_run.bold = True
        meta.add_run(normalize_datetime_str(pub_date) or "-")

        if doi:
            meta.add_run("    ")
            meta_run = meta.add_run("DOI: ")
            meta_run.bold = True
            add_hyperlink(meta, doi_to_url(doi), doi)

        # Notes: keep abstract_source if present
        # MUST: remove must_have_abstract=False (and generally don't show it unless TRUE-ish)
        extra_bits = []
        if abstract_source:
            extra_bits.append(f"abstract_source={abstract_source}")
        if is_truthy_flag(must_have_abstract):
            extra_bits.append("must_have_abstract=True")

        if extra_bits:
            extra = doc.add_paragraph()
            extra.paragraph_format.space_before = Pt(0)
            extra.paragraph_format.space_after = Pt(6)
            extra_run = extra.add_run("Notes: ")
            extra_run.bold = True
            extra.add_run("; ".join(extra_bits))

        # Abstract
        abs_p = doc.add_paragraph()
        abs_p_format = abs_p.paragraph_format
        abs_p_format.space_before = Pt(0)
        abs_p_format.space_after = Pt(10)
        abs_p_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        abs_p_format.line_spacing = 1.15

        abs_label = abs_p.add_run("Abstract: ")
        abs_label.bold = True

        # Here: digit-only lines -> subscript
        add_abstract_with_subscripts(abs_p, abstract)

        # Divider
        if idx != len(rows):
            add_divider_line(doc)

    doc.save(output_docx)

csv_to_word(INPUT_CSV, OUTPUT_DOCX, report_title=TITLE)
print(f"Done -> {OUTPUT_DOCX}")
