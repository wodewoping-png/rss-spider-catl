# -*- coding: utf-8 -*-
"""
Pick the latest weekly CSV in output/weekly named like:
  news_with_abstract_2025-12-22.csv
and generate a Word file with the SAME base name:
  news_with_abstract_2025-12-22.docx
saved back to output/weekly/

Dependencies:
    pip install python-docx
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


# ----------------------------
# Helpers: hyperlink + styling
# ----------------------------
def add_hyperlink(paragraph, url: str, text: str, color_hex="1155CC", underline=True):
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
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")


def add_divider_line(doc: Document):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
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
    return (s or "").strip()


def is_truthy_flag(s: str) -> bool:
    s = (s or "").strip().lower()
    return s in {"true", "1", "yes", "y", "t"}


# ----------------------------
# Abstract cleanup + subscript
# ----------------------------
def strip_leading_abstract(text: str) -> str:
    """Remove leading 'Abstract' (case-insensitive) at the very beginning."""
    if not text:
        return text
    s = text.lstrip()
    s = re.sub(r"^(abstract)\s*[:.\-–—]*\s*", "", s, flags=re.IGNORECASE)
    return s


def abstract_to_runs(abstract: str):
    """
    Convert abstract into chunks (text, is_subscript).
    Rule: any line that becomes digits-only after strip() -> subscript.
    Collapses line breaks into spaces.
    """
    if not abstract:
        return [("(empty)", False)]

    abstract = strip_leading_abstract(abstract)

    text = abstract.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")

    chunks = []
    pending_space = False

    for line in lines:
        s = line.strip()
        if s == "":
            pending_space = True
            continue

        if pending_space and chunks:
            chunks.append((" ", False))
            pending_space = False

        if s.isdigit():
            chunks.append((s, True))  # attach directly as subscript
        else:
            if chunks:
                prev_text, prev_sub = chunks[-1]
                if not prev_sub and prev_text and not prev_text.endswith(" "):
                    if prev_text[-1] not in {"-", "−", "/"}:
                        chunks.append((" ", False))
            chunks.append((s, False))

    # de-dup spaces
    cleaned = []
    for t, sub in chunks:
        if cleaned and t == " " and cleaned[-1][0] == " ":
            continue
        cleaned.append((t, sub))

    return cleaned


def add_abstract_with_subscripts(paragraph, abstract: str):
    for t, is_sub in abstract_to_runs(abstract):
        run = paragraph.add_run(t)
        if is_sub:
            run.font.subscript = True


# ----------------------------
# CSV picker (prefix + date + mtime tie-break)
# ----------------------------
def _extract_date_from_name(filename: str):
    """
    Expect formats like:
      news_with_abstract_2025-12-22.csv
    Also tolerates 20251222 or 2025_12_22.
    Returns datetime.date or None.
    """
    stem = Path(filename).stem
    m = re.search(r"(\d{4})[-_]?(\d{2})[-_]?(\d{2})", stem)
    if not m:
        return None
    y, mo, d = map(int, m.groups())
    try:
        return datetime(y, mo, d).date()
    except ValueError:
        return None


def pick_latest_weekly_csv(folder="output/weekly", prefix="news_with_abstract_") -> Path:
    """
    1) filter by startswith(prefix) and .csv
    2) pick newest by embedded date; tie-breaker by mtime
    3) if no embedded date, fallback to newest mtime
    """
    folder = Path(folder)
    if not folder.exists():
        raise FileNotFoundError(f"Folder not found: {folder.resolve()}")

    candidates = [p for p in folder.glob("*.csv") if p.name.startswith(prefix)]
    if not candidates:
        raise FileNotFoundError(f"No CSV starting with '{prefix}' in {folder.resolve()}")

    with_dates = []
    without_dates = []
    for p in candidates:
        d = _extract_date_from_name(p.name)
        if d is not None:
            with_dates.append((d, p))
        else:
            without_dates.append(p)

    if with_dates:
        with_dates.sort(key=lambda x: (x[0], x[1].stat().st_mtime), reverse=True)
        return with_dates[0][1]

    return max(without_dates, key=lambda p: p.stat().st_mtime)


# ----------------------------
# Main conversion
# ----------------------------
def csv_to_word(input_csv: str, output_docx: str, report_title: str = "Literature Digest", encoding: str = "utf-8-sig"):
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

        # Record header
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

        p.add_run(f"[{idx}] ").bold = True
        if link:
            add_hyperlink(p, link, title or "(no title)")
        else:
            p.add_run(title or "(no title)").bold = True

        # Meta
        meta = doc.add_paragraph()
        meta.paragraph_format.space_before = Pt(0)
        meta.paragraph_format.space_after = Pt(6)
        meta.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

        mr = meta.add_run("Source: ")
        mr.bold = True
        meta.add_run(source or "-")

        meta.add_run("    ")
        mr = meta.add_run("Pub date: ")
        mr.bold = True
        meta.add_run(normalize_datetime_str(pub_date) or "-")

        if doi:
            meta.add_run("    ")
            mr = meta.add_run("DOI: ")
            mr.bold = True
            add_hyperlink(meta, doi_to_url(doi), doi)

        # Notes:
        # - keep abstract_source if present
        # - do NOT output must_have_abstract=False; only show it if True-ish
        extra_bits = []
        if abstract_source:
            extra_bits.append(f"abstract_source={abstract_source}")
        if is_truthy_flag(must_have_abstract):
            extra_bits.append("must_have_abstract=True")

        if extra_bits:
            extra = doc.add_paragraph()
            extra.paragraph_format.space_before = Pt(0)
            extra.paragraph_format.space_after = Pt(6)
            er = extra.add_run("Notes: ")
            er.bold = True
            extra.add_run("; ".join(extra_bits))

        # Abstract
        abs_p = doc.add_paragraph()
        abs_p.paragraph_format.space_before = Pt(0)
        abs_p.paragraph_format.space_after = Pt(10)
        abs_p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        abs_p.paragraph_format.line_spacing = 1.15

        abs_p.add_run("Abstract: ").bold = True
        add_abstract_with_subscripts(abs_p, abstract)

        if idx != len(rows):
            add_divider_line(doc)

    doc.save(output_docx)


if __name__ == "__main__":
    weekly_dir = Path("output/weekly")
    csv_path = pick_latest_weekly_csv(folder=str(weekly_dir), prefix="news_with_abstract_")

    # Output .docx in the SAME folder with same base name:
    docx_path = csv_path.with_suffix(".docx")

    TITLE = "Tech Tracking Digest"
    csv_to_word(str(csv_path), str(docx_path), report_title=TITLE)

    print(f"Picked CSV:  {csv_path}")
    print(f"Wrote DOCX: {docx_path}")
