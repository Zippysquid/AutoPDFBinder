#!/usr/bin/env python
import os
import sys
import subprocess
import time
import logging
import shutil
from pathlib import Path
from typing import List, Tuple, Dict
from time import sleep

# For Word-to-PDF conversion via COM automation
import comtypes
from comtypes.client import CreateObject

# Required packages:
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
import PyPDF2
import fitz  # PyMuPDF

###############################################################################
# Configuration & Constants
###############################################################################
SCRIPT_DIR = Path(__file__).parent.absolute()
OUTPUT_DIR = SCRIPT_DIR / "output"
LOG_FILE = SCRIPT_DIR / "script_log.txt"
FINAL_PDF = SCRIPT_DIR / "final_output.pdf"

BATES_START = 1         # Final PDF page numbering will start at 001
BATES_FONT_SIZE = 14
TIMEOUT_BASE = 1.0

# Design Configuration
DESIGN_CONFIG = {
    'primary_color': RGBColor(0, 51, 102),      # Navy blue
    'accent_color': RGBColor(200, 200, 200),    # Light gray
    'header_bg': RGBColor(245, 245, 245),       # Very light gray
    'alt_row_color': RGBColor(250, 250, 250),   # Almost white
    'text_gray': RGBColor(100, 100, 100),       # Medium gray
    'cover_style': 'modern',  # 'modern', 'classic', 'minimal'
    'company_name': 'Document Archive',
    'department': 'Document Management System',
    'show_metadata': True,
    'use_icons': True  # Enable/disable icons in TOC
}

###############################################################################
# Logging Setup
###############################################################################
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [LOG] %(levelname)s: %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8", mode='w'),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.info
log_error = logging.error
log_debug = logging.debug

###############################################################################
# Dependency Installation
###############################################################################
def install_missing_packages() -> None:
    required_packages = {
        "python-docx": "docx",
        "pypdf2": "PyPDF2",
        "pymupdf": "fitz",
        "comtypes": "comtypes"
    }
    log(f"Checking dependencies with Python: {sys.executable}")
    missing = False
    for pkg_name, import_name in required_packages.items():
        try:
            __import__(import_name)
            log(f"Package {pkg_name} ({import_name}) is installed.")
        except ImportError:
            log(f"Installing {pkg_name}...")
            try:
                subprocess.run(
                    [sys.executable, "-m", "pip", "install", "--user", pkg_name],
                    capture_output=True, text=True, check=True, timeout=300
                )
                log(f"Installed {pkg_name}")
            except subprocess.CalledProcessError as e:
                log_error(f"Failed to install {pkg_name}: {e.stderr}")
                missing = True
    if missing:
        log_error("Critical: Some packages failed to install.")
        sys.exit(1)

install_missing_packages()

###############################################################################
# Helper Functions
###############################################################################
def remove_table_borders(table: docx.table.Table) -> None:
    """Remove table-level borders from a docx table."""
    for border in table._element.xpath('.//w:tblBorders'):
        border.getparent().remove(border)

def remove_cell_borders(table: docx.table.Table) -> None:
    """Remove cell-level borders from a docx table."""
    for row in table.rows:
        for cell in row.cells:
            for border in cell._tc.xpath('.//w:tcBorders'):
                border.getparent().remove(border)

def set_cell_margins(cell, top=0, left=0, bottom=0, right=0):
    """
    Set margins (in dxa units) for a table cell.
    1 dxa = 1/20th of a point.
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn('w:tcMar'))
    if tcMar is None:
        tcMar = OxmlElement('w:tcMar')
        tcPr.append(tcMar)
    for margin, value in (('top', top), ('left', left), ('bottom', bottom), ('right', right)):
        element = tcMar.find(qn(f'w:{margin}'))
        if element is None:
            element = OxmlElement(f'w:{margin}')
            tcMar.append(element)
        element.set(qn('w:w'), str(value))
        element.set(qn('w:type'), 'dxa')

def set_cell_background(cell, color: RGBColor):
    """Set background color for a table cell."""
    # Convert RGBColor to hex
    hex_color = '%02x%02x%02x' % (color[0], color[1], color[2])

    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_table_borders(table, color: RGBColor = None, width: int = 1):
    """Add subtle borders to a table."""
    if color is None:
        color = RGBColor(200, 200, 200)

    hex_color = '%02x%02x%02x' % (color[0], color[1], color[2])

    # Add table borders
    tbl = table._element
    tblPr = tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')

    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), str(width * 8))  # width in eighths of a point
        border.set(qn('w:color'), hex_color)
        tblBorders.append(border)

    tblPr.append(tblBorders)

def get_file_type(filename: str) -> str:
    """Determine file type from extension."""
    ext = Path(filename).suffix.lower()
    types = {
        '.pdf': 'PDF Document',
        '.docx': 'Word Document',
        '.doc': 'Word Document (Legacy)',
        '.xlsx': 'Excel Spreadsheet',
        '.pptx': 'PowerPoint Presentation'
    }
    return types.get(ext, 'Document')

def get_file_size_if_available(filepath: Path) -> str:
    """Get formatted file size if file exists."""
    try:
        if filepath.exists():
            size = filepath.stat().st_size
            for unit in ['B', 'KB', 'MB', 'GB']:
                if size < 1024.0:
                    return f"{size:.1f} {unit}"
                size /= 1024.0
    except:
        pass
    return "N/A"

def apply_document_styles(doc: docx.Document) -> None:
    """Apply consistent styling (font, margins, spacing) to the document."""
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(8)
    style.paragraph_format.line_spacing = 1.15
    section = doc.sections[0]
    section.page_height = Inches(11)
    section.page_width = Inches(8.5)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    """Convert a DOCX file to PDF using Microsoft Word COM automation."""
    log(f"Converting {docx_path} -> {pdf_path}")
    if not docx_path.exists():
        raise FileNotFoundError(f"{docx_path} not found.")
    word = None
    try:
        word = CreateObject('Word.Application')
        word.DisplayAlerts = 0  # Disable alerts (e.g., Save As prompts)
        word.Visible = False
        doc = word.Documents.Open(str(docx_path), ReadOnly=True)
        doc.SaveAs(str(pdf_path), FileFormat=17)
        doc.Close()
        time.sleep(TIMEOUT_BASE)
    except Exception as e:
        log_error(f"Error converting {docx_path}: {e}")
        raise
    finally:
        if word is not None:
            try:
                word.NormalTemplate.Saved = True  # Prevent normal.dot update prompt
            except Exception as e:
                log_debug(f"Error setting NormalTemplate.Saved: {e}")
            word.Quit()
    if not pdf_path.exists():
        raise FileNotFoundError(f"Conversion failed: {pdf_path} not created.")
    log(f"Converted {docx_path.name} -> {pdf_path.name}")

def pdf_page_count(pdf_path: Path) -> int:
    """Return the number of pages in a PDF file."""
    try:
        reader = PyPDF2.PdfReader(str(pdf_path))
        return len(reader.pages)
    except Exception as e:
        log_error(f"Error counting pages in {pdf_path}: {e}")
        return 0

def merge_pdfs(pdf_paths: List[Path], final_pdf: Path) -> None:
    """Merge PDFs in the provided order into a final PDF."""
    from PyPDF2 import PdfMerger
    log(f"Merging PDFs into {final_pdf.name}")
    merger = PdfMerger()
    for p in pdf_paths:
        if p.exists():
            pages = pdf_page_count(p)
            if pages > 0:
                log(f"Appending {p.name} ({pages} pages)")
                merger.append(str(p))
            else:
                log(f"Skipping {p.name}, 0 pages.")
        else:
            log_error(f"File missing: {p.name}")
    merger.write(str(final_pdf))
    merger.close()
    if not final_pdf.exists():
        raise FileNotFoundError(f"Merged PDF not created: {final_pdf}")
    log(f"Merged PDF saved: {final_pdf.name}")

def apply_bates_numbering(pdf_path: Path, start_number: int = BATES_START, font_size: int = BATES_FONT_SIZE) -> None:
    """Apply sequential Bates numbering with a subtle background to each page."""
    log(f"Applying Bates numbering to {pdf_path.name}")
    doc = fitz.open(str(pdf_path))
    for i, page in enumerate(doc):
        bates_num = f"{start_number + i:03d}"
        rect = page.rect
        x = rect.width - 80
        y = rect.height - 40
        page.draw_rect(
            fitz.Rect(x - 10, y - 25, x + 60, y + 5),
            color=(0.9, 0.9, 0.9),
            fill=(0.9, 0.9, 0.9)
        )
        page.insert_text((x, y), bates_num,
                         fontname="Helvetica",
                         fontsize=font_size,
                         color=(0, 0, 0))
    temp_pdf = pdf_path.with_suffix('.temp.pdf')
    doc.save(str(temp_pdf))
    doc.close()
    temp_pdf.replace(pdf_path)
    log(f"Bates numbering applied to {pdf_path.name}")

def add_pdf_bookmarks(pdf_path: Path, toc_list: List[List]) -> None:
    """
    Add bookmarks (outline) to the merged PDF.
    Each entry in toc_list is [level, title, page_number] (1-indexed).
    """
    log("Adding PDF bookmarks (outline)...")
    try:
        doc = fitz.open(str(pdf_path))
        doc.set_toc(toc_list)
        new_toc = doc.get_toc()
        log_debug(f"New TOC: {new_toc}")

        # Always do a full save to avoid incremental save issues
        # Save to a temporary file first
        temp_pdf = pdf_path.with_suffix('.temp.pdf')
        doc.save(str(temp_pdf), garbage=4, deflate=True)
        doc.close()

        # Replace the original with the temp file
        temp_pdf.replace(pdf_path)
        log("PDF bookmarks added.")
    except Exception as e:
        log(f"Warning: Could not add PDF bookmarks: {e}")
        log("The PDF has been created successfully but without bookmarks.")
        # Don't fail the entire process just because bookmarks couldn't be added
        if 'doc' in locals():
            doc.close()

###############################################################################
# Document Creation Functions
###############################################################################
def create_cover_page(cover_docx: Path, number: str, file_name: str) -> None:
    """
    Generate an enhanced cover page with professional design elements.
    """
    log(f"Creating cover page: {cover_docx.name}")
    doc = docx.Document()
    apply_document_styles(doc)

    # Add a styled header section
    header_table = doc.add_table(rows=1, cols=1)
    header_table.autofit = False
    header_cell = header_table.cell(0, 0)
    header_cell.width = Inches(6.5)

    # Set header background
    set_cell_background(header_cell, DESIGN_CONFIG['primary_color'])
    set_cell_margins(header_cell, top=200, bottom=200)

    # Add organization name in header - changed to show document number
    header_para = header_cell.paragraphs[0]
    header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_run = header_para.add_run(f"DOCUMENT {number}")
    header_run.font.color.rgb = RGBColor(255, 255, 255)
    header_run.font.size = Pt(14)
    header_run.font.bold = True
    header_run.font.name = "Calibri"

    # Add vertical spacing
    for _ in range(5):
        doc.add_paragraph()

    # Main document title with better typography (no truncation) (no truncation)
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    title_run = title_para.add_run(file_name)
    title_run.font.name = "Calibri"
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_para.paragraph_format.space_after = Pt(24)

    # Add file metadata in a styled box
    if DESIGN_CONFIG['show_metadata']:
        doc.add_paragraph()  # spacing

        # Create metadata table
        meta_table = doc.add_table(rows=3, cols=2)
        meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        meta_table.autofit = False

        # Style the metadata table with subtle borders
        set_table_borders(meta_table, color=DESIGN_CONFIG['accent_color'], width=1)

        # Get actual file path
        actual_file = Path(file_name) if Path(file_name).exists() else None

        # Add metadata rows
        metadata = [
            ("Document Type:", get_file_type(file_name)),
            ("File Size:", get_file_size_if_available(actual_file) if actual_file else "N/A"),
            ("Generated:", time.strftime("%B %d, %Y at %I:%M %p"))
        ]

        for i, (label, value) in enumerate(metadata):
            label_cell = meta_table.cell(i, 0)
            value_cell = meta_table.cell(i, 1)

            # Set cell widths
            label_cell.width = Inches(1.5)
            value_cell.width = Inches(3.5)

            # Style label
            label_para = label_cell.paragraphs[0]
            label_run = label_para.add_run(label)
            label_run.font.bold = True
            label_run.font.size = Pt(10)
            label_run.font.name = "Calibri"
            label_run.font.color.rgb = DESIGN_CONFIG['text_gray']

            # Style value
            value_para = value_cell.paragraphs[0]
            value_run = value_para.add_run(value)
            value_run.font.size = Pt(10)
            value_run.font.name = "Calibri"

            # Add padding to cells
            set_cell_margins(label_cell, top=100, bottom=100, left=100, right=50)
            set_cell_margins(value_cell, top=100, bottom=100, left=50, right=100)

    # Remove the footer section since it's not needed
    # Just keep the decorative line
    for _ in range(6):
        doc.add_paragraph()

    # Add a subtle bottom border
    bottom_line = doc.add_paragraph()
    bottom_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = bottom_line.add_run("_" * 60)
    line_run.font.color.rgb = DESIGN_CONFIG['accent_color']
    line_run.font.size = Pt(8)

    safe_save_docx(doc, cover_docx)

def create_contents_page(contents_docx: Path, all_items: List[Tuple[str, Path]], bates_map: Dict[str, int],
                         processed: Dict[str, Tuple[Path, int, Path, int]] = None,
                         real_files: List[Tuple[str, Path]] = None,
                         contents_pages: int = 0) -> None:
    """
    Create an enhanced table of contents with modern design elements.
    """
    log(f"Creating contents page: {contents_docx.name}")
    doc = docx.Document()
    apply_document_styles(doc)

    # Create a styled header section
    header_table = doc.add_table(rows=2, cols=1)
    header_table.autofit = False
    remove_table_borders(header_table)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Main title
    title_cell = header_table.cell(0, 0)
    title_cell.width = Inches(6.5)
    title_para = title_cell.paragraphs[0]
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run("TABLE OF CONTENTS")
    title_run.font.name = "Calibri"
    title_run.font.size = Pt(18)
    title_run.font.bold = True
    title_run.font.color.rgb = DESIGN_CONFIG['primary_color']

    # Subtitle with date and statistics
    subtitle_cell = header_table.cell(1, 0)
    subtitle_cell.width = Inches(6.5)
    subtitle_para = subtitle_cell.paragraphs[0]
    subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Calculate statistics
    total_docs = len([num for num, p in all_items if p.is_file()])
    total_folders = len([num for num, p in all_items if p.is_dir()])

    subtitle_text = f"{time.strftime('%B %d, %Y')} | {total_docs} Document"
    if total_docs != 1:
        subtitle_text += "s"
    if total_folders > 0:
        subtitle_text += f" in {total_folders} Folder"
        if total_folders != 1:
            subtitle_text += "s"

    subtitle_run = subtitle_para.add_run(subtitle_text)
    subtitle_run.font.size = Pt(10)
    subtitle_run.font.italic = True
    subtitle_run.font.color.rgb = DESIGN_CONFIG['text_gray']
    subtitle_run.font.name = "Calibri"

    # Add decorative divider
    doc.add_paragraph()

    # Create a thin line
    line_table = doc.add_table(rows=1, cols=1)
    line_table.autofit = False
    line_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    line_cell = line_table.cell(0, 0)
    line_cell.width = Inches(6.5)
    line_cell.height = Pt(1)
    set_cell_background(line_cell, DESIGN_CONFIG['accent_color'])
    remove_table_borders(line_table)

    doc.add_paragraph()  # spacing

    # Determine if we should use icons
    use_icons = DESIGN_CONFIG['use_icons']
    num_cols = 3 if use_icons else 2

    # Create main content table
    table = doc.add_table(rows=len(all_items) + 1, cols=num_cols)
    table.style = "Table Grid"
    remove_table_borders(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set column widths
    if use_icons:
        widths = [Inches(0.4), Inches(5.1), Inches(1.0)]  # Icon, Name, Page
    else:
        widths = [Inches(5.5), Inches(1.0)]  # Name, Page

    for col_idx, width in enumerate(widths):
        for cell in table.columns[col_idx].cells:
            cell.width = width

    # Style header row
    header_row = table.rows[0]
    if use_icons:
        header_texts = ["", "Document", "Page"]
    else:
        header_texts = ["Document", "Page"]

    for i, text in enumerate(header_texts):
        cell = header_row.cells[i]
        set_cell_background(cell, DESIGN_CONFIG['header_bg'])
        set_cell_margins(cell, top=100, bottom=100, left=100 if i == 0 else 50, right=100 if i == len(header_texts)-1 else 50)

        if text:  # Skip empty icon column header
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT if text == "Document" else WD_ALIGN_PARAGRAPH.RIGHT
            run = para.add_run(text)
            run.font.bold = True
            run.font.size = Pt(11)
            run.font.color.rgb = DESIGN_CONFIG['primary_color']
            run.font.name = "Calibri"

    # Add content rows with enhanced styling
    for idx, (num, path_obj) in enumerate(all_items):
        row_idx = idx + 1
        row = table.rows[row_idx]

        # Alternate row coloring for better readability
        if row_idx % 2 == 0:
            for cell in row.cells:
                set_cell_background(cell, DESIGN_CONFIG['alt_row_color'])

        col_offset = 0

        # Icon column (if enabled)
        if use_icons:
            icon_cell = row.cells[0]
            icon_para = icon_cell.paragraphs[0]
            icon_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_margins(icon_cell, top=50, bottom=50)

            if path_obj.is_dir():
                icon = "📁"
            else:
                ext = path_obj.suffix.lower()
                if ext == ".pdf":
                    icon = "📄"
                elif ext in [".docx", ".doc"]:
                    icon = "📝"
                else:
                    icon = "📎"

            icon_run = icon_para.add_run(icon)
            icon_run.font.size = Pt(11)
            col_offset = 1

        # Document name column with smart indentation
        name_cell = row.cells[col_offset]
        name_para = name_cell.paragraphs[0]
        set_cell_margins(name_cell, top=50, bottom=50, left=100, right=50)

        # Calculate indentation level
        indent_level = num.count(".")

        # Create visual hierarchy with indentation
        if indent_level > 0:
            # Add visual indent
            indent_text = "  " * indent_level + "└─ "
            indent_run = name_para.add_run(indent_text)
            indent_run.font.color.rgb = RGBColor(200, 200, 200)
            indent_run.font.name = "Courier New"
            indent_run.font.size = Pt(10)

        # Add document number
        num_run = name_para.add_run(f"{num}. ")
        num_run.font.name = "Calibri"
        num_run.font.size = Pt(11)
        num_run.font.color.rgb = DESIGN_CONFIG['text_gray']

        # Add document name (full name, no truncation)
        name_run = name_para.add_run(path_obj.name)
        name_run.font.name = "Calibri"
        name_run.font.size = Pt(11)

        if path_obj.is_dir():
            name_run.font.bold = True
            name_run.font.color.rgb = DESIGN_CONFIG['primary_color']

        # Page number column
        page_cell = row.cells[col_offset + 1]
        page_para = page_cell.paragraphs[0]
        page_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_cell_margins(page_cell, top=50, bottom=50, left=50, right=100)

        if num in bates_map and path_obj.is_file():
            page_run = page_para.add_run(f"{bates_map[num]:03d}")
            page_run.font.name = "Calibri"
            page_run.font.size = Pt(11)

    # Add footer section
    doc.add_paragraph()
    doc.add_paragraph()

    # Footer with total pages
    if bates_map:
        footer_para = doc.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Calculate actual total pages
        if processed:
            total_pages = contents_pages  # Start with TOC pages
            for num, _ in real_files:
                cover_pdf, cover_pages, file_pdf, file_pages = processed[num]
                total_pages += cover_pages + file_pages
        else:
            total_pages = max(bates_map.values()) if bates_map else 0
        footer_run = footer_para.add_run(f"Total Pages: {total_pages}")
        footer_run.font.size = Pt(9)
        footer_run.font.italic = True
        footer_run.font.color.rgb = DESIGN_CONFIG['text_gray']
        footer_run.font.name = "Calibri"

    safe_save_docx(doc, contents_docx)
    log(f"Contents page saved: {contents_docx.name}")

###############################################################################
# Main Execution Flow
###############################################################################
def cleanup_output_dir() -> None:
    """Clean up any leftover files from previous runs."""
    patterns = ["cover_*.docx", "cover_*.pdf", "file_*.pdf", "contents_dummy.*", "contents.docx", "contents.pdf"]
    for pattern in patterns:
        for temp in OUTPUT_DIR.glob(pattern):
            try:
                temp.unlink()
                log_debug(f"Cleaned up existing file: {temp}")
            except PermissionError:
                log(f"Warning: Could not delete {temp.name} (file may be open)")
            except Exception as e:
                log_debug(f"Could not delete {temp.name}: {e}")

def safe_save_docx(doc: docx.Document, path: Path, retry_count: int = 3) -> None:
    """Save a docx file with retry logic for permission errors."""
    for attempt in range(retry_count):
        try:
            doc.save(str(path))
            return
        except PermissionError:
            if attempt < retry_count - 1:
                log(f"Permission denied for {path.name}, retrying in 2 seconds...")
                time.sleep(2)
                # Try to delete the file if it exists
                if path.exists():
                    try:
                        path.unlink()
                    except:
                        pass
            else:
                raise

def main() -> None:
    try:
        OUTPUT_DIR.mkdir(exist_ok=True)
        log(f"Output directory ready: {OUTPUT_DIR}")

        # Clean up any existing temporary files
        cleanup_output_dir()

        all_items: List[Tuple[str, Path]] = []
        def scan_dir(dirpath: Path, prefix: str = ""):
            if OUTPUT_DIR in dirpath.parents or dirpath == OUTPUT_DIR:
                return
            files = sorted(
                [f for f in dirpath.iterdir() if f.is_file() and f.suffix.lower() in [".docx", ".pdf"] and f != FINAL_PDF],
                key=lambda x: x.name.lower()
            )
            i = 1
            for f in files:
                num = f"{prefix}{i}"
                all_items.append((num, f))
                i += 1
            subdirs = sorted(
                [d for d in dirpath.iterdir() if d.is_dir() and d != OUTPUT_DIR],
                key=lambda x: x.name.lower()
            )
            j = 1
            for d in subdirs:
                sub_prefix = f"{prefix}{j}."
                all_items.append((sub_prefix[:-1], d))
                j += 1
                scan_dir(d, prefix=sub_prefix)
        scan_dir(SCRIPT_DIR, "")
        real_files: List[Tuple[str, Path]] = [
            (num, p) for num, p in all_items
            if p.is_file() and p.suffix.lower() in [".docx", ".pdf"] and p != FINAL_PDF
        ]
        processed: Dict[str, Tuple[Path, int, Path, int]] = {}
        for num, p in real_files:
            cover_docx = OUTPUT_DIR / f"cover_{num}.docx"
            cover_pdf = OUTPUT_DIR / f"cover_{num}.pdf"
            create_cover_page(cover_docx, num, p.name)
            convert_docx_to_pdf(cover_docx, cover_pdf)
            cover_pages = pdf_page_count(cover_pdf)
            if p.suffix.lower() == ".docx":
                file_pdf = OUTPUT_DIR / f"file_{num}.pdf"
                convert_docx_to_pdf(p, file_pdf)
            else:
                file_pdf = p
            file_pages = pdf_page_count(file_pdf)
            processed[num] = (cover_pdf, cover_pages, file_pdf, file_pages)
            log_debug(f"Processed {num}: cover={cover_pdf.name} ({cover_pages} pages), file={file_pdf.name} ({file_pages} pages)")
        dummy_contents_docx = OUTPUT_DIR / "contents_dummy.docx"
        create_contents_page(dummy_contents_docx, all_items, {}, {}, real_files, 0)
        dummy_contents_pdf = OUTPUT_DIR / "contents_dummy.pdf"
        convert_docx_to_pdf(dummy_contents_docx, dummy_contents_pdf)
        contents_pages = pdf_page_count(dummy_contents_pdf)
        log_debug(f"Dummy contents PDF pages: {contents_pages}")
        current_page = BATES_START + contents_pages
        bates_map: Dict[str, int] = {}
        for num, _ in real_files:
            cover_pdf, cover_pages, file_pdf, file_pages = processed[num]
            bates_map[num] = current_page
            current_page += cover_pages + file_pages
        log_debug(f"Bates mapping: {bates_map}")
        contents_docx = OUTPUT_DIR / "contents.docx"
        create_contents_page(contents_docx, all_items, bates_map, processed, real_files, contents_pages)
        contents_pdf = OUTPUT_DIR / "contents.pdf"
        convert_docx_to_pdf(contents_docx, contents_pdf)
        final_order: List[Path] = [contents_pdf]
        for num, _ in real_files:
            cover_pdf, _, file_pdf, _ = processed[num]
            final_order.extend([cover_pdf, file_pdf])
        merge_pdfs(final_order, FINAL_PDF)
        apply_bates_numbering(FINAL_PDF, start_number=BATES_START, font_size=BATES_FONT_SIZE)
        toc_list = []
        for num, p in real_files:
            title = f"{num} - {p.name}"
            toc_list.append([1, title, bates_map[num]])
        add_pdf_bookmarks(FINAL_PDF, toc_list)
        for pattern in ["cover_*.docx", "cover_*.pdf", "file_*.pdf", "contents_dummy.*", "contents.docx", "contents.pdf"]:
            for temp in OUTPUT_DIR.glob(pattern):
                try:
                    temp.unlink()
                    log_debug(f"Removed temporary file: {temp}")
                except Exception:
                    pass
        log("Process complete. Check final_output.pdf and script_log.txt.")
    except Exception as e:
        log_error(f"Fatal error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
