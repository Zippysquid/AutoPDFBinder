#!/usr/bin/env python
"""
AutoPDFBinder - Enhanced Legal Document PDF Bundler
Automatically converts directories of Word and PDF files into a single,
print-ready PDF with table of contents, cover pages, Bates numbering,
and bookmarks.
"""
import os
import sys
import subprocess
import time
import logging
import shutil
import json
import argparse
from pathlib import Path
from typing import List, Tuple, Dict, Optional
from dataclasses import dataclass, asdict

# For Word-to-PDF conversion via COM automation (Windows only)
try:
    import comtypes
    from comtypes.client import CreateObject
    HAS_COM = True
except (ImportError, OSError):
    HAS_COM = False
    # COM not available (not on Windows or comtypes not installed)
    pass

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

# AI Integration (optional)
try:
    import anthropic
    import base64
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

###############################################################################
# Configuration & Constants
###############################################################################
SCRIPT_DIR = Path(__file__).parent.absolute()
OUTPUT_DIR = SCRIPT_DIR / "output"
LOG_FILE = SCRIPT_DIR / "script_log.txt"
FINAL_PDF = SCRIPT_DIR / "final_output.pdf"
CONFIG_FILE = SCRIPT_DIR / "config.json"

BATES_START = 1
BATES_FONT_SIZE = 14
TIMEOUT_BASE = 1.0

# Design Configuration (can be overridden by config.json)
DESIGN_CONFIG = {
    'primary_color': (0, 51, 102),      # Navy blue (RGB tuple)
    'accent_color': (200, 200, 200),    # Light gray
    'header_bg': (245, 245, 245),       # Very light gray
    'alt_row_color': (250, 250, 250),   # Almost white
    'text_gray': (100, 100, 100),       # Medium gray
    'cover_style': 'modern',
    'company_name': 'Document Archive',
    'department': 'Document Management System',
    'show_metadata': True,
    'use_icons': True,
    'add_toc_links': True,              # Enable clickable TOC links
    'nested_bookmarks': True,            # Enable nested bookmark hierarchy
    'ai_features': {
        'enabled': False,                # Enable AI document analysis
        'max_pages_to_analyze': 3,       # Number of pages to send to Claude
        'cache_results': True,           # Cache AI analysis results
        'model': 'claude-3-5-sonnet-20241022'
    }
}

@dataclass
class ProcessingStats:
    """Track processing statistics"""
    total_files: int = 0
    successful_files: int = 0
    failed_files: int = 0
    total_pages: int = 0
    start_time: float = 0.0
    end_time: float = 0.0
    failed_file_list: List[str] = None

    def __post_init__(self):
        if self.failed_file_list is None:
            self.failed_file_list = []

    def duration(self) -> float:
        return self.end_time - self.start_time if self.end_time else time.time() - self.start_time

stats = ProcessingStats()

###############################################################################
# Logging Setup
###############################################################################
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8", mode='w'),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.info
log_error = logging.error
log_debug = logging.debug
log_warning = logging.warning

###############################################################################
# Configuration Management
###############################################################################
def load_config() -> Dict:
    """Load configuration from config.json if it exists"""
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
                log(f"Loaded configuration from {CONFIG_FILE.name}")
                return config
        except Exception as e:
            log_warning(f"Could not load config file: {e}. Using defaults.")
    return {}

def save_default_config():
    """Save default configuration to config.json"""
    default_config = {
        "bates_start": BATES_START,
        "bates_font_size": BATES_FONT_SIZE,
        "design": DESIGN_CONFIG,
        "output_filename": "final_output.pdf"
    }
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(default_config, f, indent=2)
        log(f"Created default config file: {CONFIG_FILE.name}")
    except Exception as e:
        log_warning(f"Could not save config file: {e}")

def apply_config(config: Dict):
    """Apply configuration settings"""
    global BATES_START, BATES_FONT_SIZE, FINAL_PDF

    if 'bates_start' in config:
        BATES_START = config['bates_start']
    if 'bates_font_size' in config:
        BATES_FONT_SIZE = config['bates_font_size']
    if 'output_filename' in config:
        FINAL_PDF = SCRIPT_DIR / config['output_filename']
    if 'design' in config:
        DESIGN_CONFIG.update(config['design'])

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
            log(f"✓ Package {pkg_name} is installed")
        except ImportError:
            log(f"Installing {pkg_name}...")
            try:
                subprocess.run(
                    [sys.executable, "-m", "pip", "install", "--user", pkg_name],
                    capture_output=True, text=True, check=True, timeout=300
                )
                log(f"✓ Installed {pkg_name}")
            except subprocess.CalledProcessError as e:
                log_error(f"✗ Failed to install {pkg_name}: {e.stderr}")
                missing = True
    if missing:
        log_error("Critical: Some packages failed to install.")
        sys.exit(1)

###############################################################################
# Progress Tracking
###############################################################################
def print_progress(current: int, total: int, task: str = "Processing"):
    """Print progress bar"""
    if total == 0:
        return
    percent = int((current / total) * 100)
    bar_length = 40
    filled = int((bar_length * current) / total)
    bar = '█' * filled + '░' * (bar_length - filled)
    print(f"\r{task}: |{bar}| {percent}% ({current}/{total})", end='', flush=True)
    if current == total:
        print()  # New line when complete

###############################################################################
# Helper Functions
###############################################################################
def rgb_to_color(rgb_tuple) -> RGBColor:
    """Convert RGB tuple to RGBColor"""
    return RGBColor(*rgb_tuple)

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
    """Set margins (in dxa units) for a table cell."""
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

def set_cell_background(cell, color):
    """Set background color for a table cell."""
    # Color should be an RGB tuple (r, g, b)
    hex_color = '%02x%02x%02x' % color

    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_table_borders(table, color = None, width: int = 1):
    """Add subtle borders to a table."""
    if color is None:
        color = (200, 200, 200)

    # Color should be an RGB tuple (r, g, b)
    hex_color = '%02x%02x%02x' % color

    tbl = table._element
    tblPr = tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')

    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), str(width * 8))
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

###############################################################################
# AI Document Analysis
###############################################################################
def extract_pdf_preview(pdf_path: Path, max_pages: int = 3) -> List[bytes]:
    """Extract first N pages of PDF as PNG images for Claude Vision"""
    try:
        doc = fitz.open(str(pdf_path))
        images = []
        for i in range(min(max_pages, len(doc))):
            page = doc[i]
            # Render at 150 DPI for good quality
            pix = page.get_pixmap(dpi=150)
            images.append(pix.tobytes("png"))
        doc.close()
        return images
    except Exception as e:
        log_error(f"Error extracting preview from {pdf_path.name}: {e}")
        return []

def analyze_document_with_ai(pdf_path: Path, filename: str) -> Optional[Dict]:
    """Use Claude API to analyze document and extract metadata"""
    if not HAS_ANTHROPIC:
        return None

    if not DESIGN_CONFIG.get('ai_features', {}).get('enabled', False):
        return None

    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        log_warning("AI features enabled but ANTHROPIC_API_KEY not set")
        return None

    try:
        # Extract preview images
        max_pages = DESIGN_CONFIG['ai_features'].get('max_pages_to_analyze', 3)
        images = extract_pdf_preview(pdf_path, max_pages)

        if not images:
            return None

        # Prepare Claude API call
        client = anthropic.Anthropic(api_key=api_key)

        # Build message content with images
        content = []
        for img in images:
            content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": base64.b64encode(img).decode()
                }
            })

        # Add analysis prompt
        content.append({
            "type": "text",
            "text": f"""Analyze this document (filename: {filename}).

Return ONLY valid JSON with this exact structure:
{{
  "title": "Actual document title from content (or null)",
  "description": "One-line description max 100 chars (or null)",
  "document_type": "Contract|Invoice|Report|Letter|Agreement|Memo|Other",
  "date": "YYYY-MM-DD format if visible (or null)",
  "parties": ["Party 1", "Party 2"] or null,
  "reference_number": "Document/case/reference number (or null)",
  "category": "Legal|Financial|Administrative|Technical|Other",
  "quality_issues": "Brief note if pages corrupt/illegible (or null)"
}}

Rules:
- Use null for any information not clearly visible
- Keep description concise and professional
- Extract exact dates in YYYY-MM-DD format
- Identify all parties mentioned (people/organizations)
- Be accurate - don't guess or infer beyond what's visible"""
        })

        # Call Claude API
        model = DESIGN_CONFIG['ai_features'].get('model', 'claude-3-5-sonnet-20241022')
        log_debug(f"Analyzing {filename} with Claude {model}...")

        message = client.messages.create(
            model=model,
            max_tokens=1024,
            messages=[{"role": "user", "content": content}]
        )

        # Parse JSON response
        response_text = message.content[0].text
        ai_data = json.loads(response_text)

        log(f"✓ AI analyzed: {filename}")
        log_debug(f"  Title: {ai_data.get('title', 'N/A')}")
        log_debug(f"  Type: {ai_data.get('document_type', 'N/A')}")

        return ai_data

    except json.JSONDecodeError as e:
        log_warning(f"AI returned invalid JSON for {filename}: {e}")
        return None
    except Exception as e:
        log_warning(f"AI analysis failed for {filename}: {e}")
        return None

def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> bool:
    """Convert a DOCX file to PDF using Microsoft Word COM automation."""
    if not HAS_COM:
        log_error(f"Cannot convert {docx_path.name}: COM automation not available (Windows with MS Word required)")
        return False

    log(f"Converting {docx_path.name} -> {pdf_path.name}")
    if not docx_path.exists():
        log_error(f"{docx_path} not found.")
        return False
    word = None
    try:
        word = CreateObject('Word.Application')
        word.DisplayAlerts = 0
        word.Visible = False
        doc = word.Documents.Open(str(docx_path), ReadOnly=True)
        doc.SaveAs(str(pdf_path), FileFormat=17)
        doc.Close()
        time.sleep(TIMEOUT_BASE)
    except Exception as e:
        log_error(f"Error converting {docx_path.name}: {e}")
        return False
    finally:
        if word is not None:
            try:
                word.NormalTemplate.Saved = True
            except:
                pass
            word.Quit()
    if not pdf_path.exists():
        log_error(f"Conversion failed: {pdf_path.name} not created")
        return False
    log(f"✓ Converted {docx_path.name}")
    return True

def pdf_page_count(pdf_path: Path) -> int:
    """Return the number of pages in a PDF file."""
    try:
        reader = PyPDF2.PdfReader(str(pdf_path))
        return len(reader.pages)
    except Exception as e:
        log_error(f"Error counting pages in {pdf_path.name}: {e}")
        return 0

def merge_pdfs(pdf_paths: List[Path], final_pdf: Path) -> None:
    """Merge PDFs in the provided order into a final PDF."""
    from PyPDF2 import PdfMerger
    log(f"Merging {len(pdf_paths)} PDFs into {final_pdf.name}")
    merger = PdfMerger()
    for idx, p in enumerate(pdf_paths, 1):
        print_progress(idx, len(pdf_paths), "Merging PDFs")
        if p.exists():
            pages = pdf_page_count(p)
            if pages > 0:
                log_debug(f"Appending {p.name} ({pages} pages)")
                merger.append(str(p))
                stats.total_pages += pages
            else:
                log_warning(f"Skipping {p.name}, 0 pages")
        else:
            log_error(f"File missing: {p.name}")
    merger.write(str(final_pdf))
    merger.close()
    if not final_pdf.exists():
        raise FileNotFoundError(f"Merged PDF not created: {final_pdf}")
    log(f"✓ Merged PDF saved: {final_pdf.name}")

def apply_bates_numbering(pdf_path: Path, start_number: int = BATES_START, font_size: int = BATES_FONT_SIZE) -> None:
    """Apply sequential Bates numbering with a subtle background to each page."""
    log(f"Applying Bates numbering to {pdf_path.name}")
    doc = fitz.open(str(pdf_path))
    total_pages = len(doc)
    for i, page in enumerate(doc):
        print_progress(i + 1, total_pages, "Adding Bates numbers")
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
    log(f"✓ Bates numbering applied")

def add_pdf_bookmarks(pdf_path: Path, toc_list: List[List]) -> None:
    """Add bookmarks (outline) to the merged PDF with proper nesting."""
    log("Adding PDF bookmarks (outline)...")
    try:
        doc = fitz.open(str(pdf_path))
        doc.set_toc(toc_list)
        new_toc = doc.get_toc()
        log_debug(f"TOC entries: {len(new_toc)}")

        temp_pdf = pdf_path.with_suffix('.temp.pdf')
        doc.save(str(temp_pdf), garbage=4, deflate=True)
        doc.close()

        temp_pdf.replace(pdf_path)
        log("✓ PDF bookmarks added")
    except Exception as e:
        log_warning(f"Could not add PDF bookmarks: {e}")
        log("The PDF has been created successfully but without bookmarks")
        if 'doc' in locals():
            doc.close()

def add_toc_links(pdf_path: Path, link_entries: List[Tuple[str, int]], contents_pages: int) -> None:
    """Insert clickable links on the table of contents pages."""
    log("Adding clickable links to table of contents...")
    try:
        doc = fitz.open(str(pdf_path))
        links_added = 0

        for text, target_page in link_entries:
            for page_num in range(contents_pages):
                page = doc[page_num]
                # Search for the text on the page
                rects = page.search_for(text)
                for rect in rects:
                    # Create a link annotation
                    link = {
                        "kind": fitz.LINK_GOTO,
                        "page": target_page - 1,  # Convert to 0-indexed
                        "from": rect
                    }
                    page.insert_link(link)
                    links_added += 1
                    log_debug(f"Added link: '{text[:30]}...' -> page {target_page}")

        temp_pdf = pdf_path.with_suffix('.links.pdf')
        doc.save(str(temp_pdf))
        doc.close()
        temp_pdf.replace(pdf_path)
        log(f"✓ Added {links_added} clickable TOC links")
    except Exception as e:
        log_warning(f"Could not add TOC links: {e}")
        if 'doc' in locals():
            doc.close()

###############################################################################
# Document Creation Functions
###############################################################################
def create_cover_page_pdf(cover_pdf: Path, number: str, file_name: str, actual_path: Optional[Path] = None, ai_data: Optional[Dict] = None) -> None:
    """Generate a cover page directly as PDF using PyMuPDF (cross-platform)."""
    log_debug(f"Creating cover page (PDF): {cover_pdf.name}")

    doc = fitz.open()
    page = doc.new_page(width=612, height=792)  # US Letter

    # Colors from DESIGN_CONFIG
    primary_color = tuple(c/255.0 for c in DESIGN_CONFIG['primary_color'])
    text_gray = tuple(c/255.0 for c in DESIGN_CONFIG['text_gray'])

    # Header background (navy blue bar)
    header_rect = fitz.Rect(50, 50, 562, 120)
    page.draw_rect(header_rect, color=primary_color, fill=primary_color)

    # Document number and category in header
    header_text = f"DOCUMENT {number}"
    if ai_data and ai_data.get('category'):
        header_text += f" | {ai_data['category']}"
    page.insert_text((306, 95), header_text,
                     fontname="helv-bold", fontsize=14, color=(1, 1, 1), align=1)

    # Main title (use AI-extracted title if available)
    title_y = 300
    display_title = ai_data.get('title') if ai_data and ai_data.get('title') else file_name
    page.insert_text((306, title_y), display_title,
                     fontname="helv-bold", fontsize=20, color=(0, 0, 0), align=1)

    # AI description (if available)
    if ai_data and ai_data.get('description'):
        page.insert_text((306, title_y + 30), ai_data['description'],
                         fontname="helv-oblique", fontsize=11, color=text_gray, align=1)

    # Metadata box (enhanced with AI data)
    if DESIGN_CONFIG['show_metadata']:
        meta_y = 400
        metadata = []

        # Add AI-extracted metadata first
        if ai_data:
            if ai_data.get('document_type'):
                metadata.append(("Document Type:", ai_data['document_type']))
            if ai_data.get('date'):
                metadata.append(("Date:", ai_data['date']))
            if ai_data.get('parties'):
                parties_str = ', '.join(ai_data['parties'][:2])  # First 2 parties
                if len(ai_data['parties']) > 2:
                    parties_str += f" +{len(ai_data['parties']) - 2} more"
                metadata.append(("Parties:", parties_str))
            if ai_data.get('reference_number'):
                metadata.append(("Reference:", ai_data['reference_number']))

        # Add standard metadata
        if not any(label == "Document Type:" for label, _ in metadata):
            metadata.append(("Document Type:", get_file_type(file_name)))

        metadata.append(("File Size:", get_file_size_if_available(actual_path) if actual_path else "N/A"))
        metadata.append(("Generated:", time.strftime("%B %d, %Y at %I:%M %p")))

        # Quality warning if AI detected issues
        if ai_data and ai_data.get('quality_issues'):
            metadata.append(("⚠ Warning:", ai_data['quality_issues']))

        for label, value in metadata:
            if value:  # Only show non-null values
                page.insert_text((120, meta_y), label, fontname="helv-bold", fontsize=10, color=text_gray)
                page.insert_text((250, meta_y), str(value), fontname="helv", fontsize=10, color=(0, 0, 0))
                meta_y += 22

    # Bottom line
    line_y = 700
    page.draw_line((100, line_y), (512, line_y), color=(0.8, 0.8, 0.8), width=1)

    # AI attribution (if used)
    if ai_data:
        page.insert_text((306, 750), "Analyzed by Claude AI",
                         fontname="helv-oblique", fontsize=8, color=(0.7, 0.7, 0.7), align=1)

    doc.save(str(cover_pdf))
    doc.close()

def create_cover_page(cover_docx: Path, number: str, file_name: str, actual_path: Optional[Path] = None) -> None:
    """Generate an enhanced cover page with professional design elements."""
    log_debug(f"Creating cover page: {cover_docx.name}")
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

    # Add document number in header
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

    # Main document title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    title_run = title_para.add_run(file_name)
    title_run.font.name = "Calibri"
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_para.paragraph_format.space_after = Pt(24)

    # Add file metadata
    if DESIGN_CONFIG['show_metadata']:
        doc.add_paragraph()

        meta_table = doc.add_table(rows=3, cols=2)
        meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        meta_table.autofit = False

        set_table_borders(meta_table, color=DESIGN_CONFIG['accent_color'], width=1)

        # Add metadata rows
        metadata = [
            ("Document Type:", get_file_type(file_name)),
            ("File Size:", get_file_size_if_available(actual_path) if actual_path else "N/A"),
            ("Generated:", time.strftime("%B %d, %Y at %I:%M %p"))
        ]

        for i, (label, value) in enumerate(metadata):
            label_cell = meta_table.cell(i, 0)
            value_cell = meta_table.cell(i, 1)

            label_cell.width = Inches(1.5)
            value_cell.width = Inches(3.5)

            label_para = label_cell.paragraphs[0]
            label_run = label_para.add_run(label)
            label_run.font.bold = True
            label_run.font.size = Pt(10)
            label_run.font.name = "Calibri"
            label_run.font.color.rgb = rgb_to_color(DESIGN_CONFIG['text_gray'])

            value_para = value_cell.paragraphs[0]
            value_run = value_para.add_run(value)
            value_run.font.size = Pt(10)
            value_run.font.name = "Calibri"

            set_cell_margins(label_cell, top=100, bottom=100, left=100, right=50)
            set_cell_margins(value_cell, top=100, bottom=100, left=50, right=100)

    for _ in range(6):
        doc.add_paragraph()

    # Add a subtle bottom border
    bottom_line = doc.add_paragraph()
    bottom_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = bottom_line.add_run("_" * 60)
    line_run.font.color.rgb = rgb_to_color(DESIGN_CONFIG['accent_color'])
    line_run.font.size = Pt(8)

    safe_save_docx(doc, cover_docx)

def create_contents_page(contents_docx: Path, all_items: List[Tuple[str, Path]],
                         bates_map: Dict[str, int],
                         processed: Dict[str, Tuple[Path, int, Path, int]] = None,
                         real_files: List[Tuple[str, Path]] = None,
                         contents_pages: int = 0) -> None:
    """Create an enhanced table of contents with modern design elements."""
    log_debug(f"Creating contents page: {contents_docx.name}")
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
    title_run.font.color.rgb = rgb_to_color(DESIGN_CONFIG['primary_color'])

    # Subtitle with date and statistics
    subtitle_cell = header_table.cell(1, 0)
    subtitle_cell.width = Inches(6.5)
    subtitle_para = subtitle_cell.paragraphs[0]
    subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

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
    subtitle_run.font.color.rgb = rgb_to_color(DESIGN_CONFIG['text_gray'])
    subtitle_run.font.name = "Calibri"

    # Add decorative divider
    doc.add_paragraph()

    line_table = doc.add_table(rows=1, cols=1)
    line_table.autofit = False
    line_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    line_cell = line_table.cell(0, 0)
    line_cell.width = Inches(6.5)
    line_cell.height = Pt(1)
    set_cell_background(line_cell, DESIGN_CONFIG['accent_color'])
    remove_table_borders(line_table)

    doc.add_paragraph()

    # Create main content table
    use_icons = DESIGN_CONFIG['use_icons']
    num_cols = 3 if use_icons else 2

    table = doc.add_table(rows=len(all_items) + 1, cols=num_cols)
    table.style = "Table Grid"
    remove_table_borders(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set column widths
    if use_icons:
        widths = [Inches(0.4), Inches(5.1), Inches(1.0)]
    else:
        widths = [Inches(5.5), Inches(1.0)]

    for col_idx, width in enumerate(widths):
        for cell in table.columns[col_idx].cells:
            cell.width = width

    # Style header row
    header_row = table.rows[0]
    header_texts = ["", "Document", "Page"] if use_icons else ["Document", "Page"]

    for i, text in enumerate(header_texts):
        cell = header_row.cells[i]
        set_cell_background(cell, DESIGN_CONFIG['header_bg'])
        set_cell_margins(cell, top=100, bottom=100, left=100 if i == 0 else 50, right=100 if i == len(header_texts)-1 else 50)

        if text:
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT if text == "Document" else WD_ALIGN_PARAGRAPH.RIGHT
            run = para.add_run(text)
            run.font.bold = True
            run.font.size = Pt(11)
            run.font.color.rgb = rgb_to_color(DESIGN_CONFIG['primary_color'])
            run.font.name = "Calibri"

    # Add content rows
    for idx, (num, path_obj) in enumerate(all_items):
        row_idx = idx + 1
        row = table.rows[row_idx]

        if row_idx % 2 == 0:
            for cell in row.cells:
                set_cell_background(cell, DESIGN_CONFIG['alt_row_color'])

        col_offset = 0

        # Icon column
        if use_icons:
            icon_cell = row.cells[0]
            icon_para = icon_cell.paragraphs[0]
            icon_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_margins(icon_cell, top=50, bottom=50)

            if path_obj.is_dir():
                icon = "📁"
            else:
                ext = path_obj.suffix.lower()
                icon = "📄" if ext == ".pdf" else "📝" if ext in [".docx", ".doc"] else "📎"

            icon_run = icon_para.add_run(icon)
            icon_run.font.size = Pt(11)
            col_offset = 1

        # Document name column
        name_cell = row.cells[col_offset]
        name_para = name_cell.paragraphs[0]
        set_cell_margins(name_cell, top=50, bottom=50, left=100, right=50)

        indent_level = num.count(".")

        if indent_level > 0:
            indent_text = "  " * indent_level + "└─ "
            indent_run = name_para.add_run(indent_text)
            indent_run.font.color.rgb = RGBColor(200, 200, 200)
            indent_run.font.name = "Courier New"
            indent_run.font.size = Pt(10)

        num_run = name_para.add_run(f"{num}. ")
        num_run.font.name = "Calibri"
        num_run.font.size = Pt(11)
        num_run.font.color.rgb = rgb_to_color(DESIGN_CONFIG['text_gray'])

        name_run = name_para.add_run(path_obj.name)
        name_run.font.name = "Calibri"
        name_run.font.size = Pt(11)

        if path_obj.is_dir():
            name_run.font.bold = True
            name_run.font.color.rgb = rgb_to_color(DESIGN_CONFIG['primary_color'])

        # Page number column
        page_cell = row.cells[col_offset + 1]
        page_para = page_cell.paragraphs[0]
        page_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_cell_margins(page_cell, top=50, bottom=50, left=50, right=100)

        if num in bates_map and path_obj.is_file():
            page_run = page_para.add_run(f"{bates_map[num]:03d}")
            page_run.font.name = "Calibri"
            page_run.font.size = Pt(11)

    # Add footer
    doc.add_paragraph()
    doc.add_paragraph()

    if bates_map:
        footer_para = doc.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if processed and real_files:
            total_pages = contents_pages
            for num, _ in real_files:
                cover_pdf, cover_pages, file_pdf, file_pages = processed[num]
                total_pages += cover_pages + file_pages
        else:
            total_pages = max(bates_map.values()) if bates_map else 0

        footer_run = footer_para.add_run(f"Total Pages: {total_pages}")
        footer_run.font.size = Pt(9)
        footer_run.font.italic = True
        footer_run.font.color.rgb = rgb_to_color(DESIGN_CONFIG['text_gray'])
        footer_run.font.name = "Calibri"

    safe_save_docx(doc, contents_docx)

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
                log_debug(f"Cleaned up: {temp.name}")
            except PermissionError:
                log_warning(f"Could not delete {temp.name} (file may be open)")
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
                log_warning(f"Permission denied for {path.name}, retrying in 2 seconds...")
                time.sleep(2)
                if path.exists():
                    try:
                        path.unlink()
                    except:
                        pass
            else:
                raise

def print_summary():
    """Print processing summary"""
    duration = stats.duration()
    print("\n" + "="*60)
    print("PROCESSING SUMMARY")
    print("="*60)
    print(f"Total files processed: {stats.total_files}")
    print(f"Successful: {stats.successful_files}")
    print(f"Failed: {stats.failed_files}")
    print(f"Total pages in final PDF: {stats.total_pages}")
    print(f"Processing time: {duration:.2f} seconds")

    if stats.failed_files > 0:
        print(f"\nFailed files:")
        for failed in stats.failed_file_list:
            print(f"  - {failed}")

    print(f"\nOutput file: {FINAL_PDF}")
    print(f"Log file: {LOG_FILE}")
    print("="*60)

def build_nested_toc(all_items: List[Tuple[str, Path]], real_files: List[Tuple[str, Path]],
                     bates_map: Dict[str, int]) -> List[List]:
    """Build hierarchical bookmark structure for nested folders"""
    if not DESIGN_CONFIG.get('nested_bookmarks', False):
        # Flat structure (original behavior)
        toc_list = []
        for num, p in real_files:
            if num in bates_map:
                title = f"{num} - {p.name}"
                toc_list.append([1, title, bates_map[num]])
        return toc_list

    # Nested structure
    toc_list = []
    folder_map = {}  # Track folder entries

    for num, path_obj in all_items:
        level = num.count('.') + 1

        if path_obj.is_dir():
            # Add folder bookmark
            title = f"{num} - {path_obj.name}"
            # Use the first SUCCESSFUL file's page in this folder as the target
            first_file_page = None
            for file_num, file_path in real_files:
                if file_num.startswith(num + "."):
                    page = bates_map.get(file_num)
                    if page is not None:  # Found a successful file
                        first_file_page = page
                        break
            if first_file_page:
                toc_list.append([level, title, first_file_page])
                folder_map[num] = title
        elif path_obj.is_file() and num in bates_map:
            # Add file bookmark
            title = f"{num} - {path_obj.name}"
            toc_list.append([level, title, bates_map[num]])

    return toc_list

def main() -> None:
    parser = argparse.ArgumentParser(description='AutoPDFBinder - Legal Document PDF Bundler')
    parser.add_argument('--create-config', action='store_true', help='Create default config.json file')
    parser.add_argument('--output', type=str, help='Output PDF filename')
    parser.add_argument('--bates-start', type=int, help='Starting Bates number')
    parser.add_argument('--no-icons', action='store_true', help='Disable icons in TOC')
    parser.add_argument('--no-toc-links', action='store_true', help='Disable clickable TOC links')
    parser.add_argument('--use-ai', action='store_true', help='Enable AI document analysis (requires ANTHROPIC_API_KEY)')
    parser.add_argument('--ai-pages', type=int, default=3, help='Number of pages to analyze with AI (default: 3)')

    args = parser.parse_args()

    if args.create_config:
        save_default_config()
        return

    # Install missing packages (only when actually running, not for --help or --create-config)
    install_missing_packages()

    try:
        stats.start_time = time.time()

        # Load configuration
        config = load_config()
        apply_config(config)

        # Apply command-line overrides
        if args.output:
            global FINAL_PDF
            FINAL_PDF = SCRIPT_DIR / args.output
        if args.bates_start:
            global BATES_START
            BATES_START = args.bates_start
        if args.no_icons:
            DESIGN_CONFIG['use_icons'] = False
        if args.no_toc_links:
            DESIGN_CONFIG['add_toc_links'] = False
        if args.use_ai:
            if not HAS_ANTHROPIC:
                log_error("AI features requested but anthropic package not installed")
                log_error("Install with: pip install anthropic")
                sys.exit(1)
            DESIGN_CONFIG['ai_features']['enabled'] = True
            DESIGN_CONFIG['ai_features']['max_pages_to_analyze'] = args.ai_pages
            log(f"AI features enabled (analyzing first {args.ai_pages} pages)")

        OUTPUT_DIR.mkdir(exist_ok=True)
        log(f"Output directory: {OUTPUT_DIR}")

        cleanup_output_dir()

        # Scan directory
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

        stats.total_files = len(real_files)
        log(f"Found {stats.total_files} files to process")

        if stats.total_files == 0:
            log("No files to process. Exiting.")
            return

        # Process files with error recovery
        processed: Dict[str, Tuple[Path, int, Path, int]] = {}
        ai_cache: Dict[str, Optional[Dict]] = {}  # Cache AI results

        for idx, (num, p) in enumerate(real_files, 1):
            try:
                print_progress(idx, stats.total_files, "Processing files")

                # AI Analysis (if enabled)
                ai_data = None
                if DESIGN_CONFIG.get('ai_features', {}).get('enabled', False):
                    # Only analyze PDFs (DOCX would need conversion first)
                    if p.suffix.lower() == '.pdf':
                        ai_data = analyze_document_with_ai(p, p.name)
                        ai_cache[num] = ai_data

                cover_pdf = OUTPUT_DIR / f"cover_{num}.pdf"

                # Create cover page (use PDF method if COM not available)
                if HAS_COM:
                    cover_docx = OUTPUT_DIR / f"cover_{num}.docx"
                    create_cover_page(cover_docx, num, p.name, p)
                    if not convert_docx_to_pdf(cover_docx, cover_pdf):
                        raise Exception("Cover page conversion failed")
                else:
                    # Direct PDF creation (cross-platform) with AI data
                    create_cover_page_pdf(cover_pdf, num, p.name, p, ai_data=ai_data)

                cover_pages = pdf_page_count(cover_pdf)

                if p.suffix.lower() == ".docx":
                    if not HAS_COM:
                        raise Exception("DOCX conversion requires Windows + MS Word")
                    file_pdf = OUTPUT_DIR / f"file_{num}.pdf"
                    if not convert_docx_to_pdf(p, file_pdf):
                        raise Exception("Document conversion failed")
                else:
                    file_pdf = p

                file_pages = pdf_page_count(file_pdf)

                if file_pages == 0:
                    raise Exception("Document has 0 pages")

                processed[num] = (cover_pdf, cover_pages, file_pdf, file_pages)
                stats.successful_files += 1
                log_debug(f"✓ Processed {num}: {p.name}")

            except Exception as e:
                stats.failed_files += 1
                stats.failed_file_list.append(f"{p.name}: {str(e)}")
                log_error(f"✗ Failed to process {p.name}: {e}")
                # Continue with next file instead of crashing
                continue

        if stats.successful_files == 0:
            log_error("No files were successfully processed. Cannot create output PDF.")
            return

        # Update real_files to only include successfully processed files
        real_files = [(num, p) for num, p in real_files if num in processed]

        # Filter all_items to only include successfully processed files and their parent folders
        # Keep directories and files that were successfully processed
        processed_nums = set(processed.keys())
        filtered_all_items = []
        for num, path_obj in all_items:
            if path_obj.is_dir():
                # Keep folder if any of its children were successfully processed
                has_successful_child = any(
                    file_num.startswith(num + ".")
                    for file_num in processed_nums
                )
                if has_successful_child:
                    filtered_all_items.append((num, path_obj))
            elif num in processed_nums:
                # Keep file if it was successfully processed
                filtered_all_items.append((num, path_obj))

        all_items = filtered_all_items

        # Create contents page
        dummy_contents_docx = OUTPUT_DIR / "contents_dummy.docx"
        create_contents_page(dummy_contents_docx, all_items, {}, {}, real_files, 0)
        dummy_contents_pdf = OUTPUT_DIR / "contents_dummy.pdf"
        convert_docx_to_pdf(dummy_contents_docx, dummy_contents_pdf)
        contents_pages = pdf_page_count(dummy_contents_pdf)
        log_debug(f"Contents pages: {contents_pages}")

        # Build Bates map
        current_page = BATES_START + contents_pages
        bates_map: Dict[str, int] = {}
        for num, _ in real_files:
            cover_pdf, cover_pages, file_pdf, file_pages = processed[num]
            bates_map[num] = current_page
            current_page += cover_pages + file_pages
        log_debug(f"Bates map created for {len(bates_map)} documents")

        # Create final contents page
        contents_docx = OUTPUT_DIR / "contents.docx"
        create_contents_page(contents_docx, all_items, bates_map, processed, real_files, contents_pages)
        contents_pdf = OUTPUT_DIR / "contents.pdf"
        convert_docx_to_pdf(contents_docx, contents_pdf)

        # Merge all PDFs
        final_order: List[Path] = [contents_pdf]
        for num, _ in real_files:
            cover_pdf, _, file_pdf, _ = processed[num]
            final_order.extend([cover_pdf, file_pdf])

        merge_pdfs(final_order, FINAL_PDF)

        # Apply Bates numbering
        apply_bates_numbering(FINAL_PDF, start_number=BATES_START, font_size=BATES_FONT_SIZE)

        # Add bookmarks with proper nesting
        toc_list = build_nested_toc(all_items, real_files, bates_map)
        add_pdf_bookmarks(FINAL_PDF, toc_list)

        # Add clickable TOC links
        if DESIGN_CONFIG.get('add_toc_links', True):
            link_entries = []
            for num, p in real_files:
                if num in bates_map:
                    indent = "  " * num.count(".")
                    if indent:
                        indent += "└─ "
                    link_text = f"{indent}{num}. {p.name}"
                    link_entries.append((link_text, bates_map[num]))
            add_toc_links(FINAL_PDF, link_entries, contents_pages)

        # Cleanup
        cleanup_output_dir()

        stats.end_time = time.time()

        log("✓ Process complete!")
        print_summary()

    except KeyboardInterrupt:
        log("\nProcess interrupted by user")
        sys.exit(1)
    except Exception as e:
        log_error(f"Fatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
