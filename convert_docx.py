#!/usr/bin/env python3
"""Convert DOCX files to PDF using python-docx and reportlab"""
import sys
from pathlib import Path
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

def convert_docx_to_pdf(docx_path: Path, pdf_path: Path):
    """Convert a DOCX file to PDF"""
    try:
        # Load DOCX
        doc = Document(str(docx_path))

        # Create PDF
        pdf = SimpleDocTemplate(
            str(pdf_path),
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=18,
        )

        # Container for the 'Flowable' objects
        elements = []

        # Define styles
        styles = getSampleStyleSheet()
        normal_style = styles['Normal']
        heading_style = styles['Heading1']

        # Process DOCX paragraphs
        for para in doc.paragraphs:
            if para.text.strip():
                # Check if it's a heading
                if para.style.name.startswith('Heading'):
                    p = Paragraph(para.text, heading_style)
                else:
                    p = Paragraph(para.text, normal_style)
                elements.append(p)
                elements.append(Spacer(1, 0.2 * inch))

        # Build PDF
        pdf.build(elements)
        print(f"✓ Converted: {docx_path.name} -> {pdf_path.name}")
        return True

    except Exception as e:
        print(f"✗ Failed to convert {docx_path.name}: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: convert_docx.py <input.docx> <output.pdf>")
        sys.exit(1)

    docx_file = Path(sys.argv[1])
    pdf_file = Path(sys.argv[2])

    if not docx_file.exists():
        print(f"Error: {docx_file} not found")
        sys.exit(1)

    success = convert_docx_to_pdf(docx_file, pdf_file)
    sys.exit(0 if success else 1)
