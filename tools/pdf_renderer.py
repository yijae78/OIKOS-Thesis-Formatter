"""PDF renderer — Word COM으로 docx→PDF 변환, PyMuPDF로 페이지별 이미지 생성."""

import os
import hashlib
import time
from pathlib import Path

CACHE_DIR = Path(__file__).parent.parent / "cache"
CACHE_DIR.mkdir(exist_ok=True)


def docx_to_pdf(docx_path: Path, pdf_path: Path = None, force: bool = False) -> Path:
    """Convert docx to PDF using Word COM automation.

    Returns the PDF path. Caches by file modification time.
    """
    docx_path = Path(docx_path).resolve()
    if pdf_path is None:
        name = docx_path.stem
        pdf_path = CACHE_DIR / f"{name}.pdf"

    # Check cache: skip if PDF is newer than docx
    if not force and pdf_path.exists():
        if pdf_path.stat().st_mtime >= docx_path.stat().st_mtime:
            return pdf_path

    import win32com.client
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(str(docx_path), ReadOnly=True)
        doc.ExportAsFixedFormat(str(pdf_path), 17)  # wdExportFormatPDF
        doc.Close(False)
        doc = None
    finally:
        if doc:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass

    return pdf_path


def render_pdf_pages(pdf_path: Path, dpi: int = 144, force: bool = False) -> list[Path]:
    """Render all PDF pages as PNG images. Returns list of image paths.

    Caches by PDF modification time. Use force=True to regenerate.
    """
    import fitz

    pdf_path = Path(pdf_path)
    cache_key = hashlib.md5(f"{pdf_path.name}_{pdf_path.stat().st_mtime}".encode()).hexdigest()[:10]
    pages_dir = CACHE_DIR / f"pages_{cache_key}"

    # Check if already rendered
    if not force and pages_dir.exists():
        existing = sorted(pages_dir.glob("page_*.png"))
        if existing:
            return existing

    pages_dir.mkdir(exist_ok=True)

    pdf = fitz.open(str(pdf_path))
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    paths = []

    for i in range(pdf.page_count):
        page = pdf[i]
        pix = page.get_pixmap(matrix=mat)
        out = pages_dir / f"page_{i + 1:03d}.png"
        pix.save(str(out))
        paths.append(out)

    pdf.close()
    return paths


def get_pdf_page_count(pdf_path: Path) -> int:
    """Get total page count from a PDF."""
    import fitz
    pdf = fitz.open(str(pdf_path))
    count = pdf.page_count
    pdf.close()
    return count


def ensure_pdf_images(docx_path: Path, dpi: int = 144, force: bool = False) -> list[Path]:
    """Full pipeline: docx → PDF → page images. Returns list of PNG paths.
    Use force=True to regenerate even if cache exists."""
    pdf_path = docx_to_pdf(docx_path, force=force)
    return render_pdf_pages(pdf_path, dpi=dpi, force=force)
