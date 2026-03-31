"""Page number manager — control page numbering format and start values per section."""

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pathlib import Path


# Valid page number formats in Word
PAGE_FORMATS = {
    'decimal': 'decimal',           # 1, 2, 3...
    'lowerRoman': 'lowerRoman',     # i, ii, iii...
    'upperRoman': 'upperRoman',     # I, II, III...
    'lowerLetter': 'lowerLetter',   # a, b, c...
    'upperLetter': 'upperLetter',   # A, B, C...
}


def get_page_number_settings(doc: Document) -> list[dict]:
    """Get page number settings for all sections."""
    settings = []
    for i, section in enumerate(doc.sections):
        sec_props = section._sectPr
        pg_num_type = sec_props.find(qn('w:pgNumType'))

        fmt = None
        start = None
        if pg_num_type is not None:
            fmt = pg_num_type.get(qn('w:fmt'))
            start_val = pg_num_type.get(qn('w:start'))
            start = int(start_val) if start_val is not None else None

        settings.append({
            'section': i + 1,
            'format': fmt,
            'start': start,
            'description': _describe_numbering(fmt, start),
        })
    return settings


def set_page_number_format(doc: Document, section_index: int, fmt: str = None, start: int = None) -> None:
    """Set page number format and/or start value for a specific section.

    Args:
        doc: The Document object
        section_index: 0-based section index
        fmt: Page number format (decimal, lowerRoman, upperRoman, lowerLetter, upperLetter)
        start: Starting page number (integer)
    """
    if section_index < 0 or section_index >= len(doc.sections):
        raise ValueError(f"Section index {section_index} out of range (0-{len(doc.sections) - 1})")

    if fmt and fmt not in PAGE_FORMATS:
        raise ValueError(f"Invalid format '{fmt}'. Valid: {list(PAGE_FORMATS.keys())}")

    section = doc.sections[section_index]
    sec_props = section._sectPr

    # Find or create pgNumType element
    pg_num_type = sec_props.find(qn('w:pgNumType'))
    if pg_num_type is None:
        pg_num_type = OxmlElement('w:pgNumType')
        sec_props.append(pg_num_type)

    # Set format if specified
    if fmt is not None:
        pg_num_type.set(qn('w:fmt'), fmt)

    # Set start if specified
    if start is not None:
        pg_num_type.set(qn('w:start'), str(start))


def remove_page_number_start(doc: Document, section_index: int) -> None:
    """Remove explicit start value so section continues from previous."""
    if section_index < 0 or section_index >= len(doc.sections):
        raise ValueError(f"Section index {section_index} out of range")

    section = doc.sections[section_index]
    sec_props = section._sectPr
    pg_num_type = sec_props.find(qn('w:pgNumType'))

    if pg_num_type is not None:
        start_attr = qn('w:start')
        if start_attr in pg_num_type.attrib:
            del pg_num_type.attrib[start_attr]
            print(f"[OK] Section {section_index + 1}: Removed explicit start value (will continue from previous)")


def set_page_numbers_for_range(doc: Document, start_section: int, end_section: int,
                                fmt: str = None, start: int = None) -> None:
    """Set page number format for a range of sections (1-based, inclusive)."""
    for sec_num in range(start_section, end_section + 1):
        sec_idx = sec_num - 1
        if 0 <= sec_idx < len(doc.sections):
            set_page_number_format(doc, sec_idx, fmt=fmt, start=start if sec_num == start_section else None)
            print(f"[OK] Section {sec_num}: format={fmt or '(unchanged)'}, "
                  f"start={start if sec_num == start_section else '(continue)'}")


def print_page_number_report(doc: Document) -> None:
    """Print a formatted report of all page number settings."""
    settings = get_page_number_settings(doc)

    print(f"\n{'─' * 60}")
    print("PAGE NUMBER SETTINGS")
    print(f"{'─' * 60}")
    print(f"{'Section':>8} {'Format':>14} {'Start':>8} {'Description'}")
    print(f"{'─' * 8} {'─' * 14} {'─' * 8} {'─' * 25}")

    for s in settings:
        fmt = s['format'] or '(inherit)'
        start = str(s['start']) if s['start'] is not None else '-'
        print(f"{s['section']:>8} {fmt:>14} {start:>8} {s['description']}")


def _describe_numbering(fmt, start):
    """Human-readable description of page numbering."""
    if fmt is None and start is None:
        return "Continues from previous"

    parts = []
    if fmt == 'lowerRoman':
        parts.append("Roman numerals (i, ii, iii)")
    elif fmt == 'upperRoman':
        parts.append("Roman numerals (I, II, III)")
    elif fmt == 'decimal':
        parts.append("Arabic numbers (1, 2, 3)")
    elif fmt == 'lowerLetter':
        parts.append("Letters (a, b, c)")
    elif fmt == 'upperLetter':
        parts.append("Letters (A, B, C)")
    elif fmt:
        parts.append(f"Format: {fmt}")

    if start is not None:
        parts.append(f"starts at {start}")

    return ", ".join(parts) if parts else "Continues from previous"
