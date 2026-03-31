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


def get_toc_entries(doc: Document) -> list[dict]:
    """Get all TOC entries with their cached page numbers.

    Returns list of dicts with keys:
        para_index, text, cached_page, bookmark, style
    """
    entries = []
    for i, para in enumerate(doc.paragraphs):
        style_name = para.style.name if para.style else ''
        if not (style_name.startswith('toc') or style_name == 'table of figures'):
            continue

        text = para.text.strip()
        if not text:
            continue

        # Extract PAGEREF bookmark and cached value
        # Runs may be inside <w:hyperlink>, so search recursively
        bookmark = None
        cached_page = None
        runs = para._element.findall('.//' + qn('w:r'))
        in_field_result = False

        for r in runs:
            fld = r.find(qn('w:fldChar'))
            if fld is not None:
                ftype = fld.get(qn('w:fldCharType'))
                if ftype == 'separate':
                    in_field_result = True
                elif ftype == 'end':
                    in_field_result = False
            else:
                instr = r.find(qn('w:instrText'))
                if instr is not None and instr.text and 'PAGEREF' in instr.text:
                    parts = instr.text.split()
                    for p in parts:
                        if p.startswith('_Toc'):
                            bookmark = p
                elif in_field_result:
                    t = r.find(qn('w:t'))
                    if t is not None and t.text:
                        cached_page = t.text

        # Split display text and page number
        display_text = text.rsplit('\t', 1)[0] if '\t' in text else text

        entries.append({
            'para_index': i,
            'display_text': display_text,
            'cached_page': cached_page,
            'bookmark': bookmark,
            'style': style_name,
        })
    return entries


def update_toc_cached_value(doc: Document, para_index: int, new_value: str) -> bool:
    """Update the cached PAGEREF value for a single TOC paragraph.

    Args:
        doc: The Document object
        para_index: 0-based paragraph index of the TOC entry
        new_value: New page number string (e.g. 'iii', '42')

    Returns True if updated, False if no PAGEREF field found.
    """
    para = doc.paragraphs[para_index]
    runs = para._element.findall('.//' + qn('w:r'))
    in_field_result = False

    for r in runs:
        fld = r.find(qn('w:fldChar'))
        if fld is not None:
            ftype = fld.get(qn('w:fldCharType'))
            if ftype == 'separate':
                in_field_result = True
            elif ftype == 'end':
                in_field_result = False
        elif in_field_result:
            t = r.find(qn('w:t'))
            if t is not None and t.text:
                old = t.text
                t.text = new_value
                return True
    return False


def fix_toc_page_numbers(doc: Document, corrections: dict[int, str]) -> list[str]:
    """Batch-update TOC PAGEREF cached values.

    Args:
        doc: The Document object
        corrections: {para_index: new_page_string} mapping

    Returns list of log messages.
    """
    logs = []
    entries = {e['para_index']: e for e in get_toc_entries(doc)}

    for para_idx, new_val in corrections.items():
        entry = entries.get(para_idx)
        if entry is None:
            logs.append(f"[SKIP] para[{para_idx}]: not a TOC entry")
            continue

        old_val = entry['cached_page']
        if old_val == new_val:
            logs.append(f"[SKIP] para[{para_idx}] {entry['display_text'][:30]}: already {new_val}")
            continue

        if update_toc_cached_value(doc, para_idx, new_val):
            logs.append(f"[OK] para[{para_idx}] {entry['display_text'][:30]}: {old_val} -> {new_val}")
        else:
            logs.append(f"[FAIL] para[{para_idx}] {entry['display_text'][:30]}: no PAGEREF field")

    return logs


INT_TO_ROMAN = [
    (1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'),
    (100, 'c'), (90, 'xc'), (50, 'l'), (40, 'xl'),
    (10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i'),
]


def int_to_lower_roman(n: int) -> str:
    """Convert integer to lowercase Roman numeral string."""
    result = []
    for value, numeral in INT_TO_ROMAN:
        while n >= value:
            result.append(numeral)
            n -= value
    return ''.join(result)


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
