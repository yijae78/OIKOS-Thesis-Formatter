"""Document structure analyzer — makes Word's invisible structure visible."""

from docx import Document
from docx.oxml.ns import qn
from pathlib import Path


def load_document(filepath: Path) -> Document:
    """Load a Word document."""
    return Document(str(filepath))


def get_section_info(doc: Document) -> list[dict]:
    """Extract detailed info for each section."""
    sections = []
    for i, section in enumerate(doc.sections):
        sec_props = section._sectPr

        # Page number format
        pg_num_type = sec_props.find(qn('w:pgNumType'))
        page_fmt = None
        page_start = None
        if pg_num_type is not None:
            page_fmt = pg_num_type.get(qn('w:fmt'))
            start_val = pg_num_type.get(qn('w:start'))
            page_start = int(start_val) if start_val is not None else None

        # Section break type
        sec_type_elem = sec_props.find(qn('w:type'))
        break_type = sec_type_elem.get(qn('w:val')) if sec_type_elem is not None else 'nextPage'

        # Page size
        pg_sz = sec_props.find(qn('w:pgSz'))
        width_emu = int(pg_sz.get(qn('w:w'), 0)) if pg_sz is not None else 0
        height_emu = int(pg_sz.get(qn('w:h'), 0)) if pg_sz is not None else 0

        # Margins
        pg_mar = sec_props.find(qn('w:pgMar'))
        margins = {}
        if pg_mar is not None:
            for attr in ['top', 'bottom', 'left', 'right', 'gutter']:
                val = pg_mar.get(qn(f'w:{attr}'))
                margins[attr] = round(int(val) / 1440, 2) if val else None  # twips to inches

        # Different first page
        title_pg = sec_props.find(qn('w:titlePg'))
        diff_first = title_pg is not None

        # Header/footer references
        header_refs = sec_props.findall(qn('w:headerReference'))
        footer_refs = sec_props.findall(qn('w:footerReference'))

        sections.append({
            'index': i,
            'number': i + 1,
            'break_type': break_type,
            'page_format': page_fmt,
            'page_start': page_start,
            'width_inches': round(width_emu / 1440, 2) if width_emu else None,
            'height_inches': round(height_emu / 1440, 2) if height_emu else None,
            'margins': margins,
            'different_first_page': diff_first,
            'header_count': len(header_refs),
            'footer_count': len(footer_refs),
        })
    return sections


def get_heading_outline(doc: Document) -> list[dict]:
    """Extract document outline from headings."""
    outline = []
    for i, para in enumerate(doc.paragraphs):
        style_name = para.style.name if para.style else "None"
        text = para.text.strip()
        if not text:
            continue

        level = None
        if 'Heading: Main' in style_name:
            level = 2
        elif style_name == 'Heading 3':
            level = 3
        elif style_name == 'Heading 4':
            level = 4
        elif style_name.startswith('Heading'):
            try:
                level = int(style_name.split()[-1])
            except ValueError:
                pass

        if level is not None:
            outline.append({
                'paragraph_index': i,
                'level': level,
                'style': style_name,
                'text': text,
            })
    return outline


def get_style_summary(doc: Document) -> dict[str, int]:
    """Count paragraphs by style."""
    counts: dict[str, int] = {}
    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "(No Style)"
        counts[style_name] = counts.get(style_name, 0) + 1
    return dict(sorted(counts.items(), key=lambda x: -x[1]))


def get_paragraphs_in_section(doc: Document, section_index: int) -> list[dict]:
    """Get all paragraphs belonging to a specific section.

    Word stores section properties at the end of the last paragraph of each section
    (in the paragraph's pPr/sectPr) except for the last section which is in the body's sectPr.
    """
    body = doc.element.body
    paragraphs = body.findall(qn('w:p'))
    tables = body.findall(qn('w:tbl'))

    # Find section boundaries by looking for sectPr in paragraphs
    section_breaks = []
    for idx, p in enumerate(paragraphs):
        pPr = p.find(qn('w:pPr'))
        if pPr is not None:
            sectPr = pPr.find(qn('w:sectPr'))
            if sectPr is not None:
                section_breaks.append(idx)

    # Build section ranges
    start = 0
    section_ranges = []
    for break_idx in section_breaks:
        section_ranges.append((start, break_idx + 1))
        start = break_idx + 1
    # Last section (body-level sectPr)
    section_ranges.append((start, len(paragraphs)))

    if section_index < 0 or section_index >= len(section_ranges):
        return []

    s, e = section_ranges[section_index]
    result = []
    for idx in range(s, e):
        p = paragraphs[idx]
        # Extract text directly from XML
        texts = []
        for r in p.findall(qn('w:r')):
            for t in r.findall(qn('w:t')):
                if t.text:
                    texts.append(t.text)
        text = ''.join(texts)

        # Get style from XML
        pPr = p.find(qn('w:pPr'))
        style = "(No Style)"
        if pPr is not None:
            pStyle = pPr.find(qn('w:pStyle'))
            if pStyle is not None:
                style = pStyle.get(qn('w:val'), "(No Style)")

        result.append({
            'index': idx,
            'style': style,
            'text': text[:100] + ('...' if len(text) > 100 else ''),
        })
    return result


def get_header_footer_content(doc: Document) -> list[dict]:
    """Extract header and footer text content per section."""
    results = []
    for i, section in enumerate(doc.sections):
        sec_info = {'section': i + 1, 'headers': {}, 'footers': {}}

        for hf_type in ['header', 'footer']:
            for variant in ['first_page', 'even_page', '']:
                attr_name = f"{variant}_{hf_type}" if variant else hf_type
                try:
                    obj = getattr(section, attr_name)
                    if obj and obj.paragraphs:
                        text = ' | '.join(p.text for p in obj.paragraphs if p.text.strip())
                        if text:
                            sec_info[f'{hf_type}s'][attr_name] = text
                except Exception:
                    pass

        results.append(sec_info)
    return results


def print_full_structure(filepath: Path) -> None:
    """Print comprehensive document structure report."""
    doc = load_document(filepath)

    print("=" * 70)
    print("THESIS DOCUMENT STRUCTURE REPORT")
    print("=" * 70)

    # Basic stats
    styles = get_style_summary(doc)
    print(f"\nTotal paragraphs: {len(doc.paragraphs)}")
    print(f"Total sections: {len(doc.sections)}")
    print(f"Total tables: {len(doc.tables)}")

    # Styles
    print(f"\n{'─' * 50}")
    print("STYLES USED:")
    print(f"{'─' * 50}")
    for style, count in styles.items():
        print(f"  {count:>4}x  {style}")

    # Sections
    sections = get_section_info(doc)
    print(f"\n{'─' * 50}")
    print("SECTIONS:")
    print(f"{'─' * 50}")
    print(f"{'#':>3} {'Break':>12} {'PageFmt':>12} {'Start':>6} {'TopMargin':>10} {'DiffFirst':>10}")
    print(f"{'─' * 3} {'─' * 12} {'─' * 12} {'─' * 6} {'─' * 10} {'─' * 10}")
    for s in sections:
        fmt = s['page_format'] or '(inherit)'
        start = str(s['page_start']) if s['page_start'] is not None else '-'
        top = f"{s['margins'].get('top', '-')}in" if s['margins'].get('top') else '-'
        diff = 'Yes' if s['different_first_page'] else 'No'
        print(f"{s['number']:>3} {s['break_type']:>12} {fmt:>12} {start:>6} {top:>10} {diff:>10}")

    # Outline
    outline = get_heading_outline(doc)
    print(f"\n{'─' * 50}")
    print("DOCUMENT OUTLINE:")
    print(f"{'─' * 50}")
    for h in outline:
        indent = "  " * (h['level'] - 1)
        print(f"  {indent}[H{h['level']}] {h['text'][:80]}")

    print(f"\n{'=' * 70}")
    print("END OF REPORT")
    print(f"{'=' * 70}")


if __name__ == '__main__':
    import sys
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from tools import WORKING_FILE, ensure_working_copy

    ensure_working_copy()
    print_full_structure(WORKING_FILE)
