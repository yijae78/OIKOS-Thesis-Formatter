"""Section manager — reorder, move, and manipulate document sections."""

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pathlib import Path
from copy import deepcopy
import lxml.etree as etree


def get_section_boundaries(doc: Document) -> list[dict]:
    """Find the XML element boundaries for each section.

    Returns a list of dicts with start/end element indices in the body,
    and the section properties element.
    """
    body = doc.element.body
    children = list(body)  # All direct children (paragraphs, tables, etc.)

    # Sections are delimited by sectPr elements:
    # - In paragraph pPr (for all sections except the last)
    # - In body (for the last section)
    boundaries = []
    current_start = 0

    for idx, child in enumerate(children):
        # Check if this element contains a section break
        if child.tag == qn('w:p'):
            pPr = child.find(qn('w:pPr'))
            if pPr is not None:
                sectPr = pPr.find(qn('w:sectPr'))
                if sectPr is not None:
                    boundaries.append({
                        'section_index': len(boundaries),
                        'start_child': current_start,
                        'end_child': idx,  # inclusive — this paragraph ends the section
                        'sectPr_location': 'paragraph',
                        'element_count': idx - current_start + 1,
                    })
                    current_start = idx + 1

    # Last section (body-level sectPr)
    body_sectPr = body.find(qn('w:sectPr'))
    if body_sectPr is not None:
        boundaries.append({
            'section_index': len(boundaries),
            'start_child': current_start,
            'end_child': len(children) - 1,  # -1 because body sectPr is separate
            'sectPr_location': 'body',
            'element_count': len(children) - current_start,
        })

    return boundaries


def get_section_content_summary(doc: Document) -> list[dict]:
    """Get a summary of content in each section."""
    boundaries = get_section_boundaries(doc)
    body = doc.element.body
    children = list(body)

    summaries = []
    for sec in boundaries:
        paragraphs = []
        tables = 0

        for i in range(sec['start_child'], min(sec['end_child'] + 1, len(children))):
            child = children[i]
            if child.tag == qn('w:p'):
                # Extract text directly from XML to avoid 'part' attribute issues
                texts = []
                for r in child.findall(qn('w:r')):
                    for t in r.findall(qn('w:t')):
                        if t.text:
                            texts.append(t.text)
                text = ''.join(texts).strip()

                # Get style name from XML
                pPr = child.find(qn('w:pPr'))
                style = "(No Style)"
                if pPr is not None:
                    pStyle = pPr.find(qn('w:pStyle'))
                    if pStyle is not None:
                        style = pStyle.get(qn('w:val'), "(No Style)")

                if text:
                    paragraphs.append({
                        'style': style,
                        'text': text[:80],
                    })
            elif child.tag == qn('w:tbl'):
                tables += 1

        # Get first heading as section title
        title = None
        for p in paragraphs:
            if 'Heading' in p['style'] or 'heading' in p['style'].lower():
                title = p['text']
                break

        summaries.append({
            'section': sec['section_index'] + 1,
            'title': title or '(no heading)',
            'paragraph_count': len(paragraphs),
            'table_count': tables,
            'element_count': sec['element_count'],
            'first_text': paragraphs[0]['text'] if paragraphs else '(empty)',
        })

    return summaries


def swap_sections(doc: Document, section_a: int, section_b: int) -> None:
    """Swap two sections (1-based section numbers).

    This is a complex operation that moves all XML elements between sections.
    IMPORTANT: Always backup before calling this.
    """
    if section_a == section_b:
        return

    boundaries = get_section_boundaries(doc)
    total = len(boundaries)

    # Convert to 0-based
    idx_a = section_a - 1
    idx_b = section_b - 1

    if idx_a < 0 or idx_a >= total or idx_b < 0 or idx_b >= total:
        raise ValueError(f"Section numbers must be 1-{total}. Got {section_a} and {section_b}")

    # Ensure a < b for consistent ordering
    if idx_a > idx_b:
        idx_a, idx_b = idx_b, idx_a

    body = doc.element.body
    children = list(body)

    sec_a = boundaries[idx_a]
    sec_b = boundaries[idx_b]

    # Extract elements for both sections (deep copy)
    elements_a = [deepcopy(children[i]) for i in range(sec_a['start_child'], sec_a['end_child'] + 1)]
    elements_b = [deepcopy(children[i]) for i in range(sec_b['start_child'], sec_b['end_child'] + 1)]

    # Replace section B's elements with section A's
    # (do B first since it comes later, so indices stay valid)
    for i in range(sec_b['end_child'], sec_b['start_child'] - 1, -1):
        body.remove(children[i])

    insert_point_b = sec_b['start_child']
    # After removing B elements, if B was after A, insertion point might shift
    # but since we work on the actual tree and haven't touched A yet, we need
    # to recalculate
    current_children = list(body)
    # Find where B used to start
    # Since we removed B elements, we insert A elements at the same position
    for i, elem in enumerate(elements_a):
        if insert_point_b + i < len(current_children):
            current_children[insert_point_b + i].addprevious(elem)
        else:
            body.append(elem)

    # Now replace section A's elements with section B's
    current_children = list(body)
    for i in range(sec_a['end_child'], sec_a['start_child'] - 1, -1):
        if i < len(current_children):
            body.remove(current_children[i])

    current_children = list(body)
    for i, elem in enumerate(elements_b):
        if sec_a['start_child'] + i < len(current_children):
            current_children[sec_a['start_child'] + i].addprevious(elem)
        else:
            body.append(elem)

    print(f"[OK] Swapped section {section_a} ↔ section {section_b}")


def move_section(doc: Document, from_section: int, to_position: int) -> None:
    """Move a section to a new position (1-based).

    Extracts all elements of the source section and reinserts them at the target position.
    IMPORTANT: Always backup before calling this.
    """
    boundaries = get_section_boundaries(doc)
    total = len(boundaries)

    from_idx = from_section - 1
    to_idx = to_position - 1

    if from_idx < 0 or from_idx >= total:
        raise ValueError(f"Source section {from_section} out of range (1-{total})")
    if to_idx < 0 or to_idx >= total:
        raise ValueError(f"Target position {to_position} out of range (1-{total})")
    if from_idx == to_idx:
        print(f"[SKIP] Section {from_section} is already at position {to_position}")
        return

    body = doc.element.body
    children = list(body)

    src = boundaries[from_idx]

    # Extract source elements
    extracted = []
    for i in range(src['start_child'], src['end_child'] + 1):
        extracted.append(children[i])

    # Remove from current position
    for elem in extracted:
        body.remove(elem)

    # Recalculate boundaries after removal
    boundaries_new = get_section_boundaries(doc)
    current_children = list(body)

    # Find insertion point
    if to_idx == 0:
        # Insert at the very beginning
        if current_children:
            ref_elem = current_children[0]
            for elem in reversed(extracted):
                ref_elem.addprevious(elem)
        else:
            for elem in extracted:
                body.append(elem)
    elif to_idx >= len(boundaries_new):
        # Insert at the end (before body sectPr)
        body_sectPr = body.find(qn('w:sectPr'))
        if body_sectPr is not None:
            for elem in reversed(extracted):
                body_sectPr.addprevious(elem)
        else:
            for elem in extracted:
                body.append(elem)
    else:
        # Insert before the target section
        target = boundaries_new[to_idx]
        current_children = list(body)
        ref_elem = current_children[target['start_child']]
        for elem in reversed(extracted):
            ref_elem.addprevious(elem)

    print(f"[OK] Moved section {from_section} → position {to_position}")


def change_section_break_type(doc: Document, section_index: int, break_type: str) -> None:
    """Change the break type for a section (0-based index).

    Valid break types: nextPage, continuous, evenPage, oddPage, nextColumn
    """
    valid_types = ['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn']
    if break_type not in valid_types:
        raise ValueError(f"Invalid break type '{break_type}'. Valid: {valid_types}")

    if section_index < 0 or section_index >= len(doc.sections):
        raise ValueError(f"Section index {section_index} out of range")

    section = doc.sections[section_index]
    sec_props = section._sectPr

    sec_type = sec_props.find(qn('w:type'))
    if sec_type is None:
        sec_type = OxmlElement('w:type')
        sec_props.insert(0, sec_type)

    sec_type.set(qn('w:val'), break_type)
    print(f"[OK] Section {section_index + 1}: break type → {break_type}")


def print_section_summary(doc: Document) -> None:
    """Print formatted summary of all sections with content preview."""
    summaries = get_section_content_summary(doc)

    print(f"\n{'─' * 70}")
    print("SECTION CONTENT SUMMARY")
    print(f"{'─' * 70}")
    print(f"{'#':>3} {'Title':<35} {'Paras':>6} {'Tables':>7} {'First Text'}")
    print(f"{'─' * 3} {'─' * 35} {'─' * 6} {'─' * 7} {'─' * 25}")

    for s in summaries:
        title = s['title'][:33] + '..' if len(s['title']) > 35 else s['title']
        first = s['first_text'][:23] + '..' if len(s['first_text']) > 25 else s['first_text']
        # Handle encoding issues with special characters
        line = f"{s['section']:>3} {title:<35} {s['paragraph_count']:>6} {s['table_count']:>7} {first}"
        try:
            print(line)
        except UnicodeEncodeError:
            print(line.encode('utf-8', errors='replace').decode('utf-8', errors='replace'))
