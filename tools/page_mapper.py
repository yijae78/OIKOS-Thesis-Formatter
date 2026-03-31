"""Map document headings to exact PDF page numbers using text search."""

import re
import fitz
from pathlib import Path


def _extract_search_keys(text: str) -> list[str]:
    """Extract multiple search keys from heading text, ordered by specificity."""
    keys = []
    text = text.strip()
    if not text or len(text) < 2:
        return keys

    # Korean chapter: "제 1 장서론 (Introduction)" → search "제 1 장"
    m = re.match(r'(제\s*\d+\s*장)', text)
    if m:
        keys.append(m.group(1))

    # English chapter: "CHAPTER 1INTRODUCTION" → search "CHAPTER 1"
    m = re.match(r'(CHAPTER\s*\d+)', text, re.IGNORECASE)
    if m:
        keys.append(m.group(1))

    # Korean keywords that are unique enough
    for kw in ['요약', '결론', '제언', '서론', '감사의 말', '헌정', '개요', '목차',
               '인용 문헌', '약어 목록', '표 목록', '그림 목록']:
        if kw in text:
            keys.append(kw)

    # English keywords
    for kw in ['ABSTRACT', 'TABLE OF CONTENTS', 'ENGLISH SUMMARY',
               'REFERENCES CITED', 'INTRODUCTION', 'LITERATURE REVIEW',
               'METHODOLOGY', 'CASE ANALYSIS', 'DISCUSSION', 'CONCLUSION',
               'RECOMMENDATIONS', 'DEDICATION', 'ACKNOWLEDGEMENTS',
               'LIST OF TABLES', 'LIST OF FIGURES', 'LIST OF ABBREVIATIONS']:
        if kw.lower() in text.upper():
            keys.append(kw)

    # Subsection numbers: "2.1.", "3.1.1." etc.
    m = re.match(r'(\d+\.\d+[\.\d]*)', text)
    if m:
        keys.append(m.group(1))

    # Fallback: first 15 chars, then first 8 chars
    if len(text) >= 15:
        keys.append(text[:15])
    if len(text) >= 8:
        keys.append(text[:8])

    return keys


def build_heading_page_map(pdf_path: Path, outline: list) -> dict:
    """Build mapping: paragraph_index → PDF page number.

    Searches each heading's text in the PDF to find its exact page.
    Returns dict {para_index: pdf_page_number (1-based)}.
    """
    pdf = fitz.open(str(pdf_path))
    heading_map = {}

    # Pre-extract page text for faster searching
    page_texts = []
    for pg_num in range(pdf.page_count):
        page_texts.append(pdf[pg_num].get_text().upper())

    for entry in outline:
        text = entry.get('display_text', entry['text']).strip()
        if not text or len(text) < 2:
            continue

        search_keys = _extract_search_keys(text)
        found = False

        for key in search_keys:
            if found:
                break
            key_upper = key.upper()
            # First: fast text-in-page search (case insensitive)
            for pg_num, pg_text in enumerate(page_texts):
                if key_upper in pg_text:
                    heading_map[entry['i']] = pg_num + 1
                    found = True
                    break

        if not found:
            # Last resort: fitz search_for with original text[:20]
            short = text[:20]
            for pg_num in range(pdf.page_count):
                results = pdf[pg_num].search_for(short)
                if results:
                    heading_map[entry['i']] = pg_num + 1
                    break

    pdf.close()
    return heading_map


def build_section_page_map(pdf_path: Path) -> dict:
    """Build mapping: section_keyword → PDF page number for major sections."""
    pdf = fitz.open(str(pdf_path))
    section_map = {}

    patterns = [
        (r'제\s*(\d+)\s*장', 'chapter'),
        (r'CHAPTER\s*(\d+)', 'en_chapter'),
        (r'ABSTRACT', 'abstract'),
        (r'개요', 'abstract_kr'),
        (r'목차', 'toc'),
        (r'ENGLISH\s*SUMMARY', 'en_summary'),
        (r'인용\s*문헌', 'references'),
        (r'부록', 'appendix'),
    ]

    for page_num in range(pdf.page_count):
        page = pdf[page_num]
        blocks = page.get_text('blocks')
        for block in blocks[:5]:  # Only check top blocks
            text = block[4].strip() if len(block) > 4 else ''
            y_pos = block[1] if len(block) > 1 else 999

            if y_pos > 250:  # Skip if not near top
                continue

            for pattern, key_type in patterns:
                m = re.search(pattern, text)
                if m:
                    if key_type == 'chapter':
                        key = f'ch{m.group(1)}'
                    elif key_type == 'en_chapter':
                        key = f'en_ch{m.group(1)}'
                    else:
                        key = key_type
                    if key not in section_map:
                        section_map[key] = page_num + 1

    pdf.close()
    return section_map
