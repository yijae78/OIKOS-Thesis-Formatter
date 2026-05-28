"""Fix page numbering after cover page insertion.
- S0-S2 (cover): no page numbers, unlinked empty footers
- S3 (front matter start): lowerRoman start=1 (reset after cover)
- S21 (body start): decimal start=1 (already set, verify)
- Update cached TOC page numbers"""
import os
import shutil
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from lxml import etree

BASE = os.path.join(os.path.dirname(__file__), '..')
WORKING = os.path.join(BASE, 'output', 'thesis_english.docx')


def ensure_empty_footer(section):
    """Ensure section has its own empty footer (not linked to previous)."""
    # Unlink footer from previous
    footer = section.footer
    footer.is_linked_to_previous = False
    # Clear all paragraphs
    for para in footer.paragraphs:
        for child in list(para._element):
            if child.tag != qn('w:pPr'):
                para._element.remove(child)

    # Unlink header too
    header = section.header
    header.is_linked_to_previous = False
    for para in header.paragraphs:
        for child in list(para._element):
            if child.tag != qn('w:pPr'):
                para._element.remove(child)


def set_page_number_start(section, fmt, start):
    """Set page number format and start value for a section."""
    sectPr = section._sectPr
    # Remove existing pgNumType
    for pnt in sectPr.findall(qn('w:pgNumType')):
        sectPr.remove(pnt)
    # Add new pgNumType
    pnt = etree.SubElement(sectPr, qn('w:pgNumType'))
    pnt.set(qn('w:fmt'), fmt)
    pnt.set(qn('w:start'), str(start))


def remove_page_number_format(section):
    """Remove page number type entirely (no numbering)."""
    sectPr = section._sectPr
    for pnt in sectPr.findall(qn('w:pgNumType')):
        sectPr.remove(pnt)


def main():
    # Backup
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup = os.path.join(BASE, 'backups', f'thesis_pre_pgnum_{ts}.docx')
    shutil.copy2(WORKING, backup)
    print(f"Backup: {os.path.basename(backup)}")

    doc = Document(WORKING)

    # === 1. Fix cover sections S0-S2: no page numbers ===
    print("\n--- Fixing cover sections S0-S2 ---")
    for i in range(3):
        s = doc.sections[i]
        ensure_empty_footer(s)
        remove_page_number_format(s)
        print(f"  S{i}: footer unlinked & cleared, pgNumType removed")

    # === 2. Fix S3: front matter starts at page i ===
    print("\n--- Fixing S3 (front matter start) ---")
    set_page_number_start(doc.sections[3], 'lowerRoman', 1)
    print("  S3: pgNumType set to lowerRoman start=1")

    # === 3. Verify S21 (body start): decimal start=1 ===
    print("\n--- Checking S21 (body start) ---")
    s21 = doc.sections[21]
    sectPr = s21._sectPr
    pnt = sectPr.find(qn('w:pgNumType'))
    if pnt is not None:
        fmt = pnt.get(qn('w:fmt'), '-')
        start = pnt.get(qn('w:start'), '-')
        print(f"  S21: pgNumType fmt={fmt} start={start}")
        if start != '1':
            pnt.set(qn('w:start'), '1')
            print(f"  S21: start corrected to 1")
    else:
        # Need to add decimal start=1
        set_page_number_start(s21, 'decimal', 1)
        print("  S21: pgNumType added decimal start=1")

    # === 4. Update cached page number values in footers ===
    # The PAGE field cached values need updating.
    # We can't calculate real page numbers, but we can set the start values
    # so Word renders them correctly when opened.
    print("\n--- Updating cached footer page numbers ---")
    # Find all footer parts and update cached PAGE field values
    updated_count = 0
    for i, s in enumerate(doc.sections):
        sectPr = s._sectPr
        ftr_refs = sectPr.findall(qn('w:footerReference'))
        for fref in ftr_refs:
            rid = fref.get(qn('r:id'))
            try:
                ftr_part = s.part.rels[rid].target_part
                ftr_xml = ftr_part.blob
                tree = etree.fromstring(ftr_xml)
                # Find cached text elements in PAGE fields
                # PAGE fields have fldChar(begin) → instrText → fldChar(separate) → text → fldChar(end)
                changed = False
                runs = tree.findall(f'.//{qn("w:r")}')
                in_field = False
                after_separate = False
                for run in runs:
                    fld_chars = run.findall(qn('w:fldChar'))
                    for fc in fld_chars:
                        fld_type = fc.get(qn('w:fldCharType'))
                        if fld_type == 'begin':
                            in_field = True
                            after_separate = False
                        elif fld_type == 'separate':
                            after_separate = True
                        elif fld_type == 'end':
                            in_field = False
                            after_separate = False

                    # If we're after the separator, the text is the cached value
                    if after_separate and not fld_chars:
                        t_elems = run.findall(qn('w:t'))
                        for t in t_elems:
                            # Clear cached value so Word recalculates on open
                            if t.text and t.text.strip():
                                t.text = ''
                                changed = True

                if changed:
                    ftr_part._blob = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
                    updated_count += 1
            except Exception as e:
                pass

    print(f"  Cleared cached values in {updated_count} footers")

    # Also clear cached values in headers (page numbers sometimes in headers)
    for i, s in enumerate(doc.sections):
        sectPr = s._sectPr
        hdr_refs = sectPr.findall(qn('w:headerReference'))
        for href in hdr_refs:
            rid = href.get(qn('r:id'))
            try:
                hdr_part = s.part.rels[rid].target_part
                hdr_xml = hdr_part.blob
                tree = etree.fromstring(hdr_xml)
                changed = False
                runs = tree.findall(f'.//{qn("w:r")}')
                in_field = False
                after_separate = False
                for run in runs:
                    fld_chars = run.findall(qn('w:fldChar'))
                    for fc in fld_chars:
                        fld_type = fc.get(qn('w:fldCharType'))
                        if fld_type == 'begin':
                            in_field = True
                            after_separate = False
                        elif fld_type == 'separate':
                            after_separate = True
                        elif fld_type == 'end':
                            in_field = False
                            after_separate = False
                    if after_separate and not fld_chars:
                        t_elems = run.findall(qn('w:t'))
                        for t in t_elems:
                            if t.text and t.text.strip():
                                t.text = ''
                                changed = True
                if changed:
                    hdr_part._blob = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
                    updated_count += 1
            except:
                pass

    # Save
    doc.save(WORKING)
    print(f"\nSaved: {os.path.basename(WORKING)}")

    # Verify
    doc2 = Document(WORKING)
    print("\n--- Verification ---")
    for i in range(min(5, len(doc2.sections))):
        s = doc2.sections[i]
        sectPr = s._sectPr
        brk = sectPr.find(qn('w:type'))
        brk_val = brk.get(qn('w:val')) if brk is not None else 'nextPage'
        pnt = sectPr.find(qn('w:pgNumType'))
        pg_info = ''
        if pnt is not None:
            pg_info = f" pgNum={pnt.get(qn('w:fmt'),'?')}/start={pnt.get(qn('w:start'),'?')}"
        f_link = s.footer.is_linked_to_previous
        print(f"  S{i}: break={brk_val}{pg_info} f_link={f_link}")

    # Check S21
    s21 = doc2.sections[21]
    pnt = s21._sectPr.find(qn('w:pgNumType'))
    if pnt is not None:
        print(f"  S21: pgNum={pnt.get(qn('w:fmt'),'?')}/start={pnt.get(qn('w:start'),'?')}")


if __name__ == '__main__':
    main()
