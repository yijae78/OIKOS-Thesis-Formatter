"""Insert 3 cover/signature pages before the existing document content.
Page 1: Cover page (English only)
Page 2: Inner cover (OIKOS University)
Page 3: Dissertation Approval Sheet with 5 signatures"""
import os
import shutil
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, Cm, Emu
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.image import Image as DocxImage
from lxml import etree

BASE = os.path.join(os.path.dirname(__file__), '..')
WORKING = os.path.join(BASE, 'output', 'thesis_english.docx')
SIG_DIR = os.path.join(BASE, '_temp_figs')


def make_sectPr_xml(break_type='nextPage'):
    """Create section properties XML matching existing document."""
    return f'''<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:type w:val="{break_type}"/>
        <w:pgSz w:w="12240" w:h="15840"/>
        <w:pgMar w:top="2880" w:right="1440" w:bottom="1440"
                 w:left="2160" w:header="1080" w:footer="1080" w:gutter="0"/>
    </w:sectPr>'''


def add_centered_text(para, text, font_size=12, bold=False, italic=False):
    """Add centered text to a paragraph."""
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(font_size)
    run.bold = bold
    run.italic = italic
    return run


def add_empty_para(body, ref_element):
    """Insert an empty paragraph before ref_element."""
    p = etree.SubElement(etree.Element('tmp'), qn('w:p'))
    body.insert(list(body).index(ref_element), p)
    return p


def insert_image_in_element(doc, p_elem, image_path, width_inches):
    """Insert an inline image into a paragraph element."""
    img = DocxImage.from_file(image_path)
    width_emu = int(Inches(width_inches))
    height_emu = int(width_emu * img.px_height / img.px_width)

    rId, _ = doc.part.get_or_add_image(image_path)
    pic_id = abs(hash(image_path)) % 100000 + 1
    fname = os.path.basename(image_path)

    nsmap = {
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }

    inline = etree.SubElement(
        etree.Element('tmp'),
        '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline',
        nsmap=nsmap
    )
    inline.set('distT', '0')
    inline.set('distB', '0')
    inline.set('distL', '0')
    inline.set('distR', '0')

    extent = etree.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
    extent.set('cx', str(width_emu))
    extent.set('cy', str(height_emu))

    effect = etree.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}effectExtent')
    for attr in ['l', 't', 'r', 'b']:
        effect.set(attr, '0')

    docPr = etree.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr')
    docPr.set('id', str(pic_id))
    docPr.set('name', fname)

    graphic = etree.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/main}graphic')
    graphicData = etree.SubElement(graphic, '{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData')
    graphicData.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

    pic = etree.SubElement(graphicData, '{http://schemas.openxmlformats.org/drawingml/2006/picture}pic')
    nvPicPr = etree.SubElement(pic, '{http://schemas.openxmlformats.org/drawingml/2006/picture}nvPicPr')
    cNvPr = etree.SubElement(nvPicPr, '{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPr')
    cNvPr.set('id', '0')
    cNvPr.set('name', fname)
    etree.SubElement(nvPicPr, '{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPicPr')

    blipFill = etree.SubElement(pic, '{http://schemas.openxmlformats.org/drawingml/2006/picture}blipFill')
    blip = etree.SubElement(blipFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
    blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', rId)
    stretch = etree.SubElement(blipFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}stretch')
    etree.SubElement(stretch, '{http://schemas.openxmlformats.org/drawingml/2006/main}fillRect')

    spPr = etree.SubElement(pic, '{http://schemas.openxmlformats.org/drawingml/2006/picture}spPr')
    xfrm = etree.SubElement(spPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
    off = etree.SubElement(xfrm, '{http://schemas.openxmlformats.org/drawingml/2006/main}off')
    off.set('x', '0')
    off.set('y', '0')
    ext = etree.SubElement(xfrm, '{http://schemas.openxmlformats.org/drawingml/2006/main}ext')
    ext.set('cx', str(width_emu))
    ext.set('cy', str(height_emu))
    prstGeom = etree.SubElement(spPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom')
    prstGeom.set('prst', 'rect')
    etree.SubElement(prstGeom, '{http://schemas.openxmlformats.org/drawingml/2006/main}avLst')

    run_elem = etree.SubElement(p_elem, qn('w:r'))
    drawing_elem = etree.SubElement(run_elem, qn('w:drawing'))
    drawing_elem.append(inline)


def create_paragraph(body, ref, text, align='center', font_size=12, bold=False, italic=False, space_before=0, space_after=0):
    """Create and insert a paragraph before ref element."""
    p = etree.SubElement(etree.Element('tmp'), qn('w:p'))
    body.insert(list(body).index(ref), p)

    # Paragraph properties
    pPr = etree.SubElement(p, qn('w:pPr'))
    if align == 'center':
        jc = etree.SubElement(pPr, qn('w:jc'))
        jc.set(qn('w:val'), 'center')
    if space_before or space_after:
        spacing = etree.SubElement(pPr, qn('w:spacing'))
        if space_before:
            spacing.set(qn('w:before'), str(space_before))
        if space_after:
            spacing.set(qn('w:after'), str(space_after))

    if text:
        # Run
        r = etree.SubElement(p, qn('w:r'))
        rPr = etree.SubElement(r, qn('w:rPr'))
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        sz = etree.SubElement(rPr, qn('w:sz'))
        sz.set(qn('w:val'), str(font_size * 2))  # half-points
        szCs = etree.SubElement(rPr, qn('w:szCs'))
        szCs.set(qn('w:val'), str(font_size * 2))
        if bold:
            etree.SubElement(rPr, qn('w:b'))
        if italic:
            etree.SubElement(rPr, qn('w:i'))
        t = etree.SubElement(r, qn('w:t'))
        t.text = text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    return p


def add_section_break(p_elem):
    """Add a nextPage section break to the last paragraph's properties."""
    pPr = p_elem.find(qn('w:pPr'))
    if pPr is None:
        pPr = etree.SubElement(p_elem, qn('w:pPr'))
        p_elem.insert(0, pPr)

    sectPr = etree.SubElement(pPr, qn('w:sectPr'))
    pg_type = etree.SubElement(sectPr, qn('w:type'))
    pg_type.set(qn('w:val'), 'nextPage')
    pgSz = etree.SubElement(sectPr, qn('w:pgSz'))
    pgSz.set(qn('w:w'), '12240')
    pgSz.set(qn('w:h'), '15840')
    pgMar = etree.SubElement(sectPr, qn('w:pgMar'))
    pgMar.set(qn('w:top'), '2880')
    pgMar.set(qn('w:right'), '1440')
    pgMar.set(qn('w:bottom'), '1440')
    pgMar.set(qn('w:left'), '2160')
    pgMar.set(qn('w:header'), '1080')
    pgMar.set(qn('w:footer'), '1080')
    pgMar.set(qn('w:gutter'), '0')


def main():
    # Backup
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup = os.path.join(BASE, 'backups', f'thesis_pre_cover_{ts}.docx')
    shutil.copy2(WORKING, backup)
    print(f"Backup: {os.path.basename(backup)}")

    doc = Document(WORKING)
    body = doc.element.body
    first_elem = body[0]  # Reference point for insertion

    # ============================================================
    # PAGE 1: Cover Page (외표지 - English only)
    # ============================================================
    print("Creating Page 1: Cover...")

    # Spacer at top
    for _ in range(4):
        create_paragraph(body, first_elem, '', font_size=12, space_after=200)

    # Title (quoted)
    create_paragraph(body, first_elem,
        '\u300cA Study on the Mechanism by Which AI Utilization Training Requests',
        font_size=14, italic=True, space_after=0)
    create_paragraph(body, first_elem,
        'Fail to Translate into Organizational Performance',
        font_size=14, italic=True, space_after=0)
    create_paragraph(body, first_elem,
        '\u2013 Focusing on Seizing within Dynamic Capabilities Theory\u300d',
        font_size=14, italic=True, space_after=600)

    # Spacer
    create_paragraph(body, first_elem, '', font_size=12, space_after=400)

    # By
    create_paragraph(body, first_elem, 'By', font_size=12, space_after=200)
    create_paragraph(body, first_elem, 'Yijae Shin', font_size=14, bold=True, space_after=600)

    # University info
    create_paragraph(body, first_elem, 'A Dissertation Presented to the Faculty of',
                     font_size=12, space_after=0)
    create_paragraph(body, first_elem, 'OIKOS UNIVERSITY',
                     font_size=14, bold=True, space_after=200)
    create_paragraph(body, first_elem, 'In Partial Fulfillment of the',
                     font_size=12, space_after=0)
    create_paragraph(body, first_elem,
        'Requirements for the Degree of Doctor of Business Administration',
        font_size=12, space_after=600)

    # Date
    create_paragraph(body, first_elem, '', font_size=12, space_after=200)
    p_last1 = create_paragraph(body, first_elem, 'February 2026',
                               font_size=12, space_after=0)

    # Section break after page 1
    add_section_break(p_last1)
    print("  Page 1 created")

    # ============================================================
    # PAGE 2: Inner Cover (내표지)
    # ============================================================
    print("Creating Page 2: Inner Cover...")

    for _ in range(2):
        create_paragraph(body, first_elem, '', font_size=12, space_after=200)

    create_paragraph(body, first_elem, 'OIKOS UNIVERSITY',
                     font_size=16, bold=True, space_after=400)

    create_paragraph(body, first_elem,
        '\u201cA Study on the Mechanism by Which AI Utilization Training Requests',
        font_size=12, space_after=0)
    create_paragraph(body, first_elem,
        'Fail to Translate into Organizational Performance\u2013 Focusing',
        font_size=12, space_after=0)
    create_paragraph(body, first_elem,
        'on Seizing within Dynamic Capabilities Theory\u201d',
        font_size=12, space_after=600)

    create_paragraph(body, first_elem, 'A DISSERTATION', font_size=12, bold=True, space_after=100)
    create_paragraph(body, first_elem, 'SUBMITTED TO THE FACULTY OF', font_size=12, space_after=100)
    create_paragraph(body, first_elem, 'OIKOS UNIVERSITY', font_size=12, bold=True, space_after=100)
    create_paragraph(body, first_elem, 'IN CANDIDACY FOR THE DEGREE OF', font_size=12, space_after=100)
    create_paragraph(body, first_elem, 'DOCTOR OF BUSINESS ADMINISTRATION', font_size=12, bold=True, space_after=400)

    create_paragraph(body, first_elem, 'by', font_size=12, space_after=100)
    create_paragraph(body, first_elem, 'Yijae Shin', font_size=14, bold=True, space_after=400)

    create_paragraph(body, first_elem, 'OAKLAND, CALIFORNIA', font_size=12, space_after=100)
    create_paragraph(body, first_elem, 'May 2026', font_size=12, space_after=400)

    create_paragraph(body, first_elem, 'Copyright 2026', font_size=10, space_after=0)
    create_paragraph(body, first_elem, 'Yijae Shin', font_size=10, space_after=0)
    p_last2 = create_paragraph(body, first_elem, 'ALL RIGHTS RESERVED', font_size=10, space_after=0)

    add_section_break(p_last2)
    print("  Page 2 created")

    # ============================================================
    # PAGE 3: Dissertation Approval Sheet (서명)
    # ============================================================
    print("Creating Page 3: Approval Sheet...")

    create_paragraph(body, first_elem, '', font_size=12, space_after=200)
    create_paragraph(body, first_elem, 'DISSERTATION APPROVAL SHEET',
                     font_size=14, bold=True, space_after=400)

    create_paragraph(body, first_elem, 'This dissertation, entitled', font_size=12, space_after=200)

    create_paragraph(body, first_elem,
        '\u201cA Study on the Mechanism by Which AI Utilization Training Requests',
        font_size=12, italic=True, space_after=0)
    create_paragraph(body, first_elem,
        'Fail to Translate into Organizational Performance\u2013 Focusing',
        font_size=12, italic=True, space_after=0)
    create_paragraph(body, first_elem,
        'on Seizing within Dynamic Capabilities Theory\u201d',
        font_size=12, italic=True, space_after=400)

    create_paragraph(body, first_elem,
        'And submitted in candidacy for the degree of', font_size=12, space_after=100)
    create_paragraph(body, first_elem,
        'Doctor of Business Administration', font_size=12, bold=True, space_after=200)

    create_paragraph(body, first_elem,
        'Has been read and approved', font_size=12, space_after=100)
    create_paragraph(body, first_elem,
        'by the undersigned members of the faculty of', font_size=12, space_after=100)
    create_paragraph(body, first_elem,
        'Oikos University', font_size=12, bold=True, space_after=400)

    # Signature lines with images
    sig_files = [
        ('signature_0_608x98.png', 'Chairman'),
        ('signature_1_718x140.png', 'Member'),
        ('signature_2_930x110.png', 'Member'),
        ('signature_3_638x88.png', 'Member'),
        ('signature_4_730x140.png', 'Member'),
    ]

    for sig_file, role in sig_files:
        sig_path = os.path.join(SIG_DIR, sig_file)
        # Create paragraph for signature image
        p_sig = create_paragraph(body, first_elem, '', font_size=12, space_after=0)
        # Insert signature image
        insert_image_in_element(doc, p_sig, sig_path, 2.0)
        # Role label
        create_paragraph(body, first_elem, role, font_size=11, space_after=200)

    # Date and author
    create_paragraph(body, first_elem, '', font_size=12, space_after=200)
    create_paragraph(body, first_elem, 'May 2026', font_size=12, space_after=100)
    p_last3 = create_paragraph(body, first_elem,
        'written by Yijae Shin', font_size=12, space_after=0)

    add_section_break(p_last3)
    print("  Page 3 created")

    # Save
    doc.save(WORKING)
    print(f"\nSaved: {os.path.basename(WORKING)}")

    # Verify
    doc2 = Document(WORKING)
    print(f"\nVerification:")
    print(f"  Sections: {len(doc2.sections)}")
    print(f"  Paragraphs: {len(doc2.paragraphs)}")
    # Check first few paragraphs
    for i in range(min(5, len(doc2.paragraphs))):
        p = doc2.paragraphs[i]
        txt = p.text.strip()[:60] if p.text.strip() else '(empty)'
        print(f"  P{i}: {txt}")
    # Check document opens correctly
    print(f"  Document opens: OK")


if __name__ == '__main__':
    main()
