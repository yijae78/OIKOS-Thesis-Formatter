"""Replace existing figures and insert missing ones into thesis_english.docx."""
import os
import shutil
from datetime import datetime
from docx import Document
from docx.shared import Inches, Emu
from docx.oxml.ns import qn
from lxml import etree

BASE = os.path.join(os.path.dirname(__file__), '..')
WORKING = os.path.join(BASE, 'output', 'thesis_english.docx')
FIG_DIR = os.path.join(BASE, 'output', 'figures')


def replace_image_blob(doc, rel_id, new_image_path):
    """Replace the binary blob of an existing image part."""
    rel = doc.part.rels[rel_id]
    with open(new_image_path, 'rb') as f:
        new_blob = f.read()
    rel.target_part._blob = new_blob
    print(f"  Replaced {rel_id} with {os.path.basename(new_image_path)} ({len(new_blob):,}b)")


def insert_picture_in_paragraph(doc, para, image_path, width_inches):
    """Insert a picture into an existing paragraph using python-docx internals."""
    from docx.image.image import Image as DocxImage

    # Get image dimensions
    img = DocxImage.from_file(image_path)
    width_emu = int(Inches(width_inches))
    height_emu = int(width_emu * img.px_height / img.px_width)

    # Add image to document package and get relationship ID
    rId, image_part = doc.part.get_or_add_image(image_path)

    # Build inline XML using qn() for proper namespace handling
    # Use a unique id
    pic_id = abs(hash(image_path)) % 100000 + 1
    fname = os.path.basename(image_path)

    # Namespaces
    nsmap = {
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }

    # Build element tree manually
    inline = etree.SubElement(
        etree.Element('tmp'), '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline',
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
    effect.set('l', '0')
    effect.set('t', '0')
    effect.set('r', '0')
    effect.set('b', '0')

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

    # Create w:r > w:drawing > wp:inline structure
    run_elem = etree.SubElement(para._element, qn('w:r'))
    drawing_elem = etree.SubElement(run_elem, qn('w:drawing'))
    drawing_elem.append(inline)

    print(f"  Inserted {fname} at paragraph (rId={rId}, {width_emu}x{height_emu} EMU)")


def main():
    # Backup
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup = os.path.join(BASE, 'backups', f'thesis_pre_figswap_{ts}.docx')
    shutil.copy2(WORKING, backup)
    print(f"Backup: {os.path.basename(backup)}")

    doc = Document(WORKING)

    # === 1. Replace existing images ===
    print("\n--- Replacing existing images ---")
    replace_image_blob(doc, 'rId75', os.path.join(FIG_DIR, 'fig_4_1.png'))
    replace_image_blob(doc, 'rId77', os.path.join(FIG_DIR, 'fig_4_2.png'))

    # === 2. Insert missing images ===
    print("\n--- Inserting missing images ---")

    # Fig 1.1: Caption P385-386, insert at P387 (empty)
    print("Figure 1.1 -> P387:")
    insert_picture_in_paragraph(doc, doc.paragraphs[387],
                                 os.path.join(FIG_DIR, 'fig_1_1.png'), 5.5)

    # Fig 5.1: Caption P745-746, insert at P744 (empty Normal before caption)
    print("Figure 5.1 -> P744:")
    insert_picture_in_paragraph(doc, doc.paragraphs[744],
                                 os.path.join(FIG_DIR, 'fig_5_1.png'), 5.0)

    # Fig 5.2: Caption P784-785, insert at P783 (empty Normal)
    print("Figure 5.2 -> P783:")
    insert_picture_in_paragraph(doc, doc.paragraphs[783],
                                 os.path.join(FIG_DIR, 'fig_5_2.png'), 5.5)

    # Fig 5.3: Caption P799-800, insert at P801 (empty)
    print("Figure 5.3 -> P801:")
    insert_picture_in_paragraph(doc, doc.paragraphs[801],
                                 os.path.join(FIG_DIR, 'fig_5_3.png'), 5.5)

    # Fig 6.1: Caption P844-845, insert at P846 (empty)
    print("Figure 6.1 -> P846:")
    insert_picture_in_paragraph(doc, doc.paragraphs[846],
                                 os.path.join(FIG_DIR, 'fig_6_1.png'), 5.0)

    # Save
    doc.save(WORKING)
    print(f"\nSaved: {os.path.basename(WORKING)}")

    # Verify
    doc2 = Document(WORKING)
    body = doc2.element.body
    drawings = body.findall(f'.//{qn("w:drawing")}')
    print(f"\nVerification:")
    print(f"  Total drawing elements: {len(drawings)}")
    for idx in [387, 669, 690, 744, 783, 801, 846]:
        p = doc2.paragraphs[idx]
        n = len(p._element.findall(f'.//{qn("w:drawing")}'))
        print(f"  P{idx}: {n} image(s)")


if __name__ == '__main__':
    main()
