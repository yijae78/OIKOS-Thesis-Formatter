"""Build final thesis PDF v4: All v3 fixes + user feedback fixes.
New fixes:
  #6 Signature page: bigger signatures (2.0in)
  #7 Table 3.1 top border
  #8 Figure 5.1/5.2 caption position (title above image)
  #9 Section 5.7 spacing
  #10 Appendix D: two images on same page
"""
import sys
import io
import os
import time
import shutil
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
MAIN_DOCX = os.path.join(BASE, 'output', 'thesis_english.docx')
COVER_DOCX = os.path.join(BASE, 'output', 'cover_pages.docx')
BODY_PDF = os.path.join(BASE, 'output', 'thesis_body.pdf')
COVER_PDF = os.path.join(BASE, 'output', 'thesis_cover.pdf')
FINAL_PDF = os.path.join(BASE, 'output', 'Shin_Yijae_DBA_Dissertation_OIKOS_2026.pdf')
FINAL_DOCX = os.path.join(BASE, 'output', 'Shin_Yijae_DBA_Dissertation_OIKOS_2026.docx')
FIG_DIR = os.path.join(BASE, 'output', 'figures')
SIG_DIR = os.path.join(BASE, '_temp_figs')
BACKUP_DIR = os.path.join(BASE, 'backups')
ORIGINAL_PDF = os.path.join(BASE, 'mydata(논문최종본)',
    '#신이재 DBA박사논문 최최종본---양면 제본용(2026.3.31.)--재수정본--전면쪽위치수정.pdf')


# =================================================================
# Fix #1 + #2 + #6: Cover with balanced title + BIGGER signatures
# =================================================================
def create_cover():
    """Create 3-page cover: balanced title, inner cover, compact approval with LARGE signatures."""
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_SECTION_START

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    def add_line(text, size=12, bold=False, space_before=0, space_after=0):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format
        pf.space_before = Pt(space_before)
        pf.space_after = Pt(space_after)
        pf.line_spacing = Pt(size + 2)
        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(size)
            run.bold = bold
        return p

    # === PAGE 1: COVER ===
    sec = doc.sections[0]
    sec.top_margin = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1)
    sec.right_margin = Inches(1)

    add_line('', space_after=30)
    add_line('\u300cA Study on the Mechanism by Which AI Utilization', 14, True, space_after=2)
    add_line('Training Requests Fail to Translate into', 14, True, space_after=2)
    add_line('Organizational Performance', 14, True, space_after=2)
    add_line('\u2013 Focusing on Seizing within Dynamic', 14, True, space_after=2)
    add_line('Capabilities Theory\u300d', 14, True, space_after=0)
    add_line('', space_after=30)
    add_line('By', 12, False, space_after=4)
    add_line('Yijae Shin', 12, False, space_after=0)
    add_line('', space_after=70)
    add_line('A Dissertation Presented to the Faculty of', 12, False, space_after=6)
    add_line('OIKOS UNIVERSITY', 14, True, space_after=6)
    add_line('In Partial Fulfillment of the', 12, False, space_after=6)
    add_line('Requirements for the Degree of', 12, False, space_after=2)
    add_line('Doctor of Business Administration', 12, False, space_after=0)
    add_line('', space_after=190)
    add_line('February 2026', 12, False)

    # === PAGE 2: INNER COVER ===
    sec2 = doc.add_section(WD_SECTION_START.NEW_PAGE)
    sec2.top_margin = Inches(1)
    sec2.bottom_margin = Inches(1)
    sec2.left_margin = Inches(1)
    sec2.right_margin = Inches(1)

    add_line('OIKOS UNIVERSITY', 14, True, space_after=30)
    add_line('"A Study on the Mechanism by Which AI Utilization', 12, False, space_after=4)
    add_line('Training Requests Fail to Translate into', 12, False, space_after=4)
    add_line('Organizational Performance\u2013 Focusing', 12, False, space_after=4)
    add_line('on Seizing within Dynamic Capabilities Theory"', 12, False, space_after=0)
    add_line('', space_after=24)
    add_line('A DISSERTATION', 12, True, space_after=6)
    add_line('SUBMITTED TO THE FACULTY OF', 12, False, space_after=6)
    add_line('OIKOS UNIVERSITY', 12, True, space_after=6)
    add_line('IN CANDIDACY FOR THE DEGREE OF', 12, False, space_after=6)
    add_line('DOCTOR OF BUSINESS ADMINISTRATION', 12, True, space_after=0)
    add_line('', space_after=40)
    add_line('by', 12, False, space_after=4)
    add_line('Yijae Shin', 14, True, space_after=0)
    add_line('', space_after=60)
    add_line('OAKLAND, CALIFORNIA', 12, False, space_after=4)
    add_line('May 2026', 12, False, space_after=0)
    add_line('', space_after=40)
    add_line('Copyright 2026', 10, False, space_after=4)
    add_line('Yijae Shin', 10, False, space_after=4)
    add_line('ALL RIGHTS RESERVED', 10, False)

    # PAGE 3 (approval) is now extracted from original Korean PDF — not generated here

    doc.save(COVER_DOCX)
    print(f'  Cover DOCX: {len(doc.sections)} sections (2 pages, approval from original PDF)')


# =================================================================
# Fix #3: Move Keywords to same page as Abstract end
# =================================================================
def fix_keywords(doc):
    """Remove page break before Keywords so it stays on Abstract's last page."""
    from docx.oxml.ns import qn

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if text.startswith('Keywords:') and 'AI Utilization Training' in text:
            print(f'  Found Keywords at paragraph {i}: "{text[:60]}..."')
            removed = 0
            for j in range(i - 1, max(i - 5, 0), -1):
                prev = doc.paragraphs[j]
                prev_text = prev.text.strip()
                for run in prev.runs:
                    br_elements = run._element.findall(qn('w:br'))
                    for br in br_elements:
                        if br.get(qn('w:type')) == 'page':
                            run._element.remove(br)
                            removed += 1
                pPr = prev._element.find(qn('w:pPr'))
                if pPr is not None:
                    sectPr = pPr.find(qn('w:sectPr'))
                    if sectPr is not None:
                        pPr.remove(sectPr)
                        removed += 1
                if not prev_text and not prev._element.findall(f'.//{qn("w:drawing")}'):
                    prev._element.getparent().remove(prev._element)
                    removed += 1
                if prev_text:
                    break
            pPr = para._element.find(qn('w:pPr'))
            if pPr is not None:
                pb = pPr.find(qn('w:pageBreakBefore'))
                if pb is not None:
                    pPr.remove(pb)
                    removed += 1
            print(f'  Total removals: {removed}')
            return True
    print('  WARNING: Keywords paragraph not found!')
    return False


# =================================================================
# Fix #4: Replace figures with new improved versions
# =================================================================
def replace_figures(doc):
    """Replace figure images with minimal XML surgery to avoid corruption."""
    from docx.oxml.ns import qn
    from PIL import Image as PILImage

    ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    ns_wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    ns_wpg = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
    ns_wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'

    fig_rel_map = {
        'rId88': ('fig_1_1.png', 5.5),
        'rId75': ('fig_4_1.png', 5.5),
        'rId77': ('fig_4_2.png', 4.2),
        'rId89': ('fig_5_1.png', 5.0),
        'rId90': ('fig_5_2.png', 5.5),
        'rId91': ('fig_5_3.png', 5.5),
        'rId92': ('fig_6_1.png', 5.0),
    }

    replaced = 0
    for rel_id, (fig_file, width_in) in fig_rel_map.items():
        fig_path = os.path.join(FIG_DIR, fig_file)
        if not os.path.exists(fig_path) or rel_id not in doc.part.rels:
            print(f'    SKIP: {fig_file} (missing)')
            continue
        with open(fig_path, 'rb') as f:
            new_blob = f.read()
        doc.part.rels[rel_id].target_part._blob = new_blob
        replaced += 1
        print(f'    Blob: {fig_file} via {rel_id} ({len(new_blob):,}b)')

    # Fix Fig 4.1 (rId75): remove srcRect crop + remove overlay shapes + update dimensions
    print('  --- Fig 4.1 surgery ---')
    img_41 = PILImage.open(os.path.join(FIG_DIR, 'fig_4_1.png'))
    w41, h41 = img_41.size
    cx41 = int(5.5 * 914400)
    cy41 = int(cx41 * h41 / w41)

    for para in doc.paragraphs:
        blips = para._element.findall(f'.//{{{ns_a}}}blip')
        for blip in blips:
            if blip.get(f'{{{ns_r}}}embed') == 'rId75':
                drawing = para._element.findall(f'.//{qn("w:drawing")}')[0]
                for sr in drawing.findall(f'.//{{{ns_a}}}srcRect'):
                    sr.getparent().remove(sr)
                wgp = drawing.findall(f'.//{{{ns_wpg}}}wgp')
                if wgp:
                    for grpSp in wgp[0].findall(f'{{{ns_wpg}}}grpSp'):
                        wgp[0].remove(grpSp)
                    for wsp in wgp[0].findall(f'{{{ns_wps}}}wsp'):
                        wgp[0].remove(wsp)
                for ext in drawing.findall(f'.//{{{ns_wp}}}extent'):
                    ext.set('cx', str(cx41)); ext.set('cy', str(cy41))
                for grpSpPr in drawing.findall(f'.//{{{ns_wpg}}}grpSpPr'):
                    for xfrm in grpSpPr.findall(f'{{{ns_a}}}xfrm'):
                        for aext in xfrm.findall(f'{{{ns_a}}}ext'):
                            aext.set('cx', str(cx41)); aext.set('cy', str(cy41))
                        for chExt in xfrm.findall(f'{{{ns_a}}}chExt'):
                            chExt.set('cx', str(cx41)); chExt.set('cy', str(cy41))
                        for off in xfrm.findall(f'{{{ns_a}}}off'):
                            off.set('x', '0'); off.set('y', '0')
                        for chOff in xfrm.findall(f'{{{ns_a}}}chOff'):
                            chOff.set('x', '0'); chOff.set('y', '0')
                for pic_spPr in drawing.findall(f'.//{{{ns_a}}}spPr'):
                    for xfrm in pic_spPr.findall(f'{{{ns_a}}}xfrm'):
                        for aext in xfrm.findall(f'{{{ns_a}}}ext'):
                            aext.set('cx', str(cx41)); aext.set('cy', str(cy41))
                        for off in xfrm.findall(f'{{{ns_a}}}off'):
                            off.set('x', '0'); off.set('y', '0')
                print(f'    4.1 dims: {cx41/914400:.2f}in x {cy41/914400:.2f}in')
                break
        else:
            continue
        break

    # Fix Fig 4.2 (rId77)
    print('  --- Fig 4.2 surgery ---')
    img_42 = PILImage.open(os.path.join(FIG_DIR, 'fig_4_2.png'))
    w42, h42 = img_42.size
    cx42 = int(4.2 * 914400)
    cy42 = int(cx42 * h42 / w42)

    for para in doc.paragraphs:
        blips = para._element.findall(f'.//{{{ns_a}}}blip')
        for blip in blips:
            if blip.get(f'{{{ns_r}}}embed') == 'rId77':
                drawing = para._element.findall(f'.//{qn("w:drawing")}')[0]
                for sr in drawing.findall(f'.//{{{ns_a}}}srcRect'):
                    sr.getparent().remove(sr)
                for ext in drawing.findall(f'.//{{{ns_wp}}}extent'):
                    ext.set('cx', str(cx42)); ext.set('cy', str(cy42))
                for spPr in drawing.findall(f'.//{{{ns_a}}}spPr'):
                    for xfrm in spPr.findall(f'{{{ns_a}}}xfrm'):
                        for aext in xfrm.findall(f'{{{ns_a}}}ext'):
                            aext.set('cx', str(cx42)); aext.set('cy', str(cy42))
                print(f'    4.2 dims: {cx42/914400:.2f}in x {cy42/914400:.2f}in')
                break
        else:
            continue
        break

    # Fix Fig 5.3 (rId91): update dims
    print('  --- Fig 5.3 dims ---')
    img_53 = PILImage.open(os.path.join(FIG_DIR, 'fig_5_3.png'))
    w53, h53 = img_53.size
    cx53 = int(5.5 * 914400)
    cy53 = int(cx53 * h53 / w53)
    for para in doc.paragraphs:
        blips = para._element.findall(f'.//{{{ns_a}}}blip')
        for blip in blips:
            if blip.get(f'{{{ns_r}}}embed') == 'rId91':
                drawing = para._element.findall(f'.//{qn("w:drawing")}')[0]
                for ext in drawing.findall(f'.//{{{ns_wp}}}extent'):
                    ext.set('cx', str(cx53)); ext.set('cy', str(cy53))
                for spPr in drawing.findall(f'.//{{{ns_a}}}spPr'):
                    for xfrm in spPr.findall(f'{{{ns_a}}}xfrm'):
                        for aext in xfrm.findall(f'{{{ns_a}}}ext'):
                            aext.set('cx', str(cx53)); aext.set('cy', str(cy53))
                print(f'    5.3 dims: {cx53/914400:.2f}in x {cy53/914400:.2f}in')
                break
        else:
            continue
        break

    print(f'  Total blobs replaced: {replaced}')


# =================================================================
# Fix #5a: Remove duplicate "REFERENCES CITED"
# =================================================================
def fix_duplicate_references_cited(doc):
    from docx.oxml.ns import qn
    removed = 0
    for i, para in enumerate(doc.paragraphs):
        if i < 900:
            continue
        text = para.text.strip()
        style = para.style.name if para.style else ''
        if text.upper() == 'REFERENCES CITED' and 'EndNote' in style:
            para._element.getparent().remove(para._element)
            removed += 1
            print(f'  Removed duplicate REFERENCES CITED at para {i}')
    if removed == 0:
        print('  No duplicate REFERENCES CITED found')


# =================================================================
# Fix #5b: APPENDICES spacing
# =================================================================
def fix_appendices(doc):
    from docx.oxml.ns import qn
    from docx.shared import Pt
    from lxml import etree

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if i < 500:
            continue
        if not (text.lower() == 'appendices' or text.lower().startswith('appendices')):
            continue
        style_name = para.style.name if para.style else ''
        print(f'  Found APPENDICES at paragraph {i}: style="{style_name}"')

        pf = para.paragraph_format
        if pf.space_before and pf.space_before > Pt(36):
            pf.space_before = Pt(12)

        removed = 0
        for j in range(i - 1, max(i - 10, 0), -1):
            prev = doc.paragraphs[j]
            prev_text = prev.text.strip()
            if not prev_text and not prev._element.findall(f'.//{qn("w:drawing")}'):
                pPr_prev = prev._element.find(qn('w:pPr'))
                has_sectPr = pPr_prev is not None and pPr_prev.find(qn('w:sectPr')) is not None
                if has_sectPr:
                    continue
                prev._element.getparent().remove(prev._element)
                removed += 1
            elif prev_text:
                break

        pPr = para._element.find(qn('w:pPr'))
        if pPr is None:
            pPr = etree.SubElement(para._element, qn('w:pPr'))
            para._element.insert(0, pPr)
        pb = pPr.find(qn('w:pageBreakBefore'))
        if pb is None:
            etree.SubElement(pPr, qn('w:pageBreakBefore'))

        spacing = pPr.find(qn('w:spacing'))
        if spacing is not None:
            before_val = spacing.get(qn('w:before'))
            if before_val and int(before_val) > 720:
                spacing.set(qn('w:before'), '240')

        print(f'  APPENDICES fix applied (removed {removed} empty paras)')
        return True

    print('  WARNING: APPENDICES not found!')
    return False


# =================================================================
# Fix #7: Table 3.1 top border
# =================================================================
def fix_table_3_1_border(doc):
    """Add visible top border to Table 3.1 (Table index 1: 'Category')."""
    from docx.oxml.ns import qn
    from lxml import etree

    # Table 1 is Table 3.1 (preceded by "List and Volume of Data Sources")
    if len(doc.tables) < 2:
        print('  WARNING: Not enough tables!')
        return False

    tbl = doc.tables[1]
    print(f'  Table 3.1: {len(tbl.rows)} rows x {len(tbl.columns)} cols, first_cell="{tbl.rows[0].cells[0].text[:30]}"')

    # Ensure table-level top border is thick and visible
    tblPr = tbl._tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = etree.SubElement(tbl._tbl, qn('w:tblPr'))
        tbl._tbl.insert(0, tblPr)

    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is None:
        tblBorders = etree.SubElement(tblPr, qn('w:tblBorders'))

    # Set top border to thick single line
    top_el = tblBorders.find(qn('w:top'))
    if top_el is None:
        top_el = etree.SubElement(tblBorders, qn('w:top'))
    top_el.set(qn('w:val'), 'single')
    top_el.set(qn('w:sz'), '12')  # 1.5pt
    top_el.set(qn('w:space'), '0')
    top_el.set(qn('w:color'), '000000')

    # Also add explicit cell-level top borders on first row for certainty
    for cell in tbl.rows[0].cells:
        tcPr = cell._tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = etree.SubElement(cell._tc, qn('w:tcPr'))
            cell._tc.insert(0, tcPr)
        tcBorders = tcPr.find(qn('w:tcBorders'))
        if tcBorders is None:
            tcBorders = etree.SubElement(tcPr, qn('w:tcBorders'))
        top_cell = tcBorders.find(qn('w:top'))
        if top_cell is None:
            top_cell = etree.SubElement(tcBorders, qn('w:top'))
        top_cell.set(qn('w:val'), 'single')
        top_cell.set(qn('w:sz'), '12')
        top_cell.set(qn('w:space'), '0')
        top_cell.set(qn('w:color'), '000000')

    print('  Table 3.1: top border set to 1.5pt solid black')
    return True


# =================================================================
# Fix #8: Figure 5.1/5.2 caption position (title above image)
# =================================================================
def fix_figure_caption_positions(doc):
    """Move Figure 5.1 and 5.2 captions to be ABOVE their images (consistent with other figures).
    Also set keepWithNext on caption + subtitle to prevent page break before image."""
    from docx.oxml.ns import qn
    from lxml import etree

    def set_keep_with_next(element):
        pPr = element.find(qn('w:pPr'))
        if pPr is None:
            pPr = etree.SubElement(element, qn('w:pPr'))
            element.insert(0, pPr)
        kwn = pPr.find(qn('w:keepNext'))
        if kwn is None:
            kwn = etree.SubElement(pPr, qn('w:keepNext'))
        kwn.set(qn('w:val'), '1')

    def set_page_break_before(element):
        pPr = element.find(qn('w:pPr'))
        if pPr is None:
            pPr = etree.SubElement(element, qn('w:pPr'))
            element.insert(0, pPr)
        pbb = pPr.find(qn('w:pageBreakBefore'))
        if pbb is None:
            pbb = etree.SubElement(pPr, qn('w:pageBreakBefore'))
        pbb.set(qn('w:val'), '1')

    fixes = 0
    for fig_name in ['Figure 5.1', 'Figure 5.2']:
        # Find caption element directly
        caption_el = None
        image_el = None

        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if text.startswith(fig_name) and len(text) < 200:
                caption_el = para._element
                # Check if image is BEFORE caption (wrong order)
                if i > 0:
                    prev = doc.paragraphs[i - 1]
                    if len(prev._element.findall(f'.//{qn("w:drawing")}')) > 0:
                        image_el = prev._element
                break

        if caption_el is None:
            print(f'  WARNING: {fig_name} caption not found!')
            continue

        if image_el is not None:
            # Move caption before image
            parent = caption_el.getparent()
            parent.remove(caption_el)
            image_el.addprevious(caption_el)
            fixes += 1
            print(f'  {fig_name}: moved caption ABOVE image (was below)')

        # Ensure caption and image are on the same page:
        set_keep_with_next(caption_el)
        # Remove ALL page breaks from the image paragraph (inside runs)
        img_sib = caption_el.getnext()
        if img_sib is not None:
            # Remove pageBreakBefore property
            img_pPr = img_sib.find(qn('w:pPr'))
            if img_pPr is not None:
                for pbb in img_pPr.findall(qn('w:pageBreakBefore')):
                    img_pPr.remove(pbb)
            # Remove page break runs inside the image paragraph
            for run in img_sib.findall(f'.//{qn("w:r")}'):
                for br in run.findall(qn('w:br')):
                    if br.get(qn('w:type')) == 'page':
                        run.remove(br)
                        print(f'  {fig_name}: removed page break run from image para')
                # Clean up empty runs
                if len(run) == 0:
                    run.getparent().remove(run)
        print(f'  {fig_name}: caption/image sync done')

    print(f'  Caption position fixes: {fixes}')
    return fixes > 0


# =================================================================
# Fix #9: Section 5.7 spacing
# =================================================================
def fix_section_5_7_spacing(doc):
    """Reduce excessive whitespace above section 5.7 heading."""
    from docx.oxml.ns import qn
    from docx.shared import Pt
    from lxml import etree

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if '5.7.' in text and i > 500 and i < 900:
            style_name = para.style.name if para.style else ''
            if 'Heading' not in style_name:
                continue

            print(f'  Found 5.7 at P{i}: "{text[:60]}" style={style_name}')

            # Reduce space_before
            pPr = para._element.find(qn('w:pPr'))
            if pPr is not None:
                spacing = pPr.find(qn('w:spacing'))
                if spacing is not None:
                    before = spacing.get(qn('w:before'))
                    if before and int(before) > 480:
                        spacing.set(qn('w:before'), '240')
                        print(f'    Reduced spacing from {before} to 240 twips')

            # Remove empty paragraphs before
            removed = 0
            for j in range(i - 1, max(i - 5, 0), -1):
                prev = doc.paragraphs[j]
                prev_text = prev.text.strip()
                if not prev_text and not prev._element.findall(f'.//{qn("w:drawing")}'):
                    pPr_prev = prev._element.find(qn('w:pPr'))
                    has_sectPr = pPr_prev is not None and pPr_prev.find(qn('w:sectPr')) is not None
                    if has_sectPr:
                        continue
                    prev._element.getparent().remove(prev._element)
                    removed += 1
                elif prev_text:
                    break

            print(f'    Removed {removed} empty paragraphs before heading')
            return True

    print('  WARNING: Section 5.7 not found!')
    return False


# =================================================================
# Fix #11: Remove blank page after LIST OF TABLES
# =================================================================
def fix_lot_blank_page(doc):
    """Remove empty paragraphs between LIST OF TABLES entries and LIST OF FIGURES."""
    from docx.oxml.ns import qn

    lot_idx = None
    lof_idx = None
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        style = para.style.name if para.style else ''
        if 'LIST OF TABLES' in text and 'toc' not in style.lower():
            lot_idx = i
        if 'LIST OF FIGURES' in text and 'toc' not in style.lower():
            lof_idx = i
            break

    if lot_idx is None or lof_idx is None:
        print('  WARNING: LOT or LOF not found')
        return

    print(f'  LOT at P{lot_idx}, LOF at P{lof_idx}')

    # Remove empty paragraphs between LOT entries and LOF heading
    # Walk backwards from LOF to find empties
    removed = 0
    for j in range(lof_idx - 1, lot_idx, -1):
        para = doc.paragraphs[j]
        text = para.text.strip()
        has_drawing = len(para._element.findall(f'.//{qn("w:drawing")}')) > 0
        pPr = para._element.find(qn('w:pPr'))
        has_sectPr = pPr is not None and pPr.find(qn('w:sectPr')) is not None

        if not text and not has_drawing and not has_sectPr:
            para._element.getparent().remove(para._element)
            removed += 1
        elif has_sectPr and not text:
            # Keep section break paragraphs
            pass
        else:
            break  # Hit actual content

    print(f'  Removed {removed} empty paragraphs between LOT and LOF')

    # Add pageBreakBefore to LOF heading so it starts on a new page
    from lxml import etree
    lof_para = doc.paragraphs[lof_idx]
    lof_pPr = lof_para._element.find(qn('w:pPr'))
    if lof_pPr is None:
        lof_pPr = etree.SubElement(lof_para._element, qn('w:pPr'))
        lof_para._element.insert(0, lof_pPr)
    pb = lof_pPr.find(qn('w:pageBreakBefore'))
    if pb is None:
        etree.SubElement(lof_pPr, qn('w:pageBreakBefore'))
    print('  Added pageBreakBefore to LIST OF FIGURES')


# =================================================================
# Fix #10: Appendix D - two images on same page
# =================================================================
def fix_appendix_d_layout(doc):
    """Ensure both Appendix D images fit on the same page by reducing size and spacing."""
    from docx.oxml.ns import qn
    from docx.shared import Pt

    ns_wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    # Find Appendix D heading - store element reference, not index
    app_d_el = None
    appendix_d_idx = None
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if i > 900 and ('Appendix D' in text or 'APPENDIX D' in text):
            appendix_d_idx = i
            app_d_el = para._element
            print(f'  Found Appendix D at P{i}: "{text[:60]}"')
            break

    if app_d_el is None:
        print('  WARNING: Appendix D not found!')
        return False

    # Find the two image paragraphs after Appendix D (by XML sibling walk)
    img_paras = []
    sibling = app_d_el
    checked = 0
    while sibling is not None and checked < 10:
        drawings = sibling.findall(f'.//{qn("w:drawing")}')
        if drawings:
            # Find matching paragraph object
            for para in doc.paragraphs:
                if para._element is sibling:
                    img_paras.append((0, para, drawings))
                    break
        sibling = sibling.getnext()
        checked += 1

    print(f'  Found {len(img_paras)} images after Appendix D')

    if len(img_paras) < 2:
        print('  WARNING: Less than 2 images found!')
        return False

    # Remove empty paragraphs BEFORE Appendix D (cause of top whitespace)
    removed_before = 0
    prev = app_d_el.getprevious()
    for _ in range(20):
        if prev is None:
            break
        prev_text = prev.text or ''
        # Also check for run text
        all_text = ''.join(t.text or '' for t in prev.iter()).strip()
        has_drawing = len(prev.findall(f'.//{qn("w:drawing")}')) > 0
        if not all_text and not has_drawing:
            pPr_p = prev.find(qn('w:pPr'))
            has_sectPr = pPr_p is not None and pPr_p.find(qn('w:sectPr')) is not None
            if not has_sectPr:
                to_remove = prev
                prev = prev.getprevious()
                to_remove.getparent().remove(to_remove)
                removed_before += 1
                continue
            else:
                break
        else:
            break
        prev = prev.getprevious() if prev is not None else None
    print(f'    Removed {removed_before} empty paras before Appendix D')

    # Force Appendix D heading to start on a new page (pageBreakBefore)
    from lxml import etree
    app_d_para = None
    for para in doc.paragraphs:
        if para._element is app_d_el:
            app_d_para = para
            break
    pPr = app_d_para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = etree.SubElement(app_d_para._element, qn('w:pPr'))
        app_d_para._element.insert(0, pPr)
    pb_before = pPr.find(qn('w:pageBreakBefore'))
    if pb_before is None:
        pb_before = etree.SubElement(pPr, qn('w:pageBreakBefore'))
    print('    Added pageBreakBefore to Appendix D heading')

    # Set keepWithNext on heading, subtitle, and first image
    # so title + subtitle + images stay together on the same page
    el = app_d_el
    for _ in range(4):
        if el is None:
            break
        pPr_el = el.find(qn('w:pPr'))
        if pPr_el is None:
            pPr_el = etree.SubElement(el, qn('w:pPr'))
            el.insert(0, pPr_el)
        kwn = pPr_el.find(qn('w:keepNext'))
        if kwn is None:
            etree.SubElement(pPr_el, qn('w:keepNext'))
        el = el.getnext()

    # Resize images to fill page width (5.8in, close to original)
    target_width_in = 5.8
    for idx, (p_idx, para, drawings) in enumerate(img_paras):
        for d in drawings:
            for ext in d.findall(f'.//{{{ns_wp}}}extent'):
                old_cx = int(ext.get('cx', 0))
                old_cy = int(ext.get('cy', 0))
                new_cx = int(target_width_in * 914400)
                ratio = new_cx / old_cx if old_cx else 1
                new_cy = int(old_cy * ratio)
                ext.set('cx', str(new_cx))
                ext.set('cy', str(new_cy))
                print(f'    Image {idx+1}: {old_cx/914400:.2f}in -> {new_cx/914400:.2f}in')

            for spPr in d.findall(f'.//{{{ns_a}}}spPr'):
                for xfrm in spPr.findall(f'{{{ns_a}}}xfrm'):
                    for aext in xfrm.findall(f'{{{ns_a}}}ext'):
                        old_cx = int(aext.get('cx', 0))
                        old_cy = int(aext.get('cy', 0))
                        new_cx = int(target_width_in * 914400)
                        ratio = new_cx / old_cx if old_cx else 1
                        new_cy = int(old_cy * ratio)
                        aext.set('cx', str(new_cx))
                        aext.set('cy', str(new_cy))

        # Tight spacing around image paragraphs
        pf = para.paragraph_format
        pf.space_before = Pt(2)
        pf.space_after = Pt(2)

    # Tight spacing on Appendix D heading and subtitle
    if app_d_para:
        app_d_para.paragraph_format.space_before = Pt(6)
        app_d_para.paragraph_format.space_after = Pt(3)
    # Subtitle (next sibling)
    subtitle_el = app_d_el.getnext()
    if subtitle_el is not None:
        for para in doc.paragraphs:
            if para._element is subtitle_el:
                para.paragraph_format.space_before = Pt(3)
                para.paragraph_format.space_after = Pt(3)
                break

    # Remove empty paragraphs between title and images (walk siblings)
    el = app_d_el.getnext()
    for _ in range(10):
        if el is None:
            break
        next_el = el.getnext()
        all_text = ''.join(t.text or '' for t in el.iter()).strip()
        has_drawing = len(el.findall(f'.//{qn("w:drawing")}')) > 0
        if not all_text and not has_drawing:
            pPr_p = el.find(qn('w:pPr'))
            has_sectPr = pPr_p is not None and pPr_p.find(qn('w:sectPr')) is not None
            if not has_sectPr:
                el.getparent().remove(el)
                print(f'    Removed empty para between title and images')
        el = next_el

    print('  Appendix D: images resized and spacing reduced')
    return True


# =================================================================
# Export PDF via Word COM
# =================================================================
def export_pdf_via_word(docx_path, pdf_path, update_fields=False, save_copy=None):
    import win32com.client

    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        doc = word.Documents.Open(docx_path)
        time.sleep(2)
        pages = doc.ComputeStatistics(2)
        print(f'  Opened: {pages} pages')

        if update_fields:
            doc.Range().Fields.Update()
            if doc.TablesOfContents.Count > 0:
                for i in range(1, doc.TablesOfContents.Count + 1):
                    doc.TablesOfContents(i).Update()
                print(f'  Updated {doc.TablesOfContents.Count} TOC(s)')

            for i in range(1, doc.Sections.Count + 1):
                try:
                    for j in range(1, 4):
                        try: doc.Sections(i).Headers(j).Range.Fields.Update()
                        except: pass
                        try: doc.Sections(i).Footers(j).Range.Fields.Update()
                        except: pass
                except: pass

            time.sleep(1)

        if save_copy:
            doc.SaveAs2(save_copy)
            print(f'  Saved copy: {os.path.basename(save_copy)}')

        pages_final = doc.ComputeStatistics(2)
        print(f'  Final: {pages_final} pages')

        doc.ExportAsFixedFormat(pdf_path, 17, False, 0, 0)
        print(f'  PDF: {os.path.basename(pdf_path)}')

        doc.Close(0)
    except Exception as e:
        print(f'  ERROR: {e}')
        import traceback
        traceback.print_exc()
        try: doc.Close(0)
        except: pass
    finally:
        word.Quit()


# =================================================================
# Merge PDFs
# =================================================================
def merge_pdfs(cover_pdf, body_pdf, output_pdf):
    """Merge: 2-page cover + approval page from original Korean PDF + body."""
    from PyPDF2 import PdfMerger, PdfReader, PdfWriter

    # Extract approval page (page 3, index 2) from original Korean PDF
    approval_pdf = os.path.join(os.path.dirname(output_pdf), '_approval_page.pdf')
    reader = PdfReader(ORIGINAL_PDF)
    writer = PdfWriter()
    writer.add_page(reader.pages[2])  # page 3 (0-indexed)
    with open(approval_pdf, 'wb') as f:
        writer.write(f)
    print(f'  Extracted approval page from original PDF')

    merger = PdfMerger()
    merger.append(cover_pdf)       # pages 1-2 (cover + inner cover)
    merger.append(approval_pdf)    # page 3 (original approval sheet)
    merger.append(body_pdf)        # pages 4+ (body)
    merger.write(output_pdf)
    merger.close()

    os.remove(approval_pdf)
    print(f'  Merged: {os.path.basename(output_pdf)} ({os.path.getsize(output_pdf):,} bytes)')


# =================================================================
# Verification
# =================================================================
def verify_final(pdf_path):
    import fitz
    import re

    doc = fitz.open(pdf_path)
    total = len(doc)
    print(f'  Total pages: {total}')

    offset = None
    for i in range(total):
        txt = doc[i].get_text().strip()
        lines = txt.split('\n')
        if lines and lines[0].strip() == '1' and 'CHAPTER 1' in txt:
            offset = i
            break

    if offset is None:
        print('  WARNING: Could not find body page 1!')
        doc.close()
        return

    print(f'  Body starts at PDF page {offset + 1} (offset={offset})')

    # Check first pages
    print('\n  --- First pages ---')
    for i in range(min(8, total)):
        txt = doc[i].get_text()[:100].strip().replace('\n', ' | ')
        print(f'    PDF pg {i+1}: {txt[:80]}')

    # Keywords check
    print('\n  --- Keywords check ---')
    for i in range(total):
        txt = doc[i].get_text()
        if 'Keywords:' in txt and 'AI Utilization' in txt:
            has_abstract_text = len(txt) > 300
            status = 'GOOD (with Abstract)' if has_abstract_text else 'BAD (separate page)'
            print(f'    Keywords on PDF page {i+1}: {status}')
            break

    # APPENDICES check
    print('\n  --- APPENDICES check ---')
    for i in range(total):
        txt = doc[i].get_text().strip()
        if txt.startswith('APPENDICES') or '\nAPPENDICES\n' in txt:
            print(f'    APPENDICES on PDF page {i+1}')
            break

    # Figure verification
    print('\n  --- Figure verification ---')
    figures = {
        'Figure 1.1': 33, 'Figure 4.1': 107, 'Figure 4.2': 113,
        'Figure 5.1': 127, 'Figure 5.2': 136, 'Figure 5.3': 140,
        'Figure 6.1': 150,
    }
    for name, listed in figures.items():
        found = False
        for delta in range(-3, 4):
            check = listed + offset - 1 + delta
            if 0 <= check < total:
                txt = doc[check].get_text()
                has_img = len(doc[check].get_images()) > 0
                if name in txt and has_img:
                    bp = check - offset + 1
                    status = 'OK' if bp == listed else f'SHIFTED (listed:{listed}, actual:{bp})'
                    print(f'    {name}: body page {bp} [IMAGE] -> {status}')
                    found = True
                    break
        if not found:
            # Try wider search
            for check in range(max(0, listed + offset - 5), min(total, listed + offset + 5)):
                txt = doc[check].get_text()
                if name in txt:
                    bp = check - offset + 1
                    has_img = len(doc[check].get_images()) > 0
                    img_str = '[IMAGE]' if has_img else '[NO IMG]'
                    print(f'    {name}: body page {bp} {img_str} -> FOUND')
                    found = True
                    break
            if not found:
                print(f'    {name}: NOT FOUND!')

    # Table verification
    print('\n  --- Table verification ---')
    tables = {'Table 3.1': 69, 'Table 3.2': 73, 'Table 3.3': 75}
    for name, listed in tables.items():
        idx = listed + offset - 1
        if 0 <= idx < total:
            txt = doc[idx].get_text()
            if name in txt:
                print(f'    {name}: body page {listed} -> OK')
            else:
                print(f'    {name}: NOT on page {listed}!')

    # TOC spot-check
    print('\n  --- TOC page number spot-check ---')
    for i in range(total):
        txt = doc[i].get_text()
        if 'TABLE OF CONTENTS' in txt or 'ABLE OF CONTENTS' in txt:
            for line in txt.split('\n'):
                line = line.strip()
                m = re.search(r'(CHAPTER \d+|REFERENCES|APPENDICES).*?(\d+)\s*$', line)
                if m:
                    section = m.group(1)
                    toc_page = int(m.group(2))
                    check_idx = toc_page + offset - 1
                    if 0 <= check_idx < total:
                        check_txt = doc[check_idx].get_text()[:200]
                        found = section in check_txt or section.split()[1] in check_txt
                        status = 'OK' if found else 'CHECK'
                        print(f'    TOC: {section} -> page {toc_page} [{status}]')
            break

    # Appendix D images check
    print('\n  --- Appendix D images check ---')
    for i in range(total - 5, total):
        txt = doc[i].get_text()
        imgs = doc[i].get_images()
        if 'Appendix D' in txt or 'appendix' in txt.lower():
            print(f'    PDF pg {i+1}: {len(imgs)} images, text preview: "{txt[:60].strip()}"')
        elif len(imgs) > 0 and i > total - 5:
            print(f'    PDF pg {i+1}: {len(imgs)} images')

    doc.close()


# =================================================================
# Main
# =================================================================
def main():
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')

    # Backup
    backup = os.path.join(BACKUP_DIR, f'thesis_pre_v4build_{ts}.docx')
    shutil.copy2(MAIN_DOCX, backup)
    print(f'Backup: {os.path.basename(backup)}')

    # Step 1: Create cover with bigger signatures
    print('\n=== Step 1: Create cover (bigger signatures) ===')
    create_cover()

    # Step 2: Fix thesis DOCX
    print('\n=== Step 2: Fix thesis DOCX ===')
    from docx import Document
    doc = Document(MAIN_DOCX)

    print('\n--- Fix #3: Keywords position ---')
    fix_keywords(doc)

    print('\n--- Fix #4: Replace figures ---')
    replace_figures(doc)

    print('\n--- Fix #5a: Duplicate REFERENCES CITED ---')
    fix_duplicate_references_cited(doc)

    print('\n--- Fix #5b: APPENDICES spacing ---')
    fix_appendices(doc)

    print('\n--- Fix #7: Table 3.1 top border ---')
    fix_table_3_1_border(doc)

    print('\n--- Fix #8: Figure 5.1/5.2 caption positions ---')
    fix_figure_caption_positions(doc)

    print('\n--- Fix #9: Section 5.7 spacing ---')
    fix_section_5_7_spacing(doc)

    print('\n--- Fix #11: Remove blank page after LIST OF TABLES ---')
    fix_lot_blank_page(doc)

    print('\n--- Fix #10: Appendix D layout ---')
    fix_appendix_d_layout(doc)

    doc.save(MAIN_DOCX)
    print(f'\nSaved: {os.path.basename(MAIN_DOCX)}')

    # Step 3: Export cover PDF
    print('\n=== Step 3: Export cover PDF ===')
    export_pdf_via_word(COVER_DOCX, COVER_PDF)

    # Step 4: Export body PDF (with field update + save copy)
    print('\n=== Step 4: Export body PDF ===')
    export_pdf_via_word(MAIN_DOCX, BODY_PDF, update_fields=True, save_copy=FINAL_DOCX)

    # Step 5: Merge
    print('\n=== Step 5: Merge PDFs ===')
    merge_pdfs(COVER_PDF, BODY_PDF, FINAL_PDF)

    # Step 6: Verify
    print('\n=== Step 6: Verification ===')
    verify_final(FINAL_PDF)

    print('\n=== Done ===')
    for f in [FINAL_PDF, FINAL_DOCX]:
        if os.path.exists(f):
            print(f'  {os.path.basename(f)}: {os.path.getsize(f):,} bytes')


if __name__ == '__main__':
    main()
