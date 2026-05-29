"""Build final thesis PDF v3: All fixes applied in one pass.
Fixes: #1 Cover title balance, #2 Blank page + approval overflow,
       #3 Keywords position, #4 Figure replacement, #5 APPENDICES spacing.
Then exports PDF and verifies all page numbers."""
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
FINAL_PDF = os.path.join(BASE, 'output', 'thesis_english.pdf')
FINAL_DOCX = os.path.join(BASE, 'output', 'thesis_english_final.docx')
FIG_DIR = os.path.join(BASE, 'output', 'figures')
SIG_DIR = os.path.join(BASE, '_temp_figs')
BACKUP_DIR = os.path.join(BASE, 'backups')


# =================================================================
# Fix #1 + #2: Create cover with balanced title and compact approval
# =================================================================
def create_cover():
    """Create 3-page cover: balanced title, inner cover, compact approval."""
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

    # === PAGE 1: COVER (balanced - title moved up) ===
    sec = doc.sections[0]
    sec.top_margin = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1)
    sec.right_margin = Inches(1)

    # Reduced top space (was 60pt -> 30pt to move title up)
    add_line('', space_after=30)

    # Title block
    add_line('\u300cA Study on the Mechanism by Which AI Utilization', 14, True, space_after=2)
    add_line('Training Requests Fail to Translate into', 14, True, space_after=2)
    add_line('Organizational Performance', 14, True, space_after=2)
    add_line('\u2013 Focusing on Seizing within Dynamic', 14, True, space_after=2)
    add_line('Capabilities Theory\u300d', 14, True, space_after=0)

    # Gap to "By"
    add_line('', space_after=30)

    add_line('By', 12, False, space_after=4)
    add_line('Yijae Shin', 12, False, space_after=0)

    # Gap to affiliation
    add_line('', space_after=70)

    add_line('A Dissertation Presented to the Faculty of', 12, False, space_after=6)
    add_line('OIKOS UNIVERSITY', 14, True, space_after=6)
    add_line('In Partial Fulfillment of the', 12, False, space_after=6)
    add_line('Requirements for the Degree of', 12, False, space_after=2)
    add_line('Doctor of Business Administration', 12, False, space_after=0)

    # Gap to date (pushed down more for balance)
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

    # === PAGE 3: APPROVAL SHEET (compact - all 5 sigs on one page) ===
    sec3 = doc.add_section(WD_SECTION_START.NEW_PAGE)
    sec3.top_margin = Inches(0.8)
    sec3.bottom_margin = Inches(0.6)
    sec3.left_margin = Inches(1)
    sec3.right_margin = Inches(1)

    add_line('DISSERTATION APPROVAL SHEET', 14, True, space_after=8)
    add_line('This dissertation, entitled', 11, False, space_after=4)
    add_line('"A Study on the Mechanism by Which AI Utilization', 11, False, space_after=2)
    add_line('Training Requests Fail to Translate into', 11, False, space_after=2)
    add_line('Organizational Performance\u2013 Focusing', 11, False, space_after=2)
    add_line('on Seizing within Dynamic Capabilities Theory"', 11, False, space_after=0)
    add_line('', space_after=4)
    add_line('And submitted in candidacy for the degree of', 11, False, space_after=2)
    add_line('Doctor of Business Administration', 11, False, space_after=2)
    add_line('Has been read and approved', 11, False, space_after=2)
    add_line('by the undersigned members of the faculty of', 11, False, space_after=2)
    add_line('Oikos University', 11, False, space_after=0)
    add_line('', space_after=6)

    # 5 signatures - compact spacing
    sigs = [
        ('signature_0_608x98.png', 'Chairman'),
        ('signature_1_718x140.png', 'Member'),
        ('signature_2_930x110.png', 'Member'),
        ('signature_3_638x88.png', 'Member'),
        ('signature_4_730x140.png', 'Member'),
    ]
    for sig_name, title in sigs:
        sig_path = os.path.join(SIG_DIR, sig_name)
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p_img.paragraph_format
        pf.space_before = Pt(2)
        pf.space_after = Pt(0)
        run = p_img.add_run()
        run.add_picture(sig_path, width=Inches(1.3))

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf2 = p_title.paragraph_format
        pf2.space_before = Pt(0)
        pf2.space_after = Pt(2)
        run2 = p_title.add_run(title)
        run2.font.name = 'Times New Roman'
        run2.font.size = Pt(10)

    doc.save(COVER_DOCX)
    print(f'  Cover DOCX: {len(doc.sections)} sections')


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

            # Check preceding paragraphs for page breaks or empty paras
            # Remove empty paragraphs and page breaks before Keywords
            removed = 0
            for j in range(i - 1, max(i - 5, 0), -1):
                prev = doc.paragraphs[j]
                prev_text = prev.text.strip()

                # Remove page break runs
                for run in prev.runs:
                    br_elements = run._element.findall(qn('w:br'))
                    for br in br_elements:
                        if br.get(qn('w:type')) == 'page':
                            run._element.remove(br)
                            removed += 1
                            print(f'    Removed page break from paragraph {j}')

                # Remove section breaks that cause new pages
                pPr = prev._element.find(qn('w:pPr'))
                if pPr is not None:
                    sectPr = pPr.find(qn('w:sectPr'))
                    if sectPr is not None:
                        pPr.remove(sectPr)
                        removed += 1
                        print(f'    Removed section break from paragraph {j}')

                # Remove empty paragraphs before Keywords
                if not prev_text and not prev._element.findall(f'.//{qn("w:drawing")}'):
                    prev._element.getparent().remove(prev._element)
                    removed += 1
                    print(f'    Removed empty paragraph {j}')

                if prev_text:
                    break  # Stop at first non-empty paragraph

            # Also check if Keywords paragraph itself has a page break before
            pPr = para._element.find(qn('w:pPr'))
            if pPr is not None:
                pb = pPr.find(qn('w:pageBreakBefore'))
                if pb is not None:
                    pPr.remove(pb)
                    removed += 1
                    print(f'    Removed pageBreakBefore from Keywords paragraph')

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
    from lxml import etree

    ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    ns_wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    ns_wpg = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
    ns_wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'

    # ALL figures: blob replacement first
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
    print('  --- Fig 4.1 surgery (anchor+group) ---')
    img_41 = PILImage.open(os.path.join(FIG_DIR, 'fig_4_1.png'))
    w41, h41 = img_41.size  # 1355x1355 = square
    cx41 = int(5.5 * 914400)  # 5.5 inches
    cy41 = int(cx41 * h41 / w41)  # square = same

    for para in doc.paragraphs:
        blips = para._element.findall(f'.//{{{ns_a}}}blip')
        for blip in blips:
            if blip.get(f'{{{ns_r}}}embed') == 'rId75':
                drawing = para._element.findall(f'.//{qn("w:drawing")}')[0]

                # 1. Remove srcRect (crop)
                for sr in drawing.findall(f'.//{{{ns_a}}}srcRect'):
                    sr.getparent().remove(sr)
                    print('    Removed srcRect crop')

                # 2. Remove overlay group shapes (wpg:grpSp) inside the main group
                wgp = drawing.findall(f'.//{{{ns_wpg}}}wgp')
                if wgp:
                    for grpSp in wgp[0].findall(f'{{{ns_wpg}}}grpSp'):
                        wgp[0].remove(grpSp)
                        print('    Removed overlay grpSp')
                    for wsp in wgp[0].findall(f'{{{ns_wps}}}wsp'):
                        wgp[0].remove(wsp)
                        print('    Removed overlay wsp')

                # 3. Update ALL extent/dimension elements
                # wp:extent on anchor
                for ext in drawing.findall(f'.//{{{ns_wp}}}extent'):
                    ext.set('cx', str(cx41))
                    ext.set('cy', str(cy41))
                # a:ext in grpSpPr (group transform)
                for grpSpPr in drawing.findall(f'.//{{{ns_wpg}}}grpSpPr'):
                    for xfrm in grpSpPr.findall(f'{{{ns_a}}}xfrm'):
                        for aext in xfrm.findall(f'{{{ns_a}}}ext'):
                            aext.set('cx', str(cx41))
                            aext.set('cy', str(cy41))
                        for chExt in xfrm.findall(f'{{{ns_a}}}chExt'):
                            chExt.set('cx', str(cx41))
                            chExt.set('cy', str(cy41))
                        for off in xfrm.findall(f'{{{ns_a}}}off'):
                            off.set('x', '0')
                            off.set('y', '0')
                        for chOff in xfrm.findall(f'{{{ns_a}}}chOff'):
                            chOff.set('x', '0')
                            chOff.set('y', '0')
                # a:ext in pic spPr
                for pic_spPr in drawing.findall(f'.//{{{ns_a}}}spPr'):
                    for xfrm in pic_spPr.findall(f'{{{ns_a}}}xfrm'):
                        for aext in xfrm.findall(f'{{{ns_a}}}ext'):
                            aext.set('cx', str(cx41))
                            aext.set('cy', str(cy41))
                        for off in xfrm.findall(f'{{{ns_a}}}off'):
                            off.set('x', '0')
                            off.set('y', '0')

                print(f'    Updated dims: {cx41/914400:.2f}in x {cy41/914400:.2f}in')
                break
        else:
            continue
        break

    # Fix Fig 4.2 (rId77): remove srcRect crop + update dimensions
    print('  --- Fig 4.2 surgery (anchor) ---')
    img_42 = PILImage.open(os.path.join(FIG_DIR, 'fig_4_2.png'))
    w42, h42 = img_42.size  # 985x1705
    cx42 = int(4.2 * 914400)
    cy42 = int(cx42 * h42 / w42)

    for para in doc.paragraphs:
        blips = para._element.findall(f'.//{{{ns_a}}}blip')
        for blip in blips:
            if blip.get(f'{{{ns_r}}}embed') == 'rId77':
                drawing = para._element.findall(f'.//{qn("w:drawing")}')[0]

                # Remove srcRect
                for sr in drawing.findall(f'.//{{{ns_a}}}srcRect'):
                    sr.getparent().remove(sr)
                    print('    Removed srcRect crop')

                # Update extent
                for ext in drawing.findall(f'.//{{{ns_wp}}}extent'):
                    ext.set('cx', str(cx42))
                    ext.set('cy', str(cy42))
                for spPr in drawing.findall(f'.//{{{ns_a}}}spPr'):
                    for xfrm in spPr.findall(f'{{{ns_a}}}xfrm'):
                        for aext in xfrm.findall(f'{{{ns_a}}}ext'):
                            aext.set('cx', str(cx42))
                            aext.set('cy', str(cy42))

                print(f'    Updated dims: {cx42/914400:.2f}in x {cy42/914400:.2f}in')
                break
        else:
            continue
        break

    # Fig 5.3 (rId91): inline, dims already match ratio. Just verify.
    print('  --- Fig 5.3 check (inline) ---')
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
                    ext.set('cx', str(cx53))
                    ext.set('cy', str(cy53))
                for spPr in drawing.findall(f'.//{{{ns_a}}}spPr'):
                    for xfrm in spPr.findall(f'{{{ns_a}}}xfrm'):
                        for aext in xfrm.findall(f'{{{ns_a}}}ext'):
                            aext.set('cx', str(cx53))
                            aext.set('cy', str(cy53))
                print(f'    Updated dims: {cx53/914400:.2f}in x {cy53/914400:.2f}in')
                break
        else:
            continue
        break

    print(f'  Total blobs replaced: {replaced}')


# =================================================================
# Fix #5a: Remove duplicate "REFERENCES CITED" (EndNote artifact)
# =================================================================
def fix_duplicate_references_cited(doc):
    """Remove duplicate 'REFERENCES CITED' paragraph near Appendices (EndNote artifact)."""
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
            print(f'  Removed duplicate "REFERENCES CITED" at para {i} (style: {style})')
    if removed == 0:
        print('  No duplicate REFERENCES CITED found')
    return removed > 0


# =================================================================
# Fix #5b: Fix APPENDICES page spacing
# =================================================================
def fix_appendices(doc):
    """Remove excessive spacing before APPENDICES body heading (skip TOC entries)."""
    from docx.oxml.ns import qn
    from docx.shared import Pt

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        # Skip TOC entries (they appear early, before para 500)
        # The actual APPENDICES body heading is around para 1059
        if i < 500:
            continue
        if not (text.lower() == 'appendices' or text.lower().startswith('appendices')):
            continue

        style_name = para.style.name if para.style else ''
        print(f'  Found APPENDICES at paragraph {i}: style="{style_name}", text="{text[:40]}"')

        # Reduce space_before on the heading
        pf = para.paragraph_format
        if pf.space_before and pf.space_before > Pt(36):
            old = pf.space_before
            pf.space_before = Pt(12)
            print(f'    Reduced space_before from {old} to {Pt(12)}')

        # Remove preceding empty paragraphs BUT preserve section breaks
        removed = 0
        for j in range(i - 1, max(i - 10, 0), -1):
            prev = doc.paragraphs[j]
            prev_text = prev.text.strip()
            if not prev_text and not prev._element.findall(f'.//{qn("w:drawing")}'):
                # Check if this paragraph carries a section break - preserve it
                pPr_prev = prev._element.find(qn('w:pPr'))
                has_sectPr = pPr_prev is not None and pPr_prev.find(qn('w:sectPr')) is not None
                if has_sectPr:
                    print(f'    Keeping paragraph {j} (has section break)')
                    continue
                prev._element.getparent().remove(prev._element)
                removed += 1
                print(f'    Removed empty paragraph {j}')
            elif prev_text:
                break

        # Ensure APPENDICES starts on a new page with pageBreakBefore
        from lxml import etree
        pPr = para._element.find(qn('w:pPr'))
        if pPr is None:
            pPr = etree.SubElement(para._element, qn('w:pPr'))
            para._element.insert(0, pPr)
        pb = pPr.find(qn('w:pageBreakBefore'))
        if pb is None:
            pb = etree.SubElement(pPr, qn('w:pageBreakBefore'))
            print('    Added pageBreakBefore')

        # Fix XML spacing
        spacing = pPr.find(qn('w:spacing'))
        if spacing is not None:
            before_val = spacing.get(qn('w:before'))
            if before_val and int(before_val) > 720:
                spacing.set(qn('w:before'), '240')
                print(f'    Reduced XML spacing from {before_val} to 240 twips')

        print(f'  APPENDICES fix applied (removed {removed} empty paras)')
        return True

    print('  WARNING: APPENDICES body heading not found (checked para 500+)!')
    return False


# =================================================================
# Export PDF via Word COM
# =================================================================
def export_pdf_via_word(docx_path, pdf_path, update_fields=False, save_copy=None):
    """Open DOCX in Word COM, optionally update fields, export PDF."""
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
    from PyPDF2 import PdfMerger
    merger = PdfMerger()
    merger.append(cover_pdf)
    merger.append(body_pdf)
    merger.write(output_pdf)
    merger.close()
    print(f'  Merged: {os.path.basename(output_pdf)} ({os.path.getsize(output_pdf):,} bytes)')


# =================================================================
# Verification
# =================================================================
def verify_final(pdf_path):
    """Verify all page numbers, figures, and tables match."""
    import fitz
    import re

    doc = fitz.open(pdf_path)
    total = len(doc)
    print(f'  Total pages: {total}')

    # Find body offset
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

    # Check first few pages
    print('\n  --- First pages ---')
    for i in range(min(8, total)):
        txt = doc[i].get_text()[:100].strip().replace('\n', ' | ')
        print(f'    PDF pg {i+1}: {txt[:80]}')

    # Check Keywords position
    print('\n  --- Keywords check ---')
    for i in range(total):
        txt = doc[i].get_text()
        if 'Keywords:' in txt and 'AI Utilization' in txt:
            # Check if Abstract text is also on this page
            has_abstract_text = len(txt) > 300  # Keywords-only page would be short
            status = 'GOOD (with Abstract text)' if has_abstract_text else 'BAD (separate page)'
            print(f'    Keywords on PDF page {i+1}: {status}')
            break

    # Check APPENDICES position
    print('\n  --- APPENDICES check ---')
    for i in range(total):
        txt = doc[i].get_text().strip()
        if txt.startswith('APPENDICES') or '\nAPPENDICES\n' in txt:
            # Check if content starts near top of page
            lines = txt.split('\n')
            first_content = lines[0].strip() if lines else ''
            print(f'    APPENDICES on PDF page {i+1}: starts with "{first_content[:50]}"')
            break

    # Verify figures
    print('\n  --- Figure verification ---')
    figures = {
        'Figure 1.1': 33, 'Figure 4.1': 107, 'Figure 4.2': 113,
        'Figure 5.1': 127, 'Figure 5.2': 136, 'Figure 5.3': 140,
        'Figure 6.1': 150,
    }
    for name, listed in figures.items():
        idx = listed + offset - 1
        found = False
        for delta in range(-2, 3):
            check = idx + delta
            if 0 <= check < total:
                txt = doc[check].get_text()
                has_img = len(doc[check].get_images()) > 0
                if name.replace('Figure ', 'Figure ') in txt and has_img:
                    bp = check - offset + 1
                    status = 'OK' if bp == listed else f'MISMATCH (listed:{listed}, actual:{bp})'
                    print(f'    {name}: body page {bp} [HAS IMAGE] -> {status}')
                    found = True
                    break
        if not found:
            print(f'    {name}: NOT FOUND near expected page {listed}!')

    # Verify tables
    print('\n  --- Table verification ---')
    tables = {'Table 3.1': 69, 'Table 3.2': 73, 'Table 3.3': 75}
    for name, listed in tables.items():
        idx = listed + offset - 1
        if 0 <= idx < total:
            txt = doc[idx].get_text()
            if name in txt:
                print(f'    {name}: body page {listed} -> OK')
            else:
                print(f'    {name}: NOT on expected page {listed}!')

    # Verify TOC entries
    print('\n  --- TOC page number spot-check ---')
    for i in range(total):
        txt = doc[i].get_text()
        if 'TABLE OF CONTENTS' in txt or 'ABLE OF CONTENTS' in txt:
            # Extract some TOC entries and verify
            for line in txt.split('\n'):
                line = line.strip()
                # Match patterns like "CHAPTER 1  INTRODUCTION ...1"
                m = re.search(r'(CHAPTER \d+|REFERENCES|APPENDICES).*?(\d+)\s*$', line)
                if m:
                    section = m.group(1)
                    toc_page = int(m.group(2))
                    # Verify
                    check_idx = toc_page + offset - 1
                    if 0 <= check_idx < total:
                        check_txt = doc[check_idx].get_text()[:200]
                        found = section in check_txt or section.split()[1] in check_txt
                        status = 'OK' if found else 'CHECK'
                        print(f'    TOC: {section} -> page {toc_page} [{status}]')
            break

    doc.close()


# =================================================================
# Main
# =================================================================
def main():
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')

    # Backup
    backup = os.path.join(BACKUP_DIR, f'thesis_pre_v3build_{ts}.docx')
    shutil.copy2(MAIN_DOCX, backup)
    print(f'Backup: {os.path.basename(backup)}')

    # Step 1: Fix cover
    print('\n=== Step 1: Create fixed cover ===')
    create_cover()

    # Step 2: Fix thesis DOCX
    print('\n=== Step 2: Fix thesis DOCX ===')
    from docx import Document
    doc = Document(MAIN_DOCX)

    print('\n--- Fix #3: Keywords position ---')
    fix_keywords(doc)

    print('\n--- Fix #4: Replace figures ---')
    replace_figures(doc)

    print('\n--- Fix #5a: Remove duplicate REFERENCES CITED ---')
    fix_duplicate_references_cited(doc)

    print('\n--- Fix #5b: APPENDICES spacing ---')
    fix_appendices(doc)

    doc.save(MAIN_DOCX)
    print(f'\nSaved: {os.path.basename(MAIN_DOCX)}')

    # Step 3: Export cover PDF
    print('\n=== Step 3: Export cover PDF ===')
    export_pdf_via_word(COVER_DOCX, COVER_PDF)

    # Step 4: Export body PDF
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
