"""Build final thesis PDF: body PDF + cover PDF -> merged final.
Strategy: never modify the main thesis structurally. Export PDFs separately, then merge."""
import sys
import io
import os
import time
import shutil

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
MAIN_DOCX = os.path.join(BASE, 'output', 'thesis_english.docx')
COVER_DOCX = os.path.join(BASE, 'output', 'cover_pages.docx')
BODY_PDF = os.path.join(BASE, 'output', 'thesis_body.pdf')
COVER_PDF = os.path.join(BASE, 'output', 'thesis_cover.pdf')
FINAL_PDF = os.path.join(BASE, 'output', 'thesis_english.pdf')
FINAL_DOCX = os.path.join(BASE, 'output', 'thesis_english_final.docx')
SIG_DIR = os.path.join(BASE, '_temp_figs')


def create_cover_docx():
    """Create a standalone 3-page cover document using python-docx proper API."""
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_ORIENT

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    # === PAGE 1: COVER ===
    for _ in range(7):
        doc.add_paragraph('')

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('\u300cA Study on the Mechanism by Which AI Utilization Training Requests\n'
                     'Fail to Translate into Organizational Performance\n'
                     '\u2013 Focusing on Seizing within Dynamic Capabilities Theory\u300d')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.bold = True

    for _ in range(4):
        doc.add_paragraph('')

    cover_lines = [
        ('By', 12, False),
        ('Yijae Shin', 14, True),
        ('', 12, False),
        ('A Dissertation Presented to the Faculty of', 12, False),
        ('OIKOS UNIVERSITY', 14, True),
        ('In Partial Fulfillment of the', 12, False),
        ('Requirements for the Degree of Doctor of Business Administration', 12, False),
        ('', 12, False),
        ('', 12, False),
        ('February 2026', 12, False),
    ]
    for text, size, bold in cover_lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(size)
            run.bold = bold

    doc.add_page_break()

    # === PAGE 2: INNER COVER ===
    for _ in range(4):
        doc.add_paragraph('')

    inner_lines = [
        ('OIKOS UNIVERSITY', 16, True),
        ('', 12, False),
        ('"A Study on the Mechanism by Which AI Utilization Training Requests', 12, False),
        ('Fail to Translate into Organizational Performance\u2013 Focusing', 12, False),
        ('on Seizing within Dynamic Capabilities Theory"', 12, False),
        ('', 12, False),
        ('A DISSERTATION', 12, True),
        ('SUBMITTED TO THE FACULTY OF', 12, False),
        ('OIKOS UNIVERSITY', 12, True),
        ('IN CANDIDACY FOR THE DEGREE OF', 12, False),
        ('DOCTOR OF BUSINESS ADMINISTRATION', 12, True),
        ('', 12, False),
        ('by', 12, False),
        ('Yijae Shin', 14, True),
        ('', 12, False),
        ('OAKLAND, CALIFORNIA', 12, False),
        ('May 2026', 12, False),
        ('', 12, False),
        ('Copyright 2026', 10, False),
        ('Yijae Shin', 10, False),
        ('ALL RIGHTS RESERVED', 10, False),
    ]
    for text, size, bold in inner_lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(size)
            run.bold = bold

    doc.add_page_break()

    # === PAGE 3: APPROVAL SHEET ===
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('DISSERTATION APPROVAL SHEET')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.bold = True

    doc.add_paragraph('')

    approval_lines = [
        'This dissertation, entitled',
        '"A Study on the Mechanism by Which AI Utilization Training Requests',
        'Fail to Translate into Organizational Performance\u2013 Focusing',
        'on Seizing within Dynamic Capabilities Theory"',
        '',
        'And submitted in candidacy for the degree of',
        'Doctor of Business Administration',
        'Has been read and approved',
        'by the undersigned members of the faculty of',
        'Oikos University',
    ]
    for text in approval_lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)

    doc.add_paragraph('')

    # Signatures
    sig_files = [
        ('signature_0_608x98.png', 'Chairman'),
        ('signature_1_718x140.png', 'Member'),
        ('signature_2_930x110.png', 'Member'),
        ('signature_3_638x88.png', 'Member'),
        ('signature_4_730x140.png', 'Member'),
    ]
    for sig_name, title in sig_files:
        sig_path = os.path.join(SIG_DIR, sig_name)
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_img.add_run()
        run.add_picture(sig_path, width=Inches(2.5))

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_title.add_run(title)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)

    doc.save(COVER_DOCX)
    print(f'Cover DOCX created: {os.path.basename(COVER_DOCX)}')


def export_pdf_via_word(docx_path, pdf_path, update_fields=False, cleanup_abstract=False):
    """Open a DOCX in Word COM and export to PDF."""
    import win32com.client

    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        doc = word.Documents.Open(docx_path)
        time.sleep(2)
        pages = doc.ComputeStatistics(2)
        print(f'  Opened: {pages} pages')

        if cleanup_abstract:
            # Delete empty second abstract section
            find = doc.Range().Find
            find.ClearFormatting()
            find.Text = 'TABLE OF CONTENTS'
            found1 = find.Execute()

            find2 = doc.Range().Find
            find2.ClearFormatting()
            find2.Text = 'Keywords: AI Utilization Training'
            found2 = find2.Execute()

            if found1 and found2:
                toc_pos = find.Parent.Start
                kw_end = find2.Parent.Paragraphs(1).Range.End
                cleanup = doc.Range(kw_end, toc_pos)
                if len(cleanup.Text.strip()) < 50:
                    cleanup.Delete()
                    print('  Cleaned empty abstract section')

        if update_fields:
            doc.Range().Fields.Update()
            if doc.TablesOfContents.Count > 0:
                for i in range(1, doc.TablesOfContents.Count + 1):
                    doc.TablesOfContents(i).Update()
                print(f'  Updated {doc.TablesOfContents.Count} TOC(s)')

            for i in range(1, doc.Sections.Count + 1):
                try:
                    for j in range(1, 4):
                        try:
                            doc.Sections(i).Headers(j).Range.Fields.Update()
                        except:
                            pass
                        try:
                            doc.Sections(i).Footers(j).Range.Fields.Update()
                        except:
                            pass
                except:
                    pass

            time.sleep(1)

        # Also save an updated DOCX copy (without structural changes)
        if update_fields:
            doc.SaveAs2(FINAL_DOCX)
            print(f'  Saved updated DOCX: {os.path.basename(FINAL_DOCX)}')

        pages_final = doc.ComputeStatistics(2)
        print(f'  Final: {pages_final} pages')

        doc.ExportAsFixedFormat(pdf_path, 17, False, 0, 0)
        print(f'  PDF exported: {os.path.basename(pdf_path)}')

        doc.Close(0)
    except Exception as e:
        print(f'  ERROR: {e}')
        import traceback
        traceback.print_exc()
        try:
            doc.Close(0)
        except:
            pass
    finally:
        word.Quit()


def merge_pdfs(cover_pdf, body_pdf, output_pdf):
    """Merge cover and body PDFs."""
    from PyPDF2 import PdfMerger

    merger = PdfMerger()
    merger.append(cover_pdf)
    merger.append(body_pdf)
    merger.write(output_pdf)
    merger.close()
    print(f'Merged PDF: {os.path.basename(output_pdf)} ({os.path.getsize(output_pdf):,} bytes)')


def main():
    # Step 1: Create cover document
    print('=== Step 1: Create cover document ===')
    create_cover_docx()

    # Step 2: Export cover to PDF
    print('\n=== Step 2: Export cover PDF ===')
    export_pdf_via_word(COVER_DOCX, COVER_PDF)

    # Step 3: Export body to PDF (with field updates)
    print('\n=== Step 3: Export body PDF ===')
    export_pdf_via_word(MAIN_DOCX, BODY_PDF, update_fields=True, cleanup_abstract=True)

    # Step 4: Merge PDFs
    print('\n=== Step 4: Merge PDFs ===')
    merge_pdfs(COVER_PDF, BODY_PDF, FINAL_PDF)

    # Verify
    print('\n=== Verification ===')
    import fitz
    doc = fitz.open(FINAL_PDF)
    print(f'Total pages: {len(doc)}')
    for i in range(min(10, len(doc))):
        txt = doc[i].get_text()[:80].strip().replace('\n', ' ')
        print(f'  Page {i+1}: {txt}')
    doc.close()

    print('\n=== Output Files ===')
    for f in [FINAL_PDF, FINAL_DOCX]:
        if os.path.exists(f):
            print(f'  {os.path.basename(f)}: {os.path.getsize(f):,} bytes')


if __name__ == '__main__':
    main()
