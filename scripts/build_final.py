"""Build final thesis: clean abstract, insert covers via Word COM, update fields, export PDF."""
import sys
import io
import os
import time
import win32com.client

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
MAIN_DOCX = os.path.join(BASE, 'output', 'thesis_english.docx')
FINAL_DOCX = os.path.join(BASE, 'output', 'thesis_english_final.docx')
PDF_PATH = os.path.join(BASE, 'output', 'thesis_english.pdf')
SIG_DIR = os.path.join(BASE, '_temp_figs')


def main():
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        # Open main document
        print('Opening main document...')
        doc = word.Documents.Open(MAIN_DOCX)
        time.sleep(2)
        print(f'  Pages: {doc.ComputeStatistics(2)}')

        # === 1. Delete empty second abstract section ===
        print('Cleaning empty abstract section...')
        find = doc.Range().Find
        find.ClearFormatting()
        find.Text = 'TABLE OF CONTENTS'
        find.Execute()
        toc_pos = find.Parent.Start

        find2 = doc.Range().Find
        find2.ClearFormatting()
        find2.Text = 'Keywords: AI Utilization Training'
        find2.Execute()
        kw_end = find2.Parent.Paragraphs(1).Range.End

        cleanup = doc.Range(kw_end, toc_pos)
        if len(cleanup.Text.strip()) < 50:
            cleanup.Delete()
            print('  Deleted empty section')

        time.sleep(1)
        print(f'  Pages after cleanup: {doc.ComputeStatistics(2)}')

        # === 2. Insert cover pages ===
        print('Inserting cover pages...')
        sel = word.Selection
        sel.HomeKey(6)  # wdStory

        # --- PAGE 1: COVER ---
        sel.ParagraphFormat.Alignment = 1  # Center
        sel.Font.Name = 'Times New Roman'
        sel.Font.Size = 12
        sel.Font.Bold = False

        for _ in range(8):
            sel.TypeParagraph()

        sel.Font.Size = 14
        sel.Font.Bold = True
        sel.TypeText('\u300cA Study on the Mechanism by Which AI Utilization Training Requests')
        sel.TypeParagraph()
        sel.TypeText('Fail to Translate into Organizational Performance')
        sel.TypeParagraph()
        sel.TypeText('\u2013 Focusing on Seizing within Dynamic Capabilities Theory\u300d')
        sel.TypeParagraph()

        sel.Font.Bold = False
        sel.Font.Size = 12
        for _ in range(4):
            sel.TypeParagraph()

        sel.TypeText('By')
        sel.TypeParagraph()
        sel.Font.Size = 14
        sel.Font.Bold = True
        sel.TypeText('Yijae Shin')
        sel.TypeParagraph()
        sel.Font.Bold = False
        sel.Font.Size = 12
        sel.TypeParagraph()
        sel.TypeText('A Dissertation Presented to the Faculty of')
        sel.TypeParagraph()
        sel.Font.Bold = True
        sel.TypeText('OIKOS UNIVERSITY')
        sel.TypeParagraph()
        sel.Font.Bold = False
        sel.TypeText('In Partial Fulfillment of the')
        sel.TypeParagraph()
        sel.TypeText('Requirements for the Degree of Doctor of Business Administration')
        sel.TypeParagraph()
        sel.TypeParagraph()
        sel.TypeText('February 2026')

        sel.InsertBreak(2)  # wdSectionBreakNextPage

        # --- PAGE 2: INNER COVER ---
        sel.ParagraphFormat.Alignment = 1
        sel.Font.Name = 'Times New Roman'
        sel.Font.Size = 12
        sel.Font.Bold = False

        for _ in range(5):
            sel.TypeParagraph()

        sel.Font.Size = 16
        sel.Font.Bold = True
        sel.TypeText('OIKOS UNIVERSITY')
        sel.TypeParagraph()
        sel.Font.Size = 12
        sel.Font.Bold = False
        sel.TypeParagraph()
        sel.TypeText('"A Study on the Mechanism by Which AI Utilization Training Requests')
        sel.TypeParagraph()
        sel.TypeText('Fail to Translate into Organizational Performance\u2013 Focusing')
        sel.TypeParagraph()
        sel.TypeText('on Seizing within Dynamic Capabilities Theory"')
        sel.TypeParagraph()
        sel.TypeParagraph()
        sel.Font.Bold = True
        sel.TypeText('A DISSERTATION')
        sel.TypeParagraph()
        sel.Font.Bold = False
        sel.TypeText('SUBMITTED TO THE FACULTY OF')
        sel.TypeParagraph()
        sel.Font.Bold = True
        sel.TypeText('OIKOS UNIVERSITY')
        sel.TypeParagraph()
        sel.Font.Bold = False
        sel.TypeText('IN CANDIDACY FOR THE DEGREE OF')
        sel.TypeParagraph()
        sel.Font.Bold = True
        sel.TypeText('DOCTOR OF BUSINESS ADMINISTRATION')
        sel.TypeParagraph()
        sel.Font.Bold = False
        sel.TypeParagraph()
        sel.TypeText('by')
        sel.TypeParagraph()
        sel.Font.Size = 14
        sel.Font.Bold = True
        sel.TypeText('Yijae Shin')
        sel.TypeParagraph()
        sel.Font.Bold = False
        sel.Font.Size = 12
        sel.TypeParagraph()
        sel.TypeText('OAKLAND, CALIFORNIA')
        sel.TypeParagraph()
        sel.TypeText('May 2026')
        sel.TypeParagraph()
        sel.TypeParagraph()
        sel.Font.Size = 10
        sel.TypeText('Copyright 2026')
        sel.TypeParagraph()
        sel.TypeText('Yijae Shin')
        sel.TypeParagraph()
        sel.TypeText('ALL RIGHTS RESERVED')

        sel.InsertBreak(2)  # wdSectionBreakNextPage

        # --- PAGE 3: APPROVAL SHEET ---
        sel.ParagraphFormat.Alignment = 1
        sel.Font.Name = 'Times New Roman'
        sel.Font.Size = 14
        sel.Font.Bold = True
        sel.TypeText('DISSERTATION APPROVAL SHEET')
        sel.TypeParagraph()
        sel.Font.Bold = False
        sel.Font.Size = 12
        sel.TypeParagraph()
        sel.TypeText('This dissertation, entitled')
        sel.TypeParagraph()
        sel.TypeText('"A Study on the Mechanism by Which AI Utilization Training Requests')
        sel.TypeParagraph()
        sel.TypeText('Fail to Translate into Organizational Performance\u2013 Focusing')
        sel.TypeParagraph()
        sel.TypeText('on Seizing within Dynamic Capabilities Theory"')
        sel.TypeParagraph()
        sel.TypeParagraph()
        sel.TypeText('And submitted in candidacy for the degree of')
        sel.TypeParagraph()
        sel.TypeText('Doctor of Business Administration')
        sel.TypeParagraph()
        sel.TypeText('Has been read and approved')
        sel.TypeParagraph()
        sel.TypeText('by the undersigned members of the faculty of')
        sel.TypeParagraph()
        sel.TypeText('Oikos University')
        sel.TypeParagraph()
        sel.TypeParagraph()

        # Insert 5 signatures
        sig_files = [
            ('signature_0_608x98.png', 'Chairman'),
            ('signature_1_718x140.png', 'Member'),
            ('signature_2_930x110.png', 'Member'),
            ('signature_3_638x88.png', 'Member'),
            ('signature_4_730x140.png', 'Member'),
        ]

        for sig_name, title in sig_files:
            sig_path = os.path.join(SIG_DIR, sig_name)
            shape = sel.InlineShapes.AddPicture(sig_path)
            # Scale to ~2.5 inches width
            orig_w = shape.Width
            shape.Width = 180  # points (~2.5 inches)
            if orig_w > 0:
                shape.Height = int(shape.Height * 180 / orig_w)
            sel.MoveRight(1, 1)  # Move past image
            sel.TypeParagraph()
            sel.TypeText(title)
            sel.TypeParagraph()

        # Section break before main content
        sel.InsertBreak(2)  # wdSectionBreakNextPage

        time.sleep(2)
        print(f'  Cover inserted. Pages: {doc.ComputeStatistics(2)}')

        # === 3. Fix page numbering for cover sections ===
        print('Fixing page numbering...')
        total_sections = doc.Sections.Count
        print(f'  Total sections: {total_sections}')

        # Cover sections (1-3): no page numbers, unlinked
        for i in range(1, 5):  # sections 1-4 (cover + transition)
            try:
                sec = doc.Sections(i)
                for j in range(1, 4):
                    try:
                        sec.Headers(j).LinkToPrevious = False
                    except:
                        pass
                    try:
                        sec.Footers(j).LinkToPrevious = False
                    except:
                        pass
                # Clear footers in cover sections
                if i <= 3:
                    try:
                        sec.Footers(1).Range.Text = ''
                        sec.Footers(2).Range.Text = ''
                    except:
                        pass
            except Exception as e:
                print(f'  S{i} warning: {e}')

        # Section after covers (front matter): restart roman numerals
        try:
            front_sec = doc.Sections(4)
            front_sec.Footers(1).LinkToPrevious = False
            if front_sec.Footers(1).PageNumbers.Count == 0:
                front_sec.Footers(1).PageNumbers.Add(4, True)  # wdAlignPageNumberCenter
            front_sec.Footers(1).PageNumbers.RestartNumberingAtSection = True
            front_sec.Footers(1).PageNumbers.StartingNumber = 1
            front_sec.Footers(1).PageNumbers.NumberStyle = 2  # wdPageNumberStyleLowercaseRoman
            print('  S4: roman numeral page numbers set')
        except Exception as e:
            print(f'  S4 page num error: {e}')

        print('  Page numbering configured')

        # === 4. Update all fields ===
        print('Updating all fields...')
        doc.Range().Fields.Update()

        if doc.TablesOfContents.Count > 0:
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(i).Update()
            print(f'  Updated {doc.TablesOfContents.Count} TOC(s)')

        # Update header/footer fields
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

        time.sleep(2)
        final_pages = doc.ComputeStatistics(2)
        print(f'  Final pages: {final_pages}')

        # === 5. Save and export ===
        doc.SaveAs2(FINAL_DOCX)
        print(f'Saved: {os.path.basename(FINAL_DOCX)}')

        doc.ExportAsFixedFormat(PDF_PATH, 17, False, 0, 0)
        print(f'PDF: {os.path.basename(PDF_PATH)}')

        doc.Close(0)

    except Exception as e:
        print(f'ERROR: {e}')
        import traceback
        traceback.print_exc()
        try:
            doc.Close(0)
        except:
            pass
    finally:
        word.Quit()
        print('Word closed')

    # Verify
    for f in [FINAL_DOCX, PDF_PATH]:
        if os.path.exists(f):
            print(f'{os.path.basename(f)}: {os.path.getsize(f):,} bytes')
        else:
            print(f'{os.path.basename(f)}: NOT CREATED')


if __name__ == '__main__':
    main()
