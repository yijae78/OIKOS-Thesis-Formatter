"""Create cover pages matching original PDF layout coordinates exactly."""
import sys
import io
import os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.shared import Pt, Inches, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
SIG_DIR = os.path.join(BASE, '_temp_figs')
OUTPUT = os.path.join(BASE, 'output', 'cover_pages.docx')

# Original PDF coordinates (points): page 612x792
# 1 point = 1/72 inch. Margin = 1 inch = 72 points.
# So content area starts at y=72 (top margin)
# To place text at y=145, need ~73pt of space after margin


def add_line(doc, text, size=12, bold=False, space_before=0, space_after=0):
    """Add a centered paragraph with precise spacing."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = Pt(size + 2)  # tight line spacing
    if text:
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size)
        run.bold = bold
    return p


def page1_cover(doc):
    """Page 1: Cover page. Original y-coords: title@145, By@294, affil@410, date@675."""
    sec = doc.sections[0]
    sec.top_margin = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1)
    sec.right_margin = Inches(1)

    # Space from top margin (72pt) to title start (145pt) = 73pt
    add_line(doc, '', space_after=60)

    # Title block (y=145-264 in original, ~120pt range)
    add_line(doc, '\u300cA Study on the Mechanism by Which AI Utilization', 14, True, space_after=2)
    add_line(doc, 'Training Requests Fail to Translate into', 14, True, space_after=2)
    add_line(doc, 'Organizational Performance', 14, True, space_after=2)
    add_line(doc, '\u2013 Focusing on Seizing within Dynamic', 14, True, space_after=2)
    add_line(doc, 'Capabilities Theory\u300d', 14, True, space_after=0)

    # Gap to "By" (y=294, about 30pt below title end ~264)
    add_line(doc, '', space_after=20)

    add_line(doc, 'By', 12, False, space_after=4)
    add_line(doc, 'Yijae Shin', 12, False, space_after=0)

    # Gap to affiliation (y=410, about 80pt below author ~331)
    add_line(doc, '', space_after=60)

    add_line(doc, 'A Dissertation Presented to the Faculty of', 12, False, space_after=6)
    add_line(doc, 'OIKOS UNIVERSITY', 14, True, space_after=6)
    add_line(doc, 'In Partial Fulfillment of the', 12, False, space_after=6)
    add_line(doc, 'Requirements for the Degree of', 12, False, space_after=2)
    add_line(doc, 'Doctor of Business Administration', 12, False, space_after=0)

    # Gap to date (y=675, about 180pt below affil end ~496)
    add_line(doc, '', space_after=170)

    add_line(doc, 'February 2026', 12, False)


def page2_inner_cover(doc):
    """Page 2: Inner cover. Original: OIKOS@80, title@148, DISSERTATION@251."""
    sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
    sec.top_margin = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1)
    sec.right_margin = Inches(1)

    # OIKOS at y=80 (very close to top margin 72pt)
    add_line(doc, 'OIKOS UNIVERSITY', 14, True, space_after=30)

    # Title block at y=148
    add_line(doc, '"A Study on the Mechanism by Which AI Utilization', 12, False, space_after=4)
    add_line(doc, 'Training Requests Fail to Translate into', 12, False, space_after=4)
    add_line(doc, 'Organizational Performance\u2013 Focusing', 12, False, space_after=4)
    add_line(doc, 'on Seizing within Dynamic Capabilities Theory"', 12, False, space_after=0)

    # Gap to DISSERTATION block (y=251)
    add_line(doc, '', space_after=24)

    add_line(doc, 'A DISSERTATION', 12, True, space_after=6)
    add_line(doc, 'SUBMITTED TO THE FACULTY OF', 12, False, space_after=6)
    add_line(doc, 'OIKOS UNIVERSITY', 12, True, space_after=6)
    add_line(doc, 'IN CANDIDACY FOR THE DEGREE OF', 12, False, space_after=6)
    add_line(doc, 'DOCTOR OF BUSINESS ADMINISTRATION', 12, True, space_after=0)

    # Gap to "by" (y=419)
    add_line(doc, '', space_after=40)

    add_line(doc, 'by', 12, False, space_after=4)
    add_line(doc, 'Yijae Shin', 14, True, space_after=0)

    # Gap to location (y=539)
    add_line(doc, '', space_after=60)

    add_line(doc, 'OAKLAND, CALIFORNIA', 12, False, space_after=4)
    add_line(doc, 'May 2026', 12, False, space_after=0)

    # Gap to copyright (y=635)
    add_line(doc, '', space_after=40)

    add_line(doc, 'Copyright 2026', 10, False, space_after=4)
    add_line(doc, 'Yijae Shin', 10, False, space_after=4)
    add_line(doc, 'ALL RIGHTS RESERVED', 10, False)


def page3_approval(doc):
    """Page 3: Approval sheet. Original: title@95, dissertation text@129, sigs@435+."""
    sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
    sec.top_margin = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1)
    sec.right_margin = Inches(1)

    # Title at y=95 (close to top margin)
    add_line(doc, 'DISSERTATION APPROVAL SHEET', 14, True, space_after=12)

    # Dissertation text at y=129
    add_line(doc, 'This dissertation, entitled', 12, False, space_after=8)

    # Title in quotes
    add_line(doc, '"A Study on the Mechanism by Which AI Utilization', 12, False, space_after=2)
    add_line(doc, 'Training Requests Fail to Translate into', 12, False, space_after=2)
    add_line(doc, 'Organizational Performance\u2013 Focusing', 12, False, space_after=2)
    add_line(doc, 'on Seizing within Dynamic Capabilities Theory"', 12, False, space_after=0)

    add_line(doc, '', space_after=8)

    add_line(doc, 'And submitted in candidacy for the degree of', 12, False, space_after=4)
    add_line(doc, 'Doctor of Business Administration', 12, False, space_after=4)
    add_line(doc, 'Has been read and approved', 12, False, space_after=4)
    add_line(doc, 'by the undersigned members of the faculty of', 12, False, space_after=4)
    add_line(doc, 'Oikos University', 12, False, space_after=0)

    add_line(doc, '', space_after=14)

    # Signatures (5 total, need to fit on one page)
    # Original: Chairman@435, Member@482, Member@530, Member@578, Member@620
    # Spacing between sigs: ~48pt each
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
        pf.space_before = Pt(4)
        pf.space_after = Pt(0)
        run = p_img.add_run()
        run.add_picture(sig_path, width=Inches(1.5))

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf2 = p_title.paragraph_format
        pf2.space_before = Pt(0)
        pf2.space_after = Pt(6)
        run2 = p_title.add_run(title)
        run2.font.name = 'Times New Roman'
        run2.font.size = Pt(11)

    # May 2026 at bottom
    add_line(doc, '', space_after=8)
    add_line(doc, 'May 2026', 12, False)


def main():
    doc = Document()

    # Set default style
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    page1_cover(doc)
    page2_inner_cover(doc)
    page3_approval(doc)

    doc.save(OUTPUT)
    print(f'Cover document created: {len(doc.sections)} sections')


if __name__ == '__main__':
    main()
