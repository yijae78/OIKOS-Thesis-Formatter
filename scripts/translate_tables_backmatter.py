"""Translate Korean text in tables (0,2,3,4,5) and back matter headings
in thesis_english.docx. Table 1 (abbreviations) is already English."""
import os
import re
import shutil
from datetime import datetime
from docx import Document
from docx.oxml.ns import qn


def clear_cell_set_text(cell, new_text, font_name='Times New Roman'):
    """Clear all content in a cell and set new text, preserving pPr."""
    for paragraph in cell.paragraphs:
        elem = paragraph._element
        # Remove all children except paragraph properties
        for child in list(elem):
            if child.tag != qn('w:pPr'):
                elem.remove(child)
    # Add text to the first paragraph only
    first_para = cell.paragraphs[0]
    if new_text:
        run = first_para.add_run(new_text)
        run.font.name = font_name


def clear_paragraph_content(para, new_text, font_name='Times New Roman'):
    """Clear ALL content in a paragraph and set new text."""
    elem = para._element
    for child in list(elem):
        if child.tag != qn('w:pPr'):
            elem.remove(child)
    run = para.add_run(new_text)
    run.font.name = font_name


def translate_table_0(doc):
    """Title page table: remove Korean from cell[0,0], remove Korean name from cell[1,0]."""
    tbl = doc.tables[0]
    changed = 0

    # Cell [0,0]: Has English title (paras 0) + Korean title (para 2).
    # Keep para 0 (English), clear para 2 (Korean) and para 3.
    cell_00 = tbl.rows[0].cells[0]
    for pi, para in enumerate(cell_00.paragraphs):
        if pi >= 2:
            # Remove all content from Korean paragraphs
            elem = para._element
            for child in list(elem):
                if child.tag != qn('w:pPr'):
                    elem.remove(child)
            changed += 1

    # Cell [1,0]: "By" / "Yijae Shin" / "신이재" - remove last paragraph content
    cell_10 = tbl.rows[1].cells[0]
    for pi, para in enumerate(cell_10.paragraphs):
        if pi == 2:  # The Korean name paragraph
            elem = para._element
            for child in list(elem):
                if child.tag != qn('w:pPr'):
                    elem.remove(child)
            changed += 1

    print(f"  Table 0: {changed} changes (removed Korean title + name)")
    return changed


def translate_table_2(doc):
    """Data list table (5x6) - Table 3.1."""
    tbl = doc.tables[2]
    translations = {
        # Row 0 headers
        (0, 0): "Category",
        (0, 1): "Data Type",
        (0, 2): "Count",
        (0, 3): "Date of Creation/\nExtraction",
        (0, 4): "Inclusion Criteria",
        (0, 5): "Exclusion Criteria",
        # Row 1
        (1, 0): "Appendix A",
        (1, 1): "Official Correspondence for AI Utilization Training Request",
        (1, 2): "1 item",
        # (1, 3) keep date
        (1, 4): "Organizational-level AI utilization training request document submitted and received through official approval procedures",
        (1, 5): "Excludes unofficial proposals and verbal requests",
        # Row 2
        (2, 0): "Appendix B",
        (2, 1): "External Consulting Report (Summary)",
        (2, 2): "1 item",
        # (2, 3) keep date
        (2, 4): "External report that officially diagnosed the organization\u2019s digital/AI utilization status and training needs",
        (2, 5): "Excludes unofficial proposals and verbal requests",
        # Row 3
        (3, 0): "Appendix C",
        (3, 1): "External Consulting Report (Final Version Excerpt)",
        (3, 2): "1 item",
        # (3, 3) keep date
        (3, 4): "Portions directly relevant to research analysis from the official final deliverable expanding upon the summary",
        (3, 5): "Excludes unofficial revisions and unfinalized drafts",
        # Row 4
        (4, 0): "Appendix D",
        (4, 1): "Electronic Approval System Search Records",
        (4, 2): "2 items",
        # (4, 3) keep dates
        (4, 4): "Search result screens retrieved using identical search terms from the official records management and electronic approval system",
        (4, 5): "Excludes data not registered in the system and unofficial discussion records",
    }

    changed = 0
    for (r, c), text in translations.items():
        clear_cell_set_text(tbl.rows[r].cells[c], text)
        changed += 1

    print(f"  Table 2: {changed} cells translated")
    return changed


def translate_table_3(doc):
    """Analysis stages table (6x4) - Table 3.2."""
    tbl = doc.tables[3]
    translations = {
        # Row 0 headers
        (0, 0): "Analysis Stage",
        (0, 1): "Analysis Content",
        (0, 2): "Data Sources",
        (0, 3): "Corresponding Appendix",
        # Row 1
        (1, 0): "Judgment Formation Stage",
        (1, 1): "Verification of whether organizational-level formal judgment was formed regarding the AI utilization training request",
        (1, 2): "Internal organizational correspondence and approval documents requesting digital/AI capability enhancement training",
        (1, 3): "Appendix A",
        # Row 2
        (2, 0): "Institutional Context Diagnosis Stage",
        (2, 1): "Analysis of the organization\u2019s digital/AI awareness level and whether institutional conditions for converting judgment into decision-making exist",
        (2, 2): "External consulting reports on AI utilization (summary and final version excerpt/summary)",
        (2, 3): "Appendices B, C",
        # Row 3
        (3, 0): "Conversion Attempt and Stagnation Stage",
        (3, 1): "Verification of whether judgment was converted into a formal decision-making agenda item and whether conversion indicators such as schedule confirmation and assignment of responsible parties were observed",
        (3, 2): "Reconstructed table of \u2018AI training\u2019 search results from the organizational records management system",
        (3, 3): "Appendix D",
        # Row 4
        (4, 0): "Non-decision State Analysis Stage",
        (4, 1): "Analysis of whether the absence of formal decision constitutes a sustained and repeated non-decision state rather than a temporary omission",
        (4, 2): "Official request documents from Appendix A and reconstructed search result data from Appendix D",
        (4, 3): "Appendices A, D",
        # Row 5
        (5, 0): "Comprehensive Interpretation and Verification Stage",
        (5, 1): "Integrated interpretation of the judgment\u2013decision-making conversion failure mechanism through cross-verification among multiple data sources",
        (5, 2): "Comprehensive comparison of official documents, system records, and consulting reports",
        (5, 3): "Appendices A\u2013D",
    }

    changed = 0
    for (r, c), text in translations.items():
        clear_cell_set_text(tbl.rows[r].cells[c], text)
        changed += 1

    print(f"  Table 3: {changed} cells translated")
    return changed


def translate_table_4(doc):
    """Operational indicators table (7x4) - Table 3.3."""
    tbl = doc.tables[4]
    translations = {
        # Row 0 headers
        (0, 0): "Analytical Concept",
        (0, 1): "Operational Indicator",
        (0, 2): "Source Data",
        (0, 3): "Interpretive Memo",
        # Row 1
        (1, 0): "Organizational-Level Judgment Formation",
        (1, 1): "Whether the AI utilization training request was drafted as an official document and approved through the authorization chain",
        (1, 2): "AI utilization training request correspondence, approval authorization language",
        (1, 3): "The matter was incorporated into the organization\u2019s formal review procedure rather than remaining an individual proposal, demonstrating that the judgment was arranged into a state of official organizational consideration through the institutional channel of authorization approval.",
        # Row 2
        (2, 0): "Institutional Conditions for Conversion",
        (2, 1): "Whether a responsible department or accountable party was explicitly designated following the judgment",
        (2, 2): "External consulting reports, internal diagnostic materials",
        (2, 3): "It was confirmed that although the organization recognized the necessity of AI utilization, the allocation of responsibility and implementation structures for moving the judgment into procedural flow remained institutionally uninvoked.",
        # Row 3
        (3, 0): "Placement on Decision-Making Agenda",
        (3, 1): "Whether the matter was placed as an agenda item before a formal decision-making body",
        (3, 2): "Personnel committee schedule records, electronic approval system screens",
        (3, 3): "The judgment existed, yet a state in which it failed to enter the organizational arrangement of formal decision-making procedures was maintained.",
        # Row 4
        (4, 0): "Absence of Conversion Indicators",
        (4, 1): "Whether implementation indicators such as schedule confirmation, assignment of responsible parties, or initiation of follow-up procedures were present",
        (4, 2): "Electronic approval system records, absence of follow-up documents",
        (4, 3): "Institutional signals indicating implementation were not observed, revealing not a mere delay but a state in which procedural initiation conditions remained unmet.",
        # Row 5
        (5, 0): "Persistence of Non-decision State",
        (5, 1): "Whether the state of decision absence persisted and recurred over time",
        (5, 2): "Extended absence of system records following official correspondence, observation records",
        (5, 3): "The absence of decision-making is interpreted not as a temporary vacuum but as a conditional state in which the judgment was maintained without being translated into the organization\u2019s formal decision structure.",
        # Row 6
        (6, 0): "Comprehensive Interpretation",
        (6, 1): "Consistency of interpretation and potential for cross-verification among multiple data sources",
        (6, 2): "Appendices A\u2013D in their entirety",
        (6, 3): "Disparate data sources point to the identical judgment\u2013stagnation arrangement, demonstrating that this constitutes a state resulting from the non-activation of structural mechanisms initiating decision-making procedures rather than an issue of individual implementation capability.",
    }

    changed = 0
    for (r, c), text in translations.items():
        clear_cell_set_text(tbl.rows[r].cells[c], text)
        changed += 1

    print(f"  Table 4: {changed} cells translated")
    return changed


def translate_table_5(doc):
    """Appendix descriptions table (5x5)."""
    tbl = doc.tables[5]
    translations = {
        # Row 0 headers
        (0, 0): "Category",
        (0, 1): "Appendix Title",
        (0, 2): "Document Type",
        (0, 3): "Date",
        (0, 4): "Main Research Contribution",
        # Row 1
        (1, 0): "Appendix A",
        (1, 1): "Official Correspondence Requesting Digital/AI Capability Enhancement Training for Organizational Staff",
        (1, 2): "Official administrative document\n\n(Provisional receipt number de-identified)",
        (1, 3): "2025. 12. 24.",
        (1, 4): "Decision-making starting point data demonstrating that the organizational-level AI adoption demand was formalized",
        # Row 2
        (2, 0): "Appendix B",
        (2, 1): "External Consulting Report on AI Utilization\n(Summary)",
        (2, 2): "External professional consulting summary",
        (2, 3): "2025. 07.",
        (2, 4): "Preliminary diagnostic data presenting the organization\u2019s digital maturity assessment and the strategic necessity of AI adoption",
        # Row 3
        (3, 0): "Appendix C",
        (3, 1): "External Consulting Report on AI Utilization\n(Final Version Excerpt/Summary)",
        (3, 2): "Analysis-relevant excerpts from the external professional consulting final report",
        (3, 3): "2025. 07.",
        (3, 4): "In-depth analysis of business structure and limitations of AI adoption, utilized as contextual data for Dynamic Capabilities Theory (Seizing)",
        # Row 4
        (4, 0): "Appendix D",
        (4, 1): "Reconstructed Table of \u2018AI Training\u2019 Search Results from Organizational Records Management System",
        (4, 2): "Internal system search results reconstructed summary table\n(De-identified)",
        (4, 3): "2026. 01. 22. / 2026. 02. 09.\n(Sequential query records: de-identified)",
        (4, 4): "Process evidence of structural disconnection showing the absence of implementation-stage records despite the official request (A)",
    }

    changed = 0
    for (r, c), text in translations.items():
        clear_cell_set_text(tbl.rows[r].cells[c], text)
        changed += 1

    print(f"  Table 5: {changed} cells translated")
    return changed


def translate_back_matter(doc):
    """Translate Korean headings and content in References and Appendices sections."""
    changed = 0
    kr_pattern = re.compile(r'[\uac00-\ud7af]')

    back_matter_translations = {
        # References heading
        1060: "REFERENCES CITED",
        # Korean references - romanize authors, add [English translation]
        1124: "Kwon, S., & Yang, J. (2022). An empirical study on supply chain survival strategies in the next normal era: Focusing on the moderating effects of SME risk management orientation and social capital [In Korean]. Journal of the Korean Academy of Business Administration, 35(6), 1161\u20131186. https://doi.org/10.18032/kabaa.2022.35.6.1161",
        1126: "Noh, S. (2017). Problem solving and organizational learning in Korean SMEs: Focusing on subcontracting firms [In Korean] [Doctoral dissertation, Hanyang University].",
        1128: "Jeong, M. (2017). A study on the impact of corporate organizational learning on firm performance: Focusing on the mediating effect of intangible assets [In Korean]. Journal of Digital Convergence, 15(11), 97\u2013105. https://doi.org/10.14400/JDC.2017.15.11.97",
        1130: "Heo, M., & Kim, W. (2024). A conceptual study on the antecedents and effects of organizational structure change in Korean firms [In Korean]. Journal of Business Education Research, 39(3), 97\u2013120. https://doi.org/10.23839/kabe.2024.39.3.97",
        # Appendices heading
        1149: "Appendices",
        # Appendices introductory paragraph
        1150: "The data presented in these appendices consist of existing administrative records and consulting deliverables to which the researcher held legitimate job-related access, de-identified and excerpted/summarized to the extent necessary for analytical purposes. In accordance with organizational security regulations and contractual restrictions, the source data in their entirety are not disclosed externally, and this study presents only the structural significance and analysis-relevant portions of the data.",
        # Appendix D label
        1174: "Appendix D",
        # Appendix D subtitle
        1175: "Reconstructed Summary Table of Internal System Search Results (De-identified)",
    }

    for idx, new_text in back_matter_translations.items():
        if idx < len(doc.paragraphs):
            para = doc.paragraphs[idx]
            clear_paragraph_content(para, new_text)
            changed += 1

    print(f"  Back matter: {changed} paragraphs translated")
    return changed


def verify_korean_residuals(doc):
    """Count Korean characters remaining in tables and back matter."""
    kr_pattern = re.compile(r'[\uac00-\ud7af]')
    residuals = []

    # Check tables 0, 2, 3, 4, 5
    for ti in [0, 2, 3, 4, 5]:
        tbl = doc.tables[ti]
        for ri, row in enumerate(tbl.rows):
            for ci, cell in enumerate(row.cells):
                text = cell.text.strip()
                if kr_pattern.search(text):
                    residuals.append(f"Table {ti} [{ri},{ci}]: {text[:80]}")

    # Check back matter paragraphs (P1060+)
    for i in range(1060, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        text = p.text.strip()
        if kr_pattern.search(text):
            residuals.append(f"P{i} [{p.style.name}]: {text[:80]}")

    return residuals


def main():
    base = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
    working = os.path.join(base, 'output', 'thesis_english.docx')

    if not os.path.exists(working):
        print(f"ERROR: File not found: {working}")
        return

    # Backup
    backup_dir = os.path.join(base, 'backups')
    os.makedirs(backup_dir, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup = os.path.join(backup_dir, f'thesis_pre_tables_{ts}.docx')
    shutil.copy2(working, backup)
    print(f"Backup: {os.path.basename(backup)}")

    # Open document
    doc = Document(working)
    print(f"Tables: {len(doc.tables)}, Paragraphs: {len(doc.paragraphs)}")

    # Translate tables
    print("\n--- Translating Tables ---")
    total = 0
    total += translate_table_0(doc)
    total += translate_table_2(doc)
    total += translate_table_3(doc)
    total += translate_table_4(doc)
    total += translate_table_5(doc)

    # Translate back matter
    print("\n--- Translating Back Matter ---")
    total += translate_back_matter(doc)

    # Save
    doc.save(working)
    print(f"\nSaved: {os.path.basename(working)}")
    print(f"Total changes: {total}")

    # Verify
    print("\n--- Verification ---")
    doc2 = Document(working)
    residuals = verify_korean_residuals(doc2)
    print(f"Korean residuals in tables + back matter: {len(residuals)}")
    if residuals:
        for r in residuals:
            print(f"  {r}")

    # Sample check: show translated table cells
    print("\n--- Sample: Table 2, Row 0 (headers) ---")
    tbl2 = doc2.tables[2]
    for ci in range(6):
        print(f"  [{0},{ci}]: {tbl2.rows[0].cells[ci].text.strip()}")

    print("\n--- Sample: Table 4, Row 1 ---")
    tbl4 = doc2.tables[4]
    for ci in range(4):
        text = tbl4.rows[1].cells[ci].text.strip()[:80]
        print(f"  [{1},{ci}]: {text}")

    print("\n--- Sample: Back matter headings ---")
    for idx in [1060, 1149, 1174]:
        if idx < len(doc2.paragraphs):
            print(f"  P{idx}: {doc2.paragraphs[idx].text.strip()[:100]}")

    return total


if __name__ == '__main__':
    main()
