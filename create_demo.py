"""
Generate a demo .docx file with deliberately messy formatting
to demonstrate what DOCX Rebuilder fixes.
"""

import zipfile
import copy
from pathlib import Path
from lxml import etree
from docx import Document
from docx.shared import Pt, Inches, Twips, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

def create_messy_demo():
    """Create a legal document with realistic formatting problems."""
    doc = Document()

    # Set default style to Times New Roman 11pt
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    # ========================================
    # TITLE - clean
    # ========================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('COMMERCIAL LEASE AGREEMENT')
    run.bold = True
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('(Multi-Tenant Office Building)')
    run.font.size = Pt(11)
    run.font.name = 'Times New Roman'

    doc.add_paragraph()  # blank line

    # ========================================
    # RECITAL PARAGRAPH - has mixed fonts (the mess)
    # ========================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    run = p.add_run('\tTHIS COMMERCIAL LEASE AGREEMENT (this ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # "Agreement" in bold - but someone used Cambria by accident
    run = p.add_run('"Agreement"')
    run.bold = True
    run.font.name = 'Cambria'  # WRONG FONT - should be Times New Roman
    run.font.size = Pt(11)

    run = p.add_run(') is entered into as of _______________, 2024 (the ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('"Effective Date"')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('), by and between ABC PROPERTIES, LLC, a Delaware limited liability company (')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # "Landlord" in bold - but someone used Arial
    run = p.add_run('"Landlord"')
    run.bold = True
    run.font.name = 'Arial'  # WRONG FONT
    run.font.size = Pt(11)

    run = p.add_run('), and XYZ CORPORATION, a California corporation (')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('"Tenant"')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run(').')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    doc.add_paragraph()  # blank line

    # ========================================
    # SECTION 1 - uses inconsistent indentation
    # ========================================
    # Section heading - bold, underline
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run('\t1.\tPREMISES AND TERM')
    run.bold = True
    run.underline = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # 1.1 - correct indent
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Twips(720)  # 0.5 inch

    run = p.add_run('1.1\tPremises.')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('  Landlord hereby leases to Tenant, and Tenant hereby leases from Landlord, those certain premises consisting of approximately 5,000 rentable square feet located on the third (3rd) floor of the building located at 123 Main Street, San Francisco, California 94105 (the ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('"Building"')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('), as more particularly described on ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('Exhibit A')
    run.bold = True
    run.underline = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run(' attached hereto (the ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('"Premises"')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run(').')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # 1.2 - WRONG indent (slightly off - 700 instead of 720)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Twips(700)  # SLIGHTLY OFF

    run = p.add_run('1.2\tTerm.')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('  The initial term of this Lease shall commence on the date (the ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('"Commencement Date"')
    run.bold = True
    run.font.name = 'Times New Roman'
    # WRONG SIZE - 10.5pt instead of 11pt
    run.font.size = Pt(10.5)

    run = p.add_run(') that is the later of (i) January 1, 2025, and (ii) the date on which Landlord delivers the Premises to Tenant in the condition required by Section 4, and shall expire on December 31, 2029 (the ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('"Expiration Date"')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('), unless sooner terminated in accordance with the provisions of this Lease.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # 1.3 - WRONG indent (750 instead of 720) and LEFT aligned instead of JUSTIFIED
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # WRONG - should be JUSTIFY

    p.paragraph_format.left_indent = Twips(750)  # WRONG - should be 720

    run = p.add_run('1.3\tRenewal Option.')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('  Provided that Tenant is not then in default beyond any applicable cure period, Tenant shall have the option to extend the Term for one (1) additional period of five (5) years (the ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('"Renewal Term"')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run(') by delivering written notice to Landlord no later than nine (9) months prior to the Expiration Date.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    doc.add_paragraph()

    # ========================================
    # SECTION 2 - more inconsistencies
    # ========================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run('\t2.\tRENT')
    run.bold = True
    run.underline = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # 2.1 - correct
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Twips(720)

    run = p.add_run('2.1\tBase Rent.')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('  Commencing on the Commencement Date and continuing throughout the Term, Tenant shall pay to Landlord monthly base rent in the amount of ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # Someone pasted from Excel with Calibri
    run = p.add_run('$25,000.00')
    run.bold = True
    run.font.name = 'Calibri'  # WRONG FONT - pasted from spreadsheet
    run.font.size = Pt(11)

    run = p.add_run(' per month (the ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('"Base Rent"')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('), subject to the annual increases set forth in Section 2.2 below.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # 2.2 - indented correctly, but mixed fonts in the middle
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Twips(720)

    run = p.add_run('2.2\tAnnual Increases.')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('  On each anniversary of the Commencement Date during the Term, Base Rent shall increase by three percent ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # Someone typed in Courier New
    run = p.add_run('(3%)')
    run.font.name = 'Courier New'  # WRONG FONT
    run.font.size = Pt(11)

    run = p.add_run(' over the Base Rent payable during the immediately preceding Lease Year.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # 2.3 - wrong indent AND wrong alignment
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # WRONG
    p.paragraph_format.left_indent = Twips(680)  # WRONG

    run = p.add_run('2.3\tAdditional Rent.')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('  In addition to Base Rent, Tenant shall pay to Landlord as additional rent all other amounts due under this Lease, including without limitation Tenant\'s Proportionate Share of Operating Expenses as provided in Section 5.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    doc.add_paragraph()

    # ========================================
    # SECTION 3 - with a table (rent schedule)
    # ========================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run('\t3.\tRENT SCHEDULE')
    run.bold = True
    run.underline = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run('\tThe following table sets forth the Base Rent payable during each Lease Year:')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    doc.add_paragraph()

    # Add table with mixed formatting
    table = doc.add_table(rows=6, cols=3)
    table.style = 'Table Grid'

    # Header row
    headers = ['Lease Year', 'Monthly Base Rent', 'Annual Base Rent']
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(10)

    # Data rows with MIXED FONTS in some cells
    data = [
        ('Year 1', '$25,000.00', '$300,000.00'),
        ('Year 2', '$25,750.00', '$309,000.00'),
        ('Year 3', '$26,522.50', '$318,270.00'),
        ('Year 4', '$27,318.18', '$327,818.10'),
        ('Year 5', '$28,137.72', '$337,652.64'),
    ]

    for row_idx, row_data in enumerate(data):
        for col_idx, cell_text in enumerate(row_data):
            cell = table.rows[row_idx + 1].cells[col_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(cell_text)
            run.font.size = Pt(10)

            # Mess up fonts in some cells
            if row_idx == 1 and col_idx == 1:
                run.font.name = 'Arial'  # WRONG
            elif row_idx == 3 and col_idx == 2:
                run.font.name = 'Calibri'  # WRONG
            else:
                run.font.name = 'Times New Roman'

    doc.add_paragraph()

    # ========================================
    # SECTION 4 - with sub-sub-sections
    # ========================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run('\t4.\tCONDITION OF PREMISES')
    run.bold = True
    run.underline = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # 4.1 with subsections (a), (b), (c)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Twips(720)

    run = p.add_run('4.1\tLandlord\'s Work.')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('  Landlord shall, at Landlord\'s sole cost and expense, perform the following work in the Premises prior to the Commencement Date:')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # Sub-items with inconsistent indentation
    # (a) - correct indent
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Twips(1440)  # 1 inch

    run = p.add_run('(a)\tReplace all existing carpeting with building-standard carpet tile;')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # (b) - WRONG indent (1400 instead of 1440)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Twips(1400)  # WRONG

    run = p.add_run('(b)\tRepaint all walls and ceilings with building-standard paint;')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # (c) - WRONG indent AND wrong font
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Twips(1500)  # WRONG

    run = p.add_run('(c)')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('\tReplace all HVAC filters and confirm proper operation of the heating, ventilation, and air conditioning systems serving the Premises;')
    run.font.name = 'Arial'  # WRONG
    run.font.size = Pt(11)

    # (d) - correct
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Twips(1440)

    run = p.add_run('(d)\tClean and repair all window coverings and replace any damaged blinds.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    doc.add_paragraph()

    # ========================================
    # SECTION 5 - with some runs that have wrong colors
    # ========================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run('\t5.\tOPERATING EXPENSES')
    run.bold = True
    run.underline = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Twips(720)

    run = p.add_run('5.1\tTenant\'s Proportionate Share.')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('  Tenant\'s Proportionate Share of Operating Expenses shall be ')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # Someone left track-changes red text that was accepted but color stuck
    run = p.add_run('twelve and one-half percent (12.5%)')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)  # RED - leftover from tracked changes

    run = p.add_run(', calculated by dividing the rentable square footage of the Premises (5,000 RSF) by the total rentable square footage of the Building (40,000 RSF).')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # 5.2 - indent is off
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Twips(710)  # SLIGHTLY OFF

    run = p.add_run('5.2\tEstimated Payments.')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    run = p.add_run('  Landlord shall provide Tenant with a written estimate of Operating Expenses for each calendar year. Tenant shall pay one-twelfth (1/12) of Tenant\'s Proportionate Share of the estimated Operating Expenses on a monthly basis, together with each installment of Base Rent.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    doc.add_paragraph()

    # ========================================
    # SIGNATURE BLOCK
    # ========================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('[Signature Page Follows]')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)
    run.italic = True

    doc.add_page_break()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run('IN WITNESS WHEREOF, the parties have executed this Agreement as of the date first written above.')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    doc.add_paragraph()
    doc.add_paragraph()

    # Landlord sig block
    p = doc.add_paragraph()
    run = p.add_run('LANDLORD:')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    p = doc.add_paragraph()
    run = p.add_run('ABC PROPERTIES, LLC')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run('By: ________________________________')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    p = doc.add_paragraph()
    run = p.add_run('Name: ______________________________')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    p = doc.add_paragraph()
    run = p.add_run('Title: ______________________________')
    # WRONG - someone switched to Calibri
    run.font.name = 'Calibri'
    run.font.size = Pt(11)

    doc.add_paragraph()
    doc.add_paragraph()

    # Tenant sig block
    p = doc.add_paragraph()
    run = p.add_run('TENANT:')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    p = doc.add_paragraph()
    run = p.add_run('XYZ CORPORATION')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run('By: ________________________________')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    p = doc.add_paragraph()
    run = p.add_run('Name: ______________________________')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    p = doc.add_paragraph()
    run = p.add_run('Title: ______________________________')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)

    # Save
    output = Path('demo_messy.docx')
    doc.save(output)
    print(f'Created: {output}')
    print(f'This document has deliberate formatting problems:')
    print(f'  - Mixed fonts: Cambria, Arial, Calibri, Courier New mixed with Times New Roman')
    print(f'  - Inconsistent indentation: 680, 700, 710, 750 twips instead of 720')
    print(f'  - Wrong alignment: some paragraphs LEFT instead of JUSTIFIED')
    print(f'  - Wrong font size: 10.5pt mixed in with 11pt')
    print(f'  - Leftover red text from tracked changes')
    print(f'  - Table cells with wrong fonts')
    print(f'  - Sub-section (a)(b)(c)(d) with inconsistent indent levels')


if __name__ == '__main__':
    create_messy_demo()
