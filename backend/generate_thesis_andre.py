# backend/generate_thesis_andre.py
"""
Script Generator Tesis Magister Manajemen Lengkap untuk Andre Juliansya
Format Standar Tesis Pascasarjana Magister Manajemen Universitas Tridinanti Palembang (Rujukan Yurahadi).
Membuat:
1. TESIS_LENGKAP_ANDRE_JULIANSYA.docx (Front Matter + Bab 1-5 + Dapus + Lampiran)
2. Dokumen per-bab (Bagian Awal, Bab 1, Bab 2, Bab 3, Bab 4, Bab 5, Lampiran)
3. DATASET_SMARTPLS_92_RESPONDEN.xlsx (92 responden x 48 indikator)
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

OUTPUT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

def set_page_setup(doc):
    for section in doc.sections:
        section.page_width = Inches(8.27)     # A4: 21.0 cm
        section.page_height = Inches(11.69)   # 29.7 cm
        section.top_margin = Inches(1.575)    # 4.0 cm
        section.left_margin = Inches(1.575)   # 4.0 cm
        section.bottom_margin = Inches(1.181) # 3.0 cm
        section.right_margin = Inches(1.181)  # 3.0 cm

def add_p(doc, text="", bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0, space_before=0, space_after=0, hanging=0, font_size=12):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if indent > 0:
        p.paragraph_format.first_line_indent = Inches(indent)
    if hanging > 0:
        p.paragraph_format.first_line_indent = Inches(-hanging)
        p.paragraph_format.left_indent = Inches(hanging)
    if text:
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.name = "Times New Roman"
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_h1(doc, text):
    return add_p(doc, text, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=14, space_after=6)

def add_h2(doc, text):
    return add_p(doc, text, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10, space_after=4)

def add_h3(doc, text):
    return add_p(doc, text, bold=True, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=6, space_after=2)

def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(r'<w:tcBorders %s/>' % nsdecls('w'))
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f'w:{edge}'
            element = parse_xml(r'<%s %s w:val="%s" w:sz="%s" w:space="0" w:color="%s"/>' %
                                (tag, nsdecls('w'), edge_data.get('val', 'single'),
                                 edge_data.get('sz', '4'), edge_data.get('color', 'auto')))
            tcBorders.append(element)
        else:
            tag = f'w:{edge}'
            element = parse_xml(r'<%s %s w:val="none"/>' % (tag, nsdecls('w')))
            tcBorders.append(element)
    tcPr.append(tcBorders)

def apply_open_table(table):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    rows = table.rows
    num_rows = len(rows)
    for r_idx, row in enumerate(rows):
        for cell in row.cells:
            if r_idx == 0:
                set_cell_border(cell, top=dict(sz="12", val="single"), bottom=dict(sz="8", val="single"))
            elif r_idx == num_rows - 1:
                set_cell_border(cell, bottom=dict(sz="12", val="single"))
            else:
                set_cell_border(cell)
            for p in cell.paragraphs:
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(10 if r_idx > 0 else 10)
                    if r_idx == 0:
                        run.bold = True

def add_table_data(doc, title, headers, data, source=""):
    add_p(doc, title, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=6, space_after=2)
    table = doc.add_table(rows=len(data) + 1, cols=len(headers))
    for col_idx, h in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = str(h)
    for row_idx, row_data in enumerate(data):
        for col_idx, val in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = str(val)
    apply_open_table(table)
    if source:
        add_p(doc, source, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=6, font_size=10)
    else:
        add_p(doc, "", space_after=4)

print("Helper functions ready.")
