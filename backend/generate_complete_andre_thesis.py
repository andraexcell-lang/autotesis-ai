# backend/generate_complete_andre_thesis.py
"""
Script Pembangun Naskah Master Tesis Magister Manajemen Andre Juliansya
Universitas Tridinanti Palembang (2026)
Memproduksi:
1. TESIS_LENGKAP_ANDRE_JULIANSYA.docx
2. BAGIAN_AWAL_ANDRE_JULIANSYA.docx
3. BAB_1_PENDAHULUAN_ANDRE_JULIANSYA.docx
4. BAB_2_KAJIAN_PUSTAKA_ANDRE_JULIANSYA.docx
5. BAB_3_METODE_PENELITIAN_ANDRE_JULIANSYA.docx
6. BAB_4_HASIL_DAN_PEMBAHASAN_ANDRE_JULIANSYA.docx
7. BAB_5_KESIMPULAN_DAN_DAPUS_ANDRE_JULIANSYA.docx
8. LAMPIRAN_LENGKAP_ANDRE_JULIANSYA.docx
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

import build_andre_thesis

OUTPUT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

def setup_page(doc):
    for s in doc.sections:
        s.page_width = Inches(8.27)     # A4
        s.page_height = Inches(11.69)
        s.top_margin = Inches(1.575)    # 4 cm
        s.left_margin = Inches(1.575)   # 4 cm
        s.bottom_margin = Inches(1.181) # 3 cm
        s.right_margin = Inches(1.181)  # 3 cm

def p(doc, text="", bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0, space_before=0, space_after=0, hanging=0, font_size=12):
    para = doc.add_paragraph()
    para.alignment = align
    para.paragraph_format.line_spacing = 1.5
    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after = Pt(space_after)
    if indent > 0:
        para.paragraph_format.first_line_indent = Inches(indent)
    if hanging > 0:
        para.paragraph_format.first_line_indent = Inches(-hanging)
        para.paragraph_format.left_indent = Inches(hanging)
    if text:
        r = para.add_run(text)
        r.bold = bold
        r.italic = italic
        r.font.name = "Times New Roman"
        r.font.size = Pt(font_size)
        r.font.color.rgb = RGBColor(0, 0, 0)
    return para

def h1(doc, text):
    return p(doc, text, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=14, space_after=6)

def h2(doc, text):
    return p(doc, text, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10, space_after=4)

def h3(doc, text):
    return p(doc, text, bold=True, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=6, space_after=2)

def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(r'<w:tcBorders %s/>' % nsdecls('w'))
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f'w:{edge}'
            el = parse_xml(r'<%s %s w:val="%s" w:sz="%s" w:space="0" w:color="%s"/>' %
                           (tag, nsdecls('w'), edge_data.get('val', 'single'),
                            edge_data.get('sz', '4'), edge_data.get('color', 'auto')))
            tcBorders.append(el)
        else:
            tag = f'w:{edge}'
            el = parse_xml(r'<%s %s w:val="none"/>' % (tag, nsdecls('w')))
            tcBorders.append(el)
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
            for para in cell.paragraphs:
                para.paragraph_format.line_spacing = 1.15
                para.paragraph_format.space_before = Pt(2)
                para.paragraph_format.space_after = Pt(2)
                for run in para.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(10)
                    if r_idx == 0:
                        run.bold = True

def add_table_data(doc, title, headers, data, source="", alignments=None):
    p(doc, title, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=8, space_after=3)
    table = doc.add_table(rows=len(data) + 1, cols=len(headers))
    for c_idx, h in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.text = str(h)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r_idx, row_data in enumerate(data):
        for c_idx, val in enumerate(row_data):
            cell = table.cell(r_idx + 1, c_idx)
            cell.text = str(val)
            if alignments and c_idx < len(alignments):
                cell.paragraphs[0].alignment = alignments[c_idx]
    apply_open_table(table)
    if source:
        p(doc, source, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=8, font_size=10)
    else:
        p(doc, "", space_after=4)

print("Setup helper selesai.")
