import docx
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

doc = Document()
for s in doc.sections:
    s.page_width = Cm(21.0)
    s.page_height = Cm(29.7)
    s.top_margin = Cm(4.0)
    s.left_margin = Cm(4.0)
    s.bottom_margin = Cm(3.0)
    s.right_margin = Cm(3.0)

t = doc.add_table(rows=3, cols=3)
t.alignment = WD_TABLE_ALIGNMENT.CENTER
t.autofit = True
tblPr = t._tbl.tblPr

# 1. Paskan ke jendela (100% width = 5000 pct)
tblPr.append(parse_xml(r'<w:tblW %s w:w="5000" w:type="pct"/>' % nsdecls('w')))
# 2. Paskan ke konten (autofit layout)
tblPr.append(parse_xml(r'<w:tblLayout %s w:type="autofit"/>' % nsdecls('w')))

# 3. Semua batas garis (All Borders)
tblBorders = parse_xml(
    r'<w:tblBorders %s>'
    r'  <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
    r'  <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
    r'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
    r'  <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
    r'  <w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
    r'  <w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
    r'</w:tblBorders>' % nsdecls('w')
)
tblPr.append(tblBorders)

for row in t.rows:
    for cell in row.cells:
        cell.text = "Data Test"
        p = cell.paragraphs[0]
        p.paragraph_format.line_spacing = 1.0
        p.runs[0].font.name = "Times New Roman"
        p.runs[0].font.size = Pt(12)

out_file = r"e:\SKILL UNTUK AI\system_prompts_leaks-main\word-addin-tesismagister\output\TEST_BORDER_TABLE.docx"
doc.save(out_file)
print("SUCCESSFULLY GENERATED:", out_file)
