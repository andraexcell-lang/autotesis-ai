# backend/build_andre_thesis.py
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

OUTPUT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

# TABULASI DATA 92 RESPONDEN
# MK: 15 butir
MK_RAW = """
5 5 5 4 4 4 4 4 5 4 5 5 4 4 4
4 5 4 4 4 5 5 5 5 5 5 5 5 4 4
4 4 4 4 3 4 4 4 4 4 4 4 4 3 4
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
4 4 4 4 3 3 4 4 3 3 4 3 3 3 4
5 5 5 5 5 4 4 5 4 5 4 4 5 4 4
3 3 4 4 4 3 3 4 3 4 3 3 4 4 4
4 4 4 4 4 4 4 4 4 4 5 5 4 4 2
1 2 2 2 2 1 2 2 1 2 1 2 2 2 4
3 3 4 4 3 3 3 3 3 3 3 4 3 4 3
3 3 3 3 3 3 3 3 3 3 3 3 3 3 4
4 5 4 4 4 4 4 4 4 4 4 4 4 4 5
5 5 5 5 5 5 5 5 5 5 5 5 5 5 4
4 4 4 5 4 4 3 4 4 4 4 4 4 4 5
4 5 4 5 4 5 4 5 5 5 4 4 4 4 5
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
5 5 5 5 5 5 5 5 5 5 5 5 5 5 3
4 4 3 3 4 3 4 3 4 3 4 4 4 4 4
4 4 4 3 4 4 3 5 4 4 4 4 3 4 4
5 4 5 4 5 4 5 5 4 5 5 5 4 4 5
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
4 4 4 3 4 4 4 4 4 4 4 4 5 4 4
4 4 4 4 4 4 4 4 4 4 4 4 4 3 4
2 2 2 2 2 2 2 2 2 2 2 2 3 2 2
2 2 2 2 3 2 3 3 3 2 3 2 2 2 2
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
3 3 3 3 3 3 4 3 3 4 4 3 3 4 3
4 3 3 4 4 4 4 4 4 4 4 3 3 3 4
3 3 3 3 2 2 2 3 2 3 3 2 3 3 3
3 3 3 3 3 3 3 3 2 2 2 2 2 2 2
4 4 4 4 4 4 5 4 4 4 4 5 4 4 4
5 4 4 5 5 5 5 4 4 4 4 5 5 4 5
3 3 2 3 3 3 3 3 3 3 3 3 3 3 3
4 4 4 4 4 3 4 5 4 4 4 4 4 4 4
5 5 4 4 5 5 5 5 5 4 5 4 4 5 5
4 5 4 4 4 4 4 4 4 5 4 4 4 5 4
4 4 4 4 4 4 4 3 4 4 3 4 3 4 4
5 4 5 5 5 4 5 5 5 5 5 5 5 5 5
4 4 4 4 4 5 4 4 5 4 4 3 4 5 5
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
4 3 4 3 4 2 3 4 4 3 3 3 3 4 4
2 2 2 3 3 2 2 3 2 3 3 2 3 2 2
5 5 4 5 4 4 4 4 5 5 4 5 4 4 5
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
3 3 3 3 3 3 2 3 3 3 2 2 3 2 2
4 4 4 4 4 4 4 4 3 3 4 4 3 4 4
3 2 3 4 3 3 3 3 3 4 4 3 4 4 3
4 3 3 4 4 4 4 4 4 4 4 4 4 4 4
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
4 4 4 4 4 4 4 3 4 4 4 4 4 4 4
4 4 5 5 5 4 4 5 5 4 5 5 4 4 5
2 2 2 2 2 2 3 3 2 2 2 3 3 3 2
4 4 4 5 4 4 4 4 4 4 4 4 4 4 4
4 4 5 4 4 5 5 5 4 4 5 5 5 4 5
4 5 5 4 4 5 5 4 4 4 5 4 5 4 4
4 4 4 4 3 3 3 4 4 4 4 4 4 4 4
4 5 4 4 4 4 4 5 5 4 4 4 4 4 4
4 3 3 4 3 3 3 3 3 3 3 3 4 3 3
5 5 4 5 4 5 4 5 5 5 4 5 5 5 4
4 4 4 3 4 4 4 4 4 4 4 4 4 3 3
4 4 4 5 5 4 5 5 5 4 5 4 4 4 5
4 4 4 3 3 3 4 4 3 3 4 3 4 3 3
2 3 3 3 3 3 3 3 3 3 3 3 3 3 3
3 3 3 3 3 3 3 3 3 4 3 3 3 3 3
3 4 5 4 3 5 4 4 4 4 5 4 4 3 4
4 5 5 5 5 5 5 5 5 5 5 5 5 5 5
5 4 4 4 4 4 4 5 4 4 4 4 5 5 4
4 4 4 3 3 3 4 4 4 3 4 4 4 4 4
4 4 4 3 4 4 4 3 3 3 4 4 4 3 4
3 4 4 3 3 4 3 3 3 3 3 3 3 3 4
5 5 4 5 5 5 5 5 4 5 5 5 4 4 4
4 3 3 3 4 3 3 3 3 3 3 3 3 4 3
4 4 4 4 4 4 4 3 4 4 4 4 4 4 4
3 3 3 3 3 3 3 3 3 3 3 3 3 4 3
4 5 4 4 4 4 4 4 4 4 4 4 4 4 4
4 4 4 4 3 4 4 3 4 3 3 4 4 5 5
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
5 4 5 4 4 5 4 4 5 4 4 5 3 5 4
4 4 4 4 3 4 4 4 4 4 4 4 3 3 3
3 4 3 4 3 3 3 4 3 4 4 3 4 4 4
4 4 4 4 4 4 4 4 3 4 4 4 4 3 4
3 3 4 4 3 4 4 4 4 4 3 4 3 4 4
3 3 4 4 4 4 4 4 4 3 5 4 4 4 3
4 3 4 4 3 4 3 4 4 4 3 4 4 4 4
4 4 4 3 4 4 4 4 4 4 4 4 5 5 5
4 4 4 5 5 5 4 4 5 5 5 5 5 5 5
4 5 4 5 4 4 4 5 4 4 4 5 4 3 4
3 3 3 3 3 3 3 4 3 3 3 3 5 4 4
4 3 3 4 5 4 4 4 5 5 4 4 4 5 4
4 5 4 4 4 4 4 5 5 4 4 5 3 4 3
3 3 3 4 3 4 3 3 3 3 4 4 3 3 3
4 4 4 4 3 3 3 3 3 3 3 4
"""

# DK: 9 butir
DK_RAW = """
5 5 5 5 5 5 5 5 5
3 3 3 3 4 3 4 4 3
5 5 5 5 5 5 5 5 5
5 5 5 5 5 5 5 5 5
3 2 3 3 3 3 3 3 3
5 4 5 4 4 4 5 5 5
3 3 3 3 3 3 3 3 3
4 4 4 5 4 4 4 4 5
4 3 4 4 4 4 4 3 4
3 4 3 3 4 4 4 4 3
4 4 4 4 5 4 4 5 4
5 5 5 5 5 5 5 5 5
5 5 5 5 5 4 5 5 5
3 3 3 3 3 3 3 3 3
4 4 4 4 3 3 4 4 4
2 3 3 3 3 3 3 3 3
5 5 5 5 5 5 5 5 5
4 4 3 3 3 4 3 4 3
4 5 5 4 4 4 4 5 5
5 5 5 5 4 5 5 5 5
4 4 4 4 4 4 4 4 4
5 4 5 4 4 4 4 5 5
4 4 4 4 3 4 4 4 4
3 3 3 2 3 3 2 2 2
5 5 5 4 4 4 5 4 5
5 5 5 5 5 5 5 4 5
2 3 3 3 3 2 2 3 3
5 4 5 5 5 4 4 5 5
4 4 4 4 5 5 4 4 4
4 4 4 4 4 3 4 4 4
5 5 5 5 5 5 5 5 5
4 3 4 3 3 3 4 3 3
3 3 4 3 3 4 4 3 4
4 4 4 4 4 3 4 5 5
5 5 5 5 5 5 5 3 4
4 4 3 3 3 4 3 3 3
3 3 3 3 3 3 2 5 5
5 5 5 5 5 5 5 5 4
5 5 4 5 4 4 5 4 4
4 5 4 4 4 4 4 3 3
4 3 4 3 3 3 3 3 4
4 3 3 3 4 3 3 4 3
4 4 4 5 5 3 3 4 4
4 5 4 4 4 4 4 3 3
3 2 3 3 3 3 2 3 3
3 3 3 2 3 3 3 4 5
5 5 5 5 4 5 4 5 5
5 5 5 5 5 5 5 4 5
4 4 4 4 5 4 4 3 4
4 4 4 3 4 4 4 2 3
3 3 3 2 2 3 2 5 4
5 5 5 5 4 5 5 4 4
5 4 4 4 4 4 4 3 4
3 3 3 3 4 2 3 4 3
3 3 2 3 3 2 3 4 5
5 5 5 5 5 5 5 5 5
5 5 5 5 5 5 5 3 4
4 4 3 4 3 3 3 5 5
5 5 4 5 5 4 4 2 2
2 2 1 2 2 2 2 4 5
4 4 4 4 4 4 4 4 5
5 4 4 4 4 5 5 3 4
3 3 3 3 3 3 3 4 4
4 4 4 3 4 4 4 3 3
3 3 3 3 4 3 3 4 4
5 4 4 5 4 5 4 5 5
5 5 5 5 5 5 5 4 3
4 3 4 4 4 3 4 5 5
5 5 5 5 5 5 5 3 3
3 3 3 3 2 2 3 4 5
5 5 5 5 5 4 5 4 4
3 3 3 4 4 3 4 4 4
4 4 4 4 4 4 3 4 5
4 5 5 4 5 4 5 5 4
4 4 4 4 4 4 5 4 5
4 5 5 4 4 4 5 4 4
5 5 5 4 4 5 5 5 5
5 5 5 4 4 5 3 4 4
4 3 4 4 4 3 3 3 3
3 3 3 3 4 3 4 4 3
4 4 4 4 4 4 5 4 5
5 5 4 4 5 5 3 3 3
3 3 4 3 4 3 3 2 2
3 2 2 3 3 2 4 3 4
4 4 4 3 4 3 5 5 4
5 5 4 5 5 5 5 5 5
5 5 4 5 5 4 5 5 5
4 4 5 5 5 5 3 4 4
4 4 4 4 4 4 4 5 5
4 4 5 5 4 4 4 4 5
5 5 5 5 5 5 3 3 3
3 3 3 4 4 3
"""

# EE: 9 butir
EE_RAW = """
5 5 5 5 5 5 5 5 5
3 4 4 3 4 4 4 4 4
5 5 5 5 5 5 5 5 5
5 5 5 5 5 5 5 5 5
4 3 4 4 3 3 2 3 3
5 5 5 5 4 4 5 5 5
4 3 3 3 3 3 3 3 3
5 4 4 4 4 4 4 4 4
2 3 3 2 3 3 2 2 2
4 3 3 4 4 4 3 3 4
4 3 4 4 4 4 4 4 3
5 5 5 5 5 5 5 5 5
5 5 5 5 5 5 5 5 5
2 3 3 3 3 3 2 3 3
4 4 5 5 4 4 3 4 5
5 5 5 5 5 5 5 5 5
5 5 5 5 5 5 5 5 5
4 3 3 3 4 4 4 3 4
4 3 4 4 4 4 4 3 4
4 5 5 5 5 5 5 4 5
4 4 4 4 5 4 4 4 5
4 4 4 5 4 4 5 4 4
4 4 4 3 4 3 4 4 4
3 2 3 2 2 3 2 2 2
4 4 3 4 5 4 4 4 3
5 5 4 5 5 5 5 5 5
2 2 2 3 3 3 2 3 2
3 3 3 3 4 3 4 4 4
3 3 3 3 3 3 2 3 3
3 3 4 3 3 2 3 4 3
4 4 4 4 4 4 4 4 4
5 5 4 4 4 5 4 4 4
3 3 3 3 3 2 3 3 4
4 4 4 4 4 5 5 4 4
5 5 5 5 5 5 5 5 3
3 4 3 3 4 4 4 4 4
4 4 3 4 3 4 4 4 5
5 5 5 5 5 5 5 5 4
5 5 4 5 4 5 5 5 5
5 5 5 5 5 4 4 5 3
3 3 3 4 3 3 3 4 3
3 3 3 3 3 3 3 3 5
4 4 5 4 4 4 5 4 5
4 4 5 4 4 5 4 5 3
3 3 3 3 3 3 3 3 3
3 3 3 3 3 3 4 5 4
4 5 4 4 4 4 5 3 3
4 4 4 4 3 3 3 4 4
4 4 4 5 4 5 4 4 4
4 4 4 4 4 4 4 4 5
4 4 4 4 5 4 5 4 5
4 4 4 4 4 4 4 5 4
4 5 3 5 4 4 4 3 3
4 3 3 3 3 3 3 4 4
4 5 5 4 4 5 4 4 4
4 5 5 4 4 5 4 5 5
5 5 5 5 5 5 5 3 4
4 3 3 3 3 3 4 5 5
5 5 5 5 5 5 5 4 4
4 4 4 4 4 4 5 4 4
5 4 4 4 4 4 4 4 4
4 4 4 4 4 4 4 3 2
2 3 3 3 4 3 3 3 4
3 4 3 4 3 4 3 4 4
4 5 4 4 4 4 4 5 5
5 5 5 5 5 5 5 5 5
5 5 5 5 5 5 3 3 4
4 4 4 4 4 4 5 5 5
5 5 5 5 5 5 3 3 3
3 3 4 3 3 3 4 5 4
4 4 4 4 4 4 2 2 3
2 3 3 3 2 3 4 5 4
4 4 5 4 4 4 4 4 3
4 4 3 3 3 4 4 4 4
3 3 4 4 3 4 5 3 5
4 4 4 4 3 4 5 5 5
5 5 5 5 5 5 4 4 4
4 4 4 4 4 4 3 3 3
4 3 4 3 2 3 2 3 3
3 3 3 3 3 2 4 3 3
4 4 4 3 4 4 3 4 4
4 3 3 3 3 4 3 3 3
2 4 3 3 3 3 3 2 3
3 3 3 3 3 3 4 4 5
4 4 4 4 4 4 4 4 5
5 4 5 4 4 5 3 4 4
4 4 4 4 4 4 4 4 4
4 5 4 4 5 3 3 3 3
3 3 3 4 4 4 4 5 4
4 5 4 5 5 4 4 3 4
4 4 3 4 4 3 3 3 4
3 3 3 3 3
"""

# KP: 15 butir
KP_RAW = """
5 5 5 5 5 5 4 5 5 5 5 5 5 5 5
4 4 4 4 3 4 3 4 4 4 4 4 4 3 4
5 5 4 5 5 5 5 5 5 5 4 4 5 5 5
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
3 3 4 3 2 3 3 3 3 3 3 4 4 3 3
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
3 2 4 4 3 3 3 3 2 3 3 3 2 3 3
5 4 4 4 5 4 5 5 4 4 4 5 5 4 4
3 2 3 3 3 3 3 3 3 2 3 3 3 2 3
3 3 4 3 4 4 4 3 3 4 4 4 3 4 4
4 4 4 4 5 4 4 4 4 4 4 4 4 4 4
5 5 5 5 5 4 5 5 5 5 5 5 5 5 5
5 4 4 4 4 5 5 4 5 5 5 5 5 5 5
4 4 4 4 3 3 3 3 4 4 3 4 3 4 4
5 4 4 5 4 5 5 4 4 4 5 5 5 4 4
5 4 5 4 5 4 5 4 5 5 4 5 5 5 5
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
3 4 4 3 4 3 4 3 3 3 3 4 3 3 4
4 4 4 3 4 4 4 4 4 4 4 3 4 4 4
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
5 4 4 4 5 4 4 5 4 4 4 4 4 4 4
4 4 4 4 4 4 3 4 4 4 3 4 4 4 4
4 4 4 4 4 4 4 3 3 3 4 4 3 2 1
2 1 2 2 2 1 2 2 2 1 2 2 2 4 4
4 3 4 3 4 4 4 4 4 4 4 4 4 5 5
5 5 5 5 4 4 5 5 5 5 5 4 5 2 2
3 2 2 2 2 2 2 2 2 2 3 2 2 4 4
3 4 4 4 4 3 4 4 3 4 3 4 3 4 3
3 3 3 4 3 3 3 4 3 3 2 3 3 3 3
3 3 3 3 4 4 3 3 3 3 3 3 3 3 5
5 4 4 5 4 4 4 5 5 4 4 4 4 4 5
4 5 4 5 5 5 5 5 5 5 4 5 5 5 3
2 3 3 3 3 3 3 3 3 3 3 3 3 3 4
4 4 4 4 3 4 5 4 4 4 4 5 5 4 4
5 5 5 5 5 5 5 5 5 5 5 5 5 5 3
4 4 4 4 4 4 4 4 4 4 3 4 4 3 4
4 4 4 4 4 4 3 4 4 4 4 3 3 4 5
5 5 5 5 5 5 5 5 5 5 5 5 5 5 5
5 5 5 4 5 5 5 5 5 5 5 5 5 5 4
5 4 4 4 4 4 5 4 4 4 4 4 4 4 3
3 3 4 3 3 3 3 3 3 3 3 4 3 3 3
4 2 3 3 3 3 3 3 3 3 4 3 3 5 4
5 5 5 4 5 5 5 5 5 4 5 5 5 4 4
4 5 5 4 4 4 5 5 5 5 4 4 4 2 3
2 3 2 3 1 2 2 2 2 2 3 2 2 4 4
4 3 4 3 4 3 3 3 3 3 4 3 3 4 5
4 4 5 4 4 4 4 4 4 4 4 3 4 4 4
4 4 4 4 4 4 3 4 3 4 4 5 4 4 4
4 5 4 4 4 4 5 5 4 4 5 5 4 5 4
4 5 4 5 5 4 5 5 4 4 5 4 4 4 3
4 4 4 4 4 4 4 4 4 3 4 3 5 4 4
4 4 4 4 4 3 4 4 4 4 3 4 3 3 4
4 3 4 4 4 3 4 4 4 4 3 4 3 3 3
3 4 4 3 4 3 3 3 4 4 3 4 4 3 3
3 3 4 4 4 4 3 4 4 5 4 5 5 4 5
4 4 4 4 4 5 4 5 4 4 4 3 5 4 5
5 4 5 4 4 4 4 5 4 4 4 4 5 4 4
4 5 4 4 5 4 5 4 4 4 5 4 5 4 5
5 4 5 4 5 5 4 5 5 3 3 2 3 3 3
3 3 3 3 3 3 3 3 3 4 4 4 4 4 4
4 5 4 4 4 4 4 4 4 4 4 4 5 4 4
5 5 5 4 5 4 4 4 4 3 3 3 3 3 3
2 3 3 3 2 3 3 3 3 3 4 4 3 4 3
3 4 4 4 4 3 4 3 4 4 4 4 4 4 4
4 4 4 3 3 4 4 4 4 5 5 5 5 5 4
5 5 5 5 5 5 4 5 5 5 5 5 5 5 5
5 5 5 5 5 5 5 5 3 3 4 3 3 3 3
3 4 3 4 4 4 3 3 5 5 4 5 5 5 5
5 5 5 5 5 5 5 5 4 3 3 3 3 3 3
4 3 3 3 4 3 3 3 3 4 3 4 4 4 3
4 3 4 4 4 4 4 4 3 3 3 3 2 3 3
3 3 2 2 2 3 3 3 4 5 5 4 5 4 5
4 5 5 5 4 5 5 4 4 3 3 4 3 4 3
3 4 3 4 4 4 4 3 3 3 4 3 4 4 3
4 4 3 3 3 4 3 4 4 4 5 4 4 4 4
4 4 4 4 4 4 5 4 5 5 5 5 5 5 5
5 5 5 5 5 5 5 5 4 5 5 5 5 5 5
5 4 5 4 5 5 5 5 4 4 4 4 4 4 4
4 4 4 4 4 3 4 3 2 2 3 3 2 3 3
2 2 2 3 2 3 3 2 4 4 5 4 4 4 5
4 4 4 4 4 4 5 4 3 3 3 4 4 3 4
4 4 4 4 3 4 4 4 3 3 3 2 3 3 3
2 3 3 3 3 3 3 3 3 2 3 3 3 3 3
2 3 2 3 2 3 2 4 4 4 4 4 4 4 4
4 4 4 4 4 4 3 5 5 5 5 5 5 5 5
5 4 5 4 5 5 5 4 4 4 4 4 4 4 4
4 4 4 4 4 5 4 4 4 5 5 4 5 4 5
4 5 4 4 4 5 5 4 3 4 4 4 5 3 4
4 4 4 4 3 4 4 5 4 4 4 4 5 4 5
4 4 4 5 5 5 4 4 4 4 4 4 5 4 4
4 4 4 4 5 4 4 3 4 3 3 4 4 4 4
4 3 4 4 4 3
"""

def parse_matrix(raw_str):
    lines = [line.strip().split() for line in raw_str.strip().split("\n") if line.strip()]
    flat = []
    for line in lines:
        flat.extend([int(x) for x in line])
    return flat

mk_vals = parse_matrix(MK_RAW)
dk_vals = parse_matrix(DK_RAW)
ee_vals = parse_matrix(EE_RAW)
kp_vals = parse_matrix(KP_RAW)

# Strukturkan per responden
respondents = []
for i in range(92):
    r_mk = mk_vals[i*15 : (i+1)*15]
    r_dk = dk_vals[i*9 : (i+1)*9]
    r_ee = ee_vals[i*9 : (i+1)*9]
    r_kp = kp_vals[i*15 : (i+1)*15]
    respondents.append({
        "No": i + 1,
        "MK": r_mk,
        "DK": r_dk,
        "EE": r_ee,
        "KP": r_kp
    })

# BUAT EXCEL DATASET
def generate_excel():
    wb = openpyxl.Workbook()
    # Sheet 1: Master Dataset SmartPLS
    ws = wb.active
    ws.title = "Dataset_SmartPLS_92"
    
    headers = ["No_Resp"]
    headers += [f"MK{i+1:02d}" for i in range(15)]
    headers += [f"DK{i+1:02d}" for i in range(9)]
    headers += [f"EE{i+1:02d}" for i in range(9)]
    headers += [f"KP{i+1:02d}" for i in range(15)]
    
    ws.append(headers)
    for r in respondents:
        row = [r["No"]] + r["MK"] + r["DK"] + r["EE"] + r["KP"]
        ws.append(row)
        
    # Styling header
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for row in ws.iter_rows(min_row=2, max_row=93, min_col=1, max_col=49):
        for cell in row:
            cell.font = Font(name="Times New Roman", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
    # Sheet per variabel
    var_specs = [
        ("Motivasi_Kerja", [f"MK{i+1:02d}" for i in range(15)], "MK"),
        ("Disiplin_Kerja", [f"DK{i+1:02d}" for i in range(9)], "DK"),
        ("Employee_Engagement", [f"EE{i+1:02d}" for i in range(9)], "EE"),
        ("Kinerja_Pegawai", [f"KP{i+1:02d}" for i in range(15)], "KP"),
    ]
    for title, v_headers, key in var_specs:
        ws_v = wb.create_sheet(title=title)
        ws_v.append(["No"] + v_headers + ["Total", "Mean"])
        for r in respondents:
            scores = r[key]
            tot = sum(scores)
            avg = round(tot / len(scores), 2)
            ws_v.append([r["No"]] + scores + [tot, avg])
        for cell in ws_v[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
    excel_path = os.path.join(OUTPUT_DIR, "DATASET_SMARTPLS_92_RESPONDEN.xlsx")
    wb.save(excel_path)
    print(f"Saved Excel dataset to: {excel_path}")


if __name__ == "__main__":
    # Efek samping (cetak + tulis Excel) hanya saat dijalankan langsung.
    # Impor modul oleh server TIDAK boleh menghasilkan berkas data diam-diam.
    print(f"Total MK values: {len(mk_vals)} (Expected {92*15} = {92*15})")
    print(f"Total DK values: {len(dk_vals)} (Expected {92*9} = {92*9})")
    print(f"Total EE values: {len(ee_vals)} (Expected {92*9} = {92*9})")
    print(f"Total KP values: {len(kp_vals)} (Expected {92*15} = {92*15})")
    print(f"Parsed {len(respondents)} respondents successfully.")
    generate_excel()
