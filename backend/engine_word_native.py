# backend/engine_word_native.py
"""
Native Microsoft Word Automation Engine via COM API (win32com.client).
Allows real-time typing, formatting, and table generation directly into Microsoft Word,
bypassing all WebView2 browser sandbox, loopback, or HTTPS restrictions.
"""

import os
import time

def get_or_create_word(visible=True):
    """Gets running Word instance or creates a new one, ensuring at least one document is open."""
    import win32com.client
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception:
        word = win32com.client.Dispatch("Word.Application")
    
    word.Visible = visible
    if word.Documents.Count == 0:
        word.Documents.Add()
    return word

def set_page_setup_native(doc):
    """Sets page setup: A4, Top 4cm, Left 4cm, Bottom 3cm, Right 3cm."""
    # Points conversion: 1 cm = 28.3465 pt
    cm_to_pt = 28.3465
    for section in doc.Sections:
        ps = section.PageSetup
        ps.PageWidth = 21.0 * cm_to_pt
        ps.PageHeight = 29.7 * cm_to_pt
        ps.TopMargin = 4.0 * cm_to_pt
        ps.LeftMargin = 4.0 * cm_to_pt
        ps.BottomMargin = 3.0 * cm_to_pt
        ps.RightMargin = 3.0 * cm_to_pt

def write_paragraphs_native(paragraphs, target_doc=None):
    """
    Writes structured paragraphs to the active Word document.
    paragraphs is a list of dicts: {"type": "heading1"|"heading2"|"body"|"body_center"|"bullet", "text": "..."}
    """
    word = get_or_create_word(visible=True)
    doc = target_doc if target_doc else word.ActiveDocument
    set_page_setup_native(doc)

    # wdAlignParagraph constants
    WD_ALIGN_LEFT = 0
    WD_ALIGN_CENTER = 1
    WD_ALIGN_RIGHT = 2
    WD_ALIGN_JUSTIFY = 3

    # Move selection to end of document
    selection = word.Selection
    selection.EndKey(Unit=6)  # 6 = wdStory (end of document)

    for item in paragraphs:
        p_type = item.get("type", "body")
        raw_text = item.get("text", "")
        lines = raw_text.split("\n")

        for line in lines:
            if not line.strip() and p_type != "blank":
                continue

            selection.EndKey(Unit=6)
            sel_range = selection.Range
            sel_range.Text = line.strip() + "\r"

            # Apply font formatting
            sel_range.Font.Name = "Times New Roman"
            sel_range.Font.Size = 12
            sel_range.Font.ColorIndex = 1  # wdBlack

            # Paragraph format
            pf = sel_range.ParagraphFormat
            pf.LineSpacingRule = 2  # wdLineSpaceDouble (2.0 spasi)
            pf.SpaceBefore = 0
            pf.SpaceAfter = 0

            if p_type in ["heading1", "title"]:
                pf.Alignment = WD_ALIGN_CENTER
                sel_range.Font.Bold = True
                pf.SpaceBefore = 12
                pf.SpaceAfter = 6
                pf.FirstLineIndent = 0
            elif p_type == "heading2":
                pf.Alignment = WD_ALIGN_LEFT
                sel_range.Font.Bold = True
                pf.SpaceBefore = 12
                pf.SpaceAfter = 4
                pf.FirstLineIndent = 0
            elif p_type == "body_center":
                pf.Alignment = WD_ALIGN_CENTER
                sel_range.Font.Bold = False
                pf.FirstLineIndent = 0
            elif p_type == "bullet":
                pf.Alignment = WD_ALIGN_LEFT
                sel_range.Font.Bold = False
                pf.LeftIndent = 36  # ~1.27 cm
                pf.FirstLineIndent = 0
            elif p_type == "blank":
                pf.SpaceBefore = 0
                pf.SpaceAfter = 0
                pf.FirstLineIndent = 0
            else:
                # Regular academic body
                pf.Alignment = WD_ALIGN_JUSTIFY
                sel_range.Font.Bold = False
                pf.FirstLineIndent = 36  # 1.27 cm / 0.5 inch indent

            selection.EndKey(Unit=6)

    # Bring Word to foreground
    try:
        word.Activate()
    except Exception:
        pass

    return True

def insert_table_native(headers, data, title=None):
    """Inserts an academic open-border table directly into the active Word document."""
    word = get_or_create_word(visible=True)
    doc = word.ActiveDocument

    selection = word.Selection
    selection.EndKey(Unit=6)

    if title:
        write_paragraphs_native([{"type": "heading2", "text": title}])

    num_rows = len(data) + 1
    num_cols = len(headers)

    table = doc.Tables.Add(Range=selection.Range, NumRows=num_rows, NumColumns=num_cols)
    table.Borders.Enable = False  # Start borderless

    # Constants
    WD_BORDER_TOP = -1
    WD_BORDER_BOTTOM = -3
    WD_LINE_STYLE_SINGLE = 1

    # Header Row
    header_row = table.Rows(1)
    for col_idx, h in enumerate(headers):
        cell = header_row.Cells(col_idx + 1)
        cell.Range.Text = str(h)
        cell.Range.Font.Name = "Times New Roman"
        cell.Range.Font.Size = 11
        cell.Range.Font.Bold = True
        cell.Range.ParagraphFormat.Alignment = 1  # Center
        cell.Range.ParagraphFormat.LineSpacingRule = 0  # Single

    # Apply Header Borders (Top & Bottom)
    header_row.Borders(WD_BORDER_TOP).LineStyle = WD_LINE_STYLE_SINGLE
    header_row.Borders(WD_BORDER_TOP).LineWidth = 8
    header_row.Borders(WD_BORDER_BOTTOM).LineStyle = WD_LINE_STYLE_SINGLE
    header_row.Borders(WD_BORDER_BOTTOM).LineWidth = 8

    # Data Rows
    for r_idx, row_vals in enumerate(data):
        row = table.Rows(r_idx + 2)
        for c_idx, val in enumerate(row_vals):
            cell = row.Cells(c_idx + 1)
            cell.Range.Text = str(val)
            cell.Range.Font.Name = "Times New Roman"
            cell.Range.Font.Size = 11
            cell.Range.Font.Bold = False
            cell.Range.ParagraphFormat.Alignment = 0 if c_idx == 0 else 1  # Left for 1st col, center for rest
            cell.Range.ParagraphFormat.LineSpacingRule = 0

    # Bottom border on last row
    last_row = table.Rows(num_rows)
    last_row.Borders(WD_BORDER_BOTTOM).LineStyle = WD_LINE_STYLE_SINGLE
    last_row.Borders(WD_BORDER_BOTTOM).LineWidth = 8

    # Table autofit
    table.AutoFitBehavior(2)  # wdAutoFitWindow

    # Add a blank paragraph after table
    selection.EndKey(Unit=6)
    selection.TypeParagraph()
    return True

def open_docx_file_in_word(filepath):
    """Opens a generated .docx file in Microsoft Word and brings Word to the foreground."""
    abs_path = os.path.abspath(filepath)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"File not found: {abs_path}")
    
    word = get_or_create_word(visible=True)
    doc = word.Documents.Open(abs_path)
    try:
        word.Activate()
    except Exception:
        pass
    return doc.Name
