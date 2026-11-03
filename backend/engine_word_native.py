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
    import pythoncom

    word = None
    # 1. Coba GetActiveObject standar
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception:
        # Jika gen_py corrupt (CLSIDToPackageMap/CLSIDToClassMap error), fallback ke dynamic IDispatch
        try:
            disp = pythoncom.GetActiveObject("Word.Application")
            idisp = disp.QueryInterface(pythoncom.IID_IDispatch)
            word = win32com.client.dynamic.Dispatch(idisp)
        except Exception:
            pass

    # 2. Jika Word belum aktif, Dispatch / dynamic Dispatch
    if word is None:
        try:
            word = win32com.client.Dispatch("Word.Application")
        except Exception:
            try:
                word = win32com.client.dynamic.Dispatch("Word.Application")
            except Exception:
                # Bersihkan gen_py cache jika corrupt
                import shutil
                gen_path = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp", "gen_py")
                if os.path.exists(gen_path):
                    shutil.rmtree(gen_path, ignore_errors=True)
                word = win32com.client.Dispatch("Word.Application")
    
    try:
        word.Visible = visible
    except Exception:
        pass
    try:
        if word.Documents.Count == 0:
            word.Documents.Add()
    except Exception:
        pass
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


# ---- Konstanta Word yang dipakai penomoran halaman (tanpa perlu typelib) ----
WD_SECTION_BREAK_NEXT_PAGE = 2      # wdSectionBreakNextPage
WD_COLLAPSE_START = 1               # wdCollapseStart
WD_FOOTER_PRIMARY = 1               # wdHeaderFooterPrimary
WD_FOOTER_FIRST_PAGE = 2            # wdHeaderFooterFirstPage
WD_ALIGN_PAGE_CENTER = 1            # wdAlignPageNumberCenter
WD_NUM_ARABIC = 0                   # wdPageNumberStyleArabic
WD_NUM_LOWER_ROMAN = 2              # wdPageNumberStyleLowercaseRoman (i, ii, iii) — bukan 4 (huruf a, b, c)


def ensure_page_numbering(doc=None, verbose=True):
    """Atur NOMOR HALAMAN otomatis: bagian depan = angka romawi kecil (i, ii, ...),
    naskah (BAB I dst.) = angka Arab mulai 1, dengan PEMBATAS BAB (section break)
    otomatis di antara keduanya.

    - Idempoten: aman dipanggil setiap kali sub-bab disisipkan.
    - Pembatas hanya dibuat bila bagian depan (front matter) memang ada dan berada
      sebelum BAB I; page-break/paragraf kosong di batas itu dibersihkan lebih dulu
      supaya tidak muncul halaman kosong.
    - Halaman judul tidak diberi nomor (tetap dihitung sebagai i).

    Mengembalikan dict ringkas hasil kerja.
    """
    if doc is None:
        word = get_or_create_word(visible=True)
        doc = word.ActiveDocument

    hasil = {"dibuat_pembatas": False, "section_depan": None, "section_naskah": None,
             "jumlah_section": 0, "sukses": False}

    def _scan():
        """Satu lintasan: temukan INDEKS paragraf judul cover & judul BAB I sebenarnya.

        Entri DAFTAR ISI pun berbunyi "HALAMAN JUDUL"/"BAB I", sehingga dibedakan
        lewat TINGKAT OUTLINE (judul asli = tingkat 1 karena ditandai `outlineLvl`;
        entri daftar isi = teks biasa). Bila tidak ada yang bertingkat, dipakai
        kecocokan teks sebagai cadangan (judul BAB I = kecocokan terakhir).

        Catatan performa: paragraf ditelusuri berurutan lewat Paragraph.Next()
        (bukan Paragraphs(i), yang biayanya naik kuadratik seiring jumlah paragraf)
        dan lintasan BERHENTI begitu kedua judul bertingkat 1 sudah ditemukan.
        Hasilnya sama persis dengan pemindaian penuh, tetapi biasanya hanya
        menyentuh belasan sampai beberapa ratus paragraf pertama.
        """
        n = int(doc.Paragraphs.Count)
        i_judul = None
        i_bab1 = None
        i_bab1_teks = None

        def _baris1(p):
            try:
                t = p.Range.Text or ""
            except Exception:
                return ""
            return (t.replace("\v", "\n").split("\n")[0] or "").strip().upper()

        try:
            p = doc.Paragraphs(1)
        except Exception:
            p = None
        i = 1
        while p is not None and i <= n:
            try:
                lvl = int(p.OutlineLevel)
            except Exception:
                lvl = 10
            if lvl == 1:
                baris1 = _baris1(p)
                if baris1 == "HALAMAN JUDUL":
                    if i_judul is None:
                        i_judul = i
                elif baris1 == "BAB I":
                    if i_bab1 is None:
                        i_bab1 = i
            if i_judul is not None and i_bab1 is not None:
                break
            try:
                p = p.Next()
            except Exception:
                p = None
            i += 1

        if i_bab1 is None:
            # Cadangan (perilaku lama): kemunculan teks "BAB I" TERAKHIR, tanpa
            # memperhatikan tingkat outline. Hanya dijalankan bila benar-benar perlu.
            try:
                p = doc.Paragraphs(1)
            except Exception:
                p = None
            i = 1
            while p is not None and i <= n:
                if _baris1(p) == "BAB I":
                    i_bab1_teks = i
                try:
                    p = p.Next()
                except Exception:
                    p = None
                i += 1
            i_bab1 = i_bab1_teks
        return i_judul, i_bab1

    i_judul, i_bab1 = _scan()

    # ---------- 1) Bersihkan paragraf kosong/page-break tepat sebelum BAB I ----------
    if i_bab1 is not None:
        for _ in range(4):
            if i_bab1 - 1 < 1:
                break
            try:
                p_prev = doc.Paragraphs(i_bab1 - 1)
                teks_prev = (p_prev.Range.Text or "").strip()
            except Exception:
                break
            if teks_prev:
                break
            try:
                p_prev.Range.Delete()
                i_bab1 -= 1
            except Exception:
                break

    # ---------- 2) Pembatas bab (section break) antara bagian depan & naskah ----------
    if (i_judul is not None and i_bab1 is not None and i_judul < i_bab1):
        try:
            sec_judul = int(doc.Paragraphs(i_judul).Range.Sections(1).Index)
        except Exception:
            sec_judul = 1
        try:
            sec_bab = int(doc.Paragraphs(i_bab1).Range.Sections(1).Index)
        except Exception:
            sec_bab = sec_judul
        if sec_bab == sec_judul:
            try:
                rng = doc.Paragraphs(i_bab1).Range
                rng.Collapse(WD_COLLAPSE_START)
                rng.InsertBreak(WD_SECTION_BREAK_NEXT_PAGE)
                hasil["dibuat_pembatas"] = True
                i_bab1 += 1          # paragraf pembatas menyisip di depan judul BAB I
            except Exception as ex:
                if verbose:
                    print(f"[halaman] gagal membuat pembatas bab: {ex}")
        # Judul bab sudah mulai halaman baru lewat pembatas -> matikan page-break manual
        try:
            doc.Paragraphs(i_bab1).Range.ParagraphFormat.PageBreakBefore = False
        except Exception:
            pass

    # ---------- 3) Tentukan batas section depan vs naskah ----------
    n_sec = int(doc.Sections.Count)
    hasil["jumlah_section"] = n_sec
    idx_depan = None
    if i_judul is not None:
        try:
            idx_depan = int(doc.Paragraphs(i_judul).Range.Sections(1).Index)
        except Exception:
            idx_depan = 1
    if hasil["dibuat_pembatas"] and idx_depan is not None:
        idx_naskah = idx_depan + 1
    elif i_bab1 is not None:
        try:
            idx_naskah = int(doc.Paragraphs(i_bab1).Range.Sections(1).Index)
        except Exception:
            idx_naskah = None
        if idx_naskah is not None and idx_depan is not None and idx_naskah <= idx_depan:
            idx_naskah = idx_depan + 1
    else:
        idx_naskah = None
    hasil["section_depan"] = idx_depan
    hasil["section_naskah"] = idx_naskah

    # ---------- 4) Pasang nomor halaman di tengah footer tiap section ----------
    for k in range(1, n_sec + 1):
        sec = doc.Sections(k)
        ps = sec.PageSetup
        ftr = sec.Footers(WD_FOOTER_PRIMARY)
        try:
            ftr.LinkToPrevious = False
        except Exception:
            pass
        try:
            while int(ftr.PageNumbers.Count) > 0:
                ftr.PageNumbers(1).Delete()
        except Exception:
            pass
        try:
            if (ftr.Range.Text or "").strip():
                ftr.Range.Delete()
        except Exception:
            pass
        try:
            ftr.Range.ParagraphFormat.Alignment = WD_ALIGN_PAGE_CENTER
        except Exception:
            pass
        # "depan" = section bagian depan (front matter). Bila tidak ada front matter
        # (HALAMAN JUDUL tidak ditemukan), seluruh section memakai angka Arab.
        depan = (idx_depan is not None) and ((idx_naskah is None) or (k < idx_naskah))
        awal = (k == (idx_depan if depan else (idx_naskah if idx_naskah is not None else 1)))
        try:
            # Gaya & awal penomoran diatur pada KOLEKSI "PageNumbers" (bukan objek
            # PageNumber tunggal) — sesuai object model Word.
            pns = ftr.PageNumbers
            pns.Add(WD_ALIGN_PAGE_CENTER, True)
            pns.NumberStyle = WD_NUM_LOWER_ROMAN if depan else WD_NUM_ARABIC
            try:
                # Penomoran hanya dimulai ulang pada section pertama bagian depan dan
                # pada section pertama naskah; section berikutnya MELANJUTKAN nomor.
                pns.RestartNumberingAtSection = bool(awal)
                if awal:
                    pns.StartingNumber = 1
            except Exception:
                pass
            # Halaman judul tidak diberi nomor (tetap dihitung sebagai i)
            ps.DifferentFirstPageHeaderFooter = bool(depan and awal)
            if depan and awal:
                try:
                    sec.Footers(WD_FOOTER_FIRST_PAGE).Range.Delete()
                except Exception:
                    pass
        except Exception as ex:
            if verbose:
                print(f"[halaman] gagal memasang nomor halaman section {k}: {ex}")
            continue
        try:
            ftr.Range.Font.Name = "Times New Roman"
            ftr.Range.Font.Size = 12
        except Exception:
            pass

    hasil["sukses"] = True
    if verbose:
        print(f"[halaman] jumlah section={n_sec} | bagian depan={hasil['section_depan']} "
              f"(romawi) | naskah={hasil['section_naskah']} (Arab mulai 1) | pembatas dibuat={hasil['dibuat_pembatas']}")
    return hasil

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
