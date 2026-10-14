# backend/build_full_thesis_docx.py
"""
Script Pembangun Naskah Tesis Lengkap dan Dokumen Per-Bab:
ANDRE JULIANSYA (NPM: 255441003) - Universitas Tridinanti Palembang 2026.
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

import build_andre_thesis

OUTPUT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

import re
from docx.enum.section import WD_SECTION_START

def format_proper_instansi(name):
    """
    Mengubah nama instansi yang ALL CAPS menjadi format baku EYD (Title Case)
    dengan tetap mempertahankan akronim dan singkatan resmi huruf kapital.
    Contoh: 'SEKRETARIAT DPRD KOTA PALEMBANG' -> 'Sekretariat DPRD Kota Palembang'
            'BPKAD KABUPATEN PALI' -> 'BPKAD Kabupaten PALI'
    """
    if not name:
        return "Instansi Pemerintah Daerah"
    UPPER_ACRONYMS = {
        "DPRD", "DPR", "BPKAD", "DPMPTSP", "RSUD", "PALI", "BAPPEDA", "BKPSDM", 
        "BAPENDA", "DINSOS", "DINKES", "DISDIK", "SDM", "ASN", "PNS", "TNI", 
        "POLRI", "BUMN", "BUMD", "SPM", "IKU", "SAKIP", "LAKIP", "OPD", "UPTD", 
        "BLUD", "PT", "CV", "FEB", "MM", "SPJ", "LPJ", "BOSP", "IGD", "RME", 
        "BPJS", "SNP", "SOP", "SUMSEL", "BPSDM", "KPP", "PLN", "UPDK", "BPS"
    }
    LOWER_WORDS = {"dan", "di", "ke", "dari", "pada", "untuk", "terhadap", "dengan", "yang", "sebagai"}
    
    words = name.strip().split()
    formatted = []
    for i, w in enumerate(words):
        w_clean = re.sub(r'[^a-zA-Z0-9]', '', w).upper()
        if w_clean in UPPER_ACRONYMS:
            prefix = re.match(r'^[^a-zA-Z0-9]+', w)
            suffix = re.search(r'[^a-zA-Z0-9]+$', w)
            pre_str = prefix.group(0) if prefix else ""
            suf_str = suffix.group(0) if suffix else ""
            formatted.append(f"{pre_str}{w_clean}{suf_str}")
        elif i > 0 and w.lower() in LOWER_WORDS:
            formatted.append(w.lower())
        else:
            formatted.append(w.capitalize())
    return " ".join(formatted)

def add_rich_runs_to_para(para, text, bold_default=False, italic_default=False, font_size=12):
    """
    Menambahkan run ke paragraf dengan mem-parse teks berhuruf miring (italic).
    Menghilangkan tanda asterisk (*) mentah atau tanda kurung (*...*)
    dan memastikan istilah asing dicetak miring (italic = True) di Microsoft Word.
    """
    if not text:
        return

    # Daftar istilah asing baku yang wajib dicetak miring
    AUTO_ITALIC_TERMS = [
        r"pseudo-compliance",
        r"evidence-based policy",
        r"research gap",
        r"intervening",
        r"quiet quitting",
        r"discretionary effort",
        r"employee engagement",
        r"performance gap",
        r"clinical pathway",
        r"emergency response time",
        r"patient safety",
        r"burnout",
        r"outer model",
        r"inner model",
        r"bootstrapping",
        r"composite reliability",
        r"average variance extracted",
        r"heterotrait-monotrait ratio",
        r"cross loadings",
        r"f-square",
        r"q-square",
        r"goodness of fit"
    ]

    # Bersihkan bentuk aneh seperti (*kata*) atau (*kata) menjadi (kata) yang akan diitalic
    normalized = text
    # Pola: (*term*) -> (*term*)
    # Pola: cari teks dengan *...* ATAU istilah asing
    token_pattern = re.compile(
        r'\*([^*]+)\*|' + '|'.join([r'\b(' + t + r')\b' for t in AUTO_ITALIC_TERMS]),
        re.IGNORECASE
    )

    last_idx = 0
    for match in token_pattern.finditer(normalized):
        start, end = match.span()
        # Teks sebelum match
        if start > last_idx:
            pre_text = normalized[last_idx:start]
            r_pre = para.add_run(pre_text)
            r_pre.bold = bold_default
            r_pre.italic = italic_default
            r_pre.font.name = "Times New Roman"
            r_pre.font.size = Pt(font_size)
            r_pre.font.color.rgb = RGBColor(0, 0, 0)

        # Teks match yang harus dicetak miring
        matched_str = match.group(1) if match.group(1) is not None else match.group(0)
        # Hilangkan sisa asterisk jika ada
        matched_str = matched_str.replace("*", "")

        r_match = para.add_run(matched_str)
        r_match.bold = bold_default
        r_match.italic = True
        r_match.font.name = "Times New Roman"
        r_match.font.size = Pt(font_size)
        r_match.font.color.rgb = RGBColor(0, 0, 0)

        last_idx = end

    # Sisa teks setelah match terakhir
    if last_idx < len(normalized):
        post_text = normalized[last_idx:]
        r_post = para.add_run(post_text)
        r_post.bold = bold_default
        r_post.italic = italic_default
        r_post.font.name = "Times New Roman"
        r_post.font.size = Pt(font_size)
        r_post.font.color.rgb = RGBColor(0, 0, 0)

def setup_page(doc, continuous=False):
    for s in doc.sections:
        s.page_width = Cm(21.0)      # A4
        s.page_height = Cm(29.7)
        s.top_margin = Cm(4.0)       # 4 cm
        s.left_margin = Cm(4.0)      # 4 cm
        s.bottom_margin = Cm(3.0)    # 3 cm
        s.right_margin = Cm(3.0)     # 3 cm
        if continuous:
            s.start_type = WD_SECTION_START.CONTINUOUS

def p(doc, text="", bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0.5, space_before=0, space_after=6, hanging=0, font_size=12, line_spacing=2.0):
    para = doc.add_paragraph()
    para.alignment = align
    para.paragraph_format.line_spacing = line_spacing
    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after = Pt(space_after)
    if hanging > 0:
        para.paragraph_format.left_indent = Inches(hanging)
        para.paragraph_format.first_line_indent = Inches(-hanging)
    elif indent > 0:
        para.paragraph_format.first_line_indent = Inches(indent)
    
    if text:
        add_rich_runs_to_para(para, text, bold_default=bold, italic_default=italic, font_size=font_size)
    return para

def p_numbered(doc, num_str, title_str="", text_str="", space_after=3, line_spacing=2.0):
    """
    Penomoran standar ilmiah (Hanging Indent) 2 spasi:
    - Angka (1., 2., dst.) di tepi kiri margin (tidak menjorok).
    - Teks baris pertama mulai di posisi 0.35 inci melalui tab stop.
    - Baris kedua dan seterusnya sejajar rata kiri dengan awal baris pertama (0.35 inci).
    """
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para.paragraph_format.line_spacing = line_spacing
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(space_after)
    
    indent_width = Inches(0.35)
    para.paragraph_format.left_indent = indent_width
    para.paragraph_format.first_line_indent = -indent_width
    para.paragraph_format.tab_stops.add_tab_stop(indent_width)

    # 1. Angka di margin kiri + TAB (\t) agar baris pertama melompat tepat ke posisi indentasi baris kedua
    r_num = para.add_run(f"{num_str.strip()}\t")
    r_num.font.name = "Times New Roman"
    r_num.font.size = Pt(12)
    r_num.bold = False

    # 2. Judul poin (jika ada)
    if title_str:
        clean_title = title_str.strip().replace("(*", "(").replace("*)", ")").replace("*", "")
        add_rich_runs_to_para(para, f"{clean_title}: ", bold_default=True, font_size=12)

    # 3. Isi uraian dengan dukungan kata asing miring
    if text_str:
        add_rich_runs_to_para(para, text_str.strip(), font_size=12)
    return para

def h1(doc, text):
    return p(doc, text, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_before=14, space_after=6, line_spacing=2.0)

def h2(doc, text):
    return p(doc, text, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_before=12, space_after=4, line_spacing=2.0)

def h3(doc, text):
    return p(doc, text, bold=True, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_before=8, space_after=2, line_spacing=2.0)

def apply_full_borders_table(table):
    """
    Tabel Standar Ilmiah:
    - Tampilkan SEMUA GARIS (semua batas: top, bottom, left, right, insideH, insideV)
    - Paskan ke jendela (100% margin lebar / 5000 pct) & paskan ke konten (autofit layout)
    - Isi tabel: Font Times New Roman 12 pt, 1 Spasi (Single Spacing)
    """
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    tblPr = table._tbl.tblPr

    # 1. Paskan ke jendela (100% lebar margin / 5000 pct)
    existing_tblW = tblPr.find(qn('w:tblW'))
    if existing_tblW is not None:
        tblPr.remove(existing_tblW)
    tblPr.append(parse_xml(r'<w:tblW %s w:w="5000" w:type="pct"/>' % nsdecls('w')))

    # 2. Paskan ke konten (autofit layout)
    existing_tblLayout = tblPr.find(qn('w:tblLayout'))
    if existing_tblLayout is not None:
        tblPr.remove(existing_tblLayout)
    tblPr.append(parse_xml(r'<w:tblLayout %s w:type="autofit"/>' % nsdecls('w')))

    # 3. Tampilkan semua garis (Semua Batas / All Borders)
    existing_tblBorders = tblPr.find(qn('w:tblBorders'))
    if existing_tblBorders is not None:
        tblPr.remove(existing_tblBorders)
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

    # 4. Standar Isi Tabel: Font Times New Roman 12 pt, 1 Spasi (Single Spacing)
    for row_idx, row in enumerate(table.rows):
        is_header = (row_idx == 0)
        trPr = row._tr.get_or_add_trPr()
        if is_header:
            trPr.append(parse_xml(r'<w:tblHeader %s/>' % nsdecls('w')))
        trPr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))

        for cell in row.cells:
            for para in cell.paragraphs:
                para.paragraph_format.line_spacing = 1.0
                para.paragraph_format.space_before = Pt(2)
                para.paragraph_format.space_after = Pt(2)
                for run in para.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(12)
                    run.bold = is_header

# Alias kompatibilitas
apply_open_table = apply_full_borders_table

def add_table_data(doc, title, headers, data, source="", alignments=None):
    if title:
        # Judul Tabel CENTER Rata Tengah, 1 Spasi (Bukan 2 Spasi)
        p(doc, title, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_before=12, space_after=4, line_spacing=1.0)
    table = doc.add_table(rows=len(data) + 1, cols=len(headers))
    header_row = table.rows[0]
    for c_idx, h in enumerate(headers):
        cell = header_row.cells[c_idx]
        cell.text = str(h)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r_idx, row_data in enumerate(data):
        row = table.rows[r_idx + 1]
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = str(val)
            if alignments and c_idx < len(alignments):
                cell.paragraphs[0].alignment = alignments[c_idx]
            else:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_full_borders_table(table)
    
    # Di bawah tabel: tulisan "Sumber: ..."
    # JANGAN HURUF MIRING (italic=False), 1 SPASI (line_spacing=1.0)
    if source:
        source_clean = source.strip()
        if not source_clean.lower().startswith("sumber"):
            source_clean = f"Sumber: {source_clean}"
        p(doc, source_clean, italic=False, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_before=3, space_after=8, font_size=10, line_spacing=1.0)
    else:
        p(doc, "", indent=0, space_after=4, line_spacing=1.0)

# ==================== BAGIAN AWAL ====================
def build_bagian_awal(doc, config=None):
    if config is None:
        config = {}
    judul = config.get("judul", "PENGARUH GAYA KEPEMIMPINAN DAN BUDAYA ORGANISASI TERHADAP DISIPLIN KERJA SERTA DAMPAKNYA PADA KINERJA PEGAWAI BANK PEMBANGUNAN DAERAH SUMATERA SELATAN DAN BANGKA BELITUNG CABANG PEMBANTU PARAMESWARA").strip().upper()
    penulis = config.get("penulis", "RUSMIANITA").strip().upper()
    nim = config.get("nim", "255441029").strip()
    konsentrasi = config.get("konsentrasi", "Manajemen Sumber Daya Manusia").strip()
    instansi = config.get("instansi", "Bank Pembangunan Daerah Sumatera Selatan dan Bangka Belitung Cabang Pembantu Parameswara").strip()
    tahun = config.get("tahun", "2026").strip()
    x1 = config.get("x1", "Gaya Kepemimpinan").strip()
    x2 = config.get("x2", "Budaya Organisasi").strip()
    m = config.get("m", "Disiplin Kerja").strip()
    y = config.get("y", "Kinerja Pegawai").strip()

    # Halaman Judul
    p(doc, judul, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=20, space_after=14, font_size=14)
    p(doc, "HALAMAN JUDUL", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    p(doc, "TESIS", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10, font_size=14)
    p(doc, "Ditulis untuk memenuhi Sebagian Persyaratan dalam menempuh Gelar Magister Manajemen", italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)
    p(doc, "Disusun Oleh :", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    p(doc, f"Nama : {penulis}", align=WD_ALIGN_PARAGRAPH.CENTER)
    p(doc, f"NPM : {nim}", align=WD_ALIGN_PARAGRAPH.CENTER)
    p(doc, f"Konsentrasi : {konsentrasi}", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)
    p(doc, f"PROGRAM STUDI MAGISTER MANAJEMEN\nFAKULTAS EKONOMI DAN BISNIS\nUNIVERSITAS TRIDINANTI\n{tahun}", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=24)
    doc.add_page_break()

    # Lembar Persetujuan Komisi Pembimbing
    h1(doc, "LEMBAR PERSETUJUAN KOMISI PEMBIMBING")
    p(doc, "Tesis ini Telah Disetujui Untuk Diujikan", italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    p(doc, judul, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    p(doc, f"Nama : {penulis}\nNPM : {nim}\nKonsentrasi : {konsentrasi}\nJudul Tesis : {judul.title()}", space_after=12)
    p(doc, "Menyetujui :\nPembimbing Tesis", bold=True, space_after=8)
    p(doc, "Tanggal ...................   Pembimbing I : ...................................................\nDr. Sari Sakarina, SE.,MM.,CHRM\nNIDN : 0214038501\n", space_after=8)
    p(doc, "Tanggal ...................   Pembimbing II : ...................................................\nDr. Yuni Adinda Putri, S.E., M.Si\nNIDN : 0207069401\n", space_after=12)
    p(doc, "Mengetahui :", bold=True, space_after=6)
    p(doc, "Dekan FEB                                                    Kaprodi MM\n\n\n\nDr. Msy. Mikial, SE., M.Si., Ak.,CA.,CSRS       Dr. Sari Sakarina, SE.,MM.,CHRM\nNIDN : 0205026401                                         NIDN : 0214038501\nTanggal ...........................                         Tanggal ...........................", space_after=12)
    doc.add_page_break()

    # Lembar Persetujuan Komisi Penguji
    h1(doc, "LEMBAR PERSETUJUAN KOMISI PENGUJI TESIS")
    p(doc, judul, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    p(doc, f"Dipertahankan di depan Komisi Penguji Tesis\nProgram Studi Magister Manajemen Fakultas Ekonomi dan Bisnis\nUniversitas Tridinanti", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    p(doc, f"Nama : {penulis}\nNPM : {nim}\nKonsentrasi : {konsentrasi}", space_after=8)
    p(doc, "Menyetujui :\nKetua Penguji : Tanggal Tanda Tangan\n....................... .......................\nNIDN : \n\nAnggota Penguji :\n1. Dr. Sari Sakarina, SE.,MM. CHRM ....................... .......................\n   NIDN : 0214038501\n2. Dr. Yuni Adinda Putri, S.E., M.Si ....................... .......................\n   NIDN : 0207069401", space_after=12)
    p(doc, f"Mengetahui :\nPalembang, {tahun}                                              Palembang, {tahun}\nDekan FEB                                                    Kaprodi MM\n\n\n\nDr. Msy. Mikial, SE., M.Si., Ak.,CA.,CSRS       Dr. Sari Sakarina, SE.,MM.,CHRM\nNIDN : 0205026401                                         NIDN : 0214038501", space_after=12)
    doc.add_page_break()

    # Kata Pengantar
    h1(doc, "KATA PENGANTAR")
    p(doc, f"Puji dan Syukur kepada Tuhan Yang Maha Esa, atas limpahan dan Hidayah-Nya sehingga Tesis yang Berjudul “{judul.title()}”.", indent=0.5, space_after=6)
    p(doc, "Tesis ini ditulis untuk memenuhi sebagian persyaratan mendapatkan gelar Magister Manajemen pada Universitas Tridinanti Palembang Program Studi Magister Manajemen. Penyelesaian Tesis ini melibatkan banyak pihak yang telah memberikan bantuan, baik langsung maupun tidak langsung, moril maupun materil. Oleh karena itu pada kesempatan ini penulis menyampaikan penghargaan dan rasa terima kasih kepada:", indent=0.5, space_after=6)
    points = [
        "1. Ketua Yayasan Pendidikan Nasional Tridinanti Palembang Bapak Prof.H. Mahmud Hasjim, MME.",
        "2. Rektor Universitas Tridinanti Palembang Bapak Prof. Dr. Ir. H. Edizal AE, MS.",
        "3. Dekan Fakultas Ekonomi & Bisnis Universitas Tridinanti Palembang Ibu Dr. Msy. Mikial, SE., M.Si., Ak.,CA.,CSRS.",
        "4. Wakil Dekan I Fakultas Ekonomi & Bisnis Universitas Tridinanti Palembang Ibu Dr. M. Ima Andriyani, SE., M.Si.",
        "5. Direktur Program Pascasarjana Magister Manajemen Universitas Tridinanti Palembang Bapak Dr. Djatmiko Noviantoro, S.E., M.Si.",
        "6. Kaprodi Program Pascasarjana Magister Manajemen Universitas Tridinanti Palembang Ibu Dr. Sari Sakarina, SE., MM., CHRM, sekaligus selaku Pembimbing I yang senantiasa meluangkan waktu, pikiran dan tenaganya untuk memberikan bimbingan dan arahan dalam menyelesaikan Tesis ini.",
        "7. Ibu Dr. Yuni Adinda Putri, S.E., M.Si, selaku Pembimbing II yang senantiasa meluangkan waktu, pikiran dan tenaganya untuk memberikan bimbingan dan arahan dalam menyelesaikan Tesis ini.",
        f"8. Pimpinan dan segenap pegawai pada {instansi} yang telah bersedia meluangkan waktu dan membantu proses pengumpulan data penelitian ini.",
        "9. Kedua Orang tua tercinta, bapak dan ibu atas doa dan kasih sayangnya yang tulus.",
        "10. Saudara, keluarga dan rekan-rekan yang senantiasa memberikan dukungan dan motivasi moril.",
        "11. Rekan-rekan mahasiswa Angkatan MM Universitas Tridinanti yang senantiasa saling menguatkan.",
        "12. Semua Pihak yang telah membimbing, membantu dan mendorong penyelesaian Tesis ini."
    ]
    for pt in points:
        p(doc, pt, indent=0.25, space_after=2)
    p(doc, f"Penulis menyadari mungkin dalam Tesis ini masih terdapat kekurangan, untuk itu diharapkan tanggapan dan masukan dari berbagai pihak sebagai bahan perbaikan dengan harapan Tesis ini dapat bermanfaat bagi pengembangan ilmu di bidang {konsentrasi}.", indent=0.5, space_before=6, space_after=12)
    p(doc, f"Palembang, {tahun}\n\n\n{penulis}", align=WD_ALIGN_PARAGRAPH.RIGHT)
    doc.add_page_break()

    # Abstrak Indonesia
    h1(doc, "ABSTRAK")
    p(doc, f"{penulis}, {judul.title()}, dibawah bimbingan Ibu Dr. Sari Sakarina, SE., MM., CHRM, dan Ibu Dr. Yuni Adinda Putri, S.E., M.Si.", bold=True, space_after=6)
    p(doc, f"Penelitian ini bertujuan untuk menganalisis dan membuktikan secara empiris pengaruh {x1} dan {x2} terhadap {m} serta dampaknya terhadap {y} pada {instansi}. Penelitian ini dilatarbelakangi oleh fenomena dinamika kesenjangan pencapaian target kinerja operasional dan tuntutan peningkatan kedisiplinan kerja pegawai. Pendekatan penelitian yang digunakan adalah kuantitatif asosiatif kausal. Metode pengumpulan data dilakukan melalui penyebaran kuesioner dengan skala Likert 1 sampai 5. Data dianalisis menggunakan metode Partial Least Squares - Structural Equation Modeling (PLS-SEM) dengan perangkat lunak SmartPLS.", indent=0.5, space_after=6)
    p(doc, f"Hasil evaluasi model pengukuran (Outer Model) dan model struktural (Inner Model) membuktikan bahwa seluruh indikator valid dan reliabel. Pengujian hipotesis melalui prosedur bootstrapping membuktikan bahwa: (1) {x1} berpengaruh positif dan signifikan terhadap {m}; (2) {x2} berpengaruh positif dan signifikan terhadap {m}; (3) {x1} berpengaruh positif dan signifikan terhadap {y}; (4) {x2} berpengaruh positif dan signifikan terhadap {y}; (5) {m} berpengaruh positif dan signifikan terhadap {y}; (6) {m} memediasi secara signifikan pengaruh {x1} terhadap {y}; dan (7) {m} memediasi secara signifikan pengaruh {x2} terhadap {y}. Implikasi manajerial menyarankan pimpinan instansi untuk memperkuat kepemimpinan suportif dan pembudayaan nilai organisasi dalam menegakkan kedisiplinan guna mencapai kinerja yang optimal.", indent=0.5, space_after=6)
    p(doc, f"Kata kunci : {x1}, {x2}, {m}, {y}", bold=True, space_before=4)
    doc.add_page_break()

    # Abstract English
    h1(doc, "ABSTRACT")
    p(doc, f"{penulis}, The Effect of {x1} and {x2} on {y} with {m} as Intervening Variable at {instansi}; supervised by Dr. Sari Sakarina, SE., MM., CHRM, and Dr. Yuni Adinda Putri, S.E., M.Si.", bold=True, space_after=6)
    p(doc, f"This study aims to examine and empirically prove the effect of {x1} and {x2} on {m} and its impact on {y} at {instansi}. This research was motivated by empirical phenomena regarding performance target gaps and the crucial importance of work discipline. A quantitative causal associative approach was implemented. Data collection was conducted using Likert-scale questionnaires (1-5). Data analysis was performed using Partial Least Squares - Structural Equation Modeling (PLS-SEM) via SmartPLS.", indent=0.5, space_after=6)
    p(doc, f"The results of measurement model and structural model evaluation confirm high construct reliability and convergent validity. Bootstrapping analysis results demonstrate that: (1) {x1} has a positive and significant effect on {m}; (2) {x2} has a positive and significant effect on {m}; (3) {x1} has a positive and significant effect on {y}; (4) {x2} has a positive and significant effect on {y}; (5) {m} has a positive and significant effect on {y}; (6) {m} significantly mediates the relationship between {x1} and {y}; and (7) {m} significantly mediates the relationship between {x2} and {y}. Managerial implications suggest enhancing supportive leadership and reinforcing organizational culture to foster work discipline and achieve peak organizational performance.", indent=0.5, space_after=6)
    p(doc, f"Keywords: {x1}, {x2}, {m}, {y}", bold=True, italic=True, space_before=4)
    doc.add_page_break()

    # Surat Pernyataan
    h1(doc, "SURAT PERNYATAAN")
    p(doc, f"Saya yang bertanda tangan di bawah ini:\nNama : {penulis}\nNPM : {nim}\nProgram studi : Magister Manajemen Universitas Tridinanti Palembang\nJudul tesis : {judul.title()}\n\nDengan ini menyatakan dengan sesungguhnya bahwa:\n1. Tesis yang saya susun sebagai salah satu syarat untuk memperoleh gelar Magister Manajemen (MM) dari Fakultas Ekonomi & Bisnis Program Studi Magister Manajemen Universitas Tridinanti Palembang seluruhnya adalah hasil karya sendiri.\n2. Apabila ada bagian-bagian tertentu dalam penulisan tesis yang saya kutip dari hasil karya orang lain, akan saya tulis sumbernya dengan jelas, sesuai norma, kaidah dan etika penulisan ilmiah.\n3. Jika dikemudian hari ditemukan seluruh atau sebagian tesis yang saya susun ini bukan hasil karya saya sendiri atau adanya plagiat dalam bagian-bagian tertentu, saya bersedia menerima sanksi pencabutan gelar Akademik Magister Manajemen (MM) yang saya sandang dan sanksi-sanksi lainnya sesuai dengan peraturan perundang-undangan yang berlaku.\n\nDemikianlah pernyataan ini saya buat dengan sesungguhnya dalam keadaan sadar dengan rasa tanggung jawab tanpa paksaan dari pihak manapun.", space_after=12)
    p(doc, f"Palembang, {tahun}\nYang menyatakan,\n\n(Materai Rp 10.000)\n\n{penulis}", align=WD_ALIGN_PARAGRAPH.RIGHT)
    doc.add_page_break()

    # Riwayat Hidup
    h1(doc, "DAFTAR RIWAYAT HIDUP")
    p(doc, "DATA PRIBADI", bold=True, space_after=2)
    p(doc, f"Nama: {penulis.title()}\nNPM: {nim}\nKonsentrasi: {konsentrasi}\nProgram Studi: Magister Manajemen\nFakultas: Ekonomi dan Bisnis\nUniversitas: Universitas Tridinanti", space_after=8)
    p(doc, "RIWAYAT PENDIDIKAN", bold=True, space_after=2)
    p(doc, "S-1: Fakultas Ekonomi dan Bisnis\nSMA: SMA Negeri Unggulan\nSMP: SMP Negeri\nSD: SD Negeri", space_after=12)
    doc.add_page_break()

    # Daftar Isi
    h1(doc, "DAFTAR ISI")
    toc_items = [
        ("HALAMAN JUDUL", "i"),
        ("LEMBAR PERSETUJUAN KOMISI PEMBIMBING", "ii"),
        ("LEMBAR PERSETUJUAN KOMISI PENGUJI TESIS", "iii"),
        ("KATA PENGANTAR", "iv"),
        ("ABSTRAK", "vi"),
        ("ABSTRACT", "vii"),
        ("SURAT PERNYATAAN", "viii"),
        ("DAFTAR RIWAYAT HIDUP", "ix"),
        ("DAFTAR ISI", "x"),
        ("DAFTAR TABEL", "xii"),
        ("DAFTAR GAMBAR", "xiii"),
        ("DAFTAR LAMPIRAN", "xiv"),
        ("BAB I PENDAHULUAN", "1"),
        ("  1.1. Latar Belakang Masalah", "1"),
        ("  1.2. Identifikasi Masalah", "6"),
        ("  1.3. Pembatasan Masalah", "7"),
        ("  1.4. Perumusan Masalah", "7"),
        ("  1.5. Tujuan Penelitian", "8"),
        ("  1.6. Kegunaan Penelitian", "9"),
        ("BAB II KAJIAN TEORITIS DAN HIPOTESIS PENELITIAN", "12"),
        ("  2.1. Kajian Teoritis", "12"),
        ("    2.1.1. Kinerja", "12"),
        ("    2.1.2. Employee Engagement", "18"),
        ("    2.1.3. Motivasi Kerja", "21"),
        ("    2.1.4. Disiplin Kerja", "24"),
        ("  2.2. Hasil Penelitian Yang Relevan", "27"),
        ("  2.3. Kerangka Berpikir", "31"),
        ("  2.4. Hipotesis Penelitian", "37"),
        ("BAB III METODE PENELITIAN", "38"),
        ("  3.1. Tempat dan Waktu Penelitian", "38"),
        ("  3.2. Populasi dan Sampel", "38"),
        ("    3.2.1 Populasi", "38"),
        ("    3.2.2. Sampel", "39"),
        ("  3.3. Variabel dan Definisi Operasional", "40"),
        ("  3.4 Instrumen Penelitian", "45"),
        ("  3.5. Teknik Analisis Data", "45"),
        ("  3.6. Pengujian Hipotesis", "50"),
        ("BAB IV HASIL ANALISIS DAN PEMBAHASAN", "52"),
        ("  4.1. Hasil Analisis", "52"),
        ("  4.2. Pembahasan Hasil", "85"),
        ("BAB V KESIMPULAN, IMPLIKASI DAN SARAN", "100"),
        ("  5.1. Kesimpulan", "100"),
        ("  5.2. Implikasi Kebijakan", "101"),
        ("  5.3. Saran", "104"),
        ("DAFTAR PUSTAKA", "107"),
        ("LAMPIRAN", "111"),
    ]
    for title, page in toc_items:
        p(doc, f"{title} {' .' * 30} {page}", space_after=1, font_size=11)
    doc.add_page_break()

    # Daftar Tabel
    h1(doc, "DAFTAR TABEL")
    tables_list = [
        ("Tabel 1. 1 Capaian IKU Persentase Penerbitan Izin Tepat Waktu DPMPTSP Provinsi Sumatera Selatan", "2"),
        ("Tabel 1. 2 Rekapitulasi Absensi dan Pelanggaran Jam Kerja Pegawai DPMPTSP Provinsi Sumatera Selatan Tahun 2025", "4"),
        ("Tabel 2. 1 Hasil Penelitian Yang Relevan", "27"),
        ("Tabel 3. 1 Jadwal Penelitian", "38"),
        ("Tabel 3. 2 Kisi-Kisi Instrumen Variabel Kinerja", "41"),
        ("Tabel 3. 3 Kisi-Kisi Instrumen Variabel Employee Engagement", "42"),
        ("Tabel 3. 4 Kisi-Kisi Instrumen Variabel Motivasi Kerja", "43"),
        ("Tabel 3. 5 Kisi-kisi Instrumen Variabel Disiplin Kerja", "44"),
        ("Tabel 3. 6 Scoring Untuk Jawaban Kuesioner", "45"),
        ("Tabel 3. 7 Kriteria Keputusan Hasil Uji t-statistic", "51"),
        ("Tabel 4. 1 Deskripsi Karakteristik Responden berdasarkan Jenis Kelamin", "52"),
        ("Tabel 4. 2 Deskripsi Karakteristik Responden berdasarkan Usia", "54"),
        ("Tabel 4. 3 Deskripsi Karakteristik Responden berdasarkan Pendidikan", "55"),
        ("Tabel 4. 4 Kategori Jawaban", "57"),
        ("Tabel 4. 5 Kategori Jawaban Responden", "58"),
        ("Tabel 4. 6 Outer Loading Pls Algorithm", "62"),
        ("Tabel 4. 7 Average Variance Extracted (AVE)", "64"),
        ("Tabel 4. 8 Akar Kuadrat AVE", "66"),
        ("Tabel 4. 9 Validitas Diskriminan", "66"),
        ("Tabel 4. 10 Nilai Cross Loading", "68"),
        ("Tabel 4. 11 Hasil Uji Reliabilitas", "71"),
        ("Tabel 4. 12 Nilai R-Square (R2)", "73"),
        ("Tabel 4. 13 Hasil F2", "75"),
        ("Tabel 4. 14 Pengaruh Langsung", "77"),
        ("Tabel 4. 15 Pengaruh Tidak Langsung", "83"),
    ]
    for title, page in tables_list:
        p(doc, f"{title} {' .' * 25} {page}", space_after=1, font_size=11)
    doc.add_page_break()

    # Daftar Gambar & Daftar Lampiran
    h1(doc, "DAFTAR GAMBAR")
    p(doc, "Gambar 2. 1 Kerangka Berpikir .......................................................................... 36", space_after=1, font_size=11)
    p(doc, "Gambar 4. 1 Full Model Setelah Dikalkulasikan .................................................. 61", space_after=1, font_size=11)
    p(doc, "Gambar 4. 2 Hasil uji T-Statistik Antar Variabel .................................................. 80", space_after=1, font_size=11)
    p(doc, "", space_after=8)

    h1(doc, "DAFTAR LAMPIRAN")
    lamp_list = [
        ("Lampiran 1 Kuesioner", "111"),
        ("Lampiran 2 Hasil Deskriptif Jawaban Responden", "115"),
        ("Lampiran 3 Hasil Olah Data Menggunakan SmartPLS", "117"),
        ("Lampiran 4 Tabulasi Jawaban Responden Variabel Motivasi Kerja", "123"),
        ("Lampiran 5 Tabulasi Jawaban Responden Variabel Disiplin Kerja", "125"),
        ("Lampiran 6 Tabulasi Jawaban Responden Variabel Employee Engagement", "127"),
        ("Lampiran 7 Tabulasi Jawaban Responden Variabel Kinerja Pegawai", "129"),
        ("Lampiran 8 Hasil Similarity Turnitin", "131"),
    ]
    for title, page in lamp_list:
        p(doc, f"{title} {' .' * 25} {page}", space_after=1, font_size=11)
    doc.add_page_break()

print("Bagian awal siap.")
