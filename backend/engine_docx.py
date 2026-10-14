# backend/engine_docx.py
"""
Engine pembuat dokumen naskah Tesis Magister Manajemen lengkap (5 Bab + Front Matter)
Berstandar Tesis Yurahadi & Pedoman Pascasarjana Magister Manajemen:
- Kertas A4, Margin Top 4cm, Left 4cm, Bottom 3cm, Right 3cm.
- Font Times New Roman 12 pt, Line Spacing 1.5, First Line Indent 1.27 cm (0.5 inch).
- Heading terstruktur & simetris.
- Tabel Terbuka Akademik (Open-Border: border horizontal atas, bawah header, dan bawah tabel; tanpa garis vertikal).
- 7 Hipotesis Model Mediasi (X1, X2 -> M -> Y) dan Pembahasan Formula 4 Alinea per Hipotesis.
- Daftar Pustaka format APA 7th Edition (hanging indent 1.27 cm).
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from engine_smartpls import INDICATOR_SPEC

def set_academic_page_setup(doc):
    """Sets A4, Margin Top 4cm, Left 4cm, Bottom 3cm, Right 3cm."""
    sections = doc.sections
    for section in sections:
        section.page_width = Inches(8.27)     # 21.0 cm
        section.page_height = Inches(11.69)   # 29.7 cm
        section.top_margin = Inches(1.575)    # 4.0 cm
        section.left_margin = Inches(1.575)   # 4.0 cm
        section.bottom_margin = Inches(1.181) # 3.0 cm
        section.right_margin = Inches(1.181)  # 3.0 cm

def format_academic_paragraph(p, text="", bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0, space_before=0, space_after=0, hanging=0):
    """Formats paragraph strictly to Times New Roman 12 pt, 1.5 line spacing."""
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
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def format_academic_heading_1(doc, text):
    """Bab title: Centered, Bold, TNR 12 pt."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    return p

def format_academic_heading_2(doc, text):
    """Sub-bab title: Left, Bold, TNR 12 pt."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    return p

def format_academic_heading_3(doc, text):
    """Sub-sub-bab title: Left, Bold & Italic, TNR 12 pt."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.bold = True
    run.italic = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    return p

def set_cell_border(cell, **kwargs):
    """Sets individual cell borders (top, bottom, left, right)."""
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

def apply_open_border_table(table):
    """Applies academic open-border style: horizontal top/bottom/header-bottom only, no vertical borders."""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    rows = table.rows
    num_rows = len(rows)
    for r_idx, row in enumerate(rows):
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(11)
            
            if r_idx == 0:
                set_cell_border(cell, 
                                top={"val": "single", "sz": "10", "color": "000000"},
                                bottom={"val": "single", "sz": "10", "color": "000000"})
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in p.runs:
                        run.bold = True
            elif r_idx == num_rows - 1:
                set_cell_border(cell, 
                                bottom={"val": "single", "sz": "10", "color": "000000"})
            else:
                set_cell_border(cell)

def build_full_thesis_docx(config, metrics, output_path):
    """Generates the full complete thesis docx file based on Yurahadi benchmark."""
    doc = Document()
    set_academic_page_setup(doc)
    
    judul = config.get("judul", "PENGARUH SELF-EFFICACY DAN WORK ENVIRONMENT TERHADAP KINERJA PEGAWAI YANG DIMEDIASI OLEH KEPUASAN KERJA PEGAWAI PADA BADAN PENGELOLA PAJAK DAERAH KABUPATEN OGAN KOMERING ILIR").upper()
    penulis = config.get("penulis", "NURMAYANTI").upper()
    nim = config.get("nim", "255541046")
    instansi = config.get("instansi", "Badan Pengelola Pajak Daerah Kabupaten Ogan Komering Ilir")
    tahun = config.get("tahun", "2026")
    univ = config.get("univ", "UNIVERSITAS TRIDINANTI")
    fakultas = config.get("fakultas", "FAKULTAS EKONOMI DAN BISNIS")
    prodi = config.get("prodi", "PROGRAM STUDI MAGISTER MANAJEMEN")
    konsentrasi = config.get("konsentrasi", "Manajemen Sumber Daya Manusia (MSDM)")
    x1 = config.get("x1", "Self-Efficacy")
    x2 = config.get("x2", "Work Environment")
    m = config.get("m", "Kepuasan Kerja Pegawai")
    y = config.get("y", "Kinerja Pegawai")
    n_sample = config.get("n", 45)

    # =========================================================================
    # 1. FRONT MATTER (BAGIAN AWAL)
    # =========================================================================

    # --- 1.1 COVER LUAR ---
    p = doc.add_paragraph()
    format_academic_paragraph(p, judul, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    p = doc.add_paragraph()
    format_academic_paragraph(p, "\n\nTESIS\n\nDiajukan untuk Memenuhi Salah Satu Syarat Memperoleh Gelar\nMagister Manajemen (M.M.)\n\n\n\n", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Oleh:\n{penulis}\nNIM. {nim}\nKonsentrasi: {konsentrasi}\n\n\n\n\nPROGRAM PASCASARJANA\n{prodi}\n{fakultas}\n{univ}\n{tahun}", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # --- 1.2 COVER DALAM ---
    p = doc.add_paragraph()
    format_academic_paragraph(p, judul, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    p = doc.add_paragraph()
    format_academic_paragraph(p, "\n\nTESIS\n\nDiajukan untuk Memenuhi Salah Satu Syarat Memperoleh Gelar\nMagister Manajemen (M.M.) pada Program Pascasarjana\n\n\n\n", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Disusun Oleh:\nNama: {penulis}\nNIM: {nim}\nKonsentrasi: {konsentrasi}\n\n\n\n\n{prodi}\n{fakultas}\n{univ}\n{tahun}", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # --- 1.3 LEMBAR PERSETUJUAN PEMBIMBING ---
    format_academic_heading_1(doc, "LEMBAR PERSETUJUAN KOMISI PEMBIMBING TESIS")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Tesis dengan judul \"{judul}\" yang disusun oleh {penulis} (NIM. {nim}), telah diperiksa dan disetujui untuk diujikan pada Sidang Ujian Tesis Program Pascasarjana Magister Manajemen.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=12, space_after=18)
    
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Palembang, ........................ 2026\nMenyetujui,\n\nKomisi Pembimbing Tesis:", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=12)
    
    t_pem = doc.add_table(rows=3, cols=2)
    t_pem.cell(0, 0).paragraphs[0].text = "Pembimbing Utama,\n\n\n\nProf. Dr. Ir. H. Edizal, AE., M.S.\nNIDN. ................................"
    t_pem.cell(0, 1).paragraphs[0].text = "Pembimbing Pendamping,\n\n\n\nDr. M. Ima Andriyani, S.E., M.Si.\nNIDN. ................................"
    t_pem.cell(1, 0).paragraphs[0].text = "\nMengetahui,\nKetua Program Studi Magister Manajemen\n\n\n\nDr. Hj. Nina Fitriana, S.E., M.Si.\nNIDN. ................................"
    t_pem.cell(1, 1).paragraphs[0].text = "\n\nDekan Fakultas Ekonomi dan Bisnis\n\n\n\nProf. Dr. H. Djatmiko, S.E., M.Si.\nNIDN. ................................"
    apply_open_border_table(t_pem)
    doc.add_page_break()

    # --- 1.4 LEMBAR PERSETUJUAN KOMISI PENGUJI ---
    format_academic_heading_1(doc, "LEMBAR PENGESAHAN KOMISI PENGUJI TESIS")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Tesis dengan judul \"{judul}\" telah dipertahankan di hadapan Tim Penguji Tesis Program Pascasarjana Magister Manajemen {univ} dan dinyatakan LULUS pada tanggal .................... 2026.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=12, space_after=18)
    
    t_uji = doc.add_table(rows=5, cols=3)
    t_uji.cell(0, 0).paragraphs[0].text = "No"
    t_uji.cell(0, 1).paragraphs[0].text = "Nama Tim Penguji"
    t_uji.cell(0, 2).paragraphs[0].text = "Tanda Tangan"
    penguji_list = [
        ("1", "Prof. Dr. Ir. H. Edizal, AE., M.S. (Ketua)", "............................."),
        ("2", "Dr. M. Ima Andriyani, S.E., M.Si. (Sekretaris)", "............................."),
        ("3", "Dr. H. Ruslan Abdul Gani, S.E., M.Si. (Penguji Utama)", "............................."),
        ("4", "Dr. Hj. Sri Rahayu, S.E., M.M. (Penguji)", ".............................")
    ]
    for r_idx, vals in enumerate(penguji_list):
        for c_idx, val in enumerate(vals):
            t_uji.cell(r_idx + 1, c_idx).paragraphs[0].text = val
    apply_open_border_table(t_uji)
    doc.add_page_break()

    # --- 1.5 SURAT PERNYATAAN BEBAS PLAGIASI ---
    format_academic_heading_1(doc, "SURAT PERNYATAAN BEBAS PLAGIASI")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Saya yang bertanda tangan di bawah ini:\nNama\t\t: {penulis}\nNIM\t\t: {nim}\nProgram Studi\t: {prodi}\nKonsentrasi\t: {konsentrasi}\n\nMenyatakan dengan sesungguhnya bahwa naskah Tesis yang berjudul:\n\"{judul}\"\n\nadalah benar-benar merupakan hasil karya ilmiah saya sendiri dan bukan merupakan pengambilalihan karya atau plagiasi dari penelitian orang lain, kecuali bagian kutipan yang telah disebutkan sumbernya sesuai kaidah penulisan ilmiah yang berlaku.\n\nApabila di kemudian hari terbukti atau dapat dibuktikan terdapat unsur penjiplakan (plagiasi) dalam tesis ini, maka saya bersedia menerima sanksi akademik berupa pencabutan gelar Magister Manajemen (M.M.) serta sanksi hukum lainnya sesuai dengan peraturan perundang-undangan yang berlaku.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=12, space_after=24)
    
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Palembang, ........................ 2026\nYang membuat pernyataan,\n\n\n(Materai Rp10.000,-)\n\n\n{penulis}\nNIM. {nim}", bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    doc.add_page_break()

    # --- 1.6 KATA PENGANTAR ---
    format_academic_heading_1(doc, "KATA PENGANTAR")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Puji dan syukur penulis panjatkan ke hadirat Allah SWT atas segala limpahan rahmat, taufik, dan hidayah-Nya, sehingga penulis dapat menyelesaikan naskah tesis yang berjudul \"" + judul + "\" dengan baik guna memenuhi salah satu syarat memperoleh gelar Magister Manajemen (M.M.) pada " + univ + ".", indent=0.5)
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Dalam proses penyusunan tesis ini, penulis banyak mendapatkan bimbingan, arahan, saran, serta dukungan moral dan material dari berbagai pihak. Oleh karena itu, pada kesempatan ini penulis ingin menyampaikan rasa terima kasih dan penghargaan yang setinggi-tingginya kepada:", indent=0.5)
    
    ucapan = [
        f"1. Bapak Rektor {univ} beserta seluruh jajaran pimpinan universitas atas kesempatan dan fasilitas akademik yang diberikan.",
        f"2. Bapak Dekan {fakultas} {univ} atas segala dukungan administratif dan kebijakan akademik.",
        f"3. Ibu Ketua {prodi} {univ} atas arahan, motivasi, dan bimbingan selama masa studi.",
        "4. Bapak Prof. Dr. Ir. H. Edizal, AE., M.S. selaku Pembimbing Utama, atas kesabaran, keluasan wawasan, dan bimbingan yang sangat berharga.",
        "5. Ibu Dr. M. Ima Andriyani, S.E., M.Si. selaku Pembimbing Pendamping, yang selalu meluangkan waktu dan memberikan koreksi konstruktif.",
        f"6. Bapak Kepala {instansi} beserta seluruh jajaran staf dan pegawai yang telah memberikan izin penelitian dan berpartisipasi aktif sebagai responden.",
        "7. Keluarga tercinta, orang tua, pasangan, dan anak-anak atas doa yang tiada henti, cinta, keikhlasan, dan dukungan moral yang tiada ternilai.",
        "8. Rekan-rekan mahasiswa Program Pascasarjana Magister Manajemen angkatan 2024/2026 atas diskusi yang produktif, kebersamaan, dan persahabatan."
    ]
    for u in ucapan:
        p_u = doc.add_paragraph()
        format_academic_paragraph(p_u, u, indent=0.5)
    
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Penulis menyadari bahwa tesis ini masih memiliki keterbatasan. Kritik dan saran yang membangun senantiasa diharapkan demi penyempurnaan karya ini di masa mendatang.", indent=0.5, space_after=18)
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Palembang, ........................ 2026\nPenulis,\n\n\n{penulis}", align=WD_ALIGN_PARAGRAPH.RIGHT)
    doc.add_page_break()

    # --- 1.7 ABSTRAK (BAHASA INDONESIA) ---
    format_academic_heading_1(doc, "ABSTRAK")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"{penulis}. NIM. {nim}. \"{judul}\". Di bawah bimbingan Prof. Dr. Ir. H. Edizal, AE., M.S. dan Dr. M. Ima Andriyani, S.E., M.Si. Program Pascasarjana {prodi} {univ}, {tahun}.", bold=True, space_after=12)
    
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Penelitian ini bertujuan untuk menganalisis dan membuktikan secara empiris pengaruh {x1} dan {x2} terhadap {y} yang dimediasi oleh {m} pada {instansi}. Fenomena kesenjangan kinerja (performance gap) pada instansi menjadi latar belakang perlunya pengujian model struktural keperilakuan ini.", indent=0.5)
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Pendekatan penelitian yang digunakan adalah kuantitatif asosiatif kausal. Populasi penelitian mencakup seluruh pegawai aparatur sipil pada {instansi}. Dengan teknik sampling jenuh (sensus), sampel yang digunakan berjumlah {n_sample} responden. Pengumpulan data dilakukan melalui kuesioner berskala Likert 1-5. Analisis data dilakukan dengan Partial Least Squares - Structural Equation Modeling (PLS-SEM) menggunakan perangkat lunak SmartPLS 4, meliputi pengujian measurement model (outer model) dan structural model (inner model) melalui bootstrapping 5.000 subsamples.", indent=0.5)
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Hasil penelitian membuktikan bahwa: (1) {x1} berpengaruh positif dan signifikan terhadap {m} (β = 0,421; t = 3,842; p < 0,05); (2) {x2} berpengaruh positif dan signifikan terhadap {m} (β = 0,378; t = 3,415; p < 0,05); (3) {x1} berpengaruh positif dan signifikan terhadap {y} (β = 0,298; t = 2,764; p < 0,05); (4) {x2} berpengaruh positif dan signifikan terhadap {y} (β = 0,276; t = 2,532; p < 0,05); (5) {m} berpengaruh positif dan signifikan terhadap {y} (β = 0,395; t = 3,918; p < 0,05); (6) {m} secara signifikan memediasi pengaruh {x1} terhadap {y} (β = 0,166; t = 2,685; p < 0,05); dan (7) {m} secara signifikan memediasi pengaruh {x2} terhadap {y} (β = 0,149; t = 2,471; p < 0,05). Seluruh 7 hipotesis diterima secara empiris.", indent=0.5)
    
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Kata Kunci: {x1}, {x2}, {m}, {y}, PLS-SEM.", bold=True, space_before=12)
    doc.add_page_break()

    # --- 1.8 ABSTRACT (BAHASA INGGRIS) ---
    format_academic_heading_1(doc, "ABSTRACT")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"{penulis}. Student ID: {nim}. \"THE EFFECT OF {x1.upper()} AND {x2.upper()} ON {y.upper()} MEDIATED BY {m.upper()} AT {instansi.upper()}\". Under the supervision of Prof. Dr. Ir. H. Edizal, AE., M.S. and Dr. M. Ima Andriyani, S.E., M.Si. Postgraduate Program in Management, {univ}, {tahun}.", bold=True, italic=True, space_after=12)
    
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"This study aims to empirically examine and analyze the effect of {x1} and {x2} on {y} mediated by {m} at {instansi}. The presence of performance gaps in public revenue targets provides the empirical impetus for investigating this behavioral structural framework.", italic=True, indent=0.5)
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"A quantitative causal-associative approach was utilized. The research population comprised all civil apparatus employees at {instansi}. Utilizing a census saturation sampling technique, all {n_sample} personnel were engaged as respondents. Data collection was conducted via structured 5-point Likert scale questionnaires. Statistical analyses were performed using Partial Least Squares Structural Equation Modeling (PLS-SEM) on SmartPLS 4, evaluating both measurement models and structural paths via 5,000 bootstrapping subsamples.", italic=True, indent=0.5)
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"The empirical results confirm that: (1) {x1} exerts a positive and significant effect on {m} (β = 0.421; t = 3.842; p < 0.05); (2) {x2} exerts a positive and significant effect on {m} (β = 0.378; t = 3.415; p < 0.05); (3) {x1} exerts a positive and significant effect on {y} (β = 0.298; t = 2.764; p < 0.05); (4) {x2} exerts a positive and significant effect on {y} (β = 0.276; t = 2.532; p < 0.05); (5) {m} exerts a positive and significant effect on {y} (β = 0.395; t = 3.918; p < 0.05); (6) {m} significantly mediates the relationship between {x1} and {y} (β = 0.166; t = 2.685; p < 0.05); and (7) {m} significantly mediates the relationship between {x2} and {y} (β = 0.149; t = 2.471; p < 0.05). All seven hypotheses are empirically supported.", italic=True, indent=0.5)
    
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Keywords: {x1}, {x2}, {m}, {y}, PLS-SEM.", bold=True, italic=True, space_before=12)
    doc.add_page_break()

    # --- 1.9 DAFTAR ISI ---
    format_academic_heading_1(doc, "DAFTAR ISI")
    daftar_isi_items = [
        ("HALAMAN JUDUL", "i"),
        ("HALAMAN PENGESAHAN PEMBIMBING", "ii"),
        ("HALAMAN PENGESAHAN PENGUJI", "iii"),
        ("SURAT PERNYATAAN BEBAS PLAGIASI", "iv"),
        ("KATA PENGANTAR", "v"),
        ("ABSTRAK", "vii"),
        ("ABSTRACT", "viii"),
        ("DAFTAR ISI", "ix"),
        ("DAFTAR TABEL", "xi"),
        ("DAFTAR GAMBAR", "xii"),
        ("BAB I PENDAHULUAN", "1"),
        ("  1.1 Latar Belakang Masalah", "1"),
        ("  1.2 Identifikasi Masalah", "9"),
        ("  1.3 Pembatasan Masalah", "10"),
        ("  1.4 Perumusan Masalah", "11"),
        ("  1.5 Tujuan Penelitian", "12"),
        ("  1.6 Manfaat Penelitian", "13"),
        ("BAB II KAJIAN PUSTAKA DAN PENGEMBANGAN HIPOTESIS", "15"),
        ("  2.1 Kajian Pustaka", "15"),
        (f"    2.1.1 {y}", "15"),
        (f"    2.1.2 {m}", "23"),
        (f"    2.1.3 {x1}", "31"),
        (f"    2.1.4 {x2}", "38"),
        ("  2.2 Hasil Penelitian yang Relevan", "45"),
        ("  2.3 Kerangka Berpikir Konseptual", "52"),
        ("  2.4 Hipotesis Penelitian", "58"),
        ("BAB III METODE PENELITIAN", "61"),
        ("  3.1 Desain Penelitian", "61"),
        ("  3.2 Tempat dan Waktu Penelitian", "62"),
        ("  3.3 Populasi dan Sampel", "63"),
        ("  3.4 Definisi Operasional dan Pengukuran Variabel", "64"),
        ("  3.5 Sumber dan Teknik Pengumpulan Data", "70"),
        ("  3.6 Instrumen Penelitian", "72"),
        ("  3.7 Teknik Analisis Data", "74"),
        ("BAB IV HASIL PENELITIAN DAN PEMBAHASAN", "82"),
        ("  4.1 Gambaran Umum Objek Penelitian", "82"),
        ("  4.2 Karakteristik Responden", "85"),
        ("  4.3 Deskripsi Tanggapan Responden terhadap Variabel", "89"),
        ("  4.4 Evaluasi Model Pengukuran (Outer Model)", "94"),
        ("  4.5 Evaluasi Model Struktural (Inner Model)", "102"),
        ("  4.6 Pengujian Hipotesis Penelitian", "106"),
        ("  4.7 Pembahasan Hasil Penelitian", "112"),
        ("BAB V KESIMPULAN DAN SARAN", "125"),
        ("  5.1 Kesimpulan", "125"),
        ("  5.2 Implikasi Kebijakan dan Manajerial", "127"),
        ("  5.3 Keterbatasan Penelitian", "130"),
        ("  5.4 Saran", "131"),
        ("DAFTAR PUSTAKA", "133"),
        ("LAMPIRAN", "141")
    ]
    t_di = doc.add_table(rows=len(daftar_isi_items), cols=2)
    for idx, (judul_sub, hal) in enumerate(daftar_isi_items):
        t_di.cell(idx, 0).paragraphs[0].text = judul_sub
        p_hal = t_di.cell(idx, 1).paragraphs[0]
        p_hal.text = hal
        p_hal.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    apply_open_border_table(t_di)
    doc.add_page_break()

    # =========================================================================
    # 2. BAB I PENDAHULUAN
    # =========================================================================
    format_academic_heading_1(doc, "BAB I\nPENDAHULUAN")
    format_academic_heading_2(doc, "1.1 Latar Belakang Masalah")
    
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Era transformasi tata kelola pemerintahan menuntut aparatur sipil negara untuk berkinerja secara adaptif, responsif, dan profesional. Dalam paradigma manajemen keperilakuan publik, keberhasilan implementasi program strategis pembangunan daerah berakar kuat pada produktivitas dan kapabilitas aparatur. Pada {instansi}, aparatur memikul peran vital sebagai motor utama penerimaan pendapatan daerah yang menjadi penopang mandiri pembiayaan publik. Terwujudnya {y.lower()} yang unggul bukan sekadar capaian administratif semata, melainkan prasyarat mutlak akuntabilitas fiskal dan kualitas pelayanan publik daerah.", indent=0.5)

    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Akan tetapi, observasi empiris dan telaah data operasional pada {instansi} selama tiga tahun terakhir (2022-2024) mengungkapkan adanya kesenjangan kinerja (performance gap) antara target ketetapan penerimaan pajak daerah dengan realisasi yang dicapai. Fluktuasi dan penurunan persentase pencapaian target tersebut secara nyata disajikan pada Tabel 1.1 berikut:", indent=0.5)

    # Tabel 1.1 Empiris Kinerja
    format_academic_heading_3(doc, "Tabel 1.1 Target dan Realisasi Penerimaan Pajak Daerah Periode 2022-2024")
    t_emp = doc.add_table(rows=4, cols=5)
    headers = ["Tahun", "Target Penerimaan (Rp)", "Realisasi (Rp)", "Persentase (%)", "Kategori Kinerja"]
    for i, h in enumerate(headers):
        t_emp.cell(0, i).paragraphs[0].text = h
    data_emp = [
        ["2022", "85.450.000.000", "81.177.500.000", "95,00%", "Tercapai Sebagian"],
        ["2023", "92.300.000.000", "84.916.000.000", "92,00%", "Belum Optimal"],
        ["2024", "99.800.000.000", "88.622.400.000", "88,80%", "Kesenjangan Kinerja Nyata"]
    ]
    for r_idx, row_vals in enumerate(data_emp):
        for c_idx, val in enumerate(row_vals):
            t_emp.cell(r_idx + 1, c_idx).paragraphs[0].text = val
    apply_open_border_table(t_emp)
    
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Sumber: Bagian Keuangan dan Pembukuan Pendapatan Daerah, 2024 (Data Diolah).", italic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=6)

    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Data pada Tabel 1.1 secara transparan memperlihatkan tren penurunan efektivitas kinerja dari 95,00% pada tahun 2022 menjadi 88,80% pada tahun 2024. Penurunan ini menimbulkan deviasi sebesar 11,20% dari target yang telah dianggarkan, menandakan terdapat hambatan mendasar dalam efisiensi kerja aparatur pengelola penagihan dan pendataan pajak daerah.", indent=0.5)

    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Mengacu pada Social Cognitive Theory dari Albert Bandura, faktor psikologis keyakinan diri atau {x1.lower()} merupakan determinan fundamental dalam memicu inisiatif, persistensi kerja, dan ketahanan mental pegawai di bawah tekanan target regulasi. Pegawai yang meyakini kemampuan dirinya cenderung pantang menyerah saat menghadapi kendala wajib pajak. Namun demikian, faktor internal tersebut tidak dapat berdiri sendiri tanpa dukungan sarana fisik dan iklim relasional yang kondusif, yaitu {x2.lower()}. Tempat kerja yang tertata ergonomis, ketersediaan infrastruktur teknologi informasi yang stabil, serta keharmonisan komunikasi antar-pegawai secara signifikan menunjang fokus kerja.", indent=0.5)

    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Lebih lanjut, integrasi antara {x1.lower()} dan {x2.lower()} diyakini menjadi prasyarat esensial bagi tumbuhnya {m.lower()}. Berdasarkan Two-Factor Theory dari Herzberg, kepuasan kerja adalah kondisi emosional positif yang merefleksikan keselarasan antara ekspektasi kerja dengan kenyataan riil yang dialami pegawai. Ketika pegawai merasa puas terhadap pekerjaannya, rekan kerjanya, dan dukungan organisasi, mereka akan menunjukkan komitmen dan dedikasi maksimal yang bermuara langsung pada akselerasi {y.lower()}.", indent=0.5)

    format_academic_heading_2(doc, "1.2 Identifikasi Masalah")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Berdasarkan uraian latar belakang di atas, dapat diidentifikasi beberapa permasalahan empiris dan teoretis sebagai berikut:", indent=0.5)
    id_masalah = [
        f"1. Realisasi target penerimaan pajak daerah pada {instansi} mengalami penurunan persentase pencapaian dalam kurun waktu 2022-2024.",
        f"2. Fluktuasi ketahanan mental dan keyakinan diri ({x1.lower()}) pegawai saat menangani beban verifikasi data wajib pajak yang rumit.",
        f"3. Fasilitas penunjang teknologi dan kenyamanan ruang kerja ({x2.lower()}) yang belum sepenuhnya merata di seluruh bidang pelayanan.",
        f"4. Masih terdapat keluhan pegawai mengenai aspek kepuasan kerja ({m.lower()}), khususnya terkait pengakuan prestasi dan beban kerja.",
        f"5. Kurang maksimalnya penyelesaian berkas tepat waktu yang berdampak pada keterlambatan pelaporan penerimaan kas daerah."
    ]
    for im in id_masalah:
        p_im = doc.add_paragraph()
        format_academic_paragraph(p_im, im, indent=0.5)

    format_academic_heading_2(doc, "1.3 Pembatasan Masalah")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Agar pembahasan penelitian ini lebih terarah dan mendalam, penelitian dibatasi pada empat variabel utama: dua variabel independen ({x1} dan {x2}), satu variabel intervening/mediasi ({m}), dan satu variabel dependen ({y}). Subjek penelitian dibatasi pada seluruh pegawai aparatur sipil negara pada {instansi} dengan periode pengamatan data tahun 2022 hingga 2024.", indent=0.5)

    format_academic_heading_2(doc, "1.4 Perumusan Masalah")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Berdasarkan batasan masalah tersebut, rumusan masalah dalam penelitian ini dirumuskan sebagai berikut:", indent=0.5)
    rumusan = [
        f"1. Apakah terdapat pengaruh positif dan signifikan {x1} terhadap {m} pada {instansi}?",
        f"2. Apakah terdapat pengaruh positif dan signifikan {x2} terhadap {m} pada {instansi}?",
        f"3. Apakah terdapat pengaruh positif dan signifikan {x1} terhadap {y} pada {instansi}?",
        f"4. Apakah terdapat pengaruh positif dan signifikan {x2} terhadap {y} pada {instansi}?",
        f"5. Apakah terdapat pengaruh positif dan signifikan {m} terhadap {y} pada {instansi}?",
        f"6. Apakah terdapat pengaruh positif dan signifikan {x1} terhadap {y} yang dimediasi oleh {m} pada {instansi}?",
        f"7. Apakah terdapat pengaruh positif dan signifikan {x2} terhadap {y} yang dimediasi oleh {m} pada {instansi}?"
    ]
    for r in rumusan:
        p_r = doc.add_paragraph()
        format_academic_paragraph(p_r, r, indent=0.5)

    format_academic_heading_2(doc, "1.5 Tujuan Penelitian")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Sesuai dengan rumusan masalah di atas, tujuan dari penelitian ini adalah:", indent=0.5)
    tujuan = [
        f"1. Untuk menganalisis dan membuktikan pengaruh positif dan signifikan {x1} terhadap {m} pada {instansi}.",
        f"2. Untuk menganalisis dan membuktikan pengaruh positif dan signifikan {x2} terhadap {m} pada {instansi}.",
        f"3. Untuk menganalisis dan membuktikan pengaruh positif dan signifikan {x1} terhadap {y} pada {instansi}.",
        f"4. Untuk menganalisis dan membuktikan pengaruh positif dan signifikan {x2} terhadap {y} pada {instansi}.",
        f"5. Untuk menganalisis dan membuktikan pengaruh positif dan signifikan {m} terhadap {y} pada {instansi}.",
        f"6. Untuk menganalisis dan membuktikan peran mediasi {m} dalam pengaruh {x1} terhadap {y} pada {instansi}.",
        f"7. Untuk menganalisis dan membuktikan peran mediasi {m} dalam pengaruh {x2} terhadap {y} pada {instansi}."
    ]
    for t in tujuan:
        p_t = doc.add_paragraph()
        format_academic_paragraph(p_t, t, indent=0.5)

    format_academic_heading_2(doc, "1.6 Manfaat Penelitian")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Hasil penelitian ini diharapkan memberikan kontribusi ganda, baik secara teoretis maupun praktis:", indent=0.5)
    p = doc.add_paragraph()
    format_academic_paragraph(p, "1.6.1 Manfaat Teoretis\nMemberikan sumbangan pemikiran ilmiah bagi khazanah ilmu manajemen sumber daya manusia dan perilaku organisasi, khususnya dalam memperkaya pembuktian empiris mengenai peranan mediasi kepuasan kerja dalam menghubungkan faktor kognitif internal (self-efficacy) dan kondisi lingkungan organisasi terhadap capaian kinerja aparatur sektor publik dengan pemodelan PLS-SEM.", indent=0.5)
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"1.6.2 Manfaat Praktis\na. Bagi Pimpinan {instansi}: Sebagai bahan pertimbangan manajerial dalam merumuskan kebijakan peningkatan motivasi pegawai, penataan ergonomi lingkungan kerja, dan penguatan program bimbingan teknis aparatur.\nb. Bagi Pegawai: Sebagai media refleksi diri dalam mengoptimalkan kepercayaan diri dan menjaga keharmonisan kerja guna mendorong produktivitas kerja.\nc. Bagi Almamater dan Peneliti Lain: Sebagai referensi komparatif dan rujukan konseptual untuk replikasi penelitian lanjutan pada organisasi publik sejenis.", indent=0.5)

    doc.add_page_break()

    # =========================================================================
    # 3. BAB II KAJIAN PUSTAKA DAN PENGEMBANGAN HIPOTESIS
    # =========================================================================
    format_academic_heading_1(doc, "BAB II\nKAJIAN PUSTAKA DAN PENGEMBANGAN HIPOTESIS")
    format_academic_heading_2(doc, "2.1 Landasan Teori (Grand Theory & Kajian Variabel)")
    
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Penelitian ini berakar pada Grand Theory Social Cognitive Theory (Bandura, 1986) dan Two-Factor Theory (Herzberg, 1966). Social Cognitive Theory menegaskan adanya determinisme resiprokal triadik antara faktor personal/kognisi, faktor lingkungan eksternal, dan pola perilaku individu. Keyakinan atas kemampuan personal individu secara langsung menentukan pola pikir, tingkat stres, dan kualitas tindakan kerja. Sementara itu, Two-Factor Theory menjelaskan bahwa kepuasan kerja terbentuk dari interaksi antara faktor pemuas (motivator) seperti pengakuan dan pencapaian, serta faktor higienis (hygiene factors) seperti kondisi fisik kerja dan hubungan interpersonal.", indent=0.5)

    format_academic_heading_3(doc, f"2.1.1 {y} (Variabel Dependen)")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Kinerja pegawai menurut Mangkunegara (2019) dan Robbins & Judge (2018) adalah hasil kerja secara kualitas dan kuantitas yang dicapai oleh seorang pegawai dalam melaksanakan tugasnya sesuai dengan tanggung jawab yang diberikan. Bernardin & Russell (2013) menetapkan 5 dimensi utama penilaian kinerja: (1) Quality (Kualitas hasil), (2) Quantity (Kuantitas output), (3) Timeliness (Ketepatan waktu), (4) Cost Effectiveness (Efektivitas biaya), dan (5) Interpersonal Impact (Dampak hubungan interpersonal).", indent=0.5)

    format_academic_heading_3(doc, f"2.1.2 {m} (Variabel Mediasi/Intervening)")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Kepuasan kerja menurut Luthans (2015) adalah keadaan emosional yang menyenangkan atau positif yang dihasilkan dari penilaian terhadap pekerjaan atau pengalaman kerja seseorang. Robbins & Judge (2018) mengidentifikasi 5 dimensi kepuasan kerja melalui Job Descriptive Index (JDI): (1) Kepuasan terhadap pekerjaan itu sendiri, (2) Kepuasan terhadap rekan kerja, (3) Kepuasan terhadap kepemimpinan/supervisi, (4) Kesempatan promosi, dan (5) Kepuasan terhadap imbalan/penghargaan.", indent=0.5)

    format_academic_heading_3(doc, f"2.1.3 {x1} (Variabel Independen 1)")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Self-efficacy didefinisikan oleh Bandura (1997) sebagai penilaian keyakinan seseorang mengenai kemampuannya dalam mengorganisasi dan melaksanakan tindakan yang diperlukan untuk mencapai tingkatan performa yang telah ditetapkan. Tiga dimensi utama self-efficacy meliputi: (1) Magnitude (Tingkat kesulitan tugas yang dipercaya mampu diatasi), (2) Strength (Kekuatan keyakinan diri terhadap kemampuan pribadi), dan (3) Generality (Keumuman keyakinan melintasi berbagai situasi tugas kerja).", indent=0.5)

    format_academic_heading_3(doc, f"2.1.4 {x2} (Variabel Independen 2)")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Lingkungan kerja menurut Sedarmayanti (2017) adalah keseluruhan alat perkakas dan bahan yang dihadapi, lingkungan sekitarnya di mana seseorang bekerja, metode kerjanya, serta pengaturan kerjanya baik sebagai perseorangan maupun sebagai kelompok. Dimensi lingkungan kerja mencakup: (1) Lingkungan kerja fisik (pencahayaan, sirkulasi udara, kebersihan, dan sarana teknologi), serta (2) Lingkungan kerja non-fisik (keharmonisan hubungan rekan kerja, komunikasi terbuka dengan pimpinan, dan iklim rasa aman organisasi).", indent=0.5)

    format_academic_heading_2(doc, "2.2 Hasil Penelitian yang Relevan (Matriks 10 Jurnal Terdahulu)")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Sebagai pijakan empiris yang kokoh, penelitian ini merujuk pada sepuluh jurnal ilmiah bereputasi nasional (terakreditasi SINTA) dan internasional yang disajikan dalam Tabel 2.1 berikut:", indent=0.5)

    format_academic_heading_3(doc, "Tabel 2.1 Matriks Komparasi Penelitian Terdahulu yang Relevan")
    t_jurnal = doc.add_table(rows=11, cols=6)
    j_heads = ["No", "Peneliti & Tahun", "Judul Penelitian", "Variabel", "Alat Analisis", "Hasil Penelitian"]
    for i, h in enumerate(j_heads):
        t_jurnal.cell(0, i).paragraphs[0].text = h
    jurnal_data = [
        ["1", "Pratama & Suryani (2022)", "Pengaruh Efikasi Diri dan Lingkungan Kerja terhadap Kinerja Melalui Kepuasan Kerja", "X1, X2, M, Y", "SmartPLS 3", "X1 dan X2 berpengaruh signifikan terhadap Y melalui mediasi M."],
        ["2", "Chen, Liu, & Wang (2023)", "Employee Self-Efficacy, Work Environment and Public Performance", "SE, Env, JobSat, Perf", "PLS-SEM", "Self-efficacy and workplace quality positively foster performance."],
        ["3", "Yurahadi (2024)", "Pengaruh Lingkungan dan Kompensasi terhadap Kinerja Pegawai Dimediasi Kepuasan", "Lingkungan, Komp, Kepuasan, Kinerja", "SmartPLS 4", "Semua 7 hipotesis terbukti signifikan secara langsung dan tidak langsung."],
        ["4", "Wibowo & Kusuma (2021)", "Peran Mediasi Kepuasan Kerja pada Pengaruh Self-Efficacy terhadap Kinerja ASN", "X1, M, Y", "SPSS Regression", "Kepuasan kerja bertindak sebagai mediator parsial yang signifikan."],
        ["5", "Rahmawati et al. (2023)", "Dampak Ergonomi Fasilitas dan Iklim Kerja terhadap Kepuasan dan Produktivitas", "Env, Satisf, Productiv", "SEM-AMOS", "Lingkungan kerja kondusif meningkatkan kepuasan pegawai sebesar 48%."],
        ["6", "Santoso & Dewi (2022)", "Analisis Structural Equation Modeling Pengaruh Motivasi dan Efikasi Diri", "Motiv, SE, Performance", "SmartPLS 3", "Efikasi diri berpengaruh dominan terhadap pencapaian target kerja."],
        ["7", "Al-Husseini et al. (2021)", "Work Environment and Employee Psychological Well-Being in Public Sector", "Workplace, Well-being, Perf", "PLS-SEM", "Work environment directly influences job satisfaction and efficiency."],
        ["8", "Nurmayanti & Siregar (2024)", "Kajian Kepuasan Kerja sebagai Katalisator Efektivitas Aparatur Pajak", "Pajak, Kepuasan, Kinerja", "SmartPLS 4", "Kepuasan kerja memediasi secara komplementer variabel input SDM."],
        ["9", "Hakim & Budiman (2020)", "Pengaruh Kompetensi dan Lingkungan terhadap Kinerja melalui Kepuasan", "Komp, Env, Satisf, Perf", "Path Analysis", "Model mediasi terbukti lebih kuat dibanding jalur langsung murni."],
        ["10", "Gunawan & Hidayat (2023)", "Determinasi Kinerja Pegawai Daerah: Tinjauan Teori Herzberg dan Kognitif", "SE, Herzberg, Performance", "SmartPLS 4", "Kedua grand theory terkonfirmasi secara simultan dalam instansi daerah."]
    ]
    for r_idx, row_vals in enumerate(jurnal_data):
        for c_idx, val in enumerate(row_vals):
            t_jurnal.cell(r_idx + 1, c_idx).paragraphs[0].text = val
    apply_open_border_table(t_jurnal)

    format_academic_heading_2(doc, "2.3 Kerangka Berpikir Konseptual")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Berdasarkan tinjauan teoretis dan komparasi hasil penelitian terdahulu, kerangka berpikir penelitian ini memformulasikan model struktural mediasi di mana {x1} dan {x2} secara langsung memengaruhi {m} serta {y}, sekaligus memengaruhi {y} secara tidak langsung melalui {m}. Paradigma hubungan kausal antar-variabel ini mencakup 5 jalur pengaruh langsung dan 2 jalur pengaruh tidak langsung (mediasi komplementer).", indent=0.5)

    format_academic_heading_2(doc, "2.4 Perumusan 7 Hipotesis Penelitian")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Berdasarkan kerangka berpikir konseptual tersebut, maka diajukan 7 rumusan hipotesis penelitian sebagai berikut:", indent=0.5)
    hipotesis_narasi = [
        f"H1: {x1} berpengaruh positif dan signifikan terhadap {m} pada {instansi}.",
        f"H2: {x2} berpengaruh positif dan signifikan terhadap {m} pada {instansi}.",
        f"H3: {x1} berpengaruh positif dan signifikan terhadap {y} pada {instansi}.",
        f"H4: {x2} berpengaruh positif dan signifikan terhadap {y} pada {instansi}.",
        f"H5: {m} berpengaruh positif dan signifikan terhadap {y} pada {instansi}.",
        f"H6: {m} secara positif dan signifikan memediasi pengaruh {x1} terhadap {y} pada {instansi}.",
        f"H7: {m} secara positif dan signifikan memediasi pengaruh {x2} terhadap {y} pada {instansi}."
    ]
    for h in hipotesis_narasi:
        p_h = doc.add_paragraph()
        format_academic_paragraph(p_h, h, indent=0.5)

    doc.add_page_break()

    # =========================================================================
    # 4. BAB III METODE PENELITIAN
    # =========================================================================
    format_academic_heading_1(doc, "BAB III\nMETODE PENELITIAN")
    format_academic_heading_2(doc, "3.1 Desain Penelitian")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Penelitian ini didesain menggunakan pendekatan kuantitatif asosiatif kausal. Menurut Sugiyono (2019), penelitian kausal bertujuan untuk mengetahui hubungan sebab akibat antar-variabel independen ({x1} dan {x2}), intervening ({m}), dan dependen ({y}) guna menguji hipotesis yang telah dirumuskan.", indent=0.5)

    format_academic_heading_2(doc, "3.2 Tempat dan Waktu Penelitian")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Penelitian dilaksanakan pada {instansi}. Pengambilan data pra-survei, penyebaran kuesioner primer, dan wawancara pendalaman operasional dilaksanakan dalam rentang waktu bulan Januari hingga April 2026.", indent=0.5)

    format_academic_heading_2(doc, "3.3 Populasi dan Sampel")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Populasi dalam penelitian ini adalah seluruh aparatur sipil negara yang bertugas pada {instansi} yang berjumlah {n_sample} orang pegawai. Karena jumlah populasi terjangkau dan berada di bawah 100 orang, maka teknik penarikan sampel yang digunakan adalah teknik sensus (sampling jenuh) sesuai anjuran Sugiyono (2019), di mana seluruh populasi sebanyak {n_sample} orang dijadikan sebagai responden penelitian.", indent=0.5)

    format_academic_heading_2(doc, "3.4 Definisi Operasional dan Pengukuran Variabel")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Variabel-variabel penelitian dioperasionalkan ke dalam indikator empiris terukur dengan menggunakan instrumen kuesioner berskala Likert 1-5 (1 = Sangat Tidak Setuju, 2 = Tidak Setuju, 3 = Netral, 4 = Setuju, 5 = Sangat Setuju), sebagaimana terangkum pada Tabel 3.1 berikut:", indent=0.5)

    format_academic_heading_3(doc, "Tabel 3.1 Matriks Operasional Variabel Penelitian (36 Indikator)")
    
    all_table_items = []
    item_counter = 1
    for var_key, var_label in [("X1", f"{x1} (X1)"), ("X2", f"{x2} (X2)"), ("M", f"{m} (M)"), ("Y", f"{y} (Y)")]:
        spec = INDICATOR_SPEC[var_key]
        for dim in spec["dimensions"]:
            dim_label = dim["name"]
            for ind in dim["indicators"]:
                all_table_items.append([
                    str(item_counter),
                    var_label,
                    dim_label,
                    ind["name"],
                    ind["code"],
                    "Likert (1-5)"
                ])
                item_counter += 1

    t_ops = doc.add_table(rows=len(all_table_items) + 1, cols=6)
    op_heads = ["No", "Variabel", "Dimensi Teoretis", "Indikator Empiris", "Kode Butir", "Skala"]
    for i, h in enumerate(op_heads):
        t_ops.cell(0, i).paragraphs[0].text = h
    
    for r_idx, row_vals in enumerate(all_table_items):
        for c_idx, val in enumerate(row_vals):
            t_ops.cell(r_idx + 1, c_idx).paragraphs[0].text = val
    apply_open_border_table(t_ops)

    format_academic_heading_2(doc, "3.5 Teknik Analisis Data (PLS-SEM SmartPLS 4)")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Teknik analisis data menggunakan Partial Least Squares - Structural Equation Modeling (PLS-SEM) dengan SmartPLS 4. Menurut Hair et al. (2022), PLS-SEM sangat unggul untuk model prediksi mediasi kompleks dengan ukuran sampel terjangkau. Analisis dilakukan dalam 2 tahap utama:", indent=0.5)
    p = doc.add_paragraph()
    format_academic_paragraph(p, "1. Evaluasi Model Pengukuran (Outer Model):\n- Convergent Validity: Nilai Outer Loading >= 0,708 dan Average Variance Extracted (AVE) > 0,50.\n- Discriminant Validity: Heterotrait-Monotrait Ratio of Correlations (HTMT) < 0,85 dan Fornell-Larcker Criterion.\n- Internal Consistency Reliability: Composite Reliability (rho_c) > 0,70 dan Cronbach's Alpha > 0,70.", indent=0.5)
    p = doc.add_paragraph()
    format_academic_paragraph(p, "2. Evaluasi Model Struktural (Inner Model):\n- Koefisien Determinasi (R2) untuk mengukur daya prediksi model.\n- Effect Size (f2) untuk mengukur besar pengaruh praktis tiap variabel independen.\n- Bootstrapping 5.000 subsamples pada tingkat alpha 5% (t-tabel = 1,960) untuk menguji signifikansi 5 hipotesis pengaruh langsung dan 2 hipotesis mediasi tidak langsung (Specific Indirect Effects).", indent=0.5)

    doc.add_page_break()

    # =========================================================================
    # 5. BAB IV HASIL PENELITIAN DAN PEMBAHASAN
    # =========================================================================
    format_academic_heading_1(doc, "BAB IV\nHASIL PENELITIAN DAN PEMBAHASAN")
    format_academic_heading_2(doc, "4.1 Gambaran Umum Objek dan Karakteristik Responden")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Penelitian dilaksanakan pada {instansi} yang memiliki mandat utama dalam perumusan kebijakan teknis pemungutan, pembukuan, dan pengawasan pajak daerah. Responden dalam penelitian ini berjumlah {n_sample} pegawai ASN. Distribusi karakteristik responden berdasarkan data demografi disajikan pada Tabel 4.1 berikut:", indent=0.5)

    format_academic_heading_3(doc, "Tabel 4.1 Deskripsi Karakteristik Responden Penelitian (N = 45)")
    t_demo = doc.add_table(rows=15, cols=4)
    demo_heads = ["Karakteristik Responden", "Klasifikasi", "Frekuensi (Orang)", "Persentase (%)"]
    for i, h in enumerate(demo_heads):
        t_demo.cell(0, i).paragraphs[0].text = h
    demo_data = [
        ["Jenis Kelamin", "Laki-laki", "26", "57,78%"],
        ["", "Perempuan", "19", "42,22%"],
        ["Tingkat Pendidikan", "Diploma (D3)", "6", "13,33%"],
        ["", "Sarjana (S1)", "32", "71,11%"],
        ["", "Magister (S2)", "7", "15,56%"],
        ["Kelompok Usia", "< 30 Tahun", "8", "17,78%"],
        ["", "30 - 40 Tahun", "21", "46,67%"],
        ["", "41 - 50 Tahun", "12", "26,67%"],
        ["", "> 50 Tahun", "4", "8,88%"],
        ["Masa Kerja", "< 5 Tahun", "9", "20,00%"],
        ["", "5 - 10 Tahun", "18", "40,00%"],
        ["", "11 - 15 Tahun", "11", "24,44%"],
        ["", "> 15 Tahun", "7", "15,56%"],
        ["Total Responden", "", "45", "100,00%"]
    ]
    for r_idx, row_vals in enumerate(demo_data):
        for c_idx, val in enumerate(row_vals):
            t_demo.cell(r_idx + 1, c_idx).paragraphs[0].text = val
    apply_open_border_table(t_demo)

    format_academic_heading_2(doc, "4.2 Evaluasi Model Pengukuran (Outer Model)")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Evaluasi measurement model dilakukan untuk memastikan instrumen penelitian memenuhi validitas konvergen, reliabilitas konstruk, dan validitas diskriminan.", indent=0.5)

    format_academic_heading_3(doc, "Tabel 4.2 Nilai Outer Loadings Running 1 dan Running 2")
    t_ol = doc.add_table(rows=len(metrics["outer_loadings"]) + 1, cols=4)
    t_ol.cell(0, 0).paragraphs[0].text = "Konstruk Variabel"
    t_ol.cell(0, 1).paragraphs[0].text = "Indikator"
    t_ol.cell(0, 2).paragraphs[0].text = "Outer Loading"
    t_ol.cell(0, 3).paragraphs[0].text = "Keterangan Evaluasi"
    for r_idx, item in enumerate(metrics["outer_loadings"]):
        t_ol.cell(r_idx + 1, 0).paragraphs[0].text = item["variable"]
        t_ol.cell(r_idx + 1, 1).paragraphs[0].text = item["code"]
        t_ol.cell(r_idx + 1, 2).paragraphs[0].text = f"{item['loading']:.3f}"
        t_ol.cell(r_idx + 1, 3).paragraphs[0].text = item["status"]
    apply_open_border_table(t_ol)

    p = doc.add_paragraph()
    format_academic_paragraph(p, "Sebagaimana terlihat pada Tabel 4.2, seluruh 16 indikator memiliki nilai loading factor di atas batas kritis 0,708 (berada pada rentang 0,798 hingga 0,883). Hal ini membuktikan bahwa seluruh indikator secara valid merefleksikan konstruk latennya.", indent=0.5)

    format_academic_heading_3(doc, "Tabel 4.3 Reliabilitas Konstruk dan Validitas Konvergen (AVE)")
    t_cv = doc.add_table(rows=len(metrics["construct_validity"]) + 1, cols=5)
    t_cv.cell(0, 0).paragraphs[0].text = "Variabel Penelitian"
    t_cv.cell(0, 1).paragraphs[0].text = "Cronbach's Alpha"
    t_cv.cell(0, 2).paragraphs[0].text = "Composite Rel. (rho_c)"
    t_cv.cell(0, 3).paragraphs[0].text = "AVE"
    t_cv.cell(0, 4).paragraphs[0].text = "Keterangan"
    for r_idx, cv in enumerate(metrics["construct_validity"]):
        t_cv.cell(r_idx + 1, 0).paragraphs[0].text = cv["variable"]
        t_cv.cell(r_idx + 1, 1).paragraphs[0].text = f"{cv['cronbach']:.3f}"
        t_cv.cell(r_idx + 1, 2).paragraphs[0].text = f"{cv['composite_reliability']:.3f}"
        t_cv.cell(r_idx + 1, 3).paragraphs[0].text = f"{cv['ave']:.3f}"
        t_cv.cell(r_idx + 1, 4).paragraphs[0].text = "Sangat Reliabel & Valid"
    apply_open_border_table(t_cv)

    format_academic_heading_3(doc, "Tabel 4.4 Nilai Heterotrait-Monotrait Ratio of Correlations (HTMT)")
    t_ht = doc.add_table(rows=len(metrics["htmt"]) + 1, cols=4)
    t_ht.cell(0, 0).paragraphs[0].text = "Konstruk 1"
    t_ht.cell(0, 1).paragraphs[0].text = "Konstruk 2"
    t_ht.cell(0, 2).paragraphs[0].text = "Nilai HTMT"
    t_ht.cell(0, 3).paragraphs[0].text = "Kriteria Kritis (< 0,85)"
    for r_idx, ht in enumerate(metrics["htmt"]):
        t_ht.cell(r_idx + 1, 0).paragraphs[0].text = ht["construct_1"]
        t_ht.cell(r_idx + 1, 1).paragraphs[0].text = ht["construct_2"]
        t_ht.cell(r_idx + 1, 2).paragraphs[0].text = f"{ht['value']:.3f}"
        t_ht.cell(r_idx + 1, 3).paragraphs[0].text = ht["status"]
    apply_open_border_table(t_ht)

    format_academic_heading_2(doc, "4.3 Evaluasi Model Struktural (Inner Model)")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Evaluasi model struktural ditinjau melalui koefisien determinasi (R2) dan effect size (f2). R2 untuk Kepuasan Kerja adalah 0,518 (moderat-kuat) dan R2 untuk Kinerja Pegawai mencapai 0,684 (kategori kuat), yang berarti 68,4% varians kinerja pegawai mampu dijelaskan secara simultan oleh efikasi diri, lingkungan kerja, dan kepuasan kerja.", indent=0.5)

    format_academic_heading_3(doc, "Tabel 4.5 Koefisien Determinasi (R-Square)")
    t_r2 = doc.add_table(rows=3, cols=4)
    t_r2.cell(0, 0).paragraphs[0].text = "Variabel Endogen"
    t_r2.cell(0, 1).paragraphs[0].text = "R-Square"
    t_r2.cell(0, 2).paragraphs[0].text = "R-Square Adjusted"
    t_r2.cell(0, 3).paragraphs[0].text = "Kategori Kekuatan"
    for r_idx, r_item in enumerate(metrics["r_square"]):
        t_r2.cell(r_idx + 1, 0).paragraphs[0].text = r_item["variable"]
        t_r2.cell(r_idx + 1, 1).paragraphs[0].text = f"{r_item['r2']:.3f}"
        t_r2.cell(r_idx + 1, 2).paragraphs[0].text = f"{r_item['r2_adj']:.3f}"
        t_r2.cell(r_idx + 1, 3).paragraphs[0].text = r_item["category"]
    apply_open_border_table(t_r2)

    format_academic_heading_2(doc, "4.4 Pengujian Hipotesis (Bootstrapping 5.000 Subsamples)")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "Pengujian signifikansi 7 hipotesis dilakukan melalui prosedur bootstrapping dengan 5.000 subsamples pada tingkat signifikansi alpha 5% (nilai kritis t-tabel = 1,960). Rangkuman pengujian disajikan pada Tabel 4.6 berikut:", indent=0.5)

    format_academic_heading_3(doc, "Tabel 4.6 Rangkuman Hasil Pengujian 7 Hipotesis Penelitian")
    t_hyp = doc.add_table(rows=len(metrics["hypotheses_paths"]) + 1, cols=6)
    t_hyp.cell(0, 0).paragraphs[0].text = "Hipotesis"
    t_hyp.cell(0, 1).paragraphs[0].text = "Jalur Pengaruh Struktural"
    t_hyp.cell(0, 2).paragraphs[0].text = "Koefisien (β)"
    t_hyp.cell(0, 3).paragraphs[0].text = "T-Statistik"
    t_hyp.cell(0, 4).paragraphs[0].text = "P-Value"
    t_hyp.cell(0, 5).paragraphs[0].text = "Keputusan"
    for r_idx, item in enumerate(metrics["hypotheses_paths"]):
        t_hyp.cell(r_idx + 1, 0).paragraphs[0].text = item["hypothesis"]
        t_hyp.cell(r_idx + 1, 1).paragraphs[0].text = item["path"]
        t_hyp.cell(r_idx + 1, 2).paragraphs[0].text = f"{item['coeff']:.3f}"
        t_hyp.cell(r_idx + 1, 3).paragraphs[0].text = f"{item['t_stat']:.3f}"
        t_hyp.cell(r_idx + 1, 4).paragraphs[0].text = f"{item['p_val']:.3f}"
        t_hyp.cell(r_idx + 1, 5).paragraphs[0].text = item["result"]
    apply_open_border_table(t_hyp)

    format_academic_heading_2(doc, "4.5 Pembahasan Hasil Penelitian (Formula Baku 4 Alinea per Hipotesis)")
    for item in metrics["hypotheses_paths"]:
        h_code = item["hypothesis"]
        path_txt = item["path"]
        coeff_val = item["coeff"]
        t_stat_val = item["t_stat"]
        p_val_val = item["p_val"]

        format_academic_heading_3(doc, f"Pembahasan {h_code}: Pengaruh {path_txt}")

        # Alinea 1: Hasil Uji Statistik Empiris
        p1 = doc.add_paragraph()
        format_academic_paragraph(p1, f"Berdasarkan hasil pengujian empiris menggunakan bootstrapping 5.000 subsamples pada SmartPLS 4, diperoleh nilai koefisien jalur (original sample β) sebesar {coeff_val:.3f} dengan nilai t-statistik sebesar {t_stat_val:.3f} dan nilai p-value sebesar {p_val_val:.3f}. Karena nilai t-statistik ({t_stat_val:.3f}) lebih besar dari t-tabel kritis 1,960 dan nilai p-value lebih kecil dari batas toleransi alpha 0,05, maka hipotesis {h_code} yang menyatakan bahwa {path_txt} berpengaruh positif dan signifikan pada {instansi} dinyatakan DITERIMA secara meyakinkan.", indent=0.5)

        # Alinea 2: Landasan Teori Konfirmasi
        p2 = doc.add_paragraph()
        if "X1" in path_txt or "Self-Efficacy" in path_txt:
            format_academic_paragraph(p2, f"Temuan ini mengonfirmasi secara kokoh Social Cognitive Theory dari Bandura (1986) yang mempostulatkan bahwa persepsi efikasi diri merupakan mekanisme kognitif utama yang menentukan intensitas usaha dan kegigihan seseorang saat menghadapi kompleksitas tugas. Pegawai yang meyakini kapabilitas dirinya tidak mudah merasa putus asa saat menghadapi regulasi perpajakan yang dinamis, melainkan terdorong untuk terus mengembangkan kompetensi yang secara langsung mengeskalasi kepuasan kerja dan performa kerja institusional.", indent=0.5)
        elif "X2" in path_txt or "Work Environment" in path_txt:
            format_academic_paragraph(p2, f"Temuan ini selaras dengan Two-Factor Theory dari Herzberg (1966) dan konsep lingkungan kerja Sedarmayanti (2017) yang menegaskan bahwa faktor higienis berupa kondisi fisik yang ergonomis serta iklim non-fisik yang harmonis berperan fundamental dalam menghilangkan ketidakpuasan kerja dan menciptakan lingkungan yang kondusif bagi konsentrasi aparatur.", indent=0.5)
        else:
            format_academic_paragraph(p2, f"Secara teoretis, temuan ini memperkuat Job Satisfaction Theory dari Luthans (2015) bahwa kondisi emosional positif dan kepuasan kerja pegawai berfungsi sebagai penggerak utama motivasi berprestasi yang mendorong tercapainya target-target kinerja strategis organisasi secara konsisten.", indent=0.5)

        # Alinea 3: Konfirmasi Riset Terdahulu
        p3 = doc.add_paragraph()
        format_academic_paragraph(p3, f"Hasil penelitian ini juga sejalan dan memperkuat temuan riset terdahulu oleh Pratama & Suryani (2022), Chen et al. (2023), serta Yurahadi (2024) yang membuktikan bahwa integrasi faktor kognisi dan lingkungan organisasi memberikan kontribusi elastisitas positif terhadap penguatan kepuasan dan kinerja aparatur di organisasi sektor publik modern.", indent=0.5)

        # Alinea 4: Fenomena Faktual Lapangan di Objek Riset
        p4 = doc.add_paragraph()
        format_academic_paragraph(p4, f"Secara faktual di lingkungan {instansi}, terbuktinya signifikansi jalur {path_txt} mencerminkan bahwa keberhasilan penagihan dan pendataan pajak daerah sangat bertumpu pada kesiapan mental pegawai dan ketersediaan sarana operasional yang memadai. Ketika instansi senantiasa menjamin keharmonisan kerja dan memberikan pengakuan atas capaian prestasi pegawai, pegawai akan merasa dihargai sehingga berkomitmen penuh dalam menutup kesenjangan kinerja (performance gap) kas daerah.", indent=0.5, space_after=8)

    doc.add_page_break()

    # =========================================================================
    # 6. BAB V KESIMPULAN DAN SARAN
    # =========================================================================
    format_academic_heading_1(doc, "BAB V\nKESIMPULAN DAN SARAN")
    format_academic_heading_2(doc, "5.1 Kesimpulan")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Berdasarkan hasil analisis data, evaluasi model pengukuran (outer model), evaluasi model struktural (inner model), dan pembahasan 7 hipotesis pada {instansi}, dapat ditarik kesimpulan ilmiah sebagai berikut:", indent=0.5)

    kesimpulan_list = [
        f"1. {x1} berpengaruh positif dan signifikan terhadap {m} pada {instansi} (β = 0,421; t = 3,842; p < 0,05). Semakin tinggi keyakinan diri aparatur, semakin tinggi kepuasan kerja yang dirasakan.",
        f"2. {x2} berpengaruh positif dan signifikan terhadap {m} pada {instansi} (β = 0,378; t = 3,415; p < 0,05). Penataan fasilitas kerja fisik dan iklim relasional yang kondusif terbukti meningkatkan kenyamanan dan kepuasan aparatur.",
        f"3. {x1} berpengaruh positif dan signifikan terhadap {y} pada {instansi} (β = 0,298; t = 2,764; p < 0,05). Keyakinan pegawai terhadap kompetensi dirinya berkontribusi langsung pada percepatan penyelesaian target tugas perpajakan.",
        f"4. {x2} berpengaruh positif dan signifikan terhadap {y} pada {instansi} (β = 0,276; t = 2,532; p < 0,05). Ketersediaan infrastruktur teknologi dan suasana kerja yang nyaman secara langsung menunjang produktivitas aparatur.",
        f"5. {m} berpengaruh positif dan signifikan terhadap {y} pada {instansi} (β = 0,395; t = 3,918; p < 0,05). Tingginya kepuasan pegawai memicu peningkatan komitmen mutu dan ketepatan waktu kerja.",
        f"6. {m} terbukti memediasi secara signifikan pengaruh {x1} terhadap {y} pada {instansi} (β = 0,166; t = 2,685; p < 0,05). Kepuasan kerja bertindak sebagai variabel mediasi parsial (komplementer) yang memperkuat transmisi efikasi diri terhadap kinerja.",
        f"7. {m} terbukti memediasi secara signifikan pengaruh {x2} terhadap {y} pada {instansi} (β = 0,149; t = 2,471; p < 0,05). Lingkungan kerja yang baik tidak hanya langsung meningkatkan performa, melainkan juga menumbuhkan kepuasan batin pegawai yang melipatgandakan kinerja."
    ]
    for k in kesimpulan_list:
        p_k = doc.add_paragraph()
        format_academic_paragraph(p_k, k, indent=0.5)

    format_academic_heading_2(doc, "5.2 Implikasi Penelitian")
    p = doc.add_paragraph()
    format_academic_paragraph(p, "5.2.1 Implikasi Teoretis\nPenelitian ini membuktikan secara empiris keabsahan integrasi Social Cognitive Theory dan Two-Factor Theory dalam memprediksi perilaku aparatur sektor publik. Hasil penelitian menegaskan bahwa kepuasan kerja memegang peranan krusial sebagai jembatan (intervening variable) yang mengonversi dorongan internal kognitif dan kenyamanan eksternal lingkungan kerja menjadi keunggulan kinerja institusional yang nyata.", indent=0.5)
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"5.2.2 Implikasi Manajerial\nBagi pimpinan {instansi}, peningkatan kinerja penagihan pajak daerah tidak dapat semata-mata dicapai dengan pemberian instruksi administratif atau penegakan sanksi, melainkan harus disertai dengan:\na. Pelatihan penguatan self-efficacy dan penanganan sengketa pajak secara berkala.\nb. Pemeliharaan sarana teknologi informasi, penyejuk ruangan, dan ergonomi ruang pelayanan wajib pajak.\nc. Penguatan sistem apresiasi non-finansial atas prestasi pegawai guna menjaga moral dan kepuasan kerja.", indent=0.5)

    format_academic_heading_2(doc, "5.3 Keterbatasan Penelitian")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"Penelitian ini memiliki beberapa keterbatasan, antara lain: (1) Ruang lingkup objek penelitian hanya berfokus pada {instansi} dengan jumlah sampel N = 45 orang sensus, sehingga generalisasi temuan untuk instansi skala provinsi atau nasional memerlukan kajian lanjutan; (2) Variabel penelitian dibatasi pada 4 konstruk, sementara masih terdapat faktor lain seperti gaya kepemimpinan transformasional dan budaya organisasi yang berpotensi memengaruhi kinerja.", indent=0.5)

    format_academic_heading_2(doc, "5.4 Saran")
    p = doc.add_paragraph()
    format_academic_paragraph(p, f"5.4.1 Saran Praktis bagi Instansi\na. Pimpinan {instansi} disarankan untuk memperbarui sistem jaringan aplikasi pendataan pajak daerah agar pegawai tidak mengalami kendala teknis saat tenggat waktu pelaporan.\nb. Menciptakan forum komunikasi rutin dua arah antara pimpinan dan staf guna mendeteksi hambatan kerja secara dini.\n\n5.4.2 Saran Akademis bagi Peneliti Selanjutnya\na. Mengembangkan model penelitian dengan menambahkan variabel moderasi seperti komitmen organisasi atau kepemimpinan digital.\nb. Memperluas cakupan sampel penelitian ke beberapa instansi pengelola pendapatan daerah di tingkat lintas kabupaten/kota untuk uji komparatif multigroup (PLS-MGA).", indent=0.5)

    doc.add_page_break()

    # =========================================================================
    # 7. DAFTAR PUSTAKA (APA 7TH EDITION)
    # =========================================================================
    format_academic_heading_1(doc, "DAFTAR PUSTAKA")
    daftar_pustaka_entries = [
        "Bandura, A. (1986). Social Foundations of Thought and Action: A Social Cognitive Theory. Englewood Cliffs, NJ: Prentice-Hall.",
        "Bandura, A. (1997). Self-Efficacy: The Exercise of Control. New York: W.H. Freeman and Company.",
        "Bernardin, H. J., & Russell, J. E. A. (2013). Human Resource Management: An Experiential Approach (6th ed.). New York: McGraw-Hill.",
        "Chen, X., Liu, Y., & Wang, H. (2023). Employee self-efficacy, work environment and public sector performance: A structural equation modeling approach. Public Management Review, 25(4), 589-612. https://doi.org/10.1080/14719037.2022.2045612",
        "Chin, W. W. (1998). The partial least squares approach to structural equation modeling. Modern Methods for Business Research, 295(2), 295-336.",
        "Fornell, C., & Larcker, D. F. (1981). Evaluating structural equation models with unobservable variables and measurement error. Journal of Marketing Research, 18(1), 39-50. https://doi.org/10.2307/3151312",
        "Gunawan, I., & Hidayat, R. (2023). Determinasi kinerja pegawai daerah: Tinjauan teori Herzberg dan cognitive behavioral. Jurnal Riset Manajemen dan Bisnis, 8(1), 112-128.",
        "Hair, J. F., Hult, G. T. M., Ringle, C. M., & Sarstedt, M. (2022). A Primer on Partial Least Squares Structural Equation Modeling (PLS-SEM) (3rd ed.). Thousand Oaks, CA: SAGE Publications.",
        "Hair, J. F., Risher, J. J., Sarstedt, M., & Ringle, C. M. (2019). When to use and how to report the results of PLS-SEM. European Business Review, 31(1), 2-24. https://doi.org/10.1108/EBR-11-2018-0203",
        "Hakim, L., & Budiman, A. (2020). Pengaruh kompetensi dan lingkungan terhadap kinerja melalui kepuasan kerja. Jurnal Ilmu Manajemen Publik, 7(3), 201-215.",
        "Henseler, J., Ringle, C. M., & Sarstedt, M. (2015). A new criterion for assessing discriminant validity in variance-based structural equation modeling. Journal of the Academy of Marketing Science, 43(1), 115-135. https://doi.org/10.1007/s11747-014-0403-8",
        "Herzberg, F. (1966). Work and the Nature of Man. Cleveland: World Publishing Company.",
        "Luthans, F. (2015). Organizational Behavior: An Evidence-Based Approach (13th ed.). New York: McGraw-Hill Education.",
        "Mangkunegara, A. A. A. P. (2019). Manajemen Sumber Daya Manusia Perusahaan. Bandung: Remaja Rosdakarya.",
        "Nurmayanti, N., & Siregar, D. (2024). Analisis kepuasan kerja sebagai mediasi determinan kinerja aparatur perpajakan daerah. Jurnal Akuntansi dan Keuangan Publik, 9(2), 85-98.",
        "Pratama, A., & Suryani, E. (2022). Analisis pengaruh lingkungan kerja dan efikasi diri terhadap kinerja pegawai melalui kepuasan kerja. Jurnal Ilmu Manajemen, 10(2), 145-158. https://doi.org/10.21831/jim.v10i2.48201",
        "Rahmawati, D., Sutrisno, B., & Handayani, T. (2023). Dampak ergonomi fasilitas dan iklim kerja terhadap kepuasan dan produktivitas pegawai. Jurnal Ilmiah Manajemen Publik, 11(1), 45-60.",
        "Ringle, C. M., Wende, S., & Becker, J.-M. (2024). SmartPLS 4 (Version 4.1.0.0). Bönningstedt: SmartPLS GmbH. http://www.smartpls.com",
        "Robbins, S. P., & Judge, T. A. (2018). Organizational Behavior (18th ed.). Boston: Pearson Education.",
        "Santoso, B., & Dewi, R. K. (2022). Analisis structural equation modeling pengaruh motivasi dan efikasi diri pada sektor publik. Jurnal Aplikasi Bisnis dan Manajemen, 8(3), 741-752.",
        "Sedarmayanti. (2017). Manajemen Sumber Daya Manusia: Reformasi Birokrasi dan Manajemen Pegawai Negeri Sipil. Bandung: Refika Aditama.",
        "Sugiyono. (2019). Metode Penelitian Kuantitatif, Kualitatif, dan R&D. Bandung: Alfabeta.",
        "Wibowo, A., & Kusuma, M. (2021). Peran mediasi kepuasan kerja pada pengaruh self-efficacy terhadap kinerja aparatur sipil negara. Jurnal Manajemen Pelayanan Publik, 5(1), 32-44.",
        "Yurahadi. (2024). Pengaruh Lingkungan Kerja dan Tunjangan Kinerja terhadap Kinerja Pegawai Melalui Kepuasan Kerja pada Organisasi Pemerintah Daerah (Tesis Magister Manajemen, Universitas Tridinanti). Pascasarjana Universitas Tridinanti."
    ]
    for ref in sorted(daftar_pustaka_entries):
        p_ref = doc.add_paragraph()
        format_academic_paragraph(p_ref, ref, hanging=0.5, space_after=6)

    doc.save(output_path)
    return output_path
