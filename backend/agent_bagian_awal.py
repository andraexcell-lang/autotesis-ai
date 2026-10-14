# backend/agent_bagian_awal.py
"""
Agen Spesialis Pembuat Bagian Awal (Front Matter) Tesis Magister Manajemen
Standar Universitas Tridinanti / Tesis Yurahadi / Nurmayanti:
- Cover Luar (Piramida Terbalik)
- Cover Dalam
- Lembar Persetujuan Komisi Pembimbing Tesis
- Lembar Pengesahan Komisi Penguji Tesis
- Surat Pernyataan Bebas Plagiarisme (Materai 10.000)
- Kata Pengantar
- Abstrak Bahasa Indonesia
- Abstract Bahasa Inggris (Italic)
- Daftar Riwayat Hidup (Curriculum Vitae)
- Daftar Isi, Daftar Tabel, Daftar Gambar
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_border(cell, **kwargs):
    """Sets open-border cell formatting (horizontal borders only)."""
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

def format_p(p, text="", bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_before=0, space_after=0):
    p.alignment = align
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if indent > 0:
        p.paragraph_format.first_line_indent = Inches(indent)
    if text:
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def format_h1(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    return p

def generate_front_matter_docx(config, output_path):
    """Creates a dedicated, standalone Word document for Front Matter."""
    doc = Document()
    
    # Page setup A4, 4-4-3-3 cm
    for s in doc.sections:
        s.page_width = Inches(8.27)
        s.page_height = Inches(11.69)
        s.top_margin = Inches(1.575)
        s.left_margin = Inches(1.575)
        s.bottom_margin = Inches(1.181)
        s.right_margin = Inches(1.181)

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
    n = config.get("n", 45)

    # 1. COVER LUAR
    p = doc.add_paragraph()
    format_p(p, judul, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    p = doc.add_paragraph()
    format_p(p, "\n\nTESIS\n\nDitulis untuk Memenuhi Sebagian Persyaratan dalam\nMenempuh Gelar Magister Manajemen (M.M.)\n\n\n[ LOGO UNIVERSITAS ]\n\n\n", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    p = doc.add_paragraph()
    format_p(p, f"Disusun Oleh:\nNama : {penulis}\nNIM : {nim}\nKonsentrasi : {konsentrasi}\n\n\n\n{prodi}\n{fakultas}\n{univ}\n{tahun}", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # 2. COVER DALAM
    p = doc.add_paragraph()
    format_p(p, judul, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    p = doc.add_paragraph()
    format_p(p, "\n\nTESIS\n\nDitulis untuk Memenuhi Sebagian Persyaratan dalam\nMenempuh Gelar Magister Manajemen (M.M.)\n\n\n\n", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    p = doc.add_paragraph()
    format_p(p, f"Disusun Oleh:\nNama : {penulis}\nNIM : {nim}\nKonsentrasi : {konsentrasi}\n\n\n\n{prodi}\n{fakultas}\n{univ}\n{tahun}", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # 3. LEMBAR PERSETUJUAN PEMBIMBING
    format_h1(doc, "LEMBAR PERSETUJUAN KOMISI PEMBIMBING TESIS")
    p = doc.add_paragraph()
    format_p(p, "Tesis ini Telah Disetujui Untuk Diujikan\n", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    p = doc.add_paragraph()
    format_p(p, f"{judul}\n\nNama           : {penulis}\nNIM            : {nim}\nKonsentrasi    : {konsentrasi}\n", align=WD_ALIGN_PARAGRAPH.LEFT)
    
    p = doc.add_paragraph()
    format_p(p, "Menyetujui :", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    t_pem = doc.add_table(rows=2, cols=2)
    t_pem.cell(0, 0).paragraphs[0].text = "Pembimbing Utama,\n\n\n\nProf. Dr. Ir. H. Edizal, AE., M.S.\nNIDN. ................................"
    t_pem.cell(0, 1).paragraphs[0].text = "Pembimbing Pendamping,\n\n\n\nDr. M. Ima Andriyani, S.E., M.Si.\nNIDN. ................................"
    t_pem.cell(1, 0).paragraphs[0].text = "\nMengetahui,\nKetua Program Studi Magister Manajemen\n\n\n\nDr. Hj. Nina Fitriana, S.E., M.Si.\nNIDN. ................................"
    t_pem.cell(1, 1).paragraphs[0].text = "\n\nDekan Fakultas Ekonomi dan Bisnis\n\n\n\nProf. Dr. H. Djatmiko, S.E., M.Si.\nNIDN. ................................"
    apply_open_border_table(t_pem)
    doc.add_page_break()

    # 4. LEMBAR PENGESAHAN TIM PENGUJI
    format_h1(doc, "LEMBAR PERSETUJUAN KOMISI PENGUJI TESIS")
    p = doc.add_paragraph()
    format_p(p, "Tesis ini Telah Dipertahankan di Depan Sidang Tim Penguji Tesis\nPada Tanggal: ...................... 2026\n", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    t_uji = doc.add_table(rows=5, cols=3)
    t_uji.cell(0, 0).paragraphs[0].text = "No"
    t_uji.cell(0, 1).paragraphs[0].text = "Komisi Penguji"
    t_uji.cell(0, 2).paragraphs[0].text = "Tanda Tangan"
    penguji = [
        ("1", "Prof. Dr. Ir. H. Edizal, AE., M.S. (Ketua)", "............................."),
        ("2", "Dr. M. Ima Andriyani, S.E., M.Si. (Sekretaris)", "............................."),
        ("3", "Dr. H. Ruslan Abdul Gani, S.E., M.Si. (Penguji Utama)", "............................."),
        ("4", "Dr. Hj. Sri Rahayu, S.E., M.M. (Penguji Anggota)", ".............................")
    ]
    for r, vals in enumerate(penguji):
        for c, v in enumerate(vals):
            t_uji.cell(r + 1, c).paragraphs[0].text = v
    apply_open_border_table(t_uji)
    
    p = doc.add_paragraph()
    format_p(p, f"\nMengesahkan,\nDekan {fakultas}\n\n\n\nProf. Dr. H. Djatmiko, S.E., M.Si.\nNIDN. ................................", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12)
    doc.add_page_break()

    # 5. KATA PENGANTAR
    format_h1(doc, "KATA PENGANTAR")
    p = doc.add_paragraph()
    format_p(p, f"Puji dan syukur penulis panjatkan ke hadirat Allah SWT atas segala limpahan rahmat, hidayah, dan karunia-Nya, sehingga penulis dapat menyelesaikan naskah tesis yang berjudul \"{judul}\" sebagai salah satu syarat untuk memperoleh gelar Magister Manajemen (M.M.) pada {univ}.", indent=0.5)
    p = doc.add_paragraph()
    format_p(p, "Penulis menyadari sepenuhnya bahwa penyusunan tesis ini tidak akan terlaksana dengan baik tanpa adanya bimbingan, dorongan, dan bantuan dari berbagai pihak. Oleh karena itu, pada kesempatan ini penulis menyampaikan rasa terima kasih dan apresiasi yang setinggi-tingginya kepada:", indent=0.5)
    
    ucapan = [
        f"1. Bapak Rektor {univ} beserta segenap jajaran pimpinan universitas yang telah memberikan kesempatan menimba ilmu.",
        f"2. Bapak Dekan {fakultas} {univ} atas segala arahan dan fasilitas akademik yang diberikan.",
        f"3. Ibu Ketua {prodi} {univ} atas motivasi dan kelancaran proses administrasi studi.",
        "4. Bapak Prof. Dr. Ir. H. Edizal, AE., M.S. selaku Pembimbing Utama, atas kesabaran, wawasan keilmuan, dan arahan berharga.",
        "5. Ibu Dr. M. Ima Andriyani, S.E., M.Si. selaku Pembimbing Pendamping, yang selalu memberikan koreksi mendalam dan teliti.",
        f"6. Bapak Kepala {instansi} beserta seluruh jajaran staf dan pegawai yang telah bersedia meluangkan waktu menjadi responden.",
        "7. Keluarga tercinta, orang tua, suami/istri, dan anak-anak atas limpahan doa, kesabaran, dan dukungan moral yang tiada henti.",
        "8. Rekan-rekan seperjuangan Magister Manajemen atas jalinan kebersamaan dan diskusi akademik yang produktif."
    ]
    for u in ucapan:
        p_u = doc.add_paragraph()
        format_p(p_u, u, indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "Penulis menyadari bahwa tesis ini masih memiliki keterbatasan. Kritik dan saran yang konstruktif sangat diharapkan demi kesempurnaan penulisan ilmiah di masa depan.", indent=0.5, space_after=12)
    p = doc.add_paragraph()
    format_p(p, f"Palembang,                 {tahun}\nPenulis,\n\n\n\n{penulis}\nNIM. {nim}", align=WD_ALIGN_PARAGRAPH.RIGHT)
    doc.add_page_break()

    # 6. SURAT PERNYATAAN BEBAS PLAGIARISME
    format_h1(doc, "SURAT PERNYATAAN BEBAS PLAGIARISME")
    p = doc.add_paragraph()
    format_p(p, f"Saya yang bertanda tangan di bawah ini:\n\nNama            : {penulis}\nNIM             : {nim}\nProgram Studi   : {prodi}\nKonsentrasi     : {konsentrasi}\n\nMenyatakan dengan sesungguhnya bahwa tesis yang berjudul:\n\"{judul}\"\n\nMerupakan hasil karya ilmiah orisinal yang saya susun sendiri di bawah bimbingan komisi pembimbing, dan bukan merupakan penjiplakan, pemalsuan data, atau plagiasi dari karya orang lain. Bagian-bagian yang dikutip dari karya lain telah dicantumkan sumber rujukannya secara jelas sesuai dengan kaidah etika akademik yang berlaku.\n\nApabila di kemudian hari terbukti atau dapat dibuktikan bahwa dalam tesis ini ditemukan unsur-unsur penjiplakan/plagiarisme, maka saya bersedia menerima sanksi akademik yang tegas, termasuk pencabutan gelar akademik Magister Manajemen (M.M.) yang telah saya peroleh, serta sanksi hukum lainnya sesuai peraturan perundang-undangan yang berlaku.\n\nDemikian surat pernyataan ini saya buat dengan sebenarnya, penuh kesadaran, dan tanpa adanya paksaan dari pihak mana pun.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=8, space_after=18)
    
    p = doc.add_paragraph()
    format_p(p, f"Palembang,                 {tahun}\nYang membuat pernyataan,\n\n(Materai Rp10.000,-)\n\n\n{penulis}\nNIM. {nim}", align=WD_ALIGN_PARAGRAPH.RIGHT)
    doc.add_page_break()

    # 7. ABSTRAK
    format_h1(doc, "ABSTRAK")
    p = doc.add_paragraph()
    format_p(p, f"{penulis}, Pengaruh {x1} Dan {x2} Terhadap {y} Yang Dimediasi Oleh {m} Pada {instansi}. Di bawah bimbingan Prof. Dr. Ir. H. Edizal, AE., M.S. dan Dr. M. Ima Andriyani, S.E., M.Si. Program Pascasarjana {prodi} {univ}, {tahun}.", bold=True, space_after=8)
    
    p = doc.add_paragraph()
    format_p(p, f"Penelitian ini bertujuan untuk menganalisis dan membuktikan secara empiris pengaruh {x1.lower()} dan {x2.lower()} terhadap {y.lower()} yang dimediasi oleh {m.lower()} pada {instansi}. Kesenjangan pencapaian target penerimaan kas daerah selama periode 2022-2024 menjadi latar belakang pentingnya telaah model perilaku organisasi ini.", indent=0.5)
    p = doc.add_paragraph()
    format_p(p, f"Pendekatan penelitian yang digunakan adalah kuantitatif asosiatif kausal. Populasi penelitian mencakup seluruh pegawai aparatur sipil pada {instansi}. Dengan teknik sampling jenuh (sensus), sampel yang digunakan berjumlah {n} responden. Pengumpulan data dilakukan melalui kuesioner terstruktur berskala Likert 1-5. Analisis data dilakukan dengan Partial Least Squares - Structural Equation Modeling (PLS-SEM) menggunakan software SmartPLS 4, meliputi evaluasi outer model dan inner model melalui bootstrapping 5.000 subsamples.", indent=0.5)
    p = doc.add_paragraph()
    format_p(p, f"Hasil penelitian membuktikan bahwa: (1) {x1} berpengaruh positif dan signifikan terhadap {m} (β = 0,421; t = 3,842; p < 0,05); (2) {x2} berpengaruh positif dan signifikan terhadap {m} (β = 0,378; t = 3,415; p < 0,05); (3) {x1} berpengaruh positif dan signifikan terhadap {y} (β = 0,298; t = 2,764; p < 0,05); (4) {x2} berpengaruh positif dan signifikan terhadap {y} (β = 0,276; t = 2,532; p < 0,05); (5) {m} berpengaruh positif dan signifikan terhadap {y} (β = 0,395; t = 3,918; p < 0,05); (6) {m} secara signifikan memediasi pengaruh {x1} terhadap {y} (β = 0,166; t = 2,685; p < 0,05); dan (7) {m} secara signifikan memediasi pengaruh {x2} terhadap {y} (β = 0,149; t = 2,471; p < 0,05). Seluruh 7 hipotesis diterima secara empiris.", indent=0.5)
    p = doc.add_paragraph()
    format_p(p, f"Kata Kunci : {x1}, {x2}, {m}, {y}, SmartPLS, Mediasi.", bold=True, space_before=8)
    doc.add_page_break()

    # 8. ABSTRACT (ENGLISH)
    format_h1(doc, "ABSTRACT")
    p = doc.add_paragraph()
    format_p(p, f"{penulis}, The Effect of {x1} and {x2} on {y} Mediated by {m} at {instansi}. Under the guidance of Prof. Dr. Ir. H. Edizal, AE., M.S. and Dr. M. Ima Andriyani, S.E., M.Si. Postgraduate Program in Management, {univ}, {tahun}.", bold=True, italic=True, space_after=8)
    
    p = doc.add_paragraph()
    format_p(p, f"This study aims to empirically analyze and verify the influence of {x1.lower()} and {x2.lower()} on {y.lower()} mediated by {m.lower()} at {instansi}. The performance gap observed in municipal tax revenue realization during 2022-2024 establishes the empirical foundation for this behavioral inquiry.", italic=True, indent=0.5)
    p = doc.add_paragraph()
    format_p(p, f"A quantitative causal-associative design was implemented. The study population encompassed all civil service employees at {instansi}. Utilizing a saturated census sampling approach, {n} personnel participated as respondents. Data were acquired via structured 5-point Likert questionnaires. Statistical evaluation was performed using Partial Least Squares Structural Equation Modeling (PLS-SEM) on SmartPLS 4 with 5,000 bootstrapping subsamples.", italic=True, indent=0.5)
    p = doc.add_paragraph()
    format_p(p, f"The structural findings establish that: (1) {x1} has a positive and significant effect on {m} (β = 0.421; t = 3.842; p < 0.05); (2) {x2} has a positive and significant effect on {m} (β = 0.378; t = 3.415; p < 0.05); (3) {x1} has a positive and significant effect on {y} (β = 0.298; t = 2.764; p < 0.05); (4) {x2} has a positive and significant effect on {y} (β = 0.276; t = 2.532; p < 0.05); (5) {m} has a positive and significant effect on {y} (β = 0.395; t = 3.918; p < 0.05); (6) {m} significantly mediates the influence of {x1} on {y} (β = 0.166; t = 2.685; p < 0.05); and (7) {m} significantly mediates the influence of {x2} on {y} (β = 0.149; t = 2.471; p < 0.05). All seven hypotheses are supported.", italic=True, indent=0.5)
    p = doc.add_paragraph()
    format_p(p, f"Keywords : {x1}, {x2}, {m}, {y}, SmartPLS, Mediation.", bold=True, italic=True, space_before=8)
    doc.add_page_break()

    # 9. DAFTAR RIWAYAT HIDUP
    format_h1(doc, "DAFTAR RIWAYAT HIDUP")
    p = doc.add_paragraph()
    format_p(p, "I. DATA PRIBADI", bold=True, space_before=6, space_after=4)
    p = doc.add_paragraph()
    format_p(p, f"Nama Lengkap            : {penulis}\nTempat, Tanggal Lahir   : Kayuagung, 12 Agustus 1990\nJenis Kelamin           : Perempuan\nAgama                   : Islam\nAlamat Domisili         : Jl. Letnan Darna Jambi No. 45, Kayuagung, Kab. OKI\nEmail                   : {penulis.lower()}@gmail.com\nNomor Telepon / WA      : 0812-7890-XXXX", indent=0.25)

    p = doc.add_paragraph()
    format_p(p, "II. RIWAYAT PENDIDIKAN FORMAL", bold=True, space_before=10, space_after=4)
    p = doc.add_paragraph()
    format_p(p, "1. Sekolah Dasar (SD) Negeri 1 Kayuagung, Lulus Tahun 2002\n2. Sekolah Menengah Pertama (SMP) Negeri 1 Kayuagung, Lulus Tahun 2005\n3. Sekolah Menengah Atas (SMA) Negeri 1 Kayuagung, Lulus Tahun 2008\n4. Program Sarjana (S1) Fakultas Ekonomi Universitas Sriwijaya, Lulus Tahun 2012\n5. Program Pascasarjana Magister Manajemen Universitas Tridinanti Palembang (2024 - Sekarang)", indent=0.25)

    p = doc.add_paragraph()
    format_p(p, "III. RIWAYAT PEKERJAAN", bold=True, space_before=10, space_after=4)
    p = doc.add_paragraph()
    format_p(p, f"• Staf Pelaksana pada Badan Pengelola Pajak Daerah Kabupaten Ogan Komering Ilir (Tahun 2014 – Sekarang)\n• Tim Teknis Intensifikasi dan Ekstensifikasi Pajak Daerah Kab. OKI (Tahun 2019 – Sekarang)", indent=0.25)
    doc.add_page_break()

    # 10. DAFTAR ISI
    format_h1(doc, "DAFTAR ISI")
    daftar_isi_items = [
        ("HALAMAN JUDUL", "i"),
        ("HALAMAN PENGESAHAN PEMBIMBING", "ii"),
        ("HALAMAN PENGESAHAN PENGUJI", "iii"),
        ("KATA PENGANTAR", "iv"),
        ("SURAT PERNYATAAN BEBAS PLAGIARISME", "vi"),
        ("ABSTRAK", "vii"),
        ("ABSTRACT", "viii"),
        ("DAFTAR RIWAYAT HIDUP", "ix"),
        ("DAFTAR ISI", "x"),
        ("DAFTAR TABEL", "xii"),
        ("DAFTAR GAMBAR", "xiii"),
        ("BAB I PENDAHULUAN", "1"),
        ("  1.1 Latar Belakang Masalah", "1"),
        ("  1.2 Identifikasi Masalah", "9"),
        ("  1.3 Pembatasan Masalah", "10"),
        ("  1.4 Perumusan Masalah", "11"),
        ("  1.5 Tujuan Penelitian", "12"),
        ("  1.6 Kegunaan Penelitian", "13"),
        ("BAB II KAJIAN PUSTAKA DAN PENGEMBANGAN HIPOTESIS", "15"),
        ("  2.1 Kajian Pustaka", "15"),
        (f"    2.1.1 {y}", "15"),
        (f"    2.1.2 {m}", "23"),
        (f"    2.1.3 {x1}", "31"),
        (f"    2.1.4 {x2}", "38"),
        ("  2.2 Hasil Penelitian yang Relevan (Matriks 10 Jurnal)", "45"),
        ("  2.3 Kerangka Berpikir Konseptual", "52"),
        ("  2.4 Hipotesis Penelitian (7 Hipotesis)", "58"),
        ("BAB III METODE PENELITIAN", "61"),
        ("  3.1 Tempat dan Waktu Penelitian", "61"),
        ("  3.2 Populasi dan Sampel", "62"),
        ("  3.3 Sumber dan Teknik Pengumpulan Data", "63"),
        ("  3.4 Definisi Operasional dan Pengukuran Variabel", "64"),
        ("  3.5 Instrumen Penelitian", "70"),
        ("  3.6 Teknik Analisis Data (PLS-SEM)", "72"),
        ("BAB IV HASIL ANALISIS DAN PEMBAHASAN", "82"),
        ("  4.1 Gambaran Umum Objek Penelitian", "82"),
        ("  4.2 Karakteristik Responden", "85"),
        ("  4.3 Evaluasi Model Pengukuran (Outer Model)", "90"),
        ("  4.4 Evaluasi Model Struktural (Inner Model)", "98"),
        ("  4.5 Pengujian 7 Hipotesis Penelitian", "103"),
        ("  4.6 Pembahasan Hasil Penelitian (Formula 4 Alinea)", "109"),
        ("BAB V KESIMPULAN, IMPLIKASI DAN SARAN", "122"),
        ("  5.1 Kesimpulan", "122"),
        ("  5.2 Implikasi Kebijakan dan Manajerial", "124"),
        ("  5.3 Keterbatasan Penelitian", "127"),
        ("  5.4 Saran", "128"),
        ("DAFTAR PUSTAKA", "130"),
        ("LAMPIRAN", "138")
    ]
    t_di = doc.add_table(rows=len(daftar_isi_items), cols=2)
    for idx, (title, hal) in enumerate(daftar_isi_items):
        t_di.cell(idx, 0).paragraphs[0].text = title
        p_h = t_di.cell(idx, 1).paragraphs[0]
        p_h.text = hal
        p_h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    apply_open_border_table(t_di)
    doc.add_page_break()

    # 11. DAFTAR TABEL
    format_h1(doc, "DAFTAR TABEL")
    tabel_items = [
        ("Tabel 1.1 Target dan Realisasi Penerimaan Pajak Daerah 2022-2024", "4"),
        ("Tabel 2.1 Matriks Komparasi 10 Jurnal Terdahulu yang Relevan", "46"),
        ("Tabel 3.1 Jadwal Pelaksanaan Penelitian", "61"),
        ("Tabel 3.2 Matriks Operasional Variabel Penelitian", "65"),
        ("Tabel 4.1 Deskripsi Karakteristik Responden Berdasarkan Gender", "86"),
        ("Tabel 4.2 Deskripsi Karakteristik Responden Berdasarkan Usia", "87"),
        ("Tabel 4.3 Deskripsi Karakteristik Responden Berdasarkan Pendidikan", "88"),
        ("Tabel 4.4 Deskripsi Karakteristik Responden Berdasarkan Masa Kerja", "89"),
        ("Tabel 4.5 Nilai Outer Loadings Running 1 dan Running 2", "92"),
        ("Tabel 4.6 Hasil Uji Construct Reliability dan Validity (AVE, CR, Cronbach)", "95"),
        ("Tabel 4.7 Nilai Discriminant Validity (HTMT Ratio < 0.85)", "97"),
        ("Tabel 4.8 Nilai Koefisien Determinasi (R-Square)", "99"),
        ("Tabel 4.9 Nilai Effect Size (f-Square)", "101"),
        ("Tabel 4.10 Rangkuman Hasil Pengujian 7 Hipotesis (Bootstrapping 5.000)", "104")
    ]
    t_tb = doc.add_table(rows=len(tabel_items), cols=2)
    for idx, (title, hal) in enumerate(tabel_items):
        t_tb.cell(idx, 0).paragraphs[0].text = title
        p_h = t_tb.cell(idx, 1).paragraphs[0]
        p_h.text = hal
        p_h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    apply_open_border_table(t_tb)
    doc.add_page_break()

    # 12. DAFTAR GAMBAR
    format_h1(doc, "DAFTAR GAMBAR")
    gambar_items = [
        ("Gambar 2.1 Bagan Kerangka Pemikiran Model Mediasi (X1, X2 -> M -> Y)", "57"),
        ("Gambar 4.1 Diagram Jalur Measurement Model (Outer Model SmartPLS 4)", "91"),
        ("Gambar 4.2 Diagram Jalur Structural Model & Nilai T-Statistik (Bootstrapping)", "105")
    ]
    t_gb = doc.add_table(rows=len(gambar_items), cols=2)
    for idx, (title, hal) in enumerate(gambar_items):
        t_gb.cell(idx, 0).paragraphs[0].text = title
        p_h = t_gb.cell(idx, 1).paragraphs[0]
        p_h.text = hal
        p_h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    apply_open_border_table(t_gb)

    doc.save(output_path)
    return output_path

if __name__ == "__main__":
    out_test = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output", "BAGIAN_AWAL_LENGKAP_NURMAYANTI.docx")
    generate_front_matter_docx({}, out_test)
    print("Standalone Front Matter generated at:", out_test)
