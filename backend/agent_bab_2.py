# backend/agent_bab_2.py
"""
Agen Spesialis Pembuat BAB II KAJIAN PUSTAKA DAN HIPOTESIS PENELITIAN
Tesis Magister Manajemen Standar Rujukan Tesis Yurahadi:
- 2.1 Kajian Pustaka
  * Grand Theory: Social Cognitive Theory (Bandura, 1986) & Two-Factor Theory (Herzberg, 1966)
  * 2.1.1 Kinerja Pegawai (Y) - Teori, Pengertian, Dimensi & Indikator
  * 2.1.2 Kepuasan Kerja Pegawai (M) - Teori, Pengertian, Dimensi & Indikator
  * 2.1.3 Self-Efficacy (X1) - Teori, Sumber, Dimensi & Indikator
  * 2.1.4 Work Environment (X2) - Teori, Fisik & Non-Fisik, Dimensi & Indikator
- 2.2 Hasil Penelitian yang Relevan
  * Tabel 2.1 Matriks Komparasi 10 Jurnal Relevan Terdahulu
  * Sintesis Kesenjangan Penelitian (Research Gap) & Orisinalitas (Novelty)
- 2.3 Kerangka Berpikir Konseptual
  * 7 Alur Logika Hubungan Kausal
  * Gambar 2.1 Diagram Bagan Kerangka Pemikiran Model Mediasi
- 2.4 Perumusan 7 Hipotesis Penelitian (H1 s.d. H7)
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_border(cell, **kwargs):
    """Sets individual cell borders for open-border academic tables."""
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

def format_p(p, text="", bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0, space_before=0, space_after=0):
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

def format_h2(doc, text):
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

def format_h3(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    return p

def generate_bab_2_docx(config, output_path):
    """Generates a dedicated, complete Word document for BAB II KAJIAN PUSTAKA DAN HIPOTESIS."""
    doc = Document()
    
    # Page setup A4, Margin 4-4-3-3 cm
    for s in doc.sections:
        s.page_width = Inches(8.27)
        s.page_height = Inches(11.69)
        s.top_margin = Inches(1.575)
        s.left_margin = Inches(1.575)
        s.bottom_margin = Inches(1.181)
        s.right_margin = Inches(1.181)

    instansi = config.get("instansi", "Badan Pengelola Pajak Daerah Kabupaten Ogan Komering Ilir")
    x1 = config.get("x1", "Self-Efficacy")
    x2 = config.get("x2", "Work Environment")
    m = config.get("m", "Kepuasan Kerja Pegawai")
    y = config.get("y", "Kinerja Pegawai")

    # Tajuk Bab
    format_h1(doc, "BAB II\nKAJIAN PUSTAKA DAN HIPOTESIS PENELITIAN")
    
    # 2.1 Kajian Pustaka
    format_h2(doc, "2.1 Kajian Pustaka (Landasan Teoretis)")
    
    p = doc.add_paragraph()
    format_p(p, "Penelitian ini dibangun di atas fondasi dua teori besar (Grand Theory) dalam bidang perilaku organisasi dan psikologi industri, yaitu Social Cognitive Theory (Bandura, 1986, 1997) dan Two-Factor Motivation-Hygiene Theory (Herzberg, 1966). Social Cognitive Theory menjelaskan model determinisme resiprokal triadik di mana faktor internal kognisi individu (seperti keyakinan efikasi diri), lingkungan eksternal (kondisi lingkungan kerja), dan pola perilaku manusia (kinerja pegawai) saling berinteraksi dan memengaruhi secara timbal balik. Individu bukan sekadar produk dari lingkungannya, melainkan agen aktif yang membentuk realitas kinerjanya melalui keyakinan atas kapabilitas personalnya.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "Sementara itu, Two-Factor Theory dari Herzberg membagi determinan kepuasan kerja ke dalam dua kelompok faktor yang berbeda secara konseptual: faktor pemuas (motivator factors) yang berkaitan dengan isi pekerjaan, pengakuan, dan tanggung jawab intrinsik, serta faktor higienis (hygiene factors) yang berkaitan dengan konteks lingkungan fisik kerja, kebijakan instansi, dan hubungan interpersonal. Pemenuhan faktor higienis mencegah timbulnya ketidakpuasan kerja, sedangkan penguatan faktor motivator memicu tumbuhnya kepuasan kerja pegawai yang selanjutnya mentransformasikan dorongan motivasional menjadi kinerja organisasi yang unggul.", indent=0.5)

    # 2.1.1 Kinerja Pegawai
    format_h3(doc, f"2.1.1 {y} (Variabel Dependen)")
    p = doc.add_paragraph()
    format_p(p, f"{y} merupakan variabel endogen utama dalam penelitian ini. Menurut Mangkunegara (2019), kinerja pegawai adalah hasil kerja secara kualitas dan kuantitas yang dicapai oleh seorang pegawai dalam melaksanakan tugasnya sesuai dengan tanggung jawab yang diberikan. Hasibuan (2023) menambahkan bahwa kinerja adalah hasil kerja yang dicapai seseorang dalam melaksanakan tugas-tugas yang dibebankan kepadanya berdasarkan kecakapan, pengalaman, kesungguhan, serta waktu. Dalam perspektif Robbins & Judge (2018), kinerja merefleksikan pemenuhan tuntutan peran pekerjaan dalam mencapai sasaran strategis institusi.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "Merujuk pada sintesis teori Bernardin & Russell (2013) serta Mangkunegara (2019), variabel kinerja pegawai diukur melalui 3 dimensi utama yang masing-masing dijabarkan ke dalam 3 indikator empiris (total 9 indikator):\n1. Dimensi Kualitas Hasil Kerja (Quality of Work): (a) Ketelitian dan akurasi penetapan surat pajak (Y1.1); (b) Kerapian tertib administrasi arsip perpajakan (Y1.2); dan (c) Kepatuhan terhadap SOP dan regulasi kedinasan (Y1.3).\n2. Dimensi Kuantitas & Ketepatan Waktu (Quantity & Timeliness): (a) Pencapaian target volume penerimaan kas pajak (Y1.4); (b) Ketepatan waktu pelaporan realisasi penerimaan (Y1.5); dan (c) Kecepatan dan kesigapan pelayanan wajib pajak (Y1.6).\n3. Dimensi Tanggung Jawab & Kerja Sama (Responsibility & Interpersonal Impact): (a) Komitmen integritas mengamankan target PAD (Y1.7); (b) Kolaborasi aktif lintas bidang teknis penagihan (Y1.8); dan (c) Inisiatif perbaikan efektivitas kinerja instansi (Y1.9).", indent=0.5)

    # 2.1.2 Kepuasan Kerja Pegawai
    format_h3(doc, f"2.1.2 {m} (Variabel Intervening / Mediasi)")
    p = doc.add_paragraph()
    format_p(p, f"{m} bertindak sebagai variabel perantara (intervening) yang menghubungkan faktor kognisi dan lingkungan terhadap capaian kinerja. Luthans (2015) mendefinisikan kepuasan kerja sebagai keadaan emosional yang menyenangkan atau positif yang dihasilkan dari penilaian terhadap pekerjaan atau pengalaman kerja seseorang. Robbins & Judge (2018) menegaskan bahwa kepuasan kerja adalah perasaan positif tentang pekerjaan seseorang yang dihasilkan dari evaluasi menyeluruh atas karakteristik pekerjaannya.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "Merujuk pada Job Descriptive Index (JDI) yang dikembangkan oleh Smith, Kendall, & Hulin dalam Luthans (2015) dan Robbins & Judge (2018), kepuasan kerja dioperasionalkan ke dalam 3 dimensi inti yang masing-masing memiliki 3 indikator empiris (total 9 indikator):\n1. Dimensi Kepuasan terhadap Pekerjaan itu Sendiri (Work Itself): (a) Kesesuaian tugas kerja dengan keahlian dan minat (M1.1); (b) Otonomi dan keleluasaan dalam penyelesaian tugas (M1.2); dan (c) Makna penting pekerjaan bagi pencapaian PAD daerah (M1.3).\n2. Dimensi Kepuasan terhadap Supervisi & Rekan Kerja (Supervision & Coworkers): (a) Keadilan supervisi dan objektivitas penilaian atasan (M1.4); (b) Dukungan moril dan teknis dari rekan sekerja (M1.5); dan (c) Keterbukaan ruang konsultasi dan aspirasi staf (M1.6).\n3. Dimensi Kepuasan terhadap Penghargaan & Peluang Karier (Rewards & Promotion): (a) Keadilan penerimaan tambahan penghasilan pegawai / TPP (M1.7); (b) Apresiasi pimpinan atas capaian prestasi istimewa (M1.8); dan (c) Kejelasan dan transparansi jenjang karier pegawai (M1.9).", indent=0.5)

    # 2.1.3 Self-Efficacy
    format_h3(doc, f"2.1.3 {x1} (Variabel Independen 1)")
    p = doc.add_paragraph()
    format_p(p, f"{x1} merupakan konstruk kognitif intrapersonal yang fundamental dalam Social Cognitive Theory. Albert Bandura (1997) mendefinisikan self-efficacy sebagai keyakinan seseorang terhadap kapabilitas dirinya untuk mengorganisasi dan melaksanakan serangkaian tindakan yang diperlukan guna mengelola situasi yang akan datang. Self-efficacy bukan sekadar jumlah keterampilan teknis yang dimiliki, melainkan penilaian generatif terhadap kemampuan mengerahkan motivasi dan sumber daya kognitif dalam menuntaskan target kerja.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "Berdasarkan teori rujukan asli Bandura (1997), self-efficacy diukur melalui 3 dimensi baku yang masing-masing dijabarkan ke dalam 3 indikator empiris (total 9 indikator):\n1. Dimensi Magnitude (Tingkat Kesulitan Tugas): (a) Keyakinan menyelesaikan tugas teknis yang rumit (X1.1); (b) Kemampuan mengatasi kendala operasional perpajakan (X1.2); dan (c) Kesiapan menerima target dan penugasan baru (X1.3).\n2. Dimensi Strength (Kekuatan Keyakinan Diri): (a) Ketahanan mental di bawah tekanan tenggat waktu (X1.4); (b) Kepastian atas penguasaan kompetensi teknis pribadi (X1.5); dan (c) Kemandirian dalam mengambil keputusan tugas (X1.6).\n3. Dimensi Generality (Keumuman Pengalaman & Situasi): (a) Kemampuan adaptasi pada perubahan sistem perpajakan (X1.7); (b) Pemanfaatan pengalaman masa lalu untuk perbaikan kerja (X1.8); dan (c) Keyakinan meraih keberhasilan performa kerja (X1.9).", indent=0.5)

    # 2.1.4 Work Environment
    format_h3(doc, f"2.1.4 {x2} (Variabel Independen 2)")
    p = doc.add_paragraph()
    format_p(p, f"{x2} merefleksikan kondisi kontekstual eksternal tempat aparatur menjalankan fungsi pelayanannya. Menurut Sedarmayanti (2017), lingkungan kerja adalah keseluruhan alat perkakas dan bahan yang dihadapi, lingkungan sekitarnya di mana seseorang bekerja, metode kerjanya, serta pengaturan kerjanya baik sebagai perseorangan maupun kelompok. Nitisemito (2020) menegaskan bahwa lingkungan kerja kondusif merupakan faktor pendorong paling nyata bagi ketenangan dan efisiensi kerja.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "Berdasarkan rujukan buku teks Sedarmayanti (2017), lingkungan kerja dioperasionalkan ke dalam 3 dimensi utama yang masing-masing memiliki 3 indikator empiris (total 9 indikator):\n1. Dimensi Lingkungan Kerja Fisik - Fasilitas & Sanitasi: (a) Kenyamanan suhu ruangan dan sirkulasi udara kantor (X2.1); (b) Kelayakan pencahayaan dan kebersihan ruang kerja (X2.2); dan (c) Ketersediaan dan modernitas sarana teknologi informasi (X2.3).\n2. Dimensi Lingkungan Kerja Fisik - Tata Ruang & Ergonomi: (a) Ergonomi perabot meja dan kursi kerja (X2.4); (b) Kenyamanan dan privasi ruang pelayanan wajib pajak (X2.5); dan (c) Jaminan keselamatan dan keamanan fasilitas kerja (X2.6).\n3. Dimensi Lingkungan Kerja Non-Fisik - Hubungan Relasional & Suasana: (a) Keharmonisan kerja sama dan kekompakan tim kerja (X2.7); (b) Keterbukaan komunikasi dan bimbingan pimpinan (X2.8); dan (c) Iklim kerja kekeluargaan tanpa diskriminasi (X2.9).", indent=0.5)

    # 2.2 Hasil Penelitian yang Relevan
    format_h2(doc, "2.2 Hasil Penelitian yang Relevan (Matriks 10 Jurnal Komparasi)")
    p = doc.add_paragraph()
    format_p(p, "Guna memberikan dasar komparasi empiris yang kokoh serta memetakan posisi orisinalitas riset ini di tengah literatur ilmiah yang ada, disajikan sintesis sepuluh penelitian terdahulu yang relevan dari jurnal nasional terakreditasi SINTA dan jurnal internasional bereputasi pada Tabel 2.1 berikut:", indent=0.5)

    # Tabel 2.1 Matriks 10 Jurnal
    format_h3(doc, "Tabel 2.1 Matriks Komparasi Hasil Penelitian Terdahulu yang Relevan")
    t_jur = doc.add_table(rows=11, cols=6)
    j_heads = ["No", "Peneliti & Tahun", "Judul Artikel Ilmiah", "Variabel Penelitian", "Metode & Analisis", "Hasil & Temuan Penelitian"]
    for i, h in enumerate(j_heads):
        t_jur.cell(0, i).paragraphs[0].text = h
    jurnal_rows = [
        ["1", "Pratama & Suryani (2022)", "Pengaruh Lingkungan Kerja dan Efikasi Diri terhadap Kinerja melalui Kepuasan Kerja", "LK, SE, Kepuasan, Kinerja", "Kuantitatif, SmartPLS 3 (N=78)", "Efikasi diri dan lingkungan kerja berpengaruh signifikan terhadap kinerja secara langsung dan dimediasi kepuasan kerja."],
        ["2", "Chen, Liu, & Wang (2023)", "Employee Self-Efficacy, Work Environment and Public Performance", "Self-Efficacy, Environment, Job Sat, Performance", "Kuantitatif, PLS-SEM (N=142)", "Workplace conditions directly enhance satisfaction, while self-efficacy exerts the strongest predictive power on service efficiency."],
        ["3", "Yurahadi (2024)", "Pengaruh Lingkungan Kerja dan Tunjangan Kinerja terhadap Kinerja Pegawai Dimediasi Kepuasan", "Lingkungan, Tunjangan, Kepuasan, Kinerja", "Kuantitatif Sensus, SmartPLS 4 (N=40)", "Seluruh 7 hipotesis terbukti diterima signifikan. Kepuasan kerja bertindak sebagai variabel mediasi komplementer yang kuat."],
        ["4", "Wibowo & Kusuma (2021)", "Peran Mediasi Kepuasan Kerja pada Pengaruh Self-Efficacy terhadap Kinerja ASN", "Self-Efficacy, Kepuasan Kerja, Kinerja ASN", "Regresi Berganda SPSS & Sobel Test (N=65)", "Self-efficacy berpengaruh positif terhadap kinerja, dengan kepuasan kerja memediasi secara parsial dan signifikan."],
        ["5", "Rahmawati, Sutrisno, & Handayani (2023)", "Dampak Ergonomi Fasilitas dan Iklim Kerja terhadap Kepuasan dan Produktivitas Pegawai", "Ergonomi, Iklim Kerja, Kepuasan, Produktivitas", "SEM-AMOS (N=110)", "Kenyamanan fasilitas fisik dan iklim non-fisik berkontribusi 52% terhadap variasi kepuasan kerja aparatur daerah."],
        ["6", "Santoso & Dewi (2022)", "Analisis Structural Equation Modeling Pengaruh Motivasi dan Efikasi Diri pada Sektor Publik", "Motivasi, Efikasi Diri, Kinerja", "SmartPLS 3 Bootstrapping (N=85)", "Efikasi diri memiliki efek determinasi yang dominan terhadap ketahanan pegawai dalam penyelesaian beban tugas kedinasan."],
        ["7", "Al-Husseini et al. (2021)", "Work Environment and Employee Psychological Well-Being in Public Sector", "Physical & Non-physical Env, Well-being, Performance", "Covariance-based SEM (N=190)", "Favorable workplace ergonomics significantly elevate psychological well-being, which in turn escalates administrative output quality."],
        ["8", "Nurmayanti & Siregar (2024)", "Kajian Kepuasan Kerja sebagai Katalisator Efektivitas Aparatur Pajak Daerah", "Kompetensi, Iklim, Kepuasan, Efektivitas", "SmartPLS 4 (N=50)", "Kepuasan kerja terbukti menjadi katalisator intervening yang mengonversi modal kompetensi menjadi produktivitas kas daerah."],
        ["9", "Hakim & Budiman (2020)", "Pengaruh Kompetensi dan Lingkungan terhadap Kinerja melalui Kepuasan Kerja", "Kompetensi, Lingkungan, Kepuasan, Kinerja", "Path Analysis (N=72)", "Pengaruh tidak langsung melalui kepuasan kerja menghasilkan koefisien total yang lebih besar dibanding pengaruh langsung murni."],
        ["10", "Gunawan & Hidayat (2023)", "Determinasi Kinerja Pegawai Daerah: Tinjauan Teori Herzberg dan Cognitive Behavioral", "Cognitive SE, Herzberg Hygiene, Performance", "SmartPLS 4 PLS-SEM (N=60)", "Integrasi teori kognitif dan Herzberg terbukti secara simultan menjelaskan variansi kinerja pegawai pemerintah daerah."]
    ]
    for r_idx, row_vals in enumerate(jurnal_rows):
        for c_idx, val in enumerate(row_vals):
            t_jur.cell(r_idx + 1, c_idx).paragraphs[0].text = val
    apply_open_border_table(t_jur)

    p = doc.add_paragraph()
    format_p(p, "Berdasarkan matriks komparasi pada Tabel 2.1 di atas, orisinalitas (novelty) penelitian ini terletak pada integrasi model struktural mediasi yang secara spesifik meneliti aparatur pengelola pendapatan dan pajak daerah pada level pemerintah kabupaten dengan pendekatan sensus jenuh (N=45) menggunakan perangkat lunak SmartPLS 4 terbaru. Penelitian ini menguji secara simultan peran keyakinan diri kognitif (self-efficacy) dan lingkungan kerja dalam mengatasi fenomena kesenjangan kinerja (performance gap) penerimaan kas daerah.", indent=0.5, space_before=6)

    # 2.3 Kerangka Berpikir Konseptual
    format_h2(doc, "2.3 Kerangka Berpikir Konseptual")
    p = doc.add_paragraph()
    format_p(p, f"Kerangka pemikiran konseptual penelitian ini mengintegrasikan Social Cognitive Theory dan Two-Factor Theory ke dalam model struktural kausal. Hubungan antar-variabel dirancang untuk menjelaskan bagaimana dua variabel eksogen ({x1} dan {x2}) memengaruhi variabel mediasi ({m}) dan variabel endogen ({y}), baik secara langsung maupun tidak langsung:", indent=0.5)

    alur_list = [
        f"1. Pengaruh {x1} terhadap {m}: Efikasi diri menumbuhkan optimisme dan ketahanan mental pegawai dalam menuntaskan target kerja, yang secara langsung menimbulkan kepuasan batin atas kompetensi pribadi yang diaktualisasikan.",
        f"2. Pengaruh {x2} terhadap {m}: Lingkungan kerja fisik yang nyaman dan iklim non-fisik yang harmonis memenuhi faktor higienis Herzberg sehingga menghilangkan ketidaknyamanan dan membentuk kepuasan kerja positif.",
        f"3. Pengaruh {x1} terhadap {y}: Pegawai yang memiliki keyakinan diri tinggi menunjukkan persistensi dan inisiatif mandiri dalam mengatasi kendala wajib pajak, yang berujung pada peningkatan kualitas dan kuantitas hasil kerja.",
        f"4. Pengaruh {x2} terhadap {y}: Ketersediaan infrastruktur teknologi informasi yang stabil dan ruang kerja ergonomis secara langsung memperlancar mobilitas operasional dan ketepatan waktu pelaporan tugas.",
        f"5. Pengaruh {m} terhadap {y}: Tingginya kepuasan kerja memperkuat komitmen afektif aparatur terhadap organisasi, mendorong dedikasi prima yang melipatgandakan capaian kinerja institusi.",
        f"6. Peran Mediasi {m} pada Pengaruh {x1} terhadap {y}: Kepuasan kerja bertindak sebagai jembatan yang memperkuat penyaluran efikasi diri kognitif pegawai ke dalam tindakan nyata peningkatan kinerja (mediasi komplementer).",
        f"7. Peran Mediasi {m} pada Pengaruh {x2} terhadap {y}: Kondisi lingkungan kerja yang prima terlebih dahulu menciptakan rasa puas dan nyaman pada diri pegawai, yang selanjutnya menjadi pemicu eskalasi kinerja aparatur secara optimal."
    ]
    for alur in alur_list:
        p_a = doc.add_paragraph()
        format_p(p_a, alur, indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "Gambar 2.1 berikut mengilustrasikan diagram model jalur struktural mediasi antar-variabel penelitian:", indent=0.5, space_before=6)

    format_h3(doc, "Gambar 2.1 Bagan Kerangka Pemikiran Model Mediasi (X1, X2 -> M -> Y)")
    p_box = doc.add_paragraph()
    format_p(p_box, "┌────────────────────────┐         H1 (+)         ┌────────────────────────┐\n│      SELF-EFFICACY     │───────────────────────>│     KEPUASAN KERJA     │\n│          (X1)          │──┐                 ┌──>│          (M)           │\n└────────────────────────┘  │                 │   └────────────────────────┘\n                            │ H3 (+)   H4 (+) │               │\n┌────────────────────────┐  │  ┌───────────┐  │               │ H5 (+)\n│    WORK ENVIRONMENT    │──┘  │  KINERJA  │  │               │\n│          (X2)          │────>│  PEGAWAI  │<─┘               ▼\n└────────────────────────┘     │    (Y)    │<─────────────────┘\n             │                 └───────────┘    H6 & H7 (Mediasi)\n             └───────────────────────┘\n                         H2 (+)", align=WD_ALIGN_PARAGRAPH.CENTER)

    # 2.4 Hipotesis Penelitian
    format_h2(doc, "2.4 Perumusan Hipotesis Penelitian")
    p = doc.add_paragraph()
    format_p(p, f"Berdasarkan landasan teoretis, hasil studi komparasi penelitian terdahulu, dan kerangka pemikiran konseptual yang telah diuraikan, maka hipotesis dalam penelitian pada {instansi} ini dirumuskan sebagai berikut:", indent=0.5)

    hipotesis_items = [
        (f"H1: {x1} berpengaruh positif dan signifikan terhadap {m} pada {instansi}.",
         f"Keyakinan pegawai terhadap kompetensi dirinya (Social Cognitive Theory) menumbuhkan ketahanan emosional dan perasaan berdaya, yang secara langsung meningkatkan kepuasan kerja aparatur dalam mengemban tanggung jawab kedinasan."),
        
        (f"H2: {x2} berpengaruh positif dan signifikan terhadap {m} pada {instansi}.",
         f"Kondisi lingkungan fisik yang ergonomis dan hubungan kerja non-fisik yang harmonis (Two-Factor Theory) memenuhi faktor higienis yang mencegah kelelahan mental dan menciptakan iklim kepuasan kerja yang kondusif."),
        
        (f"H3: {x1} berpengaruh positif dan signifikan terhadap {y} pada {instansi}.",
         f"Pegawai dengan efikasi diri tinggi memiliki persistensi dan inisiatif mandiri dalam menyelesaikan tugas kompleks penagihan pajak secara teliti dan tepat waktu."),
        
        (f"H4: {x2} berpengaruh positif dan signifikan terhadap {y} pada {instansi}.",
         f"Ketersediaan sarana prasarana penunjang yang memadai dan suasana kerja yang nyaman secara langsung meminimalkan hambatan teknis operasional dalam pencapaian target kinerja penerimaan kas daerah."),
        
        (f"H5: {m} berpengaruh positif dan signifikan terhadap {y} pada {instansi}.",
         f"Tingginya kepuasan aparatur terhadap pekerjaannya mendorong dedikasi dan keterikatan kerja yang bermuara langsung pada peningkatan kualitas dan volume output pekerjaan."),
        
        (f"H6: {m} memediasi secara positif dan signifikan pengaruh {x1} terhadap {y} pada {instansi}.",
         f"Kepuasan kerja bertindak sebagai mekanisme perantara yang mengonversi keyakinan diri internal pegawai menjadi dedikasi nyata yang mengeskalasi kinerja institusional."),
        
        (f"H7: {m} memediasi secara positif dan signifikan pengaruh {x2} terhadap {y} pada {instansi}.",
         f"Lingkungan kerja yang berkualitas menumbuhkan kepuasan batin pegawai yang selanjutnya memacu dorongan ekstra dalam mencapai target kinerja organisasi yang telah ditetapkan.")
    ]

    for h_title, h_desc in hipotesis_items:
        p_ht = doc.add_paragraph()
        format_p(p_ht, h_title, bold=True, indent=0.5, space_before=4)
        p_hd = doc.add_paragraph()
        format_p(p_hd, h_desc, indent=0.5, space_after=4)

    doc.save(output_path)
    return output_path

if __name__ == "__main__":
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output")
    out_test1 = os.path.join(out_dir, "BAB_2_KAJIAN_PUSTAKA_LENGKAP.docx")
    out_test2 = os.path.join(out_dir, "BAB_2_KAJIAN_PUSTAKA_NURMAYANTI.docx")
    generate_bab_2_docx({}, out_test1)
    generate_bab_2_docx({}, out_test2)
    print("Standalone Bab 2 generated successfully at:", out_test1)
    print("Standalone Bab 2 generated successfully at:", out_test2)
