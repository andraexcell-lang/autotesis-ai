# backend/agent_bab_3.py
"""
Agen Spesialis Pembuat BAB III METODE PENELITIAN
Tesis Magister Manajemen Standar Rujukan Tesis Yurahadi:
- 3.1 Tempat dan Waktu Penelitian
  * Tabel 3.1 Jadwal Pelaksanaan Penelitian
- 3.2 Populasi dan Sampel (Sensus Sampling N=45)
- 3.3 Sumber dan Teknik Pengumpulan Data (Primer & Sekunder; Kuesioner, Observasi, Dokumentasi)
- 3.4 Variabel dan Definisi Operasional
  * Definisi Konseptual & Operasional (Y, M, X1, X2)
  * Tabel 3.2 Matriks Kisi-Kisi Operasional Variabel (16 Indikator Likert 1-5)
- 3.5 Instrumen Penelitian (Skala Likert 5 Poin & Kategori Interpretasi Skor)
- 3.6 Teknik Analisis Data (PLS-SEM SmartPLS 4: Outer Model & Inner Model)
- 3.7 Pengujian Hipotesis
  * Prosedur Bootstrapping 5.000 Subsamples
  * Tabel 3.3 Kriteria Pengambilan Keputusan Uji Hipotesis (T-Statistik & P-Value)
  * Prosedur Pengujian Efek Mediasi
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from engine_smartpls import INDICATOR_SPEC

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

def generate_bab_3_docx(config, output_path):
    """Generates a dedicated, complete Word document for BAB III METODE PENELITIAN."""
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
    n = config.get("n", 45)
    tahun = config.get("tahun", "2026")

    # Tajuk Bab
    format_h1(doc, "BAB III\nMETODE PENELITIAN")
    
    # 3.1 Tempat dan Waktu Penelitian
    format_h2(doc, "3.1 Tempat dan Waktu Penelitian")
    p = doc.add_paragraph()
    format_p(p, f"Penelitian ini dilaksanakan pada {instansi}. Objek penelitian dipilih secara purposif dengan pertimbangan bahwa instansi tersebut memegang peranan krusial sebagai penopang penerimaan Pendapatan Asli Daerah (PAD), serta memiliki fenomena kesenjangan kinerja operasional (performance gap) pada periode 2022 hingga 2024 yang relevan untuk diinvestigasi.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, f"Waktu pelaksanaan penelitian berlangsung selama enam bulan, terhitung mulai bulan Januari {tahun} sampai dengan bulan Juni {tahun}. Rincian tahapan kegiatan penelitian disajikan pada Tabel 3.1 berikut:", indent=0.5)

    # Tabel 3.1 Jadwal Pelaksanaan Penelitian
    format_h3(doc, "Tabel 3.1 Jadwal Pelaksanaan Penelitian")
    t_jadwal = doc.add_table(rows=7, cols=3)
    j_heads = ["No", "Tahapan Kegiatan Penelitian", "Alokasi Waktu Pelaksanaan"]
    for i, h in enumerate(j_heads):
        t_jadwal.cell(0, i).paragraphs[0].text = h
    jadwal_data = [
        ["1", "Tahap Pra-Survei dan Observasi Masalah Empiris", f"Januari {tahun}"],
        ["2", "Penyusunan dan Konsultasi Proposal Tesis", f"Februari {tahun}"],
        ["3", "Seminar Proposal dan Ujian Kualifikasi Riset", f"Maret {tahun}"],
        ["4", "Penyebaran Kuesioner dan Pengumpulan Data Responden", f"April {tahun}"],
        ["5", "Tabulasi Data dan Analisis Statistik SmartPLS 4", f"Mei {tahun}"],
        ["6", "Penyusunan Naskah Lengkap dan Ujian Sidang Tesis", f"Juni {tahun}"]
    ]
    for r_idx, row_vals in enumerate(jadwal_data):
        for c_idx, val in enumerate(row_vals):
            t_jadwal.cell(r_idx + 1, c_idx).paragraphs[0].text = val
    apply_open_border_table(t_jadwal)

    p = doc.add_paragraph()
    format_p(p, "Sumber: Rencana Kerja Penelitian Pascasarjana (Data Diolah Peneliti, 2026).", italic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=6)

    # 3.2 Populasi dan Sampel
    format_h2(doc, "3.2 Populasi dan Sampel")
    p = doc.add_paragraph()
    format_p(p, f"3.2.1 Populasi\nMenurut Sugiyono (2019), populasi adalah wilayah generalisasi yang terdiri atas objek/subjek yang mempunyai kuantitas dan karakteristik tertentu yang ditetapkan oleh peneliti untuk dipelajari dan kemudian ditarik kesimpulannya. Populasi dalam penelitian ini adalah seluruh aparatur sipil negara (ASN) yang bertugas aktif pada {instansi}, yaitu sebanyak {n} orang pegawai.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, f"3.2.2 Sampel\nSampel adalah bagian dari jumlah dan karakteristik yang dimiliki oleh populasi tersebut. Mengingat jumlah total populasi pegawai terjangkau dan berada di bawah 100 orang (N = {n} orang), maka teknik penarikan sampel yang digunakan dalam penelitian ini adalah Teknik Sampling Jenuh (Sensus) sebagaimana direkomendasikan oleh Sugiyono (2019) dan Arikunto (2019). Dengan menggunakan teknik sensus, seluruh populasi yang berjumlah {n} orang pegawai dijadikan sebagai responden penelitian. Penggunaan teknik sensus mengeliminasi kesalahan penarikan sampel (sampling error) serta memberikan daya representasi yang sempurna bagi pengujian model struktural pada instansi.", indent=0.5)

    # 3.3 Sumber dan Teknik Pengumpulan Data
    format_h2(doc, "3.3 Sumber dan Teknik Pengumpulan Data")
    p = doc.add_paragraph()
    format_p(p, "3.3.1 Sumber Data\nData yang digunakan dalam penelitian ini mencakup dua jenis sumber data:\n1. Data Primer: Data yang diperoleh secara langsung dari sumber pertama (responden) melalui pengisian instrumen kuesioner terstruktur berskala Likert 1-5 mengenai persepsi efikasi diri, lingkungan kerja, kepuasan kerja, dan kinerja pegawai.\n2. Data Sekunder: Data yang diperoleh secara tidak langsung melalui penelusuran dokumen kedinasan, meliputi data target dan realisasi penerimaan pajak daerah 2022–2024, Laporan Kinerja Instansi Pemerintah (LKjIP), profil kepegawaian, struktur organisasi, dan literatur pendukung.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "3.3.2 Teknik Pengumpulan Data\nPengumpulan data dilakukan melalui kombinasi teknik ilmiah berikut:\n1. Kuesioner (Angket): Menyebarkan daftar pernyataan terstruktur kepada 45 responden pegawai, baik melalui formulir digital (Google Form) maupun kuesioner cetak langsung guna memastikan keterisian 100%.\n2. Observasi: Melakukan pengamatan langsung secara sistematis terhadap kondisi fisik tempat kerja, fasilitas teknologi informasi, tata ruang kantor, dan dinamika interaksi kerja pegawai di lingkungan instansi.\n3. Studi Dokumentasi: Mengumpulkan dan mempelajari arsip resmi, laporan realisasi anggaran, peraturan bupati terkait pajak daerah, serta data statistik kepegawaian yang relevan.", indent=0.5)

    # 3.4 Variabel dan Definisi Operasional
    format_h2(doc, "3.4 Variabel dan Definisi Operasional")
    p = doc.add_paragraph()
    format_p(p, f"Penelitian ini menganalisis empat variabel yang terdiri dari dua variabel eksogen/independen ({x1} dan {x2}), satu variabel intervening/mediasi ({m}), dan satu variabel endogen/dependen ({y}). Definisi konseptual dan operasional masing-masing variabel diuraikan sebagai berikut:", indent=0.5)

    definisi_teks = [
        f"1. {y} (Y): Secara konseptual adalah hasil kerja kualitas dan kuantitas yang dicapai aparatur dalam melaksanakan tugas tanggung jawab kedinasan (Mangkunegara, 2019; Bernardin & Russell, 2013). Dioperasionalkan melalui 3 dimensi (Kualitas Hasil Kerja, Kuantitas & Ketepatan Waktu, serta Tanggung Jawab & Kerja Sama) dengan total 9 butir indikator empiris terukur (Y1.1 s.d. Y1.9).",
        f"2. {m} (M): Secara konseptual adalah keadaan emosional menyenangkan yang merefleksikan evaluasi positif pegawai terhadap pengalaman kerjanya (Luthans, 2015; Robbins & Judge, 2018). Dioperasionalkan melalui 3 dimensi (Pekerjaan itu Sendiri, Supervisi & Rekan Kerja, serta Penghargaan & Peluang Karier) dengan total 9 butir indikator empiris terukur (M1.1 s.d. M1.9).",
        f"3. {x1} (X1): Secara konseptual adalah penilaian keyakinan individu atas kapabilitas dirinya dalam mengorganisasi dan melaksanakan tindakan guna mencapai target tertentu (Bandura, 1997). Dioperasionalkan melalui 3 dimensi (Magnitude, Strength, dan Generality) dengan total 9 butir indikator empiris terukur (X1.1 s.d. X1.9).",
        f"4. {x2} (X2): Secara konseptual adalah keseluruhan faktor fisik dan non-fisik di sekitar tempat kerja yang memengaruhi kenyamanan dan produktivitas aparatur (Sedarmayanti, 2017). Dioperasionalkan melalui 3 dimensi (Lingkungan Fisik Fasilitas & Sanitasi, Lingkungan Fisik Tata Ruang & Ergonomi, serta Lingkungan Non-Fisik Hubungan Relasional) dengan total 9 butir indikator empiris terukur (X2.1 s.d. X2.9)."
    ]
    for def_item in definisi_teks:
        p_d = doc.add_paragraph()
        format_p(p_d, def_item, indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "Matriks operasionalisasi variabel, dimensi, indikator empiris terukur, butir pernyataan instrumen kuesioner, kode butir, dan skala pengukuran dirangkum secara komprehensif pada Tabel 3.2 berikut:", indent=0.5, space_before=4)

    # Tabel 3.2 Matriks Operasional Variabel (36 Butir Indikator)
    format_h3(doc, "Tabel 3.2 Matriks Kisi-Kisi Operasional Variabel Penelitian (36 Indikator)")
    
    # Flatten items from INDICATOR_SPEC
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
                    ind["statement"],
                    ind["code"],
                    "Skala Likert (1-5)"
                ])
                item_counter += 1

    t_ops = doc.add_table(rows=len(all_table_items) + 1, cols=7)
    o_heads = ["No", "Variabel", "Dimensi Teoretis", "Indikator Empiris", "Pernyataan Kuesioner", "Kode", "Skala"]
    for i, h in enumerate(o_heads):
        t_ops.cell(0, i).paragraphs[0].text = h
    
    for r_idx, row_vals in enumerate(all_table_items):
        for c_idx, val in enumerate(row_vals):
            t_ops.cell(r_idx + 1, c_idx).paragraphs[0].text = val
    apply_open_border_table(t_ops)

    # 3.5 Instrumen Penelitian
    format_h2(doc, "3.5 Instrumen Penelitian")
    p = doc.add_paragraph()
    format_p(p, "Pengukuran instrumen kuesioner menggunakan Skala Likert 5 Poin (Likert Scale) untuk mengukur sikap, persepsi, dan pendapat responden. Pemberian skor bobot jawaban ditetapkan sebagai berikut:\n1. Sangat Tidak Setuju (STS) diberi skor 1\n2. Tidak Setuju (TS) diberi skor 2\n3. Netral / Ragu-ragu (N) diberi skor 3\n4. Setuju (S) diberi skor 4\n5. Sangat Setuju (SS) diberi skor 5", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "Guna menginterpretasikan nilai rata-rata (mean score) tanggapan responden per variabel, digunakan rumus rentang skala interval (Sudjana, 2016):\nRentang Skala = (Skor Maksimal - Skor Minimal) / Jumlah Kategori = (5 - 1) / 5 = 0,80.\nKategori interpretasi skor rata-rata ditetapkan sebagai berikut:\n• 1,00 – 1,80 : Sangat Rendah / Sangat Buruk\n• 1,81 – 2,60 : Rendah / Buruk\n• 2,61 – 3,40 : Sedang / Cukup\n• 3,41 – 4,20 : Tinggi / Baik\n• 4,21 – 5,00 : Sangat Tinggi / Sangat Baik", indent=0.5)

    # 3.6 Teknik Analisis Data (PLS-SEM SmartPLS 4)
    format_h2(doc, "3.6 Teknik Analisis Data (PLS-SEM SmartPLS 4)")
    p = doc.add_paragraph()
    format_p(p, "Analisis data kuantitatif dalam penelitian ini dilakukan dengan metode Partial Least Squares - Structural Equation Modeling (PLS-SEM) menggunakan perangkat lunak SmartPLS 4. Menurut Hair et al. (2022), PLS-SEM merupakan teknik pemodelan persamaan struktural berbasis varians yang sangat unggul untuk pengujian model prediksi mediasi yang kompleks, memiliki ketahanan tinggi terhadap asumsi distribusi data normal maupun non-normal, serta sangat andal diaplikasikan pada ukuran sampel penelitian sensus (N = 45). Prosedur evaluasi model PLS-SEM dilakukan melalui dua tahapan utama:", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "1. Evaluasi Measurement Model (Outer Model)\nEvaluasi outer model bertujuan untuk menguji validitas dan reliabilitas instrumen yang merefleksikan konstruk laten, mencakup:\na. Convergent Validity: Dinilai berdasarkan nilai Outer Loadings dan Average Variance Extracted (AVE). Indikator dinyatakan memenuhi validitas konvergen jika nilai loading factor >= 0,708 (atau toleransi >= 0,700 untuk riset eksplanatori) dan nilai AVE untuk setiap konstruk laten > 0,500 (Hair et al., 2022).\nb. Discriminant Validity: Memastikan bahwa konstruk laten secara empiris unik dan berbeda dari konstruk lainnya. Dinilai melalui kriteria Heterotrait-Monotrait Ratio of Correlations (HTMT) dengan batas kritis HTMT < 0,850 (Henseler et al., 2015), serta evaluasi Cross Loadings dan Fornell-Larcker Criterion.\nc. Internal Consistency Reliability: Mengukur keandalan dan konsistensi internal instrumen melalui Composite Reliability (rho_c) dan Cronbach's Alpha. Suatu konstruk dinyatakan reliabel jika kedua nilai tersebut >= 0,700 (Hair et al., 2022).", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "2. Evaluasi Structural Model (Inner Model)\nEvaluasi inner model bertujuan untuk menguji hubungan prediktif antar-konstruk laten, dinilai melalui indikator berikut:\na. Koefisien Determinasi (R-Square / R2): Mengukur persentase variasi variabel endogen yang dapat dijelaskan oleh variabel eksogen. Nilai R2 diklasifikasikan menjadi: 0,67 (Kuat), 0,33 (Moderat), dan 0,19 (Lemah) (Chin, 1998).\nb. Effect Size (f-Square / f2): Menilai besaran kontribusi praktis masing-masing variabel eksogen terhadap nilai R2 variabel endogen. Nilai f2 dikelompokkan menjadi: 0,02 (Kecil), 0,15 (Menengah), dan 0,35 (Besar) (Cohen, 1988).\nc. Predictive Relevance (Q-Square / Q2): Menguji relevansi prediktif model out-of-sample menggunakan prosedur Blindfolding. Nilai Q2 > 0 menunjukkan bahwa model memiliki relevansi prediktif yang baik (Hair et al., 2019).\nd. Goodness of Fit Model: Dinilai melalui Standardized Root Mean Square Residual (SRMR) di mana nilai SRMR < 0,080 mengindikasikan model memiliki kecocokan yang baik (good fit).", indent=0.5)

    # 3.7 Pengujian Hipotesis
    format_h2(doc, "3.7 Pengujian Hipotesis Penelitian")
    p = doc.add_paragraph()
    format_p(p, "Pengujian signifikansi koefisien jalur struktural (Path Coefficients) dilakukan menggunakan prosedur non-parametrik Bootstrapping dengan 5.000 subsamples (Hair et al., 2022) pada tingkat signifikansi alpha 5% (tingkat kepercayaan 95%, uji dua sisi / two-tailed). Kriteria pengambilan keputusan penerimaan atau penolakan hipotesis disajikan pada Tabel 3.3 berikut:", indent=0.5)

    # Tabel 3.3 Kriteria Uji Hipotesis
    format_h3(doc, "Tabel 3.3 Kriteria Pengambilan Keputusan Pengujian Hipotesis")
    t_kep = doc.add_table(rows=3, cols=4)
    k_heads = ["No", "Parameter Statistik", "Nilai Batas Kritis", "Keputusan Uji"]
    for i, h in enumerate(k_heads):
        t_kep.cell(0, i).paragraphs[0].text = h
    kep_data = [
        ["1", "T-Statistik (t-hitung) vs T-Tabel", "T-Statistik ≥ 1,960", "Hipotesis Diterima (Signifikan)"],
        ["2", "P-Value (Tingkat Signifikansi α = 5%)", "P-Value ≤ 0,050", "Hipotesis Diterima (Signifikan)"]
    ]
    for r_idx, row_vals in enumerate(kep_data):
        for c_idx, val in enumerate(row_vals):
            t_kep.cell(r_idx + 1, c_idx).paragraphs[0].text = val
    apply_open_border_table(t_kep)

    p = doc.add_paragraph()
    format_p(p, "Untuk pengujian hipotesis mediasi (H6 dan H7), prosedur pengujian dilakukan dengan mengevaluasi Specific Indirect Effects pada luaran bootstrapping SmartPLS 4. Jika nilai t-statistik pengaruh tidak langsung >= 1,960 dan p-value <= 0,050, maka peran mediasi dinyatakan terbukti signifikan secara empiris. Sifat mediasi diidentifikasi apakah bersifat mediasi penuh (full mediation) atau mediasi parsial komplementer (complementary partial mediation) berdasarkan signifikansi pengaruh langsungnya.", indent=0.5, space_before=6)

    doc.save(output_path)
    return output_path

if __name__ == "__main__":
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output")
    out_test1 = os.path.join(out_dir, "BAB_3_METODE_PENELITIAN_LENGKAP.docx")
    out_test2 = os.path.join(out_dir, "BAB_3_METODE_PENELITIAN_NURMAYANTI.docx")
    generate_bab_3_docx({}, out_test1)
    generate_bab_3_docx({}, out_test2)
    print("Standalone Bab 3 generated successfully at:", out_test1)
    print("Standalone Bab 3 generated successfully at:", out_test2)
