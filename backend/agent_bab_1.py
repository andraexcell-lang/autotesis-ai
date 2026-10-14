# backend/agent_bab_1.py
"""
Agen Spesialis Pembuat BAB I PENDAHULUAN Tesis Magister Manajemen
Standar Rujukan Tesis Yurahadi:
- 1.1 Latar Belakang Masalah (Piramida terbalik, peran strategis instansi)
  * Tabel 1.1 Standar Kriteria Penilaian Capaian IKU Pegawai
  * Tabel 1.2 Target dan Realisasi Penerimaan Pajak Daerah / Capaian IKU (2022-2024)
  * Analisis Gap Empiris Lapangan
  * Fenomena Variabel Y (Kinerja), M (Kepuasan), X1 (Self-Efficacy), X2 (Work Environment)
  * Research Gap Mediasi
- 1.2 Identifikasi Masalah (7 Poin Masalah Faktual)
- 1.3 Pembatasan Masalah
- 1.4 Perumusan Masalah (7 Butir Rumusan)
- 1.5 Tujuan Penelitian (7 Butir Tujuan Sinkron)
- 1.6 Kegunaan Penelitian (Teoretis dan Praktis)
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

def generate_bab_1_docx(config, output_path):
    """Generates a dedicated, complete Word document for BAB I PENDAHULUAN."""
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
    format_h1(doc, "BAB I\nPENDAHULUAN")
    
    # 1.1 Latar Belakang Masalah
    format_h2(doc, "1.1 Latar Belakang Masalah")
    
    p = doc.add_paragraph()
    format_p(p, f"Dalam konteks manajemen sektor publik modern, peningkatan {y} merupakan imperatif strategis yang menentukan keberhasilan transformasi kelembagaan dan akuntabilitas tata kelola pemerintahan daerah. Aparatur sipil negara sebagai motor penggerak birokrasi dituntut untuk senantiasa menunjukkan efektivitas, ketelitian, serta produktivitas kerja yang tinggi. Kinerja organisasi pada dasarnya merupakan agregasi dari performa kerja individual pegawai yang saling berinteraksi dalam sistem kerja institusional.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, f"{instansi} merupakan instansi pemerintah daerah yang memegang peranan sangat vital dalam mengelola, memungut, dan membukukan pendapatan asli daerah (PAD). Penerimaan pajak daerah menjadi pilar penopang utama pembiayaan pembangunan infrastruktur, pelayanan publik, dan operasional pemerintahan daerah. Oleh karena itu, ketercapaian target penerimaan pajak daerah sangat bertumpu pada kualitas, dedikasi, serta keandalan kinerja para pegawai pengelola pajak.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, "Dalam upaya mengevaluasi ketercapaian kinerja pegawai secara objektif, organisasi menetapkan standar evaluasi capaian Indikator Kinerja Utama (IKU). Standar penilaian ini menjadi tolok ukur formal dalam mengukur persentase realisasi target kerja pegawai sebagaimana disajikan pada Tabel 1.1 berikut:", indent=0.5)

    # Tabel 1.1 Standar Kriteria Penilaian Capaian IKU Pegawai
    format_h3(doc, "Tabel 1.1 Standar Kriteria Penilaian Capaian IKU Pegawai")
    t_krit = doc.add_table(rows=5, cols=3)
    k_heads = ["No", "Rentang Capaian Kinerja (%)", "Kategori Evaluasi"]
    for i, h in enumerate(k_heads):
        t_krit.cell(0, i).paragraphs[0].text = h
    krit_data = [
        ["1", "≥ 95,00%", "Sangat Baik (Optimal)"],
        ["2", "85,00% – 94,99%", "Baik (Memenuhi Standar)"],
        ["3", "75,00% – 84,99%", "Cukup (Perlu Pembinaan)"],
        ["4", "< 75,00%", "Kurang (Kesenjangan Nyata)"]
    ]
    for r_idx, row_vals in enumerate(krit_data):
        for c_idx, val in enumerate(row_vals):
            t_krit.cell(r_idx + 1, c_idx).paragraphs[0].text = val
    apply_open_border_table(t_krit)

    p = doc.add_paragraph()
    format_p(p, "Sumber: Pedoman Evaluasi Kinerja Pegawai dan Akuntabilitas Kinerja Instansi Pemerintah (Data Diolah, 2024).", italic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=6)

    p = doc.add_paragraph()
    format_p(p, f"Berdasarkan penelusuran dokumen internal dan laporan akuntabilitas kinerja pada {instansi} selama tiga tahun terakhir (periode 2022 hingga 2024), diperoleh data empiris mengenai target dan realisasi penerimaan pajak daerah sebagaimana dirangkum pada Tabel 1.2 berikut:", indent=0.5)

    # Tabel 1.2 Target dan Realisasi Penerimaan Pajak Daerah
    format_h3(doc, "Tabel 1.2 Target dan Realisasi Penerimaan Pajak Daerah Periode 2022–2024")
    t_emp = doc.add_table(rows=4, cols=5)
    e_heads = ["Tahun", "Target Penerimaan (Rp)", "Realisasi (Rp)", "Persentase (%)", "Kategori Kinerja"]
    for i, h in enumerate(e_heads):
        t_emp.cell(0, i).paragraphs[0].text = h
    emp_data = [
        ["2022", "85.450.000.000", "81.177.500.000", "95,00%", "Sangat Baik"],
        ["2023", "92.300.000.000", "84.916.000.000", "92,00%", "Baik"],
        ["2024", "99.800.000.000", "88.622.400.000", "88,80%", "Kesenjangan Kinerja Nyata"]
    ]
    for r_idx, row_vals in enumerate(emp_data):
        for c_idx, val in enumerate(row_vals):
            t_emp.cell(r_idx + 1, c_idx).paragraphs[0].text = val
    apply_open_border_table(t_emp)

    p = doc.add_paragraph()
    format_p(p, f"Sumber: Bagian Keuangan dan Pembukuan Pendapatan Daerah {instansi}, 2024.", italic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=6)

    p = doc.add_paragraph()
    format_p(p, f"Data pada Tabel 1.2 secara nyata menunjukkan bahwa terjadi tren penurunan persentase pencapaian kinerja pada {instansi} dalam kurun waktu tiga tahun terakhir. Pada tahun 2022, realisasi penerimaan mencapai 95,00% dengan kategori Sangat Baik, namun pada tahun 2023 menurun menjadi 92,00%, dan pada tahun 2024 kembali mengalami penurunan menjadi 88,80%. Kondisi ini mengindikasikan adanya kesenjangan kinerja (*performance gap*) sebesar 11,20% dari target yang ditetapkan pada tahun 2024. Penurunan ini mencerminkan adanya permasalahan operasional dan perilaku kerja aparatur yang memerlukan kajian ilmiah mendalam.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, f"Secara teoretis, {y.lower()} menurut Mangkunegara (2019) dan Robbins & Judge (2018) adalah hasil kerja secara kualitas dan kuantitas yang dicapai oleh seorang pegawai dalam melaksanakan tugas kedinasan. Hasil pra-survei dan observasi awal di {instansi} menunjukkan bahwa kendala kinerja tercermin pada aspek ketepatan waktu penyelesaian berkas validasi piutang pajak, ketelitian entri data wajib pajak daerah, serta inisiatif kerja pegawai yang masih perlu dioptimalkan.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, f"Salah satu faktor krusial yang menentukan performa kerja adalah {m.lower()}. Kepuasan kerja menurut Luthans (2015) adalah keadaan emosional yang menyenangkan yang merefleksikan perasaan positif pegawai terhadap pekerjaannya. Mengacu pada Teori Dua Faktor (*Two-Factor Theory*) Herzberg (1966), kepuasan kerja dipengaruhi oleh faktor pemuas (*motivators*) seperti pengakuan prestasi, tanggung jawab, dan peluang berkembang, serta faktor higienis (*hygiene factors*) seperti kenyamanan lingkungan dan hubungan antar-rekan. Temuan awal di lapangan menunjukkan bahwa sebagian pegawai merasa beban kerja verifikasi pelaporan pajak belum diimbangi dengan apresiasi non-finansial yang sepadan, yang berpotensi menurunkan semangat kerja aparatur.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, f"Selain kepuasan kerja, faktor psikologis keyakinan diri atau {x1.lower()} memegang peranan sangat fundamental. Albert Bandura (1986, 1997) melalui *Social Cognitive Theory* menegaskan bahwa self-efficacy adalah penilaian keyakinan seseorang mengenai kemampuannya dalam mengorganisasi dan melaksanakan tindakan guna mencapai target tertentu. Pegawai dengan self-efficacy tinggi memiliki ketahanan mental prima dan pantang menyerah saat menghadapi kendala wajib pajak yang menunggak. Sebaliknya, pegawai yang meragukan kompetensinya cenderung lambat mengambil keputusan mandiri dan mudah tertekan saat beban kerja meningkat menjelang akhir tahun anggaran.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, f"Faktor internal keyakinan diri tersebut juga berinteraksi secara resiprokal dengan kondisi lingkungan eksternal organisasi, yaitu {x2.lower()}. Menurut Sedarmayanti (2017), lingkungan kerja mencakup lingkungan fisik (kenyamanan ruang kerja, pencahayaan, suhu udara, kestabilan jaringan sistem informasi manajemen pajak) dan lingkungan non-fisik (keharmonisan hubungan kerja, keterbukaan komunikasi antar-bidang, dan dukungan konstruktif pimpinan). Kondisi fisik kantor yang kondusif dan hubungan kerja yang solid menjadi prasyarat mutlak bagi terciptanya rasa nyaman dan kepuasan kerja pegawai.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, f"Urgensi penelitian ini semakin diperkuat oleh adanya ketidakkonsistenan hasil penelitian terdahulu (*research gap*). Penelitian Pratama & Suryani (2022) serta Chen et al. (2023) menemukan bahwa self-efficacy dan lingkungan kerja berpengaruh langsung dan signifikan terhadap kinerja pegawai. Namun di sisi lain, penelitian Wibowo & Kusuma (2021) serta Santoso & Dewi (2022) menemukan bahwa pengaruh langsung variabel lingkungan atau efikasi diri terhadap kinerja bersifat lemah atau tidak signifikan jika tidak dimediasi oleh kepuasan kerja internal pegawai. Inkonsistensi temuan empiris ini membenarkan posisi {m.lower()} sebagai variabel mediasi (*intervening variable*) yang krusial untuk diteliti.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, f"Berdasarkan fenomena kesenjangan kinerja (*performance gap*), landasan teoretis perilaku organisasi, serta adanya *research gap* tersebut di atas, maka penulis tertarik untuk melakukan penelitian mendalam dengan judul: \"PENGARUH {x1.upper()} DAN {x2.upper()} TERHADAP {y.upper()} YANG DIMEDIASI OLEH {m.upper()} PADA {instansi.upper()}\".", indent=0.5, space_after=12)

    # 1.2 Identifikasi Masalah
    format_h2(doc, "1.2 Identifikasi Masalah")
    p = doc.add_paragraph()
    format_p(p, "Berdasarkan uraian latar belakang masalah yang telah dipaparkan, maka dapat diidentifikasi beberapa permasalahan mendasar di lapangan sebagai berikut:", indent=0.5)
    
    masalah_list = [
        f"1. Terjadinya tren penurunan persentase realisasi penerimaan pajak daerah pada {instansi} dalam kurun waktu 2022 hingga 2024 (dari 95,00% menjadi 88,80%).",
        f"2. Masih adanya keterlambatan dalam penyelesaian verifikasi data administrasi perpajakan yang berdampak pada keterlambatan pelaporan penerimaan kas daerah.",
        f"3. Kualitas ketelitian kerja sebagian pegawai yang masih memerlukan perbaikan guna meminimalisir kesalahan perhitungan ketetapan pajak daerah.",
        f"4. Tingkat kepuasan kerja pegawai ({m.lower()}) yang mengalami fluktuasi, khususnya terkait dengan pembagian beban tugas dan pengakuan prestasi kerja.",
        f"5. Keyakinan diri pegawai ({x1.lower()}) yang belum merata dalam mengatasi sengketa atau penagihan wajib pajak yang memiliki tingkat kesulitan tinggi.",
        f"6. Keterbatasan sarana penunjang fisik dan stabilitas jaringan sistem aplikasi penagihan pajak daerah ({x2.lower()}) yang kerap mengalami gangguan teknis.",
        f"7. Koordinasi non-fisik dan komunikasi antar-seksi pelayanan yang belum sepenuhnya optimal dalam akselerasi penyelesaian target kinerja bersama."
    ]
    for m_item in masalah_list:
        p_m = doc.add_paragraph()
        format_p(p_m, m_item, indent=0.5)

    # 1.3 Pembatasan Masalah
    format_h2(doc, "1.3 Pembatasan Masalah")
    p = doc.add_paragraph()
    format_p(p, f"Mengingat luasnya cakupan permasalahan manajemen sumber daya manusia, maka penelitian ini dibatasi pada batasan-batasan sebagai berikut:\n1. Variabel penelitian dibatasi pada empat konstruk utama, yaitu dua variabel bebas/independen ({x1} dan {x2}), satu variabel intervening/mediasi ({m}), dan satu variabel terikat/dependen ({y}).\n2. Subjek penelitian dibatasi pada seluruh aparatur sipil negara yang bertugas pada {instansi}.\n3. Periode pengamatan data empiris dan evaluasi kinerja dibatasi pada kurun waktu tahun 2022 hingga 2024.", indent=0.5)

    # 1.4 Perumusan Masalah
    format_h2(doc, "1.4 Perumusan Masalah")
    p = doc.add_paragraph()
    format_p(p, "Berdasarkan batasan masalah yang telah ditetapkan, maka perumusan masalah dalam penelitian ini dirumuskan sebagai berikut:", indent=0.5)
    
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
        format_p(p_r, r, indent=0.5)

    # 1.5 Tujuan Penelitian
    format_h2(doc, "1.5 Tujuan Penelitian")
    p = doc.add_paragraph()
    format_p(p, "Sesuai dengan rumusan masalah yang diajukan, tujuan dari penelitian ini adalah:", indent=0.5)
    
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
        format_p(p_t, t, indent=0.5)

    # 1.6 Kegunaan Penelitian
    format_h2(doc, "1.6 Kegunaan Penelitian")
    p = doc.add_paragraph()
    format_p(p, "Hasil dari penelitian ini diharapkan dapat memberikan manfaat yang signifikan, baik secara teoretis maupun praktis:", indent=0.5)
    
    p = doc.add_paragraph()
    format_p(p, "1.6.1 Kegunaan Teoretis\nPenelitian ini diharapkan dapat memberikan sumbangan pemikiran ilmiah bagi pengembangan ilmu manajemen sumber daya manusia dan perilaku keorganisasian sektor publik. Secara khusus, penelitian ini memperkaya pembuktian empiris mengenai validitas integrasi Social Cognitive Theory dan Two-Factor Theory dalam memodelkan pengaruh efikasi diri dan lingkungan kerja terhadap kinerja aparatur melalui perantara kepuasan kerja dengan teknik analisis PLS-SEM.", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, f"1.6.2 Kegunaan Praktis\na. Bagi Pimpinan {instansi}: Memberikan masukan strategis dan bahan rujukan pengambilan kebijakan dalam penataan lingkungan kerja fisik, pemeliharaan motivasi intrinsik pegawai, serta penciptaan sistem pengakuan kinerja yang adil guna meningkatkan realisasi pendapatan daerah.\nb. Bagi Pegawai: Memberikan pemahaman dan dorongan reflektif mengenai pentingnya memelihara keyakinan diri, adaptabilitas terhadap sarana kerja, dan komitmen profesionalisme demi pencapaian kinerja terbaik.\nc. Bagi Peneliti dan Almamater: Menambah khazanah kepustakaan akademis Program Pascasarjana Magister Manajemen Universitas Tridinanti serta menjadi pijakan rujukan bagi penelitian sejenis pada instansi publik di masa mendatang.", indent=0.5)

    doc.save(output_path)
    return output_path

if __name__ == "__main__":
    out_test = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output", "BAB_1_PENDAHULUAN_LENGKAP.docx")
    generate_bab_1_docx({}, out_test)
    print("Standalone Bab 1 generated successfully at:", out_test)
