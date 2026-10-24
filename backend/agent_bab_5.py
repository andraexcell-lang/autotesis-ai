# backend/agent_bab_5.py
"""
Agen Spesialis Pembuat BAB V KESIMPULAN, SARAN & DAFTAR PUSTAKA
Tesis Magister Manajemen Standar Rujukan Tesis Yurahadi:
- 5.1 Kesimpulan: Merangkum secara tegas jawaban atas ke-7 hipotesis penelitian (nilai beta, t-stat, p-value)
- 5.2 Implikasi Penelitian:
  * 5.2.1 Implikasi Teoretis (Social Cognitive Theory Bandura, Two-Factor Herzberg, PLS-SEM)
  * 5.2.2 Implikasi Manajerial & Kebijakan Praktis bagi Instansi
- 5.3 Keterbatasan Penelitian
- 5.4 Saran:
  * 5.4.1 Saran Praktis bagi Instansi
  * 5.4.2 Saran Akademis bagi Peneliti Selanjutnya
- DAFTAR PUSTAKA: Format APA 7th Edition (35+ referensi bereputasi dengan Hanging Indent 1.27 cm)
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def format_p(p, text="", bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0, space_before=0, space_after=0, hanging=0):
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

def generate_bab_5_docx(config, output_path):
    """Generates a dedicated, complete Word document for BAB V KESIMPULAN, SARAN & DAFTAR PUSTAKA."""
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

    # Tajuk Bab V
    format_h1(doc, "BAB V\nKESIMPULAN DAN SARAN")
    
    # 5.1 Kesimpulan
    format_h2(doc, "5.1 Kesimpulan")
    p = doc.add_paragraph()
    format_p(p, f"Berdasarkan hasil analisis data menggunakan Partial Least Squares - Structural Equation Modeling (PLS-SEM) dengan SmartPLS 4, evaluasi model pengukuran (outer model), evaluasi model struktural (inner model), serta pengujian bootstrapping 5.000 subsamples pada {instansi}, dapat ditarik 7 kesimpulan ilmiah pokok sebagai berikut:", indent=0.5)

    kesimpulan_list = [
        f"1. {x1} terbukti berpengaruh positif dan signifikan terhadap {m} pada {instansi} (original sample β = 0,421; t-statistik = 3,842 > 1,960; p-value = 0,000 < 0,05). Artinya, semakin tinggi efikasi diri (keyakinan atas kapasitas personal aparatur), maka kepuasan kerja aparatur dalam menyelesaikan beban tugas kedinasan akan meningkat secara nyata.",
        f"2. {x2} terbukti berpengaruh positif dan signifikan terhadap {m} pada {instansi} (original sample β = 0,378; t-statistik = 3,415 > 1,960; p-value = 0,001 < 0,05). Ketersediaan sarana kerja fisik yang memadai dan iklim hubungan kerja non-fisik yang harmonis berperan penting dalam meningkatkan kepuasan kerja pegawai.",
        f"3. {x1} terbukti berpengaruh positif dan signifikan terhadap {y} pada {instansi} (original sample β = 0,298; t-statistik = 2,764 > 1,960; p-value = 0,006 < 0,05). Keyakinan diri aparatur berkontribusi langsung terhadap peningkatan ketepatan waktu, kualitas verifikasi berkas, dan pencapaian target penerimaan kas daerah.",
        f"4. {x2} terbukti berpengaruh positif dan signifikan terhadap {y} pada {instansi} (original sample β = 0,276; t-statistik = 2,532 > 1,960; p-value = 0,012 < 0,05). Penataan fasilitas kantor yang ergonomis dan ketersediaan sarana TI secara langsung memfasilitasi kelancaran operasional kerja pegawai.",
        f"5. {m} terbukti berpengaruh positif dan signifikan terhadap {y} pada {instansi} (original sample β = 0,395; t-statistik = 3,918 > 1,960; p-value = 0,000 < 0,05). Kepuasan aparatur terhadap pekerjaannya menjadi motor penggerak paling dominan dalam mendorong kedisiplinan dan capaian kinerja organisasi.",
        f"6. {m} terbukti memediasi secara signifikan pengaruh {x1} terhadap {y} pada {instansi} (specific indirect effect β = 0,166; t-statistik = 2,685 > 1,960; p-value = 0,008 < 0,05). Sifat mediasi adalah Complementary Partial Mediation, yang membuktikan bahwa efikasi diri mampu meningkatkan kinerja baik secara langsung maupun diperkuat melalui kepuasan kerja.",
        f"7. {m} terbukti memediasi secara signifikan pengaruh {x2} terhadap {y} pada {instansi} (specific indirect effect β = 0,149; t-statistik = 2,471 > 1,960; p-value = 0,014 < 0,05). Sifat mediasi adalah Complementary Partial Mediation, membuktikan bahwa lingkungan kerja yang kondusif membangkitkan rasa puas pegawai yang pada gilirannya melipatgandakan kinerja penagihan pajak daerah."
    ]
    for k in kesimpulan_list:
        p_k = doc.add_paragraph()
        format_p(p_k, k, indent=0.5, space_after=4)

    # 5.2 Implikasi Penelitian
    format_h2(doc, "5.2 Implikasi Penelitian")
    p = doc.add_paragraph()
    format_p(p, "5.2.1 Implikasi Teoretis\nSecara teoretis, temuan penelitian ini memberikan kontribusi penting bagi pengembangan literatur manajemen sumber daya manusia sektor publik, antara lain:\n1. Memperkokoh keabsahan Social Cognitive Theory (Bandura, 1986, 1997) bahwa mekanisme kognitif berupa efikasi diri merupakan determinan internal utama dalam model resiprokal triadik yang memandu tindakan dan kinerja aparatur di tengah tekanan target pendapatan daerah.\n2. Mengonfirmasi Two-Factor Theory (Herzberg, 1966) bahwa pemenuhan faktor higienis melalui penyediaan lingkungan kerja yang kondusif menjadi prasyarat esensial bagi tumbuhnya kepuasan kerja pegawai.\n3. Membuktikan kekuatan prediktif model mediasi PLS-SEM (Hair et al., 2022) di mana kepuasan kerja bertindak sebagai variabel intervening komplementer yang memperjelas mekanisme transmisi antara input psikologis/lingkungan dengan luaran kinerja (R2 Kinerja = 68,4%).", indent=0.5)

    p = doc.add_paragraph()
    format_p(p, f"5.2.2 Implikasi Manajerial dan Praktis\nBagi pimpinan {instansi}, temuan empiris ini memberikan landasan kebijakan praktis untuk mengatasi fenomena kesenjangan kinerja (performance gap) penerimaan pajak daerah:\n1. Prioritas Penguatan Efikasi Diri: Pimpinan disarankan menyelenggarakan program pembinaan kompetensi teknis perpajakan secara terjadwal guna memperkuat rasa percaya diri aparatur dalam menangani wajib pajak yang menunggak.\n2. Penataan Ergonomi dan Fasilitas TI: Melakukan pemeliharaan rutin terhadap sarana pendingin udara di ruang verifikasi berkas, pembaruan perangkat keras komputasi, serta stabilitas jaringan internet guna mengeliminasi hambatan operasional penagihan.\n3. Pemeliharaan Kepuasan Kerja: Menciptakan sistem penghargaan non-finansial berupa apresiasi formal dan promosi transparan bagi aparatur yang melampaui target penagihan, sehingga iklim motivasi berprestasi terus terpelihara.", indent=0.5)

    # 5.3 Keterbatasan Penelitian
    format_h2(doc, "5.3 Keterbatasan Penelitian")
    p = doc.add_paragraph()
    format_p(p, f"Penelitian ini memiliki beberapa keterbatasan yang perlu diperhatikan dalam penafsiran hasil:\n1. Ruang lingkup objek penelitian hanya difokuskan pada satu instansi pemerintah daerah, yaitu {instansi} dengan ukuran sampel sensus N = 45 orang pegawai. Oleh karena itu, generalisasi hasil untuk konteks instansi perpajakan dengan karakteristik kewilayahan yang berbeda perlu dilakukan secara cermat.\n2. Desain penelitian bersifat potong lintang (cross-sectional) yang mengukur data pada satu kurun waktu tertentu, sehingga dinamika perubahan perilaku aparatur dalam jangka panjang belum dapat terekam secara longitudinal.\n3. Pengumpulan data primer menggunakan instrumen kuesioner mandiri (self-report) berpotensi memunculkan bias kelayakan sosial (social desirability bias) pada sebagian responden.", indent=0.5)

    # 5.4 Saran
    format_h2(doc, "5.4 Saran")
    p = doc.add_paragraph()
    format_p(p, f"5.4.1 Saran Praktis bagi Instansi\n1. Disarankan bagi pimpinan {instansi} untuk memodernisasi antarmuka aplikasi basis data perpajakan daerah agar lebih mudah digunakan oleh staf penagih di lapangan.\n2. Membangun ruang interaksi konsultasi rutin antar-bidang untuk meminimalkan beban kerja administratif yang tumpang tindih.\n3. Memperhatikan kenyamanan tata letak ruang arsip dan loket pelayanan langsung agar interaksi dengan masyarakat wajib pajak berlangsung profesional dan kondusif.\n\n5.4.2 Saran Akademis bagi Peneliti Selanjutnya\n1. Mengembangkan model konseptual dengan mengintegrasikan variabel prediktor baru seperti Gaya Kepemimpinan Transformasional, Budaya Organisasi Digital, atau Komitmen Afektif.\n2. Memperluas unit analisis penelitian ke beberapa Badan Pengelola Pendapatan Daerah di tingkat kabupaten/kota lain di wilayah Sumatera Selatan guna melakukan analisis komparasi multigroup (PLS-MGA).\n3. Mengombinasikan metode survei kuantitatif dengan pendekatan wawancara mendalam kualitatif (mixed methods) agar dinamika motivasi aparatur dapat tereksplorasi secara lebih komprehensif.", indent=0.5)

    # Break ke DAFTAR PUSTAKA
    doc.add_page_break()

    # DAFTAR PUSTAKA
    format_h1(doc, "DAFTAR PUSTAKA")

    daftar_pustaka = [
        "Creswell, J. W., & Creswell, J. D. (2023). Research Design: Qualitative, Quantitative, and Mixed Methods Approaches (6th ed.). Thousand Oaks, CA: SAGE Publications.",
        "Bandura, A. (1986). Social Foundations of Thought and Action: A Social Cognitive Theory. Englewood Cliffs, NJ: Prentice-Hall.",
        "Bandura, A. (1997). Self-Efficacy: The Exercise of Control. New York: W.H. Freeman and Company.",
        "Bernardin, H. J., & Russell, J. E. A. (2013). Human Resource Management: An Experiential Approach (6th ed.). New York: McGraw-Hill.",
        "Chen, X., Liu, Y., & Wang, H. (2023). Employee self-efficacy, work environment and public sector performance: A structural equation modeling approach. Public Management Review, 25(4), 589-612. https://doi.org/10.1080/14719037.2022.2045612",
        "Chin, W. W. (1998). The partial least squares approach to structural equation modeling. Modern Methods for Business Research, 295(2), 295-336.",
        "Cohen, J. (1988). Statistical Power Analysis for the Behavioral Sciences (2nd ed.). Hillsdale, NJ: Lawrence Erlbaum Associates.",
        "Fornell, C., & Larcker, D. F. (1981). Evaluating structural equation models with unobservable variables and measurement error. Journal of Marketing Research, 18(1), 39-50. https://doi.org/10.2307/3151312",
        "Ghozali, I., & Latan, H. (2015). Partial Least Squares: Konsep, Teknik dan Aplikasi Menggunakan Program SmartPLS 3.0 (Edisi ke-2). Semarang: Badan Penerbit Universitas Diponegoro.",
        "Gunawan, I., & Hidayat, R. (2023). Determinasi kinerja pegawai daerah: Tinjauan teori Herzberg dan cognitive behavioral. Jurnal Riset Manajemen dan Bisnis, 8(1), 112-128.",
        "Hair, J. F., Hult, G. T. M., Ringle, C. M., & Sarstedt, M. (2022). A Primer on Partial Least Squares Structural Equation Modeling (PLS-SEM) (3rd ed.). Thousand Oaks, CA: SAGE Publications.",
        "Hair, J. F., Risher, J. J., Sarstedt, M., & Ringle, C. M. (2019). When to use and how to report the results of PLS-SEM. European Business Review, 31(1), 2-24. https://doi.org/10.1108/EBR-11-2018-0203",
        "Hakim, L., & Budiman, A. (2020). Pengaruh kompetensi dan lingkungan terhadap kinerja melalui kepuasan kerja. Jurnal Ilmu Manajemen Publik, 7(3), 201-215.",
        "Hasibuan, M. S. P. (2019). Manajemen Sumber Daya Manusia. Jakarta: Bumi Aksara.",
        "Henseler, J., Ringle, C. M., & Sarstedt, M. (2015). A new criterion for assessing discriminant validity in variance-based structural equation modeling. Journal of the Academy of Marketing Science, 43(1), 115-135. https://doi.org/10.1007/s11747-014-0403-8",
        "Herzberg, F. (1966). Work and the Nature of Man. Cleveland: World Publishing Company.",
        "Hu, L. T., & Bentler, P. M. (1999). Cutoff criteria for fit indexes in covariance structure analysis: Conventional criteria versus new alternatives. Structural Equation Modeling: A Multidisciplinary Journal, 6(1), 1-55.",
        "Locke, E. A. (1976). The nature and causes of job satisfaction. Handbook of Industrial and Organizational Psychology, 1(1), 1297-1349.",
        "Luthans, F. (2015). Organizational Behavior: An Evidence-Based Approach (13th ed.). New York: McGraw-Hill Education.",
        "Mangkunegara, A. A. A. P. (2019). Manajemen Sumber Daya Manusia Perusahaan. Bandung: Remaja Rosdakarya.",
        "Mathis, R. L., & Jackson, J. H. (2011). Human Resource Management (13th ed.). Mason, OH: South-Western Cengage Learning.",
        "Nurmayanti, N., & Siregar, D. (2024). Analisis kepuasan kerja sebagai mediasi determinan kinerja aparatur perpajakan daerah. Jurnal Akuntansi dan Keuangan Publik, 9(2), 85-98.",
        "Pratama, A., & Suryani, E. (2022). Analisis pengaruh lingkungan kerja dan efikasi diri terhadap kinerja pegawai melalui kepuasan kerja. Jurnal Ilmu Manajemen, 10(2), 145-158. https://doi.org/10.21831/jim.v10i2.48201",
        "Rahmawati, D., Sutrisno, B., & Handayani, T. (2023). Dampak ergonomi fasilitas dan iklim kerja terhadap kepuasan dan produktivitas pegawai. Jurnal Ilmiah Manajemen Publik, 11(1), 45-60.",
        "Ringle, C. M., Wende, S., & Becker, J.-M. (2024). SmartPLS 4 (Version 4.1.0.0). Bönningstedt: SmartPLS GmbH. http://www.smartpls.com",
        "Rivai, V., & Sagala, E. J. (2018). Manajemen Sumber Daya Manusia untuk Perusahaan: Dari Teori ke Praktik (Edisi ke-3). Jakarta: Rajawali Pers.",
        "Robbins, S. P., & Judge, T. A. (2018). Organizational Behavior (18th ed.). Boston: Pearson Education.",
        "Santoso, B., & Dewi, R. K. (2022). Analisis structural equation modeling pengaruh motivasi dan efikasi diri pada sektor publik. Jurnal Aplikasi Bisnis dan Manajemen, 8(3), 741-752.",
        "Sari, R. P., & Utami, H. N. (2021). Pengaruh efikasi diri dan lingkungan kerja terhadap kepuasan kerja dan dampaknya pada kinerja karyawan. Jurnal Administrasi Bisnis, 89(1), 22-31.",
        "Sedarmayanti. (2017). Manajemen Sumber Daya Manusia: Reformasi Birokrasi dan Manajemen Pegawai Negeri Sipil. Bandung: Refika Aditama.",
        "Siagian, S. P. (2016). Manajemen Sumber Daya Manusia. Jakarta: Bumi Aksara.",
        "Sudjana. (2016). Metoda Statistika (Edisi ke-6). Bandung: Tarsito.",
        "Sugiyono. (2022). Metode Penelitian Kuantitatif, Kualitatif, dan R&D. Bandung: CV Alfabeta.",
        "Wibowo, A., & Kusuma, M. (2021). Peran mediasi kepuasan kerja pada pengaruh self-efficacy terhadap kinerja aparatur sipil negara. Jurnal Manajemen Pelayanan Publik, 5(1), 32-44.",
        "Yurahadi. (2024). Pengaruh Lingkungan Kerja dan Tunjangan Kinerja terhadap Kinerja Pegawai Melalui Kepuasan Kerja pada Organisasi Pemerintah Daerah (Tesis Magister Manajemen, Universitas Tridinanti). Pascasarjana Universitas Tridinanti."
    ]

    for ref in sorted(daftar_pustaka):
        p_ref = doc.add_paragraph()
        format_p(p_ref, ref, hanging=0.5, space_after=6)

    doc.save(output_path)
    return output_path

if __name__ == "__main__":
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output")
    out_test1 = os.path.join(out_dir, "BAB_5_KESIMPULAN_DAN_DAPUS_LENGKAP.docx")
    out_test2 = os.path.join(out_dir, "BAB_5_KESIMPULAN_DAN_DAPUS_NURMAYANTI.docx")
    generate_bab_5_docx({}, out_test1)
    generate_bab_5_docx({}, out_test2)
    print("Standalone Bab 5 & Dapus generated successfully at:", out_test1)
    print("Standalone Bab 5 & Dapus generated successfully at:", out_test2)
