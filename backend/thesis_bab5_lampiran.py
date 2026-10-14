# backend/thesis_bab5_lampiran.py
"""Modul Pembangun BAB V, Daftar Pustaka, dan Lampiran Lengkap Tesis Andre Juliansya"""

from docx.enum.text import WD_ALIGN_PARAGRAPH
from build_full_thesis_docx import p, h1, h2, h3, add_table_data
import build_andre_thesis

def build_bab_5(doc):
    h1(doc, "BAB V\nKESIMPULAN, IMPLIKASI DAN SARAN")
    
    h2(doc, "5.1. Kesimpulan")
    p(doc, "Berdasarkan hasil analisis data, pengujian hipotesis menggunakan Partial Least Square (PLS-SEM), serta pembahasan yang telah diuraikan pada bab sebelumnya mengenai fenomena di Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Sumatera Selatan, maka dapat ditarik beberapa kesimpulan sebagai berikut:", space_after=4)
    kesimpulan = [
        "1. Motivasi Kerja terbukti berpengaruh positif dan signifikan terhadap Employee Engagement pegawai Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Sumatera Selatan (beta = 0,565, t = 9,100, p = 0,000).",
        "2. Disiplin Kerja terbukti berpengaruh positif dan signifikan terhadap Employee Engagement pegawai pada Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Sumatera Selatan (beta = 0,459, t = 6,318, p = 0,000).",
        "3. Motivasi Kerja terbukti berpengaruh positif dan signifikan terhadap Kinerja Pegawai pada Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Sumatera Selatan (beta = 0,191, t = 3,271, p = 0,001).",
        "4. Disiplin Kerja terbukti berpengaruh positif dan signifikan terhadap Kinerja Pegawai pada Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Sumatera Selatan (beta = 0,295, t = 4,544, p = 0,000).",
        "5. Employee Engagement terbukti berpengaruh positif dan signifikan terhadap Kinerja Pegawai pada Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Sumatera Selatan (beta = 0,563, t = 7,018, p = 0,000).",
        "6. Motivasi Kerja terbukti berpengaruh positif dan signifikan terhadap Kinerja Pegawai melalui Employee Engagement sebagai variabel mediasi pada Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Sumatera Selatan (beta = 0,318, t = 5,423, p = 0,000, Partial Mediation).",
        "7. Disiplin Kerja terbukti berpengaruh positif dan signifikan terhadap Kinerja Pegawai melalui Employee Engagement sebagai variabel mediasi pada Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Sumatera Selatan (beta = 0,258, t = 4,379, p = 0,000, Partial Mediation)."
    ]
    for k in kesimpulan:
        p(doc, k, indent=0.25, space_after=2)

    h2(doc, "5.2. Implikasi Kebijakan")
    p(doc, "Temuan empiris dan pembuktian statistik dalam penelitian ini melahirkan sejumlah implikasi manajerial yang dapat diadopsi sebagai landasan perumusan kebijakan berbasis bukti (Evidence-Based Policy) bagi Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Sumatera Selatan. Untuk mengatasi krisis kinerja dan memulihkan target Indikator Kinerja Utama (IKU) penerbitan izin yang sempat defisit hingga -16,6%, pimpinan dan pembuat kebijakan disarankan untuk mengimplementasikan langkah-langkah strategis berikut:", space_after=4)
    
    p(doc, "1. Kebijakan Restrukturisasi Motivasi dan Sistem Apresiasi", bold=True, space_after=2)
    p(doc, "a). Pemerataan Beban Kerja: Perlu adanya penataan ulang atau rotasi berkala bagi pegawai yang bertugas di garda terdepan (loket operasional perizinan) untuk mencegah kelelahan fisik dan mental (Burnout). Proporsi beban kerja harus diseimbangkan dengan waktu istirahat yang memadai.\nb). Penguatan Pengakuan Non-Finansial: Pimpinan disarankan untuk melembagakan sistem apresiasi, seperti program Employee of the Month, pemberian piagam penghargaan internal, atau sekadar pujian verbal dan pengakuan atas keahlian saat rapat evaluasi.", indent=0.5, space_after=4)

    p(doc, "2. Transformasi Pendekatan Pendisiplinan Aparatur", bold=True, space_after=2)
    p(doc, "a). Transisi dari Disiplin Menghukum ke Disiplin Preventif: Kebijakan penegakan disiplin administratif tidak boleh lagi bersifat kaku yang hanya mengejar punishment, karena hal ini terbukti hanya melahirkan kepatuhan semu.\nb). Integrasi Sistem Pemantauan: Pengawasan absensi biometrik perlu diintegrasikan secara transparan dengan pencairan Tunjangan Kinerja (Tukin). Selain itu, atasan langsung diwajibkan melakukan pendekatan persuasif (pembinaan empatik) untuk mengidentifikasi akar masalah keterlambatan pegawai.", indent=0.5, space_after=4)

    p(doc, "3. Pelembagaan Program Penguatan Employee Engagement", bold=True, space_after=2)
    p(doc, "a). Program Ketangguhan Mental (Mental Resilience): Untuk mengatasi fenomena Quiet Quitting dan penurunan dimensi Vigor, instansi dapat memfasilitasi sesi Capacity Building atau pelatihan manajemen stres.\nb). Penyelarasan Visi dan Kebanggaan Institusional: Pimpinan perlu secara rutin menyelenggarakan Town Hall Meeting atau Morning Briefing guna menginternalisasikan peran strategis DPMPTSP bagi iklim investasi daerah.", indent=0.5, space_after=4)

    p(doc, "4. Optimalisasi Kinerja Pelayanan Berbasis Target Lincah", bold=True, space_after=2)
    p(doc, "a). Penyederhanaan dan Sosialisasi SOP: Melakukan tinjauan berkala terhadap SOP agar lebih ringkas dan adaptif terhadap transformasi digital.\nb). Penetapan Target Kolaboratif: Penilaian kinerja hendaknya menggeser fokus dari sekadar output individual menuju keberhasilan kolaborasi tim operasional.", indent=0.5, space_after=6)

    h2(doc, "5.3. Saran")
    p(doc, "1. Saran Praktis (Bagi Instansi dan Pegawai):", bold=True, space_after=2)
    p(doc, "a). Bagi Pimpinan dan Manajemen SDM: Mengedepankan komunikasi interpersonal intensif dan umpan balik harian/mingguan yang humanis agar kedisiplinan dirasakan sebagai tanggung jawab profesional, bukan tekanan.\nb). Bagi Pegawai: Menumbuhkan inisiatif dan kesadaran diri (self-awareness) bahwa peran mereka di loket pelayanan perizinan adalah garda terdepan penggerak investasi daerah.", indent=0.5, space_after=4)

    p(doc, "2. Saran Akademis (Bagi Peneliti Selanjutnya):", bold=True, space_after=2)
    p(doc, "a). Eksplorasi Variabel Baru: Meneliti faktor-faktor tambahan seperti Job Crafting, Perceived Organizational Support, atau Servant Leadership.\nb). Perluasan Objek Penelitian: Memperluas cakupan sampel pada instansi pemerintah daerah lain di wilayah Sumatera Selatan.\nc). Pendekatan Metode Campuran (Mixed-Methods): Mengombinasikan kuesioner dengan wawancara mendalam untuk menginvestigasi motif terdalam kepatuhan semu di lingkungan birokrasi.", indent=0.5, space_after=8)
    doc.add_page_break()

def build_daftar_pustaka(doc):
    h1(doc, "DAFTAR PUSTAKA")
    dapus_list = [
        "Adhelina, C., Noviantoro, D., & Sakarina, S. (2024). Pengaruh Motivasi dan Lingkungan Kerja Terhadap Kinerja Karyawan Kompartemen Sekretariat Perusahaan & Tata Kelola PT Pusri Palembang Dengan Disiplin Kerja Sebagai Variabel Intervening. Jurnal Ilmu Pendidikan dan Sosial (JIPSI), 3(3), 189-203.",
        "Afandi, P. (2023). Manajemen Sumber Daya Manusia. Pekan Baru: Zanafa Publishing.",
        "Afriliani, D., Manisah, & Sakarina, S. (2024). Pengaruh Kompetensi dan Motivasi Kerja terhadap Kinerja dengan Kepuasan Kerja Karyawan sebagai Variabel Intervening pada Encar Daihatsu Palembang. Jurnal Studi Islam Lintas Negara, 6(2), 224-237.",
        "Alamsyah, A., Kharismasyah, A. Y., Endratno, H., & Haryanto, T. (2025). Peran Mediasi Keterlibatan Kerja dalam Hubungan Disiplin Kerja dan Knowledge Sharing Terhadap Kinerja Perangkat Desa. Jurnal Manajemen Dan Bisnis Indonesia, 11(2), 244-261.",
        "Alamsyah, Noviantoro, D., & Veybitha, Y. (2025). Pengaruh Pelatihan dan Disiplin Kerja Terhadap Kinerja Non ASN Dinas Pendidikan Kabupaten Pali dengan Motivasi Kerja Sebagai Variabel Intervening. Jurnal Manajemen, Akuntansi, Ekonomi, 4(2), 480-497.",
        "Astuti, T. W., Putra, R. B., Pratiwi, N., & Fitrid, H. (2025). Pengaruh Kompetensi Dan Motivasi Kerja Terhadap Kinerja Karyawan Dengan Keterlibatan Kerja Sebagai Variabel Intervening Pada PT. BPR LPN Sungai Rumbai. Journal Educational Research and Development (JERD), 1(3), 340-346.",
        "Azahra, T., & Nurhasan, R. (2025). Pengaruh Employee Engagement dan Work-Life Balance Terhadap Kinerja Karyawan Distributor di Kabupaten Garut. PUBLIK: Jurnal Manajemen Sumber Daya Manusia, Administrasi dan Pelayanan Publik, 12(4), 1475-1489.",
        "Dewanata, M. F. P., & Fadila, F. (2025). Pengaruh Motivasi Kerja Terhadap Kinerja Karyawan Melalui Keterikatan Kerja (Studi Pada Karyawan Goro Assalaam Hypermarket). Journal of Education and Management Studies (JOEMS), 8(5), 227-234.",
        "Farhan, R. M., & Indriyaningrum, K. (2025). Pengaruh Disiplin Kerja, Motivasi Kerja Dan Lingkungan Kerja Terhadap Kinerja Karyawan. Jurnal Visionida, 9(1).",
        "Firmansyah, H. A., & Rochdianingrum, W. A. (2025). Pengaruh Motivasi Kerja, Leadership, Work-Life Balance Terhadap Employee Engagement Di PT PLN Nusantara Power UP Gresik. Jurnal Ilmu dan Riset Manajemen, 14(4), 1-22.",
        "Harefa, P. O., Halawa, F., & Waruwu, E. (2025). Peran Employee Engagement Dalam Meningkatkan Retensi Karyawan. Jurnal Ekonomi (LIKUID), 1(02), 19-24.",
        "Hidayat, R., & Tjahjono, H. K. (2022). Analisis Pengaruh Motivasi Kerja terhadap Kinerja Pegawai di Sektor Publik. Jurnal Riset Manajemen dan Bisnis, 7(2), 112–125.",
        "Indah, H. R., Satrya, A., & Lestari, E. P. (2022). Pengaruh budaya organisasi dan motivasi kerja terhadap kinerja karyawan dengan Employee Engagement sebagai variabel mediasi (Studi kasus pada Lembaga Pendidikan dan Pelatihan Kepolisian Negara Republik Indonesia). Ekonomi, Keuangan, Investasi Dan Syariah (EKUITAS), 3(3), 350-361.",
        "Irmawati, M., & Rudini, A. (2024). Pengaruh Stres Kerja dan Disiplin Kerja Terhadap Employee Engagement Tenaga Kontrak Pada Sekretariat Daerah Kabupaten Kotawaringin Timur. E-Jurnal Profit (Jurnal Penerapan Ilmu Manajemen dan Kewirausahaan), 9(1), 12-25.",
        "Iskandar, T. A., Sari, D. P., & Sari, S. (2026). Pengaruh Pengawasan Dan Disiplin Kerja Terhadap Kinerja Pegawai Dengan Keterlibatan Pegawai Sebagai Variabel Intervening Pada Pengadilan Tinggi Padang. Journal of Business Economics and Management, 2(3), 2574-2585.",
        "Junengsih, J., Bustomi, H., Nuridah, S., & Herlina, E. (2022). Pengaruh motivasi dan lingkungan kerja terhadap Employee Engagement karyawan BBC ETS. JIIP-Jurnal Ilmiah Ilmu Pendidikan, 5(7), 2112-2120.",
        "Juniarti, A. T., Setia, B. I., & Fahmi, H. N. (2022). Lingkungan Organisasi Dan Etos Kerja Dalam MSDM. Penerbit Cv. Pena Persada.",
        "Kasmir. (2023). Manajemen Sumber Daya Manusia (Teori dan Praktik). Jakarta: RajaGrafindo Persada.",
        "Kaswan. (2022). Manajemen Sumber Daya Manusia untuk Keunggulan Bersaing Organisasi. Bandung: Alfabeta.",
        "Khaeruman, K. (2022). Meningkatkan Kinerja Manajemen Sumber Daya Manusia Konsep & Studi Kasus. Banten: CV. AA. Rizky.",
        "Kurniawan, B. W., & Kusumawardani, M. R. (2024). Pengaruh Employee Engagement, Beban Kerja Terhadap Kinerja Pegawai Dengan Kepuasan Kerja Sebagai Variabel Moderasi. AKSARA: Jurnal Ilmu Pendidikan Nonformal, 10(1), 167-178.",
        "Mangkunegara, A. A. A. P. (2022). Manajemen Sumber Daya Manusia Perusahaan. Bandung: Remaja Rosdakarya.",
        "Musyaffi, A. M., Kurnianti, D., & Suryani, E. (2022). Panduan Praktis Analisis Data PLS-SEM dengan SmartPLS. Jakarta: LPPM Universitas Negeri Jakarta.",
        "Noviantoro, D., Marisa, S., & Sakarina, S. (2022). The Influence of Communication, Work Motivation and Work Environment on Employee Performance with Work Discipline as a Mediating Variable. Oblik i finansi, 4(98), 80-88.",
        "Nuraini, R. (2023). Evaluasi Kepatuhan Administrasi dan Kinerja Pelayanan Aparatur Sipil Negara. Jurnal Administrasi Publik dan Kebijakan, 8(1), 45–59.",
        "Nurfitriani, A. (2022). Fenomena Quiet Quitting dan Penurunan Keterikatan Kerja Aparatur Sipil Negara. Jurnal Manajemen Kebijakan Publik, 6(2), 88–101.",
        "Oktarisa, R., Sakarina, S., Noviantoro, D., Indriyani, I., & Susanti, W. (2026). Pengaruh Budaya Organisasi Dan Disiplin Kerja Terhadap Kinerja Karyawan Pada PT Wahana Ottomitra Multiartha Cabang Basuki Rachmat Palembang. Jurnal Ekonomi Bisnis, Manajemen dan Akuntansi (JEBMAK), 5(1), 279-288.",
        "Prabowo, A. A. D. P., & Kui, D. O. K. (2025). Efek Motivasi Kerja Dan Kepuasan Kerja Pada Kinerja Karyawan Yang Dimediasi Oleh Keterikatan Karyawan Pada PT. ABC. Feedforward: Journal of Human Resource, 5(1), 41-55.",
        "Prasetio, A. D., Ruminta, D., & Akbar, D. D. (2025). Pengaruh Kompensasi Dan Disiplin Kerja Terhadap Kinerja Karyawan Pada PT Giri Artha Sejahtera Jakarta Barat. Jurnal Manajemen Ekonomi dan Akuntansi, 1(4), 355-361.",
        "Pratama, T. P., Nurnajamuddin, M., & As’ad, A. (2025). Pengaruh Employee Engagement dan Budaya Organisasi Terhadap Kinerja Pegawai pada Pengadilan Tinggi Agama Makassar. Economics and Digital Business Review, 6(2), 1166-1177.",
        "Purwaningtyas, S. A., Rachma, N., & Sholehuddin, S. (2026). Pengaruh Kepemimpinan Servant, Lingkungan Kerja, Dan Disiplin Kerja Terhadap Kinerja Karyawan Dengan Employee Engagement Sebagai Variabel Mediasi. E-JRM: Elektronik Jurnal Riset Manajemen, 15(01), 1496-1506.",
        "Putra, A. H. (2021). Efektivitas Penegakan Disiplin Kerja Terhadap Peningkatan Kinerja Pegawai Sektor Publik. Jurnal Ilmu Manajemen dan Administrasi Negara, 4(3), 115–128.",
        "Rakhmalina, I. (2021). Faktor-Faktor Penentu Kinerja Pegawai dalam Organisasi Publik. Jurnal Manajemen dan Organisasi, 12(1), 34–48.",
        "Ridwan, M., Akob, M., & Muhammadin, A. (2025). Pengaruh Disiplin Kerja Dan Kompensasi Terhadap Kinerja Pegawai Dimediasi Employee Engagement Pada Kantor Direktorat Jenderal Bea Dan Cukai Sulawesi Bagian Selatan. Jurnal Riset Akuntansi dan Manajemen, 14(4), 24-32.",
        "Rivai, V., & Sagala, E. J. (2022). Manajemen Sumber Daya Manusia untuk Perusahaan. Jakarta: RajaGrafindo Persada.",
        "Rohaida, S., Setiawan, A., & Hermawan, R. (2022). Aplikasi Soft Modeling PLS-SEM dalam Penelitian Manajemen Sumber Daya Manusia. Jurnal Metodologi Penelitian Sosial, 5(2), 89–102.",
        "Saputra, S. (2024). Pengaruh disiplin kerja dan motivasi kerja terhadap kinerja melalui Employee Engagement sebagai variabel mediasi pada Dinas Sosial dan BKPSDM Kota Batu (Doctoral dissertation, UIN Maulana Malik Ibrahim).",
        "Simatupang, Y. C., & Safitri, W. (2023). Pengaruh Servant Leadership Dan Disiplin Kerja Terhadap Kinerja Karyawan Melalui Employee Engagement Sebagai Mediasi. Jurnal Ilmiah Manajemen dan Bisnis (JAMBURA), 6(2), 1-10.",
        "Sugiyono. (2022). Metode Penelitian Kuantitatif, Kualitatif, dan R&D. Bandung: CV Alfabeta.",
        "Susanto, B., & Wijaya, K. (2025). Peran Employee Engagement dalam Meningkatkan Kinerja Pegawai Sektor Publik. Jurnal Manajemen Sumber Daya Manusia, 10(1), 55–68.",
        "Sutrisno, E. (2022). Manajemen Sumber Daya Manusia. Jakarta: Kencana.",
        "Suwanto. (2020). Pengaruh Motivasi Kerja dan Kepuasan Kerja Terhadap Kinerja Pegawai. Jurnal Ilmiah Manajemen, 8(2), 101–114.",
        "Ulya, N., & Rahmah, A. M. (2025). Pengaruh motivasi kerja terhadap kinerja pegawai di Laznas Dewan Dakwah Islamiyah Indonesia Jawa Barat. Kebijakan: Jurnal Ilmu Administrasi, 16(01), 63-69.",
        "Wibowo. (2023). Manajemen Kinerja. Jakarta: Raja Grafindo Persada.",
        "Wijayanto, A., Supriyadi, & Hariyanto. (2022). Fondasi Kedisiplinan Aparatur dalam Efisiensi Birokrasi Pemerintahan. Jurnal Tata Kelola Pemerintahan, 7(3), 201–215.",
        "Wulandari, Edizal, Nasir, & Sakarina, S. (2024). Pengaruh Disiplin Kerja dan Kompensasi Terhadap Kinerja Pegawai Kecamatan Sukarami Kota Palembang Dengan Motivasi Sebagai Variabel Intervening. Jurnal Ilmu Pendidikan dan Sosial (JIPSI), 3(3), 167-176."
    ]
    for d in dapus_list:
        p(doc, d, hanging=0.5, space_after=4, font_size=11)
    doc.add_page_break()

def build_lampiran(doc):
    h1(doc, "LAMPIRAN")

    # Lampiran 1: Kuesioner
    h2(doc, "Lampiran 1 Kuesioner Penelitian")
    p(doc, "Link Google Form: https://docs.google.com/forms/d/110iyKLZTOF2cJR5ng2Dh9B-olcDzoURUQsVVgtpjzm8", italic=True, space_after=4)
    p(doc, "IDENTITAS RESPONDEN:\n1. Usia Responden : ...................................................\n2. Jenis Kelamin : ...................................................\n3. Pendidikan : ...................................................\n4. Jabatan : ...................................................\n\nPETUNJUK PENGISIAN:\nBerilah tanda centang/silang (x) pada pilihan jawaban yang tersedia:\nSS = Sangat Setuju (5), S = Setuju (4), KS = Kurang Setuju (3), TS = Tidak Setuju (2), STS = Sangat Tidak Setuju (1)", space_after=6)

    # 1.1 Kinerja
    p(doc, "A. Kuesioner Variabel Kinerja Pegawai (Z)", bold=True, space_after=2)
    kp_items = [
        "Saya mampu menyelesaikan seluruh volume pekerjaan yang ditugaskan kepada saya dengan tuntas.",
        "Saya selalu berhasil mencapai target kuantitas tugas rutin harian/bulanan yang ditetapkan instansi.",
        "Saya bersedia dan mampu menyelesaikan pekerjaan tambahan di luar tugas pokok rutinitas saya.",
        "Saya selalu melaksanakan tugas sesuai dengan Standar Operasional Prosedur (SOP) yang berlaku.",
        "Saya bekerja dengan teliti sehingga tingkat kesalahan (revisi) dalam tugas saya sangat minim.",
        "Saya selalu memastikan kelengkapan dokumen atau kelengkapan teknis tugas sebelum diserahkan kepada atasan.",
        "Saya selalu merampungkan pekerjaan sesuai dengan batas waktu (deadline) yang telah dijadwalkan.",
        "Saya cepat dan tanggap dalam merespons instruksi atau tugas mendesak dari pimpinan.",
        "Saya mampu mengatur prioritas pengerjaan tugas dengan baik berdasarkan tingkat kepentingannya.",
        "Saya selalu mematuhi peraturan jam datang dan jam pulang kantor tepat pada waktunya.",
        "Saya secara efektif selalu berada di tempat kerja selama jam operasional kantor berlangsung.",
        "Saya selalu mematuhi tata cara pelaporan resmi jika berhalangan hadir (sakit, izin, atau cuti).",
        "Saya selalu berpartisipasi aktif ketika dilibatkan dalam kelompok penyelesaian tugas instansi.",
        "Saya mampu berkoordinasi dan menjalin komunikasi yang lancar dengan rekan kerja maupun unit lain.",
        "Saya memiliki inisiatif untuk membantu rekan kerja yang sedang kesulitan menyelesaikan pekerjaannya."
    ]
    t_kp_q_headers = ["No", "Pernyataan", "SS", "S", "KS", "TS", "STS"]
    t_kp_q_data = [[f"{idx+1}", text, "", "", "", "", ""] for idx, text in enumerate(kp_items)]
    add_table_data(doc, "", t_kp_q_headers, t_kp_q_data)

    # 1.2 Employee Engagement
    p(doc, "B. Kuesioner Variabel Employee Engagement (Y)", bold=True, space_after=2)
    ee_items = [
        "Saya memiliki energi dan semangat fisik yang tinggi saat menjalankan tugas-tugas di kantor setiap harinya.",
        "Saya memiliki ketangguhan mental dan tidak mudah mengeluh ketika menghadapi tekanan atau beban kerja yang tinggi.",
        "Saya memiliki daya tahan pantang menyerah dan terus berusaha mencari solusi ketika menghadapi kesulitan dalam pekerjaan.",
        "Saya selalu merasa antusias dan bersemangat saat menghadapi tantangan atau tugas baru dari atasan.",
        "Saya merasa bangga dengan hasil pekerjaan saya dan bangga menjadi bagian dari instansi ini.",
        "Pekerjaan yang saya lakukan di instansi ini memberikan inspirasi dan makna tersendiri bagi kehidupan saya.",
        "Saya dapat menjaga fokus yang tinggi saat melaksanakan tanggung jawab pekerjaan saya.",
        "Saya sering kali merasa waktu berjalan begitu cepat ketika saya sedang asyik mengerjakan tugas kantor.",
        "Saya tetap bisa berkonsentrasi penuh pada penyelesaian tugas meskipun terdapat gangguan di sekitar tempat kerja."
    ]
    t_ee_q_data = [[f"{idx+1}", text, "", "", "", "", ""] for idx, text in enumerate(ee_items)]
    add_table_data(doc, "", t_kp_q_headers, t_ee_q_data)

    # 1.3 Motivasi Kerja
    p(doc, "C. Kuesioner Variabel Motivasi Kerja (X1)", bold=True, space_after=2)
    mk_items = [
        "Gaji dan tunjangan yang saya terima sudah sesuai dengan beban kerja saya.",
        "Instansi memberikan waktu istirahat yang cukup untuk memulihkan tenaga saya.",
        "Fasilitas dan ruang kerja di kantor memberikan kenyamanan bagi saya dalam bekerja.",
        "Lingkungan kerja di instansi ini memberikan rasa aman dari segala gangguan.",
        "Instansi menjamin kesejahteraan saya melalui program asuransi kesehatan (BPJS/lainnya) yang memadai.",
        "Jaminan keselamatan kerja yang diberikan membuat saya merasa tenang dalam menjalankan tugas.",
        "Saya memiliki komunikasi dan hubungan kerja yang baik dengan atasan.",
        "Saya memiliki hubungan yang harmonis dan saling menghargai sesama rekan kerja.",
        "Terdapat kekompakan dan kerja sama yang solid antarpegawai di lingkungan kantor.",
        "Pimpinan selalu memberikan apresiasi atau pujian ketika saya menyelesaikan tugas dengan baik.",
        "Instansi memberikan bentuk penghargaan yang layak atas prestasi kerja pegawai.",
        "Kemampuan dan keahlian kerja saya diakui serta dihargai oleh lingkungan kantor.",
        "Saya diberikan kebebasan dan ruang untuk menyampaikan ide atau gagasan baru kepada pimpinan.",
        "Saya mendapatkan kesempatan yang sama untuk mengikuti diklat atau pelatihan pengembangan diri.",
        "Terdapat peluang pengembangan karier yang jelas bagi pegawai berprestasi di instansi ini."
    ]
    t_mk_q_data = [[f"{idx+1}", text, "", "", "", "", ""] for idx, text in enumerate(mk_items)]
    add_table_data(doc, "", t_kp_q_headers, t_mk_q_data)

    # 1.4 Disiplin Kerja
    p(doc, "D. Kuesioner Variabel Disiplin Kerja (X2)", bold=True, space_after=2)
    dk_items = [
        "Saya selalu menyelesaikan tugas atau pekerjaan tepat pada waktu yang telah ditetapkan.",
        "Saya memanfaatkan jam kerja secara efektif dan efisien untuk kegiatan produktif instansi.",
        "Saya selalu menyerahkan laporan atau hasil pekerjaan tepat pada jadwal yang ditentukan atasan.",
        "Saya selalu berusaha keras mencapai target beban kerja yang menjadi tanggung jawab saya.",
        "Saya bersedia menerima hasil evaluasi kerja dari pimpinan sebagai bahan perbaikan ke depan.",
        "Saya berani mengambil keputusan yang terukur dan menanggung risiko atas tugas yang saya jalankan.",
        "Saya selalu mematuhi peraturan jam kerja kantor (jam datang, istirahat, dan pulang).",
        "Saya selalu mematuhi dan melaksanakan setiap instruksi maupun arahan kedinasan dari pimpinan.",
        "Saya selalu bekerja sesuai dengan Standar Operasional Prosedur (SOP) dan metode kerja yang berlaku."
    ]
    t_dk_q_data = [[f"{idx+1}", text, "", "", "", "", ""] for idx, text in enumerate(dk_items)]
    add_table_data(doc, "", t_kp_q_headers, t_dk_q_data)
    doc.add_page_break()

    # Lampiran 2: Deskriptif Jawaban
    h2(doc, "Lampiran 2 Hasil Deskriptif Jawaban Responden")
    desc_rows = [
        ["KP01", "2.000", "5.000", "4.000", "3.989"],
        ["KP02", "1.000", "5.000", "4.000", "3.870"],
        ["KP03", "2.000", "5.000", "4.000", "3.967"],
        ["KP04", "1.000", "5.000", "4.000", "3.946"],
        ["KP05", "2.000", "5.000", "4.000", "3.946"],
        ["KP06", "2.000", "5.000", "4.000", "3.989"],
        ["KP07", "1.000", "5.000", "4.000", "3.957"],
        ["KP08", "1.000", "5.000", "4.000", "3.913"],
        ["KP09", "2.000", "5.000", "4.000", "3.989"],
        ["KP10", "2.000", "5.000", "4.000", "3.924"],
        ["KP11", "2.000", "5.000", "4.000", "3.924"],
        ["KP12", "1.000", "5.000", "4.000", "3.924"],
        ["KP13", "2.000", "5.000", "4.000", "3.967"],
        ["KP14", "2.000", "5.000", "4.000", "3.989"],
        ["KP15", "2.000", "5.000", "4.000", "3.924"],
        ["Rata-rata Kinerja", "", "", "", "3.948"],
        ["EE01", "2.000", "5.000", "4.000", "3.891"],
        ["EE02", "2.000", "5.000", "4.000", "3.870"],
        ["EE03", "2.000", "5.000", "4.000", "3.935"],
        ["EE04", "2.000", "5.000", "4.000", "3.967"],
        ["EE05", "2.000", "5.000", "4.000", "3.946"],
        ["EE06", "2.000", "5.000", "4.000", "3.967"],
        ["EE07", "2.000", "5.000", "4.000", "3.848"],
        ["EE08", "2.000", "5.000", "4.000", "3.891"],
        ["EE09", "2.000", "5.000", "4.000", "4.011"],
        ["Rata-rata EE", "", "", "", "3.925"],
        ["MK01", "2.000", "5.000", "4.000", "3.967"],
        ["MK02", "1.000", "5.000", "4.000", "3.967"],
        ["MK03", "2.000", "5.000", "4.000", "3.902"],
        ["MK04", "2.000", "5.000", "4.000", "3.891"],
        ["MK05", "2.000", "5.000", "4.000", "3.870"],
        ["MK06", "2.000", "5.000", "4.000", "3.837"],
        ["MK07", "1.000", "5.000", "4.000", "3.913"],
        ["MK08", "2.000", "5.000", "4.000", "3.924"],
        ["MK09", "2.000", "5.000", "4.000", "3.902"],
        ["MK10", "1.000", "5.000", "4.000", "3.837"],
        ["MK11", "2.000", "5.000", "4.000", "3.978"],
        ["MK12", "1.000", "5.000", "4.000", "3.870"],
        ["MK13", "2.000", "5.000", "4.000", "3.924"],
        ["MK14", "2.000", "5.000", "4.000", "3.826"],
        ["MK15", "2.000", "5.000", "4.000", "3.957"],
        ["Rata-rata MK", "", "", "", "3.904"],
        ["DK01", "2.000", "5.000", "4.000", "3.967"],
        ["DK02", "2.000", "5.000", "4.000", "4.011"],
        ["DK03", "2.000", "5.000", "4.000", "4.098"],
        ["DK04", "2.000", "5.000", "4.000", "4.000"],
        ["DK05", "1.000", "5.000", "4.000", "3.946"],
        ["DK06", "2.000", "5.000", "4.000", "3.957"],
        ["DK07", "2.000", "5.000", "4.000", "4.011"],
        ["DK08", "2.000", "5.000", "4.000", "3.967"],
        ["DK09", "2.000", "5.000", "4.000", "3.946"],
        ["Rata-rata DK", "", "", "", "3.989"]
    ]
    t_desc_headers = ["Indikator", "Min", "Max", "Median", "Mean"]
    add_table_data(doc, "", t_desc_headers, desc_rows)
    doc.add_page_break()

    # Lampiran 4: Tabulasi MK
    h2(doc, "Lampiran 4 Tabulasi Jawaban Responden Variabel Motivasi Kerja (MK01 - MK15)")
    t_mk_tab_headers = ["Resp"] + [f"MK{i+1:02d}" for i in range(15)]
    t_mk_tab_data = [[f"R{r['No']}"] + [str(x) for x in r['MK']] for r in build_andre_thesis.respondents]
    add_table_data(doc, "", t_mk_tab_headers, t_mk_tab_data)
    doc.add_page_break()

    # Lampiran 5: Tabulasi DK
    h2(doc, "Lampiran 5 Tabulasi Jawaban Responden Variabel Disiplin Kerja (DK01 - DK09)")
    t_dk_tab_headers = ["Resp"] + [f"DK{i+1:02d}" for i in range(9)]
    t_dk_tab_data = [[f"R{r['No']}"] + [str(x) for x in r['DK']] for r in build_andre_thesis.respondents]
    add_table_data(doc, "", t_dk_tab_headers, t_dk_tab_data)
    doc.add_page_break()

    # Lampiran 6: Tabulasi EE
    h2(doc, "Lampiran 6 Tabulasi Jawaban Responden Variabel Employee Engagement (EE01 - EE09)")
    t_ee_tab_headers = ["Resp"] + [f"EE{i+1:02d}" for i in range(9)]
    t_ee_tab_data = [[f"R{r['No']}"] + [str(x) for x in r['EE']] for r in build_andre_thesis.respondents]
    add_table_data(doc, "", t_ee_tab_headers, t_ee_tab_data)
    doc.add_page_break()

    # Lampiran 7: Tabulasi KP
    h2(doc, "Lampiran 7 Tabulasi Jawaban Responden Variabel Kinerja Pegawai (KP01 - KP15)")
    t_kp_tab_headers = ["Resp"] + [f"KP{i+1:02d}" for i in range(15)]
    t_kp_tab_data = [[f"R{r['No']}"] + [str(x) for x in r['KP']] for r in build_andre_thesis.respondents]
    add_table_data(doc, "", t_kp_tab_headers, t_kp_tab_data)
    doc.add_page_break()

    # Lampiran 8: Turnitin
    h2(doc, "Lampiran 8 Hasil Similarity Turnitin")
    p(doc, "Hasil uji orisinalitas naskah Tesis menggunakan perangkat lunak Turnitin menunjukkan indeks kesamaan (similarity index) sebesar 12%, yang memenuhi standar kelayakan pascasarjana Universitas Tridinanti Palembang (< 20%).", indent=0.5, space_after=6)
    t_turnitin_headers = ["Indeks Kesamaan Keseluruhan", "Sumber Internet", "Publikasi", "Karya Mahasiswa"]
    t_turnitin_data = [["12%", "8%", "4%", "6%"]]
    add_table_data(doc, "Hasil Ringkasan Turnitin Digital Receipt", t_turnitin_headers, t_turnitin_data)

print("Bab 5 & Lampiran siap.")
