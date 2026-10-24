# backend/thesis_bab3_bab4.py
"""Modul Pembangun BAB III dan BAB IV Tesis Andre Juliansya"""

from docx.enum.text import WD_ALIGN_PARAGRAPH
from build_full_thesis_docx import p, h1, h2, h3, add_table_data, rapikan_tabel_dokumen

def build_bab_3(doc):
    h1(doc, "BAB III\nMETODE PENELITIAN")
    
    h2(doc, "3.1. Tempat dan Waktu Penelitian")
    p(doc, "Penelitian ini dilaksanakan di Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan. Waktu penelitian selama 6 (enam) bulan mulai dari proses perencanaan, pelaksanaan, dan pelaporan hasil penelitian. Penelitian dilakukan pada bulan Mei 2026 sampai dengan bulan Oktober 2026. Supaya penelitian dapat dilaksanakan secara sistematis dan berkesinambungan, maka disusun jadwal penelitian sebagai berikut:", indent=0.5, space_after=6)

    # Tabel 3.1
    t31_headers = ["No", "Kegiatan", "Mei 2026", "Juni 2026", "Juli 2026", "Agst 2026", "Sept 2026", "Okt 2026"]
    t31_data = [
        ["1", "Penyusunan proposal", "V", "", "", "", "", ""],
        ["2", "Uji coba instrumen", "", "V", "", "", "", ""],
        ["3", "Pengumpulan data penelitian", "", "", "V", "", "", ""],
        ["4", "Pengolahan data", "", "", "", "V", "", ""],
        ["5", "Analisis data", "", "", "", "V", "", ""],
        ["6", "Menulis laporan penelitian", "", "", "", "", "V", ""],
        ["7", "Bimbingan Tesis", "", "", "", "", "V", "V"],
        ["8", "Sidang Tesis", "", "", "", "", "", "V"]
    ]
    add_table_data(doc, "Tabel 3.1\nJadwal Penelitian", t31_headers, t31_data, "Sumber : Diolah Penulis, (2026)")

    h2(doc, "3.2. Populasi dan Sampel")
    h3(doc, "3.2.1 Populasi")
    p(doc, "Populasi merupakan keseluruhan objek atau subjek yang memiliki karakteristik tertentu yang ditetapkan oleh peneliti untuk dipelajari dan kemudian ditarik kesimpulannya (Sugiyono, 2022). Setiap objek atau subjek yang menjadi bagian dalam populasi disebut sebagai unit analisis atau elemen populasi. Ditetapkan populasi dalam setiap penelitian bertujuan agar dapat menentukan besar sampel yang diambil dari anggota populasi dan membatasi berlakunya daerah generalisasi. Populasi dalam penelitian ini adalah Pegawai Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan, yang berjumlah 119 orang.", indent=0.5, space_after=6)

    h3(doc, "3.2.2. Sampel")
    p(doc, "Sampel adalah bagian dari jumlah dan karakteristik yang dimiliki oleh populasi tersebut. Pengambilan sampel dalam penelitian ini dilakukan karena keterbatasan waktu, tenaga, dan dana, namun tetap harus mampu mewakili seluruh populasi yang ada (Sugiyono 2022).", indent=0.5, space_after=4)
    p(doc, "Untuk menentukan ukuran sampel yang representatif dari total populasi Pegawai Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan yang berjumlah 119 orang, digunakan Rumus Slovin sebagai berikut:\n\nn = N / (1 + N(e)^2)\n\nKeterangan:\nn = Ukuran sampel\nN = Ukuran populasi (119 orang)\ne = Persentase kelonggaran ketidaktelitian (5% atau 0,05)\n\nPerhitungan ukuran sampel:\nn = 119 / (1 + 119(0,05)^2)\nn = 119 / (1 + 119(0,0025))\nn = 119 / (1 + 0,2975)\nn = 119 / 1,2975 = 91,71\n\nBerdasarkan hasil perhitungan di atas, diperoleh nilai sebesar 91,71 yang kemudian dibulatkan ke atas menjadi 92 orang responden.", indent=0.5, space_after=6)

    h2(doc, "3.3. Variabel dan Definisi Operasional")
    p(doc, "Sesuai masalah dan tujuan penelitian, variabel penelitian ini terdiri dari :\n1. Variabel endogen yaitu Kinerja\n2. Variabel Mediasi yaitu Employee Engagement\n3. Variabel eksogen yaitu meliputi Motivasi Kerja dan Disiplin Kerja\n\nUraian masing-masing variabel penelitian ini adalah sebagai berikut :", space_after=4)

    # Variabel Kinerja
    h3(doc, "1. Variabel Kinerja (Z)")
    p(doc, "a. Definisi Konseptual: Mangkunegara (2022) menjelaskan bahwa kinerja pegawai adalah hasil kerja yang dicapai baik secara kualitas maupun kuantitas dalam melaksanakan tugas sesuai dengan tanggung jawab yang diberikan.", indent=0.5, space_after=2)
    p(doc, "b. Definisi Operasional: Kinerja pegawai dalam penelitian ini dijelaskan sebagai tingkat keberhasilan dan kemampuan pegawai Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan dalam melaksanakan tugas dan tanggung jawabnya.", indent=0.5, space_after=4)
    t32_headers = ["No", "Dimensi", "Indikator", "No. Item Pertanyaan"]
    t32_data = [
        ["1", "Jumlah Pekerjaan (Kuantitas)", "1. Pencapaian volume pekerjaan\n2. Pencapaian target tugas rutin\n3. Penyelesaian tugas tambahan", "KP01\nKP02\nKP03"],
        ["2", "Kualitas Pekerjaan", "1. Kesesuaian kerja dengan SOP\n2. Tingkat ketelitian kerja\n3. Kelengkapan hasil pekerjaan", "KP04\nKP05\nKP06"],
        ["3", "Ketepatan Waktu", "1. Penyelesaian tugas sesuai jadwal\n2. Kecepatan merespons tugas\n3. Prioritas waktu pengerjaan", "KP07\nKP08\nKP09"],
        ["4", "Kehadiran (Absensi)", "1. Kepatuhan jam masuk dan pulang\n2. Kehadiran efektif di kantor\n3. Kepatuhan prosedur izin/cuti", "KP10\nKP11\nKP12"],
        ["5", "Kemampuan Kerja Sama", "1. Keterlibatan aktif dalam tim\n2. Kelancaran koordinasi antarpegawai\n3. Inisiatif membantu rekan kerja", "KP13\nKP14\nKP15"]
    ]
    add_table_data(doc, "Tabel 3.2\nKisi-Kisi Instrumen Variabel Kinerja", t32_headers, t32_data, "Sumber: Ulya & Rahmah (2025)")

    # Variabel Employee Engagement
    h3(doc, "2. Variabel Employee Engagement (Y)")
    p(doc, "a. Definisi Konseptual: Rivai & Sagala (2022), menjelaskan bahwa Employee Engagement merupakan keadaan ketika individu menyatukan aspek emosional, kognitif, dan fisiknya ke dalam peran kerja yang dijalankan.", indent=0.5, space_after=2)
    p(doc, "b. Definisi Operasional: Employee Engagement dalam penelitian ini dijelaskan sebagai tingkat keterikatan emosional dan psikologis pegawai Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan terhadap pekerjaan dan instansinya.", indent=0.5, space_after=4)
    t33_headers = ["No", "Dimensi", "Indikator", "No. Item Pertanyaan"]
    t33_data = [
        ["1", "Vigor", "1. Energi kerja\n2. Ketangguhan\n3. Daya tahan", "EE01\nEE02\nEE03"],
        ["2", "Dedication", "1. Antusiasme\n2. Kebanggaan\n3. Inspirasi", "EE04\nEE05\nEE06"],
        ["3", "Absorption", "1. Fokus kerja\n2. Tenggelam dalam pekerjaan\n3. Konsentrasi", "EE07\nEE08\nEE09"]
    ]
    add_table_data(doc, "Tabel 3.3\nKisi-Kisi Instrumen Variabel Employee Engagement", t33_headers, t33_data, "Sumber : Harefa, Halawa, & Waruwu (2025)")

    # Variabel Motivasi Kerja
    h3(doc, "3. Motivasi Kerja (X1)")
    p(doc, "a. Definisi Konseptual: Sutrisno (2022) menjelaskan bahwa motivasi kerja merupakan dorongan yang berasal dari dalam maupun luar diri seseorang untuk melakukan suatu aktivitas atau pekerjaan tertentu.", indent=0.5, space_after=2)
    p(doc, "b. Definisi Operasional: Motivasi kerja dalam penelitian ini dijelaskan sebagai daya dorong internal maupun eksternal yang memicu semangat, kemauan, dan ketekunan pegawai Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan dalam menjalankan tugasnya.", indent=0.5, space_after=4)
    t34_headers = ["No", "Dimensi", "Indikator", "No. Item Pertanyaan"]
    t34_data = [
        ["1", "Kebutuhan Fisiologi", "1. Kesesuaian gaji/tunjangan\n2. Kecukupan waktu istirahat\n3. Kenyamanan fasilitas kerja", "MK01\nMK02\nMK03"],
        ["2", "Kebutuhan Rasa Aman", "1. Keamanan lingkungan kerja\n2. Jaminan kesehatan pegawai\n3. Jaminan keselamatan kerja", "MK04\nMK05\nMK06"],
        ["3", "Kebutuhan Merasa Memiliki (Sosial)", "1. Hubungan baik dengan atasan\n2. Hubungan harmonis antarpegawai\n3. Kekompakan rekan kerja", "MK07\nMK08\nMK09"],
        ["4", "Kebutuhan untuk Dihargai", "1. Apresiasi dari pimpinan\n2. Penghargaan atas prestasi\n3. Pengakuan atas keahlian", "MK10\nMK11\nMK12"],
        ["5", "Kebutuhan Mengaktualisasi Diri", "1. Kesempatan menyampaikan ide\n2. Pelatihan pengembangan diri\n3. Kesempatan pengembangan karier", "MK13\nMK14\nMK15"]
    ]
    add_table_data(doc, "Tabel 3.4\nKisi-Kisi Instrumen Variabel Motivasi Kerja", t34_headers, t34_data, "Sumber : Farhan & Indriyaningrum (2025)")

    # Variabel Disiplin Kerja
    h3(doc, "4. Variabel Disiplin Kerja (X2)")
    p(doc, "a. Definisi Konseptual: Wibowo (2023) menyatakan bahwa disiplin kerja merupakan kondisi yang terbentuk melalui serangkaian perilaku yang menunjukkan nilai-nilai ketaatan, kepatuhan, keteraturan, dan ketertiban dalam organisasi.", indent=0.5, space_after=2)
    p(doc, "b. Definisi Operasional: Disiplin kerja dalam penelitian ini dijelaskan sebagai tingkat kepatuhan, kesadaran, dan ketertiban pegawai Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan dalam menjalankan kewajibannya.", indent=0.5, space_after=4)
    t35_headers = ["No", "Dimensi", "Indikator", "No. Item Pertanyaan"]
    t35_data = [
        ["1", "Ketepatan Waktu", "1. Penyelesaian tugas tepat waktu\n2. Pemanfaatan jam kerja efektif\n3. Ketepatan waktu pelaporan hasil", "DK01\nDK02\nDK03"],
        ["2", "Tanggung Jawab Kerja", "1. Pencapaian target pekerjaan\n2. Kesiapan menerima evaluasi kerja\n3. Keberanian menanggung risiko tugas", "DK04\nDK05\nDK06"],
        ["3", "Ketaatan", "1. Kepatuhan pada aturan jam kerja\n2. Kepatuhan terhadap instruksi pimpinan\n3. Kepatuhan pada Standar Operasional Prosedur (SOP)", "DK07\nDK08\nDK09"]
    ]
    add_table_data(doc, "Tabel 3.5\nKisi-kisi Instrumen Variabel Disiplin Kerja", t35_headers, t35_data, "Sumber : Prasetio, Ruminta, & Akbar (2025)")

    h2(doc, "3.4 Instrumen Penelitian")
    p(doc, "Sesuai dengan variabel diatas, terdapat 4 (empat) jenis data yang dikumpulkan, antara lain: 1) data tentang Kinerja, 2) data tentang Employee Engagement, 3) data tentang Motivasi Kerja, dan 4) data tentang Disiplin Kerja. Data yang dibutuhkan dari setiap Pegawai ditampung pada instrumen pengumpulan data berupa kuesioner berskala Likert 1-5 (Sugiyono, 2022).", indent=0.5, space_after=4)
    t36_headers = ["Jawaban Responden", "Skor"]
    t36_data = [
        ["Sangat Setuju", "5"],
        ["Setuju", "4"],
        ["Kurang Setuju", "3"],
        ["Tidak Setuju", "2"],
        ["Sangat Tidak Setuju", "1"]
    ]
    add_table_data(doc, "Tabel 3.6\nScoring Untuk Jawaban Kuesioner", t36_headers, t36_data, "Sumber: Sugiyono, (2022)")

    h2(doc, "3.5. Teknik Analisis Data")
    p(doc, "Untuk membahas pertanyaan pokok dari penelitian, peneliti menggunakan uji statistik. Metode Partial Least Square (PLS) dengan SmartPLS 3.2.9 digunakan untuk pengelolaan data dalam penelitian ini. PLS adalah model persamaan struktural (SEM) berbasis varians yang mampu mengevaluasi kumpulan hubungan yang relatif menantang untuk diukur secara bersamaan. Rohaida et al. (2022) menjelaskan bahwa PLS merupakan metode soft modeling karena tidak mengasumsikan distribusi normal ketat dan dapat digunakan pada ukuran sampel kurang dari 100 sampel.", indent=0.5, space_after=4)
    p(doc, "Terdapat tiga tahapan utama analisis PLS-SEM (Musyaffi et al., 2022):\n1. Analisis Model Pengukuran (Outer Model):\n   a) Convergent Validity: Loading Factor > 0,70 dan AVE > 0,50.\n   b) Discriminant Validity: Kriteria Fornell-Larcker dan Cross Loading.\n   c) Reliabilitas: Composite Reliability > 0,70 dan Cronbach's Alpha > 0,70.\n2. Analisis Model Struktural (Inner Model):\n   a) R Square (R²): 0,67 (kuat), 0,33 (moderat), 0,19 (lemah).\n   b) Effect Size (f²): 0,02 (kecil), 0,15 (sedang), 0,35 (besar).\n   c) Path Coefficients.\n\nPersamaan jalur sub-struktur Pertama:\nY = ρy1 X1 + ρy2 X2 + ɛ1\n\nPersamaan jalur sub-struktur Kedua:\nZ = ρz1 X1 + ρz2 X2 + ρzy Y + ɛ2", indent=0.5, space_after=6)

    h2(doc, "3.6. Pengujian Hipotesis")
    p(doc, "Pengujian hipotesis dilakukan melalui prosedur Bootstrapping untuk menghasilkan nilai t-statistik dan p-value dengan tingkat kepercayaan 95% (alpha = 0,05). Hipotesis diterima apabila t-statistik > 1,96 dan p-value < 0,05.", indent=0.5, space_after=4)
    t37_headers = ["No", "Hipotesis", "Kriteria H0 ditolak", "Kriteria H0 diterima"]
    t37_data = [
        ["1.", "H1: Motivasi Kerja berpengaruh signifikan thd Employee Engagement", "p <= 0,05 atau t >= 1,96", "p >= 0,05 atau t <= 1,96"],
        ["2.", "H2: Disiplin Kerja berpengaruh signifikan thd Employee Engagement", "p <= 0,05 atau t >= 1,96", "p >= 0,05 atau t <= 1,96"],
        ["3.", "H3: Motivasi Kerja berpengaruh signifikan thd Kinerja Pegawai", "p <= 0,05 atau t >= 1,96", "p >= 0,05 atau t <= 1,96"],
        ["4.", "H4: Disiplin Kerja berpengaruh signifikan thd Kinerja Pegawai", "p <= 0,05 atau t >= 1,96", "p >= 0,05 atau t <= 1,96"],
        ["5.", "H5: Employee Engagement berpengaruh signifikan thd Kinerja Pegawai", "p <= 0,05 atau t >= 1,96", "p >= 0,05 atau t <= 1,96"],
        ["6.", "H6: Motivasi Kerja berpengaruh signifikan thd Kinerja melalui Employee Engagement", "p <= 0,05 atau t >= 1,96", "p >= 0,05 atau t <= 1,96"],
        ["7.", "H7: Disiplin Kerja berpengaruh signifikan thd Kinerja melalui Employee Engagement", "p <= 0,05 atau t >= 1,96", "p >= 0,05 atau t <= 1,96"]
    ]
    add_table_data(doc, "Tabel 3.7\nKriteria Keputusan Hasil Uji t-statistic", t37_headers, t37_data, "Sumber : Diolah Penulis, (2026)")
    rapikan_tabel_dokumen(doc)
    doc.add_page_break()

def build_bab_4(doc):
    h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
    
    h2(doc, "4.1. Hasil Analisis")
    p(doc, "Sebelum melakukan pengujian instrumen dan analisis model struktural, langkah awal yang dilakukan dalam penelitian ini adalah menyajikan analisis deskriptif terkait data demografis responden. Analisis karakteristik responden ini bertujuan untuk memberikan gambaran umum mengenai profil latar belakang pegawai pada Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Sumatera Selatan yang menjadi sampel dalam penelitian ini.", indent=0.5, space_after=6)

    # 1. Demografi
    h3(doc, "1. Deskriptif Data Demografis Responden")
    p(doc, "Karakteristik demografis pertama yang dianalisis adalah pengelompokan responden berdasarkan jenis kelamin. Pemetaan ini berguna untuk melihat proporsi keterwakilan gender dari populasi pegawai yang berpartisipasi dalam penelitian. Rincian distribusi responden berdasarkan jenis kelamin dapat dilihat pada Tabel 4.1 berikut:", indent=0.5, space_after=4)
    t41_headers = ["Karakteristik", "Keterangan", "Total Responden", "Persentase (%)"]
    t41_data = [
        ["Jenis Kelamin", "Laki-laki", "53", "57,60"],
        ["", "Perempuan", "39", "42,40"],
        ["Jumlah", "", "92", "100"]
    ]
    add_table_data(doc, "Tabel 4.1\nDeskripsi Karakteristik Responden berdasarkan Jenis Kelamin", t41_headers, t41_data, "Sumber : Data primer diolah, (2026)")
    p(doc, "Berdasarkan data yang disajikan pada Tabel 4.1 di atas, dapat diketahui bahwa dari total 92 orang responden yang menjadi sampel penelitian, sebagian besar didominasi oleh pegawai berjenis kelamin laki-laki, yaitu sebanyak 53 orang atau mencapai persentase 57,60%. Sementara itu, pegawai berjenis kelamin perempuan berjumlah 39 orang dengan persentase sebesar 42,40%.", indent=0.5, space_after=6)

    p(doc, "Karakteristik demografis selanjutnya yang dianalisis adalah usia responden:", indent=0.5, space_after=4)
    t42_headers = ["Karakteristik", "Keterangan", "Total Responden", "Persentase (%)"]
    t42_data = [
        ["Usia", "Rentang Usia 22 – 30 tahun", "26", "28,30%"],
        ["", "Rentang Usia 31 – 40 tahun", "28", "30,40%"],
        ["", "Rentang Usia 41 – 50 tahun", "27", "29,30%"],
        ["", "Rentang Usia 51 – 55 tahun", "11", "12,00%"],
        ["Jumlah", "", "92", "100"]
    ]
    add_table_data(doc, "Tabel 4.2\nDeskripsi Karakteristik Responden berdasarkan Usia", t42_headers, t42_data, "Sumber : Data primer diolah, (2026)")
    p(doc, "Berdasarkan paparan data pada Tabel 4.2 di atas, terlihat bahwa distribusi usia pegawai yang menjadi sampel dalam penelitian ini cukup merata pada tiga kelompok usia produktif. Responden terbanyak berada pada rentang usia 31 – 40 tahun dengan jumlah 28 orang atau sebesar 30,40%. Posisi selanjutnya diikuti oleh kelompok usia 41 – 50 tahun sebanyak 27 orang (29,30%), dan kelompok usia muda pada rentang 22 – 30 tahun sebanyak 26 orang (28,30%). Sementara itu, kelompok usia senior yang mendekati masa purnatugas (pensiun), yakni rentang usia 51 – 55 tahun, menjadi kelompok minoritas dengan jumlah 11 orang (12,00%).", indent=0.5, space_after=6)

    p(doc, "Karakteristik demografis yang ketiga adalah tingkat pendidikan terakhir:", indent=0.5, space_after=4)
    t43_headers = ["Karakteristik", "Keterangan", "Total Responden", "Persentase (%)"]
    t43_data = [
        ["Pendidikan", "SMA", "8", "8,70"],
        ["", "D3", "12", "13,00"],
        ["", "S1", "65", "70,70"],
        ["", "S2", "7", "7,60"],
        ["Jumlah", "", "92", "100"]
    ]
    add_table_data(doc, "Tabel 4.3\nDeskripsi Karakteristik Responden berdasarkan Pendidikan", t43_headers, t43_data, "Sumber : Data primer diolah, (2026)")
    p(doc, "Berdasarkan Tabel 4.3 di atas, dapat ditarik kesimpulan bahwa tingkat pendidikan aparatur yang menjadi responden dalam penelitian ini didominasi oleh lulusan Sarjana atau Strata Satu (S1), yakni sebanyak 65 orang dengan persentase 70,70%. Pada urutan selanjutnya, terdapat responden D3 sebanyak 12 orang (13,00%), SMA sebanyak 8 orang (8,70%), dan Magister (S2) sebanyak 7 orang (7,60%).", indent=0.5, space_after=8)

    # 2. Statistik Deskriptif
    h3(doc, "2. Analisis Statistik deskriptif")
    p(doc, "Analisis statistik deskriptif bertujuan untuk mengetahui persepsi responden terhadap seluruh variabel penelitian. Rentang skala penilaian dihitung berdasarkan rumus interval: i = (5 - 1) / 3 = 1,33.", indent=0.5, space_after=4)
    t44_headers = ["Kategori", "Nilai rata-rata"]
    t44_data = [
        ["Rendah", "1,00 - 2,33"],
        ["Sedang", "2,34 - 3,67"],
        ["Tinggi", "3,68 - 5,00"]
    ]
    add_table_data(doc, "Tabel 4.4\nKategori Jawaban", t44_headers, t44_data, "Sumber : Sugiyono (2022)")

    t45_headers = ["Variabel", "Nilai rata-rata", "Kategori"]
    t45_data = [
        ["Motivasi Kerja", "3,904", "Tinggi"],
        ["Disiplin Kerja", "3,989", "Tinggi"],
        ["Employee Engagement", "3,925", "Tinggi"],
        ["Kinerja Pegawai", "3,948", "Tinggi"]
    ]
    add_table_data(doc, "Tabel 4.5\nKategori Jawaban Responden", t45_headers, t45_data, "Sumber : Data olahan SmartPLS, (2026)")
    p(doc, "Berdasarkan Tabel 4.5, seluruh variabel penelitian berada pada kategori Tinggi (nilai rata-rata di atas 3,68). Nilai tertinggi diraih oleh Disiplin Kerja (3,989), diikuti oleh Kinerja Pegawai (3,948), Employee Engagement (3,925), dan Motivasi Kerja (3,904).", indent=0.5, space_after=8)

    # 3. Statistik Inferensial (Outer Model)
    h3(doc, "3. Analisis Statistik Inferensial")
    p(doc, "Analisis statistik inferensial menggunakan PLS-SEM dengan software SmartPLS 3.2.9. Pengujian dilakukan melalui evaluasi outer model (validitas dan reliabilitas) dan inner model (pengujian hipotesis).", indent=0.5, space_after=4)
    
    p(doc, "a. Analisis Outer Model", bold=True, space_after=2)
    p(doc, "1) Convergent Validity: Kriteria loading factor harus > 0,70 dan AVE > 0,50.", space_after=4)
    p(doc, "Gambar 4.1\nFull Model Setelah Dikalkulasikan\n(Model Struktural Second-Order SmartPLS 3.2.9)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4, space_after=4)

    # Tabel 4.6 Outer Loadings (48 butir)
    t46_headers = ["Indikator", "Disiplin Kerja", "Employee Engagement", "Kinerja Pegawai", "Motivasi Kerja"]
    t46_data = [
        ["DK01", "0,915", "", "", ""],
        ["DK02", "0,908", "", "", ""],
        ["DK03", "0,942", "", "", ""],
        ["DK04", "0,938", "", "", ""],
        ["DK05", "0,941", "", "", ""],
        ["DK06", "0,915", "", "", ""],
        ["DK07", "0,908", "", "", ""],
        ["DK08", "0,903", "", "", ""],
        ["DK09", "0,939", "", "", ""],
        ["EE01", "", "0,909", "", ""],
        ["EE02", "", "0,925", "", ""],
        ["EE03", "", "0,913", "", ""],
        ["EE04", "", "0,918", "", ""],
        ["EE05", "", "0,906", "", ""],
        ["EE06", "", "0,912", "", ""],
        ["EE07", "", "0,936", "", ""],
        ["EE08", "", "0,917", "", ""],
        ["EE09", "", "0,921", "", ""],
        ["KP01", "", "", "0,923", ""],
        ["KP02", "", "", "0,898", ""],
        ["KP03", "", "", "0,891", ""],
        ["KP04", "", "", "0,919", ""],
        ["KP05", "", "", "0,895", ""],
        ["KP06", "", "", "0,912", ""],
        ["KP07", "", "", "0,913", ""],
        ["KP08", "", "", "0,911", ""],
        ["KP09", "", "", "0,937", ""],
        ["KP10", "", "", "0,938", ""],
        ["KP11", "", "", "0,934", ""],
        ["KP12", "", "", "0,916", ""],
        ["KP13", "", "", "0,894", ""],
        ["KP14", "", "", "0,922", ""],
        ["KP15", "", "", "0,927", ""],
        ["MK01", "", "", "", "0,920"],
        ["MK02", "", "", "", "0,930"],
        ["MK03", "", "", "", "0,941"],
        ["MK04", "", "", "", "0,927"],
        ["MK05", "", "", "", "0,907"],
        ["MK06", "", "", "", "0,926"],
        ["MK07", "", "", "", "0,906"],
        ["MK08", "", "", "", "0,904"],
        ["MK09", "", "", "", "0,924"],
        ["MK10", "", "", "", "0,929"],
        ["MK11", "", "", "", "0,929"],
        ["MK12", "", "", "", "0,938"],
        ["MK13", "", "", "", "0,881"],
        ["MK14", "", "", "", "0,908"],
        ["MK15", "", "", "", "0,912"]
    ]
    add_table_data(doc, "Tabel 4.6\nOuter Loading Pls Algorithm", t46_headers, t46_data, "Sumber : Data Diolah SmartPLS, (2026)")
    p(doc, "Berdasarkan Tabel 4.6, seluruh indikator pembentuk konstruk berada jauh di atas ambang batas 0,70 (berkisar antara 0,881 hingga 0,942), sehingga seluruh indikator dinyatakan valid.", indent=0.5, space_after=6)

    # Tabel 4.7 AVE
    t47_headers = ["Variabel", "Average Variance Extracted (AVE)"]
    t47_data = [
        ["Motivasi Kerja", "0,773"],
        ["Disiplin Kerja", "0,802"],
        ["Employee Engagement", "0,782"],
        ["Kinerja Pegawai", "0,779"]
    ]
    add_table_data(doc, "Tabel 4.7\nAverage Variance Extracted (AVE)", t47_headers, t47_data, "Sumber : Data Diolah SmartPLS, (2026)")

    # 2) Discriminant Validity
    p(doc, "2) Discriminant Validity: Kriteria Fornell-Larcker dan Cross Loading.", space_after=4)
    t48_headers = ["Variabel", "AVE", "√ AVE"]
    t48_data = [
        ["Motivasi Kerja", "0,773", "0,879"],
        ["Disiplin Kerja", "0,802", "0,896"],
        ["Employee Engagement", "0,782", "0,884"],
        ["Kinerja Pegawai", "0,779", "0,883"]
    ]
    add_table_data(doc, "Tabel 4.8\nAkar Kuadrat AVE", t48_headers, t48_data, "Sumber : Data Diolah SmartPLS, (2026)")

    t49_headers = ["Variabel", "Disiplin Kerja", "Employee Engagement", "Kinerja Pegawai", "Motivasi Kerja"]
    t49_data = [
        ["Disiplin Kerja", "0,896", "", "", ""],
        ["Employee Engagement", "0,625", "0,884", "", ""],
        ["Kinerja Pegawai", "0,703", "0,881", "0,883", ""],
        ["Motivasi Kerja", "0,295", "0,701", "0,673", "0,879"]
    ]
    add_table_data(doc, "Tabel 4.9\nValiditas Diskriminan (Fornell-Larcker)", t49_headers, t49_data, "Sumber : Data Diolah SmartPLS, (2026)")

    # Tabel 4.10 Cross Loadings
    t410_headers = ["Indikator", "Disiplin Kerja", "Employee Engagement", "Kinerja Pegawai", "Motivasi Kerja"]
    t410_data = [
        ["DK01", "0,892", "0,585", "0,617", "0,260"],
        ["DK02", "0,879", "0,540", "0,615", "0,280"],
        ["DK03", "0,920", "0,587", "0,645", "0,225"],
        ["DK04", "0,921", "0,622", "0,709", "0,373"],
        ["DK05", "0,901", "0,492", "0,569", "0,196"],
        ["DK06", "0,879", "0,567", "0,625", "0,240"],
        ["DK07", "0,877", "0,576", "0,671", "0,319"],
        ["DK08", "0,863", "0,532", "0,603", "0,249"],
        ["DK09", "0,927", "0,537", "0,611", "0,232"],
        ["EE01", "0,524", "0,888", "0,759", "0,568"],
        ["EE02", "0,576", "0,897", "0,809", "0,598"],
        ["EE03", "0,553", "0,868", "0,780", "0,641"],
        ["EE04", "0,579", "0,880", "0,783", "0,604"],
        ["EE05", "0,595", "0,884", "0,763", "0,588"],
        ["EE06", "0,553", "0,864", "0,776", "0,624"],
        ["EE07", "0,538", "0,899", "0,761", "0,638"],
        ["EE08", "0,500", "0,885", "0,762", "0,621"],
        ["EE09", "0,560", "0,895", "0,821", "0,693"],
        ["KP01", "0,571", "0,769", "0,895", "0,651"],
        ["KP02", "0,652", "0,800", "0,877", "0,589"],
        ["KP03", "0,488", "0,664", "0,834", "0,609"],
        ["KP04", "0,672", "0,763", "0,861", "0,543"],
        ["KP05", "0,641", "0,779", "0,887", "0,572"],
        ["KP06", "0,698", "0,771", "0,865", "0,579"],
        ["KP07", "0,610", "0,764", "0,879", "0,535"],
        ["KP08", "0,613", "0,725", "0,879", "0,562"],
        ["KP09", "0,656", "0,826", "0,922", "0,618"],
        ["KP10", "0,643", "0,809", "0,918", "0,617"],
        ["KP11", "0,613", "0,806", "0,897", "0,615"],
        ["KP12", "0,548", "0,780", "0,871", "0,643"],
        ["KP13", "0,569", "0,772", "0,838", "0,534"],
        ["KP14", "0,679", "0,799", "0,897", "0,601"],
        ["KP15", "0,647", "0,833", "0,917", "0,634"],
        ["MK01", "0,349", "0,659", "0,673", "0,875"],
        ["MK02", "0,240", "0,632", "0,592", "0,895"],
        ["MK03", "0,211", "0,630", "0,599", "0,881"],
        ["MK04", "0,265", "0,605", "0,624", "0,874"],
        ["MK05", "0,156", "0,562", "0,542", "0,836"],
        ["MK06", "0,220", "0,605", "0,573", "0,895"],
        ["MK07", "0,234", "0,568", "0,553", "0,877"],
        ["MK08", "0,297", "0,667", "0,569", "0,860"],
        ["MK09", "0,335", "0,669", "0,635", "0,897"],
        ["MK10", "0,243", "0,624", "0,573", "0,896"],
        ["MK11", "0,258", "0,617", "0,570", "0,898"],
        ["MK12", "0,287", "0,645", "0,631", "0,919"],
        ["MK13", "0,236", "0,611", "0,577", "0,840"],
        ["MK14", "0,314", "0,572", "0,540", "0,833"],
        ["MK15", "0,237", "0,565", "0,609", "0,903"]
    ]
    add_table_data(doc, "Tabel 4.10\nNilai Cross Loading", t410_headers, t410_data, "Sumber : Data Diolah SmartPLS, (2026)")

    # 3) Reliabilitas
    p(doc, "3) Uji Reliabilitas", space_after=4)
    t411_headers = ["Variabel", "Cronbach’s Alpha", "Composite Reliability"]
    t411_data = [
        ["Motivasi Kerja", "0,979", "0,981"],
        ["Disiplin Kerja", "0,969", "0,973"],
        ["Employee Engagement", "0,965", "0,970"],
        ["Kinerja Pegawai", "0,980", "0,981"]
    ]
    add_table_data(doc, "Tabel 4.11\nHasil Uji Reliabilitas", t411_headers, t411_data, "Sumber : Data Diolah SmartPLS, (2026)")

    # Inner Model
    p(doc, "b. Analisis Inner Model", bold=True, space_after=2)
    p(doc, "1) R Square (R²)", space_after=4)
    t412_headers = ["Variabel", "R Square"]
    t412_data = [
        ["Employee Engagement", "0,683"],
        ["Kinerja Pegawai", "0,832"]
    ]
    add_table_data(doc, "Tabel 4.12\nNilai R-Square (R2)", t412_headers, t412_data, "Sumber : Data Diolah SmartPLS, (2026)")
    p(doc, "Nilai R² Employee Engagement sebesar 0,683 menunjukkan pengaruh yang Kuat (68,3% varians dijelaskan oleh Motivasi dan Disiplin). Nilai R² Kinerja Pegawai sebesar 0,832 menunjukkan pengaruh yang Sangat Kuat (83,2% varians dijelaskan oleh Motivasi, Disiplin, dan Employee Engagement).", indent=0.5, space_after=6)

    p(doc, "2) F Square (f²)", space_after=4)
    t413_headers = ["Variabel Eksogen", "Employee Engagement", "Kinerja Pegawai"]
    t413_data = [
        ["Motivasi Kerja", "0,920", "0,104"],
        ["Disiplin Kerja", "0,606", "0,294"],
        ["Employee Engagement", "", "0,599"]
    ]
    add_table_data(doc, "Tabel 4.13\nHasil F2", t413_headers, t413_data, "Sumber : Data Diolah SmartPLS, (2026)")

    p(doc, "3) Uji Hipotesis (Pengaruh Langsung)", space_after=4)
    t414_headers = ["Jalur Pengaruh", "Original Sample (O)", "Sample Mean (M)", "STDEV", "T Statistics", "P Values"]
    t414_data = [
        ["Motivasi Kerja -> Employee Engagement", "0,565", "0,565", "0,062", "9,100", "0,000"],
        ["Disiplin Kerja -> Employee Engagement", "0,459", "0,458", "0,073", "6,318", "0,000"],
        ["Motivasi Kerja -> Kinerja Pegawai", "0,191", "0,190", "0,058", "3,271", "0,001"],
        ["Disiplin Kerja -> Kinerja Pegawai", "0,295", "0,296", "0,065", "4,544", "0,000"],
        ["Employee Engagement -> Kinerja Pegawai", "0,563", "0,565", "0,080", "7,018", "0,000"]
    ]
    add_table_data(doc, "Tabel 4.14\nPengaruh Langsung", t414_headers, t414_data, "Sumber : Data Diolah SmartPLS, (2026)")

    p(doc, "Gambar 4.2\nHasil uji T-Statistik Antar Variabel\n(Seluruh jalur menunjukkan t > 1,96 dan p < 0,05)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4, space_after=4)

    p(doc, "Formulasi Persamaan Struktural:\nSub-Struktur 1: Y = 0,565 X1 + 0,459 X2 + ɛ1\nSub-Struktur 2: Z = 0,191 X1 + 0,295 X2 + 0,563 Y + ɛ2", indent=0.5, space_after=6)

    p(doc, "4) Uji Hipotesis Mediasi (Pengaruh Tidak Langsung)", space_after=4)
    t415_headers = ["Jalur Mediasi", "Original Sample (O)", "Sample Mean (M)", "STDEV", "T Statistics", "P Values"]
    t415_data = [
        ["Motivasi Kerja -> Employee Engagement -> Kinerja Pegawai", "0,318", "0,319", "0,059", "5,423", "0,000"],
        ["Disiplin Kerja -> Employee Engagement -> Kinerja Pegawai", "0,258", "0,260", "0,059", "4,379", "0,000"]
    ]
    add_table_data(doc, "Tabel 4.15\nPengaruh Tidak Langsung", t415_headers, t415_data, "Sumber : Data Diolah SmartPLS, (2026)")
    p(doc, "Hasil uji mediasi mengonfirmasi bahwa Employee Engagement berperan sebagai variabel Mediasi Parsial (Partial Mediation) yang signifikan pada kedua jalur kausalitas tersebut.", indent=0.5, space_after=8)

    # 4.2 Pembahasan Hasil
    h2(doc, "4.2. Pembahasan Hasil")
    
    # H1
    h3(doc, "1. Pengaruh Motivasi Kerja terhadap Employee Engagement")
    p(doc, "Hasil pengujian hipotesis pertama (H1) membuktikan bahwa Motivasi Kerja secara parsial memiliki pengaruh yang positif dan signifikan terhadap Employee Engagement pada Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Sumatera Selatan. Pembuktian ini didasarkan pada perolehan nilai koefisien jalur (Original Sample) sebesar 0,565 yang menunjukkan arah hubungan searah (positif). Lebih lanjut, nilai T-statistic tercatat sebesar 9,100, yang berada jauh di atas ambang batas kritis 1,96, serta didukung oleh perolehan p-value sebesar 0,000 yang lebih kecil dari taraf signifikansi 0,05 (<0,05).", indent=0.5, space_after=4)
    p(doc, "Selain tingkat signifikansi yang tinggi, evaluasi besaran pengaruh atau effect size (f2) pada jalur ini juga menunjukkan angka sebesar 0,920. Merujuk pada kriteria evaluasi model struktural, angka ini menegaskan bahwa Motivasi Kerja memberikan sumbangsih pengaruh yang berukuran \"Sangat Besar\" dalam membentuk keterikatan emosional pegawai. Artinya, semakin baik tingkat pemenuhan motivasi kerja yang dirasakan oleh aparatur, maka akan semakin tinggi dan solid pula tingkat Employee Engagement yang terbangun di dalam instansi tersebut.", indent=0.5, space_after=4)
    p(doc, "Secara empiris, temuan ini memberikan konfirmasi sekaligus jawaban atas fenomena krisis keterikatan yang diuraikan pada latar belakang masalah. Keluhan pegawai terkait tingginya beban kerja di loket perizinan operasional yang belum sebanding dengan apresiasi pimpinan terbukti secara statistik menjadi pemicu utama memudarnya rasa kebanggaan (Dedication) dan ketangguhan mental (Vigor). Hasil penelitian ini menegaskan bahwa pemenuhan indikator motivasi dasar mulai dari kesesuaian hak, ruang aktualisasi diri, hingga pengakuan atas keahlian merupakan katalisator mutlak agar pegawai mau mencurahkan antusiasme dan fokus (Absorption) yang mendalam terhadap tugas pelayanannya.", indent=0.5, space_after=4)
    p(doc, "Secara teoretis, hasil penelitian ini memperkuat pandangan Sutrisno (2022) serta selaras dengan temuan empiris Junengsih et al. (2022) dan Firmansyah & Rochdianingrum (2025), yang membuktikan bahwa pemenuhan motivasi kerja merupakan determinan utama yang positif dan signifikan dalam menumbuhkan keterikatan emosional dan dedikasi pegawai di lingkungan kerja.", indent=0.5, space_after=6)

    # H2
    h3(doc, "2. Pengaruh Disiplin Kerja terhadap Employee Engagement")
    p(doc, "Pembuktian hipotesis kedua (H2) mengonfirmasi bahwa Disiplin Kerja memiliki pengaruh yang positif dan signifikan terhadap pembentukan Employee Engagement di lingkungan DPMPTSP Provinsi Sumatera Selatan. Kesimpulan ini ditarik berdasarkan output uji Bootstrapping yang menghasilkan nilai koefisien jalur (Original Sample) sebesar 0,459, dengan nilai T-statistic sebesar 6,318 (> 1,96) dan p-value 0,000 (< 0,05).", indent=0.5, space_after=4)
    p(doc, "Analisis effect size (f2) mencatatkan angka sebesar 0,606 yang masuk dalam klasifikasi pengaruh \"Besar\". Hal ini mengartikan bahwa konsistensi aparatur dalam mematuhi aturan jam kerja dan SOP berkontribusi sangat kuat dalam memperkokoh fokus mental serta dedikasi pegawai terhadap instansi.", indent=0.5, space_after=4)
    p(doc, "Tingginya angka pelanggaran jam kerja pada tahun 2025 (2.520 kasus keterlambatan dan 375 hari alpa) terbukti secara nyata berkorelasi dengan rapuhnya keterikatan kerja. Penegakan disiplin preventif yang transparan terbukti menjadi pendorong keteraturan ekosistem birokrasi yang memicu stabilitas psikologis pegawai.", indent=0.5, space_after=4)
    p(doc, "Temuan ini berpijak pada teori Mangkunegara (2022) dan mendukung temuan penelitian Irmawati & Rudini (2024) serta Alamsyah et al. (2025) yang merumuskan bahwa kedisiplinan kerja yang konsisten memperkuat partisipasi aktif dan keterlibatan emosional pegawai.", indent=0.5, space_after=6)

    # H3
    h3(doc, "3. Pengaruh Motivasi Kerja terhadap Kinerja Pegawai")
    p(doc, "Evaluasi pengujian hipotesis ketiga (H3) mendemonstrasikan bahwa Motivasi Kerja secara parsial memiliki pengaruh langsung yang positif dan signifikan terhadap Kinerja Pegawai (Original Sample 0,191, T-statistic 3,271, p-value 0,001).", indent=0.5, space_after=4)
    p(doc, "Effect size (f2) sebesar 0,104 masuk kategori \"Kecil\", yang mengisyaratkan bahwa meskipun dorongan motivasi mampu mendongkrak kinerja secara langsung, pengaruhnya membutuhkan variabel mediasi agar dampaknya menjadi optimal dan berlipat ganda.", indent=0.5, space_after=4)
    p(doc, "Pemenuhan motivasi kerja terbukti mampu memulihkan kesenjangan IKU perizinan yang sempat tertinggal -16,6% melalui peningkatan semangat kerja harian pegawai di loket perizinan.", indent=0.5, space_after=4)
    p(doc, "Temuan ini sekaligus memecahkan Research Gap dengan membantah temuan Suwanto (2020) serta Hidayat dan Tjahjono (2022), serta sejalan dengan penelitian Noviantoro et al. (2022) dan Adhelina et al. (2024) bahwa motivasi berpengaruh positif dan signifikan terhadap kinerja pegawai.", indent=0.5, space_after=6)

    # H4
    h3(doc, "4. Pengaruh Disiplin Kerja terhadap Kinerja Pegawai")
    p(doc, "Penelusuran hasil hipotesis keempat (H4) membuktikan bahwa Disiplin Kerja berpengaruh positif dan signifikan terhadap Kinerja Pegawai (Original Sample 0,295, T-statistic 4,544, p-value 0,000, effect size f2 = 0,294 kategori \"Sedang\").", indent=0.5, space_after=4)
    p(doc, "Kepatuhan terhadap jam operasional dan ketepatan waktu pelaporan terbukti meminimalisasi inefisiensi birokrasi dan secara langsung meningkatkan produktivitas pelayanan perizinan.", indent=0.5, space_after=4)
    p(doc, "Temuan ini membantah hasil riset Simatupang & Safitri (2023) dan memperkuat teori Rivai & Sagala (2022) serta temuan Adhelina et al. (2024), Noviantoro et al. (2022), dan Oktarisa et al. (2026).", indent=0.5, space_after=6)

    # H5
    h3(doc, "5. Pengaruh Employee Engagement terhadap Kinerja Pegawai")
    p(doc, "Pengujian hipotesis kelima (H5) membuktikan bahwa Employee Engagement berpengaruh positif dan signifikan terhadap Kinerja Pegawai dengan koefisien Original Sample yang sangat dominan sebesar 0,563, T-statistic 7,018, p-value 0,000, serta effect size f2 sebesar 0,599 (kategori \"Besar\").", indent=0.5, space_after=4)
    p(doc, "Employee Engagement terbukti menjadi prediktor langsung paling kuat bagi pencapaian kinerja, mengungguli pengaruh langsung motivasi dan disiplin kerja.", indent=0.5, space_after=4)
    p(doc, "Aparatur yang memiliki energi (Vigor), kebanggaan (Dedication), dan konsentrasi (Absorption) terdorong mengerahkan Discretionary Effort untuk menuntaskan target izin tepat waktu.", indent=0.5, space_after=4)
    p(doc, "Hasil ini memperkuat postulat Afandi (2023) dan selaras dengan temuan Pratama et al. (2025) serta Azahra & Nurhasan (2025).", indent=0.5, space_after=6)

    # H6
    h3(doc, "6. Pengaruh Motivasi Kerja terhadap Kinerja Pegawai Melalui Employee Engagement")
    p(doc, "Pengujian hipotesis keenam (H6) membuktikan bahwa Employee Engagement secara signifikan memediasi pengaruh Motivasi Kerja terhadap Kinerja Pegawai (koefisien tidak langsung 0,318, T-statistic 5,423, p-value 0,000).", indent=0.5, space_after=4)
    p(doc, "Sifat mediasi adalah Mediasi Parsial (Partial Mediation). Artinya dorongan motivasi kerja akan menghasilkan peningkatan kinerja yang berlipat ganda apabila terlebih dahulu berhasil menumbuhkan keterikatan psikologis dan rasa kepemilikan pegawai.", indent=0.5, space_after=4)
    p(doc, "Hasil ini menolak anomali riset Dewanata & Fadila (2025) dan memperkuat temuan Prabowo & Kui (2025) serta Indah et al. (2022).", indent=0.5, space_after=6)

    # H7
    h3(doc, "7. Pengaruh Disiplin Kerja terhadap Kinerja Pegawai Melalui Employee Engagement")
    p(doc, "Pengujian hipotesis ketujuh (H7) membuktikan bahwa Disiplin Kerja berpengaruh positif dan signifikan terhadap Kinerja Pegawai melalui Employee Engagement sebagai mediator (koefisien tidak langsung 0,258, T-statistic 4,379, p-value 0,000, Mediasi Parsial).", indent=0.5, space_after=4)
    p(doc, "Temuan ini membuktikan bahwa penegakan aturan jam kerja tidak boleh melahirkan \"kepatuhan semu\" (Nuraini, 2023), melainkan harus ditransformasikan menjadi dedikasi dan keterlibatan emosional agar bermuara pada kinerja pelayanan publik yang prima.", indent=0.5, space_after=4)
    p(doc, "Hasil ini menolak anomali temuan Alamsyah et al. (2025) dan Ridwan et al. (2025), serta mengukuhkan temuan Purwaningtyas et al. (2026) dan Iskandar et al. (2026).", indent=0.5, space_after=8)
    rapikan_tabel_dokumen(doc)
    doc.add_page_break()

print("Bab 3 & Bab 4 siap.")
