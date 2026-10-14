# Test code snippet for updated segments in engine_hierarchical_pipeline
import sys
sys.path.insert(0, 'word-addin-tesismagister/backend')
import academic_connector

def build_narrative_segments_macro(sektor, peran_strategis, urgensi_kinerja, instansi, y_title, y_lower, acad, seed_val):
    mode = seed_val % 4
    if mode == 0:
        seg1 = (
            f"Penyelenggaraan tata kelola sektor publik modern menuntut transformasi kelembagaan yang berorientasi "
            f"pada penciptaan nilai publik (*public value creation*) dan akuntabilitas kinerja yang berintegritas tinggi {acad['cite_macro']}. "
            f"Dalam diskursus reformasi birokrasi kontemporer, keberhasilan instansi dalam merealisasikan mandat kelembagaannya tidak lagi "
            f"semata-mata diukur dari serapan anggaran atau kepatuhan administratif mekanis, melainkan dari sejauh mana kapasitas dan integritas "
            f"Sumber Daya Manusia (SDM) mampu merespons ekspektasi pemangku kepentingan secara transparan dan berorientasi mutu {acad['cite_sdm']}. "
            f"Fondasi utama keunggulan institusi publik berpijak pada kematangan moral, motivasi pelayanan publik (*public service motivation*), "
            f"serta dedikasi aparatur dalam mengeksekusi tanggung jawab kedinasannya."
        )
        seg2 = (
            f"Sebagai representasi institusi strategis di daerah, {instansi} memikul amanah fundamental dalam {peran_strategis}. "
            f"Dalam kerangka tata kelola tersebut, efektivitas penuntasan target sasaran kelembagaan bermuara langsung pada optimalisasi {y_lower} "
            f"aparatur di lingkungan {instansi} {acad['cite_kinerja']}. Capaian {y_lower} menjadi indikator kinerja krusial karena mencerminkan derajat "
            f"keberhasilan institusi dalam menjamin bahwa {urgensi_kinerja} dapat terwujud secara konsisten, terukur, dan akuntabel. "
            f"Namun demikian, penelusuran fakta empiris di lokus penelitian memperlihatkan indikasi dinamika defisit kinerja yang cukup mengkhawatirkan "
            f"selama tiga tahun terakhir."
        )
    elif mode == 1:
        seg1 = (
            f"Dalam paradigma manajemen strategis berbasis sumber daya (*Resource-Based View*), keunggulan kompetitif dan kapabilitas dinamis "
            f"suatu organisasi bertumpu fundamental pada kualitas modal manusia (*human capital*) sebagai aset nirwujud (*intangible asset*) "
            f"yang bernilai, langka, dan sulit ditiru {acad['cite_macro']}. Pada tataran organisasi publik, pembaruan kelembagaan menempatkan aparatur "
            f"bukan sekadar pelaksana administratif rutin, melainkan aktor intelektual kunci yang menggerakkan inovasi dan ketepatan eksekusi "
            f"kebijakan di bidang {sektor} {acad['cite_sdm']}. Keberhasilan reformasi birokrasi daerah sangat ditentukan oleh sejauh mana kompetensi, "
            f"keterikatan psikologis, dan orientasi berprestasi aparatur terkelola secara berkelanjutan."
        )
        seg2 = (
            f"Dalam konstelasi pembangunan daerah, {instansi} memiliki posisi strategis sebagai motor penggerak dalam {peran_strategis}. "
            f"Realisasi mandat fungsional tersebut mensyaratkan ketersediaan aparatur dengan tingkat {y_lower} yang prima dan adaptif {acad['cite_kinerja']}. "
            f"Tingkat {y_lower} ini menjadi parameter penentu karena merefleksikan seberapa efektif {urgensi_kinerja} dapat dicapai di tengah kompleksitas "
            f"regulasi dan tingginya ekspektasi publik. Kendati demikian, telaah data internal menunjukkan bahwa capaian kinerja pada unit ini belum menunjukkan "
            f"performa optimal, bahkan terindikasi mengalami penurunan target yang signifikan."
        )
    elif mode == 2:
        seg1 = (
            f"Tuntutan masyarakat terhadap efisiensi pelayanan publik dan akuntabilitas pemerintah daerah semakin mendesak terciptanya birokrasi "
            f"yang lincah (*agile bureaucracy*) dan adaptif terhadap dinamika lingkungan strategis di bidang {sektor} {acad['cite_macro']}. "
            f"Kebijakan reformasi birokrasi tematik mengamanatkan pergeseran paradigma dari model birokrasi prosedural-mekanis menuju tata kelola "
            f"berbasis hasil (*outcome-based governance*) yang tanggap terhadap perubahan {acad['cite_sdm']}. Hambatan mendasar dalam akselerasi perubahan "
            f"ini kerap kali bukan disebabkan oleh keterbatasan regulasi formal, melainkan kesiapan mentalitas, etos kerja, serta resiliensi psikologis "
            f"aparatur dalam menginternalisasi tuntutan kerja yang terus berkembang."
        )
        seg2 = (
            f"Sebagai unit kelembagaan yang memegang otoritas penting di wilayah kerjanya, {instansi} bertanggung jawab langsung dalam {peran_strategis}. "
            f"Pada konteks ini, keberhasilan pelaksanaan tugas operasional bertumpu penuh pada {y_lower} aparatur yang konsisten dan berintegritas {acad['cite_kinerja']}. "
            f"Kualitas {y_lower} aparatur menjadi tolok ukur utama bagi tercapainya {urgensi_kinerja} secara efektif dan berkesinambungan. Namun demikian, "
            f"kenyataan empiris yang teramati di lapangan justru memperlihatkan fenomena kesenjangan yang nyata antara target yang ditetapkan dengan realisasi capaian kinerja pegawai."
        )
    else:
        seg1 = (
            f"Akselerasi transformasi digital dan tata kelola modern menuntut terciptanya keunggulan kinerja pelayanan yang presisi, responsif, "
            f"dan berdaya saing tinggi pada setiap institusi yang membidangi {sektor} {acad['cite_macro']}. Penguatan sistem kelembagaan kontemporer "
            f"mengharuskan aparatur untuk bertransformasi dari pola kerja konvensional menuju pola kerja yang kolaboratif, transparan, dan akuntabel {acad['cite_sdm']}. "
            f"Tantangan utama dalam akselerasi kinerja kelembagaan saat ini bertumpu pada kesiapan perilaku aparatur, iklim profesionalisme, "
            f"serta kapasitas adaptasi sumber daya manusia dalam menerjemahkan target strategis menjadi output pelayanan publik yang prima."
        )
        seg2 = (
            f"Dalam menopang efektivitas tata kelola pemerintahan daerah, {instansi} mengemban peranan vital dalam {peran_strategis}. "
            f"Guna mewujudkan tata kelola yang unggul, konsistensi {y_lower} aparatur menjadi prasyarat mutlak yang tidak dapat ditawar {acad['cite_kinerja']}. "
            f"Capaian {y_lower} pegawai mencerminkan sejauh mana {urgensi_kinerja} dapat dihadirkan secara profesional demi mendukung pencapaian sasaran strategis daerah. "
            f"Akan tetapi, data capaian kinerja tiga tahun terakhir pada instansi ini mengindikasikan adanya disrupsi dan penurunan efektivitas operasional yang memerlukan telaah mendalam."
        )
    return seg1, seg2

print('Macro builder loaded successfully!')
