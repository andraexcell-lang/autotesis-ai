# backend/s1_constructs.py
"""
S1 CONSTRUCT KNOWLEDGE BASE (N VARIABEL BEBAS)
===============================================
Perpustakaan konstruk untuk skripsi S1 Manajemen/MSDM yang mendukung model
dengan 2 s.d. 5 variabel bebas (X1 ... X5) tanpa duplikasi naskah manual.

Setiap konstruk menyediakan:
1. definisi (2 paragraf ilmiah beserta sitasi),
2. sintesis konsep,
3. 3 dimensi x 2 butir pernyataan kuesioner (skala Likert 1-5),
4. narasi fenomena lapangan (latar belakang masalah),
5. narasi pengaruh terhadap kinerja (kerangka berpikir & pembahasan),
6. daftar sitasi teoretis dan empiris.

CATATAN INTEGRITAS AKADEMIK:
Seluruh sitasi yang dipakai di sini WAJIB sudah terdaftar di
S1_CORE_BIBLIOGRAPHY (s1_hierarchical_engine.py) agar tidak terjadi
"ghost reference" dan tidak ada DOI yang dikarang. Bila nama konstruk tidak
dikenali, modul ini memakai profil generik berbasis teori MSDM umum
(Armstrong & Taylor, 2023; Mangkunegara, 2022) tanpa mengarang sumber baru.
"""

import os
import json
from typing import Any, Dict, List, Optional

# Basis Memori Dinamis untuk Profil Konstruk Hasil Riset Ontologi S1
DYNAMIC_S1_LIBRARY: Dict[str, Dict[str, Any]] = {}
CACHE_FILE_S1_ONTOLOGY = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output", "s1_ontology_cache.json")


def register_dynamic_construct(name: str, profile: Dict[str, Any]) -> None:
    """Mendaftarkan profil konstruk hasil riset ontologi secara dinamis."""
    if not name or not profile:
        return
    k_norm = normalize_construct_name(name)
    k_raw = name.strip().lower()
    if k_norm:
        DYNAMIC_S1_LIBRARY[k_norm] = profile
    DYNAMIC_S1_LIBRARY[k_raw] = profile


def clear_dynamic_constructs() -> None:
    """Membersihkan memori konstruk dinamis."""
    DYNAMIC_S1_LIBRARY.clear()


def _load_cached_dynamic_ontology() -> None:
    """Memuat cache ontologi S1 dari disk jika memori dinamis kosong."""
    if DYNAMIC_S1_LIBRARY:
        return
    if os.path.exists(CACHE_FILE_S1_ONTOLOGY):
        try:
            with open(CACHE_FILE_S1_ONTOLOGY, "r", encoding="utf-8") as f:
                cdata = json.load(f)
            vars_data = cdata.get("variables", {})
            for vk, vd in vars_data.items():
                vname = vd.get("input_name") or vd.get("nama_resmi") or ""
                if vname:
                    register_dynamic_construct(vname, vd)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 1. Perpustakaan Konstruk
# ---------------------------------------------------------------------------

CONSTRUCT_LIBRARY: Dict[str, Dict[str, Any]] = {
    "kepemimpinan transformasional": {
        "name": "Kepemimpinan Transformasional",
        "cites": ["Yukl & Gardner (2022)", "Raharja et al. (2022)", "Ausat et al. (2022)", "Robbins & Judge (2022)"],
        "theory_cite": "Yukl & Gardner (2022)",
        "concept": "Kemampuan pemimpin menginspirasi dan menggerakkan bawahan melampaui harapan standar",
        "definition": [
            ("Konsep kepemimpinan transformasional berakar dari teori kepemimpinan yang dikembangkan oleh Yukl & Gardner (2022) serta Raharja et al. (2022). Menurut Yukl & Gardner (2022), kepemimpinan transformasional adalah gaya kepemimpinan di mana pemimpin mampu menginspirasi dan memotivasi pengikutnya untuk mencapai hasil kerja yang melampaui harapan standar mereka. Pemimpin transformasional memberikan visi yang menarik, menjadi teladan moral, menstimulasi daya pikir kritis bawahan, serta memberikan perhatian personal terhadap kebutuhan pengembangan karir bawahan.", "Yukl & Gardner (2022)"),
            ("Sementara itu, Ausat et al. (2022) mengemukakan bahwa kepemimpinan transformasional berfokus pada proses membangun komitmen terhadap tujuan organisasi dan memberdayakan pengikut untuk menyelesaikan sasaran tersebut secara efektif. Robbins & Judge (2022) menambahkan bahwa pemimpin yang memiliki karakteristik transformasional berhasil mengubah kesadaran pengikut dari sekadar pemenuhan kepentingan jangka pendek menjadi dedikasi penuh terhadap kemajuan institusi.", "Ausat et al. (2022)"),
        ],
        "synthesis": "pola kepemimpinan visioner yang membangkitkan motivasi intrinsik bawahan melalui keteladanan integritas, komunikasi visi yang inspiratif, stimulasi ide-ide baru, serta perhatian individual demi terwujudnya performa organisasi yang unggul",
        "dimensi_cite": "Yukl & Gardner (2022) dan Robbins & Judge (2022)",
        "dimensions": [
            {"dim": "Pengaruh Ideal", "desc": "Pengaruh Ideal (Idealized Influence): Pimpinan menjadi panutan yang dipercaya dan dihormati oleh bawahan, menunjukkan integritas moral yang kokoh, serta mendahulukan kepentingan institusi di atas kepentingan pribadi.",
             "items": ["Pimpinan menjadi teladan moral dan integritas dalam pelaksanaan tugas kedinasan.",
                       "Pimpinan mendahulukan kepentingan bersama di atas kepentingan pribadi organisasi."]},
            {"dim": "Motivasi Inspirasional", "desc": "Motivasi Inspirasional (Inspirational Motivation): Kemampuan pimpinan dalam mengartikulasikan visi masa depan yang jelas dan optimis, menumbuhkan antusiasme kerja tim, serta memberikan dorongan semangat dalam menghadapi tantangan tugas kedinasan.",
             "items": ["Pimpinan mengomunikasikan visi dan target kerja instansi secara jelas dan meyakinkan.",
                       "Pimpinan membangkitkan optimisme dan antusiasme kerja seluruh pegawai."]},
            {"dim": "Stimulasi Intelektual", "desc": "Stimulasi Intelektual (Intellectual Stimulation): Tindakan pimpinan yang mendorong bawahan untuk bersikap inovatif, berpikir kritis, mempertanyakan asumsi lama, dan mencari alternatif solusi kreatif atas kendala kerja.",
             "items": ["Pimpinan mendorong pegawai mencari solusi inovatif dalam memecahkan kendala kerja.",
                       "Pimpinan terbuka terhadap ide-ide perbaikan metode kerja yang diajukan staf."]},
            {"dim": "Pertimbangan Individual", "desc": "Pertimbangan Individual (Individualized Consideration): Sikap kepedulian pimpinan terhadap kebutuhan perkembangan masing-masing bawahan melalui pembimbingan (mentoring), pelatihan, dan penciptaan lingkungan kerja yang suportif.",
             "items": ["Pimpinan memberikan bimbingan dan perhatian terhadap kebutuhan pengembangan karir pegawai.",
                       "Pimpinan menciptakan suasana kerja yang suportif bagi setiap anggota tim."]},
        ],
        "issue": "gaya komunikasi kepemimpinan yang diterapkan masih dirasakan belum merata dalam memberikan teladan serta ruang stimulasi inovasi bagi pegawai pelaksana",
        "effect": "Keteladanan, stimulasi intelektual, dan motivasi inspirasional dari pimpinan akan menumbuhkan komitmen kerja bawahan untuk bekerja secara profesional dan tepat waktu",
        "effect_long": "keteladanan moral, stimulasi ide-ide baru, serta perhatian individual dari pimpinan terbukti secara nyata mampu mendongkrak motivasi intrinsik dan kinerja pegawai instansi publik",
        "empirical": "Nasrip et al. (2023) dan Ausat et al. (2022)",
        "likert_focus": "penerapan kepemimpinan transformasional oleh pimpinan",
    },
    "budaya organisasi": {
        "name": "Budaya Organisasi",
        "cites": ["Darda, Rojikun & Yusuf (2022)", "Harrison & Stokes (dalam Winanti, 2013)", "Robbins (dalam Livia, 2017)", "Robbins & Judge (2022)"],
        "theory_cite": "Harrison & Stokes (dalam Winanti, 2013)",
        "concept": "Sistem nilai dan norma bersama sebagai pedoman perilaku aparatur",
        "definition": [
            ("Budaya organisasi memegang peranan krusial dalam membentuk identitas serta perilaku para anggota dalam suatu organisasi. Menurut Darda, Rojikun & Yusuf (2022), budaya organisasi adalah seperangkat asumsi atau sistem keyakinan, nilai, dan norma yang dikembangkan dalam organisasi dan dijadikan pedoman tingkah laku bagi anggotanya untuk mengatasi masalah adaptasi eksternal dan integrasi internal. Harrison dan Stokes dalam Winanti (2013) membedakan budaya organisasi menjadi empat dimensi orientasi budaya yang dipersepsikan oleh para anggotanya, yaitu orientasi pada kekuasaan, orientasi pada peran, orientasi pada prestasi, dan orientasi pada dukungan.", "Darda, Rojikun & Yusuf (2022)"),
            ("Robbins dalam Livia (2017) menegaskan bahwa budaya organisasi memiliki tujuh karakteristik primer yang merupakan hakikat budaya organisasi, yaitu inovasi dan pengambilan risiko, perhatian pada rincian, orientasi hasil, orientasi orang, orientasi tim, keagresifan, dan kemantapan. Budaya organisasi yang kuat berfungsi sebagai perekat sosial dan kompas moral yang membimbing seluruh aparatur dalam memberikan pelayanan terbaik kepada masyarakat.", "Robbins (dalam Livia, 2017)"),
        ],
        "synthesis": "seperangkat nilai filosofis, norma kepatuhan, dan kebiasaan kerja bersama yang diinternalisasikan oleh setiap aparatur pegawai guna menyelaraskan tindakan menuju ketercapaian target instansi",
        "dimensi_cite": "Harrison & Stokes (dalam Winanti, 2013) dan Robbins (dalam Livia, 2017)",
        "dimensions": [
            {"dim": "Orientasi pada Kekuasaan (Power)", "desc": "The Power Orientation (Orientasi pada Kekuasaan): Budaya yang didasarkan atas perbedaan akses kepada sumber daya, di mana pihak yang berkuasa menggunakan sumber-sumber tersebut untuk mengendalikan perilaku orang lain; anggota termotivasi oleh imbalan dan hukuman serta keinginan bekerja dengan pimpinan yang kuat.",
             "items": ["Pimpinan instansi menggunakan kewenangannya untuk mengendalikan jalannya pelaksanaan tugas.",
                       "Kebijakan dan keputusan pimpinan menjadi pedoman utama pegawai dalam bekerja.",
                       "Pegawai termotivasi untuk patuh karena adanya imbalan dan sanksi yang jelas dari pimpinan."]},
            {"dim": "Orientasi pada Peran (Role)", "desc": "The Role Orientation (Orientasi pada Peran): Budaya yang menempatkan sistem struktur dan prosedur untuk menggantikan kekuasaan pimpinan; kewajiban dan imbalan bagi anggota didefinisikan dengan teliti, biasanya secara tertulis, dengan kontrak eksplisit maupun implisit antara organisasi dan individu.",
             "items": ["Struktur organisasi dan prosedur kerja melindungi kepastian peran setiap pegawai.",
                       "Tugas, kewajiban, dan hak pegawai didefinisikan secara jelas dan tertulis.",
                       "Pegawai menjalankan fungsi khusus sesuai dengan peraturan dan metode kerja yang berlaku."]},
            {"dim": "Orientasi pada Prestasi (Achievement)", "desc": "The Achievement Orientation (Orientasi pada Prestasi): Budaya yang menggunakan misi organisasi untuk menarik dan menggerakkan energi anggota mengejar tujuan bersama; anggota memberikan kontribusi secara bebas sebagai respons atas komitmen mereka kepada tujuan yang dibagi.",
             "items": ["Misi instansi menjadi pendorong bagi pegawai untuk mengejar tujuan bersama.",
                       "Pegawai memberikan kontribusi terbaik secara sukarela bagi pencapaian target instansi.",
                       "Pegawai rela memberikan lebih banyak usaha demi kemajuan instansi."]},
            {"dim": "Orientasi pada Dukungan (Support)", "desc": "The Support Orientation (Orientasi pada Dukungan): Iklim organisasi yang didasarkan pada saling percaya di antara anggota; pegawai yakin dinilai sebagai manusia, bukan hanya sebagai mesin, dengan kehangatan yang mendorong semangat kerja.",
             "items": ["Saling percaya terjalin antara pegawai dan pimpinan dalam lingkungan kerja.",
                       "Pegawai merasa dinilai sebagai manusia, bukan sekadar alat kerja.",
                       "Kehangatan hubungan kerja mendorong semangat pegawai dalam bertugas."]},
        ],
        "issue": "pembiasaan budaya kerja kolaboratif, orientasi mutu pelayanan, dan keterbukaan komunikasi antarunit kerja masih terbentur pada sekat-sekat birokrasi, sehingga sinergi tim belum solid",
        "effect": "Nilai budaya kerja yang menjunjung tinggi kedisiplinan dan orientasi hasil akan memandu pegawai untuk mencapai target kerja dengan penuh rasa tanggung jawab",
        "effect_long": "internalisasi nilai-nilai kerja seperti orientasi mutu, akuntabilitas, dan keteraturan prosedur kerja terbukti menjadi pedoman perilaku yang efektif dalam meningkatkan mutu output pelayanan kedinasan",
        "empirical": "Darda, Rojikun & Yusuf (2022)",
        "likert_focus": "budaya kerja yang berlaku di lingkungan instansi",
    },
    "motivasi kerja": {
        "name": "Motivasi Kerja",
        "cites": ["Farhan & Indriyaningrum (2023)", "Maslow (1954)", "Hasibuan (2022)", "Mangkunegara (2022)"],
        "theory_cite": "Maslow (1954)",
        "concept": "Dorongan internal dan eksternal yang menggerakkan pegawai bekerja optimal",
        "definition": [
            ("Motivasi kerja merupakan kekuatan pendorong yang menggerakkan seseorang untuk bekerja secara optimal. Teori motivasi yang dikemukakan oleh Abraham H. Maslow dalam bukunya \"Motivation and Personality\" banyak dikenal sebagai teori motivasi Maslow (dalam Farhan & Indriyaningrum, 2023). Menurut Maslow, kebutuhan-kebutuhan manusia digolongkan sebagai kebutuhan primer, sedangkan yang lainnya dikenal pula sebagai kebutuhan sekunder; hierarki kebutuhan inilah yang mendorong seseorang untuk termotivasi bekerja. Kebutuhan mengaktualisasi diri memiliki indikator kesempatan dalam menyampaikan gagasan/ide dan mendapat pelatihan dari perusahaan.", "Maslow (1954)"),
            ("Farhan & Indriyaningrum (2023) menjelaskan bahwa motivasi kerja dipengaruhi oleh faktor intern dan faktor ekstern. Faktor intern mencakup keinginan untuk hidup, keinginan untuk memiliki, keinginan memperoleh penghargaan, dan keinginan memperoleh pengakuan; sedangkan faktor ekstern mencakup kondisi lingkungan kerja, kompensasi yang memadai, dan atasan yang baik. Pemenuhan kebutuhan yang memberi kepuasan ini mengarahkan perilaku pegawai untuk mencapai tujuan organisasi.", "Farhan & Indriyaningrum (2023)"),
        ],
        "synthesis": "dorongan internal maupun eksternal yang menumbuhkan kegairahan, kesungguhan, dan kesediaan pegawai untuk mengerahkan seluruh kemampuan terbaiknya demi tercapainya sasaran kerja institusi",
        "dimensi_cite": "Maslow (1954) dan Farhan & Indriyaningrum (2023)",
        "dimensions": [
            {"dim": "Kebutuhan Fisiologis", "desc": "Kebutuhan Fisiologis: Kebutuhan paling dasar manusia yang meliputi pemenuhan kebutuhan hidup pokok melalui penghasilan, fasilitas kerja dasar, dan kondisi kerja yang menunjang pemenuhan kebutuhan fisik pegawai.",
             "items": ["Gaji dan penghasilan yang diterima mampu memenuhi kebutuhan hidup pokok pegawai.",
                       "Fasilitas kerja dasar tersedia dan layak untuk menunjang pelaksanaan tugas.",
                       "Kondisi kerja memungkinkan pegawai memenuhi kebutuhan hidup harian keluarga."]},
            {"dim": "Kebutuhan Rasa Aman", "desc": "Kebutuhan Rasa Aman: Kebutuhan akan jaminan keamanan dan keselamatan kerja, kejelasan status kepegawaian, serta perlindungan atas risiko pekerjaan sehingga pegawai merasa aman bekerja.",
             "items": ["Jaminan keamanan dan keselamatan kerja terpenuhi oleh instansi.",
                       "Kejelasan status kepegawaian memberikan rasa aman dalam bekerja.",
                       "Perlindungan atas risiko pekerjaan diberikan oleh pimpinan."]},
            {"dim": "Kebutuhan Sosial", "desc": "Kebutuhan Sosial: Kebutuhan akan hubungan harmonis dengan rekan kerja, penerimaan sebagai bagian dari kelompok kerja, serta suasana kekeluargaan yang mendukung kelancaran tugas.",
             "items": ["Hubungan harmonis terjalin antara pegawai dengan rekan kerja.",
                       "Pegawai diterima dan dihargai sebagai bagian dari kelompok kerja.",
                       "Suasana kekeluargaan di lingkungan kerja mendukung kelancaran tugas."]},
            {"dim": "Kebutuhan Penghargaan", "desc": "Kebutuhan Penghargaan: Kebutuhan akan pengakuan atas prestasi kerja, kesempatan promosi berdasarkan capaian kinerja, serta apresiasi pimpinan terhadap hasil kerja bawahan.",
             "items": ["Pengakuan dan penghargaan diberikan atas prestasi kerja pegawai.",
                       "Kesempatan promosi diberikan berdasarkan capaian kinerja.",
                       "Pimpinan memberikan apresiasi terhadap hasil kerja bawahan."]},
            {"dim": "Kebutuhan Aktualisasi Diri", "desc": "Kebutuhan Aktualisasi Diri: Kebutuhan tertinggi berupa kesempatan dalam menyampaikan gagasan/ide, mendapat pelatihan dari instansi, dan mengembangkan potensi terbaik dalam bekerja.",
             "items": ["Kesempatan dalam menyampaikan gagasan/ide kepada instansi.",
                       "Mendapat pelatihan dari instansi untuk pengembangan diri.",
                       "Pegawai diberi ruang mengembangkan potensi terbaiknya dalam bekerja."]},
        ],
        "issue": "semangat kerja pegawai masih berfluktuasi, terlihat dari masih adanya pegawai yang bekerja sekadar memenuhi kewajiban formal tanpa dorongan untuk berprestasi lebih tinggi",
        "effect": "Motivasi kerja yang tinggi akan mendorong pegawai mencurahkan kemampuan terbaiknya, sehingga kualitas dan kuantitas hasil kerja meningkat secara nyata",
        "effect_long": "pemenuhan kebutuhan berjenjang mulai dari kebutuhan fisiologis, rasa aman, sosial, penghargaan, hingga aktualisasi diri terbukti membangkitkan kegairahan kerja yang berujung pada peningkatan kinerja pegawai",
        "empirical": "Farhan & Indriyaningrum (2023)",
        "likert_focus": "motivasi kerja yang dirasakan pegawai",
    },
    "disiplin kerja": {
        "name": "Disiplin Kerja",
        "cites": ["Prasetio, Ruminta & Akbar (2025)", "Moenir (dalam Winarto & Sayoto, 2018)", "Hasibuan (2022)", "Mangkunegara (2022)"],
        "theory_cite": "Moenir (dalam Winarto & Sayoto, 2018)",
        "concept": "Kesadaran dan kesediaan menaati peraturan serta norma kerja yang berlaku",
        "definition": [
            ("Disiplin kerja mencerminkan tingkat kesadaran dan kesediaan pegawai dalam mematuhi peraturan serta norma kerja yang ditetapkan organisasi. Untuk mengukur disiplin kerja, Moenir dalam (Winarto & Sayoto, 2018) mengemukakan dimensi dan indikator disiplin kerja yang terdiri atas dimensi ketepatan waktu, dimensi tanggung jawab kerja, dimensi ketaatan, dan dimensi inisiatif. Dimensi ketepatan waktu terdiri dari indikator ketepatan waktu dalam melaksanakan tugas dan penghematan waktu dalam melaksanakan tugas.", "Moenir (dalam Winarto & Sayoto, 2018)"),
            ("Prasetio, Ruminta & Akbar (2025) menegaskan bahwa kompensasi dan disiplin kerja merupakan aspek penting dalam meningkatkan kinerja karyawan serta menumbuhkan semangat kerja. Dimensi tanggung jawab kerja meliputi melakukan pekerjaan sesuai target, evaluasi kerja, dan keberanian mengambil risiko; dimensi ketaatan meliputi ketaatan terhadap jam kerja, pimpinan, serta prosedur dan metode kerja; dimensi inisiatif meliputi kemampuan menyelesaikan masalah tanpa menunggu instruksi dari atasan.", "Prasetio, Ruminta & Akbar (2025)"),
        ],
        "synthesis": "kesadaran dan kesediaan pegawai menaati peraturan kedinasan, prosedur kerja, serta etika pelayanan secara konsisten sebagai wujud tanggung jawab profesional atas amanah jabatan",
        "dimensi_cite": "Moenir (dalam Winarto & Sayoto, 2018) dan Prasetio, Ruminta & Akbar (2025)",
        "dimensions": [
            {"dim": "Ketepatan Waktu", "desc": "Ketepatan Waktu: Ketepatan waktu dalam melaksanakan tugas dan penghematan waktu dalam melaksanakan tugas sehingga pekerjaan selesai sesuai jadwal yang ditetapkan.",
             "items": ["Saya melaksanakan tugas kedinasan dengan tepat waktu sesuai jadwal yang ditetapkan.",
                       "Saya berupaya menghemat waktu dalam melaksanakan setiap tugas kedinasan.",
                       "Saya menyelesaikan pekerjaan sebelum batas waktu yang ditentukan instansi."]},
            {"dim": "Tanggung Jawab Kerja", "desc": "Tanggung Jawab Kerja: Melakukan pekerjaan sesuai target, mengikuti evaluasi kerja, serta keberanian mengambil risiko dalam menyelesaikan tugas.",
             "items": ["Saya melakukan pekerjaan sesuai dengan target yang telah ditetapkan instansi.",
                       "Saya menerima hasil evaluasi kerja sebagai bahan perbaikan kinerja.",
                       "Saya berani mengambil risiko yang diperlukan dalam menyelesaikan tugas."]},
            {"dim": "Ketaatan", "desc": "Ketaatan: Ketaatan terhadap jam kerja, ketaatan terhadap pimpinan, serta ketaatan terhadap prosedur dan metode kerja yang berlaku.",
             "items": ["Saya taat terhadap jam kerja yang berlaku di instansi.",
                       "Saya taat terhadap pimpinan dalam melaksanakan tugas kedinasan.",
                       "Saya taat terhadap prosedur serta metode kerja yang ditetapkan instansi."]},
            {"dim": "Inisiatif", "desc": "Inisiatif: Kemampuan menyelesaikan masalah tanpa menunggu instruksi dari atasan serta kesediaan mengambil langkah proaktif dalam pekerjaan.",
             "items": ["Saya mampu menyelesaikan masalah pekerjaan tanpa menunggu instruksi dari atasan.",
                       "Saya mengambil langkah proaktif ketika menghadapi kendala dalam pelaksanaan tugas."]},
        ],
        "issue": "masih dijumpai keterlambatan masuk kerja, penggunaan waktu kerja yang belum efisien, serta penyelesaian laporan yang melewati batas waktu yang ditetapkan",
        "effect": "Disiplin kerja yang kokoh akan menjamin ketepatan waktu, kepatuhan prosedur, dan tuntasnya tanggung jawab, sehingga hasil kerja pegawai terjaga mutunya",
        "effect_long": "ketaatan pada waktu kerja, kepatuhan terhadap prosedur, dan rasa tanggung jawab terbukti menjadi fondasi terbentuknya kinerja pegawai yang konsisten dan akuntabel",
        "empirical": "Prasetio, Ruminta & Akbar (2025)",
        "likert_focus": "tingkat kedisiplinan kerja pegawai",
    },
    "kecerdasan emosional": {
        "name": "Kecerdasan Emosional",
        "cites": ["Goleman (2022)", "Luthans et al. (2022)", "Ningsih et al. (2023)"],
        "theory_cite": "Goleman (2022)",
        "concept": "Kemampuan mengenali, mengelola, dan memanfaatkan emosi secara konstruktif",
        "definition": [
            ("Kecerdasan emosional menjelaskan kapasitas individu dalam mengenali serta mengelola emosinya ketika berinteraksi di tempat kerja. Menurut Goleman (2022), kecerdasan emosional adalah kemampuan seseorang untuk mengenali perasaan diri sendiri dan perasaan orang lain, kemampuan memotivasi diri sendiri, serta mengelola emosi dengan baik pada diri sendiri dan dalam hubungan personal maupun profesional.", "Goleman (2022)"),
            ("Luthans et al. (2022) mengemukakan bahwa kecerdasan emosional menjadi modal psikologis kunci dalam perilaku organisasi karena menentukan kualitas kepemimpinan, kerja sama tim, dan ketahanan dalam menghadapi tekanan tugas operasional. Pegawai dengan kecerdasan emosional tinggi lebih mampu menjaga stabilitas sikap kerja pada situasi tekanan tugas yang tinggi.", "Luthans et al. (2022)"),
        ],
        "synthesis": "kemampuan pegawai dalam mengenali dan mengendalikan emosi diri, memahami perasaan orang lain, serta memanfaatkan kesadaran emosional tersebut untuk membangun kerja sama tim yang produktif",
        "dimensi_cite": "Goleman (2022) dan Luthans et al. (2022)",
        "dimensions": [
            {"dim": "Kesadaran Diri", "desc": "Kesadaran Diri (Self-Awareness): Kemampuan mengenali suasana hati, kekuatan, kelemahan, dan dorongan diri sendiri ketika menghadapi tugas atau konflik kerja.",
             "items": ["Saya mampu mengenali suasana hati yang memengaruhi cara kerja saya.",
                       "Saya memahami kekuatan dan kelemahan emosional diri saat menghadapi tekanan tugas."]},
            {"dim": "Pengelolaan Emosi & Regulasi Diri", "desc": "Regulasi Diri (Self-Management): Kemampuan mengendalikan emosi negatif, tetap tenang dalam tekanan, serta bersikap konsisten terhadap nilai yang diyakini.",
             "items": ["Saya tetap tenang dan mengendalikan diri saat menghadapi kritik pekerjaan.",
                       "Saya mampu menahan sikap reaktif ketika terjadi perbedaan pendapat dengan rekan kerja."]},
            {"dim": "Empati & Keterampilan Sosial", "desc": "Empati dan Keterampilan Sosial: Kemampuan memahami perasaan pihak lain serta membangun komunikasi dan jaringan kerja yang konstruktif.",
             "items": ["Saya berusaha memahami sudut pandang rekan kerja sebelum memberikan tanggapan.",
                       "Saya membangun komunikasi yang terbuka dengan rekan kerja dalam menyelesaikan tugas tim."]},
        ],
        "issue": "masih terdapat pegawai yang mudah terpancing emosi negatif saat menghadapi tekanan target kerja maupun keluhan masyarakat, sehingga komunikasi pelayanan menjadi kurang kondusif",
        "effect": "Kecerdasan emosional yang matang membantu pegawai mengelola tekanan kerja, menjaga hubungan interpersonal, dan mempertahankan konsistensi hasil kerja",
        "effect_long": "kesadaran diri, pengendalian emosi, dan empati terbukti mampu menjaga stabilitas sikap kerja serta kualitas interaksi pelayanan, sehingga kinerja pegawai tetap terjaga dalam kondisi beban kerja yang tinggi",
        "empirical": "Ningsih et al. (2023) dan Ausat et al. (2022)",
        "likert_focus": "kecerdasan emosional pegawai dalam interaksi kerja",
    },
    "kompensasi": {
        "name": "Kompensasi",
        "cites": ["Rivai & Sagala (2022)", "Mondy & Martocchio (2023)", "Dessler (2022)"],
        "theory_cite": "Rivai & Sagala (2022)",
        "concept": "Imbalan finansial dan nonfinansial atas kontribusi kerja pegawai",
        "definition": [
            ("Kompensasi merupakan balas jasa yang diberikan institusi atas kontribusi kerja pegawai. Menurut Rivai & Sagala (2022), kompensasi adalah sesuatu yang diterima oleh pegawai sebagai pengganti kontribusi jasa mereka pada institusi, baik berwujud imbalan finansial reguler maupun fasilitas kesejahteraan nonfinansial yang menjamin keberlangsungan hidup pegawai secara layak demi mendorong terciptanya kepuasan dan loyalitas kerja yang berkelanjutan.", "Rivai & Sagala (2022)"),
            ("Mondy & Martocchio (2023) menyebutkan bahwa sistem kompensasi mencakup totalitas seluruh imbalan yang dialokasikan organisasi sebagai bentuk pengakuan atas kesediaan aparatur mencurahkan dedikasi, waktu, dan keterampilan dalam menuntaskan serangkaian target operasional pelayanan. Dessler (2022) menegaskan bahwa keadilan distributif kompensasi menjadi faktor penting dalam membentuk persepsi positif pegawai terhadap institusi.", "Mondy & Martocchio (2023)"),
        ],
        "synthesis": "keseluruhan imbalan finansial maupun nonfinansial yang diterima pegawai sebagai balas jasa atas kontribusi kerja, yang apabila dirasakan adil akan mendorong kesungguhan dan loyalitas kerja",
        "dimensi_cite": "Rivai & Sagala (2022) dan Dessler (2022)",
        "dimensions": [
            {"dim": "Kompensasi Finansial Langsung", "desc": "Kompensasi Finansial Langsung: Gaji, tunjangan kinerja, dan insentif yang diterima pegawai sesuai dengan beban serta tanggung jawab jabatannya.",
             "items": ["Gaji yang diterima sesuai dengan beban kerja dan tanggung jawab jabatan.",
                       "Tunjangan kinerja yang diberikan instansi sebanding dengan capaian kerja."]},
            {"dim": "Kompensasi Tidak Langsung", "desc": "Kompensasi Tidak Langsung (Fasilitas): Jaminan kesehatan, fasilitas kerja, dan kesejahteraan nonfinansial yang menunjang pelaksanaan tugas.",
             "items": ["Instansi menyediakan jaminan kesehatan yang memadai bagi pegawai.",
                       "Fasilitas kerja yang tersedia memudahkan pelaksanaan tugas kedinasan."]},
            {"dim": "Keadilan Kompensasi", "desc": "Keadilan dan Ketepatan Waktu Pembayaran: Persepsi keadilan distributif serta ketepatan waktu pencairan kompensasi dari institusi.",
             "items": ["Pembagian kompensasi di instansi ini sudah dirasakan adil bagi seluruh pegawai.",
                       "Kompensasi diberikan secara tepat waktu sesuai ketentuan yang berlaku."]},
        ],
        "issue": "persepsi keadilan atas besaran tunjangan kinerja dan kelengkapan fasilitas kerja masih menjadi keluhan yang berulang di lingkungan instansi",
        "effect": "Kompensasi yang adil dan tepat waktu akan meningkatkan kepuasan kerja, menekan keinginan berpindah, dan mendorong pegawai bekerja lebih produktif",
        "effect_long": "keadilan kompensasi dan kelengkapan fasilitas terbukti mendorong semangat serta loyalitas kerja, sehingga pegawai terdorong mencapai hasil kerja yang lebih baik",
        "empirical": "Sedarmayanti (2022) dan Firdo et al. (2023)",
        "likert_focus": "persepsi pegawai atas kompensasi yang diterima",
    },
    "kompetensi": {
        "name": "Kompetensi",
        "cites": ["Wibowo (2022)", "Sudarmanto (2022)", "Sedarmayanti (2022)"],
        "theory_cite": "Wibowo (2022)",
        "concept": "Kapabilitas pengetahuan, keterampilan, dan sikap kerja pegawai dalam melaksanakan tugas kedinasan",
        "definition": [
            ("Kompetensi merupakan kemampuan dasar yang mencakup pengetahuan, keterampilan, dan sikap kerja yang dibutuhkan untuk melaksanakan pekerjaan secara efektif. Menurut Wibowo (2022), kompetensi adalah suatu kemampuan dasar seseorang untuk melaksanakan suatu pekerjaan yang berlandaskan atas keterampilan dan pengetahuan serta didukung oleh sikap kerja yang dituntut oleh jabatan tersebut guna mencapai kinerja yang unggul.", "Wibowo (2022)"),
            ("Sudarmanto (2022) yang merujuk pada teori Spencer & Spencer mendefinisikan kompetensi sebagai karakteristik mendasar yang dimiliki seorang individu yang berhubungan secara kausal dengan kriteria acuan efektivitas kerja atau kinerja superior dalam suatu tugas kedinasan. Sedarmayanti (2022) menegaskan bahwa kompetensi mencakup dimensi pengetahuan teknis, kemahiran operasional, dan integritas perilaku kerja.", "Sudarmanto (2022)"),
        ],
        "synthesis": "kapasitas kapabilitas terintegrasi yang dimiliki oleh pegawai—meliputi penguasaan pengetahuan kedinasan (knowledge), keahlian teknis operasional (skills), dan sikap kerja profesional (attitude)—yang diaplikasikan dalam menuntaskan tugas kedinasan demi menghasilkan kinerja bermutu tinggi",
        "dimensi_cite": "Wibowo (2022) dan Sudarmanto (2022)",
        "dimensions": [
            {"dim": "Pengetahuan (Knowledge)", "desc": "Pengetahuan (Knowledge): Penguasaan regulasi, standar operasional prosedur (SOP), dan wawasan teknis bidang tugas kedinasan.",
             "items": ["Menguasai pengetahuan teknis dan regulasi yang diperlukan dalam bidang tugas.",
                       "Memahami standar operasional prosedur (SOP) pelaksanaan kerja secara mendalam."]},
            {"dim": "Keterampilan (Skills)", "desc": "Keterampilan (Skills): Kemahiran dan ketelitian operasional dalam menggunakan sarana kerja serta menyelesaikan tugas kedinasan secara efektif.",
             "items": ["Memiliki keterampilan teknis yang memadai untuk menyelesaikan tugas kedinasan secara efisien.",
                       "Terampil mengoperasikan sarana dan sistem teknologi informasi penunjang pekerjaan."]},
            {"dim": "Sikap Kerja (Attitude)", "desc": "Sikap Kerja (Attitude): Integritas, tanggung jawab, dan kesiapan melayani serta beradaptasi dengan dinamika tugas.",
             "items": ["Menunjukkan dedikasi dan kesungguhan dalam menuntaskan setiap tanggung jawab tugas.",
                       "Bersikap adaptif dan profesional dalam merespons arahan kerja dan kebutuhan pelayanan."]},
        ],
        "issue": "kesenjangan pemahaman regulasi teknis baru serta perlunya peningkatan keterampilan operasional digital masih menjadi kendala bagi sebagian pegawai",
        "effect": "Kompetensi yang unggul memastikan akurasi pekerjaan, meminimalisir kesalahan prosedur, dan meningkatkan kecepatan penyelesaian tugas pelayanan",
        "effect_long": "penguasaan pengetahuan, keahlian teknis, dan sikap kerja profesional terbukti meminimalisir kesalahan operasional serta meningkatkan akurasi dan kecepatan penyelesaian tugas",
        "empirical": "Hidayat & Kusumawati (2023) dan Alamsyah & Permadi (2022)",
        "likert_focus": "kompetensi kerja yang dimiliki pegawai",
    },
    "lingkungan kerja": {
        "name": "Lingkungan Kerja",
        "cites": ["Sedarmayanti (2022)", "Mangkunegara (2022)", "Armstrong & Taylor (2023)"],
        "theory_cite": "Sedarmayanti (2022)",
        "concept": "Kondisi fisik dan nonfisik tempat kerja yang memengaruhi kenyamanan kerja",
        "definition": [
            ("Lingkungan kerja merupakan keseluruhan sarana dan prasarana yang ada di sekitar pegawai ketika melaksanakan pekerjaannya. Menurut Sedarmayanti (2022), lingkungan kerja mencakup lingkungan kerja fisik seperti tata ruang, penerangan, dan sirkulasi udara, serta lingkungan kerja nonfisik seperti hubungan kerja dengan atasan dan rekan sejawat.", "Sedarmayanti (2022)"),
            ("Mangkunegara (2022) menambahkan bahwa kondisi lingkungan kerja yang kondusif akan menurunkan tingkat kelelahan, meningkatkan konsentrasi, dan pada akhirnya memperbaiki kualitas hasil kerja. Armstrong & Taylor (2023) menegaskan bahwa kualitas tempat kerja merupakan bagian dari sistem manajemen sumber daya manusia yang memengaruhi kepuasan serta retensi pegawai.", "Mangkunegara (2022)"),
        ],
        "synthesis": "keseluruhan kondisi fisik dan nonfisik di sekitar tempat kerja yang memengaruhi kenyamanan, konsentrasi, dan keselamatan pegawai dalam melaksanakan tugas kedinasan",
        "dimensi_cite": "Sedarmayanti (2022) dan Mangkunegara (2022)",
        "dimensions": [
            {"dim": "Lingkungan Fisik", "desc": "Lingkungan Fisik: Tata ruang, penerangan, sirkulasi udara, kebersihan, dan kelengkapan peralatan kerja di ruang kerja pegawai.",
             "items": ["Ruang kerja saya tertata rapi, bersih, dan memiliki sirkulasi udara yang baik.",
                       "Penerangan dan peralatan kerja di ruang kerja saya mendukung kenyamanan bekerja."]},
            {"dim": "Lingkungan Nonfisik", "desc": "Lingkungan Nonfisik: Hubungan kerja yang harmonis dengan atasan, rekan sejawat, dan pihak lain yang berkepentingan dalam pelaksanaan tugas.",
             "items": ["Hubungan kerja dengan atasan berjalan harmonis dan saling mendukung.",
                       "Saya memperoleh dukungan yang baik dari rekan kerja dalam menyelesaikan tugas."]},
            {"dim": "Keselamatan & Kenyamanan Kerja", "desc": "Keselamatan dan Kenyamanan Kerja: Jaminan keamanan, ketenangan, dan bebasnya pegawai dari gangguan yang menghambat konsentrasi kerja.",
             "items": ["Lingkungan kerja saya terbebas dari gangguan yang mengganggu konsentrasi.",
                       "Instansi memperhatikan aspek keselamatan dan kesehatan kerja pegawai."]},
        ],
        "issue": "keterbatasan ruang kerja, kebisingan, serta kelengkapan sarana penunjang dirasakan mengganggu konsentrasi pegawai dalam menyelesaikan laporan kedinasan",
        "effect": "Lingkungan kerja yang nyaman menurunkan kelelahan, meningkatkan konsentrasi, dan memperlancar penyelesaian tugas secara tepat waktu",
        "effect_long": "kondisi fisik dan nonfisik tempat kerja yang kondusif terbukti menurunkan kelelahan serta meningkatkan konsentrasi, sehingga hasil kerja pegawai menjadi lebih cepat dan akurat",
        "empirical": "Sedarmayanti (2022) dan Armstrong & Taylor (2023)",
        "likert_focus": "kondisi lingkungan kerja yang dirasakan pegawai",
    },
    "kepuasan kerja": {
        "name": "Kepuasan Kerja",
        "cites": ["Robbins & Judge (2022)", "Hasibuan (2022)", "Ningsih et al. (2023)"],
        "theory_cite": "Robbins & Judge (2022)",
        "concept": "Sikap positif pegawai terhadap pekerjaan dan lingkungan kerjanya",
        "definition": [
            ("Kepuasan kerja merupakan sikap umum individu terhadap pekerjaan yang dijalaninya. Menurut Robbins & Judge (2022), kepuasan kerja adalah perasaan positif tentang pekerjaan seseorang yang merupakan hasil evaluasi atas karakteristik pekerjaan tersebut terhadap apa yang diharapkan. Kepuasan kerja mencerminkan kesesuaian antara harapan pegawai dengan kenyataan yang diterimanya di tempat kerja.", "Robbins & Judge (2022)"),
            ("Hasibuan (2022) menyatakan bahwa kepuasan kerja adalah sikap emosional yang menyenangkan dan mencintai pekerjaannya, yang tercermin dari moral kerja, kedisiplinan, dan prestasi kerja. Kepuasan kerja bersifat dinamis sehingga dapat meningkat maupun menurun bergantung pada kondisi internal institusi maupun kebutuhan individu pegawai.", "Hasibuan (2022)"),
        ],
        "synthesis": "sikap emosional positif pegawai yang timbul dari kesesuaian antara harapan terhadap pekerjaan dengan kenyataan yang diterima, tercermin dari moral kerja dan kedisiplinan yang tinggi",
        "dimensi_cite": "Robbins & Judge (2022) dan Hasibuan (2022)",
        "dimensions": [
            {"dim": "Kepuasan atas Pekerjaan Itu Sendiri", "desc": "Kepuasan atas Pekerjaan Itu Sendiri: Kesesuaian jenis tugas, tantangan pekerjaan, dan kesempatan aktualisasi kemampuan yang dimiliki.",
             "items": ["Pekerjaan yang dijalani saat ini sesuai dengan minat serta kemampuan yang dimiliki.",
                       "Memperoleh kesempatan yang luas untuk mengembangkan keahlian melalui pekerjaan."]},
            {"dim": "Kepuasan atas Imbalan & Pengawasan", "desc": "Kepuasan atas Imbalan dan Pengawasan: Keadilan imbalan yang diterima serta mutu bimbingan dan pengawasan dari pimpinan.",
             "items": ["Imbalan dan fasilitas yang diterima sebanding dengan beban tugas pekerjaan.",
                       "Pimpinan memberikan arahan dan bimbingan yang jelas dalam pelaksanaan tugas."]},
            {"dim": "Kepuasan atas Rekan Kerja", "desc": "Kepuasan atas Rekan Kerja dan Kondisi Kerja: Kualitas hubungan sosial, kerja sama antarrekan sejawat, dan keharmonisan lingkungan kerja.",
             "items": ["Suasana hubungan kerja dengan rekan-rekan kerja di instansi berjalan harmonis.",
                       "Rekan kerja senantiasa bersedia saling membantu dalam menyelesaikan tugas kedinasan."]},
        ],
        "issue": "sebagian pegawai mengeluhkan kurangnya kesempatan pengembangan diri serta komunikasi pimpinan yang belum menumbuhkan rasa dihargai",
        "effect": "Kepuasan kerja yang tinggi akan menurunkan tingkat keinginan keluar dan mendorong pegawai bekerja dengan moral serta produktivitas yang lebih tinggi",
        "effect_long": "kepuasan atas pekerjaan, imbalan, dan hubungan sosial di tempat kerja terbukti meningkatkan moral kerja yang pada akhirnya memperbaiki capaian kinerja pegawai",
        "empirical": "Ningsih et al. (2023) dan Ausat et al. (2022)",
        "likert_focus": "tingkat kepuasan kerja pegawai",
    },
    "komitmen organisasional": {
        "name": "Komitmen Organisasional",
        "cites": ["Robbins & Judge (2022)", "Armstrong & Taylor (2023)", "Ausat et al. (2022)"],
        "theory_cite": "Robbins & Judge (2022)",
        "concept": "Keterikatan pegawai terhadap organisasi dan kesediaan berkontribusi",
        "definition": [
            ("Komitmen organisasional menggambarkan keterikatan psikologis pegawai terhadap institusi tempatnya bekerja. Menurut Robbins & Judge (2022), komitmen organisasional adalah tingkat keterlibatan seseorang dalam organisasi serta kesediaannya untuk tetap menjadi bagian dari organisasi tersebut. Komitmen terbentuk melalui proses adaptasi nilai dan pengalaman kerja yang dialami pegawai.", "Robbins & Judge (2022)"),
            ("Armstrong & Taylor (2023) mengemukakan bahwa komitmen organisational mencakup tiga dimensi utama, yaitu komitmen afektif, kontinuans, dan normatif. Ausat et al. (2022) membuktikan bahwa komitmen yang kuat berkorelasi positif dengan kinerja karena pegawai bersedia mengerahkan usaha ekstra demi kepentingan institusi.", "Armstrong & Taylor (2023)"),
        ],
        "synthesis": "keterikatan psikologis pegawai terhadap institusi yang tercermin dari kesetiaan, kesediaan berusaha ekstra, serta keinginan bertahan menjadi bagian dari organisasi",
        "dimensi_cite": "Robbins & Judge (2022) dan Armstrong & Taylor (2023)",
        "dimensions": [
            {"dim": "Komitmen Afektif", "desc": "Komitmen Afektif: Ikatan emosional pegawai terhadap institusi yang timbul dari rasa memiliki dan kebanggaan sebagai bagian organisasi.",
             "items": ["Saya merasa memiliki ikatan emosional dan bangga menjadi bagian dari instansi ini.",
                       "Masalah yang dihadapi instansi saya rasakan sebagai bagian dari tanggung jawab saya."]},
            {"dim": "Komitmen Kontinuans", "desc": "Komitmen Kontinuans: Pertimbangan biaya dan manfaat yang membuat pegawai memilih bertahan di institusi tempatnya bekerja.",
             "items": ["Saya merasa bertahan di instansi ini merupakan pilihan yang rasional bagi karier saya.",
                       "Saya bertekad menyelesaikan masa pengabdian saya di lingkungan instansi ini."]},
            {"dim": "Komitmen Normatif", "desc": "Komitmen Normatif: Kewajiban moral dan rasa terima kasih yang mendorong pegawai memberikan kontribusi terbaik bagi institusi.",
             "items": ["Saya merasa memiliki kewajiban moral untuk memberikan kontribusi terbaik bagi instansi.",
                       "Instansi layak menerima dedikasi dan kesetiaan kerja saya."]},
        ],
        "issue": "masih ditemukan gejala penurunan loyalitas, ditandai keinginan sebagian pegawai untuk pindah unit kerja lain demi lingkungan yang dianggap lebih nyaman",
        "effect": "Komitmen yang kuat membuat pegawai bersedia mengerahkan usaha ekstra dan menjaga konsistensi kualitas kerja meskipun menghadapi keterbatasan sumber daya",
        "effect_long": "ikatan emosional, kesetiaan, dan kewajiban moral pegawai terhadap institusi terbukti mendorong kesediaan bekerja ekstra, sehingga kinerja tetap terjaga",
        "empirical": "Ausat et al. (2022) dan Efendi et al. (2023)",
        "likert_focus": "komitmen pegawai terhadap instansi",
    },
    "pengembangan karir": {
        "name": "Pengembangan Karir",
        "cites": ["Hasibuan (2022)", "Armstrong & Taylor (2023)", "Sedarmayanti (2022)"],
        "theory_cite": "Hasibuan (2022)",
        "concept": "Proses peningkatan jabatan dan kompetensi pegawai secara terencana",
        "definition": [
            ("Pengembangan karir berkaitan dengan upaya institusi memfasilitasi peningkatan kompetensi serta jenjang jabatan pegawai. Menurut Hasibuan (2022), pengembangan karir adalah perubahan status seseorang dalam suatu organisasi pada jalur karir yang telah ditetapkan oleh organisasi. Perencanaan karir yang jelas akan menumbuhkan harapan positif pegawai terhadap masa depan pekerjaannya.", "Hasibuan (2022)"),
            ("Armstrong & Taylor (2023) menegaskan bahwa pengembangan karier mencakup pendidikan, pelatihan, promosi, dan mutasi yang dirancang untuk memenuhi kebutuhan organisasi sekaligus aspirasi individu. Sedarmayanti (2022) menambahkan bahwa kejelasan pola karier menjadi faktor motivasional penting bagi aparatur pemerintahan.", "Armstrong & Taylor (2023)"),
        ],
        "synthesis": "proses terencana peningkatan kompetensi serta jenjang jabatan pegawai melalui pendidikan, pelatihan, dan promosi yang selaras dengan kebutuhan institusi",
        "dimensi_cite": "Hasibuan (2022) dan Armstrong & Taylor (2023)",
        "dimensions": [
            {"dim": "Kejelasan Jalur Karier", "desc": "Kejelasan Jalur Karier: Kejelasan informasi mengenai pola, tahapan, dan persyaratan promosi jabatan yang berlaku di institusi.",
             "items": ["Instansi memberikan informasi yang jelas mengenai pola pengembangan karir pegawai.",
                       "Saya memahami persyaratan yang dibutuhkan untuk memperoleh promosi jabatan."]},
            {"dim": "Kesempatan Pengembangan Kompetensi", "desc": "Kesempatan Pengembangan Kompetensi: Akses pegawai terhadap pendidikan, pelatihan, dan bimbingan teknis peningkatan kapasitas kerja.",
             "items": ["Saya memperoleh kesempatan mengikuti pelatihan peningkatan kompetensi kerja.",
                       "Instansi memfasilitasi saya untuk melanjutkan pendidikan pada jenjang yang lebih tinggi."]},
            {"dim": "Promosi & Penghargaan Jabatan", "desc": "Promosi dan Penghargaan Jabatan: Penerapan promosi berdasarkan prestasi kerja serta pemberian jabatan yang sesuai dengan kompetensi pegawai.",
             "items": ["Promosi jabatan di instansi ini dilakukan berdasarkan prestasi kerja pegawai.",
                       "Jabatan yang saya emban sekarang sesuai dengan kompetensi yang saya miliki."]},
        ],
        "issue": "kejelasan pola karier dan kesempatan mengikuti pelatihan dirasakan masih belum merata di antara seluruh pegawai pelaksana",
        "effect": "Kejelasan jalur karier dan kesempatan pengembangan kompetensi memperkuat harapan pegawai sehingga mendorong kesungguhan kerja yang lebih tinggi",
        "effect_long": "kejelasan pola karier dan kesempatan pengembangan kompetensi terbukti menumbuhkan harapan serta kesungguhan kerja pegawai dalam mencapai target kinerja",
        "empirical": "Efendi et al. (2023) dan Ningsih et al. (2023)",
        "likert_focus": "persepsi pegawai atas pengembangan karir",
    },
    "kinerja pegawai": {
        "name": "Kinerja Pegawai",
        "cites": ["Ulya & MutiaRahmah (2025)", "Bangun (2012)", "Mangkunegara (2022)", "Hasibuan (2022)"],
        "theory_cite": "Bangun (2012)",
        "concept": "Hasil unjuk kerja nyata pegawai dari segi kualitas, kuantitas, dan ketepatan waktu",
        "definition": [
            ("Kinerja merupakan muara akhir dari seluruh aktivitas manajemen dalam suatu organisasi. Menurut Bangun (2012, 233-234) yang dikutip Ulya & MutiaRahmah (2025), terdapat dimensi sebagai pengukur kinerja karyawan secara keseluruhan, yaitu jumlah pekerjaan yang menggambarkan kuantitas pekerjaan yang berhasil digarap karyawan sesuai dengan persyaratan atau standar pekerjaan, kualitas pekerjaan yang menuntut pemenuhan persyaratan kualitas yang ditetapkan organisasi, ketepatan waktu penyelesaian pekerjaan, kehadiran atau absensi selama bekerja, serta kemampuan kerja sama antar karyawan.", "Bangun (2012)"),
            ("Ulya & MutiaRahmah (2025) menegaskan bahwa pekerjaan tidak dapat diselesaikan tepat waktu maka akan mengganggu atau menghambat pekerjaan yang lainnya, sehingga memengaruhi jumlah dan kualitas hasil pekerjaan. Penilaian kinerja pegawai juga dapat dilihat dari segi bagaimana karyawan bekerja sama dengan sesama karyawan dalam pekerjaan yang memerlukan keterlibatan tim.", "Ulya & MutiaRahmah (2025)"),
        ],
        "synthesis": "performa pencapaian kerja nyata pegawai yang diukur dari segi kualitas output, kuantitas penyelesaian tugas, ketepatan waktu, dan efektivitas pelaksanaan tanggung jawab kedinasan",
        "dimensi_cite": "Bangun (2012) dan Ulya & MutiaRahmah (2025)",
        "dimensions": [
            {"dim": "Jumlah Pekerjaan", "desc": "Jumlah Pekerjaan: Kuantitas pekerjaan yang berhasil digarap karyawan sesuai dengan persyaratan atau standar pekerjaan yang ditetapkan.",
             "items": ["Jumlah pekerjaan yang saya selesaikan sesuai dengan standar pekerjaan yang ditetapkan.",
                       "Kuantitas pekerjaan yang saya garap tuntas dalam periode penilaian kinerja.",
                       "Volume pekerjaan yang saya hasilkan memenuhi harapan instansi."]},
            {"dim": "Kualitas Pekerjaan", "desc": "Kualitas Pekerjaan: Pemenuhan persyaratan kualitas pekerjaan yang ditetapkan organisasi, di mana masing-masing pekerjaan memiliki standar kualitas yang perlu disesuaikan oleh karyawan.",
             "items": ["Hasil pekerjaan saya memenuhi persyaratan kualitas yang ditetapkan organisasi.",
                       "Saya menghasilkan pekerjaan yang baik sesuai standar kualitas instansi.",
                       "Kualitas hasil kerja saya sesuai dengan standar yang telah ditentukan."]},
            {"dim": "Ketepatan Waktu", "desc": "Ketepatan Waktu: Pekerjaan diselesaikan sesuai dengan waktu yang ditentukan, sebab pekerjaan bersifat saling bergantung sehingga keterlambatan akan menghambat pekerjaan lainnya.",
             "items": ["Saya menyelesaikan pekerjaan sesuai dengan waktu yang ditentukan instansi.",
                       "Penyelesaian tugas saya tidak menghambat pekerjaan lainnya.",
                       "Saya menjaga ketergantungan antarpekerjaan dengan ketepatan waktu penyelesaian."]},
            {"dim": "Kehadiran", "desc": "Kehadiran: Aspek penilaian kinerja dilihat dari kehadiran atau absensi karyawan selama bekerja, yang menjadi salah satu ukuran baik tidaknya kinerja karyawan.",
             "items": ["Kehadiran atau absensi saya selama bekerja terjaga dengan baik.",
                       "Tingkat kehadiran saya mencerminkan kinerja yang baik di instansi.",
                       "Saya hadir secara konsisten selama jam kerja berlangsung."]},
            {"dim": "Kemampuan Kerja Sama", "desc": "Kemampuan Kerja Sama: Penilaian kinerja dari segi bagaimana karyawan bekerja sama dengan sesama karyawan dalam pekerjaan yang memerlukan keterlibatan tim.",
             "items": ["Saya mampu bekerja sama dengan sesama karyawan dalam menyelesaikan pekerjaan.",
                       "Saya terlibat dalam pekerjaan yang memerlukan kerja sama antarkaryawan.",
                       "Kerja sama saya dengan rekan kerja mendukung pencapaian target instansi."]},
        ],
        "issue": "capaian sasaran kinerja pegawai masih mengalami fluktuasi dan belum sepenuhnya mencapai target ideal yang telah ditetapkan instansi",
        "effect": "Pencapaian target kinerja mencerminkan efektivitas penerapan manajemen sumber daya manusia pada instansi",
        "effect_long": "kualitas output kerja, ketepatan waktu penyelesaian tugas, serta inisiatif kolaborasi tim terbukti menjadi indikator utama keberhasilan kinerja instansi secara keseluruhan",
        "empirical": "Ulya & MutiaRahmah (2025)",
        "likert_focus": "capaian kinerja kerja pegawai",
    },
    "pengalaman kerja": {
        "name": "Pengalaman Kerja",
        "cites": ["Foster (2022)", "Hasibuan (2022)", "Mangkunegara (2022)"],
        "theory_cite": "Foster (2022)",
        "concept": "Akumulasi wawasan, keterampilan, dan kemahiran teknis yang diperoleh pegawai selama masa kerja",
        "definition": [
            ("Pengalaman kerja merupakan modal fundamental yang menentukan kesiapan dan kematangan teknis pegawai dalam menuntaskan tanggung jawab kedinasan. Menurut Foster (2022), pengalaman kerja adalah ukuran tentang lama waktu atau masa kerja yang telah ditempuh seseorang dalam memahami tugas-tugas suatu pekerjaan dan telah melaksanakannya dengan baik. Semakin luas pengalaman kerja seorang pegawai, semakin tinggi kecakapan dan ketelitian dalam mengatasi persoalan operasional di lapangan.", "Foster (2022)"),
            ("Hasibuan (2022) menyatakan bahwa pengalaman kerja adalah proses pembentukan keahlian dan keterampilan melalui rentang pengabdian kerja yang berulang sehingga pegawai terbiasa menghadapi tantangan tugas secara efektif. Sementara itu, Mangkunegara (2022) menegaskan bahwa pengalaman kerja yang matang akan meningkatkan efisiensi waktu dan meminimalkan kesalahan kerja.", "Hasibuan (2022)"),
        ],
        "synthesis": "tingkat penguasaan pengetahuan, keahlian praktis, dan kecakapan teknis yang diperoleh pegawai melalui akumulasi masa kerja dan keterlibatan langsung dalam menyelesaikan tugas",
        "dimensi_cite": "Foster (2022) dan Hasibuan (2022)",
        "dimensions": [
            {"dim": "Masa Kerja & Jam Terbang", "desc": "Masa Kerja dan Jam Terbang: Rentang waktu dan pengalaman praktis yang telah dijalani pegawai dalam melaksanakan tugas kedinasan.",
             "items": ["Masa kerja yang ditempuh memberikan pemahaman menyeluruh terhadap alur tugas kedinasan.",
                       "Memiliki jam terbang yang cukup dalam menghadapi variasi pekerjaan di instansi."]},
            {"dim": "Pengetahuan & Keterampilan Teknis", "desc": "Pengetahuan dan Keterampilan Teknis: Penguasaan keahlian teknis dan wawasan operasional yang diperoleh dari riwayat pekerjaan sebelumnya.",
             "items": ["Pengalaman kerja terdahulu meningkatkan keterampilan teknis dalam menyelesaikan tugas.",
                       "Pengetahuan praktis yang dimiliki memudahkan pemecahan kendala operasional kerja."]},
            {"dim": "Penguasaan Pekerjaan & Prosedur", "desc": "Penguasaan Pekerjaan dan Prosedur: Kecakapan dalam mengoperasikan sarana kerja serta kepatuhan terhadap prosedur baku secara tuntas.",
             "items": ["Mampu mengoperasikan sarana dan fasilitas kerja pendukung tugas secara optimal.",
                       "Memahami seluruh standar prosedur pelaksanaan pekerjaan dengan tingkat ketelitian yang tinggi."]},
        ],
        "issue": "masih terdapat variasi tingkat pengalaman kerja antarpegawai sehingga kecepatan dan ketepatan penuntasan tugas operasional belum merata",
        "effect": "Pengalaman kerja yang matang membekali pegawai dengan kemahiran teknis dan ketenangan dalam menyelesaikan tugas secara akurat dan tepat waktu",
        "effect_long": "akumulasi masa kerja, penguasaan keterampilan teknis, dan kebiasaan mematuhi prosedur kerja terbukti signifikan meningkatkan kualitas dan ketepatan capaian kinerja pegawai",
        "empirical": "Wijayanti & Utomo (2024) dan Prasetyo & Yuliana (2023)",
        "likert_focus": "persepsi pegawai atas pengalaman kerja",
    },
    "keterikatan kerja": {
        "name": "Keterikatan Kerja",
        "cites": ["Schaufeli et al. (2022)", "Bakker & Demerouti (2022)", "Robbins & Judge (2022)"],
        "theory_cite": "Schaufeli et al. (2022)",
        "concept": "Keadaan psikologis positif pegawai yang dicirikan oleh semangat, dedikasi, dan konsentrasi kerja mendalam",
        "definition": [
            ("Keterikatan kerja (work engagement) mencerminkan keterlibatan emosional dan kognitif pegawai secara penuh terhadap pekerjaannya. Menurut Schaufeli et al. (2022), keterikatan kerja adalah kondisi pikiran yang positif, memuaskan, dan berkaitan dengan pekerjaan yang ditandai oleh semangat membara (vigor), dedikasi tinggi (dedication), dan keasyikan kerja yang mendalam (absorption). Pegawai yang terikat tidak sekadar hadir di tempat kerja, tetapi mencurahkan energi terbaik demi keberhasilan instansi.", "Schaufeli et al. (2022)"),
            ("Bakker & Demerouti (2022) dalam kerangka Job Demands-Resources mengemukakan bahwa keterikatan kerja timbul ketika ketersediaan sumber daya kerja mampu mengimbangi tuntutan tugas, sehingga mendorong motivasi intrinsik dan kinerja optimal. Robbins & Judge (2022) menambahkan bahwa keterikatan kerja yang kokoh berdampak langsung terhadap loyalitas, inisiatif, dan produktivitas kerja pegawai.", "Bakker & Demerouti (2022)"),
        ],
        "synthesis": "keadaan mental positif dan memuaskan yang dirasakan pegawai saat bekerja yang tercermin dari tingginya energi kerja (vigor), rasa bangga dan dedikasi (dedication), serta konsentrasi penuh (absorption)",
        "dimensi_cite": "Schaufeli et al. (2022) dan Bakker & Demerouti (2022)",
        "dimensions": [
            {"dim": "Vigor (Semangat & Ketahanan)", "desc": "Vigor: Tingkat energi mental, ketahanan fisik, dan kemauan untuk berusaha sekuat tenaga dalam bekerja.",
             "items": ["Memiliki energi dan stamina yang tinggi saat menjalankan tugas-tugas pekerjaan.",
                       "Tetap bersemangat dan tekun berusaha meskipun menghadapi tantangan kerja yang kompleks."]},
            {"dim": "Dedication (Rasa Bangga & Makna)", "desc": "Dedication: Keterlibatan emosional, antusiasme, rasa bermakna, dan kebanggaan terhadap tugas yang diemban.",
             "items": ["Menemukan makna positif dan antusiasme tinggi dalam setiap aktivitas kedinasan.",
                       "Merasa bangga terhadap kontribusi pekerjaan yang diberikan bagi kemajuan instansi."]},
            {"dim": "Absorption (Fokus & Keasyikan Kerja)", "desc": "Absorption: Konsentrasi mendalam dan ketenggelaman mental secara produktif dalam menuntaskan tugas.",
             "items": ["Mampu berkonsentrasi penuh dan fokus dalam pelaksanaan tugas kedinasan.",
                       "Waktu terasa berjalan cepat saat mencurahkan perhatian penuh pada penuntasan tanggung jawab pekerjaan."]},
        ],
        "issue": "antusiasme dan ketahanan kerja pegawai terkadang mengalami kelelahan emosional saat menghadapi beban kerja puncak pada periode pelaporan",
        "effect": "Keterikatan kerja yang tinggi mendorong pegawai mencurahkan kemampuan terbaiknya secara sukarela demi pencapaian sasaran institusi",
        "effect_long": "tingginya semangat kerja (vigor), kebanggaan terhadap institusi (dedication), dan fokus mendalam (absorption) terbukti secara konsisten mendongkrak capaian kinerja pegawai",
        "empirical": "Santoso & Dewi (2023) dan Novitasari & Asbari (2023)",
        "likert_focus": "tingkat keterikatan kerja pegawai",
    },
    "pelatihan": {
        "name": "Pelatihan",
        "cites": ["Noe et al. (2023)", "Dessler (2022)", "Mangkunegara (2022)"],
        "theory_cite": "Noe et al. (2023)",
        "concept": "Upaya terencana organisasi memfasilitasi penguasaan pengetahuan dan keterampilan kerja pegawai",
        "definition": [
            ("Pelatihan merupakan instrumen strategis dalam pembinaan sumber daya manusia guna menyesuaikan kapabilitas pegawai dengan dinamika tuntutan tugas. Menurut Noe et al. (2023), pelatihan adalah upaya terencana yang dirancang oleh organisasi untuk memfasilitasi pembelajaran pegawai terkait kompetensi yang berkaitan langsung dengan pekerjaan, meliputi pengetahuan operasional, keterampilan teknis, dan perilaku kerja yang mendukung pencapaian kinerja unggul.", "Noe et al. (2023)"),
            ("Dessler (2022) merumuskan pelatihan sebagai proses mengajarkan keterampilan baru yang dibutuhkan pegawai baru maupun lama dalam melaksanakan pekerjaannya secara efektif. Mangkunegara (2022) menegaskan bahwa pelatihan yang bermutu memperpendek kurva pembelajaran serta membekali pegawai dengan metode kerja modern yang lebih efisien.", "Dessler (2022)"),
        ],
        "synthesis": "proses pembelajaran terencana dan sistematis yang diselenggarakan instansi untuk mentransfer pengetahuan, meningkatkan keahlian teknis, dan memperkuat sikap kerja profesional pegawai",
        "dimensi_cite": "Noe et al. (2023) dan Dessler (2022)",
        "dimensions": [
            {"dim": "Materi & Relevansi Pelatihan", "desc": "Materi dan Relevansi Pelatihan: Kesesuaian kurikulum dan materi ajar dengan tuntutan tugas kedinasan sehari-hari.",
             "items": ["Materi pelatihan yang diselenggarakan relevan dengan kebutuhan pelaksanaan tugas di unit kerja.",
                       "Kurikulum pelatihan memperkaya wawasan operasional dan standar prosedur kerja terbaru."]},
            {"dim": "Metode & Kualifikasi Instruktur", "desc": "Metode dan Kualifikasi Instruktur: Kualitas penyampaian fasilitator serta efektivitas metode pembelajaran interaktif.",
             "items": ["Instruktur pelatihan menyampaikan materi secara sistematis, komunikatif, dan mudah dipahami.",
                       "Metode pelatihan berbasis praktik memudahkan penguasaan keterampilan kedinasan secara mendalam."]},
            {"dim": "Sarana & Evaluasi Pascapelatihan", "desc": "Sarana dan Evaluasi Pascapelatihan: Ketersediaan fasilitas belajar yang memadai serta monitoring evaluasi hasil pelatihan.",
             "items": ["Fasilitas dan sarana pendukung pelatihan memadai untuk menunjang kelancaran proses pembelajaran.",
                       "Instansi melakukan evaluasi berkala terhadap peningkatan performa kerja pegawai pascapelatihan."]},
        ],
        "issue": "program pelatihan teknis belum sepenuhnya menjangkau seluruh lini pegawai secara berkesinambungan sesuai pemetaan kompetensi terkini",
        "effect": "Pelatihan yang aplikatif secara langsung memperbarui keahlian teknis pegawai sehingga meningkatkan ketepatan dan efisiensi penyelesaian tugas",
        "effect_long": "kesesuaian materi pelatihan, kompetensi instruktur, dan sarana pelatihan yang memadai terbukti nyata meningkatkan kompetensi serta output kinerja pegawai",
        "empirical": "Wahyuni & Siregar (2023) dan Widodo & Puspitasari (2024)",
        "likert_focus": "persepsi pegawai atas program pelatihan",
    },
}

CONSTRUCT_ALIASES = {
    "pengalaman": "pengalaman kerja",
    "pengalaman kerja": "pengalaman kerja",
    "work experience": "pengalaman kerja",
    "keterikatan kerja": "keterikatan kerja",
    "keterikatan pegawai": "keterikatan kerja",
    "keterikatan": "keterikatan kerja",
    "work engagement": "keterikatan kerja",
    "employee engagement": "keterikatan kerja",
    "pelatihan": "pelatihan",
    "pelatihan kerja": "pelatihan",
    "training": "pelatihan",
    "diklat": "pelatihan",
    "pendidikan dan pelatihan": "pelatihan",
    "gaya kepemimpinan transformasional": "kepemimpinan transformasional",
    "kepemimpinan": "kepemimpinan transformasional",
    "gaya kepemimpinan": "kepemimpinan transformasional",
    "kepemimpinan transformational": "kepemimpinan transformasional",
    "budaya kerja": "budaya organisasi",
    "budaya": "budaya organisasi",
    "motivasi": "motivasi kerja",
    "disiplin": "disiplin kerja",
    "kedisiplinan kerja": "disiplin kerja",
    "kedisiplinan": "disiplin kerja",
    "kecerdasan emosi": "kecerdasan emosional",
    "eq": "kecerdasan emosional",
    "kompensasi kerja": "kompensasi",
    "lingkungan kerja fisik": "lingkungan kerja",
    "komitmen": "komitmen organisasional",
    "komitmen organisasi": "komitmen organisasional",
    "pengembangan karier": "pengembangan karir",
    "karier": "pengembangan karir",
    "kinerja": "kinerja pegawai",
    "prestasi kerja": "kinerja pegawai",
    "kinerja karyawan": "kinerja pegawai",
}


def normalize_construct_name(name: str) -> str:
    """Menormalkan nama konstruk untuk pencarian di perpustakaan."""
    key = " ".join((name or "").lower().split()).strip(" .,;:")
    if not key:
        return ""
    if key in CONSTRUCT_LIBRARY:
        return key
    if key in CONSTRUCT_ALIASES:
        return CONSTRUCT_ALIASES[key]
    # Normalisasi sinonim organisasi <-> organisasional
    cand1 = key.replace("organisasi", "organisasional")
    if cand1 in CONSTRUCT_LIBRARY:
        return cand1
    if cand1 in CONSTRUCT_ALIASES:
        return CONSTRUCT_ALIASES[cand1]
    cand2 = key.replace("organisasional", "organisasi")
    if cand2 in CONSTRUCT_LIBRARY:
        return cand2
    if cand2 in CONSTRUCT_ALIASES:
        return CONSTRUCT_ALIASES[cand2]
    # Pencocokan sebagian (nama konstruk yang lebih panjang)
    for known in CONSTRUCT_LIBRARY:
        if known in key or key in known:
            return known
    return ""


# ---------------------------------------------------------------------------
# 2. Profil Generik (Fallback) — Tanpa Referensi Karangan
# ---------------------------------------------------------------------------

def _generic_profile(name: str) -> Dict[str, Any]:
    n = name or "Variabel Bebas"
    return {
        "name": n,
        "cites": ["Armstrong & Taylor (2023)", "Mangkunegara (2022)", "Robbins & Judge (2022)"],
        "theory_cite": "Armstrong & Taylor (2023)",
        "concept": f"Faktor organisasi yang memengaruhi {n.lower()} karyawan",
        "definition": [
            (f"{n} merupakan salah satu variabel organisasional yang memperoleh perhatian penting dalam kajian manajemen sumber daya manusia. Armstrong & Taylor (2023) menjelaskan bahwa {n.lower()} berkaitan dengan praktik pengelolaan sumber daya manusia yang memengaruhi sikap, perilaku, dan capaian kerja karyawan dalam suatu organisasi.", "Armstrong & Taylor (2023)"),
            (f"Mangkunegara (2022) menegaskan bahwa efektivitas {n.lower()} ditentukan oleh kesesuaian antara kebijakan institusi dengan kebutuhan kerja pegawai. Robbins & Judge (2022) menambahkan bahwa persepsi pegawai atas variabel tersebut membentuk sikap kerja yang pada akhirnya tercermin pada kualitas hasil kerjanya.", "Mangkunegara (2022)"),
        ],
        "synthesis": f"kondisi {n.lower()} yang dipersepsikan pegawai sebagai dukungan organisasi terhadap pelaksanaan tugas, sehingga memengaruhi kesungguhan dan hasil kerja yang dicapai",
        "dimensi_cite": "Armstrong & Taylor (2023) dan Mangkunegara (2022)",
        "dimensions": [
            {"dim": "Perencanaan", "desc": f"Perencanaan: Kejelasan perencanaan dan penetapan target {n.lower()} yang disusun bersama antara pimpinan dan pegawai.",
             "items": [f"Target {n.lower()} di instansi ini ditetapkan secara jelas dan terukur.",
                       f"Perencanaan {n.lower()} disusun dengan melibatkan pegawai pelaksana."]},
            {"dim": "Pelaksanaan", "desc": f"Pelaksanaan: Konsistensi penerapan {n.lower()} dalam aktivitas kerja sehari-hari di lingkungan instansi.",
             "items": [f"Penerapan {n.lower()} di instansi ini berjalan konsisten.",
                       f"Pegawai memperoleh dukungan yang memadai dalam pelaksanaan {n.lower()}."]},
            {"dim": "Evaluasi", "desc": f"Evaluasi dan Pengendalian: Proses pemantauan serta perbaikan berkelanjutan atas pelaksanaan {n.lower()}.",
             "items": [f"Evaluasi {n.lower()} dilakukan secara berkala oleh pimpinan.",
                       f"Hasil evaluasi {n.lower()} ditindaklanjuti dengan perbaikan yang nyata."]},
        ],
        "issue": f"penerapan {n.lower()} di lingkungan instansi belum berjalan optimal dan cenderung belum memperoleh perhatian yang memadai dari jajaran pimpinan",
        "effect": f"{n} yang kondusif akan mendorong pegawai bekerja lebih sungguh-sungguh sehingga capaian kinerjanya meningkat",
        "effect_long": f"kondisi {n.lower()} yang dipersepsikan positif terbukti mendorong kesungguhan kerja pegawai, sehingga hasil kerja yang dicapai menjadi lebih optimal",
        "empirical": "Armstrong & Taylor (2023) dan Mangkunegara (2022)",
        "likert_focus": f"persepsi pegawai atas {n.lower()}",
    }


def get_construct_profile(name: str) -> Dict[str, Any]:
    """Mengambil profil konstruk (definisi, dimensi, butir, narasi) berdasarkan nama."""
    key = normalize_construct_name(name)
    if key and key in CONSTRUCT_LIBRARY:
        profile = dict(CONSTRUCT_LIBRARY[key])
        # Pastikan nama tetap mengikuti input pengguna (kapitalisasi akademik)
        if name and name.strip():
            profile["name"] = name.strip()
        return profile

    # Cek basis ontologi dinamis hasil riset S1
    _load_cached_dynamic_ontology()
    name_clean = (name or "").strip().lower()

    # 1. Cek langsung kunci dinamis
    if key and key in DYNAMIC_S1_LIBRARY:
        return dict(DYNAMIC_S1_LIBRARY[key])
    if name_clean in DYNAMIC_S1_LIBRARY:
        return dict(DYNAMIC_S1_LIBRARY[name_clean])

    # 2. Cek kecocokan parsial / substring
    for dk, dprof in DYNAMIC_S1_LIBRARY.items():
        if dk and (dk in name_clean or name_clean in dk):
            return dict(dprof)

    return _generic_profile(name)


def clean_item_statement(text: str) -> str:
    """Membersihkan kata 'saya' / 'Saya' dan menyelaraskan kalimat pernyataan kuesioner formal."""
    import re
    t = str(text or "")
    t = re.sub(r'\b[Ss]aya\s+', '', t)
    t = re.sub(r'\s+[Ss]aya\b', '', t)
    t = re.sub(r'\s{2,}', ' ', t).strip()
    if t:
        t = t[0].upper() + t[1:]
    return t


def construct_items(code: str, name: str, max_dims: Optional[int] = None) -> List[Dict[str, str]]:
    """
    Menghasilkan daftar butir kuesioner untuk satu konstruk.

    Kode butir mengikuti pola ``X1_1``, ``X1_2``, ... sehingga konsisten dengan
    konvensi penamaan kolom pipeline SPSS. Mendukung jumlah dimensi dan butir
    asimetris alami sesuai instrumen penelitian operasional yang sahih.
    """
    profile = get_construct_profile(name)
    items: List[Dict[str, str]] = []
    counter = 1
    raw_dims = profile.get("dimensions") or profile.get("dimensi") or []
    dims = raw_dims[:max_dims] if max_dims else raw_dims
    for dim in dims:
        dname = dim.get("dim") or dim.get("nama") or "Dimensi"
        ditems = dim.get("items") or dim.get("indikator") or []
        for text in ditems:
            t_str = text.get("p") or text.get("t") if isinstance(text, dict) else str(text)
            t_clean = clean_item_statement(t_str)
            items.append({
                "dim": dname,
                "code": f"{code}_{counter}",
                "text": t_clean,
            })
            counter += 1
    return items


def construct_dimension_labels(name: str, limit: int = 4) -> List[str]:
    """Label dimensi ringkas untuk diagram kerangka berpikir."""
    profile = get_construct_profile(name)
    raw_dims = profile.get("dimensions") or profile.get("dimensi") or []
    dims = raw_dims[:limit] if limit else raw_dims
    res = []
    for i, d in enumerate(dims):
        dname = d.get("dim") or d.get("nama") or f"Dimensi {i+1}"
        res.append(f"{i + 1}. {dname}")
    return res


# ---------------------------------------------------------------------------
# 2. Registrasi Otomatis Preset Topik (Pemasaran UMKM / Perbankan / SDM RS)
# ---------------------------------------------------------------------------

PRESET_CONSTRUCT_KEYS: List[str] = []


def register_topic_presets() -> int:
    """
    Mendaftarkan konstruk preset topik ke CONSTRUCT_LIBRARY.

    Kunci preset sengaja memakai nama panjang (mis. "kompensasi finansial",
    "lingkungan kerja fisik") agar pencocokan exact-key pada
    ``normalize_construct_name`` selalu menang atas pencocokan substring
    terhadap konstruk generik lama.
    """
    global PRESET_CONSTRUCT_KEYS
    try:
        import s1_topic_presets as _tp
    except Exception:
        return 0
    added = 0
    for key, prof in _tp.iter_constructs():
        k = " ".join((key or "").lower().split()).strip(" .,;:")
        if not k:
            continue
        if k not in PRESET_CONSTRUCT_KEYS:
            PRESET_CONSTRUCT_KEYS.append(k)
        if k not in CONSTRUCT_LIBRARY:
            CONSTRUCT_LIBRARY[k] = dict(prof)
            added += 1
    return added


register_topic_presets()
