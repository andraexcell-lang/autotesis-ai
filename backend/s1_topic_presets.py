# backend/s1_topic_presets.py
"""
REGISTRI PRESET TOPIK SKRIPSI S1
================================
Menggabungkan preset konten (konstruk, definisi Bab II, Tabel 2.1, bibliografi)
dengan preset sektoral (narasi Latar Belakang, Tabel 1.1 & 1.2, leksikon istilah).

Fungsi publik:
- get_topic_preset(ctx)          -> dict preset sektoral atau None
- get_construct(name)            -> profil konstruk preset atau None
- iter_constructs()              -> seluruh profil konstruk preset
- iter_bab2_definitions()        -> seluruh definisi Bab II preset
- iter_bibliography()            -> seluruh entri pustaka preset
- apply_lexicon(text, preset)    -> penyesuaian istilah sektor
"""

from typing import Any, Dict, List, Optional

import s1_presets_marketing as MKT
import s1_presets_banking as BANK
import s1_presets_hr as HR

# ---------------------------------------------------------------------------
# 1. PRESET SEKTORAL — NARASI LATAR BELAKANG & TABEL EMPIRIS
# ---------------------------------------------------------------------------

SECTOR_MARKETING = {
    "id": "umkm_marketing",
    "keywords": ["USAHA MIKRO", "UMKM", "KINERJA PEMASARAN", "ORIENTASI KEWIRAUSAHAAN"],
    "sektor_label": "sektor usaha mikro, kecil, dan menengah (UMKM) kuliner",
    "disiplin_label": "manajemen pemasaran usaha",
    "tipe_lembaga": "sentra usaha mikro, kecil, dan menengah pada sektor kuliner yang menjadi tulang punggung perekonomian daerah",
    "par1_text": (
        "Peran strategis usaha mikro, kecil, dan menengah (UMKM) sebagai penyangga perekonomian daerah menuntut setiap pelaku usaha "
        "membangun kapabilitas pemasaran yang adaptif terhadap perubahan selera konsumen dan tekanan persaingan pasar. Dalam kerangka "
        "Resource-Based View, keunggulan usaha lahir dari kombinasi sumber daya internal yang bernilai, langka, dan sulit ditiru, termasuk "
        "keberanian berinovasi serta kepekaan membaca peluang pasar (Barney et al., 2023). Lumpkin et al. (2023) menegaskan bahwa pelaku usaha "
        "yang berorientasi kewirausahaan tinggi mampu mendahului pesaing dalam menghadirkan penawaran baru, sedangkan Kotler & Keller (2022) "
        "menempatkan inovasi produk sebagai inti penciptaan nilai bagi konsumen. Pada sektor kuliner yang sangat padat pesaing, kemampuan "
        "memadukan orientasi kewirausahaan dan inovasi produk menjadi penentu utama keberlanjutan usaha."
    ),
    "par2_text": (
        "{instansi} merupakan {tipe_lembaga} yang menjadi contoh empiris bagaimana tekanan persaingan, kenaikan biaya bahan baku, dan "
        "pergeseran pola konsumsi konsumen memengaruhi capaian pemasaran usaha anggotanya. Untuk menjaga keberlanjutan usaha, pembinaan "
        "diarahkan pada peningkatan kapasitas kewirausahaan, pembaruan produk, dan perluasan kanal pemasaran digital. Kendati demikian, "
        "evaluasi kinerja pembinaan menunjukkan bahwa capaian indikator pemasaran belum sepenuhnya memenuhi target yang ditetapkan "
        "(Tjiptono, 2022). Rauch et al. (2022) menegaskan bahwa kesenjangan kinerja usaha kecil umumnya berakar pada lemahnya orientasi "
        "kewirausahaan dan terbatasnya pembaruan produk yang dihasilkan pelaku usaha."
    ),
    "t11": {
        "title": "Target dan Realisasi Indikator Kinerja Pemasaran pada {instansi} Tahun 2025",
        "headers": ["No", "Indikator Kinerja Pemasaran", "Target (%)", "Realisasi (%)", "Capaian (%)", "Kategori"],
        "source": "Laporan Pembinaan dan Monitoring Kinerja {instansi}, diolah 2026",
        "rows": [
            ("Pertumbuhan Omzet Penjualan Anggota UMKM", 100.0, 84.65),
            ("Perluasan Jangkauan Pasar dan Pelanggan Baru", 100.0, 82.30),
            ("Tingkat Retensi Pelanggan Usaha Kuliner", 95.0, 79.85),
            ("Adopsi Inovasi Produk dan Kemasan Baru", 98.0, 86.40),
            ("Efektivitas Promosi pada Kanal Pemasaran Digital", 90.0, 81.15),
        ],
        "gap_sentence": "efektivitas promosi digital serta retensi pelanggan usaha kuliner",
        "gap_paragraph_tail": "kapabilitas kewirausahaan dan pembaruan produk pelaku usaha belum berjalan secara optimal",
    },
    "t12_kind": "umkm_omzet",
    "t12": {
        "title": "Rekapitulasi Omzet dan Volume Transaksi Anggota Usaha pada {instansi} Semester Terakhir",
        "source": "Rekapitulasi Transaksi Usaha Anggota {instansi}, diolah 2026",
        "lead": (
            "Selain kesenjangan capaian indikator pemasaran di atas, permasalahan empiris lain terlihat pada fluktuasi omzet dan volume "
            "transaksi anggota usaha dari bulan ke bulan. Naik turunnya omzet mencerminkan daya tahan usaha dalam menjaga minat beli "
            "konsumen pada tengah tekanan biaya bahan baku dan persaingan harga (Tjiptono, 2022). Gejolak transaksi tersebut perlu "
            "dicermati karena menjadi indikator awal melemahnya kinerja pemasaran usaha (Morgan et al., 2022). Rekapitulasi omzet dan "
            "volume transaksi anggota usaha disajikan pada Tabel 1.2 berikut:"
        ),
    },
    "par_identifikasi": (
        "Permasalahan belum optimalnya pencapaian kinerja pemasaran tersebut diduga kuat berakar pada kapabilitas internal pelaku usaha, "
        "khususnya aspek {x_join_lower}."
    ),
    "gap1_cite_pos": "Ananda & Prasetyo (2023)",
    "gap1_cite_neg": "Fauziah & Kurniawan (2022)",
    "gap1_neg_reason": "bahwa orientasi kewirausahaan tidak selalu berpengaruh langsung terhadap kinerja pemasaran karena pengaruhnya terlebih dahulu dimediasi keunggulan bersaing yang dibangun usaha",
    "gap2_cites": "Wulandari & Hidayat (2022) serta Setyawan & Rahmawati (2023)",
    "gap2_neg_cites": "Ramadhan & Wijaya (2024) dan Hakim & Salsabila (2022)",
    "gap2_neg_reason": "menemukan bahwa inovasi kemasan dan sikap proaktif tidak selalu berdampak pada kinerja pemasaran apabila kapasitas produksi serta akses pasar usaha masih terbatas",
    "lexicon": {
        "Pegawai": "Pelaku UMKM",
        "pegawai": "pelaku UMKM",
        "Aparatur": "Pelaku Usaha",
        "aparatur": "pelaku usaha",
        "Instansi": "UMKM",
        "instansi": "UMKM",
        "Masa Kerja": "Lama Usaha",
        "masa kerja": "lama usaha",
    },
}

SECTOR_BANKING = {
    "id": "bank_digital",
    "keywords": ["BANK", "NASABAH", "LAYANAN DIGITAL", "KEAMANAN SISTEM INFORMASI"],
    "sektor_label": "industri perbankan swasta pada segmen layanan digital",
    "disiplin_label": "manajemen pemasaran jasa perbankan",
    "tipe_lembaga": "bank swasta nasional yang mengandalkan kanal layanan digital sebagai tulang punggung transaksi nasabah",
    "par1_text": (
        "Transformasi digital industri perbankan telah mengubah cara nasabah berinteraksi dengan penyedia jasa keuangan, sehingga kualitas "
        "layanan digital menjadi penentu utama daya saing bank. Dalam kerangka kualitas layanan elektronik, penilaian nasabah terbentuk dari "
        "efisiensi antarmuka, keandalan transaksi, ketersediaan sistem, dan jaminan privasi data (Parasuraman et al., 2022). Zeithaml et al. "
        "(2023) menegaskan bahwa pengalaman digital yang konsisten pada setiap titik sentuh layanan menumbuhkan kepuasan yang bermuara pada "
        "loyalitas nasabah. Pada saat yang sama, keamanan sistem informasi menjadi prasyarat karena transaksi keuangan melibatkan kerentanan "
        "data pribadi yang menuntut perlindungan menyeluruh (Whitman & Mattord, 2022)."
    ),
    "par2_text": (
        "{instansi} merupakan {tipe_lembaga} yang menghadapi tantangan menjaga mutu layanan digital sekaligus mempertahankan kepercayaan "
        "nasabah di tengah meningkatnya ancaman penipuan berbasis rekayasa sosial. Untuk menjaga kinerja layanan, bank menetapkan sasaran "
        "mutu yang terukur pada aspek kecepatan respons, keberhasilan transaksi, dan penyelesaian pengaduan nasabah. Kendati demikian, "
        "evaluasi kinerja layanan menunjukkan bahwa sejumlah indikator mutu belum mencapai target yang ditetapkan (Palmatier et al., 2022). "
        "Kim et al. (2022) menegaskan bahwa penurunan persepsi keamanan dan kualitas layanan secara langsung memperlemah kepercayaan serta "
        "niat nasabah mempertahankan hubungan dengan bank."
    ),
    "t11": {
        "title": "Target dan Realisasi Indikator Mutu Layanan Digital pada {instansi} Tahun 2025",
        "headers": ["No", "Indikator Mutu Layanan", "Target (%)", "Realisasi (%)", "Capaian (%)", "Kategori"],
        "source": "Laporan Kinerja Layanan Digital {instansi}, diolah 2026",
        "rows": [
            ("Tingkat Keberhasilan Transaksi pada Kanal Digital", 100.0, 86.20),
            ("Kecepatan Respons Layanan Bantuan Nasabah", 100.0, 82.75),
            ("Penyelesaian Pengaduan Nasabah Sesuai Standar Waktu", 95.0, 80.40),
            ("Kepatuhan Pengendalian Keamanan Sistem Informasi", 98.0, 87.10),
            ("Indeks Kepuasan Nasabah atas Layanan Digital", 90.0, 83.55),
        ],
        "gap_sentence": "kecepatan respons layanan bantuan serta penyelesaian pengaduan nasabah",
        "gap_paragraph_tail": "keandalan kanal digital dan perlindungan data nasabah belum terkelola secara optimal",
    },
    "t12_kind": "bank_complaint",
    "t12": {
        "title": "Rekapitulasi Gangguan Sistem dan Keluhan Nasabah pada {instansi} Semester Terakhir",
        "source": "Rekapitulasi Pusat Layanan Nasabah {instansi}, diolah 2026",
        "lead": (
            "Selain kesenjangan capaian indikator mutu di atas, permasalahan empiris lain tercermin pada rekapitulasi gangguan sistem dan "
            "keluhan nasabah yang berfluktuasi setiap bulan. Frekuensi gangguan sistem dan keluhan transaksi merupakan cerminan langsung "
            "dari keandalan kanal layanan digital yang dirasakan nasabah (Parasuraman et al., 2022). Akumulasi keluhan yang tidak "
            "terselesaikan secara cepat berpotensi menurunkan kepercayaan nasabah terhadap bank (Kim et al., 2022). Rekapitulasi gangguan "
            "sistem dan keluhan nasabah disajikan pada Tabel 1.2 berikut:"
        ),
    },
    "par_identifikasi": (
        "Permasalahan belum optimalnya loyalitas nasabah tersebut diduga kuat berakar pada mutu layanan yang diterima nasabah, khususnya "
        "aspek {x_join_lower}."
    ),
    "gap1_cite_pos": "Mahendra & Kirana (2023)",
    "gap1_cite_neg": "Anggraini & Setiawan (2023)",
    "gap1_neg_reason": "bahwa kualitas layanan elektronik tidak selalu berpengaruh langsung terhadap retensi nasabah karena pengaruhnya bergantung pada persepsi keamanan yang terbentuk selama bertransaksi",
    "gap2_cites": "Puspitasari & Nugroho (2024) serta Yulianto & Permata (2024)",
    "gap2_neg_cites": "Nabila & Ramadhan (2022) dan Kusumawati & Handoko (2022)",
    "gap2_neg_reason": "menemukan bahwa kepercayaan dan persepsi keamanan tidak selalu berujung pada loyalitas apabila kepuasan atas pengalaman transaksi digital belum terbentuk secara konsisten",
    "lexicon": {
        "Pegawai": "Nasabah",
        "pegawai": "nasabah",
        "Aparatur": "Nasabah",
        "aparatur": "nasabah",
        "Instansi": "Bank",
        "instansi": "bank",
    },
}

SECTOR_HOSPITAL = {
    "id": "rs_turnover",
    "keywords": ["RUMAH SAKIT", "TURNOVER", "KARYAWAN", "TENAGA KESEHATAN"],
    "sektor_label": "industri pelayanan kesehatan swasta pada aspek retensi tenaga kesehatan",
    "disiplin_label": "manajemen sumber daya manusia",
    "tipe_lembaga": "rumah sakit swasta yang menyelenggarakan pelayanan kesehatan rujukan dengan tenaga kesehatan sebagai penggerak utama layanan",
    "par1_text": (
        "Keberlanjutan mutu pelayanan rumah sakit sangat bergantung pada kestabilan tenaga kesehatan yang menjalankan asuhan medis secara "
        "berkesinambungan. Dalam kerangka turnover process theory, keputusan karyawan meninggalkan organisasi terbentuk melalui rangkaian "
        "evaluasi atas kepuasan kerja, komitmen, dan peluang kerja alternatif yang tersedia (Hom et al., 2022). Griffeth et al. (2023) melalui "
        "meta-analisis menegaskan bahwa kepuasan kerja dan komitmen organisasi merupakan determinan utama niat berpindah karyawan. Pada "
        "industri padat karya profesional seperti rumah sakit, tekanan kerja, keadilan imbalan, dan kenyamanan lingkungan fisik menjadi "
        "faktor yang paling sering dikaitkan dengan tingginya angka pengunduran diri tenaga kesehatan (Robbins & Judge, 2022)."
    ),
    "par2_text": (
        "{instansi} merupakan {tipe_lembaga} yang menghadapi tantangan mempertahankan tenaga kesehatan pada unit pelayanan dengan beban "
        "kerja tinggi. Untuk menjaga kesinambungan pelayanan, manajemen menetapkan sasaran retensi karyawan yang terukur melalui indikator "
        "kehadiran, kepatuhan standar pelayanan, dan tingkat pengunduran diri tahunan. Kendati demikian, evaluasi kinerja menunjukkan bahwa "
        "sejumlah indikator retensi belum mencapai target yang ditetapkan (Holtom et al., 2022). Hom et al. (2022) menegaskan bahwa "
        "akumulasi ketidakpuasan kerja yang tidak dikelola akan berkembang menjadi niat keluar dan berujung pada pengunduran diri aktual "
        "yang membebani biaya rekrutmen serta kesinambungan asuhan pasien."
    ),
    "t11": {
        "title": "Target dan Realisasi Indikator Retensi Tenaga Kesehatan pada {instansi} Tahun 2025",
        "headers": ["No", "Indikator Retensi Karyawan", "Target (%)", "Realisasi (%)", "Capaian (%)", "Kategori"],
        "source": "Laporan Kinerja SDM dan Mutu Pelayanan {instansi}, diolah 2026",
        "rows": [
            ("Tingkat Retensi Tenaga Kesehatan Tetap", 100.0, 83.90),
            ("Kepatuhan Kehadiran pada Giliran Dinas Jaga", 100.0, 85.25),
            ("Penyelesaian Program Pembinaan dan Konseling Karyawan", 95.0, 79.60),
            ("Kesesuaian Remunerasi dengan Beban dan Risiko Kerja", 98.0, 82.45),
            ("Indeks Kepuasan Kerja Tenaga Kesehatan", 90.0, 80.85),
        ],
        "gap_sentence": "penyelesaian program pembinaan karyawan serta kesesuaian remunerasi dengan beban kerja",
        "gap_paragraph_tail": "pengelolaan beban kerja, keadilan imbalan, dan kenyamanan lingkungan kerja fisik belum berjalan secara optimal",
    },
    "t12_kind": "absensi_turnover",
    "t12": {
        "title": "Rekapitulasi Ketidakhadiran dan Pengunduran Diri Karyawan pada {instansi} Semester Terakhir",
        "source": "Bagian SDM dan Kepegawaian {instansi}, diolah 2026",
        "lead": (
            "Selain kesenjangan capaian indikator retensi di atas, permasalahan empiris lain terlihat pada rekapitulasi ketidakhadiran dan "
            "pengunduran diri karyawan yang berfluktuasi setiap bulan. Ketidakhadiran yang tinggi merupakan sinyal awal melemahnya "
            "keterikatan karyawan terhadap organisasi (Hom et al., 2022). Akumulasi izin, sakit, dan ketidakhadiran tanpa keterangan yang "
            "disertai pengunduran diri menunjukkan tekanan kerja yang belum terkelola dengan baik (Griffeth et al., 2023). Rekapitulasi "
            "ketidakhadiran dan pengunduran diri karyawan disajikan pada Tabel 1.2 berikut:"
        ),
    },
    "par_identifikasi": (
        "Permasalahan tingginya turnover intention tersebut diduga kuat berakar pada kondisi kerja yang dirasakan karyawan, khususnya "
        "aspek {x_join_lower}."
    ),
    "gap1_cite_pos": "Aditya & Pramesti (2023)",
    "gap1_cite_neg": "Nuraini & Pratiwi (2024)",
    "gap1_neg_reason": "bahwa pengaruh stres kerja terhadap niat berpindah melemah pada unit kerja dengan budaya organisasi yang suportif sehingga hasilnya tidak konsisten antar-konteks",
    "gap2_cites": "Sari & Handayani (2022) serta Sulistyowati & Hidayat (2022)",
    "gap2_neg_cites": "Wibisono & Astuti (2023) dan Yuliana & Kurnia (2023)",
    "gap2_neg_reason": "menemukan bahwa perbaikan lingkungan kerja dan imbalan tidak selalu menurunkan niat keluar apabila tekanan kerja pada unit pelayanan masih tinggi dan dukungan atasan belum memadai",
    "lexicon": {
        "Pegawai": "Karyawan",
        "pegawai": "karyawan",
        "Aparatur": "Tenaga Kesehatan",
        "aparatur": "tenaga kesehatan",
        "Instansi": "Rumah Sakit",
        "instansi": "rumah sakit",
    },
}

SECTORS = [SECTOR_MARKETING, SECTOR_BANKING, SECTOR_HOSPITAL]


# ---------------------------------------------------------------------------
# 1b. PROFIL INSTANSI (SEJARAH, VISI, MISI, STRUKTUR) PER SEKTOR
# ---------------------------------------------------------------------------

SECTOR_MARKETING["instansi_profile"] = {
    "sejarah": (
        "{instansi} merupakan wadah pembinaan dan pengembangan usaha mikro, kecil, dan menengah yang menghimpun pelaku usaha "
        "kuliner di Kota Palembang. Lembaga ini tumbuh dari kebutuhan bersama para pelaku usaha untuk memperoleh pendampingan "
        "usaha, akses pembiayaan, serta perluasan jaringan pemasaran produk. Seiring berjalannya waktu, peran lembaga ini "
        "berkembang menjadi mitra strategis pemerintah daerah dalam penguatan kapasitas kewirausahaan dan digitalisasi "
        "pemasaran usaha mikro dan kecil."
    ),
    "visi": (
        "Visi {instansi} adalah: \"Terwujudnya Usaha Mikro, Kecil, dan Menengah Kuliner yang Mandiri, Inovatif, dan Berdaya "
        "Saing di Tingkat Regional maupun Nasional\"."
    ),
    "misi": [
        "Meningkatkan kapasitas kewirausahaan pelaku usaha melalui pendampingan, pelatihan, dan business coaching yang berkelanjutan.",
        "Mendorong pengembangan inovasi produk, kemasan, dan mutu layanan agar produk anggota mampu bersaing di pasar yang semakin kompetitif.",
        "Memperluas akses pasar dan pembiayaan melalui kemitraan dengan lembaga keuangan, marketplace digital, serta jaringan distribusi regional.",
    ],
    "struktur": (
        "Struktur organisasi {instansi} dipimpin oleh seorang Ketua yang membawahi Sekretaris, Bendahara, serta tiga bidang "
        "operasional, yaitu Bidang Produksi dan Mutu Produk, Bidang Pemasaran dan Kemitraan, serta Bidang Pembinaan dan "
        "Keuangan Usaha. Pelaksanaan tugas harian didukung oleh pengurus wilayah, pendamping usaha, dan pelaku usaha anggota "
        "yang tersebar pada sentra kuliner di Kota Palembang."
    ),
}

SECTOR_BANKING["instansi_profile"] = {
    "sejarah": (
        "{instansi} merupakan bank swasta nasional yang melayani segmen ritel dan usaha kecil dengan mengandalkan kanal "
        "layanan digital sebagai tulang punggung transaksi nasabah. Bank ini beroperasi melalui jaringan kantor cabang, unit "
        "layanan syariah, serta aplikasi mobile banking yang menjangkau nasabah di berbagai wilayah. Transformasi digital yang "
        "dijalankan berfokus pada kecepatan transaksi, keandalan sistem, dan perlindungan data nasabah."
    ),
    "visi": (
        "Visi {instansi} adalah: \"Menjadi Bank Swasta Nasional Pilihan Utama Nasabah dengan Layanan Digital yang Cepat, "
        "Aman, dan Terpercaya\"."
    ),
    "misi": [
        "Menyediakan layanan perbankan digital yang efisien, andal, dan mudah diakses seluruh lapisan nasabah.",
        "Menjaga keamanan dan kerahasiaan data nasabah melalui pengendalian sistem informasi yang berlapis dan berkelanjutan.",
        "Membangun hubungan jangka panjang dengan nasabah melalui pelayanan yang transparan, responsif, dan berorientasi pada kepuasan.",
    ],
    "struktur": (
        "Struktur organisasi {instansi} dipimpin oleh Direktur Utama yang membawahi Direktorat Bisnis Ritel, Direktorat "
        "Teknologi dan Operasi Digital, Direktorat Risiko dan Kepatuhan, serta Direktorat Keuangan dan Sumber Daya Manusia. "
        "Pelaksanaan layanan harian didukung oleh Divisi Layanan Digital, Divisi Pusat Layanan Nasabah, Divisi Keamanan "
        "Informasi, serta jaringan kantor cabang di wilayah operasional."
    ),
}

SECTOR_HOSPITAL["instansi_profile"] = {
    "sejarah": (
        "{instansi} merupakan fasilitas pelayanan kesehatan swasta yang menyelenggarakan pelayanan rawat jalan, rawat inap, "
        "instalasi gawat darurat, serta layanan penunjang medis dengan tenaga kesehatan sebagai penggerak utama pelayanan. "
        "Rumah sakit ini terus mengembangkan sarana prasarana klinis dan sistem manajemen sumber daya manusia untuk menjaga "
        "kesinambungan mutu asuhan serta keselamatan pasien."
    ),
    "visi": (
        "Visi {instansi} adalah: \"Menjadi Rumah Sakit Swasta Pilihan Utama Masyarakat dengan Pelayanan Prima, Profesional, "
        "dan Berpusat pada Keselamatan Pasien\"."
    ),
    "misi": [
        "Menyelenggarakan pelayanan medis dan keperawatan yang bermutu, beretika, dan berstandar akreditasi rumah sakit.",
        "Mengembangkan kompetensi dan profesionalisme tenaga kesehatan secara berkelanjutan melalui pendidikan dan pelatihan.",
        "Menciptakan lingkungan kerja yang aman, nyaman, dan berkeadilan guna mempertahankan tenaga kesehatan berdedikasi.",
    ],
    "struktur": (
        "Struktur organisasi {instansi} dipimpin oleh Direktur Rumah Sakit yang dibantu Komite Medik, Komite Keperawatan, "
        "serta Satuan Pemeriksaan Internal. Direktur membawahi Bidang Pelayanan Medis, Bidang Keperawatan, Bidang Penunjang "
        "Medis, dan Bidang Administrasi, Keuangan, serta Sumber Daya Manusia yang mengoordinasikan instalasi rawat inap, "
        "rawat jalan, gawat darurat, dan unit penunjang klinis."
    ),
}

# ---------------------------------------------------------------------------
# 2. PEMBANGUN TABEL 1.2 KHUSUS SEKTOR
# ---------------------------------------------------------------------------

_MONTHS = ["Januari", "Februari", "Maret", "April", "Mei", "Juni"]


def _fmt_id(value: float, digits: int = 2) -> str:
    return f"{value:,.{digits}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def build_table_1_2(kind: str, rng, ctx: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Membangun payload Tabel 1.2 sesuai karakteristik sektor penelitian."""
    if kind == "umkm_omzet":
        headers = ["No", "Bulan", "Omzet (Rp Juta)", "Transaksi (Unit)",
                   "Rata-rata Nilai Transaksi (Rp Ribu)", "Pertumbuhan Omzet (%)"]
        rows, omzet_series, trx_series = [], [], []
        base = rng.uniform(118.0, 135.0)
        for idx, m in enumerate(_MONTHS):
            growth = rng.uniform(-6.5, 3.8)
            omzet = round(base * (1 + growth / 100.0) + rng.uniform(-4.5, 5.5), 2)
            if idx > 0:
                growth = round((omzet - omzet_series[-1]) / omzet_series[-1] * 100.0, 2)
            else:
                growth = 0.0
            trx = int(round(omzet * 1000 / rng.uniform(38.0, 52.0)))
            avg_val = round(omzet * 1_000_000 / max(trx, 1) / 1000.0, 1)
            omzet_series.append(omzet)
            trx_series.append(trx)
            rows.append([str(idx + 1), m, _fmt_id(omzet), f"{trx:,}".replace(",", "."),
                         _fmt_id(avg_val, 1), _fmt_id(growth)])
            base = omzet
        tot_omzet = round(sum(omzet_series), 2)
        tot_trx = sum(trx_series)
        min_idx = min(range(len(omzet_series)), key=lambda i: omzet_series[i])
        drop_month = _MONTHS[min_idx]
        drop_pct = round((omzet_series[min_idx] - omzet_series[min_idx - 1]) / omzet_series[min_idx - 1] * 100.0, 2) if min_idx > 0 else 0.0
        _sem_growth = round((omzet_series[-1] - omzet_series[0]) / omzet_series[0] * 100.0, 2)
        total_row = ["Total", "Rekapitulasi Semester", _fmt_id(tot_omzet),
                     f"{tot_trx:,}".replace(",", "."),
                     _fmt_id(tot_omzet * 1_000_000 / max(tot_trx, 1) / 1000.0, 1),
                     _fmt_id(_sem_growth)]
        concl = (
            f"Merujuk pada data Tabel 1.2 di atas, total omzet anggota usaha pada semester terakhir mencapai Rp {_fmt_id(tot_omzet)} juta "
            f"dengan volume transaksi sebanyak {tot_trx:,} unit. Penurunan omzet paling menonjol terjadi pada bulan {drop_month} "
            f"sebesar {_fmt_id(abs(drop_pct))}% dibandingkan bulan sebelumnya, sementara pertumbuhan omzet semesteran hanya tercatat "
            f"{_fmt_id(_sem_growth)}%. Pola fluktuasi ini mengindikasikan bahwa kapasitas pemasaran usaha belum stabil dalam mempertahankan "
            f"volume penjualan, sehingga penguatan orientasi kewirausahaan dan pembaruan produk menjadi kebutuhan mendesak bagi pelaku usaha."
        ).replace("{trx}", f"{tot_trx:,}")
        return {"headers": headers, "rows": rows, "total_row": total_row, "concl": concl}

    if kind == "bank_complaint":
        headers = ["No", "Bulan", "Gangguan Sistem (Kasus)", "Keluhan Transaksi (Kasus)",
                   "Keluhan Layanan (Kasus)", "Total Keluhan"]
        rows, totals = [], []
        for idx, m in enumerate(_MONTHS):
            g = rng.randint(9, 22)
            t = rng.randint(18, 36)
            l = rng.randint(11, 27)
            tot = g + t + l
            totals.append(tot)
            rows.append([str(idx + 1), m, str(g), str(t), str(l), str(tot)])
        grand = sum(totals)
        peak_idx = max(range(len(totals)), key=lambda i: totals[i])
        total_row = ["Total", "Rekapitulasi Semester",
                     str(sum(int(r[2]) for r in rows)), str(sum(int(r[3]) for r in rows)),
                     str(sum(int(r[4]) for r in rows)), str(grand)]
        concl = (
            f"Berdasarkan Tabel 1.2 di atas, akumulasi keluhan nasabah pada semester terakhir mencapai {grand} kasus yang terdiri atas "
            f"{total_row[2]} kasus gangguan sistem, {total_row[3]} kasus keluhan transaksi, dan {total_row[4]} kasus keluhan layanan. "
            f"Beban keluhan tertinggi terjadi pada bulan {_MONTHS[peak_idx]} sebanyak {totals[peak_idx]} kasus. Tingginya keluhan transaksi "
            f"dan layanan mencerminkan bahwa keandalan kanal digital serta kecepatan respons layanan bantuan belum sepenuhnya memenuhi "
            f"harapan nasabah, sehingga berpotensi menurunkan kepercayaan dan loyalitas nasabah terhadap bank."
        )
        return {"headers": headers, "rows": rows, "total_row": total_row, "concl": concl}

    if kind == "absensi_turnover":
        headers = ["No", "Bulan", "Izin (Hari)", "Sakit (Hari)", "Tanpa Keterangan (Hari)",
                   "Mengundurkan Diri (Orang)", "Total Ketidakhadiran"]
        rows, totals, resign_total = [], [], 0
        for idx, m in enumerate(_MONTHS):
            i = rng.randint(14, 26)
            s = rng.randint(10, 19)
            a = rng.randint(3, 9)
            r = rng.randint(1, 4)
            tot = i + s + a
            resign_total += r
            totals.append(tot)
            rows.append([str(idx + 1), m, str(i), str(s), str(a), str(r), str(tot)])
        grand = sum(totals)
        peak_idx = max(range(len(totals)), key=lambda i: totals[i])
        total_row = ["Total", "Rekapitulasi Semester",
                     str(sum(int(r[2]) for r in rows)), str(sum(int(r[3]) for r in rows)),
                     str(sum(int(r[4]) for r in rows)), str(resign_total), str(grand)]
        concl = (
            f"Merujuk pada Tabel 1.2 di atas, akumulasi ketidakhadiran karyawan mencapai {grand} hari kerja dalam satu semester dengan "
            f"puncak tertinggi pada bulan {_MONTHS[peak_idx]} sebanyak {totals[peak_idx]} hari. Pada periode yang sama tercatat "
            f"{resign_total} karyawan mengundurkan diri, sehingga rata-rata satu hingga empat karyawan meninggalkan organisasi setiap bulan. "
            f"Kombinasi tingginya izin, sakit, dan ketidakhadiran tanpa keterangan yang disertai pengunduran diri mengindikasikan bahwa "
            f"tekanan kerja, ketidakadilan imbalan, dan ketidaknyamanan lingkungan kerja fisik belum terkelola secara memadai oleh manajemen."
        )
        return {"headers": headers, "rows": rows, "total_row": total_row, "concl": concl}

    return None


# ---------------------------------------------------------------------------
# 3. API PRESET
# ---------------------------------------------------------------------------

BIBLIOGRAPHY_EXTRA = [
    {"author": "Barney, J. B., Ketchen, D. J., & Wright, M.", "year": "2023", "cite": "Barney et al. (2023)",
     "title": "The resource-based view of the firm: Foundations, extensions, and future horizons", "source": "Journal of Management, 49(1), 12–39. https://doi.org/10.1177/01492063221131777"},
    {"author": "Kementerian Koperasi dan Usaha Kecil dan Menengah Republik Indonesia", "year": "2023",
     "cite": "Kementerian Koperasi dan UKM (2023)",
     "title": "Laporan kinerja pembinaan usaha mikro, kecil, dan menengah tahun 2023", "source": "Kementerian Koperasi dan UKM Republik Indonesia."},
    {"author": "Otoritas Jasa Keuangan", "year": "2023", "cite": "Otoritas Jasa Keuangan (2023)",
     "title": "Statistik perbankan Indonesia dan laporan pengaduan konsumen sektor jasa keuangan 2023", "source": "Otoritas Jasa Keuangan Republik Indonesia."},
    {"author": "Kementerian Kesehatan Republik Indonesia", "year": "2023", "cite": "Kementerian Kesehatan (2023)",
     "title": "Profil kesehatan Indonesia 2023 dan laporan ketenagakerjaan fasilitas pelayanan kesehatan", "source": "Kementerian Kesehatan Republik Indonesia."},
]


def iter_constructs():
    for mod in (MKT, BANK, HR):
        for key, prof in mod.CONSTRUCTS.items():
            yield key, prof


def iter_bab2_definitions():
    for mod in (MKT, BANK, HR):
        for key, data in mod.BAB2_DEFINITIONS.items():
            yield key, data


def iter_bibliography():
    seen = set()
    for entry in BIBLIOGRAPHY_EXTRA:
        if entry["cite"] not in seen:
            seen.add(entry["cite"])
            yield entry
    for mod in (MKT, BANK, HR):
        for entry in mod.BIBLIOGRAPHY:
            if entry["cite"] in seen:
                continue
            seen.add(entry["cite"])
            yield entry


def iter_table_2_1():
    for mod in (MKT, BANK, HR):
        yield mod.__name__, mod.TABLE_2_1


TABLE_2_1_BY_SECTOR = {
    "umkm_marketing": MKT.TABLE_2_1,
    "bank_digital": BANK.TABLE_2_1,
    "rs_turnover": HR.TABLE_2_1,
}


def get_table_2_1(ctx: Dict[str, Any]) -> Optional[list]:
    """Mengembalikan 10 baris Tabel 2.1 sesuai preset sektor (None bila tidak ada)."""
    preset = get_topic_preset(ctx)
    if not preset:
        return None
    return TABLE_2_1_BY_SECTOR.get(preset.get("id"))


def get_topic_preset(ctx: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Mendeteksi preset sektoral yang paling sesuai dengan judul/instansi penelitian."""
    corpus = f"{ctx.get('instansi', '')} {ctx.get('title_upper', '')} {ctx.get('y_name', '')}".upper()
    best, best_score = None, 0
    for sec in SECTORS:
        score = sum(1 for kw in sec["keywords"] if kw in corpus)
        if score > best_score:
            best, best_score = sec, score
    if best_score >= 2:
        return best
    return None


def get_construct(name: str) -> Optional[Dict[str, Any]]:
    key = " ".join((name or "").lower().split()).strip(" .,;:")
    if not key:
        return None
    for mod in (MKT, BANK, HR):
        if key in mod.CONSTRUCTS:
            return mod.CONSTRUCTS[key]
    return None


def apply_lexicon(text: str, preset: Optional[Dict[str, Any]]) -> str:
    """Menyesuaikan istilah generik naskah dengan konteks sektor penelitian (whole-word)."""
    if not text or not preset:
        return text
    import re
    lex = preset.get("lexicon") or {}
    out = text
    for src in sorted(lex, key=len, reverse=True):
        dst = lex[src]
        out = re.sub(r"(?<![A-Za-z])" + re.escape(src) + r"(?![A-Za-z])", dst, out)
    return out
