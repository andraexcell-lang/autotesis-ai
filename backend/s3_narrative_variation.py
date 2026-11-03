# -*- coding: utf-8 -*-
"""
Mesin variasi narasi Disertasi S3 (P1 de-templating).

Masalah yang diselesaikan
-------------------------
Sebelum modul ini ada, hampir seluruh narasi Bab I--III dibentuk dari satu
template `f-string`. Ketika dua peneliti memakai mesin yang sama dengan judul
berbeda, naskah hanya berubah pada nama instansi dan nama variabel, sementara
susunan kalimatnya identik. Modul ini menyediakan tiga lapisan pemutus template:

1. ``VAR_KB`` -- basis pengetahuan semantik per konstruk (definisi, teori,
   mekanisme kausal, risiko, kondisi ideal, tuas intervensi, dimensi/indikator).
   Narasi kemudian *dirakit dari makna konstruk*, bukan dari satu kalimat baku.
2. ``profile_map`` -- menggabungkan basis pengetahuan dengan ontologi hasil
   riset judul. Ontologi yang generik (mis. "Dimensi Kapasitas X") diganti
   dimensi substantif yang benar-benar mengukur konstruk.
3. ``polish`` -- parafrase leksikal deterministik (bank sinonim akademik) yang
   di-seed dari judul + salt paragraf. Hasilnya stabil untuk judul yang sama,
   tetapi berbeda antar-judul, dan tidak menyentuh sitasi dalam tanda kurung.

Semua fungsi bersifat murni/deterministik: tidak ada `random` global, tidak ada
ketergantungan pada jam sistem, sehingga hasil build dapat direproduksi.
"""

from __future__ import annotations

import hashlib
import random
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

# ==============================================================================
# 1. UTILITAS DETERMINISTIK
# ==============================================================================


def _norm(text: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(text or "").lower()).strip()


def _seed(*parts: Any) -> int:
    raw = "||".join(str(p) for p in parts).encode("utf-8", "ignore")
    return int(hashlib.blake2b(raw, digest_size=8).hexdigest(), 16)


def rng_for(ctx: Dict[str, Any], salt: str = "") -> random.Random:
    """RNG deterministik per (judul, instansi, salt)."""
    return random.Random(
        _seed(
            ctx.get("judul") or ctx.get("title") or "",
            ctx.get("instansi") or "",
            "|".join(ctx.get("x_vars") or []),
            "|".join(ctx.get("m_vars") or []),
            "|".join(ctx.get("z_vars") or []),
            ctx.get("y_var") or "",
            salt,
        )
    )


def pick(rng: random.Random, options: Sequence[str]) -> str:
    return options[rng.randrange(len(options))]


def pick_idx(rng: random.Random, n: int) -> int:
    return rng.randrange(max(1, n))


# ==============================================================================
# 2. BASIS PENGETAHUAN SEMANTIK VARIABEL
# ==============================================================================
# Setiap entri memuat makna substantif konstruk. Field yang dipakai perakit
# narasi: family, definisi, teori, mekanisme, risiko, tinggi, intervensi,
# dimensi, sitasi. Teori ditulis (nama teori, sitasi, klaim).
# Seluruh sitasi di bawah ini sudah ada pada registry kanonik
# (CANONICAL_TEXTBOOK_REFERENCES / CANONICAL_FOUNDATIONAL_REFERENCES).

VAR_KB: List[Dict[str, Any]] = [
    {
        "key": "kompetensi",
        "aliases": ["competency", "competence", "kompetensi kerja", "kapabilitas"],
        "label": "Kompetensi",
        "family": "kapasitas",
        "definisi": (
            "kesatuan pengetahuan prosedural, keterampilan teknis, dan sikap kerja "
            "yang memungkinkan aparatur menerjemahkan tuntutan jabatan menjadi "
            "tindakan kerja yang tepat, mandiri, dan sesuai standar profesi"
        ),
        "teori": [
            ("Human Capital Theory", "Becker (1964)",
             "menempatkan pengetahuan dan keterampilan sebagai bentuk investasi produktif yang nilainya terakumulasi melalui pendidikan, pelatihan, dan pengalaman kerja"),
            ("Resource-Based View", "Daft (2022)",
             "menegaskan bahwa kapabilitas internal yang bernilai, langka, dan sulit ditiru merupakan sumber keunggulan kinerja yang berkelanjutan"),
            ("Teori Kompetensi Spencer", "Spencer & Spencer (1993)",
             "membedakan kompetensi menjadi pengetahuan, keterampilan, konsep diri, sifat, dan motif yang bersama-sama memprediksi kinerja superior"),
        ],
        "mekanisme": (
            "membekali aparatur dengan peta kognitif dan prosedur kerja yang tepat sehingga "
            "keputusan teknis diambil lebih cepat, kesalahan prosedural menurun, dan ketergantungan "
            "pada instruksi atasan berkurang"
        ),
        "risiko": (
            "kesalahan prosedur, pekerjaan yang harus diulang, serta ketergantungan berlebihan "
            "pada supervisi atasan langsung"
        ),
        "tinggi": (
            "kemandirian kerja meningkat, mutu luaran lebih konsisten, dan aparatur mampu "
            "menangani persoalan teknis yang tidak rutin"
        ),
        "intervensi": (
            "analisis kebutuhan pelatihan berbasis jabatan, rotasi penugasan, coaching atasan "
            "langsung, dan sertifikasi kompetensi"
        ),
        "dimensi": [
            {"nama": "Pengetahuan Kerja (Know-What)", "indikator": [
                "Pemahaman terhadap regulasi dan prosedur kerja yang berlaku",
                "Penguasaan substansi teknis bidang tugas",
                "Kemampuan mengaitkan pengetahuan dengan persoalan kerja nyata"]},
            {"nama": "Keterampilan Teknis (Skill)", "indikator": [
                "Ketepatan menerapkan prosedur operasional standar",
                "Kecepatan dan ketelitian menyelesaikan pekerjaan teknis",
                "Kemampuan menggunakan sarana kerja dan sistem informasi kedinasan"]},
            {"nama": "Sikap dan Perilaku Profesional (Attitude)", "indikator": [
                "Kesiapan menerima tanggung jawab pekerjaan",
                "Konsistensi menjaga mutu hasil kerja",
                "Keterbukaan terhadap koreksi dan pengembangan diri"]},
        ],
        "sitasi": "Spencer & Spencer (1993); Kasmir (2023); Afandi (2023)",
    },
    {
        "key": "beban kerja",
        "aliases": ["workload", "beban tugas", "job demand", "tuntutan kerja"],
        "label": "Beban Kerja",
        "family": "tuntutan",
        "definisi": (
            "akumulasi tuntutan tugas yang menuntut pengerahan waktu, kapasitas mental, "
            "dan energi fisik aparatur dalam batas tenggat serta standar mutu tertentu"
        ),
        "teori": [
            ("Job Demands-Resources Theory", "Bakker & Demerouti (2022)",
             "memandang tuntutan kerja sebagai tekanan yang dapat menguras energi, tetapi juga dapat merangsang pertumbuhan apabila diimbangi sumber daya kerja yang memadai"),
            ("NASA Task Load Index", "Hart & Staveland (1988)",
             "memetakan beban kerja ke dalam dimensi tuntutan mental, tuntutan fisik, tuntutan waktu, usaha, frustrasi, dan performa"),
            ("Psikologi Industri dan Organisasi", "Munandar (2021)",
             "menjelaskan bahwa beban kerja berlebih menurunkan ketelitian dan memicu kelelahan, sedangkan beban terlalu rendah menimbulkan kebosanan"),
        ],
        "mekanisme": (
            "menyerap kapasitas kognitif, waktu, dan energi fisik aparatur sehingga kualitas "
            "pemrosesan informasi dan ketahanan kerja ditentukan oleh keseimbangan antara "
            "tuntutan tugas dan sumber daya yang tersedia"
        ),
        "risiko": (
            "kelelahan kerja, penurunan ketelitian, penundaan penyelesaian tugas, dan "
            "kecenderungan melewatkan tahapan prosedur ketika tenggat mendesak"
        ),
        "tinggi": (
            "pada takaran yang terkelola, tuntutan kerja menstimulasi kewaspadaan, "
            "prioritisasi tugas, dan pengerahan usaha yang lebih terarah"
        ),
        "intervensi": (
            "penataan ulang beban tugas, kejelasan prioritas, penambahan sumber daya kerja, "
            "serta pemantauan indikator kelelahan secara berkala"
        ),
        "dimensi": [
            {"nama": "Beban Waktu (Temporal Demand)", "indikator": [
                "Kecukupan waktu menyelesaikan tugas sesuai tenggat",
                "Frekuensi pekerjaan yang harus diselesaikan secara bersamaan",
                "Tekanan waktu pada periode pelaporan atau puncak aktivitas"]},
            {"nama": "Beban Mental (Mental Demand)", "indikator": [
                "Kompleksitas analisis yang dituntut oleh pekerjaan",
                "Kebutuhan konsentrasi dan ketelitian tingkat tinggi",
                "Frekuensi pengambilan keputusan teknis yang berdampak luas"]},
            {"nama": "Beban Fisik dan Frustrasi (Effort & Frustration)", "indikator": [
                "Intensitas aktivitas fisik dan mobilitas kerja",
                "Tingkat kelelahan yang dirasakan setelah jam dinas",
                "Gangguan emosional akibat hambatan prosedural dan tekanan target"]},
        ],
        "sitasi": "Bakker & Demerouti (2022); Hart & Staveland (1988); Munandar (2021)",
    },
    {
        "key": "keterikatan kerja",
        "aliases": ["employee engagement", "work engagement", "keterlibatan kerja", "engagement"],
        "label": "Keterikatan Kerja",
        "family": "energi",
        "definisi": (
            "kondisi psikologis positif yang ditandai oleh vigor (energi dan ketangguhan), "
            "dedication (kebanggaan dan antusiasme), serta absorption (konsentrasi mendalam) "
            "dalam menjalankan tugas"
        ),
        "teori": [
            ("Job Demands-Resources Theory", "Bakker & Demerouti (2022)",
             "menjelaskan bahwa sumber daya kerja memicu keterikatan yang kemudian menyalurkan energi psikologis menjadi perilaku kerja produktif"),
            ("Teori Keterikatan Kerja", "Saks (2022)",
             "menegaskan bahwa keterikatan merupakan anteseden penting bagi perilaku kewargaan organisasi, kepuasan kerja, dan kinerja"),
            ("Keterikatan Kerja sebagai Energi Psikologis", "Schaufeli (2022)",
             "memosisikan vigor, dedication, dan absorption sebagai inti pengalaman kerja yang menggerakkan persistensi usaha"),
        ],
        "mekanisme": (
            "menyalurkan kapasitas dan tuntutan kerja menjadi energi psikologis yang membuat "
            "aparatur mencurahkan usaha diskresioner, bertahan menghadapi hambatan, dan "
            "menjaga mutu interaksi kerja"
        ),
        "risiko": (
            "keterlibatan formalistik, kelelahan emosional, dan penurunan inisiatif di luar "
            "kewajiban minimal"
        ),
        "tinggi": (
            "aparatur menunjukkan persistensi usaha, kebanggaan profesi, dan konsentrasi penuh "
            "yang mempercepat penyelesaian pekerjaan bermutu"
        ),
        "intervensi": (
            "penguatan makna tugas, umpan balik kinerja, dukungan atasan, dan penyediaan "
            "sumber daya kerja yang memadai"
        ),
        "dimensi": [
            {"nama": "Vigor (Energi dan Ketangguhan)", "indikator": [
                "Tingkat energi dan ketahanan mental saat bekerja",
                "Kegigihan berusaha ketika menghadapi hambatan tugas",
                "Kesediaan menambah usaha untuk menyelesaikan pekerjaan mendesak"]},
            {"nama": "Dedication (Kebanggaan dan Antusiasme)", "indikator": [
                "Rasa bangga terhadap institusi dan pekerjaan",
                "Antusiasme menjalankan tugas kedinasan",
                "Kesediaan menerima pekerjaan sebagai bagian dari identitas profesi"]},
            {"nama": "Absorption (Konsentrasi dan Penyerapan)", "indikator": [
                "Konsentrasi penuh saat jam dinas",
                "Keterlarutan positif dalam menyelesaikan pekerjaan",
                "Kesulitan melepaskan pikiran dari tugas yang belum tuntas"]},
        ],
        "sitasi": "Schaufeli (2022); Saks (2022); Bakker & Demerouti (2022)",
    },
    {
        "key": "modal psikologis",
        "aliases": ["psychological capital", "psycap", "modal psikologi"],
        "label": "Modal Psikologis",
        "family": "modal",
        "definisi": (
            "kapasitas psikologis positif yang terdiri atas self-efficacy, hope, resilience, "
            "dan optimism yang dapat dikembangkan untuk menghadapi tuntutan kerja"
        ),
        "teori": [
            ("Psychological Capital Theory", "Luthans et al. (2022)",
             "merumuskan modal psikologis sebagai konstruk tingkat individu yang bersifat state-like, terukur, dan dapat dikembangkan melalui intervensi"),
            ("Meta-Analisis Modal Psikologis", "Avey et al. (2011)",
             "menunjukkan bahwa modal psikologis berkorelasi positif dengan sikap kerja, perilaku kewargaan, dan kinerja"),
            ("Positive Psychology", "Seligman (2002)",
             "menekankan kekuatan psikologis positif sebagai sumber daya yang melindungi individu dari tekanan dan memperluas repertoar tindakan"),
        ],
        "mekanisme": (
            "menyediakan cadangan psikologis berupa keyakinan kemampuan diri, harapan akan "
            "kemungkinan perbaikan, daya lenting saat gagal, dan optimisme yang realists"
        ),
        "risiko": (
            "cepat menyerah saat menghadapi hambatan, pemulihan kerja yang lambat setelah "
            "tekanan tinggi, dan interpretasi pesimistis terhadap perubahan kebijakan"
        ),
        "tinggi": (
            "aparatur mampu mempertahankan fokus, bangkit dari kegagalan, dan menafsirkan "
            "tekanan sebagai tantangan yang dapat dikelola"
        ),
        "intervensi": (
            "pelatihan pengembangan modal psikologis, pembinaan harapan berbasis target, "
            "dukungan sosial atasan, dan refleksi pengalaman kerja"
        ),
        "dimensi": [
            {"nama": "Self-Efficacy (Keyakinan Kemampuan Diri)", "indikator": [
                "Keyakinan menyelesaikan tugas sulit",
                "Kepercayaan pada kemampuan menganalisis persoalan kerja",
                "Kesediaan mengambil tanggung jawab baru"]},
            {"nama": "Hope (Harapan dan Perencanaan Jalur)", "indikator": [
                "Kemampuan merumuskan alternatif solusi",
                "Orientasi pada pencapaian sasaran kerja",
                "Ketekunan mencari jalan keluar saat hambatan muncul"]},
            {"nama": "Resilience (Daya Lenting)", "indikator": [
                "Kemampuan pulih setelah mengalami tekanan kerja",
                "Ketahanan menghadapi kritik dan kegagalan",
                "Kemampuan menyesuaikan diri terhadap perubahan tugas"]},
            {"nama": "Optimism (Optimisme Realistis)", "indikator": [
                "Harapan positif terhadap hasil kerja",
                "Cara menafsirkan peristiwa negatif secara proporsional",
                "Keyakinan bahwa usaha yang tepat menghasilkan perbaikan"]},
        ],
        "sitasi": "Luthans et al. (2022); Avey et al. (2011); Seligman (2002)",
    },
    {
        "key": "kinerja pegawai",
        "aliases": ["kinerja", "performance", "employee performance", "job performance", "kinerja aparatur"],
        "label": "Kinerja Pegawai",
        "family": "luaran",
        "definisi": (
            "hasil kerja yang mencakup mutu, kuantitas, ketepatan waktu, dan efektivitas "
            "perilaku kerja aparatur dalam mencapai sasaran organisasi"
        ),
        "teori": [
            ("Goal-Setting Theory", "Wibowo (2023)",
             "menjelaskan bahwa sasaran yang spesifik dan menantang, disertai umpan balik, mengarahkan usaha dan meningkatkan kinerja"),
            ("Teori Kinerja", "Campbell & Wiernik (2022)",
             "memosisikan kinerja sebagai perilaku kerja yang dapat diamati dan dinilai, bukan semata hasil akhir"),
            ("Manajemen Kinerja", "Kasmir (2023)",
             "menegaskan bahwa kinerja harus dinilai melalui indikator yang terukur, adil, dan terkait langsung dengan tanggung jawab jabatan"),
        ],
        "mekanisme": (
            "mengonsolidasikan kapasitas, energi, dan dukungan kerja menjadi luaran yang "
            "bermutu, tepat waktu, dan selaras dengan sasaran organisasi"
        ),
        "risiko": (
            "capaian target tidak merata, mutu luaran menurun pada periode sibuk, dan "
            "ketergantungan pada perbaikan ulang"
        ),
        "tinggi": (
            "sasaran kerja tercapai secara konsisten, mutu luaran terjaga, dan aparatur "
            "menunjukkan inisiatif perbaikan"
        ),
        "intervensi": (
            "penetapan indikator kinerja yang jelas, umpan balik berkala, penghargaan berbasis "
            "capaian, dan penguatan kapasitas teknis"
        ),
        "dimensi": [
            {"nama": "Kualitas Pekerjaan", "indikator": [
                "Ketelitian dan ketepatan hasil kerja",
                "Kesesuaian luaran dengan standar operasional prosedur",
                "Kerapian dokumentasi dan pelaporan"]},
            {"nama": "Kuantitas Output", "indikator": [
                "Ketercapaian target volume pekerjaan",
                "Kemampuan menuntaskan tugas tambahan",
                "Konsistensi produktivitas antarperiode"]},
            {"nama": "Ketepatan Waktu", "indikator": [
                "Penyelesaian tugas sesuai tenggat",
                "Daya tanggap terhadap disposisi pimpinan",
                "Disiplin pemanfaatan jam kerja"]},
            {"nama": "Efektivitas dan Kerja Sama", "indikator": [
                "Efisiensi penggunaan sumber daya kantor",
                "Inisiatif memecahkan hambatan kerja",
                "Kesediaan berkoordinasi lintas unit"]},
        ],
        "sitasi": "Campbell & Wiernik (2022); Kasmir (2023); Wibowo (2023)",
    },
    {
        "key": "kepemimpinan transformasional",
        "aliases": ["transformational leadership", "kepemimpinan transformasi"],
        "label": "Kepemimpinan Transformasional",
        "family": "kepemimpinan",
        "definisi": (
            "gaya kepemimpinan yang menggerakkan bawahan melalui keteladanan, visi inspiratif, "
            "stimulasi intelektual, dan perhatian individual"
        ),
        "teori": [
            ("Full Range Leadership Theory", "Bass (1985)",
             "menjelaskan bahwa kepemimpinan transformasional mengubah nilai, keyakinan, dan aspirasi bawahan sehingga melampaui pertukaran imbalan"),
            ("Teori Kepemimpinan", "Yukl & Gardner (2024)",
             "menegaskan peran keteladanan dan komunikasi visi dalam membentuk komitmen serta kapasitas kolektif organisasi"),
        ],
        "mekanisme": (
            "membentuk identifikasi bawahan terhadap visi dan keteladanan pimpinan sehingga "
            "usaha kerja melampaui kewajiban formal"
        ),
        "risiko": (
            "keterikatan yang bergantung pada sosok pimpinan, retorika tanpa dukungan sumber daya, "
            "dan melemahnya inisiatif ketika keteladanan tidak konsisten"
        ),
        "tinggi": (
            "bawahan menunjukkan komitmen afektif, keberanian berinovasi, dan kesediaan "
            "berkorban untuk sasaran organisasi"
        ),
        "intervensi": (
            "penguatan keteladanan pimpinan, komunikasi visi, pendelegasian yang memberdayakan, "
            "dan pembinaan individual"
        ),
        "dimensi": [
            {"nama": "Idealized Influence", "indikator": [
                "Keteladanan perilaku pimpinan",
                "Konsistensi pimpinan menepati komitmen"]},
            {"nama": "Inspirational Motivation", "indikator": [
                "Kejelasan visi dan sasaran",
                "Kemampuan membangkitkan antusiasme kerja"]},
            {"nama": "Intellectual Stimulation", "indikator": [
                "Dorongan berpikir kritis dan inovatif",
                "Keterbukaan terhadap gagasan bawahan"]},
            {"nama": "Individualized Consideration", "indikator": [
                "Perhatian pada kebutuhan pengembangan bawahan",
                "Dukungan penyelesaian masalah kerja"]},
        ],
        "sitasi": "Bass (1985); Yukl & Gardner (2024); Northouse (2024)",
    },
    {
        "key": "budaya organisasi",
        "aliases": ["organizational culture", "budaya kerja", "kultur organisasi"],
        "label": "Budaya Organisasi",
        "family": "nilai",
        "definisi": (
            "sistem nilai, asumsi dasar, dan norma bersama yang membentuk cara aparatur "
            "menafsirkan pekerjaan dan berperilaku di lingkungan organisasi"
        ),
        "teori": [
            ("Teori Budaya Organisasi", "Schein & Schein (2023)",
             "membedakan artefak, nilai yang dianut, dan asumsi dasar sebagai lapisan budaya yang membimbing perilaku"),
            ("Competing Values Framework", "Cameron & Quinn (2022)",
             "memetakan budaya ke dalam orientasi clan, adhocracy, market, dan hierarchy yang memengaruhi efektivitas organisasi"),
        ],
        "mekanisme": (
            "menyediakan kerangka normatif bersama yang membantu aparatur menafsirkan "
            "prioritas, menilai perilaku yang pantas, dan menyelaraskan tindakan"
        ),
        "risiko": (
            "inkonsistensi antara nilai yang diumumkan dan praktik sehari-hari, serta "
            "toleransi terhadap penyimpangan prosedur"
        ),
        "tinggi": (
            "aparatur memiliki pegangan nilai yang jelas, kohesi sosial kuat, dan orientasi "
            "pada mutu serta akuntabilitas"
        ),
        "intervensi": (
            "penyelarasan nilai dengan sistem penghargaan, keteladanan pimpinan, dan "
            "ritual organisasi yang memperkuat norma positif"
        ),
        "dimensi": [
            {"nama": "Nilai dan Norma Bersama", "indikator": [
                "Kejelasan nilai yang dijunjung organisasi",
                "Kesesuaian perilaku sehari-hari dengan nilai tersebut"]},
            {"nama": "Asumsi Dasar dan Keyakinan", "indikator": [
                "Keyakinan tentang cara kerja yang benar",
                "Orientasi terhadap perubahan dan perbaikan"]},
            {"nama": "Artefak dan Praktik", "indikator": [
                "Simbol, ritual, dan kebiasaan kerja",
                "Konsistensi pimpinan dalam mempraktikkan nilai"]},
        ],
        "sitasi": "Schein & Schein (2023); Cameron & Quinn (2022); Sobirin (2023)",
    },
    {
        "key": "iklim kerja",
        "aliases": ["organizational climate", "iklim organisasi", "iklim kerja organisasi"],
        "label": "Iklim Kerja",
        "family": "iklim",
        "definisi": (
            "persepsi bersama aparatur mengenai kualitas lingkungan kerja, dukungan atasan, "
            "kejelasan peran, dan keadilan perlakuan dalam organisasi"
        ),
        "teori": [
            ("Teori Iklim Organisasi", "Litwin & Stringer (1968)",
             "menjelaskan bahwa iklim kerja memediasi pengaruh struktur organisasi terhadap motivasi dan perilaku anggota"),
            ("Budaya dan Iklim Organisasi", "Wirawan (2022)",
             "memosisikan iklim sebagai persepsi kolektif yang relatif stabil terhadap praktik organisasi"),
        ],
        "mekanisme": (
            "membentuk persepsi tentang apa yang dihargai dan diharapkan organisasi sehingga "
            "mengarahkan intensitas usaha serta kerja sama aparatur"
        ),
        "risiko": (
            "persepsi ketidakadilan, rendahnya dukungan atasan, dan iklim defensif yang "
            "menghambat pelaporan masalah"
        ),
        "tinggi": (
            "aparatur merasa aman menyampaikan gagasan, mendapat dukungan, dan memahami "
            "standar yang diharapkan"
        ),
        "intervensi": (
            "perbaikan kejelasan peran, keadilan prosedur, dukungan atasan, dan kanal "
            "komunikasi dua arah"
        ),
        "dimensi": [
            {"nama": "Struktur dan Kejelasan Peran", "indikator": [
                "Kejelasan pembagian tugas dan wewenang",
                "Ketersediaan prosedur kerja yang dipahami"]},
            {"nama": "Dukungan dan Kehangatan", "indikator": [
                "Dukungan atasan terhadap penyelesaian tugas",
                "Hubungan kerja yang saling membantu"]},
            {"nama": "Keadilan dan Penghargaan", "indikator": [
                "Keadilan penilaian kinerja",
                "Pengakuan atas kontribusi kerja"]},
        ],
        "sitasi": "Litwin & Stringer (1968); Wirawan (2022)",
    },
    {
        "key": "disiplin kerja",
        "aliases": ["work discipline", "kedisiplinan", "disiplin pegawai"],
        "label": "Disiplin Kerja",
        "family": "kepatuhan",
        "definisi": (
            "ketaatan aparatur terhadap aturan, prosedur, jadwal, dan tanggung jawab jabatan "
            "yang ditunjukkan secara konsisten dalam perilaku kerja"
        ),
        "teori": [
            ("Manajemen Sumber Daya Manusia", "Sutrisno (2022)",
             "menegaskan disiplin sebagai sikap menghormati dan menaati peraturan serta norma yang berlaku"),
            ("Manajemen SDM", "Hasibuan (2022)",
             "mengaitkan disiplin dengan kesadaran dan kesediaan mematuhi aturan serta tanggung jawab kerja"),
        ],
        "mekanisme": (
            "mengurangi variasi perilaku yang menyimpang dari standar sehingga pekerjaan "
            "berjalan tertib, dapat diprediksi, dan tepat waktu"
        ),
        "risiko": (
            "keterlambatan, ketidakhadiran, pelanggaran prosedur, dan penurunan ketertiban administrasi"
        ),
        "tinggi": (
            "ketertiban kerja terjaga, tenggat dipenuhi, dan standar prosedur dijalankan konsisten"
        ),
        "intervensi": (
            "penegakan aturan yang konsisten, keteladanan atasan, pembinaan, dan sistem "
            "penghargaan-hukuman yang adil"
        ),
        "dimensi": [
            {"nama": "Ketaatan pada Aturan", "indikator": [
                "Kepatuhan terhadap prosedur kerja",
                "Kesediaan menerima konsekuensi pelanggaran"]},
            {"nama": "Kedisiplinan Waktu", "indikator": [
                "Ketepatan waktu hadir dan menyelesaikan tugas",
                "Pemanfaatan jam kerja secara produktif"]},
            {"nama": "Tanggung Jawab Kerja", "indikator": [
                "Penyelesaian tugas tanpa pengawasan berlebih",
                "Kesediaan memperbaiki kesalahan kerja"]},
        ],
        "sitasi": "Sutrisno (2022); Hasibuan (2022)",
    },
    {
        "key": "motivasi kerja",
        "aliases": ["motivation", "motivasi", "motivasi pegawai"],
        "label": "Motivasi Kerja",
        "family": "dorongan",
        "definisi": (
            "daya dorong internal dan eksternal yang mengarahkan intensitas, arah, dan "
            "ketekunan usaha aparatur dalam mencapai sasaran kerja"
        ),
        "teori": [
            ("Expectancy Theory", "Vroom (1964)",
             "menjelaskan motivasi sebagai hasil perkalian harapan keberhasilan, instrumentalitas, dan valensi imbalan"),
            ("Need Theory", "McClelland (1987)",
             "menekankan kebutuhan berprestasi, berafiliasi, dan berkuasa sebagai pendorong perilaku kerja"),
            ("Teori Motivasi", "Uno (2023)",
             "merangkum motivasi sebagai proses internal yang menggerakkan dan mengarahkan tindakan"),
        ],
        "mekanisme": (
            "mengarahkan energi psikologis aparatur pada sasaran yang dianggap bernilai serta "
            "mempertahankan usaha ketika hambatan muncul"
        ),
        "risiko": (
            "usaha minimal, orientasi pada formalitas, dan penurunan ketekunan ketika imbalan "
            "tidak dirasakan adil"
        ),
        "tinggi": (
            "aparatur menunjukkan inisiatif, ketekunan, dan kesediaan bekerja melampaui tuntutan minimal"
        ),
        "intervensi": (
            "penyelarasan imbalan dengan kinerja, kejelasan jalur karier, umpan balik, dan "
            "penguatan makna pekerjaan"
        ),
        "dimensi": [
            {"nama": "Motivasi Intrinsik", "indikator": [
                "Kepuasan dari pekerjaan itu sendiri",
                "Dorongan mengembangkan kemampuan"]},
            {"nama": "Motivasi Ekstrinsik", "indikator": [
                "Ketertarikan pada imbalan dan pengakuan",
                "Kejelasan hubungan kinerja dengan konsekuensi"]},
            {"nama": "Ketekunan Usaha", "indikator": [
                "Kesediaan bekerja melebihi tuntutan minimal",
                "Ketahanan menghadapi hambatan kerja"]},
        ],
        "sitasi": "Vroom (1964); McClelland (1987); Uno (2023)",
    },
    {
        "key": "kepuasan kerja",
        "aliases": ["job satisfaction", "kepuasan pegawai"],
        "label": "Kepuasan Kerja",
        "family": "evaluasi",
        "definisi": (
            "evaluasi afektif dan kognitif aparatur terhadap pekerjaan, imbalan, supervisi, "
            "rekan kerja, dan peluang pengembangan"
        ),
        "teori": [
            ("Perilaku Organisasi", "Robbins & Judge (2022)",
             "menjelaskan kepuasan kerja sebagai sikap positif terhadap pekerjaan yang terbentuk dari evaluasi karakteristik pekerjaan"),
            ("Organizational Behavior", "Luthans et al. (2022)",
             "mengaitkan kepuasan kerja dengan komitmen, keterlibatan, dan perilaku kewargaan organisasi"),
        ],
        "mekanisme": (
            "membentuk disposisi positif terhadap pekerjaan yang meningkatkan keterlibatan "
            "dan menurunkan niat meninggalkan organisasi"
        ),
        "risiko": (
            "keluhan kerja, penurunan komitmen, dan peningkatan niat berpindah"
        ),
        "tinggi": (
            "aparatur merasa pekerjaan bermakna, diperlakukan adil, dan memiliki peluang berkembang"
        ),
        "intervensi": (
            "perbaikan keadilan imbalan, kualitas supervisi, desain pekerjaan, dan jalur pengembangan"
        ),
        "dimensi": [
            {"nama": "Kepuasan terhadap Pekerjaan", "indikator": [
                "Makna pekerjaan yang dirasakan",
                "Kesesuaian tugas dengan kemampuan"]},
            {"nama": "Kepuasan terhadap Imbalan dan Karier", "indikator": [
                "Keadilan imbalan",
                "Kejelasan peluang pengembangan"]},
            {"nama": "Kepuasan terhadap Lingkungan Sosial", "indikator": [
                "Kualitas hubungan dengan atasan",
                "Kualitas hubungan dengan rekan kerja"]},
        ],
        "sitasi": "Robbins & Judge (2022); Luthans et al. (2022)",
    },
    {
        "key": "kompensasi",
        "aliases": ["compensation", "imbalan", "remunerasi"],
        "label": "Kompensasi",
        "family": "imbalan",
        "definisi": (
            "keseluruhan imbalan finansial dan nonfinansial yang diterima aparatur sebagai "
            "pertukaran atas kontribusi kerja"
        ),
        "teori": [
            ("Manajemen Sumber Daya Manusia", "Hasibuan (2022)",
             "menjelaskan kompensasi sebagai semua pendapatan berbentuk uang, barang langsung atau tidak langsung yang diterima karyawan"),
            ("Human Resource Management", "Dessler (2023)",
             "menekankan keadilan internal dan daya saing eksternal kompensasi sebagai penentu kepuasan dan retensi"),
        ],
        "mekanisme": (
            "memenuhi harapan timbal balik aparatur sehingga memperkuat keadilan yang "
            "dirasakan dan kesediaan mencurahkan usaha"
        ),
        "risiko": (
            "persepsi ketidakadilan imbalan, penurunan motivasi, dan niat meninggalkan organisasi"
        ),
        "tinggi": (
            "aparatur mempersepsikan imbalan adil, kompetitif, dan terkait dengan kontribusi"
        ),
        "intervensi": (
            "evaluasi jabatan, penyesuaian struktur imbalan, transparansi tunjangan kinerja, "
            "dan penghargaan nonfinansial"
        ),
        "dimensi": [
            {"nama": "Kompensasi Finansial Langsung", "indikator": [
                "Kesesuaian gaji dengan beban dan tanggung jawab",
                "Keadilan tunjangan kinerja"]},
            {"nama": "Kompensasi Tidak Langsung", "indikator": [
                "Ketersediaan jaminan sosial dan kesehatan",
                "Fasilitas pendukung kesejahteraan"]},
            {"nama": "Kompensasi Nonfinansial", "indikator": [
                "Pengakuan atas prestasi",
                "Kesempatan pengembangan karier"]},
        ],
        "sitasi": "Hasibuan (2022); Dessler (2023)",
    },
    {
        "key": "lingkungan kerja",
        "aliases": ["work environment", "lingkungan kerja fisik", "kondisi kerja"],
        "label": "Lingkungan Kerja",
        "family": "kondisi",
        "definisi": (
            "keseluruhan kondisi fisik, sosial, dan psikologis tempat aparatur menjalankan "
            "tugas yang memengaruhi kenyamanan dan efektivitas kerja"
        ),
        "teori": [
            ("Manajemen SDM Sektor Publik", "Sedarmayanti (2022)",
             "membedakan lingkungan kerja fisik dan nonfisik serta kaitannya dengan produktivitas aparatur"),
            ("Manajemen Sumber Daya Manusia", "Mangkunegara (2022)",
             "menegaskan lingkungan kerja yang kondusif menurunkan kelelahan dan meningkatkan efektivitas kerja"),
        ],
        "mekanisme": (
            "mengurangi hambatan fisik dan sosial sehingga energi aparatur lebih terfokus "
            "pada penyelesaian tugas"
        ),
        "risiko": (
            "kelelahan akibat sarana yang tidak memadai, gangguan konsentrasi, dan "
            "ketegangan hubungan kerja"
        ),
        "tinggi": (
            "aparatur bekerja nyaman, aman, dan memperoleh dukungan sosial yang memadai"
        ),
        "intervensi": (
            "perbaikan tata ruang, sarana kerja, keselamatan kerja, dan iklim hubungan sosial"
        ),
        "dimensi": [
            {"nama": "Lingkungan Fisik", "indikator": [
                "Kenyamanan tata ruang dan pencahayaan",
                "Kelayakan sarana dan peralatan kerja"]},
            {"nama": "Lingkungan Sosial", "indikator": [
                "Kualitas hubungan dengan rekan kerja",
                "Dukungan atasan"]},
            {"nama": "Lingkungan Psikologis", "indikator": [
                "Rasa aman dalam bekerja",
                "Keseimbangan tuntutan dan dukungan kerja"]},
        ],
        "sitasi": "Sedarmayanti (2022); Mangkunegara (2022)",
    },
    {
        "key": "komitmen organisasi",
        "aliases": ["organizational commitment", "komitmen kerja"],
        "label": "Komitmen Organisasi",
        "family": "sikap",
        "definisi": (
            "keterikatan psikologis aparatur terhadap organisasi yang tercermin dalam "
            "komitmen afektif, kontinuans, dan normatif"
        ),
        "teori": [
            ("Organizational Behavior", "Colquitt et al. (2022)",
             "membedakan komitmen afektif, kontinuans, dan normatif sebagai bentuk keterikatan pada organisasi"),
            ("Organizational Behavior", "Kreitner & Kinicki (2022)",
             "mengaitkan komitmen dengan keinginan mempertahankan keanggotaan dan kesediaan berusaha"),
        ],
        "mekanisme": (
            "mengikat aparatur secara emosional dan normatif sehingga menurunkan niat keluar "
            "dan meningkatkan kesediaan berkontribusi"
        ),
        "risiko": (
            "komitmen instrumental yang lemah, niat berpindah, dan keterlibatan minimal"
        ),
        "tinggi": (
            "aparatur merasa terikat pada tujuan organisasi dan bersedia berusaha melampaui kewajiban"
        ),
        "intervensi": (
            "keadilan organisasi, dukungan pengembangan karier, dan penyelarasan nilai individu-organisasi"
        ),
        "dimensi": [
            {"nama": "Komitmen Afektif", "indikator": [
                "Rasa memiliki terhadap organisasi",
                "Keterikatan emosional pada tujuan organisasi"]},
            {"nama": "Komitmen Kontinuans", "indikator": [
                "Pertimbangan biaya meninggalkan organisasi",
                "Kestabilan keanggotaan"]},
            {"nama": "Komitmen Normatif", "indikator": [
                "Kewajiban moral bertahan",
                "Loyalitas terhadap institusi"]},
        ],
        "sitasi": "Colquitt et al. (2022); Kreitner & Kinicki (2022)",
    },
    {
        "key": "etos kerja",
        "aliases": ["work ethic", "etika kerja"],
        "label": "Etos Kerja",
        "family": "nilai",
        "definisi": (
            "sistem nilai dan sikap aparatur terhadap kerja yang menekankan kesungguhan, "
            "disiplin, tanggung jawab, dan makna kerja"
        ),
        "teori": [
            ("Etos Kerja Profesional", "Sinamo (2022)",
             "merumuskan etos kerja sebagai seperangkat nilai kerja profesional yang membentuk kualitas perilaku"),
            ("Membudayakan Etos Kerja", "Tasmara (2002)",
             "menekankan kerja sebagai ibadah dan amanah yang menuntut kesungguhan"),
        ],
        "mekanisme": (
            "membentuk standar internal tentang mutu usaha sehingga aparatur menilai "
            "ketekunan dan tanggung jawab sebagai kewajiban moral"
        ),
        "risiko": (
            "kerja seadanya, orientasi pada formalitas, dan rendahnya tanggung jawab atas mutu"
        ),
        "tinggi": (
            "aparatur bekerja sungguh-sungguh, disiplin, dan bertanggung jawab atas hasil"
        ),
        "intervensi": (
            "internalisasi nilai kerja, keteladanan pimpinan, dan penghargaan atas kerja bermutu"
        ),
        "dimensi": [
            {"nama": "Kesungguhan Kerja", "indikator": [
                "Intensitas usaha menyelesaikan tugas",
                "Kesediaan bekerja tuntas"]},
            {"nama": "Tanggung Jawab", "indikator": [
                "Kesediaan menanggung konsekuensi kerja",
                "Konsistensi memenuhi janji kerja"]},
            {"nama": "Makna Kerja", "indikator": [
                "Pemaknaan kerja sebagai amanah",
                "Orientasi pada manfaat bagi publik"]},
        ],
        "sitasi": "Sinamo (2022); Tasmara (2002)",
    },
    {
        "key": "pengawasan internal",
        "aliases": ["audit internal", "internal audit", "pengendalian internal", "spip", "pengawasan"],
        "label": "Pengawasan Internal",
        "family": "kontrol",
        "definisi": (
            "mekanisme pemantauan dan evaluasi independen yang dirancang untuk memastikan "
            "kepatuhan prosedur, keandalan pelaporan, dan efektivitas pengendalian"
        ),
        "teori": [
            ("Agency Theory", "Jensen & Meckling (1976)",
             "menjelaskan perlunya pemantauan untuk mengurangi asimetri informasi dan perilaku oportunistik agen"),
            ("Stewardship Theory", "Donaldson & Davis (1991)",
             "menawarkan kontrol berbasis kepercayaan dan penyelarasan kepentingan, bukan semata pengawasan represif"),
        ],
        "mekanisme": (
            "mendeteksi penyimpangan lebih awal dan menutup celah prosedural sehingga "
            "menurunkan risiko kesalahan serta penyalahgunaan wewenang"
        ),
        "risiko": (
            "temuan audit berulang, lemahnya tindak lanjut, dan formalitas pengendalian"
        ),
        "tinggi": (
            "risiko operasional menurun, rekomendasi ditindaklanjuti, dan kepatuhan prosedur menguat"
        ),
        "intervensi": (
            "penguatan independensi auditor, audit berbasis risiko, dan tindak lanjut temuan"
        ),
        "dimensi": [
            {"nama": "Independensi dan Objektivitas", "indikator": [
                "Kebebasan auditor dari tekanan kepentingan",
                "Objektivitas penilaian temuan"]},
            {"nama": "Kecukupan Prosedur", "indikator": [
                "Cakupan audit berbasis risiko",
                "Kualitas bukti dan dokumentasi"]},
            {"nama": "Tindak Lanjut", "indikator": [
                "Kecepatan penyelesaian rekomendasi",
                "Penurunan temuan berulang"]},
        ],
        "sitasi": "Jensen & Meckling (1976); Donaldson & Davis (1991)",
    },
    {
        "key": "pemberdayaan",
        "aliases": ["empowerment", "pemberdayaan pegawai", "psychological empowerment"],
        "label": "Pemberdayaan",
        "family": "kapasitas",
        "definisi": (
            "proses pemberian makna, kewenangan, dan ruang diskresi kepada aparatur agar "
            "mampu mengambil keputusan kerja secara bertanggung jawab"
        ),
        "teori": [
            ("Perilaku Organisasi", "Robbins & Judge (2022)",
             "mengaitkan pemberdayaan dengan peningkatan otonomi dan keyakinan kemampuan"),
            ("Organizational Behavior", "Luthans et al. (2022)",
             "menjelaskan pemberdayaan sebagai perluasan kendali aparatur atas pekerjaannya"),
        ],
        "mekanisme": (
            "memperluas ruang diskresi dan rasa kendali sehingga aparatur lebih proaktif "
            "dalam memecahkan masalah kerja"
        ),
        "risiko": (
            "ketergantungan pada atasan, rendahnya inisiatif, dan keputusan yang menunggu instruksi"
        ),
        "tinggi": (
            "aparatur berinisiatif, bertanggung jawab, dan mampu mengambil keputusan sesuai kewenangan"
        ),
        "intervensi": (
            "pendelegasian kewenangan, akses informasi, dan dukungan pengambilan keputusan"
        ),
        "dimensi": [
            {"nama": "Makna", "indikator": [
                "Kesesuaian tugas dengan nilai pribadi",
                "Pemaknaan kontribusi kerja"]},
            {"nama": "Kompetensi", "indikator": [
                "Keyakinan menyelesaikan tugas",
                "Penguasaan keterampilan yang dibutuhkan"]},
            {"nama": "Otonomi dan Dampak", "indikator": [
                "Kebebasan menentukan cara kerja",
                "Pengaruh hasil kerja terhadap unit"]},
        ],
        "sitasi": "Robbins & Judge (2022); Luthans et al. (2022)",
    },
    {
        "key": "pelatihan",
        "aliases": ["training", "diklat", "pengembangan sdm"],
        "label": "Pelatihan",
        "family": "kapasitas",
        "definisi": (
            "proses pembelajaran terencana untuk meningkatkan pengetahuan, keterampilan, "
            "dan sikap kerja aparatur sesuai tuntutan jabatan"
        ),
        "teori": [
            ("Human Resource Management", "Dessler (2023)",
             "menjelaskan pelatihan sebagai proses sistematis peningkatan kompetensi yang terkait dengan kebutuhan organisasi"),
            ("Human Resource Management", "Mathis et al. (2022)",
             "menekankan pentingnya analisis kebutuhan, desain, pelaksanaan, dan evaluasi pelatihan"),
        ],
        "mekanisme": (
            "menutup kesenjangan kompetensi sehingga aparatur lebih siap menghadapi "
            "tuntutan teknis dan perubahan prosedur"
        ),
        "risiko": (
            "pelatihan tidak sesuai kebutuhan, hasil tidak diterapkan, dan pemborosan anggaran"
        ),
        "tinggi": (
            "aparatur menerapkan pengetahuan baru, bekerja lebih efektif, dan meningkatkan mutu luaran"
        ),
        "intervensi": (
            "analisis kebutuhan pelatihan, desain berbasis praktik, dan evaluasi transfer pelatihan"
        ),
        "dimensi": [
            {"nama": "Reaksi dan Pembelajaran", "indikator": [
                "Kesesuaian materi dengan kebutuhan kerja",
                "Peningkatan pemahaman peserta"]},
            {"nama": "Perilaku Kerja", "indikator": [
                "Penerapan hasil pelatihan di tempat kerja",
                "Perubahan prosedur kerja ke arah lebih baik"]},
            {"nama": "Hasil", "indikator": [
                "Peningkatan mutu dan efisiensi kerja",
                "Kontribusi pada sasaran unit"]},
        ],
        "sitasi": "Dessler (2023); Mathis et al. (2022)",
    },
    {
        "key": "stres kerja",
        "aliases": ["job stress", "stres", "tekanan kerja"],
        "label": "Stres Kerja",
        "family": "tuntutan",
        "definisi": (
            "respons psikologis dan fisiologis aparatur terhadap tuntutan kerja yang "
            "melampaui kemampuan atau sumber daya yang dimiliki"
        ),
        "teori": [
            ("Psikologi Industri dan Organisasi", "Munandar (2021)",
             "menjelaskan stres kerja sebagai ketegangan yang memengaruhi emosi, proses berpikir, dan kondisi fisik"),
            ("Perilaku Organisasi", "Robbins & Judge (2022)",
             "memandang stres sebagai kondisi dinamis yang menghadapkan individu pada peluang, tuntutan, atau sumber daya yang diharapkan"),
        ],
        "mekanisme": (
            "menguras sumber daya kognitif dan emosional sehingga menurunkan konsentrasi, "
            "kualitas keputusan, dan ketahanan kerja"
        ),
        "risiko": (
            "kelelahan, mudah tersinggung, penurunan konsentrasi, dan gangguan kesehatan kerja"
        ),
        "tinggi": (
            "pada tingkat optimal, tekanan dapat memotivasi kewaspadaan dan prioritisasi kerja"
        ),
        "intervensi": (
            "manajemen tuntutan kerja, dukungan sosial, relaksasi, dan kejelasan peran"
        ),
        "dimensi": [
            {"nama": "Tuntutan Peran", "indikator": [
                "Beban tugas yang dirasakan",
                "Tekanan tenggat waktu"]},
            {"nama": "Konflik dan Ambiguitas Peran", "indikator": [
                "Ketidakjelasan harapan peran",
                "Konflik antara tuntutan kerja"]},
            {"nama": "Gejala Stres", "indikator": [
                "Kelelahan emosional",
                "Gangguan konsentrasi dan tidur"]},
        ],
        "sitasi": "Munandar (2021); Robbins & Judge (2022)",
    },
    {
        "key": "kualitas pelayanan",
        "aliases": ["service quality", "mutu layanan", "pelayanan publik"],
        "label": "Kualitas Pelayanan",
        "family": "luaran",
        "definisi": (
            "tingkat kesesuaian layanan dengan harapan pengguna yang tercermin pada keandalan, "
            "daya tanggap, jaminan, empati, dan bukti fisik layanan"
        ),
        "teori": [
            ("Manajemen", "Robbins & Coulter (2022)",
             "menekankan kualitas sebagai kesesuaian dengan standar dan kepuasan pengguna"),
            ("Manajemen SDM Sektor Publik", "Sedarmayanti (2022)",
             "mengaitkan kualitas pelayanan publik dengan kompetensi dan disiplin aparatur"),
        ],
        "mekanisme": (
            "menyelaraskan proses layanan dengan harapan pengguna sehingga kepuasan dan "
            "kepercayaan publik meningkat"
        ),
        "risiko": (
            "keluhan pengguna, ketidakpastian prosedur, dan penurunan kepercayaan publik"
        ),
        "tinggi": (
            "layanan cepat, akurat, ramah, dan konsisten dengan standar"
        ),
        "intervensi": (
            "standar pelayanan, pelatihan frontliner, dan sistem pengaduan yang responsif"
        ),
        "dimensi": [
            {"nama": "Keandalan", "indikator": [
                "Konsistensi layanan sesuai janji",
                "Akurasi hasil layanan"]},
            {"nama": "Daya Tanggap", "indikator": [
                "Kecepatan merespons kebutuhan pengguna",
                "Kesiapan membantu"]},
            {"nama": "Jaminan dan Empati", "indikator": [
                "Kompetensi petugas layanan",
                "Perhatian terhadap kebutuhan pengguna"]},
        ],
        "sitasi": "Robbins & Coulter (2022); Sedarmayanti (2022)",
    },
    {
        "key": "pencegahan kecurangan",
        "aliases": ["fraud prevention", "pencegahan fraud", "anti fraud", "pencegahan korupsi"],
        "label": "Pencegahan Kecurangan",
        "family": "kontrol",
        "definisi": (
            "rangkaian upaya preventif untuk menutup peluang, mengurangi rasionalisasi, "
            "dan menekan tekanan terjadinya kecurangan dalam organisasi"
        ),
        "teori": [
            ("Agency Theory", "Jensen & Meckling (1976)",
             "menjelaskan pentingnya pemantauan dan insentif yang selaras untuk mengurangi perilaku oportunistik"),
            ("Etika Profesi", "Keraf (2022)",
             "menekankan integritas dan tanggung jawab moral sebagai penopang kepatuhan"),
        ],
        "mekanisme": (
            "memperkuat pengendalian, transparansi, dan budaya integritas sehingga peluang "
            "serta rationalisasi kecurangan menyempit"
        ),
        "risiko": (
            "celah prosedur, konsentrasi wewenang, dan lemahnya kanal pelaporan"
        ),
        "tinggi": (
            "potensi kecurangan terdeteksi dini dan pelaporan pelanggaran meningkat"
        ),
        "intervensi": (
            "segregasi tugas, whistleblowing system, audit berbasis risiko, dan budaya integritas"
        ),
        "dimensi": [
            {"nama": "Pengendalian Internal", "indikator": [
                "Pemisahan fungsi yang memadai",
                "Keandalan prosedur otorisasi"]},
            {"nama": "Transparansi dan Akuntabilitas", "indikator": [
                "Keterbukaan informasi prosedur",
                "Ketertelusuran dokumen"]},
            {"nama": "Budaya Integritas", "indikator": [
                "Keteladanan pimpinan",
                "Ketersediaan kanal pelaporan"]},
        ],
        "sitasi": "Jensen & Meckling (1976); Keraf (2022)",
    },
]

# Indeks cepat berbasis alias
_ALIAS_INDEX: List[Tuple[str, Dict[str, Any]]] = []
for _entry in VAR_KB:
    for _alias in [_entry["key"], _entry.get("label", "")] + list(_entry.get("aliases") or []):
        if _alias:
            _ALIAS_INDEX.append((_norm(_alias), _entry))


def lookup(name: Any) -> Optional[Dict[str, Any]]:
    """Cari profil konstruk berdasarkan nama variabel (toleran sinonim)."""
    n = _norm(name)
    if not n:
        return None
    # 1) exact alias
    for alias, entry in _ALIAS_INDEX:
        if n == alias:
            return entry
    # 2) frasa terpanjang yang termuat dalam nama
    best: Optional[Tuple[int, Dict[str, Any]]] = None
    for alias, entry in _ALIAS_INDEX:
        if len(alias) >= 4 and alias in n:
            if best is None or len(alias) > best[0]:
                best = (len(alias), entry)
    if best:
        return best[1]
    # 3) nama termuat dalam alias (mis. "kompetensi" vs "kompetensi kerja")
    for alias, entry in _ALIAS_INDEX:
        if len(n) >= 4 and n in alias:
            return entry
    return None


# ==============================================================================
# 3. PROFIL TERPADU (KB + ONTOLOGI)
# ==============================================================================


def is_generic_dims(dims: Sequence[Dict[str, Any]]) -> bool:
    """Publik: apakah dimensi hanya pengulangan nama variabel (generik)?"""
    return _generic_dims(dims)


def _generic_dims(dims: Sequence[Dict[str, Any]]) -> bool:
    """Deteksi dimensi ontologi yang hanya pengulangan nama variabel."""
    if not dims:
        return True
    generic_markers = ("kapasitas ", "konsistensi ")
    for d in dims:
        nama = _norm(d.get("nama", ""))
        if not nama.startswith("dimensi "):
            return False
        if not any(m in nama for m in generic_markers):
            return False
    return True


def _ontology_definisi_ok(text: Any) -> bool:
    t = _norm(text)
    if len(t.split()) < 8:
        return False
    if "dalam perspektif manajemen strategis pada" in t:
        return False
    return True


def make_profile(name: str, vdata: Optional[Dict[str, Any]], ctx: Dict[str, Any]) -> Dict[str, Any]:
    """Gabungkan basis pengetahuan dengan data ontologi judul."""
    kb = lookup(name) or {}
    vdata = vdata or {}
    family = kb.get("family") or "kapasitas"

    definisi = vdata.get("definisi") if _ontology_definisi_ok(vdata.get("definisi")) else ""
    definisi = definisi or kb.get("definisi") or (
        f"konstruk {name} sebagaimana dioperasionalkan dalam konteks kerja organisasi"
    )

    dims = vdata.get("dimensi") or []
    if _generic_dims(dims):
        dims = kb.get("dimensi") or dims

    teori = kb.get("teori") or [
        ("Teori Perilaku Organisasi", "Robbins & Judge (2022)",
         "menjelaskan bahwa perilaku kerja dipengaruhi oleh karakteristik individu, tuntutan peran, dan konteks organisasi")
    ]

    return {
        "key": kb.get("key") or _norm(name).replace(" ", "_"),
        "label": name,
        "family": family,
        "definisi": definisi.strip().rstrip("."),
        "teori": teori,
        "mekanisme": kb.get("mekanisme") or (
            f"menggerakkan perubahan pada {name} melalui interaksi antara karakteristik individu dan tuntutan peran"
        ),
        "risiko": kb.get("risiko") or (
            f"ketidakoptimalan {name} yang tampak pada inkonsistensi perilaku kerja"
        ),
        "tinggi": kb.get("tinggi") or (
            f"kondisi {name} yang kondusif tercermin pada perilaku kerja yang lebih konsisten"
        ),
        "intervensi": kb.get("intervensi") or (
            f"penataan faktor pendukung {name} secara terencana dan berkelanjutan"
        ),
        "dimensi": dims,
        "sitasi": kb.get("sitasi") or (vdata.get("tokoh_ahli") or "Robbins & Judge (2022)"),
        "grand": (vdata.get("grand_theory") or "").strip(),
        "onto_definisi": (vdata.get("definisi") or "").strip(),
        "instansi": ctx.get("instansi", "Instansi Pemerintah"),
    }


def profile_map(ctx: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    onto_vars = (ctx.get("ontology") or {}).get("variables") or {}
    pm: Dict[str, Dict[str, Any]] = {}
    for i, x in enumerate(ctx.get("x_vars") or [], 1):
        pm[f"X{i}"] = make_profile(x, onto_vars.get(f"x{i}"), ctx)
    for j, m in enumerate(ctx.get("m_vars") or [], 1):
        key = "M" if len(ctx.get("m_vars") or []) == 1 else f"M{j}"
        pm[key] = make_profile(m, onto_vars.get("m"), ctx)
    for j, z in enumerate(ctx.get("z_vars") or [], 1):
        key = "Z" if len(ctx.get("z_vars") or []) == 1 else f"Z{j}"
        pm[key] = make_profile(z, onto_vars.get("z"), ctx)
    pm["Y"] = make_profile(ctx.get("y_var", "Kinerja"), onto_vars.get("y"), ctx)
    pm["_all"] = {  # type: ignore[assignment]
        "x": [pm[f"X{i}"] for i in range(1, len(ctx.get("x_vars") or []) + 1)],
        "m": [v for k, v in pm.items() if k.startswith("M") and k != "_all"],
        "z": [v for k, v in pm.items() if k.startswith("Z") and k != "_all"],
        "y": pm["Y"],
    }
    return pm


def profile_for_role(pm: Dict[str, Any], role: str) -> Dict[str, Any]:
    return pm.get(role.upper()) or pm.get("Y") or {}


# ==============================================================================
# 4. PERAKIT KALIMAT SEMANTIK
# ==============================================================================


def theory_sentence(p: Dict[str, Any], rng: random.Random, variant: int = 0) -> str:
    teori = p.get("teori") or []
    if not teori:
        return ""
    nama, sitasi, klaim = teori[variant % len(teori)]
    return f"{theory_entry(nama, sitasi)} {klaim}."


def _as_phrase(text: str) -> str:
    """Ubah kalimat definisi menjadi frasa yang aman disisipkan."""
    t = (text or "").strip().rstrip(".")
    return t


def join_id(items: Any) -> str:
    """Gabungkan nama variabel/dimensi secara gramatikal: 'A', 'A dan B', 'A, B, dan C'."""
    vals = [str(x).strip() for x in (items or []) if str(x).strip()]
    if not vals:
        return ""
    if len(vals) == 1:
        return vals[0]
    if len(vals) == 2:
        return f"{vals[0]} dan {vals[1]}"
    return ", ".join(vals[:-1]) + f", dan {vals[-1]}"


def premise_mayor(cause: Dict[str, Any], effect: Dict[str, Any], rng: random.Random, kind: str = "") -> str:
    """Premis mayor: dalil teoretis tentang relasi sebab-akibat."""
    c_theory = theory_sentence(cause, rng, 0)
    e_theory = theory_sentence(effect, rng, 0)
    connector = pick(rng, [
        "Kedua dalil tersebut bertemu pada satu titik:",
        "Titik temu kedua kerangka teoretis itu terletak pada proposisi bahwa",
        "Secara logis, kedua kerangka ini saling melengkapi dalam menjelaskan bahwa",
        "Konvergensi kedua teori tersebut menghasilkan proposisi bahwa",
    ])

    if kind == "mediasi":
        return (
            f"Premis Mayor: {c_theory} {e_theory} {connector} kapasitas yang melekat pada "
            f"{cause['label']} tidak berubah menjadi luaran kerja secara otomatis, melainkan "
            f"harus lebih dahulu diolah menjadi energi dan komitmen perilaku melalui "
            f"{effect['label']}."
        )
    if kind == "moderasi":
        return (
            f"Premis Mayor: {c_theory} {e_theory} {connector} kekuatan hubungan antara "
            f"{cause['label']} dan {effect['label']} tidak bersifat universal, melainkan "
            f"bergantung pada kondisi batas yang memengaruhi kemampuan aparatur mengolah "
            f"sumber daya tersebut."
        )
    if kind == "kondisional":
        return (
            f"Premis Mayor: {c_theory} {e_theory} {connector} efek transmisi {cause['label']} "
            f"terhadap {effect['label']} berjalan melalui mekanisme perantara yang intensitasnya "
            f"dapat bertambah atau berkurang sesuai kualitas sumber daya psikologis yang dimiliki aparatur."
        )
    return (
        f"Premis Mayor: {c_theory} {e_theory} {connector} {cause['label']} yang terkelola baik "
        f"akan {field(cause, 'mekanisme', rng)}, sehingga pada gilirannya {field(effect, 'tinggi', rng)}."
    )


def premise_minor(
    cause: Dict[str, Any], effect: Dict[str, Any], instansi: str, rng: random.Random, kind: str = ""
) -> str:
    """Premis minor: observasi kontekstual pada instansi."""
    opener = pick(rng, [
        "Premis Minor:",
        "Premis Minor:",
        "Premis Minor:",
    ])
    unit = pick(rng, [
        "unit teknis dan unit pelayanan",
        "bidang operasional dan sekretariat",
        "unit kerja yang berhadapan langsung dengan pengguna layanan",
        "sektor pelaksana dan unit pengawasan internal",
    ])
    if kind == "mediasi":
        return (
            f"{opener} Pada {instansi}, variasi capaian kerja pada {unit} tidak sepenuhnya "
            f"dijelaskan oleh perbedaan {cause['label']}; sebagian perbedaan itu tampak pada "
            f"kualitas {effect['label']} yang belum merata, sehingga jalur transmisi tersebut "
            f"layak diuji secara empiris."
        )
    if kind == "moderasi":
        return (
            f"{opener} Pada {instansi}, aparatur menghadapi tuntutan kerja yang relatif setara, "
            f"namun respons mereka berbeda; perbedaan itu mengindikasikan adanya kondisi psikologis "
            f"yang memperkuat atau memperlemah dampak {cause['label']} terhadap {effect['label']}."
        )
    if kind == "kondisional":
        return (
            f"{opener} Pada {instansi}, efek tidak langsung {cause['label']} terhadap {effect['label']} "
            f"diperkirakan tidak konstan, karena kualitas sumber daya psikologis aparatur berbeda "
            f"antarunit dan antarmasa kerja."
        )
    return (
        f"{opener} Pada {instansi}, masih dijumpai antara lain {field(cause, 'risiko', rng)} pada {unit}, "
        f"sementara kondisi ideal {effect['label']} belum sepenuhnya terbentuk, sehingga "
        f"hubungan kedua konstruk perlu dibuktikan secara empiris."
    )


def path_paragraph(
    cause: Dict[str, Any],
    effect: Dict[str, Any],
    instansi: str,
    rng: random.Random,
    relation: str = "langsung",
) -> str:
    """Paragraf naratif satu jalur kausal (bukan template hipotesis)."""
    theory = theory_sentence(cause, rng, 0)
    theory2 = theory_sentence(effect, rng, 1)
    bridge = pick(rng, [
        "Rantai penalaran ini dapat dijelaskan sebagai berikut.",
        "Mekanisme tersebut dapat dirunut secara bertahap.",
        "Dengan demikian, hubungan keduanya bukan relasi semu.",
        "Implikasi teoretisnya bersifat langsung dan dapat diuji.",
    ])
    if relation == "mediasi":
        return (
            f"{theory} {theory2} Dalam model penelitian ini, {cause['label']} diposisikan sebagai "
            f"anteseden yang {field(cause, 'mekanisme', rng)}. Energi yang terbentuk dari proses tersebut "
            f"kemudian menjadi bahan baku bagi {effect['label']} untuk {field(effect, 'mekanisme', rng)}. "
            f"{bridge} Apabila mata rantai itu terputus, maka risiko yang muncul adalah {field(cause, 'risiko', rng)}, dan pada saat yang "
            f"sama {effect['label']} kehilangan daya ungkitnya terhadap sasaran organisasi."
        )
    if relation == "moderasi":
        return (
            f"{theory} {theory2} Posisi {effect['label']} dalam model ini bukan sebagai akibat "
            f"langsung, melainkan sebagai kondisi batas yang menentukan sejauh mana {cause['label']} "
            f"mampu {field(cause, 'mekanisme', rng)}. Apabila {effect['label']} berada pada kondisi yang favorable, "
            f"maka {field(effect, 'tinggi', rng)}; sebaliknya, ketika kondisi tersebut rapuh, risiko yang muncul adalah {field(effect, 'risiko', rng)}. "
            f"{bridge} Karena itu, pengaruh {cause['label']} terhadap kinerja tidak dapat dibaca "
            f"sebagai konstanta, melainkan sebagai fungsi dari kualitas {effect['label']}."
        )
    return (
        f"{theory} {theory2} Secara operasional, {cause['label']} {field(cause, 'mekanisme', rng)}. "
        f"Proses itu memperbesar peluang bahwa {field(effect, 'tinggi', rng)}. {bridge} Di lingkungan {instansi}, "
        f"pengujian jalur ini penting karena masih dijumpai {field(cause, 'risiko', rng)} yang terlihat "
        f"dalam rutinitas kedinasan."
    )


def mediation_paragraph(
    cause: Dict[str, Any], mediator: Dict[str, Any], outcome: Dict[str, Any],
    instansi: str, rng: random.Random,
) -> str:
    return (
        f"Jalur tidak langsung dari {cause['label']} menuju {outcome['label']} melalui "
        f"{mediator['label']} dibangun atas tiga proposisi yang saling bergantung. Pertama, "
        f"{theory_sentence(cause, rng, 0)} Kedua, {theory_sentence(mediator, rng, 0)} Ketiga, "
        f"{theory_sentence(outcome, rng, 0)} Ketiga proposisi itu menyatu dalam satu logika: "
        f"{cause['label']} {field(cause, 'mekanisme', rng)}, lalu {mediator['label']} menyalurkan energi tersebut "
        f"agar {field(outcome, 'tinggi', rng)}. Pada konteks {instansi}, rangkaian ini menjelaskan mengapa "
        f"investasi pada {cause['label']} dapat gagal berbuah kinerja apabila risiko yang muncul adalah {field(mediator, 'risiko', rng)}."
    )


def moderation_paragraph(
    focal: Dict[str, Any], moderator: Dict[str, Any], outcome: Dict[str, Any],
    instansi: str, rng: random.Random,
) -> str:
    return (
        f"Peran {moderator['label']} sebagai variabel moderasi bertumpu pada argumen bahwa "
        f"dampak {focal['label']} terhadap {outcome['label']} tidak homogen di seluruh unit kerja. "
        f"{theory_sentence(moderator, rng, 0)} Dalam kerangka itu, {moderator['label']} berfungsi "
        f"sebagai sumber daya yang {field(moderator, 'mekanisme', rng)}. Ketika {moderator['label']} kuat, "
        f"{field(moderator, 'tinggi', rng)}, sehingga pengaruh {focal['label']} terhadap {outcome['label']} "
        f"teramplifikasi. Sebaliknya, apabila {moderator['label']} lemah, risiko yang muncul adalah "
        f"{field(moderator, 'risiko', rng)}; akibatnya, daya ungkit {focal['label']} "
        f"melemah meskipun kapasitas formal organisasi tetap tersedia. Pada {instansi}, variasi "
        f"kondisi tersebut menjadikan pengujian interaksi relevan secara teoretis dan praktis."
    )


def conditional_indirect_paragraph(
    cause: Dict[str, Any], mediator: Dict[str, Any], moderator: Dict[str, Any],
    outcome: Dict[str, Any], instansi: str, rng: random.Random,
) -> str:
    return (
        f"Mediasi termoderasi yang diuji dalam penelitian ini menggabungkan dua mekanisme: "
        f"transmisi {cause['label']} melalui {mediator['label']}, dan kondisi batas yang dimainkan "
        f"oleh {moderator['label']}. Secara teoretis, {theory_sentence(moderator, rng, 1)} "
        f"Konsekuensinya, efek tidak langsung {cause['label']} terhadap {outcome['label']} tidak "
        f"dibaca sebagai satu angka tunggal, melainkan sebagai fungsi bersyarat dari tingkat "
        f"{moderator['label']}. Apabila {moderator['label']} tinggi, jalur {mediator['label']} "
        f"bekerja lebih efektif karena {field(moderator, 'tinggi', rng)}. Sebaliknya, apabila "
        f"{moderator['label']} rendah, risiko yang muncul adalah {field(moderator, 'risiko', rng)}, "
        f"sehingga transmisi tersebut melemah. Pada {instansi}, pola bersyarat ini penting untuk "
        f"menghindari rekomendasi kebijakan yang terlalu generalisasi."
    )


# ==============================================================================
# 5. DEFINISI KONSEPTUAL & OPERASIONAL
# ==============================================================================


def conceptual_definition(p: Dict[str, Any], instansi: str, rng: random.Random) -> str:
    theory = theory_sentence(p, rng, 0)
    dims = [d.get("nama", "") for d in (p.get("dimensi") or []) if d.get("nama")]
    dim_txt = ", ".join(dims[:-1]) + ", serta " + dims[-1] if len(dims) > 1 else (dims[0] if dims else "dimensi utamanya")
    return (
        f"Secara konseptual, {p['label']} didefinisikan sebagai {_as_phrase(p['definisi'])}. "
        f"{theory} Definisi tersebut menempatkan {p['label']} bukan sebagai atribut tunggal, "
        f"melainkan sebagai konstruk yang terbentuk dari {dim_txt}. Dalam konteks {instansi}, "
        f"pemaknaan ini penting agar pengukuran tidak berhenti pada persepsi umum, tetapi "
        f"menjangkau perilaku kerja yang dapat diamati dan dinilai."
    )


def operational_definition(
    p: Dict[str, Any], instansi: str, rng: random.Random, total_items: int
) -> str:
    dims = p.get("dimensi") or []
    dim_txt = join_id(d.get("nama", "Dimensi") for d in dims) or "dimensi utama"
    return (
        f"Secara operasional, {p['label']} diukur melalui kuesioner tertutup dengan skala Likert "
        f"1 sampai 5 yang menggambarkan intensitas persepsi responden terhadap {dim_txt}. "
        f"Sebanyak {total_items} butir pernyataan disusun dari indikator setiap dimensi, lalu "
        f"diuji validitas dan reliabilitasnya sebelum digunakan pada survei utama di {instansi}. "
        f"Skor yang semakin tinggi menunjukkan kondisi {p['label']} yang semakin mendukung "
        f"efektivitas kerja, sedangkan skor rendah menunjukkan adanya ruang perbaikan pada "
        f"dimensi yang bersangkutan."
    )


# ==============================================================================
# 6. PARAFRASE LEKSIKAL DETERMINISTIK
# ==============================================================================

# Pasangan pengganti yang aman secara makna dan tata bahasa akademik.
# Hanya diterapkan pada narasi (bukan judul, tabel, sitasi, atau pernyataan hipotesis).
_LEX_PAIRS: List[Tuple[str, List[str]]] = [
    ("Penelitian ini", ["Studi ini", "Kajian ini"]),
    ("pertanyaan penelitian", ["pertanyaan studi", "pertanyaan kajian"]),
    ("hubungan kausal", ["kaitan sebab-akibat", "relasi kausal"]),
    ("strategi survei", ["desain survei", "rancangan survei"]),
    ("risiko yang muncul adalah", ["gejalanya adalah", "risikonya adalah", "yang tampak adalah"]),
    ("menuntut pengukuran", ["memerlukan pengukuran", "membutuhkan pengukuran"]),
    ("pemisahan pengaruh", ["pembedaan dampak", "pemisahan efek"]),
    ("pengujian signifikansi", ["uji kebermaknaan", "pengujian kebermaknaan"]),
    ("satu titik waktu", ["satu periode waktu", "waktu yang sama"]),
    ("unit analisis", ["satuan analisis"]),
    ("common method bias", ["bias metode umum"]),
    ("butir pernyataan", ["butir pertanyaan", "item pernyataan"]),
    ("bertujuan untuk", ["dimaksudkan untuk", "diarahkan untuk"]),
    ("sesuai dengan", ["selaras dengan", "sejalan dengan"]),
    ("dilakukan", ["dijalankan", "dikerjakan"]),
    ("menjelaskan", ["menguraikan", "memaparkan"]),
    ("penelitian", ["studi", "kajian"]),
    ("responden", ["partisipan"]),
    ("analisis", ["telaah", "pengkajian"]),
    ("instrumen", ["alat ukur", "perangkat ukur"]),
    ("kuesioner", ["angket", "daftar pertanyaan"]),
    ("yaitu", ["yakni", "antara lain"]),
    ("juga", ["pula"]),
    ("harus", ["wajib", "perlu"]),
    ("berdasarkan", ["mengacu pada", "bertolak dari"]),
    ("mengenai", ["tentang", "perihal"]),
    ("Oleh karena itu", ["Atas dasar itu", "Bertolak dari hal tersebut", "Karena itu"]),
    ("Karena itu", ["Oleh sebab itu", "Atas dasar itu", "Sejalan dengan itu"]),
    ("karena itu", ["oleh sebab itu", "atas dasar itu", "sejalan dengan itu"]),
    ("Dengan demikian", ["Sejalan dengan itu", "Dari sana dapat disimpulkan", "Atas dasar itu"]),
    ("Namun demikian", ["Akan tetapi", "Meskipun demikian", "Pada sisi lain"]),
    ("Selain itu", ["Di samping itu", "Lebih dari itu", "Tambahan pula"]),
    ("Sementara itu", ["Adapun", "Pada waktu yang sama", "Di sisi lain"]),
    ("Hal ini", ["Kondisi ini", "Fenomena ini", "Keadaan tersebut"]),
    ("merupakan", ["adalah", "menjadi"]),
    ("menunjukkan", ["memperlihatkan", "mengindikasikan", "menandakan"]),
    ("mencerminkan", ["merefleksikan", "menggambarkan"]),
    ("berpengaruh", ["berdampak", "memengaruhi"]),
    ("dalam rangka", ["guna", "untuk"]),
    ("secara simultan", ["secara serentak", "secara bersamaan"]),
    ("secara komprehensif", ["secara menyeluruh", "secara utuh"]),
    ("pada dasarnya", ["hakikatnya", "prinsipnya"]),
    ("sejalan dengan", ["konsisten dengan", "selaras dengan"]),
    ("berlandaskan", ["berpijak", "bertumpu"]),
    ("mendorong", ["memicu", "menggerakkan"]),
    ("meningkatkan", ["mengoptimalkan", "memperkuat"]),
    ("memengaruhi", ["mempengaruhi", "memberi dampak pada"]),
    ("sebagai konsekuensi", ["sebagai akibatnya", "implikasinya"]),
    ("memainkan peran", ["berperan", "mengemban peran"]),
    ("memberikan kontribusi", ["berkontribusi", "menyumbang"]),
    ("dalam konteks", ["dalam kerangka"]),
    ("menggambarkan", ["memotret", "melukiskan"]),
    ("menegaskan", ["menegaskan kembali", "menandaskan"]),
    ("menguraikan", ["menjabarkan", "memaparkan"]),
    ("tujuan", ["sasaran", "maksud"]),
]

_CITATION_PATTERN = re.compile(r"\([^()]*\b(?:19|20)\d{2}[a-z]?[^()]*\)")
_QUOTE_PATTERN = re.compile(r'"[^"]{3,}"')


def _protect(text: str) -> Tuple[str, List[str]]:
    protected: List[str] = []

    def _stash(match: re.Match) -> str:
        protected.append(match.group(0))
        return f"\u0001{len(protected) - 1}\u0001"

    t = _CITATION_PATTERN.sub(_stash, text)
    t = _QUOTE_PATTERN.sub(_stash, t)
    return t, protected


def _restore(text: str, protected: List[str]) -> str:
    for i, original in enumerate(protected):
        text = text.replace(f"\u0001{i}\u0001", original)
    return text


def _apply_pair(text: str, src: str, options: List[str], rng: random.Random, prob: float) -> str:
    if src not in text and src.lower() not in text.lower():
        return text
    pattern = re.compile(r"(?<![A-Za-z])" + re.escape(src) + r"(?![A-Za-z])")

    def _repl(match: re.Match) -> str:
        if rng.random() > prob:
            return match.group(0)
        chosen = options[rng.randrange(len(options))]
        original = match.group(0)
        # Cegah penggandaan bila opsi berawalan kata yang sama dan ekornya
        # sudah ada di teks (mis. "menegaskan kembali" -> "menegaskan kembali kembali").
        if chosen.lower().startswith(original.lower()):
            extra = chosen[len(original):]
            if extra and text[match.end():].startswith(extra):
                return match.group(0)
        if original[:1].isupper():
            chosen = chosen[:1].upper() + chosen[1:]
        return chosen

    return pattern.sub(_repl, text)


def field(prof: Dict[str, Any], key: str, rng: Optional[random.Random] = None) -> str:
    """Ambil isi profil dengan panjang bervariasi agar tidak berulang identik.

    Profil menyimpan deskripsi panjang (mis. ``risiko``). Bila kutipan panjang itu
    dipakai pada dua belas hipotesis, naskah menjadi repetitif. Fungsi ini memilih
    satu klausa, dua klausa, atau teks penuh secara deterministik.
    """
    text = str(prof.get(key, "") or "").strip()
    if not text:
        return text
    clauses = [c.strip(" ,;") for c in re.split(r"[;,]", text) if c.strip(" ,;")]
    if len(clauses) <= 1 or rng is None:
        return text
    mode = rng.choices([0, 1, 2], weights=[0.45, 0.35, 0.20])[0]
    if mode == 0:
        return clauses[0]
    if mode == 1 and len(clauses) >= 2:
        # Jangan gabungkan klausa kedua yang hanya berupa satu-dua kata
        # (mis. "waktu" pada "menyerap kapasitas kognitif, waktu").
        if len(clauses[1].split()) >= 4:
            return f"{clauses[0]}, {clauses[1]}"
        return clauses[0]
    return text


def theory_entry(nama: str, sitasi: str) -> str:
    """Gabungkan nama teori dan sitasi tanpa kurung ganda."""
    nama = (nama or "").strip()
    sitasi = (sitasi or "").strip()
    if not sitasi:
        return nama
    if "(" in sitasi and sitasi.endswith(")"):
        return f"{nama} menurut {sitasi}"
    return f"{nama} ({sitasi})"


def join_cites(first: str, rest: str) -> str:
    """Hindari sitasi ganda bila ``first`` sudah ada di dalam ``rest``."""
    first = (first or "").strip()
    rest = (rest or "").strip()
    if not first:
        return rest
    if not rest:
        return first
    key = first.split("(")[0].strip().lower()
    if key and key in rest.lower():
        return rest
    return f"{first}; {rest}"


def restructure(text: str, ctx: Dict[str, Any], salt: str = "", strength: float = 0.5) -> str:
    """Ubah struktur kalimat (fronting/pemecahan) secara deterministik.

    Hanya pola yang aman secara gramatikal yang disentuh: anak kalimat di awal
    dipindah ke akhir, dan klausa akibat dipisah menjadi kalimat baru. Tujuannya
    memutus rangkaian n-gram yang sama antarjudul, bukan mengubah makna.
    """
    if not text or len(text) < 80:
        return text
    rng = rng_for(ctx, f"restructure:{salt}")
    parts = re.split(r"(?<=[.!?])\s+", text)
    out: List[str] = []
    for s in parts:
        s = s.strip()
        if not s:
            continue
        if len(s) >= 90 and rng.random() < strength:
            m = re.match(
                r"^(Sebab|Lantaran|Meskipun|Walaupun)\s+(.+?),\s+(.+)$",
                s, re.I,
            )
            if m:
                head, clause, main = m.groups()
                main = main.rstrip(". ")
                s = f"{main[:1].upper()}{main[1:]}, {head.lower()} {clause}."
            else:
                for src, dst in [
                    (", sehingga ", ". Akibatnya, "),
                    (", karena itu ", ". Karena itu, "),
                    (", oleh sebab itu ", ". Oleh sebab itu, "),
                    (", akibatnya ", ". Akibatnya, "),
                ]:
                    if src in s:
                        s = s.replace(src, dst, 1)
                        break
        out.append(s)
    return " ".join(out)


def polish_and_restructure(text: str, ctx: Dict[str, Any], salt: str = "", strength: float = 0.7) -> str:
    """Parafrase leksikal lalu restrukturisasi kalimat."""
    return restructure(polish(text, ctx, salt, strength), ctx, f"{salt}-rs", 0.5)


def polish(text: str, ctx: Dict[str, Any], salt: str = "", strength: float = 0.7) -> str:
    """Parafrase narasi secara deterministik tanpa menyentuh sitasi/kutipan."""
    if not text or len(text) < 60:
        return text
    rng = rng_for(ctx, salt or text[:60])
    body, protected = _protect(text)
    for src, options in _LEX_PAIRS:
        body = _apply_pair(body, src, options, rng, strength)
    body = re.sub(r"[ \t]{2,}", " ", body)
    return _restore(body, protected)


# ==============================================================================
# 7. RENCANA UJI COBA INSTRUMEN
# ==============================================================================


def pilot_plan(ctx: Dict[str, Any]) -> Dict[str, str]:
    """Tentukan n uji coba dan r-tabel dua sisi (alpha 5%) secara konsisten.

    Nilai kritis r diambil dari tabel distribusi t untuk df = n - 2:
    n=30 -> df 28 -> 0,361; n=40 -> df 38 -> 0,312; n=50 -> df 48 -> 0,279;
    n=60 -> df 58 -> 0,254. Ukuran uji coba tidak boleh melebihi sampel utama.
    """
    n_main = int(ctx.get("n") or 0)
    if n_main >= 100:
        options = [50, 60]
    elif n_main >= 60:
        options = [40, 50]
    elif n_main >= 35:
        options = [30, 40]
    else:
        options = [max(15, n_main // 2)] if n_main >= 20 else [20]
    rng = rng_for(ctx, "pilot")
    n_pilot = options[rng.randrange(len(options))]
    table = {20: "0,444", 30: "0,361", 40: "0,312", 50: "0,279", 60: "0,254"}
    r_tabel = table.get(n_pilot)
    if r_tabel is None:
        # perkiraan konservatif untuk n di luar tabel
        r_tabel = "0,361" if n_pilot < 35 else "0,279"
    return {"n": str(n_pilot), "r_tabel": r_tabel}


# ==============================================================================
# 8. NARASI METODE (BAB III)
# ==============================================================================


def method_design_paragraphs(ctx: Dict[str, Any], pm: Dict[str, Any], rng: random.Random) -> List[str]:
    instansi = ctx["instansi"]
    y = pm["Y"]
    x_labels = join_id(p["label"] for p in pm["_all"]["x"])
    m_labels = ", ".join(p["label"] for p in pm["_all"]["m"]) or "mekanisme perantara"
    z_labels = ", ".join(p["label"] for p in pm["_all"]["z"]) or "kondisi kontekstual"

    p1 = pick(rng, [
        (
            f"Penelitian ini menggunakan pendekatan kuantitatif dengan strategi survei "
            f"eksplanatori untuk menguji hubungan kausal antarkonstruk pada {instansi} "
            f"(Creswell & Creswell, 2023). Pendekatan kuantitatif dipilih karena pertanyaan "
            f"penelitian menuntut pengukuran yang terstandar, pemisahan pengaruh langsung dan "
            f"tidak langsung, serta pengujian signifikansi yang dapat digeneralisasikan "
            f"(Hair et al., 2022). Instrumen utama berupa kuesioner tertutup yang mengukur "
            f"{x_labels}, {m_labels}, {z_labels}, dan {y['label']} pada satu titik waktu."
        ),
        (
            f"Kajian ini dirancang sebagai penelitian kuantitatif dengan strategi survei untuk "
            f"menguji keterkaitan kausal antarkonstruk di lingkungan {instansi} "
            f"(Creswell & Creswell, 2023). Pilihan itu berangkat dari kebutuhan memperoleh ukuran "
            f"yang terstandar, memisahkan pengaruh langsung dari pengaruh tidak langsung, serta "
            f"menilai kebermaknaan statistik yang dapat digeneralisasikan (Hair et al., 2022). "
            f"Data dikumpulkan melalui kuesioner tertutup yang memuat pengukuran {x_labels}, "
            f"{m_labels}, {z_labels}, dan {y['label']} dalam satu periode pengamatan."
        ),
        (
            f"Untuk menjawab pertanyaan penelitian, dipilih pendekatan kuantitatif dengan desain "
            f"survei eksplanatori pada {instansi} (Creswell & Creswell, 2023). Alasan pemilihan "
            f"terletak pada tuntutan pengukuran yang konsisten, penelusuran jalur langsung dan "
            f"tidak langsung secara terpisah, dan pengujian hipotesis yang memungkinkan "
            f"generalisasi (Hair et al., 2022). Adapun instrumen pokoknya adalah kuesioner "
            f"tertutup yang mengukur {x_labels}, {m_labels}, {z_labels}, dan {y['label']}."
        ),
    ])
    p2 = pick(rng, [
        (
            f"Desain cross-sectional dipakai untuk memotret konfigurasi {y['label']} pada "
            f"periode pengamatan yang sama (Cooper & Schindler, 2022). Sifat eksplanatori "
            f"terlihat dari penempatan {x_labels} sebagai anteseden, {m_labels} sebagai "
            f"mekanisme transmisi, dan {z_labels} sebagai kondisi batas. Dengan struktur itu, "
            f"penelitian tidak berhenti pada deskripsi {y['label']}, tetapi menguji rangkaian "
            f"proposisi yang menghubungkan {x_labels} dengan luaran kerja secara eksplisit."
        ),
        (
            f"Desain cross-sectional dipilih untuk menangkap keadaan {y['label']} pada rentang "
            f"waktu yang sama (Cooper & Schindler, 2022). Karakter eksplanatoris tampak dari "
            f"posisi {x_labels} sebagai anteseden, {m_labels} sebagai penyalur pengaruh, dan "
            f"{z_labels} sebagai pembatas kondisi. Susunan tersebut membuat kajian ini tidak "
            f"berhenti pada pelukisan {y['label']}, melainkan menguji rangkaian proposisi yang "
            f"menghubungkan {x_labels} dengan luaran kerja."
        ),
        (
            f"Penelitian memakai rancangan cross-sectional agar konfigurasi {y['label']} terekam "
            f"pada satu periode pengamatan (Cooper & Schindler, 2022). Sifat eksplanatorisnya "
            f"tercermin pada penempatan {x_labels} sebagai variabel pendahulu, {m_labels} sebagai "
            f"mekanisme transmisi, dan {z_labels} sebagai kondisi batas. Dengan begitu, fokus "
            f"penelitian melampaui deskripsi {y['label']} dan mengarah pada pengujian proposisi "
            f"antara {x_labels} dan luaran kerja."
        ),
    ])
    p3 = pick(rng, [
        (
            f"Unit analisis penelitian adalah aparatur pada {instansi} yang memiliki "
            f"pengalaman kerja memadai untuk menilai {y['label']} dan faktor-faktor yang "
            f"memengaruhinya. Pendekatan ini memungkinkan peneliti menangkap variasi persepsi "
            f"antarbidang, bukan hanya rata-rata organisasi (Neuman, 2022)."
        ),
        (
            f"Sasaran unit analisisnya adalah aparatur di {instansi} yang dipandang cukup "
            f"berpengalaman untuk menilai {y['label']} beserta faktor yang memengaruhinya. Corak "
            f"ini memungkinkan peneliti merekam keragaman persepsi antarunit, tidak sekadar nilai "
            f"rata-rata organisasi (Neuman, 2022)."
        ),
        (
            f"Unit yang dianalisis adalah aparatur {instansi} yang dinilai memiliki pengalaman "
            f"memadai untuk menimbang {y['label']} dan penentunya. Pendekatan ini membuka peluang "
            f"menangkap perbedaan persepsi antarunit kerja, alih-alih hanya memotret rerata "
            f"instansi (Neuman, 2022)."
        ),
    ])
    return [polish(p1, ctx, "3a-1"), polish(p2, ctx, "3a-2"), polish(p3, ctx, "3a-3")]


def method_location_paragraphs(ctx: Dict[str, Any], pm: Dict[str, Any], rng: random.Random) -> List[str]:
    instansi = ctx["instansi"]
    y = pm["Y"]
    x0 = pm["_all"]["x"][0] if pm["_all"]["x"] else {"label": "variabel eksogen"}
    p1 = pick(rng, [
        (
            f"Lokus penelitian ditetapkan pada {instansi} karena instansi ini mengelola fungsi "
            f"pelayanan dan penerimaan daerah yang menuntut akuntabilitas tinggi. Pemilihan lokus "
            f"didasarkan pada tiga pertimbangan: relevansi substantif dengan {x0['label']}, "
            f"ketersediaan data kinerja, dan akses peneliti terhadap unit kerja (Sekaran & Bougie, 2020). "
            f"Ketiga pertimbangan itu penting agar pengukuran {y['label']} tidak berjarak dari "
            f"praktik kerja yang sesungguhnya."
        ),
        (
            f"Riset ini dilaksanakan pada {instansi}, sebuah instansi yang menjalankan fungsi "
            f"pelayanan dan pemungutan penerimaan daerah dengan tuntutan akuntabilitas tinggi. "
            f"Tiga alasan mendasari pemilihan tempat penelitian: keterkaitan substantif dengan "
            f"{x0['label']}, ketersediaan data kinerja, dan kemudahan akses ke unit kerja "
            f"(Sekaran & Bougie, 2020). Ketiganya dipandang perlu agar pengukuran {y['label']} "
            f"tetap berakar pada praktik kerja sehari-hari."
        ),
        (
            f"Penelitian mengambil tempat di {instansi} karena lembaga tersebut mengemban "
            f"pelayanan publik dan pengelolaan penerimaan daerah yang menuntut "
            f"pertanggungjawaban ketat. Pemilihan ini bertumpu pada tiga hal: kesesuaian tema "
            f"dengan {x0['label']}, tersedianya data kinerja, dan keterjangkauan unit kerja bagi "
            f"peneliti (Sekaran & Bougie, 2020). Pertimbangan itu menjaga agar {y['label']} "
            f"diukur dalam konteks kerja yang nyata."
        ),
    ])
    p2 = pick(rng, [
        (
            f"Selain itu, {instansi} memiliki sebaran unit kerja dengan karakteristik tugas yang "
            f"beragam, sehingga memungkinkan peneliti mengamati variasi {y['label']} dan kondisi "
            f"kerja antarbidang (Neuman, 2022). Keragaman ini menjadi modal metodologis untuk "
            f"menguji apakah hubungan antarkonstruk bersifat seragam atau bergantung pada konteks "
            f"unit kerja. Keterbukaan pimpinan dan ketersediaan arsip kepegawaian juga menunjang "
            f"ketertelusuran data."
        ),
        (
            f"Di samping itu, unit kerja pada {instansi} memiliki karakteristik tugas yang "
            f"berbeda-beda sehingga membuka ruang bagi peneliti untuk mengamati variasi "
            f"{y['label']} dan situasi kerja antarunit (Neuman, 2022). Keragaman ini menjadi "
            f"bekal metodologis untuk menilai apakah hubungan antarkonstruk seragam atau "
            f"ditentukan konteks unit. Dukungan pimpinan serta arsip kepegawaian yang tertata "
            f"turut mempermudah penelusuran data."
        ),
        (
            f"Selain itu, struktur unit kerja {instansi} cukup beragam sehingga memungkinkan "
            f"peneliti menyimak perbedaan {y['label']} dan iklim kerja antarbidang (Neuman, 2022). "
            f"Keberagaman tersebut berguna untuk menguji apakah keterkaitan antarkonstruk berlaku "
            f"umum atau bergantung pada konteks unit. Keterbukaan pimpinan dan ketersediaan "
            f"dokumen kepegawaian juga menopang kelengkapan data."
        ),
    ])
    return [polish(p1, ctx, "3b-1"), polish(p2, ctx, "3b-2")]


def method_data_paragraphs(ctx: Dict[str, Any], pm: Dict[str, Any], rng: random.Random) -> List[str]:
    instansi = ctx["instansi"]
    n = ctx.get("n", 0)
    p1 = pick(rng, [
        (
            f"Data primer dikumpulkan melalui kuesioner yang disebarkan kepada {n} responden "
            f"terpilih di {instansi}. Kuesioner memuat identitas demografis, butir pernyataan "
            f"untuk setiap konstruk, dan ruang catatan bagi responden yang ingin menjelaskan "
            f"pengalaman kerjanya (Hair et al., 2022). Data sekunder diperoleh dari dokumen "
            f"perencanaan, laporan kinerja, dan arsip kepegawaian untuk memverifikasi konteks "
            f"dan memperkuat interpretasi temuan."
        ),
        (
            f"Data primer dijaring lewat kuesioner yang diberikan kepada {n} responden terpilih "
            f"di {instansi}. Kuesioner tersebut berisi identitas demografis, pernyataan tiap "
            f"konstruk, serta ruang bagi responden untuk menuturkan pengalaman kerjanya "
            f"(Hair et al., 2022). Adapun data sekunder ditelusuri dari dokumen perencanaan, "
            f"laporan kinerja, dan arsip kepegawaian untuk mengecek konteks dan menajamkan "
            f"interpretasi."
        ),
        (
            f"Pengumpulan data primer dilakukan dengan kuesioner kepada {n} responden di "
            f"{instansi}. Isinya mencakup data demografis, butir pernyataan untuk tiap konstruk, "
            f"dan kolom catatan bagi responden yang hendak menguraikan pengalamannya "
            f"(Hair et al., 2022). Sementara itu, data sekunder bersumber dari dokumen "
            f"perencanaan, laporan kinerja, dan arsip kepegawaian guna memverifikasi konteks "
            f"serta memperkaya pembahasan."
        ),
    ])
    p2 = pick(rng, [
        (
            f"Karena data primer bersumber dari laporan diri pada satu waktu, risiko common "
            f"method bias dikendalikan melalui prosedur pengumpulan dan pengujian statistik. "
            f"Responden dijamin anonim, urutan butir diacak antarkonstruk, dan tidak ada jawaban "
            f"yang dinilai benar salah (Podsakoff et al., 2024). Secara statistik, Harman's "
            f"single factor test digunakan untuk memastikan tidak ada satu faktor yang "
            f"mendominasi varians total (Hair et al., 2022)."
        ),
        (
            f"Karena seluruh data primer berupa penilaian diri yang dihimpun dalam satu waktu, "
            f"potensi common method bias ditekan melalui prosedur pengumpulan dan pengujian "
            f"statistik. Kerahasiaan responden dijaga, urutan butir diacak antar konstruk, dan "
            f"tidak ada jawaban yang dianggap benar atau salah (Podsakoff et al., 2024). Secara "
            f"statistik, Harman's single factor test dipakai untuk memastikan tidak ada faktor "
            f"tunggal yang mendominasi varians (Hair et al., 2022)."
        ),
        (
            f"Mengingat data primer berasal dari laporan diri pada satu titik waktu, risiko "
            f"common method bias dikendalikan lewat langkah prosedural dan statistik. Anonimitas "
            f"responden dijaga, urutan butir diacak lintas konstruk, dan penilaian tidak diberi "
            f"status benar atau salah (Podsakoff et al., 2024). Untuk pengujian statistiknya, "
            f"Harman's single factor test digunakan agar tidak ada satu faktor yang menyerap "
            f"mayoritas varians (Hair et al., 2022)."
        ),
    ])
    return [polish(p1, ctx, "3d-1"), polish(p2, ctx, "3d-2")]


def method_instrument_paragraphs(
    ctx: Dict[str, Any], pm: Dict[str, Any], rng: random.Random, total_items: int, pilot: Dict[str, str]
) -> List[str]:
    instansi = ctx["instansi"]
    p1 = pick(rng, [
        (
            f"Instrumen penelitian berupa kuesioner tertutup dengan skala Likert lima poin. "
            f"Sebanyak {total_items} butir pernyataan disusun dari dimensi dan indikator setiap "
            f"konstruk, lalu diterjemahkan ke dalam kalimat yang lugas dan tidak bermakna ganda "
            f"(Sekaran & Bougie, 2020). Setiap butir mengukur satu indikator agar hasil "
            f"pengukuran dapat ditelusuri kembali ke kerangka konseptual."
        ),
        (
            f"Instrumen yang dipakai adalah kuesioner tertutup berskala Likert lima titik. "
            f"Sebanyak {total_items} butir dirumuskan dari dimensi dan indikator tiap konstruk, "
            f"kemudian dikalimatkan secara lugas dan bebas makna ganda (Sekaran & Bougie, 2020). "
            f"Satu butir mewakili satu indikator agar hasil pengukuran dapat dirunut kembali ke "
            f"kerangka konseptual."
        ),
        (
            f"Alat pengumpul data berupa kuesioner tertutup dengan skala Likert lima pilihan. "
            f"Dari dimensi dan indikator setiap konstruk disusun {total_items} butir pernyataan "
            f"yang dirumuskan tegas dan tidak ambigu (Sekaran & Bougie, 2020). Setiap butir "
            f"menampung satu indikator sehingga pengukurannya dapat dilacak hingga kerangka "
            f"konsep."
        ),
    ])
    p2 = pick(rng, [
        (
            f"Skala 1 sampai 5 dipilih untuk memberi ruang variasi jawaban yang cukup tanpa "
            f"membebani responden. Sebelum digunakan pada survei utama di {instansi}, instrumen "
            f"melalui telaah ahli dan uji coba kepada {pilot['n']} responden di luar sampel utama "
            f"(Hair et al., 2022). Hasil uji coba dipakai untuk memilih butir yang valid dan "
            f"menghitung koefisien reliabilitas."
        ),
        (
            f"Skala 1 hingga 5 digunakan agar variasi jawaban memadai tanpa memberatkan "
            f"responden. Sebelum diterapkan pada survei utama di {instansi}, instrumen ditelaah "
            f"ahli dan diujicobakan kepada {pilot['n']} responden di luar sampel utama "
            f"(Hair et al., 2022). Hasil uji coba dipakai untuk menyeleksi butir yang valid dan "
            f"menghitung reliabilitasnya."
        ),
        (
            f"Rentang skala 1 sampai 5 ditetapkan untuk menyediakan pilihan jawaban yang cukup "
            f"tanpa membebani responden. Sebelum dipakai pada survei utama di {instansi}, "
            f"instrumen melalui penelaahan pakar dan uji coba pada {pilot['n']} responden di luar "
            f"sampel inti (Hair et al., 2022). Temuan uji coba menjadi dasar pemilihan butir "
            f"valid dan perhitungan reliabilitas."
        ),
    ])
    return [polish(p1, ctx, "3g-1"), polish(p2, ctx, "3g-2")]


def method_analysis_paragraphs(ctx: Dict[str, Any], pm: Dict[str, Any], rng: random.Random) -> List[str]:
    instansi = ctx["instansi"]
    y = pm["Y"]
    n = ctx.get("n", 0)
    m_or = ", ".join(p["label"] for p in pm["_all"]["m"]) or "perantara"
    p1 = pick(rng, [
        (
            f"Analisis data dilakukan dengan PLS-SEM menggunakan SmartPLS 4 karena model "
            f"menggabungkan jalur langsung, mediasi, dan moderasi dalam satu kerangka "
            f"(Hair et al., 2022; Sarstedt et al., 2022). Dengan n = {n}, model memenuhi "
            f"kecukupan sampel untuk estimasi parameter dan pengujian efek tidak langsung. "
            f"PLS-SEM juga tidak menuntut normalitas multivariat yang ketat dan tetap stabil "
            f"pada data persepsi berskala ordinal."
        ),
        (
            f"Analisis data memakai PLS-SEM dengan perangkat SmartPLS 4 sebab model menyatukan "
            f"jalur langsung, mediasi, dan moderasi dalam satu kerangka (Hair et al., 2022; "
            f"Sarstedt et al., 2022). Dengan n = {n}, model sudah memadai untuk mengestimasi "
            f"parameter dan menguji efek tidak langsung. PLS-SEM pula tidak menuntut asumsi "
            f"normalitas multivariat yang ketat dan tetap kokoh pada data persepsi berskala "
            f"ordinal."
        ),
        (
            f"Pengolahan data dilakukan dengan PLS-SEM melalui SmartPLS 4 karena model "
            f"menyatukan pengaruh langsung, mediasi, sekaligus moderasi (Hair et al., 2022; "
            f"Sarstedt et al., 2022). Pada n = {n}, ukuran sampel dinilai cukup untuk estimasi "
            f"parameter dan pengujian efek tidak langsung. Pendekatan ini juga lebih lentur "
            f"terhadap data persepsi berskala ordinal dan tidak menuntut normalitas multivariat "
            f"yang ketat."
        ),
    ])
    p2 = pick(rng, [
        (
            f"Evaluasi dilakukan dalam tiga tahap. Tahap pertama menilai outer model melalui "
            f"outer loading, AVE, composite reliability, dan HTMT untuk memastikan indikator "
            f"mengukur konstruknya secara valid dan reliabel. Tahap kedua menilai inner model "
            f"melalui VIF, R-square, f-square, Q-square, dan SRMR guna membaca kekuatan "
            f"prediksi serta kualitas model (Hair et al., 2022). Tahap ketiga menguji hipotesis "
            f"melalui bootstrapping 5.000 subsampel pada alpha 5 persen."
        ),
        (
            f"Evaluasi ditempuh dalam tiga tahapan. Tahap awal menilai outer model lewat outer "
            f"loading, AVE, composite reliability, dan HTMT untuk memastikan indikator valid dan "
            f"reliabel dalam mengukur konstruknya. Tahap berikutnya menilai inner model melalui "
            f"VIF, R-square, f-square, Q-square, dan SRMR untuk membaca daya prediksi dan mutu "
            f"model (Hair et al., 2022). Tahap terakhir menguji hipotesis dengan bootstrapping "
            f"5.000 subsampel pada taraf alpha 5 persen."
        ),
        (
            f"Prosedur evaluasi dibagi menjadi tiga bagian. Pertama, penilaian outer model "
            f"melalui outer loading, AVE, composite reliability, dan HTMT guna memastikan "
            f"indikator mengukur konstruk secara sahih dan andal. Kedua, penilaian inner model "
            f"lewat VIF, R-square, f-square, Q-square, dan SRMR untuk menimbang kekuatan "
            f"prediksi serta kualitas model (Hair et al., 2022). Ketiga, pengujian hipotesis "
            f"melalui bootstrapping 5.000 subsampel pada alpha 5 persen."
        ),
    ])
    p3 = pick(rng, [
        (
            f"Efek mediasi {m_or} diuji melalui specific indirect effect dengan bias-corrected "
            f"confidence interval, sedangkan efek moderasi diuji melalui latent interaction "
            f"term. Interpretasi diarahkan pada signifikansi, arah, dan besaran efek, bukan "
            f"semata nilai p, agar rekomendasi bagi {instansi} memiliki dasar yang proporsional "
            f"(Hayes, 2022)."
        ),
        (
            f"Efek mediasi {m_or} diuji lewat specific indirect effect dengan bias-corrected "
            f"confidence interval, sementara efek moderasi diuji melalui latent interaction "
            f"term. Penafsiran diarahkan pada signifikansi, arah, dan besaran efek, bukan "
            f"sekadar nilai p, sehingga saran bagi {instansi} berdiri di atas dasar yang "
            f"proporsional (Hayes, 2022)."
        ),
        (
            f"Pengujian efek mediasi {m_or} memakai specific indirect effect dengan "
            f"bias-corrected confidence interval, sedangkan efek moderasi diuji melalui latent "
            f"interaction term. Penafsirannya menekankan signifikansi, arah, dan besaran efek, "
            f"tidak hanya nilai p, agar rekomendasi untuk {instansi} berpijak pada dasar yang "
            f"sepadan (Hayes, 2022)."
        ),
    ])
    return [polish(p1, ctx, "3i-1"), polish(p2, ctx, "3i-2"), polish(p3, ctx, "3i-3")]


# ==============================================================================
# 9. PENYARINGAN AKHIR DOKUMEN (paragraf statis yang belum divariasikan)
# ==============================================================================

_FORMULA_MARKERS = ("=", "Σ", "√", "×", "β", "γ", "r_xy", "r_11", "≤", "≥", "α", "σ")


def vary_document(doc: Any, ctx: Dict[str, Any], strength: float = 0.65) -> int:
    """Parafrase paragraf naratif yang belum melewati ``pv``.

    Pembangun lama masih memuat paragraf statis (kajian teori, waktu penelitian,
    populasi/sampel). Alih-alih menulis ulang ratusan pemanggilan ``p()``, fungsi
    ini menyaring paragraf setelah dokumen dirakit. Paragraf yang sudah divariasikan
    ditandai ``_nv_varied``; judul, pernyataan hipotesis, rumus, tabel, daftar
    pustaka, dan butir pertanyaan dilewati.

    Mengembalikan jumlah paragraf yang diubah.
    """
    import re as _re

    try:
        from build_full_thesis_docx import add_rich_runs_to_para
    except Exception:  # pragma: no cover - fallback bila modul tidak tersedia
        return 0

    after_dapus = False
    changed = 0

    for idx, para in enumerate(getattr(doc, "paragraphs", [])):
        text = (para.text or "").replace("\x0b", "\n").strip()
        if not text:
            continue
        if text.upper() == "DAFTAR PUSTAKA":
            after_dapus = True
            continue
        if after_dapus:
            continue
        if getattr(para, "_nv_varied", False):
            continue
        if len(text) < 100 or "\n" in text:
            continue
        if any(m in text for m in _FORMULA_MARKERS):
            continue
        if text.startswith(("Tabel ", "Gambar ", "Sumber", "BAB ", "DAFTAR")):
            continue
        if "?" in text:
            continue
        if _re.match(r"^\d+\.\s", text) and len(text) < 600:
            continue

        runs = [r for r in para.runs if (r.text or "").strip()]
        if runs and all(bool(r.bold) for r in runs):
            continue

        new_text = polish(text, ctx, f"doc-{idx}", strength=strength)
        new_text = restructure(new_text, ctx, f"doc-{idx}", strength=0.5)
        if not new_text or new_text == text:
            continue

        # Ganti isi paragraf tanpa mengubah format paragraf (indentasi, spasi,
        # perataan). Format inline dibangun ulang oleh add_rich_runs_to_para.
        for run in list(para.runs):
            run._element.getparent().remove(run._element)
        add_rich_runs_to_para(para, new_text, bold_default=False, italic_default=False, font_size=12)
        try:
            para._nv_varied = True
        except Exception:
            pass
        changed += 1

    return changed
