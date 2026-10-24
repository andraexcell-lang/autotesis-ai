# backend/academic_dimension_researcher.py
"""
ACADEMIC DIMENSION & INDICATOR RESEARCHER
Mesin Riset Ilmiah Dimensi, Indikator, dan Teori Operasional Variabel Tesis Magister (S2).

Membangun "Blueprint Ontologi Tesis" yang mengunci:
- Teori Ahli & Grand Theory per Variabel (X1, X2, M, Y)
- Dimensi Baku
- Indikator Operasional Baku (Skala Likert 1-5)
- Rujukan Jurnal Bereputasi Utama (dengan DOI aktif)
- Pemetaan Benang Merah Otomatis:
  * Indikator Y  --> Parameter Baris Tabel 1.1 (Evaluasi Kinerja 3 Tahun)
  * Indikator M  --> Masalah Operasional Tabel 1.2 (Rekapitulasi 12 Bulan)
  * Indikator X1 & X2 --> Narasi Telaah Fenomena Lapangan
  * Tabel Kisi-Kisi Instrumen Bab III (Tabel 3.1)
  * Lampiran Kuesioner Penelitian Lengkap
"""

import hashlib
import os
import json
import re
import shutil
import urllib.request
import urllib.parse
from semantic_title_parser import parse_academic_title

# Berkas lama (satu judul) dipertahankan sebagai sumber migrasi, tetapi TIDAK
# lagi ditulis oleh judul baru. Setiap judul mendapat berkas sendiri agar
# blueprint JUNGJUNG tidak tertimpa.
ONTOLOGY_CACHE_FILE = os.path.join(os.path.dirname(__file__), "cache_thesis_ontology.json")
ONTOLOGY_CACHE_DIR = os.path.join(os.path.dirname(__file__), "ontology_cache")


def _slug_judul(judul):
    s = re.sub(r"[^a-z0-9]+", "_", (judul or "").lower()).strip("_")
    return (s[:48] or "judul")


def ontology_cache_path(judul):
    """Path berkas ontologi khusus satu judul (hash judul, bukan judul mentah)."""
    k = hashlib.md5((judul or "").strip().lower().encode("utf-8")).hexdigest()[:16]
    return os.path.join(ONTOLOGY_CACHE_DIR, f"{_slug_judul(judul)}_{k}.json")


def _judul_cocok(data, judul):
    try:
        meta = (data or {}).get("metadata", {}) or {}
        return ((meta.get("judul") or "").strip().lower()
                == (judul or "").strip().lower())
    except Exception:
        return False


def baca_ontology_cache(judul):
    """Baca blueprint judul ini; migrasikan berkas lama bila judulnya cocok.

    Migrasi memakai salinan byte-per-byte supaya hash berkas (dan cache generasi)
    tidak berubah hanya karena format JSON ditulis ulang.
    """
    judul = (judul or "").strip()
    if not judul:
        return None
    path = ontology_cache_path(judul)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if data.get("variables"):
                return data
        except Exception as e:
            print(f"[ontology] berkas judul rusak, coba berkas lama: {e}")

    if os.path.exists(ONTOLOGY_CACHE_FILE):
        try:
            with open(ONTOLOGY_CACHE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if data.get("variables") and _judul_cocok(data, judul):
                os.makedirs(ONTOLOGY_CACHE_DIR, exist_ok=True)
                shutil.copyfile(ONTOLOGY_CACHE_FILE, path)
                print(f"[ontology] migrasi blueprint lama -> {os.path.basename(path)}")
                return data
        except Exception as e:
            print(f"[ontology] gagal membaca berkas lama: {e}")
    return None


def simpan_ontology_cache(data, judul=None):
    """Tulis blueprint ke berkas judul ini secara atomik. Berkas lama tidak disentuh."""
    judul = (judul or ((data or {}).get("metadata", {}) or {}).get("judul") or "").strip()
    if not judul:
        raise ValueError("judul ontologi kosong")
    os.makedirs(ONTOLOGY_CACHE_DIR, exist_ok=True)
    path = ontology_cache_path(judul)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp, path)
    return path

# ==============================================================================
# BASIS PENGETAHUAN ONTOLOGI BAKU MSDM S2 (50+ VARIABEL RISET POPULER)
# Standar Buku Teks Internasional & Jurnal Bereputasi Terakreditasi
# ==============================================================================
MSDM_ONTOLOGY_KNOWLEDGE_BASE = {
    # 1. Kinerja Pegawai / Karyawan / Aparatur (Y)
    "kinerja": {
        "nama_resmi": "Kinerja Pegawai",
        "grand_theory": "Goal-Setting Theory (Locke & Latham, 1990)",
        "tokoh_ahli": "Mangkunegara (2022); Kasmir (2023); Afandi (2023)",
        "definisi": "Hasil kerja secara kualitas dan kuantitas yang dicapai oleh seorang pegawai dalam melaksanakan tugasnya sesuai dengan tanggung jawab yang diberikan.",
        "dimensi": [
            {"nama": "Kualitas Pekerjaan", "indikator": ["Ketelitian dan kerapian hasil tugas", "Kesesuaian hasil kerja dengan Standar Operasional Prosedur (SOP)"]},
            {"nama": "Kuantitas Output", "indikator": ["Pencapaian target volume pekerjaan rutin", "Kemampuan menuntaskan beban tugas tambahan"]},
            {"nama": "Ketepatan Waktu", "indikator": ["Penyelesaian pekerjaan sesuai tenggat waktu (deadline)", "Kecepatan merespons instruksi dan kebutuhan kedinasan"]},
            {"nama": "Efektivitas & Kerjasama", "indikator": ["Optimalisasi pemanfaatan sarana kerja", "Koordinasi dan inisiatif saling mendukung dalam tim"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Ulya & Rahmah, 2025)",
            "entri": "Ulya, N., & Rahmah, M. (2025). Pengukuran Kinerja Pegawai Pelayanan Publik: Tinjauan Teori dan Aplikasi Indikator Kinerja. Jurnal Tata Pamong, 7(1), 89–104. https://doi.org/10.33701/jtp.v7i1.3412"
        }
    },

    # 2. Kepuasan Kerja (M / X)
    "kepuasan_kerja": {
        "nama_resmi": "Kepuasan Kerja",
        "grand_theory": "Value-Percept Theory (Locke, 1976); Two-Factor Theory (Herzberg, 1966)",
        "tokoh_ahli": "Locke (1976); Hackman & Oldham (1980); Robbins & Judge (2022)",
        "definisi": "Keadaan emosional positif dan menyenangkan yang timbul dari penilaian individu terhadap pekerjaannya atau pengalaman kerjanya.",
        "dimensi": [
            {"nama": "Kepuasan terhadap Pekerjaan itu Sendiri", "indikator": ["Kesesuaian tugas dengan minat dan keahlian", "Tantangan tugas yang memicu kebanggaan profesional"]},
            {"nama": "Kepuasan terhadap Rekognisi & Apresiasi", "indikator": ["Pengakuan pimpinan atas prestasi kerja pegawai", "Keadilan pemberian insentif dan penghargaan"]},
            {"nama": "Kepuasan terhadap Supervisi & Rekan Kerja", "indikator": ["Dukungan dan bimbingan teknis dari atasan", "Hubungan interpersonal dan kekompakan rekan sekerja"]},
            {"nama": "Kepuasan terhadap Kondisi Lingkungan", "indikator": ["Kenyamanan dan kelengkapan fasilitas penunjang tugas", "Kejelasan alur koordinasi dan minimnya birokrasi kaku"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Siregar & Fadli, 2023)",
            "entri": "Siregar, F. H., & Fadli, A. (2023). Pengaruh Budaya Organisasi Terhadap Kinerja Pegawai dengan Kepuasan Kerja Sebagai Variabel Intervening. Jurnal Riset Akuntansi dan Manajemen Malahayati, 12(2), 145–158. https://doi.org/10.33024/jrm.v12i2.9812"
        }
    },

    # 3. Organizational Citizenship Behavior / OCB (X / M)
    "ocb": {
        "nama_resmi": "Organizational Citizenship Behavior (OCB)",
        "grand_theory": "Social Exchange Theory (Blau, 1964)",
        "tokoh_ahli": "Organ (1988); Podsakoff et al. (2000)",
        "definisi": "Perilaku kerja ekstra-peran (discretionary effort) individual yang tidak diatur secara langsung atau eksplisit dalam sistem imbalan formal, namun secara keseluruhan meningkatkan efektivitas fungsi organisasi.",
        "dimensi": [
            {"nama": "Altruism (Kerelaan Menolong)", "indikator": ["Kerelaan membantu rekan kerja yang kelebihan beban tugas", "Membimbing pegawai baru atau rekan yang mengalami kendala teknis"]},
            {"nama": "Conscientiousness (Kehati-hatian & Ketaatan)", "indikator": ["Kehadiran dan ketertiban kerja melampaui standar formal", "Kepatuhan sukarela terhadap aturan organisasi"]},
            {"nama": "Sportsmanship (Sportivitas)", "indikator": ["Toleransi terhadap kendala lingkungan kerja tanpa banyak mengeluh", "Fokus pada penyelesaian solusi daripada membesar-besarkan masalah"]},
            {"nama": "Courtesy (Kesantunan)", "indikator": ["Mencegah gesekan kerja melalui komunikasi yang santun dan terbuka", "Menghargai hak dan pendapat rekan kerja"]},
            {"nama": "Civic Virtue (Kebajikan Kewargaan)", "indikator": ["Partisipasi aktif dalam kegiatan penunjang organisasi", "Kepedulian terhadap reputasi dan kelangsungan institusi"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Hani & Pratama, 2023)",
            "entri": "Hani, T. T., & Pratama, A. (2023). Peran Organizational Citizenship Behavior dalam Memediasi Pengaruh Iklim Organisasi Terhadap Kinerja Pegawai. Jurnal Ilmiah Manajemen dan Kewirausahaan, 3(1), 112–125. https://doi.org/10.55606/jimak.v3i1.1219"
        }
    },

    # 4. Motivasi Kerja (X / M)
    "motivasi": {
        "nama_resmi": "Motivasi Kerja",
        "grand_theory": "Need Theory (McClelland, 1987); Public Service Motivation (Perry & Wise, 1990)",
        "tokoh_ahli": "Herzberg (1966); McClelland (1987); Kasmir (2023); Afandi (2023)",
        "definisi": "Daya dorong psikologis intrinsik dan ekstrinsik yang mengarahkan dan memelihara kegigihan perilaku individu untuk mencapai tujuan organisasi.",
        "dimensi": [
            {"nama": "Kebutuhan Berprestasi (Need for Achievement)", "indikator": ["Dorongan untuk menuntaskan target yang menantang", "Keinginan meningkatkan keahlian dan standar kualitas tugas"]},
            {"nama": "Kebutuhan Afiliasi (Need for Affiliation)", "indikator": ["Antusiasme membangun sinergi dan hubungan harmonis dengan rekan", "Kesiapan berkolaborasi dalam kerja tim lintas seksi"]},
            {"nama": "Kebutuhan Pengaruh/Kekuasaan (Need for Power)", "indikator": ["Inisiatif memberikan masukan konstruktif dalam rapat dinas", "Tanggung jawab memimpin dan mengarahkan jalannya penugasan"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Farhan & Indriyaningrum, 2025)",
            "entri": "Farhan, M., & Indriyaningrum, S. (2025). Pengaruh Motivasi Kerja dan Kepuasan Kerja terhadap Produktivitas Aparatur. Jurnal Manajemen dan Kebijakan Publik, 10(1), 45–59. https://doi.org/10.22146/jmkp.71234"
        }
    },

    # 5. Promosi Jabatan (X)
    "promosi_jabatan": {
        "nama_resmi": "Promosi Jabatan",
        "grand_theory": "Expectancy Theory (Vroom, 1964); Equity Theory (Adams, 1965)",
        "tokoh_ahli": "Hasibuan (2022); Flippo (2018); Dessler (2023)",
        "definisi": "Perpindahan posisi pegawai dari suatu jabatan ke jabatan lain yang memiliki status, tanggung jawab, wewenang, dan hak yang lebih tinggi berdasarkan prinsip meritokrasi.",
        "dimensi": [
            {"nama": "Keadilan Prosedural Promosi", "indikator": ["Transparansi kriteria dan mekanisme seleksi kenaikan jabatan", "Objektivitas penilaian berbasis portofolio dan rekam jejak"]},
            {"nama": "Kesesuaian Kompetensi & Kualifikasi", "indikator": ["Kesesuaian jabatan baru dengan latar belakang keahlian pegawai", "Uji kelayakan dan kepatutan yang terstandar"]},
            {"nama": "Peluang & Kejelasan Jenjang Karier", "indikator": ["Ketersediaan formasi jenjang karier yang proporsional", "Kepastian masa kerja dan persyaratan kenaikan pangkat"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Hendri & Yunita, 2023)",
            "entri": "Hendri, H., & Yunita, Y. (2023). Pengaruh Promosi Jabatan dan Mutasi Terhadap Kinerja Pegawai Negeri Sipil. Jurnal Ekonomi Manajemen Sistem Informasi, 4(5), 845–856. https://doi.org/10.38035/jemsi.v4i5.1412"
        }
    },

    # 6. Mutasi Kerja / Rotasi Kerja (X)
    "mutasi_kerja": {
        "nama_resmi": "Mutasi Kerja",
        "grand_theory": "Person-Environment Fit Theory (Edwards, 1996)",
        "tokoh_ahli": "Simamora (2020); Hasibuan (2022); Siswanto (2021)",
        "definisi": "Kegiatan ketenagakerjaan yang berhubungan dengan pemindahan pegawai dari suatu posisi atau unit kerja ke posisi lain yang setara guna penyegaran wawasan dan peningkatan kapabilitas organisasi.",
        "dimensi": [
            {"nama": "Prinsip Kesesuaian Tugas (Job-Fit)", "indikator": ["Penempatan tugas baru sesuai dengan kompetensi dan pengalaman", "Kemudahan beradaptasi dengan ritme dan uraian tugas baru"]},
            {"nama": "Penyegaran & Tour of Duty", "indikator": ["Pemberian kesempatan memperkaya wawasan fungsi organisasi", "Penghindaran kejenuhan kerja pada satu unit penugasan"]},
            {"nama": "Keadilan & Kebijakan Rotasi", "indikator": ["Transparansi pertimbangan rotasi tugas kedinasan", "Pemberitahuan dan persiapan sebelum alih tugas dilaksanakan"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Gunawan & Hartono, 2023)",
            "entri": "Gunawan, A., & Hartono, R. (2023). Pengaruh Mutasi Kerja dan Beban Kerja Terhadap Kepuasan dan Kinerja Aparatur Sipil Negara. Jurnal Riset Manajemen dan Bisnis, 8(1), 55–67. https://doi.org/10.36407/jrmb.v8i1.790"
        }
    },

    # 7. Pengembangan Karier (X)
    "pengembangan_karier": {
        "nama_resmi": "Pengembangan Karier",
        "grand_theory": "Career Development Theory (Hall, 1976); Super's Life-Span Theory (1980)",
        "tokoh_ahli": "Dessler (2023); Hall (2002); Mangkunegara (2022)",
        "definisi": "Proses formal dan terarah yang dirancang organisasi untuk membantu pegawai memperluas kapabilitas profesional, pengalaman, dan jenjang kepangkatan sepanjang masa kerjanya.",
        "dimensi": [
            {"nama": "Perencanaan Karier Individual", "indikator": ["Kejelasan arah dan target kemajuan jabatan pegawai", "Keselarasan aspirasi pegawai dengan kebutuhan instansi"]},
            {"nama": "Bimbingan & Mentoring Atasan", "indikator": ["Dukungan atasan dalam memberikan coaching dan feedback", "Peluang mengikuti program sertifikasi keahlian profesional"]},
            {"nama": "Akses Pelatihan Fungsional", "indikator": ["Pemerataan kesempatan mengikuti pendidikan dan pelatihan (diklat)", "Relevansi materi diklat terhadap tuntutan modernisasi birokrasi"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Hidayat & Sugiarto, 2022)",
            "entri": "Hidayat, R., & Sugiarto, E. (2022). Pengaruh Perencanaan dan Pengembangan Karir Terhadap Kepuasan Kerja dan Kinerja Pegawai. Jurnal Ilmiah Manajemen Publik, 8(2), 78–91. https://doi.org/10.25157/jimp.v8i2.6189"
        }
    },

    # 8. Budaya Organisasi (X)
    "budaya_organisasi": {
        "nama_resmi": "Budaya Organisasi",
        "grand_theory": "Organizational Culture Theory (Schein, 1985; Cameron & Quinn, 2011)",
        "tokoh_ahli": "Schein (2010); Robbins & Judge (2022); Tika (2020)",
        "definisi": "Sistem makna, nilai-nilai bersama, keyakinan dasar, dan norma perilaku yang dianut oleh anggota organisasi yang membedakan organisasi tersebut dari organisasi lainnya.",
        "dimensi": [
            {"nama": "Integritas & Akuntabilitas (Core Values)", "indikator": ["Konsistensi memegang teguh kejujuran dalam pelaporan kedinasan", "Pertanggungjawaban profesional atas setiap kewenangan"]},
            {"nama": "Orientasi Pelayanan & Kolaborasi", "indikator": ["Dedikasi memberikan pelayanan prima kepada stakeholder", "Keterbukaan bekerjasama dan berbagi informasi antar-seksi"]},
            {"nama": "Adaptabilitas & Inovasi", "indikator": ["Kesiapan mengadopsi sistem kerja digital baru", "Keterbukaan terhadap gagasan perbaikan prosedur kerja"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Lestari & Nugroho, 2024)",
            "entri": "Lestari, K., & Nugroho, B. (2024). Pengaruh Budaya Organisasi dan Komitmen Terhadap Kinerja Pegawai Instansi Pemerintah Daerah. SeMaKa: Jurnal Manajemen dan Bisnis, 3(1), 45–56. https://doi.org/10.56916/semaka.v3i1.621"
        }
    },

    # 9. Disiplin Kerja (X)
    "disiplin": {
        "nama_resmi": "Disiplin Kerja",
        "grand_theory": "Theory of Planned Behavior (Ajzen, 1991)",
        "tokoh_ahli": "Sutrisno (2022); Hasibuan (2022); Afandi (2023)",
        "definisi": "Kesadaran dan kesediaan pegawai dalam menaati semua peraturan organisasi, norma sosial yang berlaku, dan standar operasional prosedur kedinasan secara konsisten.",
        "dimensi": [
            {"nama": "Ketaatan Waktu Kerja", "indikator": ["Ketepatan jam masuk, istirahat, dan jam pulang kedinasan", "Pemanfaatan jam dinas secara efektif untuk kepentingan pekerjaan"]},
            {"nama": "Kepatuhan terhadap Prosedur (SOP)", "indikator": ["Menjalankan instruksi kedinasan sesuai pedoman baku", "Ketertiban administrasi dalam pertanggungjawaban berkas"]},
            {"nama": "Tanggung Jawab Pemeliharaan Sarana", "indikator": ["Menjaga dan memelihara perlengkapan inventaris kantor", "Kesiapan mematuhi kode etik dan tata tertib instansi"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Prasetio, Ruminta, & Akbar, 2025)",
            "entri": "Prasetio, A., Ruminta, & Akbar, M. (2025). Analisis Kedisiplinan Kerja dan Dimensinya pada Pegawai Pemerintah Daerah. Jurnal Ilmu Administrasi, 12(1), 55–68. https://doi.org/10.31113/jia.v12i1.890"
        }
    },

    # 10. Employee Engagement (M / Y)
    "engagement_komitmen": {
        "nama_resmi": "Employee Engagement",
        "grand_theory": "Job Demands-Resources (JD-R) Model (Bakker & Demerouti, 2007)",
        "tokoh_ahli": "Schaufeli et al. (2006); Kaswan (2022); Rivai & Sagala (2022)",
        "definisi": "Kondisi psikologis positif dan memuaskan yang terkait dengan pekerjaan, yang ditandai oleh semangat (vigor), dedikasi (dedication), dan penyerapan (absorption).",
        "dimensi": [
            {"nama": "Vigor (Ketangguhan Mental & Energi)", "indikator": ["Tingkat energi dan ketahanan mental saat bekerja", "Kegigihan untuk terus berusaha saat menghadapi hambatan"]},
            {"nama": "Dedication (Antusiasme & Kebanggaan)", "indikator": ["Rasa bangga yang mendalam terhadap institusi tempat bekerja", "Antusiasme dan inspirasi yang dirasakan dalam menjalankan tugas"]},
            {"nama": "Absorption (Fokus & Penyerapan)", "indikator": ["Konsentrasi kerja penuh dan mendalam saat jam dinas", "Kerap larut secara positif dalam menyelesaikan pekerjaan"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Susanto & Wijaya, 2025)",
            "entri": "Susanto, H., & Wijaya, K. (2025). Keterikatan Pegawai dan Reformasi Birokrasi di Era Transformasi Digital. Jurnal Kebijakan dan Manajemen Publik, 13(1), 101–115. https://doi.org/10.20473/jkmp.v13i1.51201"
        }
    },

    # 11. Kecerdasan Emosional (X)
    "kecerdasan_emosional": {
        "nama_resmi": "Kecerdasan Emosional",
        "grand_theory": "Emotional Intelligence Framework (Goleman, 1995)",
        "tokoh_ahli": "Goleman (2001); Mayer & Salovey (1997)",
        "definisi": "Kemampuan individu untuk mengenali perasaan diri sendiri dan orang lain, memotivasi diri, serta mengelola emosi dengan baik pada diri sendiri dan dalam hubungan dengan orang lain.",
        "dimensi": [
            {"nama": "Kesadaran Diri (Self-Awareness)", "indikator": ["Mengenali emosi dan suasana hati saat menghadapi tekanan kerja", "Memahami kekuatan dan keterbatasan kapasitas diri"]},
            {"nama": "Pengaturan Diri (Self-Regulation)", "indikator": ["Mengendalikan reaksi emosional negatif di bawah tekanan tenggat waktu", "Menjaga ketenangan dan berpikir jernih saat menghadapi konflik kerja"]},
            {"nama": "Empati & Hubungan Sosial (Social Skills)", "indikator": ["Peka terhadap sudut pandang dan perasaan rekan kerja", "Mampu menjalin komunikasi yang menenangkan dan persuasif"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Baiduri & Muliawan, 2021)",
            "entri": "Baiduri, I., & Muliawan, F. M. (2021). Bagimana Pengaruh Variabel Emosional dan Motivasi Terhadap Kinerja Pegawai. Transekonomika: Akuntansi Bisnis dan Keuangan, 1(1), 25–30. https://doi.org/10.55047/transekonomika.v1i1.10"
        }
    },

    # 12. Beban Kerja (X)
    "beban_kerja": {
        "nama_resmi": "Beban Kerja",
        "grand_theory": "Cognitive Load Theory (Sweller, 1988); NASA Task Load Index (Hart, 2006)",
        "tokoh_ahli": "Hart & Staveland (1988); Sunyoto (2020); Munandar (2021)",
        "definisi": "Sekumpulan tugas dan tanggung jawab yang harus diselesaikan oleh seorang pegawai atau unit organisasi dalam jangka waktu tertentu dengan pengerahan kapasitas fisik dan mental.",
        "dimensi": [
            {"nama": "Beban Waktu (Temporal Demand)", "indikator": ["Kepadatan tenggat waktu penyelesaian tugas harian", "Keterbatasan waktu istirahat akibat volume pekerjaan menumpuk"]},
            {"nama": "Beban Mental (Mental Demand)", "indikator": ["Tingkat konsentrasi dan ketelitian kognitif yang dituntut pekerjaan", "Kompleksitas pemecahan masalah dan regulasi teknis"]},
            {"nama": "Beban Fisik & Tekanan Kerja (Effort & Frustration)", "indikator": ["Keletihan fisik yang dirasakan setelah jam kerja", "Tingkat stres akibat penumpukan berkas yang belum tertuntaskan"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Wahyuni & Arifin, 2023)",
            "entri": "Wahyuni, S., & Arifin, M. I. (2023). Pengaruh Beban Kerja dan Stres Kerja Terhadap Kinerja Tenaga Kerja Publik. Jurnal Riset Manajemen Sains Indonesia, 14(2), 120–132. https://doi.org/10.21009/jrmsi.014.2.06"
        }
    },

    # 13. Pelatihan (X)
    "pelatihan": {
        "nama_resmi": "Pelatihan",
        "grand_theory": "Social Learning Theory (Bandura, 1977); Training Transfer Theory (Baldwin & Ford, 1988)",
        "tokoh_ahli": "Noe (2020); Dessler (2021); Mangkunegara (2022)",
        "definisi": "Proses sistematis pemberian pengetahuan, keterampilan, dan sikap kepada pegawai agar mampu melaksanakan tugas jabatan secara efektif dan efisien.",
        "dimensi": [
            {"nama": "Materi Pelatihan", "indikator": ["Relevansi materi dengan kebutuhan tugas jabatan", "Kejelasan dan kemutakhiran isi materi pelatihan"]},
            {"nama": "Metode Pelatihan", "indikator": ["Kesesuaian metode dengan gaya belajar peserta", "Keterlibatan aktif peserta dalam praktik dan simulasi"]},
            {"nama": "Instruktur/Pelatih", "indikator": ["Penguasaan kompetensi dan pengalaman instruktur", "Kemampuan instruktur menyampaikan materi secara mudah dipahami"]},
            {"nama": "Fasilitas dan Sarana", "indikator": ["Kelengkapan sarana pendukung kegiatan pelatihan", "Kenyamanan lingkungan belajar selama pelatihan"]}
        ],
        "jurnal_rujukan": {
            "sitasi": "(Siregar & Tanjung, 2023)",
            "entri": "Siregar, R., & Tanjung, M. (2023). Pengaruh Kompetensi dan Pelatihan Terhadap Kinerja Pegawai Negeri Sipil. Jurnal Manajemen dan Organisasi Terapan, 4(2), 88–101. https://doi.org/10.21009/jmot.042.07"
        }
    }
}


def detect_variable_key(var_name):
    """Mendeteksi kunci ontologi leksikal dari nama variabel penelitian."""
    v = (var_name or "").lower().strip()
    if any(k in v for k in ["kinerja", "performance", "prestasi kerja"]):
        return "kinerja"
    elif any(k in v for k in ["kepuasan", "satisfaction", "puas"]):
        return "kepuasan_kerja"
    elif any(k in v for k in ["citizenship", "ocb", "kewargaan", "ekstra peran"]):
        return "ocb"
    elif any(k in v for k in ["promosi", "kenaikan pangkat", "pangkat"]):
        return "promosi_jabatan"
    elif any(k in v for k in ["mutasi", "rotasi", "tour of duty", "alih tugas"]):
        return "mutasi_kerja"
    elif any(k in v for k in ["karier", "karir", "career"]):
        return "pengembangan_karier"
    elif any(k in v for k in ["budaya", "culture", "nilai organisasi"]):
        return "budaya_organisasi"
    elif any(k in v for k in ["disiplin", "discipline", "tata tertib"]):
        return "disiplin"
    elif any(k in v for k in ["engagement", "keterikatan", "komitmen"]):
        return "engagement_komitmen"
    elif any(k in v for k in ["emosi", "emotional", "kecerdasan"]):
        return "kecerdasan_emosional"
    elif any(k in v for k in ["beban", "workload"]):
        return "beban_kerja"
    elif any(k in v for k in ["motivasi", "motivation", "dorongan"]):
        return "motivasi"
    elif any(k in v for k in ["pelatihan", "training", "diklat", "workshop", "sosialisasi"]):
        return "pelatihan"
    else:
        # ATURAN ANTI-BUG SENYAP: variabel di luar knowledge base TIDAK BOLEH
        # diam-diam dianggap "motivasi" (merusak Bab II/III/Lampiran tanpa error).
        # Penulis tahu ada variabel asing sehingga bisa dikonfirmasi pengguna.
        return "unknown"


_KATA_INGGRIS = set(
    "the of and to in is are was were this that with for from has have been will can "
    "not but which their its as at by on or an it also such more than used are these".split()
)

# Terjemahan deterministik istilah indikator tunggal yang lazim dari sumber
# berbahasa Inggris. Ini MENERJEMAHKAN, bukan mengganti/menambah indikator:
# jumlah butir dan makna konstruk tetap sama.
_ISTILAH_INDONESIA = {
    "reaction": "Reaksi",
    "learning": "Pembelajaran",
    "behavior": "Perilaku",
    "behaviour": "Perilaku",
    "result": "Hasil",
    "results": "Hasil",
    "knowledge": "Pengetahuan",
    "skill": "Keterampilan",
    "skills": "Keterampilan",
    "attitude": "Sikap",
    "attitudes": "Sikap",
    "ability": "Kemampuan",
    "abilities": "Kemampuan",
    "motivation": "Motivasi",
    "discipline": "Disiplin",
    "communication": "Komunikasi",
    "teamwork": "Kerja Sama Tim",
    "leadership": "Kepemimpinan",
    "compensation": "Kompensasi",
    "satisfaction": "Kepuasan",
    "performance": "Kinerja",
    "training": "Pelatihan",
    "development": "Pengembangan",
    "career": "Karier",
    "trust": "Kepercayaan",
    "commitment": "Komitmen",
    "engagement": "Keterikatan",
}

# Indikator Indonesia baku (fallback deterministik bila Gemini tidak tersedia).
_FALLBACK_INDIKATOR = {
    "altruism": [
        "Membantu rekan kerja yang mengalami kesulitan menyelesaikan tugas",
        "Memberikan bantuan tanpa mengharapkan imbalan",
        "Berbagi beban pekerjaan dengan rekan satu tim",
    ],
    "conscientiousness": [
        "Mematuhi prosedur kerja yang berlaku",
        "Menyelesaikan tugas dengan teliti dan tepat waktu",
        "Hadir dan aktif sesuai jadwal kerja yang ditetapkan",
    ],
    "sportsmanship": [
        "Menahan diri untuk tidak mengeluh atas beban kerja",
        "Menerima keputusan organisasi dengan sikap positif",
        "Menjaga suasana kerja tetap kondusif",
    ],
    "courtesy": [
        "Bersikap sopan kepada rekan kerja dan atasan",
        "Berkoordinasi sebelum mengambil keputusan yang memengaruhi rekan",
        "Menjaga komunikasi kerja yang santun",
    ],
    "civic virtue": [
        "Terlibat aktif dalam kegiatan organisasi",
        "Mengikuti perkembangan informasi dan kebijakan organisasi",
        "Memberikan masukan untuk perbaikan organisasi",
    ],
}


def _indikator_perlu_indonesia(ind_list):
    """True bila ada indikator berbahasa Inggris / kalimat mentah dari sumber."""
    for s in (ind_list or []):
        t = (s or "").strip()
        if not t:
            continue
        low = t.lower().strip(" .,:;")
        # Istilah konstruk tunggal (Reaction, Learning, ...) -> perlu diterjemahkan.
        if " " not in low and low in _ISTILAH_INDONESIA:
            return True
        kata = re.findall(r"[A-Za-z]+", low)
        if kata:
            rasio = sum(1 for x in kata if x in _KATA_INGGRIS) / len(kata)
            if rasio > 0.12:
                return True
        if re.search(r"indicators? used are|the researcher|figure\s*\d|is a person who|"
                     r"closely related|towards|studies have|this study", t, re.I):
            return True
    return False


def pastikan_indikator_indonesia(onto, var_name=None, role=None):
    """Pastikan INDIKATOR berbahasa Indonesia; NAMA DIMENSI dibiarkan apa adanya.

    Dipakai untuk kasus nyata OCB: dimensi (altruism, sportsmanship, courtesy,
    civic virtue) memang berbahasa Inggris dan dipertahankan, tetapi indikator
    yang menempel masih kalimat Inggris mentah dari abstrak jurnal.
    """
    dims = (onto or {}).get("dimensi") or []
    if not dims:
        return onto

    # 0. Terjemahan deterministik istilah tunggal (jumlah butir TETAP).
    #    Contoh: Training Dessler -> Reaksi, Pembelajaran, Perilaku, Hasil.
    translated = 0
    for d in dims:
        baru = []
        for s in (d.get("indikator") or []):
            key = str(s or "").strip().rstrip(".").lower()
            if key in _ISTILAH_INDONESIA:
                baru.append(_ISTILAH_INDONESIA[key])
                translated += 1
            else:
                baru.append(s)
        d["indikator"] = baru
    if translated and not any(_indikator_perlu_indonesia(d.get("indikator") or [])
                              for d in dims):
        print(f"[indikator] {translated} istilah diterjemahkan deterministik "
              "(jumlah indikator tetap)")
        return onto

    perlu = any(_indikator_perlu_indonesia(d.get("indikator") or []) for d in dims) \
        or any(len(d.get("indikator") or []) < 2 for d in dims)
    if not perlu:
        return onto

    nama_dims = [d.get("nama", "") for d in dims if d.get("nama")]
    hasil = None
    try:
        import llm_background_writer as W
        hasil = W.write_dimension_indicators(
            var_name or onto.get("nama_resmi", ""),
            nama_dims,
            source_label=((onto.get("jurnal_rujukan") or {}).get("sitasi", "")),
            tokoh=onto.get("tokoh_ahli", ""),
        )
    except Exception as ex:
        print(f"[indikator] penjabaran ulang via Gemini gagal: {ex}")

    if hasil and len(hasil.get("dimensi_list") or []) == len(dims):
        for d, y in zip(dims, hasil["dimensi_list"]):
            d["indikator"] = list(y.get("indikator") or [])
            d["pernyataan"] = [""] * len(d["indikator"])
        print(f"[indikator] {len(dims)} dimensi variabel '{var_name}' "
              f"dijabarkan ulang ke Bahasa Indonesia (Gemini); nama dimensi tetap.")
        return onto

    # Fallback deterministik
    dipakai = 0
    for d in dims:
        key = (d.get("nama") or "").strip().lower()
        if key in _FALLBACK_INDIKATOR and (
                _indikator_perlu_indonesia(d.get("indikator") or [])
                or len(d.get("indikator") or []) < 2):
            d["indikator"] = list(_FALLBACK_INDIKATOR[key])
            d["pernyataan"] = [""] * len(d["indikator"])
            dipakai += 1
    if dipakai:
        print(f"[indikator] {dipakai} dimensi variabel '{var_name}' memakai indikator "
              f"Indonesia baku (fallback deterministik).")
    return onto


def research_variable_ontology(var_name, role="X", judul="", instansi=""):
    """
    Meriset teori ahli, dimensi, indikator, dan jurnal untuk satu variabel.

    MODE MURNI AI + SUMBER TERVERIFIKASI: dimensi/indikator WAJIB berasal dari
    sumber nyata (pustaka permanen atau telusur live OpenAlex 2022-2026 dengan
    kutipan terverifikasi). Bila tidak ada sumber yang lolos, variabel ditandai
    needs_review dan pengguna diminta unggah jurnal/buku khusus,
    bukan diisi karangan model.
    """
    try:
        import llm_ontology_researcher
        llm_res = llm_ontology_researcher.research_variable_llm(
            var_name, role=role, judul=judul, instansi=instansi)
        if llm_res:
            # B9: nama dimensi boleh tetap bahasa aslinya (mis. altruism),
            # tetapi INDIKATOR wajib Bahasa Indonesia. Bila sumber memberi
            # kalimat Inggris mentah, indikator dijabarkan ulang via Gemini.
            return pastikan_indikator_indonesia(llm_res, var_name=var_name, role=role)
    except Exception as e:
        print(f"[ontology] agen LLM gagal untuk '{var_name}': {e}")

    # Tidak ada sumber terverifikasi -> tandai, jangan isi template/karangan.
    return {
        "role": role,
        "input_name": var_name,
        "nama_resmi": var_name if var_name else "(belum diriset)",
        "grand_theory": "",
        "tokoh_ahli": "",
        "definisi": "",
        "dimensi": [],
        "jurnal_rujukan": {"sitasi": "", "entri": ""},
        "skala_pengukuran": "Skala Likert 1–5 (1: Sangat Tidak Setuju s/d 5: Sangat Setuju)",
        "engine": "ai-failed",
        "needs_review": True,
        "needs_review_note": (
            f"Variabel \"{var_name}\" belum punya sumber dimensi/indikator terverifikasi "
            "(pencarian jurnal/buku 2022–2026 belum menemukan yang lolos). "
            "Silakan unggah jurnal/buku khusus berisi indikator variabel ini, "
            "atau riset ulang."
        ),
    }


def _kuesioner_statement_for(row):
    """
    Fallback pernyataan kuesioner bila agen AI tidak menyediakannya.
    Mengubah frasa indikator menjadi kalimat pernyataan Likert yang wajar,
    bukan lagi 'mengoptimalkan <frasa mentah>'.
    """
    ind = str(row.get("indikator", "")).strip()
    if not ind:
        return "Saya menunjukkan sikap positif dalam pelaksanaan tugas kedinasan."
    # Buang tanda kurung & rapikan
    core = re.sub(r"\s*\([^)]*\)", "", ind).strip().rstrip(".")
    first = core.split()[0].lower() if core.split() else ""
    verba = ("mampu", "menyelesaikan", "memahami", "menjaga", "mematuhi",
             "menguasai", "melaksanakan", "memberikan", "membantu", "mengikuti",
             "menyusun", "mengelola", "berpartisipasi", "meningkatkan", "berbagi",
             "menahan", "menerima", "bersikap", "berkoordinasi", "terlibat",
             "menghadiri", "mendukung", "berinisiatif", "menawarkan", "aktif",
             "tetap", "tidak", "memanfaatkan", "menyesuaikan", "mempertahankan")
    # Frasa yang sudah berbentuk aktivitas/kapabilitas -> langsung sebagai predikat.
    # Kata berawalan verba Indonesia (me-, mem-, men-, meng-, ber-, ter-, di-)
    # juga diperlakukan sebagai predikat agar tidak muncul "Saya menunjukkan menawarkan...".
    teks = (core[0].lower() + core[1:]) if core else core
    if first in verba or re.match(r"^(me|mem|men|meng|ber|ter|di)", first):
        return f"Saya {teks}."
    return f"Saya menunjukkan {teks}."


# Frasa penutup berulang yang dulu menempel di hampir semua butir kuesioner.
_SUFIKS_BOILERPLATE = (
    " dalam pelaksanaan tanggung jawab kedinasan",
    " dalam pelaksanaan tugas kedinasan",
)
_VERBA_DEPAN_RE = re.compile(r"^(me|mem|men|meng|ber|ter|di)")


def _normalisasi_pernyataan(teks):
    """Rapikan butir pernyataan kuesioner agar gramatikal & konsisten (B4).

    - membuang frasa penutup berulang ("... dalam pelaksanaan tanggung jawab kedinasan");
    - 'Saya menunjukkan <verba> ...' -> 'Saya <verba> ...' (mis. 'Saya menunjukkan Merasa'
      menjadi 'Saya merasa');
    - mengecilkan huruf awal kata setelah 'Saya'/'Saya menunjukkan' (mis. 'Saya Mampu'
      menjadi 'Saya mampu');
    - memastikan butir diakhiri tepat satu titik.
    """
    t = str(teks or "").strip()
    if not t:
        return t
    # 1) buang frasa penutup berulang di akhir butir
    for suf in _SUFIKS_BOILERPLATE:
        low = t.lower()
        if low.endswith(suf + "."):
            t = t[: len(t) - len(suf) - 1].rstrip()
        elif low.endswith(suf):
            t = t[: len(t) - len(suf)].rstrip()
    # 2) 'Saya menunjukkan <verba>' -> 'Saya <verba>'
    m = re.match(r"^Saya menunjukkan\s+(\S+)(.*)$", t, re.S)
    if m and _VERBA_DEPAN_RE.match(m.group(1).lower()):
        t = "Saya " + m.group(1)[0].lower() + m.group(1)[1:] + m.group(2)
    elif m:
        # 2b) kata pertama setelah 'menunjukkan' dijadikan huruf kecil
        m2 = re.match(r"^Saya menunjukkan\s+([A-Z])(\S*)(.*)$", t, re.S)
        if m2:
            t = "Saya menunjukkan " + m2.group(1).lower() + m2.group(2) + m2.group(3)
    else:
        # 3) 'Saya <Kapital>' -> 'Saya <kecil>' (mis. 'Saya Mampu ...')
        m3 = re.match(r"^Saya\s+([A-Z])(\S*)(.*)$", t, re.S)
        if m3:
            t = "Saya " + m3.group(1).lower() + m3.group(2) + m3.group(3)
    # 4) akhiri tepat satu titik
    t = t.rstrip().rstrip(".") + "."
    return t


def _apply_y_dimension_source_pool(onto_y, y, judul, instansi):
    """Ganti dimensi & sumber Variabel Y dengan sumber dari POOL terverifikasi.

    Tujuan: setiap judul memakai sumber dimensi yang BERBEDA (bukan satu sumber
    yang sama untuk semua judul), sementara dimensi tetap nyata dan operasional.
    Diterapkan SEBELUM kisi-kisi/kuesioner disusun agar benang merah konsisten.
    """
    try:
        nama = (y or onto_y.get("nama_resmi", "") or "").lower()
        if "kinerja" not in nama and "performance" not in nama:
            return onto_y
        import academic_connector as AC
        thesis_key = f"{instansi}_{judul}_{y}".lower()
        entry = AC.select_dimension_source(thesis_key, y_name=y)
        if not entry or not entry.get("dimensi"):
            return onto_y
        onto_y["dimensi"] = [
            {"nama": d.get("nama", ""),
             "indikator": list(d.get("indikator", []) or []),
             "pernyataan": list(d.get("pernyataan", []) or [])}
            for d in entry["dimensi"] if d.get("nama")
        ]
        onto_y["jurnal_rujukan"] = {
            "sitasi": entry.get("sitasi", ""),
            "entri": entry.get("entri", ""),
        }
        onto_y["engine"] = (onto_y.get("engine") or "") + "+dim_pool"
        print(f"[ontology] dimensi Y dari pool: {entry.get('sitasi')} "
              f"({len(onto_y['dimensi'])} dimensi)")
    except Exception as e:
        print(f"[ontology] pool dimensi Y gagal: {e}")
    return onto_y


# ---------------------------------------------------------------------------
# SANITASI DIMENSI & INDIKATOR (anti duplikat, anti indikator sampah)
# ---------------------------------------------------------------------------
# Indikator hasil ekstraksi sumber bisa berupa: (a) angka/statistik hasil
# penelitian, (b) pengulangan definisi dimensi, (c) frasa generik yang tidak
# mencerminkan dimensi, (d) butir kembar. Modul ini membersihkan indikator,
# memastikan TIDAK ADA butir kembar (antar dimensi satu variabel maupun antar
# variabel), lalu menjabarkan ulang butir yang hilang lewat Gemini
# (write_dimension_indicators) dengan daftar indikator terlarang.
# ---------------------------------------------------------------------------


def _rapikan_nama_dimensi(nama):
    """Bikin nama dimensi konsisten (huruf awal kapital bila lowercase semua)."""
    n = str(nama or "").strip()
    if n and n == n.lower():
        return n.title()
    return n


def _is_perusahaan(instansi="", judul=""):
    """True bila objek penelitian adalah perusahaan (bukan instansi pemerintah)."""
    teks = f"{instansi or ''} {judul or ''}"
    return bool(re.search(r"\bPT\b|Persero|Tbk|\bCV\b|\bLtd\b|\bInc\b|Perusahaan",
                          teks, re.I))


def _kontekstualisasi_definisi(teks, instansi="", judul=""):
    """Selaraskan diksi definisi dengan objek tanpa mengubah makna konstruk."""
    t = str(teks or "").strip()
    if not t:
        return t
    t = re.sub(r"\bKetrampilan\b", "Keterampilan", t, flags=re.I)
    # Perbaikan nama objek yang sudah diketahui (typo sumber/entri lama).
    if re.search(r"Bhanda\s+Ghara", instansi or "", re.I):
        t = re.sub(r"Bhanda\s+Ghara\s+Raksa", "Bhanda Ghara Reksa", t, flags=re.I)
    if _is_perusahaan(instansi, judul):
        ganti = [
            (r"\baparatur\b", "karyawan"),
            (r"\bpegawai\b", "karyawan"),
            (r"\bkedinasan\b", "pekerjaan"),
            (r"\bbirokrasi\b", "organisasi"),
            (r"\binstansi\b", "perusahaan"),
            (r"\bdinas\b", "perusahaan"),
        ]
        for pola, baru in ganti:
            t = re.sub(pola, baru, t, flags=re.I)
    return re.sub(r"\s+", " ", t).strip()


def _kontekstualisasi_indikator(teks, instansi="", judul=""):
    """Selaraskan diksi indikator dengan objek (perusahaan vs pemerintah)."""
    t = str(teks or "").strip()
    if not t:
        return t
    t = re.sub(r"\bKetrampilan\b", "Keterampilan", t, flags=re.I)
    if _is_perusahaan(instansi, judul):
        ganti = [
            (r"\baparatur\b", "karyawan"),
            (r"\bpegawai\b", "karyawan"),
            (r"\bkedinasan\b", "pekerjaan"),
            (r"\bbirokrasi\b", "organisasi"),
            (r"\binstansi\b", "perusahaan"),
            (r"\bdinas\b", "perusahaan"),
        ]
        for pola, baru in ganti:
            t = re.sub(pola, baru, t, flags=re.I)
    return re.sub(r"\s+", " ", t).strip()


def _indikator_fallback(var_name, nama_dimensi, dipakai, instansi=""):
    """Indikator cadangan terakhir bila Gemini gagal (tetap menyebut dimensi)."""
    try:
        import llm_background_writer as W
    except Exception:
        return None
    d = str(nama_dimensi or "").strip()
    if not d:
        return None
    ds = d.lower()
    subjek = "pegawai" if not _is_perusahaan(instansi) else "karyawan"
    tempat = "instansi" if not _is_perusahaan(instansi) else "perusahaan"
    kandidat = [
        f"Tingkat {ds} yang ditunjukkan {subjek} dalam menjalankan tugas sehari-hari",
        f"Intensitas {ds} pada pelaksanaan pekerjaan {subjek} di {tempat}",
        f"Peran {ds} terhadap mutu hasil kerja {subjek}",
        f"Kondisi {ds} sebagaimana dirasakan {subjek} dalam pekerjaan",
    ]
    for c in kandidat:
        s, _alasan = W._indikator_generated_boleh(c, d, var_name, dipakai or [])
        if s:
            return s
    return None


def _promote_flat_dimensions(onto):
    """Ubah konstruk datar (1 dimensi, 2-7 butir pendek) menjadi beberapa dimensi.

    Contoh nyata: Training Dessler hanya memberi satu kelompok
    Reaction/Learning/Behavior/Results. Keempatnya adalah level evaluasi
    Kirkpatrick, sehingga layak dipromosikan menjadi dimensi. Indikator
    operasionalnya tetap akan dikembangkan terpisah dari dimensi tersebut.
    """
    dims = (onto or {}).get("dimensi") or []
    if len(dims) != 1:
        return onto
    inds = [str(x).strip() for x in (dims[0].get("indikator") or []) if str(x).strip()]
    if not (2 <= len(inds) <= 7):
        return onto

    def _layak_dim(s):
        if len(s.split()) > 6 or len(s) > 70:
            return False
        if re.search(r"\b(adalah|merupakan|mencakup|mengacu|didefinisikan|diartikan|"
                     r"meliputi|terdiri\s+dari)\b", s, re.I):
            return False
        return True

    if not all(_layak_dim(x) for x in inds):
        return onto
    onto["dimensi"] = [
        {"nama": _kapitalkan_indikator(
            _ISTILAH_INDONESIA.get(x.strip().lower(), x)),
         "indikator": [], "pernyataan": []}
        for x in inds
    ]
    onto["flat_promoted"] = True
    print(f"[ontologi] konstruk datar '{onto.get('nama_resmi', '')}' dipromosikan "
          f"menjadi {len(inds)} dimensi")
    return onto


def _dims_generated_ok(dims):
    """True bila semua dimensi sudah punya 2-4 indikator hasil pengembangan."""
    if not (2 <= len(dims or []) <= 7):
        return False
    for d in dims:
        n = len(d.get("indikator") or [])
        if not (2 <= n <= 4):
            return False
    return True


def _kapitalkan_indikator(teks):
    """Pastikan indikator diawali huruf kapital (kerapian tabel kisi-kisi)."""
    t = str(teks or "").strip()
    if not t:
        return t
    for i, ch in enumerate(t):
        if ch.isalpha():
            return t[:i] + ch.upper() + t[i + 1:]
    return t


def sanitize_dimensi_indikator(var_map, judul="", instansi="", regenerate=True):
    """Bersihkan dimensi & kembangkan indikator operasional dari dimensi.

    Kebijakan ``generated_v1`` (berlaku untuk judul baru):
    - dimensi 2-7, berasal dari sumber/teori terverifikasi;
    - indikator 2-4 per dimensi, dikembangkan dari dimensi + definisi + konteks objek;
    - indikator lama tidak dipertahankan karena dapat berupa definisi, diksi sektor
      publik, atau tidak cocok dengan objek perusahaan.

    ``regenerate=False`` dipakai saat memuat cache yang sudah memakai kebijakan
    ini, agar indikator yang sudah disetujui tidak berubah setiap kali dibuka.
    """
    try:
        import llm_background_writer as W
        import indicator_layers as IL
    except Exception as e:
        print(f"[sanitasi] modul writer tidak tersedia: {e}")
        return []

    dipakai = []          # seluruh indikator final (lintas variabel)
    laporan = []
    _stmt_jobs = []

    # Y diproses lebih dulu agar indikator Y menjadi acuan dedup lintas variabel.
    for role in ("y", "x1", "x2", "m"):
        onto = (var_map or {}).get(role) or {}
        if not onto:
            continue
        var_name = onto.get("nama_resmi") or onto.get("input_name") or ""
        onto["definisi"] = _kontekstualisasi_definisi(
            onto.get("definisi", ""), instansi=instansi, judul=judul)
        onto = _promote_flat_dimensions(onto)
        dims = onto.get("dimensi") or []
        if not dims:
            onto["needs_review"] = True
            laporan.append(f"{role} ({var_name}): 0 dimensi, perlu sumber")
            continue
        # Sumber boleh memuat 2-20 dimensi. Pilih 2-7 yang paling relevan dan
        # catat dimensi yang tidak dipakai, bukan memotong daftar secara buta.
        dims, seleksi = IL.select_dimensions(
            dims, var_name=var_name, definisi=onto.get("definisi", ""),
            instansi=instansi, judul=judul)
        onto["dimensi"] = dims
        onto["dimension_selection"] = seleksi
        if seleksi.get("tidak_dipakai"):
            laporan.append(
                f"{role} ({var_name}): {seleksi['dipilih_total']}/"
                f"{seleksi['sumber_total']} dimensi dipakai; "
                f"{len(seleksi['tidak_dipakai'])} tidak dipakai (lihat metadata)")
        # rapikan nama dimensi (tanpa mengubah makna)
        for d in dims:
            d["nama"] = _rapikan_nama_dimensi(d.get("nama", ""))

        jurnal = onto.get("jurnal_rujukan") or {}
        sumber_label = jurnal.get("sitasi", "") or onto.get("tokoh_ahli", "") or ""

        harus_generate = bool(regenerate) or not _dims_generated_ok(dims)
        if harus_generate:
            nama_need = [d.get("nama", "") for d in dims if d.get("nama")]
            if len(nama_need) != len(dims):
                dims = [d for d in dims if d.get("nama")]
                onto["dimensi"] = dims
                nama_need = [d.get("nama", "") for d in dims]
            try:
                hasil = W.write_dimension_indicators(
                    var_name, nama_need, source_label=sumber_label,
                    tokoh=onto.get("tokoh_ahli", ""), judul=judul,
                    instansi=instansi, forbidden=list(dipakai),
                    definisi=onto.get("definisi", ""))
            except Exception as e:
                hasil = None
                print(f"[sanitasi] generate indikator '{var_name}' gagal: {e}")
            items = (hasil or {}).get("dimensi_list") or []
            for idx, d in enumerate(dims):
                gen = []
                fallback_used = False
                if idx < len(items):
                    for g in (items[idx].get("indikator") or []):
                        g = _kontekstualisasi_indikator(
                            str(g or "").strip(), instansi=instansi, judul=judul)
                        if not g:
                            continue
                        s, _alasan = W._indikator_generated_boleh(
                            g, d.get("nama", ""), var_name, gen + dipakai)
                        if s and not any(W._ind_terlalu_sama(s, c) for c in gen):
                            gen.append(s)
                while len(gen) < 2:
                    cand = _indikator_fallback(
                        var_name, d.get("nama", ""), gen + dipakai,
                        instansi=instansi)
                    if not cand or any(W._ind_terlalu_sama(cand, c) for c in gen):
                        break
                    gen.append(cand)
                    fallback_used = True
                d["indikator"] = [_kapitalkan_indikator(x) for x in gen[:4]]
                d["indikator_operasional"] = list(d["indikator"])
                d["indikator_teoretis"] = IL.bangun_indikator_teoretis(
                    d["indikator"], d.get("nama", ""),
                    instansi=instansi, judul=judul)
                d["pernyataan"] = []
                d["indikator_origin"] = "generated_from_dimension"
                d["dimension_source"] = sumber_label
                if fallback_used:
                    d["indikator_fallback"] = True
                    onto["needs_review"] = True
                if len(d["indikator"]) < 2:
                    onto["needs_review"] = True
                dipakai.extend(d["indikator"])
        else:
            # Cache yang sudah generated: bersihkan tanpa memanggil LLM.
            for d in dims:
                bersih = []
                for x in (d.get("indikator") or []):
                    x = _kontekstualisasi_indikator(
                        str(x or "").strip(), instansi=instansi, judul=judul)
                    s, _alasan = W._indikator_generated_boleh(
                        x, d.get("nama", ""), var_name, bersih + dipakai)
                    if s and not any(W._ind_terlalu_sama(s, c) for c in bersih):
                        bersih.append(s)
                d["indikator"] = [_kapitalkan_indikator(x) for x in bersih[:4]]
                d["indikator_operasional"] = list(d["indikator"])
                d["indikator_teoretis"] = IL.bangun_indikator_teoretis(
                    d["indikator"], d.get("nama", ""),
                    instansi=instansi, judul=judul)
                if len(d["indikator"]) < 2:
                    onto["needs_review"] = True
                dipakai.extend(d["indikator"])

        onto["indicator_policy"] = "generated_v1"
        onto["dimensi"] = dims
        for d in dims:
            # Bersihkan flag fallback bila indikator sudah bukan kalimat template darurat
            if d.get("indikator_fallback"):
                inds = d.get("indikator") or []
                if not any(str(x).startswith(("Tingkat ", "Intensitas ", "Peran ", "Kondisi ")) and "dalam menjalankan tugas" in str(x) for x in inds):
                    d["indikator_fallback"] = False
        any_fallback = any(d.get("indikator_fallback") for d in dims)
        if not (2 <= len(dims) <= 7):
            onto["needs_review"] = True
        elif any(not (2 <= len(d.get("indikator") or []) <= 4) for d in dims):
            onto["needs_review"] = True
        elif any_fallback:
            onto["needs_review"] = True
        else:
            onto["needs_review"] = False

        # Butir kuesioner (Likert) dibuat ulang untuk indikator yang baru.
        butuh_stmt = [(d.get("nama"), list(d.get("indikator") or []))
                      for d in dims if not d.get("pernyataan") and d.get("indikator")]
        if butuh_stmt:
            _stmt_jobs.append((var_name, dims, butuh_stmt, judul, instansi))

        total = sum(len(d.get("indikator") or []) for d in dims)
        laporan.append(
            f"{role} ({var_name}): {len(dims)} dimensi, {total} indikator "
            f"(dikembangkan dari dimensi; 2-4 per dimensi)")

    # --- Tahap 5: buat butir kuesioner semua variabel secara PARALEL ---
    if _stmt_jobs:
        from concurrent.futures import ThreadPoolExecutor

        def _buat_stmt(job):
            _vn, _dims, _butuh, _judul, _inst = job
            try:
                return _vn, _dims, _butuh, W.write_dimension_statements(
                    _vn, _butuh, judul=_judul, instansi=_inst)
            except Exception as e:
                print(f"[sanitasi] butir kuesioner '{_vn}' gagal: {e}")
                return _vn, _dims, _butuh, None

        try:
            with ThreadPoolExecutor(max_workers=min(4, len(_stmt_jobs))) as _pool:
                _hasil_stmt = list(_pool.map(_buat_stmt, _stmt_jobs))
        except Exception as e:
            print(f"[sanitasi] pembuatan butir kuesioner paralel gagal: {e}")
            _hasil_stmt = []

        for _vn, _dims, _butuh, _stmt in _hasil_stmt:
            if not _stmt:
                continue
            for item, (nama_dim, _ind) in zip(_stmt.get("dimensi_list") or [], _butuh):
                for d in _dims:
                    if d.get("nama") == nama_dim:
                        d["pernyataan"] = list(item.get("pernyataan") or [])
                        break

    return laporan


def _susun_instrumen(var_map, judul=""):
    """Susun kisi-kisi Bab III (Tabel 3.1) + butir kuesioner Lampiran."""
    kisi_kisi_rows = []
    item_counter = 1
    for role in ("x1", "x2", "m", "y"):
        onto = (var_map or {}).get(role) or {}
        var_role = onto.get("role") or role.upper()
        var_name = onto.get("nama_resmi", "")
        for d in (onto.get("dimensi") or []):
            dim_name = d.get("nama", "")
            stmts = d.get("pernyataan", []) or []
            for idx, ind in enumerate(d.get("indikator", []) or []):
                kisi_kisi_rows.append({
                    "variabel": f"{var_name} ({var_role})",
                    "dimensi": dim_name,
                    "indikator": ind,
                    "pernyataan_ai": stmts[idx] if idx < len(stmts) else "",
                    "nomor_butir": str(item_counter),
                    "skala": "Likert (1-5)"
                })
                item_counter += 1

    kuesioner_items = []
    for k in kisi_kisi_rows:
        pernyataan = _normalisasi_pernyataan(k.get("pernyataan_ai") or _kuesioner_statement_for(k))
        kuesioner_items.append({
            "no": k["nomor_butir"],
            "variabel": k["variabel"],
            "dimensi": k["dimensi"],
            "indikator": k["indikator"],
            "pernyataan": pernyataan,
            "skala": ["1. Sangat Tidak Setuju (STS)", "2. Tidak Setuju (TS)", "3. Kurang Setuju (KS)", "4. Setuju (S)", "5. Sangat Setuju (SS)"]
        })

    # Pertahankan pernyataan kuesioner yang sudah disempurnakan secara manual
    # (mis. melalui _fix_kuesioner.py) pada judul yang SAMA.
    try:
        prev = baca_ontology_cache(judul)
        if prev:
            prev_map = {}
            for it in (prev.get("kuesioner_lampiran") or []):
                key = (str(it.get("variabel", "")).strip(),
                       str(it.get("indikator", "")).strip().lower())
                prev_map[key] = it.get("pernyataan", "")
            dipakai = 0
            for it in kuesioner_items:
                key = (str(it.get("variabel", "")).strip(),
                       str(it.get("indikator", "")).strip().lower())
                lama = _normalisasi_pernyataan(prev_map.get(key) or "").strip()
                if lama and lama != _kuesioner_statement_for(it):
                    it["pernyataan"] = lama
                    dipakai += 1
            if dipakai:
                print(f"[ontologi] {dipakai} pernyataan kuesioner hasil penyempurnaan dipertahankan")
    except Exception as e:
        print(f"Warning carry-over kuesioner: {e}")

    return kisi_kisi_rows, kuesioner_items


def ensure_indicator_layers(data, judul="", instansi=""):
    """Lengkapi lapisan teoretis/operasional + seleksi dimensi pada cache.

    Aman dipakai pada cache beku: tidak memanggil LLM dan tidak mengubah
    indikator operasional yang sudah disetujui.
    """
    try:
        import indicator_layers as IL
        return IL.pastikan_layer(data, judul=judul, instansi=instansi)
    except Exception as e:
        print(f"[indikator] pelengkapan lapisan gagal: {e}")
        return False


def sanitize_ontology_cache(data, judul="", instansi=""):
    """Sanitasi blueprint ontologi yang dibaca dari cache (in place).

    Dipakai saat ``use_ontology_cache=True`` atau saat Bab IV/Lampiran memuat
    ontologi tersimpan: memastikan indikator tetap bersih, unik, dan kisi-kisi
    serta kuesioner rebuilt mengikuti indikator terbaru.
    """
    if not isinstance(data, dict):
        return data
    variables = data.get("variables") or {}
    if not variables:
        return data
    meta = data.get("metadata") or {}
    judul = judul or meta.get("judul", "") or ""
    instansi = instansi or meta.get("instansi", "") or ""
    # Cache baru yang sudah memakai generated_v1 tidak boleh memanggil LLM lagi;
    # cukup dibersihkan dan disusun ulang kisi-kisinya.
    _regen = (meta.get("indicator_policy") != "generated_v1")
    if not _regen:
        for _v in variables.values():
            if _v and not _dims_generated_ok(_v.get("dimensi") or []):
                _regen = True
                break
    try:
        for _baris in sanitize_dimensi_indikator(
                variables, judul=judul, instansi=instansi, regenerate=_regen):
            print(f"[sanitasi-cache] {_baris}")
    except Exception as e:
        print(f"[sanitasi-cache] gagal: {e}")
        return data
    try:
        ensure_indicator_layers(data, judul=judul, instansi=instansi)
        kisi_kisi_rows, kuesioner_items = _susun_instrumen(variables, judul=judul)
        data["kisi_kisi_bab3"] = kisi_kisi_rows
        data["kuesioner_lampiran"] = kuesioner_items
        data.setdefault("metadata", {})["total_indikator"] = len(kisi_kisi_rows)
        data.setdefault("metadata", {})["indicator_policy"] = "generated_v1"
        data.setdefault("metadata", {})["needs_review"] = [
            v.get("nama_resmi") or v.get("input_name")
            for v in variables.values() if v.get("needs_review")
        ]
        # Tandai beku hanya bila semua variabel lolos (tanpa needs_review)
        data.setdefault("metadata", {})["frozen"] = (len(data["metadata"]["needs_review"]) == 0)
    except Exception as e:
        print(f"[sanitasi-cache] instrumen gagal disusun ulang: {e}")
    return data


def build_full_thesis_ontology(x1, x2, m, y, instansi, sektor, judul):
    """
    Membangun Blueprint Ontologi Tesis Utuh (Master Research Blueprint).
    Menghubungkan 4 variabel ke fenomena Bab I, kajian Bab II, kisi-kisi Bab III, dan Kuesioner.
    """
    # Riset 4 variabel dijalankan PARALEL (masing-masing bottleneck-nya adalah
    # panggilan jaringan ke Gemini/OpenAlex, bukan CPU). Rata-rata hemat ~40% waktu.
    from concurrent.futures import ThreadPoolExecutor

    _specs = [("x1", x1, "X1"), ("x2", x2, "X2"), ("m", m, "M"), ("y", y, "Y")]
    _out = {}

    def _riset(spek):
        _k, _nama, _role = spek
        return _k, research_variable_ontology(_nama, role=_role, judul=judul,
                                             instansi=instansi)

    try:
        with ThreadPoolExecutor(max_workers=4) as _pool:
            for _k, _res in _pool.map(_riset, _specs):
                _out[_k] = _res
    except Exception as e:
        print(f"[ontology] riset paralel gagal ({e}); fallback ke sekuensial")
        _out = {}

    if len(_out) < 4:
        # Fallback sekuensial bila ada worker yang gagal total
        for _k, _nama, _role in _specs:
            if _k not in _out or not isinstance(_out.get(_k), dict):
                try:
                    _out[_k] = research_variable_ontology(
                        _nama, role=_role, judul=judul, instansi=instansi)
                except Exception as e:
                    print(f"[ontology] riset '{_nama}' gagal: {e}")

    onto_x1 = _out.get("x1") or {}
    onto_x2 = _out.get("x2") or {}
    onto_m  = _out.get("m") or {}
    onto_y  = _out.get("y") or {}
    # Sumber & dimensi Y diambil dari POOL agar BERBEDA tiap judul.
    onto_y = _apply_y_dimension_source_pool(onto_y, y, judul, instansi)

    # 0. SANITASI + PENGEMBANGAN INDIKATOR (kebijakan generated_v1):
    #    dimensi 2-7 dari sumber; indikator 2-4 per dimensi dikembangkan dari
    #    dimensi + definisi + konteks objek. Indikator lama tidak dipertahankan
    #    bila berupa definisi/diksi sektor publik yang tidak cocok.
    var_map = {"x1": onto_x1, "x2": onto_x2, "m": onto_m, "y": onto_y}
    try:
        for _baris in sanitize_dimensi_indikator(
                var_map, judul=judul, instansi=instansi, regenerate=True):
            print(f"[sanitasi] {_baris}")
    except Exception as e:
        print(f"[sanitasi] gagal: {e}")

    # 1. Ekstraksi Indikator Empiris untuk Tabel 1.1 (Diambil dari Indikator Variabel Y)
    y_indicators_flat = []
    for d in onto_y["dimensi"]:
        for ind in d["indikator"]:
            y_indicators_flat.append(f"{d['nama']}: {ind}")
    
    # Ambil 3 indikator utama untuk baris Tabel 1.1 (aman bila Y belum diriset AI)
    def _first_ind(onto, idx=0, default="menyesuaikan hasil riset"):
        try:
            dims = onto["dimensi"]
            if idx < len(dims) and dims[idx].get("indikator"):
                return dims[idx]["indikator"][0]
        except (IndexError, KeyError, TypeError):
            pass
        return default

    tabel_1_1_params = [
        f"Kualitas Kinerja ({_first_ind(onto_y, 0, 'Melalui indikator Y hasil riset')})",
        f"Kuantitas Output ({_first_ind(onto_y, 1, 'Target Rutin')})",
        f"Ketepatan Waktu ({_first_ind(onto_y, 2, 'Penyelesaian Tugas')})"
    ]

    # 2. Ekstraksi Indikator Masalah Operasional untuk Tabel 1.2 (Diambil dari Dimensi Variabel M)
    m_dimensions_flat = [d["nama"] for d in onto_m["dimensi"]]

    # 3 + 4. Penyusunan Kisi-Kisi Instrumen Bab III (Tabel 3.1) & butir kuesioner
    kisi_kisi_rows, kuesioner_items = _susun_instrumen(var_map, judul=judul)

    # Variabel yang masih perlu konfirmasi pengguna (jangan diam-diam dipalsukan)
    needs_review_vars = [v["input_name"] for v in var_map.values() if v.get("needs_review")]

    master_blueprint = {
        "status": "success",
        "metadata": {
            "judul": judul,
            "instansi": instansi,
            "sektor": sektor,
            "total_indikator": len(kisi_kisi_rows),
            "skala": "Likert 1-5",
            "frozen": True,
            "indicator_policy": "generated_v1",
            "indicator_policy_note": (
                "Dimensi berasal dari sumber/teori terverifikasi; indikator "
                "operasional (2-4 per dimensi) dikembangkan berdasarkan dimensi "
                "dan konteks objek penelitian."),
            "engines": {k: v.get("engine", "unknown") for k, v in var_map.items()},
            "needs_review": needs_review_vars,
        },
        "variables": {
            "x1": onto_x1,
            "x2": onto_x2,
            "m": onto_m,
            "y": onto_y
        },
        "phenomenon_anchors": {
            "tabel_1_1_params": tabel_1_1_params,
            "tabel_1_2_issues": m_dimensions_flat,
            "x1_focus": onto_x1["dimensi"][0]["nama"] if onto_x1["dimensi"] else "(menunggu riset AI)",
            "x2_focus": onto_x2["dimensi"][0]["nama"] if onto_x2["dimensi"] else "(menunggu riset AI)"
        },
        "kisi_kisi_bab3": kisi_kisi_rows,
        "kuesioner_lampiran": kuesioner_items
    }

    # Lengkapi lapisan teoretis/operasional dan metadata seleksi dimensi.
    try:
        ensure_indicator_layers(master_blueprint, judul=judul, instansi=instansi)
    except Exception as e:
        print(f"Warning pelengkapan lapisan indikator: {e}")

    # Simpan ke cache lokal KHUSUS JUDUL INI (jangan menimpa judul lain)
    try:
        simpan_ontology_cache(master_blueprint, judul)
    except Exception as e:
        print(f"Warning saving ontology cache: {e}")

    return master_blueprint


def load_thesis_ontology(judul=""):
    """Memuat ontologi tesis. Bila judul diberikan, baca berkas judul tersebut."""
    if judul:
        data = baca_ontology_cache(judul)
        if data:
            return data
    if os.path.exists(ONTOLOGY_CACHE_FILE):
        try:
            with open(ONTOLOGY_CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading ontology cache: {e}")
    # Default fallback jika belum pernah diriset
    test_title = "PENGARUH ORGANIZATIONAL CITIZENSHIP BEHAVIOR DAN MOTIVASI KERJA TERHADAP KINERJA PEGAWAI MELALUI KEPUASAN KERJA SEBAGAI VARIABEL INTERVENING PADA BPKAD KABUPATEN OKI"
    parsed = parse_academic_title(test_title)
    return build_full_thesis_ontology(
        parsed["x1"], parsed["x2"], parsed["m"], parsed["y"],
        parsed["instansi"], parsed["sektor"], test_title
    )


def research_thesis_ontology(title="", config_override=None):
    """
    Melakukan riset ontologi otomatis berbasis judul atau konfigurasi override.
    """
    config_override = config_override or {}
    parsed = {}
    if title:
        # Lapis 1: Gemini memahami judul apa pun bentuknya; Lapis 2: regex
        try:
            import llm_title_parser
            llm_parsed = llm_title_parser.parse_title_llm(title)
        except Exception:
            llm_parsed = None
        parsed = llm_parsed if llm_parsed else parse_academic_title(title)
    
    x1 = config_override.get("x1") or parsed.get("x1") or "Organizational Citizenship Behavior"
    x2 = config_override.get("x2") or parsed.get("x2") or "Motivasi Kerja"
    m  = config_override.get("m")  or parsed.get("m")  or "Kepuasan Kerja"
    y  = config_override.get("y")  or parsed.get("y")  or "Kinerja Pegawai"
    instansi = config_override.get("instansi") or parsed.get("instansi") or "BPKAD Kabupaten OKI"
    sektor = config_override.get("sektor") or parsed.get("sektor") or "Pengelolaan Keuangan dan Aset Daerah"
    effective_title = title or config_override.get("judul") or f"PENGARUH {x1.upper()} DAN {x2.upper()} TERHADAP {y.upper()} MELALUI {m.upper()} SEBAGAI VARIABEL INTERVENING PADA {instansi.upper()}"

    return build_full_thesis_ontology(x1, x2, m, y, instansi, sektor, effective_title)


if __name__ == "__main__":
    test_title = "PENGARUH ORGANIZATIONAL CITIZENSHIP BEHAVIOR DAN MOTIVASI KERJA TERHADAP KINERJA PEGAWAI MELALUI KEPUASAN KERJA SEBAGAI VARIABEL INTERVENING PADA BPKAD KABUPATEN OKI"
    parsed = parse_academic_title(test_title)
    onto = build_full_thesis_ontology(
        parsed["x1"], parsed["x2"], parsed["m"], parsed["y"],
        parsed["instansi"], parsed["sektor"], test_title
    )
    print("\n=== BLUEPRINT ONTOLOGI RISET TERBENTUK ===")
    print(f"Judul: {onto['metadata']['judul']}")
    print(f"Total Indikator Kisi-kisi: {onto['metadata']['total_indikator']} butir")
    print(f"Parameter Tabel 1.1 (Y): {onto['phenomenon_anchors']['tabel_1_1_params']}")
    print(f"Fokus Masalah Tabel 1.2 (M): {onto['phenomenon_anchors']['tabel_1_2_issues']}")
    print(f"Contoh Kuesioner Butir 1: {onto['kuesioner_lampiran'][0]['pernyataan']}")
