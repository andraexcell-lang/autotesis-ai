# backend/objek_sterilisasi.py
"""
Sterilisasi istilah OBJEK PENELITIAN dari naskah BAB II (Kajian Teoritis).

Aturan (sesuai kebutuhan pengguna):
  - BAB II bersifat UMUM/teoretis: TIDAK boleh menyebut nama instansi, kota,
    unit kerja, maupun konteks sektor objek (mis. "pasien", "rekam medis",
    "gudang", "pengiriman barang", "nasabah", "siswa", "aparatur").
  - Indikator BAB II harus PENDEK (frasa benda 1-4 kata), contoh:
      "Ketelitian pegawai dalam melakukan pemeriksaan rekam medis pasien
       sesuai standar operasional"  ->  "Ketelitian pegawai"
      "Kesesuaian hasil tindakan asuhan keperawatan dengan prosedur
       rumah sakit islam"           ->  "Kesesuaian hasil"
      "Kerapian penulisan dokumentasi medis dan laporan pelayanan
       kesehatan harian"            ->  "Kerapian"

Modul ini sengaja tanpa LLM dan tanpa dependensi luar agar aman dipakai
dari engine, penulis latar belakang, maupun penulis BAB II.
"""

import re

# ---------------------------------------------------------------------------
# Daftar kata/frasa sektor yang TIDAK pernah sah di kajian teoretis umum.
# Diterapkan pada SEMUA judul (aman: istilah ini tidak mungkin muncul sebagai
# konstruk teoretis BAB II variabel manajemen).
# ---------------------------------------------------------------------------
_FRASA_OBJEK = [
    "rumah sakit", "rekam medis", "asuhan keperawatan", "rawat inap",
    "ruang inap", "rawat jalan", "rawat intensif", "pelayanan kesehatan",
    "tenaga medis", "bongkar muat", "berita acara", "surat jalan",
    "pengiriman barang", "inventaris barang",
]

_KATA_OBJEK = [
    # identitas lembaga
    "pt", "persero", "tbk", "cabang", "perusahaan", "instansi", "yayasan",
    "bumn", "pdam", "perumda", "koperasi", "bank", "spbu",
    "rs", "rsud", "rsia", "rsu",
    # kesehatan
    "pasien", "rekam", "medis", "asuhan", "keperawatan", "perawat", "dokter",
    "klinik", "puskesmas", "obat", "farmasi", "perawatan", "rawat", "klinis",
    # pergudangan / logistik / ekspedisi
    "gudang", "pergudangan", "logistik", "kargo", "bongkar", "muat",
    "pengiriman", "inventaris", "persediaan", "stok", "armada", "ekspedisi",
    "kurir", "supplier", "vendor", "mitra",
    # perbankan / keuangan
    "nasabah", "rekening", "tabungan", "atm", "debitur",
    # pendidikan
    "siswa", "murid", "pelajar", "mahasiswa", "dosen", "guru", "kurikulum",
    # pemerintah / aparatur
    "aparatur", "kelurahan", "kecamatan", "desa", "pns",
    "bpkad", "bkpsdm", "bappeda", "dispenda", "disdik", "dinkes", "dishub", "setda",
]

# Nama kota/wilayah umum (tetap diblok walau tidak ada di field instansi).
_KOTA = [
    "palembang", "jakarta", "bandung", "surabaya", "medan", "semarang",
    "makassar", "pekanbaru", "jambi", "lampung", "padang", "bengkulu",
    "batam", "balikpapan", "banjarmasin", "pontianak", "manado", "denpasar",
    "yogyakarta", "solo", "malang", "bogor", "depok", "tangerang", "bekasi",
    "sumatera", "kalimantan", "sulawesi", "papua", "jawa",
]

# Kata yang TIDAK diambil sebagai token nama entitas dari field instansi.
_STOP_TOKEN = {
    "pt", "persero", "perseroan", "tbk", "kabupaten", "kota", "provinsi",
    "daerah", "pusat", "rumah", "sakit", "kristen", "katolik",
    "protestan", "budha", "hindu", "konghucu", "rs", "rsud", "rsia", "rsu",
    "badan", "dinas", "kantor", "lembaga", "institut", "universitas",
    "politeknik", "sekolah", "tinggi", "negeri", "pemerintah", "swasta",
    "yayasan", "bank", "bpr", "koperasi", "perumda", "pdam", "camat",
    "lurah", "jalan", "jl", "no", "nomor", "gedung", "lantai", "the", "of",
    "and", "for", "in", "on", "an", "a", "to", "di", "dan", "yang", "pada",
    # Istilah umum keilmuan manajemen, pemerintahan, & administrasi (bukan entitas khusus)
    "pengembangan", "kepegawaian", "pelatihan", "pendidikan", "pelayanan",
    "pengelolaan", "keuangan", "perencanaan", "sumber", "daya", "manusia",
    "administrasi", "aset", "kinerja", "kompetensi", "organisasi", "kebijakan",
    "aparatur", "teknologi", "komunikasi", "informasi", "kesehatan", "sosial",
    "lingkungan", "pekerjaan", "umum", "perhubungan", "pendapatan", "inspektorat",
    "sekretariat", "tata", "usaha", "hukum",
}

# Preposisi yang ikut terhapus bila berada PERSIS sebelum istilah objek
# (mis. "di gudang", "pada pasien", "dari logistik").
_PREP_IKUT = r"(?:\b(?:di|ke|pada|dalam|untuk|dengan|dari|terhadap|bagi|oleh|kepada)\s+)?"

# Konjungsi/preposisi sisa yang dibersihkan bila menggantung (beruntun atau
# menjelang tanda baca / akhir kalimat) setelah istilah objek dihapus.
_RX_KONJ_SISA = re.compile(
    r"\s+\b(?:di|ke|pada|dalam|untuk|dengan|dari|terhadap|bagi|oleh|yang|"
    r"dan|atau|serta|adalah|ialah|kepada)\s+(?=\b(?:di|ke|pada|dalam|untuk|dengan|"
    r"dari|terhadap|bagi|oleh|yang|dan|atau|serta|adalah|ialah|kepada)\b|[,.;:!?]|$)",
    re.I)

_KATA_OBJEK += _KOTA
_KATA_OBJEK = sorted(set(_KATA_OBJEK), key=len, reverse=True)

# Titik potong indikator: connector, verba kerja, atau kata proses yang
# menandai awal deskripsi panjang. Bagian SEBELUM titik ini dipertahankan.
_RX_POTONG = re.compile(
    r"\b(?:dalam|untuk|dengan|pada|terhadap|sesuai|yang|melalui|sebagai|"
    r"bagi|oleh|dan|serta|karena|hingga|yaitu|antara|tanpa|di|ke|dari|"
    r"secara|akibat|saat|ketika|seperti|berupa|"
    r"menyelesaikan|melakukan|mengisi|menerima|mengoptimalkan|memenuhi|"
    r"menunjukkan|melaksanakan|mengurus|mengerjakan|bekerja|memeriksa|"
    r"mengawasi|mengecek|menjaga|memastikan|menyusun|membuat|mengelola|"
    r"menghadapi|menjalankan|menguji|mengambil|mempertahankan|"
    r"meminimalisir|mengatasi|memandang|menilai|menganggap|merasakan|"
    r"menikmati|memperoleh|mengerahkan|menguras|menuntaskan|memberikan|"
    r"menangani|mencapai|berkoordinasi|membantu|berkomunikasi|berdiri|"
    r"mengejar|mengatur|datang|mematuhi|"
    r"pencatatan|penulisan|dokumentasi|pelaporan|penyelesaian|pelaksanaan|"
    r"pemeriksaan|penanganan|penyusunan|pengisian|penggunaan|"
    r"penyediaan|pemberian)\b",
    re.I)

_MAX_KATA = 4

_PREP_UJUNG = {
    "di", "ke", "pada", "dalam", "untuk", "dengan", "dari", "terhadap",
    "bagi", "oleh", "sebagai", "yang", "dan", "atau", "serta", "akan",
    "secara", "saat", "akibat", "seperti", "berupa", "atas", "tentang",
    "mengenai", "tanpa", "hingga", "sampai", "guna", "demi", "karena",
    "sudah", "telah", "masih", "juga", "sangat", "lebih", "paling",
    "para", "sesama", "antar",
}

# Verba di awal indikator operasional ("Menilai tingkat gaji" -> "Tingkat
# gaji") supaya lapisan teoretis Bab II tetap berupa frasa benda.
_VERBA_AWAL = {
    "menilai", "memandang", "menganggap", "merasakan", "menikmati",
    "memperoleh", "mendapatkan", "menghadapi", "mengerahkan", "menguras",
    "menuntaskan", "memberikan", "menangani", "mencapai", "meminimalisir",
    "mengatasi", "berkoordinasi", "membantu", "berkomunikasi", "mematuhi",
    "menyelesaikan", "melakukan", "menggunakan", "mengikuti", "menjaga",
    "memastikan", "menyusun", "membuat", "memeriksa", "mengawasi",
    "mengecek", "mengelola", "mengisi", "menerima", "memenuhi",
    "mempertahankan", "menjalankan", "menguji", "mengambil", "mengejar",
    "mengatur", "meningkatkan", "mengembangkan", "mengoptimalkan",
}

# ---------------------------------------------------------------------------
# Kompilasi regex objek (di-cache per instansi agar cepat).
# ---------------------------------------------------------------------------
_cache_rx = {}
_cache_kata = {}


def _token_entitas(instansi=""):
    """Token kata-kunci dari nama instansi (mis. 'Bhanda', 'Ghara', 'Reksa',
    'Siti', 'Khadijah', 'Palembang') yang wajib disterilkan dari BAB II."""
    s = re.sub(r"[^a-zA-Z]+", " ", str(instansi or ""))
    out = set()
    for t in s.split():
        tl = t.lower()
        if len(tl) >= 3 and tl not in _STOP_TOKEN:
            out.add(tl)
    return out


# Turunan kata sifat dari token entitas yang tetap menunjuk objek penelitian.
# Contoh: instansi "RS Islam ..." -> "Islami"/"Islamiah" juga tidak boleh
# muncul di BAB II (tetapi "Islam" sebagai nama teori TIDAK diblok global,
# agar tesis bertema etos kerja Islam di instansi lain tetap aman).
_VARIAN_ENTITAS = {
    "islam": ("islami", "islamiah", "islamiyah", "keislaman"),
    "kristen": ("kristiani", "kekristenan"),
    "katolik": ("katolik", "katolikisme"),
    "hindu": ("hindu", "kehinduan"),
}


def daftar_kata_objek(instansi=""):
    """Set kata terlarang = kata sektor global + token nama entitas instansi."""
    instansi = str(instansi or "")
    if instansi not in _cache_kata:
        token = _token_entitas(instansi)
        varian = set()
        for t in token:
            varian.update(_VARIAN_ENTITAS.get(t, ()))
        _cache_kata[instansi] = frozenset(_KATA_OBJEK) | frozenset(token) | frozenset(varian)
    return _cache_kata[instansi]


def _regex_objek(instansi=""):
    instansi = str(instansi or "")
    if instansi not in _cache_rx:
        # Frasa panjang dicocokkan lebih dulu; tiap istilah boleh diawali
        # preposisi (di/pada/dari/...) yang ikut terhapus.
        pola = _PREP_IKUT + r"\b(?:" + "|".join(
            re.escape(k).replace(r"\ ", r"\s+") for k in sorted(
                _FRASA_OBJEK, key=len, reverse=True)) + r")\b"
        kata = sorted(daftar_kata_objek(instansi), key=len, reverse=True)
        if kata:
            pola += r"|" + _PREP_IKUT + r"\b(?:" + "|".join(
                re.escape(k) for k in kata) + r")\b"
        _cache_rx[instansi] = re.compile(pola, re.I)
    return _cache_rx[instansi]


def mengandung_objek(teks, instansi=""):
    """True bila teks memuat istilah objek/sektor/entitas penelitian."""
    t = str(teks or "")
    if not t.strip():
        return False
    if _regex_objek(instansi).search(t):
        return True
    tl = t.lower()
    for f in _FRASA_OBJEK:
        if f in tl:
            return True
    return False


def bersihkan_objek(teks, instansi=""):
    """Buang seluruh istilah objek/sektor dari teks paragraf.

    Menghapus kata terlarang (bersama preposisi di depannya) di posisi mana
    pun lalu merapikan spasi dan konjungsi menggantung agar kalimat tetap
    terbaca.
    """
    t = str(teks or "")
    if not t.strip():
        return t
    t = _regex_objek(instansi).sub(" ", t)
    t = re.sub(r"\s+", " ", t)
    # Bersihkan konjungsi/preposisi menggantung (beruntun atau menjelang tanda baca).
    for _ in range(4):
        baru = _RX_KONJ_SISA.sub(" ", t)
        if baru == t:
            break
        t = baru
    t = re.sub(r"\s+([,.;:!?])", r"\1", t)
    t = re.sub(r"([,.;:!?])\s+([,.;:!?])", r"\1\2", t)
    t = re.sub(r"\s{2,}", " ", t).strip(" ,;:-")
    return t


_DERAJAT_AWAL = {
    "minimalnya", "tingginya", "rendahnya", "banyaknya", "kurangnya",
    "adanya", "besarnya", "kecilnya", "minimnya", "maksimalnya",
}


def pendekkan_indikator(teks, instansi=""):
    """Ubah indikator menjadi frasa benda PENDEK (maks 4 kata, umum).

    Urutan:
      1) buang istilah objek/sektor,
      2) potong di titik potong pertama (preposisi/verba/kata proses),
      3) batasi 4 kata,
      4) buang preposisi menggantung di ujung.
    Hasil 1 kata pun sah (mis. "Kerapian").
    """
    t = str(teks or "").strip()
    if not t:
        return ""
    t_clean = bersihkan_objek(t, instansi)
    # Buang verba di awal (maks 2) agar hasilnya frasa benda, bukan kalimat.
    for _ in range(2):
        kata = t_clean.split()
        if kata and kata[0].lower() in _VERBA_AWAL:
            t_clean = " ".join(kata[1:])
        else:
            break
    m = _RX_POTONG.search(t_clean)
    if m:
        if m.start() == 0:
            first_w_end = len(t_clean.split()[0])
            m2 = _RX_POTONG.search(t_clean, first_w_end)
            t_cut = t_clean[: m2.start()] if m2 else t_clean
        else:
            t_cut = t_clean[: m.start()]
    else:
        t_cut = t_clean
    t_cut = re.sub(r"\s+", " ", t_cut).strip(" ,;:-")
    kata = t_cut.split()

    # Guardrail kata derajat menggantung (-nya):
    # Bila hasil pemotongan hanya 1 kata penanda derajat (mis. 'Minimalnya'),
    # sertakan kata berikutnya dari teks asal agar tidak menggantung tanpa kata benda inti.
    if len(kata) == 1 and (kata[0].lower() in _DERAJAT_AWAL or kata[0].lower().endswith("nya")):
        raw_words = re.sub(r"\s+", " ", t_clean).strip().split()
        found_idx = -1
        for idx_w, w in enumerate(raw_words):
            if w.lower() == kata[0].lower():
                found_idx = idx_w
                break
        if found_idx >= 0 and found_idx + 1 < len(raw_words):
            kata = raw_words[found_idx:min(found_idx + 3, len(raw_words))]
            while kata and kata[-1].lower() in _PREP_UJUNG:
                kata.pop()

    if len(kata) > _MAX_KATA:
        kata = kata[:_MAX_KATA]
    while kata and kata[-1].lower() in _PREP_UJUNG:
        kata.pop()
    if not kata:
        return ""
    hasil = " ".join(kata)
    for i, ch in enumerate(hasil):
        if ch.isalpha():
            return hasil[:i] + ch.upper() + hasil[i + 1:]
    return hasil
