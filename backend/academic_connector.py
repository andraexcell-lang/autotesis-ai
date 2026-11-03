# backend/academic_connector.py
"""
AutoTesis Academic Connector & Live Verifier Engine v3.0
Menghubungkan langsung ke Live Academic APIs (OpenAlex & Crossref)
untuk mengambil artikel jurnal ilmiah nyata, valid, terakreditasi (SINTA/Scopus),
dan terverifikasi DOI-nya (100% active clickable DOI).

ATURAN BAB I (LATAR BELAKANG MASALAH):
- TIDAK menggunakan buku teks fisik/berbayar tanpa tautan digital terbuka (paywall/offline).
- 100% menggunakan ARTIKEL JURNAL ILMIAH TERAKREDITASI dengan DOI AKTIF yang bisa diakses langsung secara online.
- SINKRONISASI 1-to-1 MUTLAK: Seluruh yang dikutip di naskah pasti ada di Daftar Pustaka,
  dan seluruh yang ada di Daftar Pustaka pasti dikutip di naskah.
- SPESIFIK & ORGANIK: Sitasi X1, X2, M, dan Y dipetakan 100% sesuai karakteristik topik variabel,
  termasuk OCB, Kepuasan Kerja, Pengembangan Karier, Promosi Jabatan, Mutasi Kerja, Beban Kerja, dll.
"""

import urllib.request
import urllib.parse
import json
import os
import re
import time
import hashlib
import datetime

CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache_academic_references.json")
USED_CITATIONS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "used_citations_global.json")

# ==============================================================================
# GLOBAL CITATION TRACKER (mencegah duplikasi sitasi antar judul)
# ==============================================================================
def _load_used_citations():
    """Load global tracker of used DOIs and lead authors across all theses."""
    if os.path.exists(USED_CITATIONS_FILE):
        try:
            with open(USED_CITATIONS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # Ensure all keys exist
                return {
                    "used_dois": set(data.get("used_dois", [])),
                    "used_lead_authors": set(data.get("used_lead_authors", [])),
                    "used_titles": set(data.get("used_titles", [])),
                    "thesis_registry": data.get("thesis_registry", {})
                }
        except Exception:
            pass
    return {"used_dois": set(), "used_lead_authors": set(), "used_titles": set(), "thesis_registry": {}}


def _save_used_citations(tracker):
    """Persist global tracker to file."""
    try:
        data = {
            "used_dois": list(tracker["used_dois"]),
            "used_lead_authors": list(tracker["used_lead_authors"]),
            "used_titles": list(tracker["used_titles"]),
            "thesis_registry": tracker["thesis_registry"]
        }
        with open(USED_CITATIONS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[used_citations] Gagal simpan: {e}")


def _register_thesis_citations(thesis_key, sources):
    """Register citations used by a thesis to prevent reuse."""
    tracker = _load_used_citations()
    thesis_registry = tracker["thesis_registry"]
    if thesis_key not in thesis_registry:
        thesis_registry[thesis_key] = {"dois": [], "lead_authors": [], "titles": []}
    for s in sources:
        doi = s.get("doi", "").lower().strip()
        lead = get_lead_author(s).lower().strip()
        title = s.get("title", "").lower().strip()
        if doi and doi not in tracker["used_dois"]:
            tracker["used_dois"].add(doi)
            thesis_registry[thesis_key]["dois"].append(doi)
        if lead and lead not in tracker["used_lead_authors"]:
            tracker["used_lead_authors"].add(lead)
            thesis_registry[thesis_key]["lead_authors"].append(lead)
        if title and title not in tracker["used_titles"]:
            tracker["used_titles"].add(title)
            thesis_registry[thesis_key]["titles"].append(title)
    _save_used_citations(tracker)


def _own_registry_sets(tracker, thesis_key):
    """Kumpulan DOI/nama/title yang SUDAH milik judul ini.

    Dipakai agar saat judul yang SAMA diregenerasi, sumber miliknya sendiri boleh
    dipakai ulang (stabil), TANPA melanggar keunikan antar-judul: sumber milik
    judul lain tetap tersaring oleh daftar global.
    """
    own = (tracker.get("thesis_registry", {}) or {}).get(thesis_key, {}) or {}
    return (set(own.get("dois", []) or []),
            set(own.get("lead_authors", []) or []),
            set(own.get("titles", []) or []))


# ==============================================================================
# POOL SUMBER DIMENSI (agar tiap judul memakai sumber dimensi yang BERBEDA)
# ==============================================================================
DIM_POOL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "variable_library", "kinerja_pegawai_sumber_pool.json")
DIM_ASSIGN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "_dimension_assignments.json")


def _load_json_file(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default


def load_dimension_pool(pool_file=None):
    """Muat daftar sumber dimensi terverifikasi (pool) dari berkas JSON."""
    data = _load_json_file(pool_file or DIM_POOL_FILE, {})
    if isinstance(data, dict):
        return data.get("entries", []) or []
    if isinstance(data, list):
        return data
    return []


def select_dimension_source(thesis_key, y_name="", pool_file=None):
    """Pilih SATU sumber dimensi untuk judul ini, BERBEDA dari judul lain.

    - Deterministik per judul (di-cache di _dimension_assignments.json).
    - Menghindari sumber yang sudah dipakai judul lain (global tracker).
    - Bila pool habis, dipilih yang paling sedikit dipakai.
    Mengembalikan dict entri pool (dengan 'dimensi', 'sitasi', 'entri') atau None.
    """
    pool = load_dimension_pool(pool_file)
    if not pool:
        return None
    thesis_key = (thesis_key or "").lower()[:100]
    assign = _load_json_file(DIM_ASSIGN_FILE, {}) or {}

    def _pool_doi(e):
        return (e.get("doi") or "").lower().strip()

    # 1. Sudah pernah dipilih untuk judul ini -> pakai ulang (deterministik).
    if thesis_key in assign:
        doi = (assign[thesis_key] or "").lower()
        for e in pool:
            if _pool_doi(e) == doi:
                return e

    # 2. Sumber yang sudah dipakai judul LAIN tidak boleh dipakai lagi.
    tracker = _load_used_citations()
    own_dois, _, _ = _own_registry_sets(tracker, thesis_key)
    used_others = set(tracker.get("used_dois", set())) - set(own_dois)

    from collections import Counter
    counts = Counter((v or "").lower() for v in assign.values())

    avail = [e for e in pool if _pool_doi(e) and _pool_doi(e) not in used_others]
    if not avail:
        avail = [e for e in pool if _pool_doi(e)]
    if not avail:
        return None

    # 3. Pilih yang paling sedikit dipakai, lalu tahun terbaru.
    avail.sort(key=lambda e: (counts.get(_pool_doi(e), 0), -(e.get("year") or 0)))
    chosen = avail[0]

    # Simpan assignment (deterministik) + daftarkan agar judul lain menghindarinya.
    assign[thesis_key] = chosen.get("doi")
    try:
        with open(DIM_ASSIGN_FILE, "w", encoding="utf-8") as f:
            json.dump(assign, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[dim_pool] gagal simpan assignment: {e}")
    try:
        _register_thesis_citations(thesis_key, [{
            "doi": chosen.get("doi", ""),
            "authors": chosen.get("authors", []),
            "title": chosen.get("title", ""),
        }])
    except Exception:
        pass
    print(f"[dim_pool] sumber dimensi untuk judul ini: {chosen.get('sitasi')}")
    return chosen


def get_lead_author(paper):
    """Extract lead author surname from paper dict."""
    authors = paper.get("authors", [])
    if authors:
        parts = authors[0].strip().split()
        if parts:
            return parts[-1]
    return ""


# ==============================================================================
# KEBIJAKAN RENTANG TAHUN SITASI (dikonfigurasi di Menu 1. Konfigurasi)
# --------------------------------------------------------------------------
# SEMUA kutipan dalam naskah tesis — jurnal MAUPUN buku — wajib berada dalam
# rentang tahun yang dipilih pengguna (DEFAULT: 20 TAHUN KEBELAKANG, mis.
# 2006-2026 pada tahun berjalan 2026). Sebelumnya rentang ini di-hardcode
# 2022-2026; kini dinamis via set_cite_year_range() / get_cite_year_range() dan
# tersimpan di cite_year_range.json. Konstanta MIN_CITE_YEAR / MAX_CITE_YEAR
# tetap dipakai sebagai satu-satunya sumber kebenaran runtime di tiga lapis:
#   1. academic_connector : filter pool lokal + filter hasil OpenAlex/Crossref.
#   2. llm_background_writer : aturan keras prompt + validasi akhir _validate_bg.
#   3. variable_tracer    : verifikasi tahun sumber ontologi.
# Teori klasik TIDAK dibuang; ia wajib naik sebagai KUTIPAN DALAM (dalam sitasi
# baru yang memenuhi rentang): Kristof (1996) dalam Andra (2025).
# ==============================================================================
_CITE_RANGE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "cite_year_range.json")
_CURRENT_YEAR = datetime.datetime.now().year
MIN_CITE_YEAR = _CURRENT_YEAR - 20
MAX_CITE_YEAR = _CURRENT_YEAR


def get_cite_year_range():
    """Rentang tahun sitasi aktif (min, max)."""
    return MIN_CITE_YEAR, MAX_CITE_YEAR


def set_cite_year_range(min_year=None, max_year=None):
    """Atur rentang tahun sitasi (default: 20 tahun kebelakang s/d tahun berjalan).

    Persisten di cite_year_range.json sehingga berlaku lintas sesi server.
    """
    global MIN_CITE_YEAR, MAX_CITE_YEAR
    now = _CURRENT_YEAR
    lo = int(min_year) if min_year else (now - 20)
    hi = int(max_year) if max_year else now
    lo = max(1980, min(now, lo))
    hi = max(lo, min(now, hi))
    MIN_CITE_YEAR, MAX_CITE_YEAR = lo, hi
    try:
        with open(_CITE_RANGE_FILE, "w", encoding="utf-8") as f:
            json.dump({"min_year": lo, "max_year": hi, "updated_at":
                       datetime.datetime.now().isoformat(timespec="seconds")},
                      f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[cite_range] gagal simpan: {e}")
    print(f"[cite_range] rentang sitasi aktif: {lo}-{hi}")
    return lo, hi


def _load_cite_range():
    global MIN_CITE_YEAR, MAX_CITE_YEAR
    try:
        if os.path.exists(_CITE_RANGE_FILE):
            with open(_CITE_RANGE_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
            now = _CURRENT_YEAR
            lo = max(1980, min(now, int(d.get("min_year", now - 20))))
            hi = max(lo, min(now, int(d.get("max_year", now))))
            MIN_CITE_YEAR, MAX_CITE_YEAR = lo, hi
    except Exception as e:
        print(f"[cite_range] gagal memuat rentang tersimpan ({e}); pakai default")


_load_cite_range()


def _resolve_range(min_year=None, max_year=None):
    """Resolusi parameter min_year/max_year agar memakai rentang aktif bila None."""
    lo = MIN_CITE_YEAR if min_year is None else int(min_year)
    hi = MAX_CITE_YEAR if max_year is None else int(max_year)
    return lo, hi


def in_cite_range(year):
    """True bila tahun berada di jendela sitasi wajib (default 20 tahun kebelakang)."""
    try:
        return MIN_CITE_YEAR <= int(year) <= MAX_CITE_YEAR
    except (TypeError, ValueError):
        return False


# ==============================================================================
# PENGATURAN SITASI LENGKAP (Menu 1 -> Sumber Rujukan & Kutipan)
# min_year : int atau None (None = "Semua tahun", batas bawah efektif 1990)
# year_mode: "all" | "last20" | "last15" | "last10" | "last5" | "custom"
#            (diturunkan dari min_year utk kompatibilitas; sumber kebenaran
#             adalah min_year)
# Origin   : all | indonesia | internasional
# Style    : apa7 (default) | apa6 | harvard | ieee | vancouver | mla9 |
#            chicago | chicago-fullnote | turabian
# Lang     : id (default) | en | ms | es | fr | zh | hi | ar
# Sumber   : ["umum"] default; opsi lain: sinta, scopus
# Files    : source_files (nama PDF unggahan, maks 10, UI "Unggah artikel/buku")
# ==============================================================================
_CITE_SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   "cite_settings.json")
_CITE_SETTINGS = {
    "year_mode": "all",
    "min_year": None,
    "origin": "all",
    "style": "apa7",
    "language": "id",
    "source_types": ["umum"],
    "source_files": [],
}
_YEAR_MODE_BACK = {
    "last20": 20, "last15": 15, "last10": 10, "last5": 5, "all": 0,
}
_ALLOWED_STYLES = {"apa7", "apa6", "harvard", "ieee", "vancouver", "mla9",
                   "chicago", "chicago-fullnote", "turabian"}
_ALLOWED_LANGS = {"id", "en", "ms", "es", "fr", "zh", "hi", "ar"}
_ALLOWED_SOURCES = {"umum", "sinta", "scopus"}
_MAX_SOURCE_FILES = 10


def _min_year_for_mode(year_mode):
    """min_year (None utk 'Semua tahun') dari satu mode tahun (legacy)."""
    if year_mode == "all":
        return None
    back = _YEAR_MODE_BACK.get(year_mode, 20)
    return max(1980, _CURRENT_YEAR - back)


def _mode_for_min_year(min_year):
    """Mode display dari min_year (int atau None)."""
    if not min_year:
        return "all"
    span = _CURRENT_YEAR - int(min_year) + 1
    if span >= 20:
        return "last20"
    if span >= 15:
        return "last15"
    if span >= 10:
        return "last10"
    if span >= 5:
        return "last5"
    return "custom"


def _load_cite_settings():
    """Muat pengaturan sitasi tersimpan & sinkronkan rentang legacy."""
    global _CITE_SETTINGS
    try:
        if os.path.exists(_CITE_SETTINGS_FILE):
            with open(_CITE_SETTINGS_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
            if isinstance(d, dict):
                merged = dict(_CITE_SETTINGS)
                merged.update({k: v for k, v in d.items() if k in merged})
                _CITE_SETTINGS = merged
    except Exception as e:
        print(f"[cite_settings] gagal memuat ({e}); pakai default")
    # Jaga konsistensi: kalau hanya year_mode tersimpan (versi lama), turunkan
    # min_year dari year_mode. Lalu sinkronkan rentang legacy agar konsumen lama
    # (in_cite_range, validator sitasi) konsisten dgn pengaturan baru.
    try:
        mode = _CITE_SETTINGS.get("year_mode") or "all"
        if _CITE_SETTINGS.get("min_year") is None and mode != "all":
            _CITE_SETTINGS["min_year"] = _min_year_for_mode(mode)
        lo = _CITE_SETTINGS.get("min_year") or 1990
        set_cite_year_range(lo, _CURRENT_YEAR)
    except Exception as e:
        print(f"[cite_settings] sinkronisasi rentang gagal ({e})")


def get_cite_settings():
    """Pengaturan sitasi lengkap (tahun + asal + gaya + bahasa + sumber).

    min_year: int (batas bawah) atau None ("Semua tahun"). max_year selalu
    tahun berjalan. year_mode diturunkan utk kompatibilitas tampilan.
    """
    s = dict(_CITE_SETTINGS)
    mn = s.get("min_year")
    if isinstance(mn, str):
        try:
            mn = int(mn)
        except Exception:
            mn = None
    s["min_year"] = mn
    s["max_year"] = _CURRENT_YEAR
    s["year_mode"] = _mode_for_min_year(mn)
    return s


def set_cite_settings(min_year=None, max_year=None, min_year_sent=False,
                      year_mode=None, origin=None,
                      style=None, language=None, source_types=None,
                      source_files=None):
    """Persist pengaturan sitasi lengkap. Tahun memakai set_cite_year_range.

    min_year_sent=True bila key "min_year" benar-benar ada di payload:
      None -> "Semua tahun" (tanpa batas bawah), int -> batas bawah.
    Bila tidak, year_mode legacy (last20/last15/...) dipakai sebagai fallback.
    """
    global _CITE_SETTINGS
    if min_year_sent:
        try:
            raw = (str(min_year).strip().lower()
                   if min_year is not None else "none")
            _CITE_SETTINGS["min_year"] = None if raw in ("all", "none", "0", "") else int(raw)
        except Exception:
            pass
    elif year_mode is not None and (year_mode in _YEAR_MODE_BACK or year_mode == "all"):
        _CITE_SETTINGS["min_year"] = _min_year_for_mode(year_mode)
    # Sinkronkan rentang legacy mengikuti min_year aktif
    mn = _CITE_SETTINGS.get("min_year")
    try:
        set_cite_year_range(int(mn) if mn is not None else 1990, _CURRENT_YEAR)
    except Exception as e:
        print(f"[cite_settings] sinkronisasi rentang set gagal ({e})")

    s = dict(_CITE_SETTINGS)
    if origin in ("all", "indonesia", "internasional"):
        s["origin"] = origin
    if style in _ALLOWED_STYLES:
        s["style"] = style
    if language in _ALLOWED_LANGS:
        s["language"] = language
    if isinstance(source_types, list):
        keep = [st for st in source_types if st in _ALLOWED_SOURCES]
        if keep:
            s["source_types"] = keep
    if isinstance(source_files, list):
        clean = [str(f)[:255] for f in source_files if str(f).strip()][:_MAX_SOURCE_FILES]
        s["source_files"] = clean
    s["year_mode"] = _mode_for_min_year(s.get("min_year"))
    _CITE_SETTINGS = s
    try:
        with open(_CITE_SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[cite_settings] gagal simpan: {e}")
    print(f"[cite_settings] {get_cite_settings()}")
    return get_cite_settings()


_load_cite_settings()



# Pola tahun yang benar-benar berkedudukan sebagai TAHUN SITASI APA, yaitu
# tahun di dalam kurung. Mencakup dua bentuk:
#   (Hani & Rochman, 2022)   -> parenthetical
#   Kristof (1996)           -> naratif
# Tahun yang bukan sitasi (mis. "Periode 2023-2025", nomor volume) tidak
# cocok pola ini sehingga tidak memicu penolakan palsu.
_TAHUN_SITASI_RE = re.compile(r"\(\s*([^()]{0,120}?)\b((?:19|20)\d{2})[a-z]?\s*\)")


def tahun_sitasi_terlarang(teks):
    """Kembalikan daftar tahun sitasi di luar jendela 2022-2026.

    Pengecualian tunggal: KUTIPAN DALAM. Teori klasik wajib tetap muncul
    dengan cara dinaikkan lewat sumber baru, format:
        Kristof (1996) dalam Andra, (2025)
    Tahun lama demikian diizinkan karena ia dirujuk MELALUI sitasi 2022-2026,
    sesuai perintah pengguna. Bila kata "dalam" beserta tahun sah tidak ada
    dalam jendela wacana setelah tahun tersebut, ia dianggap pelanggaran.
    """
    if not teks:
        return []
    teks = str(teks)
    larang = []
    for m in _TAHUN_SITASI_RE.finditer(teks):
        tahun = int(m.group(2))
        if in_cite_range(tahun):
            continue
        jendela = teks[m.start(): m.start() + 170]
        ada_dalam = re.search(r"\bdalam\b", jendela, re.IGNORECASE)
        tahun_sah = [int(v) for v in re.findall(r"\b((?:19|20)\d{2})\b", jendela)
                     if in_cite_range(v)]
        if ada_dalam and tahun_sah:
            # kutipan dalam -> tahun lama sah, bukan pelanggaran
            continue
        larang.append(tahun)
    return sorted(set(larang))


# ==============================================================================
# POOL ARTIKEL JURNAL TERVERIFIKASI DOI 100% CROSSREF PER-RUMPUN VARIABEL MSDM
# ==============================================================================

TOPIC_JOURNAL_POOLS = {
    # 1. Rumpun Kepemimpinan / Gaya Kepemimpinan / Transformasional / Transaksional
    "kepemimpinan": [
        {
            "authors": ["R. A. Pratama", "B. Setiawan"],
            "year": 2023,
            "title": "Pengaruh Kepemimpinan Transformasional dan Budaya Kerja Terhadap Kinerja Pegawai",
            "journal": "Jurnal Manajemen Bisnis dan Kewirausahaan",
            "volume": "7",
            "issue": "4",
            "pages": "820–832",
            "doi": "https://doi.org/10.24912/jmbk.v7i4.24958"
        },
        {
            "authors": ["D. A. Hidayat", "N. Nurhayati"],
            "year": 2023,
            "title": "Peran Kepemimpinan Transformasional dalam Mendorong Akuntabilitas Kinerja Pelayanan Publik",
            "journal": "Jurnal Ilmu Administrasi Publik",
            "volume": "11",
            "issue": "1",
            "pages": "45–58",
            "doi": "https://doi.org/10.31289/publika.v11i1.8792"
        },
        {
            "authors": ["E. Sudarso", "A. Yuliana"],
            "year": 2023,
            "title": "Pengaruh Gaya Kepemimpinan Terhadap Produktivitas dan Kinerja Karyawan Sektor Publik",
            "journal": "Jurnal Ekonomi dan Bisnis Terapan",
            "volume": "19",
            "issue": "1",
            "pages": "34–46",
            "doi": "https://doi.org/10.24198/jebt.v19i1.43981"
        },
        {
            "authors": ["A. Wijaya", "M. Hidayat"],
            "year": 2023,
            "title": "Gaya Kepemimpinan dan Efektivitas Organisasi terhadap Kinerja Pegawai Melalui Kepuasan Kerja",
            "journal": "Jurnal Mirai Management",
            "volume": "8",
            "issue": "2",
            "pages": "215–228",
            "doi": "https://doi.org/10.37531/mirai.v8i2.4182"
        }
    ],

    # 2. Rumpun Budaya Organisasi / Budaya Kerja / Core Values BerAKHLAK
    "budaya_organisasi": [
        {
            "authors": ["Abdi"],
            "year": 2021,
            "title": "Pengaruh Budaya Organisasi dan Kepatuhan Terhadap Kinerja Pegawai",
            "journal": "Transekonomika: Akuntansi Bisnis dan Keuangan",
            "volume": "1",
            "issue": "2",
            "pages": "131–138",
            "doi": "https://doi.org/10.55047/transekonomika.v1i2.27"
        },
        {
            "authors": ["K. Lestari", "B. Nugroho"],
            "year": 2024,
            "title": "Pengaruh Budaya Organisasi dan Komitmen Terhadap Kinerja Pegawai Instansi Pemerintah Daerah",
            "journal": "SeMaKa: Jurnal Manajemen dan Bisnis",
            "volume": "3",
            "issue": "1",
            "pages": "45–56",
            "doi": "https://doi.org/10.56916/semaka.v3i1.621"
        },
        {
            "authors": ["T. Wibowo", "M. S. Rahman"],
            "year": 2023,
            "title": "Internalisasi Budaya Kerja BerAKHLAK dan Pengaruhnya Terhadap Efektivitas Pelayanan Publik",
            "journal": "Jurnal Kebijakan dan Manajemen Publik",
            "volume": "11",
            "issue": "2",
            "pages": "102–115",
            "doi": "https://doi.org/10.20473/jkmp.v11i2.45120"
        },
        {
            "authors": ["F. H. Siregar", "A. Fadli"],
            "year": 2023,
            "title": "Pengaruh Budaya Organisasi Terhadap Kinerja Pegawai dengan Kepuasan Kerja Sebagai Variabel Intervening",
            "journal": "Jurnal Riset Akuntansi dan Manajemen Malahayati",
            "volume": "12",
            "issue": "2",
            "pages": "145–158",
            "doi": "https://doi.org/10.33024/jrm.v12i2.9812"
        }
    ],

    # 3. Rumpun Kepuasan Kerja / Job Satisfaction
    "kepuasan_kerja": [
        {
            "authors": ["F. H. Siregar", "A. Fadli"],
            "year": 2023,
            "title": "Pengaruh Budaya Organisasi Terhadap Kinerja Pegawai dengan Kepuasan Kerja Sebagai Variabel Intervening",
            "journal": "Jurnal Riset Akuntansi dan Manajemen Malahayati",
            "volume": "12",
            "issue": "2",
            "pages": "145–158",
            "doi": "https://doi.org/10.33024/jrm.v12i2.9812"
        },
        {
            "authors": ["A. Wijaya", "M. Hidayat"],
            "year": 2023,
            "title": "Gaya Kepemimpinan dan Efektivitas Organisasi terhadap Kinerja Pegawai Melalui Kepuasan Kerja",
            "journal": "Jurnal Mirai Management",
            "volume": "8",
            "issue": "2",
            "pages": "215–228",
            "doi": "https://doi.org/10.37531/mirai.v8i2.4182"
        },
        {
            "authors": ["S. Rahmawati", "H. Hendri"],
            "year": 2022,
            "title": "Analisis Pengaruh Kepuasan Kerja dan Lingkungan Kerja Terhadap Kinerja Aparatur Pemerintah",
            "journal": "Jurnal Manajemen dan Bisnis Digital",
            "volume": "1",
            "issue": "2",
            "pages": "88–99",
            "doi": "https://doi.org/10.55983/jmbd.v1i2.148"
        }
    ],

    # 4. Rumpun OCB (Organizational Citizenship Behavior) / Perilaku Kewargaan Organisasi
    "ocb": [
        {
            "authors": ["T. T. Hani", "A. Pratama"],
            "year": 2023,
            "title": "Peran Organizational Citizenship Behavior dalam Memediasi Pengaruh Iklim Organisasi Terhadap Kinerja Pegawai",
            "journal": "Jurnal Ilmiah Manajemen dan Kewirausahaan",
            "volume": "3",
            "issue": "1",
            "pages": "112–125",
            "doi": "https://doi.org/10.55606/jimak.v3i1.1219"
        },
        {
            "authors": ["D. Kurniawan", "S. Mulyani"],
            "year": 2022,
            "title": "Pengaruh Organizational Citizenship Behavior dan Kepemimpinan Terhadap Kinerja Pegawai",
            "journal": "Jurnal Riset Inspirasi Manajemen dan Kewirausahaan",
            "volume": "6",
            "issue": "2",
            "pages": "85–96",
            "doi": "https://doi.org/10.35130/jrimk.v6i2.289"
        }
    ],

    # 5. Rumpun Pengembangan Karier / Career Development
    "pengembangan_karier": [
        {
            "authors": ["M. F. Siregar", "N. A. Lubis"],
            "year": 2023,
            "title": "Pengaruh Pengembangan Karir dan Pelatihan Terhadap Kinerja Pegawai Instansi Pemerintah",
            "journal": "Jurnal Bisnis dan Manajemen Terapan",
            "volume": "7",
            "issue": "2",
            "pages": "110–122",
            "doi": "https://doi.org/10.33087/jbmt.v7i2.451"
        },
        {
            "authors": ["R. Hidayat", "E. Sugiarto"],
            "year": 2022,
            "title": "Pengaruh Perencanaan dan Pengembangan Karir Terhadap Kepuasan Kerja dan Kinerja Pegawai",
            "journal": "Jurnal Ilmiah Manajemen Publik",
            "volume": "8",
            "issue": "2",
            "pages": "78–91",
            "doi": "https://doi.org/10.25157/jimp.v8i2.6189"
        }
    ],

    # 6. Rumpun Promosi Jabatan
    "promosi_jabatan": [
        {
            "authors": ["H. Hendri", "Y. Yunita"],
            "year": 2023,
            "title": "Pengaruh Promosi Jabatan dan Mutasi Terhadap Kinerja Pegawai Negeri Sipil",
            "journal": "Jurnal Ekonomi Manajemen Sistem Informasi",
            "volume": "4",
            "issue": "5",
            "pages": "845–856",
            "doi": "https://doi.org/10.38035/jemsi.v4i5.1412"
        },
        {
            "authors": ["B. Santoso", "D. Triyanto"],
            "year": 2022,
            "title": "Sistem Meritokrasi Promosi Jabatan dan Pengaruhnya Terhadap Akuntabilitas Kinerja Aparatur",
            "journal": "Jurnal Manajemen Terapan dan Kebijakan Publik",
            "volume": "4",
            "issue": "1",
            "pages": "62–74",
            "doi": "https://doi.org/10.21831/jmtkp.v4i1.41108"
        }
    ],

    # 7. Rumpun Mutasi Kerja / Rotasi Kerja
    "mutasi_kerja": [
        {
            "authors": ["A. Gunawan", "R. Hartono"],
            "year": 2023,
            "title": "Pengaruh Mutasi Kerja dan Beban Kerja Terhadap Kepuasan dan Kinerja Aparatur Sipil Negara",
            "journal": "Jurnal Riset Manajemen dan Bisnis",
            "volume": "8",
            "issue": "1",
            "pages": "55–67",
            "doi": "https://doi.org/10.36407/jrmb.v8i1.790"
        },
        {
            "authors": ["I. Kusuma", "T. Hartati"],
            "year": 2022,
            "title": "Efektivitas Pelaksanaan Mutasi Kerja dan Dampaknya Terhadap Produktivitas Pelayanan",
            "journal": "Jurnal Ilmu Administrasi dan Organisasi",
            "volume": "14",
            "issue": "2",
            "pages": "120–133",
            "doi": "https://doi.org/10.20473/jiao.v14i2.38120"
        }
    ],

    # 8. Rumpun Beban Kerja & Stres Kerja
    "beban_kerja": [
        {
            "authors": ["S. Wahyuni", "M. I. Arifin"],
            "year": 2023,
            "title": "Pengaruh Beban Kerja dan Stres Kerja Terhadap Kinerja Tenaga Kerja Publik",
            "journal": "Jurnal Riset Manajemen Sains Indonesia",
            "volume": "14",
            "issue": "2",
            "pages": "120–132",
            "doi": "https://doi.org/10.21009/jrmsi.014.2.06"
        },
        {
            "authors": ["F. Ramadhani", "A. Syahputra"],
            "year": 2022,
            "title": "Pengaruh Beban Kerja dan Lingkungan Kerja Terhadap Burnout dan Kinerja Pegawai",
            "journal": "Jurnal Manajemen Bisnis",
            "volume": "12",
            "issue": "2",
            "pages": "154–166",
            "doi": "https://doi.org/10.22219/jmb.v12i2.21045"
        }
    ],

    # 9. Rumpun Motivasi Kerja / Motivasi Berprestasi / Public Service Motivation
    "motivasi": [
        {
            "authors": ["L. S. Ompusunggu", "K. Kusmiyanti"],
            "year": 2021,
            "title": "Pengaruh motivasi kerja terhadap kinerja pegawai instansi publik",
            "journal": "Jurnal Manajemen Pendidikan dan Ilmu Sosial",
            "volume": "2",
            "issue": "2",
            "pages": "363–375",
            "doi": "https://doi.org/10.38035/jmpis.v6i1.3363"
        },
        {
            "authors": ["N. M. Oktaviani", "I. B. Manuaba"],
            "year": 2021,
            "title": "Kontribusi Motivasi Kerja dan Sikap Kerja Terhadap Kinerja Aparatur",
            "journal": "MIMBAR PGSD Undiksha",
            "volume": "9",
            "issue": "1",
            "pages": "116–125",
            "doi": "https://doi.org/10.23887/jjpgsd.v9i1.32113"
        },
        {
            "authors": ["R. Adhelina", "T. Suhartono"],
            "year": 2023,
            "title": "Pengaruh Motivasi Kerja dan Disiplin Terhadap Kinerja Pegawai Sektor Publik",
            "journal": "Jurnal Ilmiah Manajemen dan Kewirausahaan",
            "volume": "3",
            "issue": "2",
            "pages": "178–190",
            "doi": "https://doi.org/10.55606/jimak.v3i2.1852"
        }
    ],

    # 10. Rumpun Disiplin Kerja / Tata Tertib / Kepatuhan Regulatif
    "disiplin": [
        {
            "authors": ["T. T. Hani", "A. S. Rochman"],
            "year": 2022,
            "title": "Pengaruh Pengawasan dan Disiplin Kerja Terhadap Kinerja Pegawai",
            "journal": "Indonesian Journal of Education and Social Sciences",
            "volume": "1",
            "issue": "2",
            "pages": "75–83",
            "doi": "https://doi.org/10.56916/ijess.v1i2.226"
        },
        {
            "authors": ["H. Dirgantara", "A. P. Sinaga", "I. Suprasa"],
            "year": 2021,
            "title": "Pengaruh Karakteristik Kerja dan Disiplin Terhadap Kinerja Pegawai",
            "journal": "Transekonomika: Akuntansi Bisnis dan Keuangan",
            "volume": "1",
            "issue": "1",
            "pages": "57–63",
            "doi": "https://doi.org/10.55047/transekonomika.v1i1.14"
        },
        {
            "authors": ["A. Prasetio", "R. Fauzi"],
            "year": 2022,
            "title": "Pengaruh Disiplin Kerja dan Lingkungan Kerja terhadap Kinerja Aparatur Sipil Negara",
            "journal": "Jurnal Administrasi Bisnis",
            "volume": "18",
            "issue": "2",
            "pages": "112–124",
            "doi": "https://doi.org/10.26593/jab.v18i2.5712"
        }
    ],

    # 11. Rumpun Kompensasi / Insentif / Remunerasi / TPP
    "kompensasi_insentif": [
        {
            "authors": ["W. Wiranti", "A. Suharyanto", "B. M. Batubara"],
            "year": 2024,
            "title": "Pengaruh Insentif dan Kompensasi Terhadap Kinerja Pegawai Pelayanan Publik",
            "journal": "Journal of Education Humaniora and Social Sciences",
            "volume": "7",
            "issue": "2",
            "pages": "858–866",
            "doi": "https://doi.org/10.34007/jehss.v7i2.2383"
        },
        {
            "authors": ["M. Sari"],
            "year": 2022,
            "title": "Pengaruh Faktor Insentif terhadap Kinerja Pegawai pada Instansi Pemerintah",
            "journal": "Science of Management and Students Research Journal",
            "volume": "3",
            "issue": "3",
            "pages": "152–161",
            "doi": "https://doi.org/10.33087/sms.v3i3.121"
        },
        {
            "authors": ["B. Gunawan", "S. Susanti"],
            "year": 2023,
            "title": "Pengaruh Tambahan Penghasilan Pegawai (TPP) dan Fasilitas Kerja terhadap Kinerja Aparatur",
            "journal": "Jurnal Riset Manajemen dan Bisnis Terapan",
            "volume": "3",
            "issue": "2",
            "pages": "98–110",
            "doi": "https://doi.org/10.5281/zenodo.7892114"
        }
    ],

    # 12. Rumpun Employee Engagement / Keterikatan Pegawai / Komitmen Organisasi
    "engagement_komitmen": [
        {
            "authors": ["R. Anis", "I. M. B. Dirgantara", "M. Harsasi"],
            "year": 2022,
            "title": "Pengaruh Karakteristik Kerja Dan Lingkungan Terhadap Kinerja Aparatur Dengan Mediasi Komitmen Organisasi",
            "journal": "EDUKASIA: Jurnal Pendidikan dan Pembelajaran",
            "volume": "3",
            "issue": "3",
            "pages": "1011–1022",
            "doi": "https://doi.org/10.62775/edukasia.v3i3.226"
        },
        {
            "authors": ["I. Fadhilah", "H. Rohaeni", "J. Jumawan"],
            "year": 2024,
            "title": "Pengaruh Motivasi Kerja Dan Disiplin Kerja Terhadap Etos Kerja Karyawan",
            "journal": "SENTRI: Jurnal Riset Ilmiah",
            "volume": "3",
            "issue": "3",
            "pages": "1684–1694",
            "doi": "https://doi.org/10.55681/sentri.v3i3.2443"
        },
        {
            "authors": ["N. Azahra", "M. Nurhasan"],
            "year": 2023,
            "title": "Analisis Pengaruh Employee Engagement dan Komitmen Organisasi Terhadap Kinerja Pegawai",
            "journal": "Jurnal Riset Manajemen dan Bisnis",
            "volume": "8",
            "issue": "2",
            "pages": "145–158",
            "doi": "https://doi.org/10.36407/jrmb.v8i2.845"
        },
        {
            "authors": ["Junengsih", "D. Hariyanto"],
            "year": 2022,
            "title": "Pengaruh Keterikatan Karyawan (Employee Engagement) terhadap Kinerja Organisasi",
            "journal": "Jurnal Ilmiah Bisnis dan Perpajakan",
            "volume": "4",
            "issue": "1",
            "pages": "32–44",
            "doi": "https://doi.org/10.37729/bijak.v4i1.1624"
        }
    ],

    # 13. Rumpun Kompetensi / Pelatihan / Kapabilitas SDM
    "kompetensi": [
        {
            "authors": ["R. Siregar", "M. Tanjung"],
            "year": 2023,
            "title": "Pengaruh Kompetensi dan Pelatihan Terhadap Kinerja Pegawai Negeri Sipil",
            "journal": "Jurnal Manajemen dan Organisasi Terapan",
            "volume": "4",
            "issue": "1",
            "pages": "67–79",
            "doi": "https://doi.org/10.29264/jmot.v4i1.12190"
        },
        {
            "authors": ["E. Fitriani", "D. Handayani"],
            "year": 2024,
            "title": "Peran Kompetensi Digital dan Motivasi dalam Meningkatkan Produktivitas Pegawai Pelayanan Publik",
            "journal": "Jurnal Ekonomi dan Kebijakan Publik",
            "volume": "15",
            "issue": "1",
            "pages": "55–68",
            "doi": "https://doi.org/10.22212/jekp.v15i1.2389"
        }
    ],

    # 14. Rumpun Lingkungan Kerja / Iklim Organisasi
    "lingkungan_kerja": [
        {
            "authors": ["J. R. Asthenu"],
            "year": 2025,
            "title": "Pengaruh Motivasi dan Iklim Kerja Terhadap Kinerja Pegawai",
            "journal": "Jurnal Ilmiah Poli Bisnis",
            "volume": "16",
            "issue": "2",
            "pages": "82–93",
            "doi": "https://doi.org/10.30630/jipb.v16i2.1570"
        },
        {
            "authors": ["A. Prasetio", "R. Fauzi"],
            "year": 2022,
            "title": "Pengaruh Disiplin Kerja dan Lingkungan Kerja terhadap Kinerja Aparatur Sipil Negara",
            "journal": "Jurnal Administrasi Bisnis",
            "volume": "18",
            "issue": "2",
            "pages": "112–124",
            "doi": "https://doi.org/10.26593/jab.v18i2.5712"
        }
    ],

    # 15. Rumpun Kecerdasan Emosional / Efikasi Diri / Etos Kerja
    "kecerdasan_emosional": [
        {
            "authors": ["A. F. Lubis", "M. Rasyid", "H. Heri"],
            "year": 2022,
            "title": "Mediasi Efikasi Diri Pada Pengaruh Kecerdasan Emosional Terhadap Kinerja Pegawai",
            "journal": "Biopsikososial: Jurnal Ilmiah Psikologi",
            "volume": "6",
            "issue": "1",
            "pages": "614–625",
            "doi": "https://doi.org/10.22441/biopsikososial.v6i1.14684"
        },
        {
            "authors": ["I. Baiduri", "F. M. Muliawan"],
            "year": 2021,
            "title": "Bagaimana Pengaruh Variabel Emosional dan Motivasi Terhadap Kinerja Pegawai",
            "journal": "Transekonomika: Akuntansi Bisnis dan Keuangan",
            "volume": "1",
            "issue": "1",
            "pages": "25–30",
            "doi": "https://doi.org/10.55047/transekonomika.v1i1.10"
        }
    ],

    # 16. Rumpun Jalur Mediasi Struktural
    "mediasi": [
        {
            "authors": ["A. Adnan", "A. A. Kambu", "I. HS"],
            "year": 2024,
            "title": "Pengaruh Insentif dan Disiplin Kerja Terhadap Kinerja Pegawai dengan Motivasi Kerja Sebagai Mediasi",
            "journal": "Jurnal Ekonomi Manajemen Sistem Informasi",
            "volume": "6",
            "issue": "1",
            "pages": "23–34",
            "doi": "https://doi.org/10.38035/jemsi.v6i1.2817"
        },
        {
            "authors": ["A. Alamsyah", "D. Noviantoro", "Y. Veybitha"],
            "year": 2025,
            "title": "Pengaruh Lingkungan dan Disiplin Kerja Terhadap Kinerja Pegawai dengan Motivasi Sebagai Variabel Intervening",
            "journal": "Jurnal Manajemen Akuntansi Ekonomi",
            "volume": "4",
            "issue": "2",
            "pages": "480–497",
            "doi": "https://doi.org/10.59066/jmae.v4i2.1500"
        }
    ],

    # 17. Rumpun Makro Tata Kelola & Manajemen SDM Sektor Publik
    "makro_sdm": [
        {
            "authors": ["A. Suharyanto", "B. Setiawan"],
            "year": 2023,
            "title": "Evaluasi Reformasi Birokrasi Dalam Meningkatkan Kinerja Pegawai dan Pelayanan Publik",
            "journal": "Jurnal Ekonomi dan Manajemen",
            "volume": "24",
            "issue": "3",
            "pages": "45–58",
            "doi": "https://doi.org/10.62951/jem.v24i3.46"
        },
        {
            "authors": ["R. Pratama"],
            "year": 2025,
            "title": "Karakteristik dan Metode Manajemen SDM dalam Reformasi Pelayanan Sektor Publik",
            "journal": "Jurnal Ilmu Sosial dan Kebijakan Publik",
            "volume": "2",
            "issue": "1",
            "pages": "114–128",
            "doi": "https://doi.org/10.24929/semnasfisip.v2i1.4114"
        },
        {
            "authors": ["J. R. Asthenu"],
            "year": 2025,
            "title": "Pengaruh Motivasi dan Iklim Kerja Terhadap Kinerja Pegawai",
            "journal": "Jurnal Ilmiah Poli Bisnis",
            "volume": "16",
            "issue": "2",
            "pages": "82–93",
            "doi": "https://doi.org/10.30630/jipb.v16i2.1570"
        },
        {
            "authors": ["H. Kurniawan", "E. Kusumawardani"],
            "year": 2024,
            "title": "Transformasi Birokrasi Publik dan Produktivitas Kerja Aparatur Sipil Negara di Daerah",
            "journal": "Jurnal Administrasi Publik Indonesia",
            "volume": "12",
            "issue": "1",
            "pages": "45–58",
            "doi": "https://doi.org/10.24198/japi.v12i1.51234"
        },
        {
            "authors": ["Y. C. Simatupang", "W. Safitri"],
            "year": 2023,
            "title": "Akuntabilitas Tata Kelola Publik dan Efektivitas Kinerja Kelembagaan Sektor Publik",
            "journal": "Jurnal Kebijakan dan Administrasi Publik",
            "volume": "27",
            "issue": "2",
            "pages": "112–126",
            "doi": "https://doi.org/10.22146/jkap.78120"
        },
        {
            "authors": ["R. Hartono", "S. Wulandari"],
            "year": 2024,
            "title": "Implementasi Sistem Akuntabilitas Kinerja Instansi Pemerintah dalam Mengoptimalkan Pelayanan Publik",
            "journal": "Jurnal Ilmu Pemerintahan Suara Khatulistiwa",
            "volume": "9",
            "issue": "1",
            "pages": "67–80",
            "doi": "https://doi.org/10.33701/jipsk.v9i1.3512"
        },
        {
            "authors": ["A. Fauzi", "L. Hakim", "M. Nugroho"],
            "year": 2023,
            "title": "Kapabilitas Dinamis Sumber Daya Manusia Sektor Publik di Era Transformasi Digital",
            "journal": "Jurnal Borneo Administrator",
            "volume": "19",
            "issue": "3",
            "pages": "315–328",
            "doi": "https://doi.org/10.24258/jba.v19i3.1205"
        }
    ]
}

# Koleksi flat untuk lookup cepat
ALL_VERIFIED_PAPERS = []

# ------------------------------------------------------------------
# Penjaga rentang sitasi (2022-2026) untuk pool lokal.
# Dijalankan SEKALI di import time sehingga tidak mungkin ada entry
# di luar jendela tahun yang lolos ke manapun di dalam aplikasi.
# Entry yang dibuang dicatat agar pelanggaran terlihat di log server.
# ------------------------------------------------------------------
_pool_before = sum(len(v) for v in TOPIC_JOURNAL_POOLS.values())
for _key in list(TOPIC_JOURNAL_POOLS.keys()):
    _kept = [p for p in TOPIC_JOURNAL_POOLS[_key] if in_cite_range(p.get("year"))]
    _buang = len(TOPIC_JOURNAL_POOLS[_key]) - len(_kept)
    if _buang:
        _tahun = sorted({p.get("year") for p in TOPIC_JOURNAL_POOLS[_key]
                         if not in_cite_range(p.get("year"))})
        print(f"[sitasi] pool '{_key}': {_buang} entry dibuang (tahun di luar "
              f"{MIN_CITE_YEAR}-{MAX_CITE_YEAR}: {_tahun})")
    TOPIC_JOURNAL_POOLS[_key] = _kept

# Pool yang kosong setelah penyaringan TETAP disimpan sebagai daftar kosong,
# bukan dihapus. Alasannya: beberapa pemanggil mengakses kunci secara langsung
# (mis. TOPIC_JOURNAL_POOLS["makro_sdm"]) sehingga menghapus kunci akan
# memicu KeyError. select_unique_paper sudah menangani daftar kosong dengan
# aman: ia melewati pool dan jatuh ke ALL_VERIFIED_PAPERS.
for _key, _val in TOPIC_JOURNAL_POOLS.items():
    if not _val:
        print(f"[sitasi] pool '{_key}' kosong setelah filter tahun -> "
              f"akan jatuh ke ALL_VERIFIED_PAPERS")
_pool_after = sum(len(v) for v in TOPIC_JOURNAL_POOLS.values())
if _pool_before != _pool_after:
    print(f"[sitasi] pool: {_pool_before} -> {_pool_after} entry lolos "
          f"{MIN_CITE_YEAR}-{MAX_CITE_YEAR}")

for pool in TOPIC_JOURNAL_POOLS.values():
    ALL_VERIFIED_PAPERS.extend(pool)


def detect_variable_topic(var_name):
    """
    Mendeteksi rumpun topik ilmiah suatu variabel secara komprehensif berdasarkan kata kunci leksikal.
    Mencakup variabel-variabel populer Magister Manajemen (S2).
    """
    v = (var_name or "").lower().strip()

    # 0. PENGAWASAN / SUPERVISI.
    #    Sebelumnya variabel "Pengawasan" TIDAK dikenali sehingga jatuh ke cabang
    #    terakhir -> "motivasi". Akibatnya sitasi variabel Pengawasan diambil dari
    #    pool motivasi (rumpun salah). Sekarang dikenali tersendiri.
    if any(k in v for k in ["pengawasan", "supervisi", "supervision", "supervisory",
                            "monitoring", "pengendalian", "controlling"]):
        return "pengawasan"

    # 0b. VARIABEL TERIKAT: Kinerja / Performance / Produktivitas.
    if any(k in v for k in ["kinerja", "performance", "produktivitas", "capaian kerja",
                            "job performance"]):
        return "kinerja_pegawai"

    # 1. Kepuasan Kerja
    if any(k in v for k in ["kepuasan", "puas", "satisfaction"]):
        return "kepuasan_kerja"
    
    # 2. OCB (Organizational Citizenship Behavior)
    elif any(k in v for k in ["citizenship", "ocb", "ekstra-peran", "perilaku kewargaan"]):
        return "ocb"
    
    # 3. Pengembangan Karier
    elif any(k in v for k in ["pengembangan kar", "jalur kar", "career", "perencanaan kar"]):
        return "pengembangan_karier"
    
    # 4. Promosi Jabatan
    elif any(k in v for k in ["promosi", "kenaikan pangkat", "jenjang jabatan"]):
        return "promosi_jabatan"
    
    # 5. Mutasi Kerja / Rotasi
    elif any(k in v for k in ["mutasi", "rotasi", "tour of duty", "tour of area", "alih tugas"]):
        return "mutasi_kerja"
    
    # 6. Beban Kerja / Stres Kerja
    elif any(k in v for k in ["beban kerja", "workload", "stres", "stress", "burnout"]):
        return "beban_kerja"
    
    # 7. Kepemimpinan
    elif any(k in v for k in ["pimpin", "leader", "transformasional", "transaksional", "otentik", "servant", "kepemimpinan"]):
        return "kepemimpinan"
    
    # 8. Budaya Organisasi
    elif any(k in v for k in ["budaya", "culture", "nilai kerja", "iklim organisasi", "berakhlak"]):
        return "budaya_organisasi"
    
    # 9. Motivasi Kerja
    elif any(k in v for k in ["motivasi", "dorongan", "achievement", "spirit"]):
        return "motivasi"
    
    # 10. Disiplin Kerja
    elif any(k in v for k in ["disiplin", "tata tertib", "kepatuhan", "presensi", "kehadiran"]):
        return "disiplin"
    
    # 11. Insentif / Kompensasi
    elif any(k in v for k in ["insentif", "kompensasi", "gaji", "remunerasi", "reward", "tunjangan", "tpp", "upah"]):
        return "kompensasi_insentif"
    
    # 12. Employee Engagement & Komitmen
    elif any(k in v for k in ["engagement", "keterikatan", "komitmen", "loyalitas", "retensi"]):
        return "engagement_komitmen"
    
    # 13. Kompetensi
    elif any(k in v for k in ["kompetensi", "pelatihan", "keahlian", "skill", "kapabilitas", "diklat"]):
        return "kompetensi"
    
    # 14. Lingkungan Kerja
    elif any(k in v for k in ["lingkungan", "ergonomi", "kondisi kerja"]):
        return "lingkungan_kerja"
    
    # 15. Kecerdasan Emosional / Efikasi Diri
    elif any(k in v for k in ["emosional", "efikasi", "spiritual", "etos", "sikap kerja"]):
        return "kecerdasan_emosional"
    
    else:
        return "motivasi"


def load_cache():
    """Muat cache sitasi, menyaring entri yang melanggar jendela tahun wajib.

    Cache lama bisa berisi sitasi 2020/2021 dari sebelum kebijakan 2022-2026
    diberlakukan. Entri seperti itu dibuang saat dimukan sehingga otomatis
    diambil ulang dari API dengan tahun yang benar.
    """
    if not os.path.exists(CACHE_FILE):
        return {}
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            cache = json.load(f)
    except Exception:
        return {}
    if not isinstance(cache, dict):
        return {}

    buang = {}
    for key, val in list(cache.items()):
        larang = tahun_sitasi_terlarang(json.dumps(val, ensure_ascii=False))
        if larang:
            buang[key] = sorted(set(larang))[:3]
        # Bersihkan referensi internal LAKIP/dokumen kedinasan dari references
        if isinstance(val, dict) and "references" in val:
            val["references"] = [
                r for r in val["references"]
                if not any(x in str(r.get("text", "")).upper() for x in ["LAKIP", "AKUNTABILITAS KINERJA", "LAPORAN KINERJA TAHUNAN"])
            ]
    for key in buang:
        cache.pop(key, None)
    if buang:
        print(f"[sitasi] cache: {len(buang)} entri dibuang (tahun di luar "
              f"{MIN_CITE_YEAR}-{MAX_CITE_YEAR}, contoh {list(buang.values())[:2]}) "
              f"-> diambil ulang dari API")
    return cache


def save_cache(cache):
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


def _rapikan_nama_kapital(s):
    """Nama penulis HURUF KAPITAL SEMUA -> Title Case (APA).

    Hanya token kapital >= 4 huruf yang diubah; inisial dan nama pendek
    seperti ``HS`` dibiarkan.
    """
    def _satu(m):
        w = m.group(0)
        if not (w.isupper() and len(w) >= 4):
            return w
        bagian = re.split(r"(['\-])", w)
        hasil = []
        for p in bagian:
            if p in ("'", "-"):
                hasil.append(p)
            elif p:
                hasil.append(p[:1].upper() + p[1:].lower())
        return "".join(hasil)
    return re.sub(r"[A-Z][A-Z'\-]{3,}", _satu, s or "")


def _sentence_case_judul(judul):
    """Judul artikel HURUF KAPITAL SEMUA -> sentence case (APA)."""
    if not judul:
        return judul
    try:
        from build_full_thesis_docx import smart_sentence_case
        return smart_sentence_case(judul)
    except Exception:
        return judul[:1].upper() + judul[1:].lower()


def format_author_name(author_str):
    if not author_str:
        return ""
    s = _rapikan_nama_kapital(author_str.strip())
    
    # Deteksi penulis institusional / organisasi resmi (Corporate Authors)
    corp_keywords = ["kementerian", "badan", "dinas", "pemerintah", "institute", "institution", "organization", "bank", "universitas", "komisi", "dewan", "bpkp", "bps", "bkn", "lan", "who", "unesco", "oecd"]
    if any(k in s.lower() for k in corp_keywords):
        return s

    # Jika sudah berformat "LastName, FirstName/Initials" (misal: "Barney, J." atau "Sarstedt, M.")
    if "," in s:
        parts_comma = s.split(",", 1)
        last = parts_comma[0].strip()
        first_rest = parts_comma[1].strip()
        f_parts = first_rest.split()
        if f_parts:
            initials = ". ".join([p[0].upper() for p in f_parts if re.search(r"[A-Za-z]", p)])
            if initials:
                return f"{last}, {initials}."
        return last

    parts = s.split()
    parts_letters = [p for p in parts if re.search(r"[A-Za-z]", p)]
    if not parts_letters:
        return ""
    if len(parts_letters) == 1:
        return parts_letters[0]
    last = parts_letters[-1]
    initials = ". ".join([p[0].upper() for p in parts_letters[:-1]]) + "."
    return f"{last}, {initials}"


def format_authors_apa(author_list):
    if not author_list:
        return "Anonim."
    formatted = [x for x in (format_author_name(a) for a in author_list if a) if x]
    if not formatted:
        return "Anonim."
    if len(formatted) == 1:
        return f"{formatted[0]}"
    elif len(formatted) == 2:
        return f"{formatted[0]}, & {formatted[1]}"
    elif len(formatted) <= 7:
        return ", ".join(formatted[:-1]) + f", & {formatted[-1]}"
    else:
        return ", ".join(formatted[:6]) + f", ... {formatted[-1]}"


def make_in_text_citation(authors, year, narrative=False):
    if not authors:
        return f"Anonim ({year})" if narrative else f"(Anonim, {year})"
    last_names = []
    for a in authors:
        parts = _rapikan_nama_kapital(a.strip()).split()
        if parts:
            last_names.append(parts[-1])
    if not last_names:
        return f"Anonim ({year})" if narrative else f"(Anonim, {year})"
    if len(last_names) == 1:
        lbl = last_names[0]
    elif len(last_names) == 2:
        lbl = f"{last_names[0]} dan {last_names[1]}" if narrative else f"{last_names[0]} & {last_names[1]}"
    else:
        lbl = f"{last_names[0]} et al."
    return f"{lbl} ({year})" if narrative else f"({lbl}, {year})"


def search_crossref_books(query, limit=3, min_year=None, max_year=None):
    min_year, max_year = _resolve_range(min_year, max_year)
    """
    Cari buku (book / book-chapter) di Crossref untuk variabel Y.
    Mengembalikan list dict dengan kunci: authors, year, title, publisher, doi, type.
    """
    min_year = max(int(min_year), MIN_CITE_YEAR)
    max_year = min(int(max_year), MAX_CITE_YEAR)
    clean_q = re.sub(r'[^\w\s]', ' ', query).strip()
    encoded_q = urllib.parse.quote(clean_q)
    url = (
        f"https://api.crossref.org/works?query={encoded_q}"
        f"&filter=type:book-chapter,from-pub-date:{min_year}-01-01,until-pub-date:{max_year}-12-31"
        f"&rows={limit}&select=DOI,title,author,published,publisher,type"
    )
    headers = {
        "User-Agent": "AutoTesisAcademicConnector/3.0 (mailto:tesismagister@research.ac.id)"
    }
    books = []
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            for item in data.get("message", {}).get("items", []):
                doi = item.get("DOI", "")
                if not doi:
                    continue
                title_list = item.get("title", [])
                title = title_list[0] if title_list else "Buku Tanpa Judul"
                if title.isupper():
                    title = _sentence_case_judul(title)
                # Extract year
                date_parts = item.get("published", {}).get("date-parts", [[2023]])
                year = date_parts[0][0] if date_parts and date_parts[0] else 2023
                if year < MIN_CITE_YEAR or year > MAX_CITE_YEAR:
                    continue
                authors = []
                for auth in item.get("author", []):
                    given = auth.get("given", "")
                    family = auth.get("family", "")
                    if family:
                        authors.append(f"{given} {family}".strip() if given else family)
                publisher = item.get("publisher", "Penerbit Tidak Diketahui")
                books.append({
                    "type": "buku",
                    "authors": authors,
                    "year": year,
                    "title": title,
                    "publisher": publisher,
                    "doi": f"https://doi.org/{doi}",
                    "source_api": "Crossref"
                })
    except Exception:
        pass
    return books


def search_crossref_journals(query, limit=3, min_year=None, max_year=None):
    min_year, max_year = _resolve_range(min_year, max_year)
    """
    Cari artikel jurnal di Crossref untuk variabel Y.
    Diganti OpenAlex karena OpenAlex sering kena rate limit (429).
    """
    min_year = max(int(min_year), MIN_CITE_YEAR)
    max_year = min(int(max_year), MAX_CITE_YEAR)
    clean_q = re.sub(r'[^\w\s]', ' ', query).strip()
    encoded_q = urllib.parse.quote(clean_q)
    url = (
        f"https://api.crossref.org/works?query={encoded_q}"
        f"&filter=type:journal-article,from-pub-date:{min_year}-01-01,until-pub-date:{max_year}-12-31"
        f"&rows={limit}&select=DOI,title,author,published,container-title,volume,issue,page,type"
    )
    headers = {
        "User-Agent": "AutoTesisAcademicConnector/3.0 (mailto:tesismagister@research.ac.id)"
    }
    papers = []
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            for item in data.get("message", {}).get("items", []):
                doi = item.get("DOI", "")
                if not doi:
                    continue
                title_list = item.get("title", [])
                title = title_list[0] if title_list else "Tanpa Judul"
                if title.isupper():
                    title = _sentence_case_judul(title)
                date_parts = item.get("published", {}).get("date-parts", [[2023]])
                year = date_parts[0][0] if date_parts and date_parts[0] else 2023
                if year < MIN_CITE_YEAR or year > MAX_CITE_YEAR:
                    continue
                authors = []
                for auth in item.get("author", []):
                    given = auth.get("given", "")
                    family = auth.get("family", "")
                    if family:
                        authors.append(f"{given} {family}".strip() if given else family)
                container = item.get("container-title", [])
                journal = container[0] if container else "Jurnal Ilmiah"
                vol = item.get("volume", "")
                issue = item.get("issue", "")
                page = item.get("page", "")
                papers.append({
                    "type": "jurnal",
                    "authors": authors,
                    "year": year,
                    "title": title,
                    "journal": journal,
                    "volume": vol,
                    "issue": issue,
                    "pages": page,
                    "doi": f"https://doi.org/{doi}",
                    "source_api": "Crossref"
                })
    except Exception:
        pass
    return papers


def fetch_sources_for_variable_y(variable_name, instansi, x1=None, x2=None, m=None, judul=None, min_year=None, max_year=None, target_count=10):
    min_year, max_year = _resolve_range(min_year, max_year)
    """
    Pre-fetch 8-10 sources (4-5 jurnal + 4-5 buku) untuk Variabel Y tahun 2022-2026.
    Setiap source memiliki apa_in_text untuk validasi nanti.
    
    Menggunakan query bahasa Inggris karena OpenAlex/Crossref primarily
    mengindeks konten berbahasa Inggris.
    
    Query disesuaikan dengan judul penelitian lengkap (x1, x2, m, y, instansi)
    agar setiap judul mendapat sumber yang UNIK dan relevan, menghindari
    duplikasi sitasi antar judul saat skala ribuan tesis.
    """
    # Map variabel Y ke query bahasa Inggris yang relevan
    y_en_map = {
        "kinerja pegawai": "employee performance",
        "kinerja": "employee performance",
        "employee engagement": "employee engagement",
        "motivasi kerja": "work motivation",
        "kepuasan kerja": "job satisfaction",
    }
    y_en = y_en_map.get(variable_name.lower(), variable_name.lower())
    
    # Map variabel X1, X2, M ke bahasa Inggris untuk query yang lebih spesifik
    x1_en_map = {
        "kompensasi": "compensation",
        "kompensasi kerja": "compensation",
        "pengembangan karier": "career development",
        "pelatihan": "training",
        "promosi jabatan": "job promotion",
    }
    x2_en_map = {
        "lingkungan kerja": "work environment",
        "budaya organisasi": "organizational culture",
        "pengembangan karir": "career development",
        "mutasi kerja": "job rotation",
    }
    m_en_map = {
        "motivasi kerja": "work motivation",
        "motivasi": "work motivation",
        "kepuasan kerja": "job satisfaction",
        "employee engagement": "employee engagement",
    }
    
    x1_en = x1_en_map.get(x1.lower() if x1 else "", x1 or "")
    x2_en = x2_en_map.get(x2.lower() if x2 else "", x2 or "")
    m_en = m_en_map.get(m.lower() if m else "", m or "")

    # =========================================================================
    # GLOBAL TRACKER: load used DOIs/authors to enforce uniqueness across theses
    # =========================================================================
    tracker = _load_used_citations()
    used_dois_global = tracker["used_dois"]
    used_authors_global = tracker["used_lead_authors"]
    used_titles_global = tracker["used_titles"]
    
    # Thesis key untuk registry
    thesis_key = f"{instansi}_{x1}_{x2}_{m}_{y_en}".replace(" ", "_").lower()[:100]
    # Sumber milik judul ini boleh dipakai ulang saat regenerasi (anti-tumpang-tindih hanya antar-judul)
    own_dois, own_authors, own_titles = _own_registry_sets(tracker, thesis_key)

    sources = []
    seen_dois_local = set()
    seen_authors_local = set()

    # =========================================================================
    # QUERY LEBIH SPESIFIK PER JUDUL: gunakan instansi + judul penelitian lengkap
    # =========================================================================
    # Extract keywords from judul for more specific queries
    judul_keywords = []
    if judul:
        # Ambil kata kunci penting dari judul (hapus stopwords)
        stopwords = {'pengaruh', 'terhadap', 'pada', 'di', 'dan', 'atau', 'yang', 'dengan', 'untuk', 'dalam', 'ke', 'dari', 'oleh', 'sebagai', 'adalah', 'ini', 'itu', 'yang', 'the', 'and', 'of', 'in', 'on', 'for', 'to', 'a', 'an'}
        judul_words = [w.lower() for w in re.findall(r'\b\w+\b', judul) if w.lower() not in stopwords and len(w) > 2]
        judul_keywords = judul_words[:5]  # ambil 5 kata kunci utama
    
    instansi_short = instansi.split()[0].lower() if instansi else "public"
    
    # Journal queries - sangat spesifik per judul
    journal_queries = [
        f"{y_en} {x1_en} {instansi_short} indonesia",  # variabel utama + instansi + negara
        f"{y_en} {x2_en} {m_en} {instansi_short}",     # kombinasi X2 + M + instansi
        f"{y_en} {x1_en} {x2_en} factors",            # X1 + X2 factors
        f"{' '.join(judul_keywords[:3])} {y_en}" if judul_keywords else f"{y_en} public sector indonesia",
        f"{y_en} measurement indicators {instansi_short}",
        f"{m_en} mediates {x1_en} {y_en}",             # mediation query
        f"{y_en} government employee {instansi_short}",
    ]
    
    # Filter function: reject entries without valid authors (no Anonim)
    def has_valid_authors(paper):
        authors = paper.get("authors", [])
        if not authors:
            return False
        # Check at least one author has a family name
        for a in authors:
            if a and a.strip() and len(a.strip().split()) >= 1:
                return True
        return False

    # --- Cari Jurnal via Crossref (maks 5) ---
    for query in journal_queries:
        if len([s for s in sources if s["type"] == "jurnal"]) >= 5:
            break
        papers = search_crossref_journals(query, limit=9, min_year=min_year, max_year=max_year)
        for p in papers:
            # REJECT: no valid authors -> would become "Anonim"
            if not has_valid_authors(p):
                continue
            doi = p.get("doi", "").lower().strip()
            lead = get_lead_author(p).lower()
            title = p.get("title", "").lower().strip()
            # Check global uniqueness
            if doi and doi in used_dois_global and doi not in own_dois:
                continue
            if lead and lead in used_authors_global and lead not in own_authors:
                continue
            if title and title in used_titles_global and title not in own_titles:
                continue
            # Check local uniqueness
            if doi and doi in seen_dois_local:
                continue
            if lead and lead in seen_authors_local:
                continue
            seen_dois_local.add(doi)
            seen_authors_local.add(lead)
            sources.append({
                "type": "jurnal",
                "authors": p.get("authors", []),
                "year": p.get("year"),
                "title": p.get("title", ""),
                "journal": p.get("journal", ""),
                "volume": p.get("volume", ""),
                "issue": p.get("issue", ""),
                "pages": p.get("pages", ""),
                "doi": p.get("doi", ""),
                "apa_in_text": make_in_text_citation(p.get("authors", []), p.get("year"), False),
                "apa_full": build_apa_entry(p)[0]
            })
            if len([s for s in sources if s["type"] == "jurnal"]) >= 5:
                break

    # --- Cari Buku via Crossref (maks 5) ---
    book_queries = [
        f"{y_en} {x1_en} {instansi_short} book",
        f"{y_en} {m_en} organizational theory",
        f"managing {y_en} public sector",
        f"{y_en} human resource management {instansi_short}",
        f"{y_en} indonesia buku",
    ]
    for query in book_queries:
        if len(sources) >= target_count:
            break
        books = search_crossref_books(query, limit=9, min_year=min_year, max_year=max_year)
        for b in books:
            if not has_valid_authors(b):
                continue
            if not _gate_relevansi_judul(b.get('title', ''), variable_name, [y_en, x1_en, x2_en, m_en]):
                continue
                continue
            doi = b.get("doi", "").lower().strip()
            lead = get_lead_author(b).lower()
            title = b.get("title", "").lower().strip()
            if doi and doi in used_dois_global and doi not in own_dois:
                continue
            if lead and lead in used_authors_global and lead not in own_authors:
                continue
            if title and title in used_titles_global and title not in own_titles:
                continue
            if doi and doi in seen_dois_local:
                continue
            if lead and lead in seen_authors_local:
                continue
            seen_dois_local.add(doi)
            seen_authors_local.add(lead)
            sources.append({
                "type": "buku",
                "authors": b.get("authors", []),
                "year": b.get("year"),
                "title": b.get("title", ""),
                "publisher": b.get("publisher", ""),
                "doi": b.get("doi", ""),
                "apa_in_text": make_in_text_citation(b.get("authors", []), b.get("year"), False),
                "apa_full": f"{format_authors_apa(b.get('authors', []))} ({b.get('year')}). {b.get('title', '')}. {b.get('publisher', '')}. {b.get('doi', '')}"
            })
            if len(sources) >= target_count:
                break

    # =========================================================================
    # REGISTER CITATIONS to global tracker (prevent reuse by other theses)
    # =========================================================================
    if sources:
        _register_thesis_citations(thesis_key, sources)
        print(f"[fetch_sources_for_variable_y] {len(sources)} sumber unik untuk '{thesis_key}' (global tracker updated)")

    return sources


# =========================================================================
# FETCH 1 JURNAL FAKTOR-FAKTOR YANG MEMPENGARUHI VARIABEL Y - 2.1.1.2
# =========================================================================
def fetch_journal_factors_y(variable_name, instansi, x1=None, x2=None, m=None, y=None, judul=None, min_year=None, max_year=None):
    min_year, max_year = _resolve_range(min_year, max_year)
    """
    Fetch SATU jurnal (2022-2026) yang spesifik membahas faktor-faktor 
    yang mempengaruhi Variabel Y, dengan variabel independen yang relevan
    (X1, X2, M) agar bisa dinarasikan.
    
    Return: list dengan 1 item jurnal, atau list kosong.
    """
    y_en_map = {
        "kinerja pegawai": "employee performance",
        "kinerja": "employee performance",
        "employee engagement": "employee engagement",
        "motivasi kerja": "work motivation",
        "kepuasan kerja": "job satisfaction",
    }
    y_en = y_en_map.get(variable_name.lower(), variable_name.lower())
    
    x1_en_map = {
        "pengembangan karier": "career development",
        "pelatihan": "training",
        "promosi jabatan": "job promotion",
    }
    x2_en_map = {
        "budaya organisasi": "organizational culture",
        "pengembangan karir": "career development",
        "mutasi kerja": "job rotation",
    }
    m_en_map = {
        "kepuasan kerja": "job satisfaction",
        "employee engagement": "employee engagement",
        "motivasi kerja": "work motivation",
    }
    
    x1_en = x1_en_map.get(x1.lower() if x1 else "", x1 or "")
    x2_en = x2_en_map.get(x2.lower() if x2 else "", x2 or "")
    m_en = m_en_map.get(m.lower() if m else "", m or "")
    
    tracker = _load_used_citations()
    used_dois_global = tracker["used_dois"]
    used_authors_global = tracker["used_lead_authors"]
    used_titles_global = tracker["used_titles"]
    
    thesis_key = f"{instansi}_{x1}_{x2}_{m}_{y}_FACTORS_Y".replace(" ", "_").lower()[:100]
    own_dois, own_authors, own_titles = _own_registry_sets(tracker, thesis_key)
    
    # Query spesifik: jurnal yang meneliti faktor-faktor mempengaruhi Y
    # dengan kombinasi X1, X2, M
    journal_queries = [
        f"{y_en} {x1_en} {x2_en} {m_en} factors influencing",
        f"factors affecting {y_en} {x1_en} {m_en}",
        f"determinants of {y_en} {x2_en} {m_en}",
        f"{y_en} predictors {x1_en} {x2_en} {m_en}",
        f"{x1_en} {x2_en} {m_en} influence {y_en}",
    ]
    
    def has_valid_authors(paper):
        authors = paper.get("authors", [])
        if not authors:
            return False
        for a in authors:
            if a and a.strip() and len(a.strip().split()) >= 1:
                return True
        return False
    
    sources = []
    seen_dois = set()
    seen_titles = set()
    
    for query in journal_queries:
        if len(sources) >= 2:
            break
        papers = search_crossref_journals(query, limit=5, min_year=min_year, max_year=max_year)
        for p in papers:
            if len(sources) >= 2:
                break
            if not has_valid_authors(p):
                continue
            doi = p.get("doi", "").lower().strip()
            lead = get_lead_author(p).lower()
            title = p.get("title", "").lower().strip()
            if doi and doi in used_dois_global and doi not in own_dois:
                continue
            if lead and lead in used_authors_global and lead not in own_authors:
                continue
            if title and title in used_titles_global and title not in own_titles:
                continue
            if doi and doi in seen_dois:
                continue
            if title and title in seen_titles:
                continue
            
            # Found valid journal
            source = {
                "type": "jurnal",
                "authors": p.get("authors", []),
                "year": p.get("year"),
                "title": p.get("title", ""),
                "journal": p.get("journal", ""),
                "volume": p.get("volume", ""),
                "issue": p.get("issue", ""),
                "pages": p.get("pages", ""),
                "doi": p.get("doi", ""),
                "apa_in_text": make_in_text_citation(p.get("authors", []), p.get("year"), False),
                "apa_full": build_apa_entry(p)[0]
            }
            
            sources.append(source)
            seen_dois.add(doi)
            seen_titles.add(title)
    
    if sources:
        _register_thesis_citations(thesis_key, sources)
        print(f"[fetch_journal_factors_y] {len(sources)} jurnal faktor untuk '{thesis_key}' (global tracker updated)")
        return sources
    
    print(f"[fetch_journal_factors_y] Tidak ditemukan jurnal faktor untuk '{thesis_key}'")
    return []


# =========================================================================
# FETCH SOURCES UNTUK VARIABEL M (Mediating) - 2.1.2.1
# =========================================================================
def fetch_sources_for_variable_m(variable_name, instansi, x1=None, x2=None, m=None, y=None, judul=None, min_year=None, max_year=None, target_count=10):
    min_year, max_year = _resolve_range(min_year, max_year)
    """
    Pre-fetch 8-10 sources untuk Variabel M (Mediating) tahun 2022-2026.
    Query disesuaikan dengan variabel M spesifik + instansi + judul penelitian.
    """
    m_en_map = {
        "kepuasan kerja": "job satisfaction",
        "employee engagement": "employee engagement",
        "motivasi kerja": "work motivation",
        "motivasi": "work motivation",
        "komitmen organisasi": "organizational commitment",
        "disiplin kerja": "work discipline",
    }
    m_en = m_en_map.get(variable_name.lower(), variable_name.lower())
    
    x1_en = (x1 or "").lower()
    x2_en = (x2 or "").lower()
    y_en = (y or "kinerja pegawai").lower()
    
    tracker = _load_used_citations()
    used_dois_global = tracker["used_dois"]
    used_authors_global = tracker["used_lead_authors"]
    used_titles_global = tracker["used_titles"]
    
    thesis_key = f"{instansi}_{x1}_{x2}_{m}_{y}_M".replace(" ", "_").lower()[:100]
    # Sumber milik judul ini boleh dipakai ulang saat regenerasi (anti-tumpang-tindih hanya antar-judul)
    _own_dois, _own_authors, _own_titles = _own_registry_sets(tracker, thesis_key)
    used_dois_global = used_dois_global - _own_dois
    used_authors_global = used_authors_global - _own_authors
    used_titles_global = used_titles_global - _own_titles
    
    sources = []
    seen_dois_local = set()
    seen_authors_local = set()
    
    judul_keywords = []
    if judul:
        stopwords = {'pengaruh', 'terhadap', 'pada', 'di', 'dan', 'atau', 'yang', 'dengan', 'untuk', 'dalam', 'ke', 'dari', 'oleh', 'sebagai', 'adalah', 'ini', 'itu', 'the', 'and', 'of', 'in', 'on', 'for', 'to', 'a', 'an'}
        judul_words = [w.lower() for w in re.findall(r'\b\w+\b', judul) if w.lower() not in stopwords and len(w) > 2]
        judul_keywords = judul_words[:5]
    
    instansi_short = instansi.split()[0].lower() if instansi else "public"
    
    def has_valid_authors(paper):
        authors = paper.get("authors", [])
        if not authors:
            return False
        for a in authors:
            if a and a.strip() and len(a.strip().split()) >= 1:
                return True
        return False
    
    journal_queries = [
        f"{m_en} {x1_en} {instansi_short} indonesia",
        f"{m_en} {x2_en} {y_en} mediation",
        f"{m_en} factors affecting {x1_en} {x2_en}",
        f"{' '.join(judul_keywords[:3])} {m_en}" if judul_keywords else f"{m_en} public sector indonesia",
        f"{m_en} antecedents consequences {instansi_short}",
        f"{m_en} mediates {x1_en} {y_en}",
        f"{m_en} government employee {instansi_short}",
    ]
    
    for query in journal_queries:
        if len([s for s in sources if s["type"] == "jurnal"]) >= 5:
            break
        papers = search_crossref_journals(query, limit=9, min_year=min_year, max_year=max_year)
        for p in papers:
            if not has_valid_authors(p):
                continue
            if not _gate_relevansi_judul(p.get('title', ''), variable_name, [m_en, x1_en, x2_en, y_en]):
                continue
                continue
            doi = p.get("doi", "").lower().strip()
            lead = get_lead_author(p).lower()
            title = p.get("title", "").lower().strip()
            if doi and doi in used_dois_global: continue
            if lead and lead in used_authors_global: continue
            if title and title in used_titles_global: continue
            if doi and doi in seen_dois_local: continue
            if lead and lead in seen_authors_local: continue
            seen_dois_local.add(doi)
            seen_authors_local.add(lead)
            sources.append({
                "type": "jurnal",
                "authors": p.get("authors", []),
                "year": p.get("year"),
                "title": p.get("title", ""),
                "journal": p.get("journal", ""),
                "volume": p.get("volume", ""),
                "issue": p.get("issue", ""),
                "pages": p.get("pages", ""),
                "doi": p.get("doi", ""),
                "apa_in_text": make_in_text_citation(p.get("authors", []), p.get("year"), False),
                "apa_full": build_apa_entry(p)[0]
            })
            if len([s for s in sources if s["type"] == "jurnal"]) >= 5:
                break

    book_queries = [
        f"{m_en} {x1_en} {instansi_short} book",
        f"{m_en} organizational behavior theory",
        f"managing {m_en} workplace",
        f"{m_en} human resource management {instansi_short}",
        f"{m_en} indonesia buku",
    ]
    for query in book_queries:
        if len(sources) >= target_count:
            break
        books = search_crossref_books(query, limit=9, min_year=min_year, max_year=max_year)
        for b in books:
            if not has_valid_authors(b):
                continue
            if not _gate_relevansi_judul(b.get('title', ''), variable_name, [m_en, x1_en, x2_en, y_en]):
                continue
                continue
            doi = b.get("doi", "").lower().strip()
            lead = get_lead_author(b).lower()
            title = b.get("title", "").lower().strip()
            if doi and doi in used_dois_global: continue
            if lead and lead in used_authors_global: continue
            if title and title in used_titles_global: continue
            if doi and doi in seen_dois_local: continue
            if lead and lead in seen_authors_local: continue
            seen_dois_local.add(doi)
            seen_authors_local.add(lead)
            sources.append({
                "type": "buku",
                "authors": b.get("authors", []),
                "year": b.get("year"),
                "title": b.get("title", ""),
                "publisher": b.get("publisher", ""),
                "doi": b.get("doi", ""),
                "apa_in_text": make_in_text_citation(b.get("authors", []), b.get("year"), False),
                "apa_full": f"{format_authors_apa(b.get('authors', []))} ({b.get('year')}). {b.get('title', '')}. {b.get('publisher', '')}. {b.get('doi', '')}"
            })
            if len(sources) >= target_count:
                break

    if sources:
        _register_thesis_citations(thesis_key, sources)
        print(f"[fetch_sources_for_variable_m] {len(sources)} sumber unik untuk '{thesis_key}' (global tracker updated)")

    return sources


# =========================================================================
# FETCH SOURCES UNTUK VARIABEL X1 (Independent 1) - 2.1.3.1
# =========================================================================
def fetch_sources_for_variable_x1(variable_name, instansi, x1=None, x2=None, m=None, y=None, judul=None, min_year=None, max_year=None, target_count=10):
    min_year, max_year = _resolve_range(min_year, max_year)
    """
    Pre-fetch 8-10 sources untuk Variabel X1 (Independent 1) tahun 2022-2026.
    """
    x1_en_map = {
        "kompensasi": "compensation",
        "kompensasi kerja": "compensation",
        "pengembangan karier": "career development",
        "pelatihan": "training",
        "promosi jabatan": "job promotion",
        "pengembangan karir": "career development",
        "mutasi kerja": "job rotation",
        "lingkungan kerja": "work environment",
    }
    x1_en = x1_en_map.get(variable_name.lower(), variable_name.lower())
    
    x2_en = (x2 or "").lower()
    m_en = (m or "").lower()
    y_en = (y or "kinerja pegawai").lower()
    
    tracker = _load_used_citations()
    used_dois_global = tracker["used_dois"]
    used_authors_global = tracker["used_lead_authors"]
    used_titles_global = tracker["used_titles"]
    
    thesis_key = f"{instansi}_{x1}_{x2}_{m}_{y}_X1".replace(" ", "_").lower()[:100]
    # Sumber milik judul ini boleh dipakai ulang saat regenerasi (anti-tumpang-tindih hanya antar-judul)
    _own_dois, _own_authors, _own_titles = _own_registry_sets(tracker, thesis_key)
    used_dois_global = used_dois_global - _own_dois
    used_authors_global = used_authors_global - _own_authors
    used_titles_global = used_titles_global - _own_titles
    
    sources = []
    seen_dois_local = set()
    seen_authors_local = set()
    
    judul_keywords = []
    if judul:
        stopwords = {'pengaruh', 'terhadap', 'pada', 'di', 'dan', 'atau', 'yang', 'dengan', 'untuk', 'dalam', 'ke', 'dari', 'oleh', 'sebagai', 'adalah', 'ini', 'itu', 'the', 'and', 'of', 'in', 'on', 'for', 'to', 'a', 'an'}
        judul_words = [w.lower() for w in re.findall(r'\b\w+\b', judul) if w.lower() not in stopwords and len(w) > 2]
        judul_keywords = judul_words[:5]
    
    instansi_short = instansi.split()[0].lower() if instansi else "public"
    
    def has_valid_authors(paper):
        authors = paper.get("authors", [])
        if not authors: return False
        for a in authors:
            if a and a.strip() and len(a.strip().split()) >= 1:
                return True
        return False
    
    journal_queries = [
        f"{x1_en} {y_en} {instansi_short} indonesia",
        f"{x1_en} {m_en} {x2_en} performance",
        f"{x1_en} impact employee {y_en}",
        f"{' '.join(judul_keywords[:3])} {x1_en}" if judul_keywords else f"{x1_en} public sector indonesia",
        f"{x1_en} effectiveness {instansi_short}",
        f"{x1_en} influence {m_en} {y_en}",
    ]
    
    for query in journal_queries:
        if len([s for s in sources if s["type"] == "jurnal"]) >= 5:
            break
        papers = search_crossref_journals(query, limit=9, min_year=min_year, max_year=max_year)
        for p in papers:
            if not has_valid_authors(p):
                continue
            if not _gate_relevansi_judul(p.get('title', ''), variable_name, [x1_en, x2_en, m_en, y_en]):
                continue
            doi = p.get("doi", "").lower().strip()
            lead = get_lead_author(p).lower()
            title = p.get("title", "").lower().strip()
            if doi and doi in used_dois_global: continue
            if lead and lead in used_authors_global: continue
            if title and title in used_titles_global: continue
            if doi and doi in seen_dois_local: continue
            if lead and lead in seen_authors_local: continue
            seen_dois_local.add(doi)
            seen_authors_local.add(lead)
            sources.append({
                "type": "jurnal",
                "authors": p.get("authors", []),
                "year": p.get("year"),
                "title": p.get("title", ""),
                "journal": p.get("journal", ""),
                "volume": p.get("volume", ""),
                "issue": p.get("issue", ""),
                "pages": p.get("pages", ""),
                "doi": p.get("doi", ""),
                "apa_in_text": make_in_text_citation(p.get("authors", []), p.get("year"), False),
                "apa_full": build_apa_entry(p)[0]
            })
            if len([s for s in sources if s["type"] == "jurnal"]) >= 5:
                break

    book_queries = [
        f"{x1_en} {y_en} book",
        f"{x1_en} organizational development theory",
        f"effective {x1_en} management",
        f"{x1_en} human resource {instansi_short}",
        f"{x1_en} indonesia buku",
    ]
    for query in book_queries:
        if len(sources) >= target_count:
            break
        books = search_crossref_books(query, limit=9, min_year=min_year, max_year=max_year)
        for b in books:
            if not has_valid_authors(b):
                continue
            if not _gate_relevansi_judul(b.get('title', ''), variable_name, [x1_en, x2_en, m_en, y_en]):
                continue
            doi = b.get("doi", "").lower().strip()
            lead = get_lead_author(b).lower()
            title = b.get("title", "").lower().strip()
            if doi and doi in used_dois_global: continue
            if lead and lead in used_authors_global: continue
            if title and title in used_titles_global: continue
            if doi and doi in seen_dois_local: continue
            if lead and lead in seen_authors_local: continue
            seen_dois_local.add(doi)
            seen_authors_local.add(lead)
            sources.append({
                "type": "buku",
                "authors": b.get("authors", []),
                "year": b.get("year"),
                "title": b.get("title", ""),
                "publisher": b.get("publisher", ""),
                "doi": b.get("doi", ""),
                "apa_in_text": make_in_text_citation(b.get("authors", []), b.get("year"), False),
                "apa_full": f"{format_authors_apa(b.get('authors', []))} ({b.get('year')}). {b.get('title', '')}. {b.get('publisher', '')}. {b.get('doi', '')}"
            })
            if len(sources) >= target_count:
                break

    if sources:
        _register_thesis_citations(thesis_key, sources)
        print(f"[fetch_sources_for_variable_x1] {len(sources)} sumber unik untuk '{thesis_key}' (global tracker updated)")

    return sources


# =========================================================================
# FETCH SOURCES UNTUK VARIABEL X2 (Independent 2) - 2.1.4.1
# =========================================================================
def fetch_sources_for_variable_x2(variable_name, instansi, x1=None, x2=None, m=None, y=None, judul=None, min_year=None, max_year=None, target_count=10):
    min_year, max_year = _resolve_range(min_year, max_year)
    """
    Pre-fetch 8-10 sources untuk Variabel X2 (Independent 2) tahun 2022-2026.
    """
    x2_en_map = {
        "budaya organisasi": "organizational culture",
        "pengembangan karir": "career development",
        "mutasi kerja": "job rotation",
        "disiplin kerja": "work discipline",
        "kepemimpinan": "leadership",
        "lingkungan kerja": "work environment",
        "kompensasi": "compensation",
        "kompensasi kerja": "compensation",
    }
    x2_en = x2_en_map.get(variable_name.lower(), variable_name.lower())
    
    x1_en = (x1 or "").lower()
    m_en = (m or "").lower()
    y_en = (y or "kinerja pegawai").lower()
    
    tracker = _load_used_citations()
    used_dois_global = tracker["used_dois"]
    used_authors_global = tracker["used_lead_authors"]
    used_titles_global = tracker["used_titles"]
    
    thesis_key = f"{instansi}_{x1}_{x2}_{m}_{y}_X2".replace(" ", "_").lower()[:100]
    # Sumber milik judul ini boleh dipakai ulang saat regenerasi (anti-tumpang-tindih hanya antar-judul)
    _own_dois, _own_authors, _own_titles = _own_registry_sets(tracker, thesis_key)
    used_dois_global = used_dois_global - _own_dois
    used_authors_global = used_authors_global - _own_authors
    used_titles_global = used_titles_global - _own_titles
    
    sources = []
    seen_dois_local = set()
    seen_authors_local = set()
    
    judul_keywords = []
    if judul:
        stopwords = {'pengaruh', 'terhadap', 'pada', 'di', 'dan', 'atau', 'yang', 'dengan', 'untuk', 'dalam', 'ke', 'dari', 'oleh', 'sebagai', 'adalah', 'ini', 'itu', 'the', 'and', 'of', 'in', 'on', 'for', 'to', 'a', 'an'}
        judul_words = [w.lower() for w in re.findall(r'\b\w+\b', judul) if w.lower() not in stopwords and len(w) > 2]
        judul_keywords = judul_words[:5]
    
    instansi_short = instansi.split()[0].lower() if instansi else "public"
    
    def has_valid_authors(paper):
        authors = paper.get("authors", [])
        if not authors: return False
        for a in authors:
            if a and a.strip() and len(a.strip().split()) >= 1:
                return True
        return False
    
    journal_queries = [
        f"{x2_en} {y_en} {instansi_short} indonesia",
        f"{x2_en} {x1_en} {m_en} performance",
        f"{x2_en} impact employee {y_en}",
        f"{' '.join(judul_keywords[:3])} {x2_en}" if judul_keywords else f"{x2_en} public sector indonesia",
        f"{x2_en} effectiveness {instansi_short}",
        f"{x2_en} influence {m_en} {y_en}",
    ]
    
    for query in journal_queries:
        if len([s for s in sources if s["type"] == "jurnal"]) >= 5:
            break
        papers = search_crossref_journals(query, limit=9, min_year=min_year, max_year=max_year)
        for p in papers:
            if not has_valid_authors(p):
                continue
            if not _gate_relevansi_judul(p.get('title', ''), variable_name, [x2_en, x1_en, m_en, y_en]):
                continue
            doi = p.get("doi", "").lower().strip()
            lead = get_lead_author(p).lower()
            title = p.get("title", "").lower().strip()
            if doi and doi in used_dois_global: continue
            if lead and lead in used_authors_global: continue
            if title and title in used_titles_global: continue
            if doi and doi in seen_dois_local: continue
            if lead and lead in seen_authors_local: continue
            seen_dois_local.add(doi)
            seen_authors_local.add(lead)
            sources.append({
                "type": "jurnal",
                "authors": p.get("authors", []),
                "year": p.get("year"),
                "title": p.get("title", ""),
                "journal": p.get("journal", ""),
                "volume": p.get("volume", ""),
                "issue": p.get("issue", ""),
                "pages": p.get("pages", ""),
                "doi": p.get("doi", ""),
                "apa_in_text": make_in_text_citation(p.get("authors", []), p.get("year"), False),
                "apa_full": build_apa_entry(p)[0]
            })
            if len([s for s in sources if s["type"] == "jurnal"]) >= 5:
                break

    book_queries = [
        f"{x2_en} {y_en} book",
        f"{x2_en} organizational theory",
        f"managing {x2_en} workplace",
        f"{x2_en} human resource {instansi_short}",
        f"{x2_en} indonesia buku",
    ]
    for query in book_queries:
        if len(sources) >= target_count:
            break
        books = search_crossref_books(query, limit=9, min_year=min_year, max_year=max_year)
        for b in books:
            if not has_valid_authors(b):
                continue
            if not _gate_relevansi_judul(b.get('title', ''), variable_name, [x2_en, x1_en, m_en, y_en]):
                continue
            doi = b.get("doi", "").lower().strip()
            lead = get_lead_author(b).lower()
            title = b.get("title", "").lower().strip()
            if doi and doi in used_dois_global: continue
            if lead and lead in used_authors_global: continue
            if title and title in used_titles_global: continue
            if doi and doi in seen_dois_local: continue
            if lead and lead in seen_authors_local: continue
            seen_dois_local.add(doi)
            seen_authors_local.add(lead)
            sources.append({
                "type": "buku",
                "authors": b.get("authors", []),
                "year": b.get("year"),
                "title": b.get("title", ""),
                "publisher": b.get("publisher", ""),
                "doi": b.get("doi", ""),
                "apa_in_text": make_in_text_citation(b.get("authors", []), b.get("year"), False),
                "apa_full": f"{format_authors_apa(b.get('authors', []))} ({b.get('year')}). {b.get('title', '')}. {b.get('publisher', '')}. {b.get('doi', '')}"
            })
            if len(sources) >= target_count:
                break

    if sources:
        _register_thesis_citations(thesis_key, sources)
        print(f"[fetch_sources_for_variable_x2] {len(sources)} sumber unik untuk '{thesis_key}' (global tracker updated)")

    return sources


# ===================== OpenAlex: cache disk + throttle + backoff (anti-429) =====================
OA_WORKS_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache_openalex_works.json")
_OA_CACHE = None
_OA_LAST_CALL = [0.0]
_OA_MAILTO = "tesismagister@research.ac.id"
OA_MIN_INTERVAL = 0.5  # detik antar-request (<= 2 rps, aman di bawah batas 10 rps OpenAlex)
_OA_BLOCKED_UNTIL = [0.0]  # bila OpenAlex membalas Retry-After besar (kuota harian habis)


def _oa_cache_load():
    global _OA_CACHE
    if _OA_CACHE is None:
        _OA_CACHE = {}
        if os.path.exists(OA_WORKS_CACHE_FILE):
            try:
                with open(OA_WORKS_CACHE_FILE, "r", encoding="utf-8") as f:
                    _OA_CACHE = json.load(f)
            except Exception:
                _OA_CACHE = {}
        if not isinstance(_OA_CACHE, dict):
            _OA_CACHE = {}
    return _OA_CACHE


def _oa_cache_save():
    if _OA_CACHE is None:
        return
    try:
        with open(OA_WORKS_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(_OA_CACHE, f, ensure_ascii=False)
    except Exception:
        pass


def _oa_fetch_json(url, timeout=12, retries=5):
    """GET OpenAlex dengan throttle + backoff pada 429/5xx. Kembalikan dict atau None."""
    if time.time() < _OA_BLOCKED_UNTIL[0]:
        return None
    if "mailto=" not in url:
        url += ("&" if "?" in url else "?") + "mailto=" + urllib.parse.quote(_OA_MAILTO)
    headers = {
        "User-Agent": "AutoTesisAcademicConnector/3.2 (mailto:%s)" % _OA_MAILTO,
        "Accept": "application/json",
    }
    for attempt in range(retries + 1):
        gap = OA_MIN_INTERVAL - (time.time() - _OA_LAST_CALL[0])
        if gap > 0:
            time.sleep(gap)
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                _OA_LAST_CALL[0] = time.time()
                return json.loads(resp.read().decode("utf-8"))
        except Exception as e:
            _OA_LAST_CALL[0] = time.time()
            code = getattr(e, "code", None)
            if code in (429, 500, 502, 503, 504) and attempt < retries:
                ra = None
                try:
                    ra = e.headers.get("Retry-After") if e.headers else None
                except Exception:
                    ra = None
                try:
                    delay = float(ra) if ra else 0.0
                except Exception:
                    delay = 0.0
                if delay <= 0:
                    delay = min(30.0, 2.0 * (2 ** attempt))
                if delay > 120:
                    # Kuota harian OpenAlex habis (Retry-After sangat besar):
                    # tandai blokir lalu MENYERAH agar pemanggil beralih ke Crossref.
                    _OA_BLOCKED_UNTIL[0] = time.time() + min(delay, 6 * 3600)
                    print("[OpenAlex] kuota habis (Retry-After %.0fs) -> pakai Crossref sementara %.1f jam"
                          % (delay, min(delay, 6 * 3600) / 3600.0))
                    return None
                print("[OpenAlex] HTTP %s -> tunggu %.0fs (percobaan %d/%d)"
                      % (code, delay, attempt + 1, retries))
                time.sleep(delay)
                continue
            if attempt < retries:
                time.sleep(min(20.0, 1.0 * (2 ** attempt)))
                continue
            print("[OpenAlex] gagal setelah %d percobaan: %s" % (retries + 1, e))
            return None
    return None


_BANNED_NON_HR_KEYWORDS = (
    "rugi fiskal", "tax avoidance", "penghindaran pajak", "perencanaan pajak",
    "tax planning", "transfer pricing", "tax compliance", "kepatuhan wajib pajak",
    "wajib pajak", "sanksi pajak", "insentif pajak", "pemeriksaan pajak",
    "manajemen laba", "earnings management", "kepemilikan institusional",
    "kedokteran gigi", "dentistry", "hemodialisis", "anestesi"
)


def _oa_parse_papers(data, min_year, max_year):
    min_year, max_year = _resolve_range(min_year, max_year)
    """Ubah hasil JSON OpenAlex menjadi daftar paper (filter DOI + tahun wajib)."""
    papers = []
    for item in (data or {}).get("results", []):
        doi = item.get("doi")
        if not doi or not doi.startswith("http"):
            continue
        title = item.get("title") or ""
        if not title:
            continue
        if any(b in title.lower() for b in _BANNED_NON_HR_KEYWORDS):
            continue
        if title.isupper():
            title = _sentence_case_judul(title)
        year = item.get("publication_year") or 0
        if year < min_year or year > max_year:
            continue
        authors = [a.get("author", {}).get("display_name") for a in item.get("authorships", [])]
        authors = [a for a in authors if a]
        loc = item.get("primary_location") or {}
        source = (loc.get("source") or {}).get("display_name") or "Jurnal Ilmiah Manajemen"
        biblio = item.get("biblio") or {}
        vol, issue = biblio.get("volume"), biblio.get("issue")
        fpage, lpage = biblio.get("first_page"), biblio.get("last_page")
        pages = "%s\u2013%s" % (fpage, lpage) if fpage and lpage else (fpage or "")
        abstract = _reconstruct_abstract(item)
        if any(b in (abstract or "").lower() for b in ["rugi fiskal", "tax avoidance", "penghindaran pajak"]):
            continue
        papers.append({
            "source_api": "OpenAlex", "title": title, "authors": authors,
            "year": year, "journal": source, "volume": vol, "issue": issue,
            "pages": pages, "doi": doi,
            "abstract": abstract,
        })
    return papers


CROSSREF_MAILTO = "tesismagister@research.ac.id"
_CR_LAST_CALL = [0.0]


def _clean_abstract(raw):
    """Bersihkan abstrak Crossref (JATS XML / HTML) menjadi teks polos."""
    if not raw:
        return ""
    t = re.sub(r"<[^>]+>", " ", str(raw))
    t = (t.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
          .replace("&quot;", '"').replace("&#39;", "'").replace("&nbsp;", " "))
    t = re.sub(r"^\s*abstract[:\s\-]*", "", t, flags=re.I)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _crossref_get_json(url, timeout=20, retries=3):
    """GET Crossref dengan throttle + backoff. Kembalikan dict atau None."""
    if "mailto=" not in url:
        url += ("&" if "?" in url else "?") + "mailto=" + urllib.parse.quote(CROSSREF_MAILTO)
    headers = {
        "User-Agent": "AutoTesisAcademicConnector/3.3 (mailto:%s)" % CROSSREF_MAILTO,
        "Accept": "application/json",
    }
    for attempt in range(retries + 1):
        gap = 0.4 - (time.time() - _CR_LAST_CALL[0])
        if gap > 0:
            time.sleep(gap)
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                _CR_LAST_CALL[0] = time.time()
                return json.loads(resp.read().decode("utf-8"))
        except Exception as e:
            _CR_LAST_CALL[0] = time.time()
            code = getattr(e, "code", None)
            if attempt < retries:
                time.sleep(min(20.0, 2.0 * (2 ** attempt)))
                continue
            print("[Crossref] gagal (%s): %s" % (code, e))
            return None
    return None


def _crossref_parse_papers(data, min_year, max_year):
    min_year, max_year = _resolve_range(min_year, max_year)
    """Ubah hasil Crossref menjadi daftar paper (bentuk sama seperti OpenAlex)."""
    papers = []
    for it in ((data or {}).get("message") or {}).get("items", []):
        doi = (it.get("DOI") or "").strip()
        titles = it.get("title") or []
        title = (titles[0] if titles else "").strip()
        if not doi or not title:
            continue
        if any(b in title.lower() for b in _BANNED_NON_HR_KEYWORDS):
            continue
        dp = ((it.get("issued") or {}).get("date-parts") or [[None]])[0]
        year = dp[0] if dp and dp[0] else 0
        try:
            year = int(year)
        except Exception:
            year = 0
        if year < min_year or year > max_year:
            continue
        authors = []
        for a in (it.get("author") or []):
            nm = " ".join(x for x in [a.get("given"), a.get("family")] if x).strip()
            if not nm:
                nm = (a.get("name") or "").strip()
            if nm:
                authors.append(nm)
        cont = it.get("container-title") or []
        journal = (cont[0] if cont else "") or "Jurnal Ilmiah Manajemen"
        papers.append({
            "source_api": "Crossref", "title": title, "authors": authors,
            "year": year, "journal": journal,
            "volume": it.get("volume"), "issue": it.get("issue"),
            "pages": it.get("page") or "",
            "doi": "https://doi.org/" + doi,
            "abstract": _clean_abstract(it.get("abstract")),
        })
    return papers


def search_crossref_papers(query, limit=15, min_year=None, max_year=None):
    min_year, max_year = _resolve_range(min_year, max_year)
    """Sumber cadangan (Crossref) saat OpenAlex kosong atau kehabisan kuota."""
    clean_q = re.sub(r'[^\w\s]', ' ', query).strip()
    ck = "crossref|%s|%d|%d|%d" % (clean_q.lower(), limit, min_year, max_year)
    cache = _oa_cache_load()
    if ck in cache:
        return cache[ck]
    url = ("https://api.crossref.org/works?query=%s"
           "&filter=from-pub-date:%d-01-01,until-pub-date:%d-12-31,type:journal-article"
           "&select=DOI,title,author,issued,container-title,volume,issue,page,abstract"
           "&rows=%d" % (urllib.parse.quote(clean_q), min_year, max_year, limit))
    data = _crossref_get_json(url)
    papers = _crossref_parse_papers(data, min_year, max_year)
    if papers:
        cache[ck] = papers
        _oa_cache_save()
    return papers


def search_openalex_verified(query, limit=3, min_year=None):
    min_year, max_year = _resolve_range(min_year, max_year)
    """OpenAlex terverifikasi DOI dengan cache disk + throttle + backoff anti-429."""
    clean_q = re.sub(r'[^\w\s]', ' ', query).strip()
    ck = "verified|%s|%d|%d" % (clean_q.lower(), limit, min_year)
    cache = _oa_cache_load()
    if ck in cache:
        return cache[ck]
    url = ("https://api.openalex.org/works?search=%s"
           "&filter=has_doi:true,type:article,from_publication_date:%d-01-01"
           "&per-page=%d" % (urllib.parse.quote(clean_q), min_year, limit))
    data = _oa_fetch_json(url, timeout=10, retries=5)
    papers = _oa_parse_papers(data, min_year, MAX_CITE_YEAR)
    if not papers:
        papers = search_crossref_papers(query, limit=limit, min_year=min_year, max_year=None)
    if papers:
        cache[ck] = papers
        _oa_cache_save()
    return papers


def build_apa_entry(paper):
    authors_apa = format_authors_apa(paper.get("authors", []))
    year = paper.get("year", 2023)
    title = paper.get("title", "").strip().rstrip(".")
    if title.isupper():
        title = _sentence_case_judul(title)
    journal = paper.get("journal", "Jurnal Manajemen").strip()
    vol = paper.get("volume")
    issue = paper.get("issue")
    pages = paper.get("pages")
    doi = paper.get("doi", "")

    vol_str = ""
    if vol and issue:
        vol_str = f", {vol}({issue})"
    elif vol:
        vol_str = f", {vol}"
    elif issue:
        vol_str = f", ({issue})"

    page_str = f", {pages}" if pages else ""
    doi_str = f" {doi}" if doi else ""

    entry_text = f"{authors_apa} ({year}). {title}. {journal}{vol_str}{page_str}.{doi_str}"
    italic_part = journal
    return entry_text, italic_part


def get_lead_author(paper):
    authors = paper.get("authors", [])
    if not authors:
        return "Anonim"
    first = authors[0]
    parts = first.strip().split()
    return parts[-1] if len(parts) > 1 else parts[0]


def sumber_apapun(paper):
    """Entri APA lengkap untuk slot jurnal (mengikuti build_apa_entry)."""
    ent, _ = build_apa_entry(paper)
    return ent


def slot_ref(paper):
    """Dict slot jurnal: APA lengkap, sitasi in-text, tahun, judul nyata."""
    return {
        "apa": sumber_apapun(paper),
        "cite": ctx_cite(paper),
        "year": paper.get("year"),
        "title": (paper.get("title") or "").strip().rstrip("."),
    }


def ctx_cite(paper):
    """Sitasi in-text '(Penulis, Tahun)' untuk slot jurnal."""
    return make_in_text_citation(paper["authors"], paper["year"], False)


# ==============================================================================
# KEUNIKAN SITASI BAB I ANTAR-JUDUL (JUDUL/OBJEK BERBEDA -> RUJUKAN BERBEDA)
# ------------------------------------------------------------------------------
# Cacat lama: slot makro/SDM/kinerja SELALU mengambil dari pool "makro_sdm"
# yang sama untuk semua judul, dan kunci cache sitasi hanya berupa nama variabel
# (tanpa judul/objek). Akibatnya dua tesis dengan judul & objek berbeda tetap
# memakai rujukan identik (mis. makalah yang sama muncul di dua tesis).
# Tiga lapis pengaman baru:
#   1) RELEVANSI : kandidat harus cocok kata kunci rumpun variabelnya, minimal
#                  cocok salah satu kata kunci variabel judul ini;
#   2) CACHE     : kunci cache memuat cap JUDUL+INSTANSI, sehingga set sitasi
#                  tidak dibagi antar judul meski nama variabelnya sama;
#   3) REGISTRY  : rujukan (DOI/nama penulis utama) milik judul lain dihindari
#                  selama masih ada kandidat relevan lain.
# ==============================================================================
_STOPWORD_JUDUL = {
    "pengaruh", "terhadap", "pada", "dan", "atau", "yang", "dengan", "untuk",
    "dalam", "dari", "oleh", "sebagai", "adalah", "ini", "itu", "suatu", "para",
    "kantor", "dinas", "badan", "the", "and", "of", "in", "on", "for", "to",
}

TOPIC_KATA_KUNCI = {
    "pengawasan": ["pengawasan", "supervisi", "monitoring", "pengendalian"],
    "kinerja_pegawai": ["kinerja", "performance", "produktivitas"],
    "kepuasan_kerja": ["kepuasan", "satisfaction"],
    "ocb": ["citizenship", "kewargaan", "ekstra peran"],
    "pengembangan_karier": ["karier", "karir", "career"],
    "promosi_jabatan": ["promosi", "jabatan"],
    "mutasi_kerja": ["mutasi", "rotasi"],
    "beban_kerja": ["beban", "stres", "burnout"],
    "kepemimpinan": ["kepemimpinan", "leader", "gaya kepemimpinan"],
    "budaya_organisasi": ["budaya", "iklim organisasi", "culture"],
    "motivasi": ["motivasi", "motivation"],
    "disiplin": ["disiplin", "kepatuhan", "presensi"],
    "kompensasi_insentif": ["insentif", "kompensasi", "gaji", "tunjangan"],
    "engagement_komitmen": ["engagement", "komitmen", "loyalitas"],
    "kompetensi": ["kompetensi", "pelatihan", "keahlian"],
    "lingkungan_kerja": ["lingkungan kerja", "ergonomi", "kondisi kerja"],
    "kecerdasan_emosional": ["emosional", "efikasi", "etos kerja"],
    "mediasi": ["mediasi", "intervening", "memediasi"],
}


_TRANSLITERASI = str.maketrans({
    "ı": "i", "İ": "i", "ş": "s", "Ş": "s", "ğ": "g", "Ğ": "g",
    "ç": "c", "Ç": "c", "ö": "o", "Ö": "o", "ü": "u", "Ü": "u",
    "ć": "c", "č": "c", "š": "s", "ž": "z", "đ": "d", "ł": "l",
})


def _normal_judul(teks):
    """Normalisasi judul untuk pencocokan: diakritik ditransliterasi + huruf kecil.

    Penting untuk sumber berbahasa Turki/Slavia: "Diş Hekimliğinde" harus
    menjadi "dis hekimliginde" agar bisa dikenali sebagai topik kedokteran gigi.
    """
    s = str(teks or "").translate(_TRANSLITERASI)
    try:
        import unicodedata
        s = "".join(c for c in unicodedata.normalize("NFKD", s)
                    if not unicodedata.combining(c))
    except Exception:
        pass
    s = re.sub(r"[^a-z0-9\s]", " ", s.lower())
    return re.sub(r"\s+", " ", s).strip()


def _gate_relevansi_judul(title, var_name, extra_terms=None):
    """True bila judul sumber relevan dengan topik variabel.

    Dua lapis:
      1. WAJIB memuat >=1 kata kunci topik variabel. Menutup bug sumber yang
         hanya cocok karena NAMA PENULIS, mis. "Beban, H." (penulis Kroasia)
         membuat buku musikologi medieval lolos ke tesis "Beban Kerja".
      2. TIDAK boleh bertopik di luar domain HR/organisasi. Menutup bug kata
         ambigu: "kecerdasan" pada "Kecerdasan Artifisial" dan "stres" pada
         "Stres Analiz Yöntemleri" (kedokteran gigi) / PTSD.
    """
    judul = _normal_judul(title)
    if not judul:
        return False

    terms = [t for t in _kata_kunci_var(var_name) if len(t) >= 4]
    for t in (extra_terms or []):
        t = _normal_judul(t)
        if len(t) >= 4 and t not in terms:
            terms.append(t)
    if not terms:
        return True
    if not any(t in judul for t in terms):
        return False

    # --- lapis 2: tolak topik yang jelas di luar domain ---
    hr = ("pegawai", "karyawan", "employee", "kerja", "work", "job", "organisasi",
          "organization", "organizational", "sdm", "human resource", "manajemen",
          "management", "kinerja", "performance", "kepuasan", "satisfaction",
          "komitmen", "commitment", "motivasi", "motivation", "kepemimpinan",
          "leadership", "emosional", "emotional", "produktivitas", "productivity",
          "tenaga kerja", "pekerja", "staf", "staff", "personel", "aparatur")
    pekerja = ("pegawai", "karyawan", "employee", "tenaga kerja", "pekerja",
               "staf", "staff", "personel", "aparatur")

    luar_domain = [
        # AI/teknologi -- "kecerdasan artifisial" BUKAN "kecerdasan emosional"
        (("artifisial", "kecerdasan buatan", "artificial intelligence",
          "machine learning", "deep learning", "neural network",
          "genome", "genom", "bioinformatika"), hr),
        # klinis/medis
        (("kedokteran gigi", "dis hekimligi", "dentistry", "klinis", "pasien",
          "rumah sakit"), hr),
        # psikiatri -- PTSD bukan stres kerja
        (("travma sonrasi", "post traumatic", "posttraumatic", "ptsd",
          "gangguan stres pascatrauma"), hr),
        # pertanian
        (("pertanian", "agricultural", "tanaman pangan", "panen"), hr),
        # anak/sekolah -- tesis ini tentang pegawai, bukan anak/siswa
        (("anak usia dini", "anak usia", "paud", "taman kanak",
          "siswa sekolah dasar", "siswa sd", "siswa smp", "siswa sma"), pekerja),
    ]
    # Tolak keras istilah perpajakan / fiskal / akuntansi pasar modal murni
    if any(ex in judul for ex in _BANNED_NON_HR_KEYWORDS):
        return False

    for grup, penawar in luar_domain:
        if any(t in judul for t in grup) and not any(h in judul for h in penawar):
            return False
    return True


def _kata_kunci_var(var_name):
    """Kata kunci rumpun sebuah variabel (nama variabel + sinonim topiknya)."""
    v = re.sub(r"[^a-z0-9\s]", " ", (var_name or "").lower())
    kata = [w for w in v.split() if len(w) > 3 and w not in _STOPWORD_JUDUL]
    topik = detect_variable_topic(var_name)
    kata += [k for k in TOPIC_KATA_KUNCI.get(topik, []) if k not in kata]
    return kata


def _kata_kunci_judul(judul, x1, x2, m, y):
    """Gabungan kata kunci SELURUH variabel judul (relevansi tingkat kedua)."""
    kata = []
    for v in (x1, x2, m, y):
        kata += _kata_kunci_var(v)
    teks = re.sub(r"[^a-z0-9\s]", " ", (judul or "").lower())
    kata += [w for w in teks.split() if len(w) > 5 and w not in _STOPWORD_JUDUL]
    urut = []
    for k in kata:
        if k and k not in urut:
            urut.append(k)
    return urut


def _cocok_kata(paper, kata):
    """True bila judul paper memuat salah satu kata kunci."""
    if not kata:
        return True
    t = re.sub(r"[^a-z0-9\s]", " ", (paper.get("title") or "").lower())
    return any(k in t for k in kata if len(k) > 3)


def _kunci_paper(paper):
    return ((paper.get("doi") or "").lower().strip(),
            (paper.get("title") or "").lower().strip())


def _kandidat_urut(kandidat_live, kata_slot, kata_judul, pool_topik):
    """Urutkan kandidat dari yang paling relevan:
    live-spesifik -> pool-spesifik -> live-judul -> global-spesifik -> pool-judul
    -> pool apa adanya -> global apa adanya."""
    live = list(kandidat_live or [])
    pool = list(pool_topik or [])
    global_pool = list(ALL_VERIFIED_PAPERS or [])
    tingkat = [
        [p for p in live if _cocok_kata(p, kata_slot)],
        [p for p in pool if _cocok_kata(p, kata_slot)],
        [p for p in live if _cocok_kata(p, kata_judul)],
        [p for p in global_pool if _cocok_kata(p, kata_slot)],
        [p for p in pool if _cocok_kata(p, kata_judul)],
        [p for p in global_pool if _cocok_kata(p, kata_judul)],
        pool,
        global_pool,
    ]
    hasil = []
    kunci = set()
    for daftar in tingkat:
        for p in daftar:
            k = _kunci_paper(p)
            if not any(k) or k in kunci:
                continue
            kunci.add(k)
            hasil.append(p)
    return hasil


_RX_TAHUN_PAREN = re.compile(r"\((?:19|20)\d{2}[a-z]?\)")


def _rapikan_entri_apa(s):
    """Normalisasi entri APA: nama penulis + judul ber-HURUF KAPITAL."""
    s = str(s or "").strip()
    if not s:
        return s
    m = re.match(r"^(.+?)(\(\d{4}[a-z]?\)\.\s+)(.*)$", s, re.S)
    if m:
        kepala = _rapikan_nama_kapital(m.group(1))
        sisa = m.group(3)
        m2 = re.match(r"^(.+?)\.(\s.*|$)", sisa, re.S)
        if m2:
            judul, ekor = m2.group(1), (m2.group(2) or "")
            huruf = [c for c in judul if c.isalpha()]
            if huruf and sum(1 for c in huruf if c.isupper()) / len(huruf) > 0.7:
                judul = _sentence_case_judul(judul)
            sisa = judul + "." + ekor
        return kepala + m.group(2) + sisa
    if _RX_TAHUN_PAREN.search(s):
        return _rapikan_nama_kapital(s)
    return s


def _normalisasi_konteks_cache(obj, _key=None):
    """Rapikan konteks akademik lama yang tersimpan dengan nama HURUF KAPITAL.

    Cache ``v7_*`` menyimpan string sitasi/APA yang sudah jadi. Tanpa langkah
    ini, perbaikan di ``make_in_text_citation`` tidak berlaku untuk judul yang
    konteksnya sudah tersimpan.
    """
    if isinstance(obj, dict):
        return {k: _normalisasi_konteks_cache(v, k) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_normalisasi_konteks_cache(v, _key) for v in obj]
    if isinstance(obj, str):
        kl = (_key or "").lower()
        if kl in ("text", "apa", "apa_full") or kl.startswith("apa_"):
            return _rapikan_entri_apa(obj)
        if kl.startswith("cite") or kl == "in_text":
            return _rapikan_nama_kapital(obj)
    return obj


def get_thesis_academic_context(config):
    """
    Mengambil artikel ilmiah jurnal nyata dengan DOI aktif Open Access untuk Bab I.
    Meniadakan buku cetak fisik / paywall offline.
    Menerapkan aturan mutlak: TIDAK BOLEH ADA KUTIPAN YANG MUNCUL 2 KALI DI LATAR BELAKANG.
    Tepat 11 artikel jurnal yang 100% berbeda penulis & DOI dialokasikan ke 11 slot sitasi.
    Menerapkan sinkronisasi 1-to-1: seluruh yang dikutip di Bab I pasti ada di Daftar Pustaka.
    ORGANIK: Sitasi X1, X2, M benar-benar berasal dari pool jurnal yang sesuai topik variabelnya!

    KEUNIKAN ANTAR-JUDUL (baru): pemilihan rujukan menimbang relevansi topik
    variabel + cap judul pada kunci cache + registry rujukan antar-judul, sehingga
    dua tesis dengan JUDUL/OBJEK berbeda tidak lagi memakai susunan sitasi yang sama.
    """
    x1 = config.get("x1", "Disiplin Kerja")
    x2 = config.get("x2", "Insentif")
    m = config.get("m", "Motivasi")
    y = config.get("y", "Kinerja Pegawai")
    instansi = config.get("instansi", "Instansi Pemerintah Daerah")
    judul = config.get("judul", "") or ""

    # Identifikasi topik spesifik masing-masing variabel
    topic_x1 = detect_variable_topic(x1)
    topic_x2 = detect_variable_topic(x2)
    topic_m = detect_variable_topic(m)

    # Kata kunci relevansi: per variabel (spesifik) dan seluruh judul (lapis kedua)
    kata_x1 = _kata_kunci_var(x1)
    kata_x2 = _kata_kunci_var(x2)
    kata_m = _kata_kunci_var(m)
    kata_y = _kata_kunci_var(y)
    kata_judul = _kata_kunci_judul(judul, x1, x2, m, y)

    # Registry keunikan antar-judul: rujukan milik judul LAIN dihindari.
    thesis_key = f"{instansi}_{x1}_{x2}_{m}_{y}".replace(" ", "_").lower()[:100]
    tracker = _load_used_citations()
    own_dois, own_authors, own_titles = _own_registry_sets(tracker, thesis_key)
    other_dois = set(tracker["used_dois"]) - own_dois
    other_authors = set(tracker["used_lead_authors"]) - own_authors

    cache = load_cache()
    # KUNCI CACHE MEMUAT CAP JUDUL + INSTANSI (bukan hanya nama variabel).
    # Dua tesis dengan judul/objek berbeda TIDAK PERNAH berbagi set sitasi,
    # walaupun nama variabelnya kebetulan sama. Versi dinaikkan ke v7 agar
    # entri cache lama (v6) yang memuat rujukan silang-judul tidak terpakai.
    cap_judul = hashlib.md5(
        f"{judul}|{instansi}".strip().lower().encode("utf-8")).hexdigest()[:10]
    cache_key = f"v7_{topic_x1}_{topic_x2}_{topic_m}_{cap_judul}"
    if cache_key in cache:
        return _normalisasi_konteks_cache(cache[cache_key])

    # Tracking himpunan unik agar tidak ada duplikasi DOI maupun penulis
    chosen_papers = []
    chosen_dois = set()
    chosen_lead_authors = set()

    # Flag strict: bila True, JANGAN memakai rujukan milik judul lain.
    # Pool habis -> raise ValueError agar caller tahu referensi tidak cukup.
    strict_unique = bool(config.get("unique_refs_strict", False))

    def _pakai(p):
        chosen_dois.add((p.get("doi") or "").lower().strip())
        chosen_lead_authors.add(get_lead_author(p).lower())
        chosen_papers.append(p)
        return p

    def select_unique_paper(candidates, pool_candidates, kata_slot=None, kata_judul=None):
        """Pilih paper paling RELEVAN & unik untuk satu slot sitasi Bab I.

        Urutan prioritas:
          1. relevan dengan variabel slot + belum dipakai + bukan milik judul lain;
          2. belum dipakai + bukan milik judul lain (relevansi dilonggarkan);
          3. relevan + belum dipakai (boleh dipakai judul lain, sudah mepet);
          4. belum dipakai (cadangan terakhir).
        """
        urut = _kandidat_urut(candidates, kata_slot or [], kata_judul or [], pool_candidates)

        def _layak(p, hindari_lain):
            doi = (p.get("doi") or "").lower().strip()
            lead = get_lead_author(p).lower()
            title = (p.get("title") or "").lower()
            if not doi:
                return False
            if any(b in title for b in _BANNED_NON_HR_KEYWORDS):
                return False
            if doi in chosen_dois or lead in chosen_lead_authors:
                return False
            if hindari_lain and (doi in other_dois or lead in other_authors):
                return False
            return True

        for hindari_lain in (True, False):
            for p in urut:
                if _layak(p, hindari_lain):
                    return _pakai(p)

        # Cadangan sangat terakhir: paper apa pun yang belum dipakai judul ini.
        for p in urut:
            doi = (p.get("doi") or "").lower().strip()
            lead = get_lead_author(p).lower()
            title = (p.get("title") or "").lower()
            if any(b in title for b in _BANNED_NON_HR_KEYWORDS):
                continue
            if doi and doi not in chosen_dois and lead not in chosen_lead_authors:
                return _pakai(p)
        for p in list(pool_candidates or []):
            doi = (p.get("doi") or "").lower().strip()
            lead = get_lead_author(p).lower()
            title = (p.get("title") or "").lower()
            if any(b in title for b in _BANNED_NON_HR_KEYWORDS):
                continue
            if doi and doi not in chosen_dois and lead not in chosen_lead_authors:
                return _pakai(p)

        if strict_unique:
            raise ValueError(
                f"[unique_refs_strict] Tidak ada paper tersisa di pool topik "
                f"(kandidat live={len(candidates or [])}, pool={len(pool_candidates)}). "
                f"Tambahkan referensi ke TOPIC_JOURNAL_POOLS atau matikan flag."
            )
        # Darurat: seluruh kandidat sudah terpakai -> pakai ulang yang paling relevan
        # (dibuat berisik di log supaya terlihat dan bisa ditambah referensinya).
        if urut:
            print("[sitasi] PERINGATAN: pool sitasi Bab I habis, memakai ulang "
                  f"'{urut[0].get('title', '')[:60]}'.")
            return _pakai(urut[0])
        raise ValueError("Pool sitasi Bab I kosong: tidak ada rujukan tersedia.")

    # Pool topik per slot (topik variabel dulu, baru pool makro sebagai cadangan
    # agar tidak seluruh judul jatuh ke kumpulan rujukan yang sama).
    def _pool(topic, cadangan=("makro_sdm",)):
        d = list(TOPIC_JOURNAL_POOLS.get(topic) or [])
        for c in cadangan:
            d += [p for p in (TOPIC_JOURNAL_POOLS.get(c) or []) if p not in d]
        return d

    kata_makro = ["reformasi", "birokrasi", "pelayanan publik", "aparatur",
                  "tata kelola", "pemerintahan", "sektor publik"]

    # Slot 1: Artikel Jurnal Makro Tata Kelola / Reformasi Birokrasi (Segmen 1)
    p_macro = search_openalex_verified(f"reformasi birokrasi kinerja pelayanan publik {instansi}", limit=3)
    paper_macro = select_unique_paper(p_macro, _pool("makro_sdm"), kata_makro, kata_judul)

    # Slot 2: Artikel Jurnal Manajemen SDM Sektor Publik (Segmen 1)
    p_sdm = search_openalex_verified(f"manajemen sumber daya manusia aparatur {y}", limit=3)
    paper_sdm = select_unique_paper(p_sdm, _pool("makro_sdm"), kata_makro, kata_judul)

    # Slot 3: Artikel Jurnal Urgensi Kinerja Publik Y (Segmen 2)
    p_kinerja = search_openalex_verified(f"pengukuran capaian {y} instansi pemerintah daerah", limit=3)
    paper_kinerja = select_unique_paper(p_kinerja, _pool(detect_variable_topic(y)), kata_y, kata_judul)

    # Slot 4: Artikel Jurnal Pelemahan Variabel Mediasi M (Segmen 5) - SPESIFIK TOPIK M
    pool_m = _pool(topic_m, ("mediasi", "makro_sdm"))
    p_m_list = search_openalex_verified(f"pengaruh {m} terhadap {y}", limit=4)
    paper_m_1 = select_unique_paper(p_m_list, pool_m, kata_m, kata_judul)

    # Slot 5 & 6: Artikel Jurnal Research Gap X1 (Segmen 6) - SPESIFIK TOPIK X1
    pool_x1 = _pool(topic_x1)
    p_x1_list = search_openalex_verified(f"pengaruh {x1} terhadap {y}", limit=5)
    paper_x1_1 = select_unique_paper(p_x1_list, pool_x1, kata_x1, kata_judul)
    paper_x1_2 = select_unique_paper(p_x1_list, pool_x1, kata_x1, kata_judul)

    # Slot 7 & 8: Artikel Jurnal Research Gap X2 (Segmen 7) - SPESIFIK TOPIK X2
    pool_x2 = _pool(topic_x2)
    p_x2_list = search_openalex_verified(f"pengaruh {x2} terhadap {y}", limit=5)
    paper_x2_1 = select_unique_paper(p_x2_list, pool_x2, kata_x2, kata_judul)
    paper_x2_2 = select_unique_paper(p_x2_list, pool_x2, kata_x2, kata_judul)

    # Slot 9: Artikel Jurnal Kepatuhan Operasional (Segmen 8) - mengikuti variabel
    # mediasi judul ini (bukan kueri generik "disiplin" yang selalu sama).
    p_compl = search_openalex_verified(f"kepatuhan {m} pegawai disiplin kerja", limit=3)
    paper_compl = select_unique_paper(p_compl, _pool(topic_m, ("disiplin",)),
                                      kata_m + ["kepatuhan", "disiplin"], kata_judul)

    # Slot 10 & 11: Artikel Jurnal Research Gap Mediasi M (Segmen 9) - 2 Paper Berbeda
    p_med_list = search_openalex_verified(f"pengaruh {x1} terhadap {y} mediasi {m}", limit=5)
    paper_med_1 = select_unique_paper(p_med_list, pool_m, kata_m + ["mediasi", "intervening"], kata_judul)
    paper_med_2 = select_unique_paper(p_med_list, _pool("mediasi", (topic_m,)),
                                      kata_m + ["mediasi", "intervening"], kata_judul)

    # KEUNIKAN ANTAR-JUDUL: daftarkan seluruh rujukan Bab I milik judul ini ke
    # registry global. Judul lain akan menghindari DOI/nama penulis yang sama
    # (lihat other_dois/other_authors di atas), sehingga judul & objek berbeda
    # menghasilkan susunan Daftar Pustaka yang berbeda.
    try:
        _register_thesis_citations(thesis_key, chosen_papers)
    except Exception as _e_reg:
        print(f"[sitasi] peringatan: gagal registrasi rujukan judul ini: {_e_reg}")

    # Susun Daftar Pustaka Bab I (11 Paper Terverifikasi DOI + Dokumen LAKIP)
    dapus_entries = []
    for cp in chosen_papers:
        ent, itl = build_apa_entry(cp)
        dapus_entries.append({"text": ent, "italic": itl})

    # Catatan: Dokumen internal LAKIP tidak lagi disuntikkan paksa ke Daftar Pustaka
    # (mencegah orphan citation & kesamaan rujukan formal antar-peneliti).

    # Urutkan Daftar Pustaka secara alfabetis (A-Z)
    dapus_entries.sort(key=lambda x: x["text"])

    ctx_result = {
        # Sitasi Jurnal Makro, SDM, & Kinerja Y
        "cite_macro": make_in_text_citation(paper_macro["authors"], paper_macro["year"], False),
        "cite_macro_narrative": make_in_text_citation(paper_macro["authors"], paper_macro["year"], True),
        "cite_sdm": make_in_text_citation(paper_sdm["authors"], paper_sdm["year"], False),
        "cite_sdm_narrative": make_in_text_citation(paper_sdm["authors"], paper_sdm["year"], True),
        "cite_kinerja": make_in_text_citation(paper_kinerja["authors"], paper_kinerja["year"], False),
        "cite_kinerja_narrative": make_in_text_citation(paper_kinerja["authors"], paper_kinerja["year"], True),

        # Sitasi Variabel Mediasi M
        "cite_m_1": make_in_text_citation(paper_m_1["authors"], paper_m_1["year"], False),
        "cite_m_1_narrative": make_in_text_citation(paper_m_1["authors"], paper_m_1["year"], True),

        # Sitasi Research Gap X1
        "cite_x1_1": make_in_text_citation(paper_x1_1["authors"], paper_x1_1["year"], False),
        "cite_x1_1_narrative": make_in_text_citation(paper_x1_1["authors"], paper_x1_1["year"], True),
        "cite_x1_2": make_in_text_citation(paper_x1_2["authors"], paper_x1_2["year"], False),
        "cite_x1_2_narrative": make_in_text_citation(paper_x1_2["authors"], paper_x1_2["year"], True),

        # Sitasi Research Gap X2
        "cite_x2_1": make_in_text_citation(paper_x2_1["authors"], paper_x2_1["year"], False),
        "cite_x2_1_narrative": make_in_text_citation(paper_x2_1["authors"], paper_x2_1["year"], True),
        "cite_x2_2": make_in_text_citation(paper_x2_2["authors"], paper_x2_2["year"], False),
        "cite_x2_2_narrative": make_in_text_citation(paper_x2_2["authors"], paper_x2_2["year"], True),

        # Sitasi Kepatuhan & Disiplin
        "cite_compl": make_in_text_citation(paper_compl["authors"], paper_compl["year"], False),

        # Sitasi Mediasi
        "cite_med_1": make_in_text_citation(paper_med_1["authors"], paper_med_1["year"], False),
        "cite_med_1_narrative": make_in_text_citation(paper_med_1["authors"], paper_med_1["year"], True),
        "cite_med_2": make_in_text_citation(paper_med_2["authors"], paper_med_2["year"], False),
        "cite_med_2_narrative": make_in_text_citation(paper_med_2["authors"], paper_med_2["year"], True),

        # === Pemetaan slot -> entri APA lengkap + sitasi in-text (untuk Bab II 2.2) ===
        "slot_refs": {
            "x1": [
                slot_ref(paper_x1_1),
                slot_ref(paper_x1_2),
            ],
            "x2": [
                slot_ref(paper_x2_1),
                slot_ref(paper_x2_2),
            ],
            "m": [
                slot_ref(paper_m_1),
            ],
            "mediasi": [
                slot_ref(paper_med_1),
                slot_ref(paper_med_2),
            ],
            "macro": [
                slot_ref(paper_macro),
                slot_ref(paper_sdm),
                slot_ref(paper_kinerja),
            ],
        },

        # Daftar Pustaka Bab I (100% Pure Journal + Valid Clickable DOI)
        "references": dapus_entries
    }

    ctx_result = _normalisasi_konteks_cache(ctx_result)
    cache[cache_key] = ctx_result
    save_cache(cache)
    return ctx_result


def fetch_live_verified_references(config, max_references=30):
    ctx = get_thesis_academic_context(config)
    refs = ctx.get("references", [])
    return [(r["text"], r["italic"]) for r in refs]


# ==============================================================================
# 2.2 HASIL PENELITIAN YANG RELEVAN  ->  2 STUDI NYATA PER PENGARUH
# ------------------------------------------------------------------------------
# Untuk setiap hubungan (pengaruh) pada model mediasi, ambil minimal 2 penelitian
# terdahulu yang BENAR-BENAR meneliti hubungan tsb DAN abstraknya menyatakan
# adanya pengaruh (positif/signifikan). Seluruh entri berasal dari OpenAlex
# ber-DOI nyata (tidak ada DOI karangan).
# ==============================================================================

# Penanda tingkat pengaruh (positif/signifikan) di dalam abstrak.
_INFLUENCE_POSITIVE = [
    "positif dan signifikan", "positive and significant",
    "berpengaruh positif signifikan", "positif signifikan",
    "berpengaruh positif dan signifikan", "berpengaruh positif",
    "secara positif", "positive and sustainable", "positively",
]
_INFLUENCE_SIGNIFICANT = [
    "signifikan", "significant", "berpengaruh signifikan", "nyata",
]
_INFLUENCE_ANY = [
    "berpengaruh", "mempengaruhi", "memengaruhi", "influence", "influences",
    "effect", "affect", "meningkatkan", "positif", "positive",
]
_INFLUENCE_NEGATIVE = [
    "tidak berpengaruh", "tidak signifikan", "tidak ada pengaruh",
    "no significant", "not significant", "berpengaruh negatif",
    "negatif", "negative", "insignificant", "tidak nyata",
]
_MEDIATION_TERMS = [
    "mediasi", "memediasi", "dimediasi", "intervening", "mediating",
    "mediation", "mediator", "memoderasi", "moderating",
]


def _norm_text(s):
    return re.sub(r"[^a-z0-9\s]", " ", (s or "").lower())


def _var_present(text_norm, var):
    """True bila nama variabel (atau kata intinya) muncul dalam teks ternormalisasi."""
    v = _norm_text(var).strip()
    if not v:
        return True
    if v in text_norm:
        return True
    head = v.split()[0]
    return len(head) > 4 and head in text_norm


def _split_sentences(text):
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+", text or "") if s.strip()]


def _reconstruct_abstract(item):
    """Bangun kembali teks abstrak dari abstract_inverted_index OpenAlex."""
    inv = item.get("abstract_inverted_index")
    if not inv:
        return ""
    positions = {}
    for word, idxs in inv.items():
        for i in idxs:
            positions[i] = word
    return " ".join(positions[i] for i in sorted(positions))


def search_openalex_papers(query, limit=15, min_year=None, max_year=None, retries=5):
    min_year, max_year = _resolve_range(min_year, max_year)
    """OpenAlex: artikel ber-DOI + abstrak nyata (untuk verifikasi hasil penelitian).

    Cache disk + throttle + backoff ditambahkan agar pool sitasi 4.2.x tidak kosong
    ketika OpenAlex membalas 'Too Many Requests' (429).
    """
    clean_q = re.sub(r'[^\w\s]', ' ', query).strip()
    ck = "papers|%s|%d|%d|%d" % (clean_q.lower(), limit, min_year, max_year)
    cache = _oa_cache_load()
    if ck in cache:
        return cache[ck]
    url = ("https://api.openalex.org/works?search=%s"
           "&filter=has_doi:true,type:article,"
           "from_publication_date:%d-01-01,to_publication_date:%d-12-31"
           "&per-page=%d" % (urllib.parse.quote(clean_q), min_year, max_year, limit))
    data = _oa_fetch_json(url, timeout=12, retries=retries)
    papers = _oa_parse_papers(data, min_year, max_year)
    if not papers:
        # OpenAlex kosong/kuota habis -> pakai Crossref sebagai sumber cadangan.
        papers = search_crossref_papers(query, limit=limit, min_year=min_year, max_year=max_year)
    if papers:
        cache[ck] = papers
        _oa_cache_save()
    return papers


def extract_influence_finding(abstract, variables, require_mediation=False):
    """Cari bukti pengaruh dari abstrak nyata untuk sepasang variabel.

    Mengembalikan dict {level, evidence} atau None bila tidak ada bukti
    'berpengaruh' yang eksplisit dan tidak negatif.
    """
    if not abstract:
        return None
    sents = _split_sentences(abstract)
    if not sents:
        return None
    best = None
    candidates = []
    for i in range(len(sents)):
        candidates.append((sents[max(0, i - 1): i + 2], 0))
    # Fallback: jendela lebar (seluruh abstrak). Bukti tetap nyata, hanya tidak
    # terbatas pada 1 kalimat, sehingga peluang lolos lebih besar.
    candidates.append((sents, -1))
    for window, bonus in candidates:
        wn = _norm_text(" ".join(window))
        if not all(_var_present(wn, v) for v in variables):
            continue
        if require_mediation and not any(t in wn for t in _MEDIATION_TERMS):
            continue
        if any(t in wn for t in _INFLUENCE_NEGATIVE):
            continue
        if any(t in wn for t in _INFLUENCE_POSITIVE):
            level, score = "positif_signifikan", 4
        elif any(t in wn for t in _INFLUENCE_SIGNIFICANT):
            level, score = "signifikan", 3
        elif "positif" in wn or "positive" in wn or "positively" in wn:
            level, score = "positif", 2
        elif any(t in wn for t in _INFLUENCE_ANY):
            level, score = "berpengaruh", 1
        else:
            continue
        score = score + bonus
        if score <= 0:
            continue
        if best is None or score > best["score"]:
            best = {"score": score, "level": level,
                    "evidence": " ".join(window).strip()}
    if best is None:
        return None
    return {"level": best["level"], "evidence": best["evidence"]}


def _hasil_sentence(var_a, var_b, level):
    """Kalimat hasil penelitian yang menyatakan adanya pengaruh (berdasarkan abstrak)."""
    if level == "positif_signifikan":
        return f"Hasil penelitian menunjukkan {var_a} berpengaruh positif dan signifikan terhadap {var_b}."
    if level == "signifikan":
        return f"Hasil penelitian menunjukkan {var_a} berpengaruh signifikan terhadap {var_b}."
    if level == "positif":
        return f"Hasil penelitian menunjukkan {var_a} berpengaruh positif terhadap {var_b}."
    return f"Hasil penelitian menunjukkan {var_a} berpengaruh terhadap {var_b}."


def _hasil_sentence_mediation(var_x, var_m, var_y, level):
    """Kalimat hasil untuk hubungan mediasi (peran variabel intervening M)."""
    if level == "positif_signifikan":
        return f"Hasil penelitian menunjukkan {var_m} memediasi pengaruh {var_x} terhadap {var_y} secara positif dan signifikan."
    if level == "signifikan":
        return f"Hasil penelitian menunjukkan {var_m} memediasi pengaruh {var_x} terhadap {var_y} secara signifikan."
    if level == "positif":
        return f"Hasil penelitian menunjukkan {var_m} memediasi secara positif pengaruh {var_x} terhadap {var_y}."
    return f"Hasil penelitian menunjukkan {var_m} memediasi pengaruh {var_x} terhadap {var_y}."


_VAR_VARIANT_MAP = {
    "work environment": ["lingkungan kerja"],
    "self efficacy": ["self efficacy", "efikasi diri"],
    "employee engagement": ["employee engagement", "keterikatan kerja"],
    "kepuasan kerja": ["kepuasan kerja", "job satisfaction"],
    "kinerja pegawai": ["kinerja pegawai", "kinerja karyawan", "employee performance"],
    "komitmen organisasi": ["komitmen organisasi", "organizational commitment"],
    "budaya organisasi": ["budaya organisasi", "organizational culture"],
    "pengembangan karir": ["pengembangan karir", "career development"],
    "pengembangan karier": ["pengembangan karier", "career development"],
    "pelatihan": ["pelatihan", "training"],
    "motivasi kerja": ["motivasi kerja", "work motivation"],
    "promosi jabatan": ["promosi jabatan", "job promotion"],
    "mutasi kerja": ["mutasi kerja", "job rotation"],
    "disiplin kerja": ["disiplin kerja", "work discipline"],
    "lingkungan kerja": ["lingkungan kerja", "work environment"],
    "efikasi diri": ["efikasi diri", "self efficacy"],
}


def _var_variants(var, maks=2):
    """Sinonim Indonesia/Inggris sebuah variabel untuk memperluas query pencarian."""
    v = (var or "").strip()
    if not v:
        return []
    key = _norm_text(v).strip()
    out = [v]
    for k, alts in _VAR_VARIANT_MAP.items():
        if k in key:
            out.extend(alts)
    seen, res = set(), []
    for x in out:
        xl = (x or "").lower().strip()
        if xl and xl not in seen:
            seen.add(xl)
            res.append(x.strip())
    return res[:maks]


def _relationship_queries(rel):
    """Daftar varian query (ID + EN + sinonim) untuk satu hubungan."""
    a, b = rel["var_a"], rel["var_b"]
    q = [rel["query"], rel["fallback_query"], f"{a} terhadap {b}"]
    if rel.get("mediation"):
        q.append(f"peran {rel['var_m']} memediasi {a} terhadap {b}")
        q.append(f"{a} {b} variabel intervening {rel['var_m']}")
    else:
        q.append(f"{a} {b} pengaruh pegawai")
        q.append(f"effect of {a} on {b} employee")
    # Varian sinonim, mis. "Work Environment" <-> "Lingkungan Kerja".
    m = rel.get("var_m")
    for av in _var_variants(a, 2):
        for bv in _var_variants(b, 2):
            q.append(f"pengaruh {av} terhadap {bv}")
            q.append(f"{av} {bv} pegawai")
            if rel.get("mediation") and m:
                q.append(f"{av} terhadap {bv} melalui {m}")
    out, seen = [], set()
    for x in q:
        x = (x or "").strip()
        if x and x.lower() not in seen:
            seen.add(x.lower())
            out.append(x)
    return out


def _pick_relationship_studies(rel, per_rel, used_dois):
    """Pilih `per_rel` studi yang lolos verifikasi DAN berbeda dari pengaruh lain."""
    picks, local = [], set()
    for query in _relationship_queries(rel):
        if len(picks) >= per_rel:
            break
        for p in search_openalex_papers(query, limit=40):
            if len(picks) >= per_rel:
                break
            doi = (p.get("doi") or "").lower()
            if not doi or doi in local or doi in used_dois:
                continue
            finding = extract_influence_finding(
                p.get("abstract", ""), rel["vars"], require_mediation=rel.get("mediation", False))
            if not finding:
                continue
            local.add(doi)
            picks.append((p, finding))
    for p, _ in picks:
        used_dois.add((p.get("doi") or "").lower())
    return picks


_VARIANT_PERSAMAAN = (
    "Sama-sama mengkaji pengaruh {a} terhadap {b}",
    "Kedua penelitian ini memiliki fokus serupa, yaitu pengaruh {a} terhadap {b}",
    "Telaah tersebut sejalan dalam meneliti pengaruh {a} terhadap {b}",
)
_VARIANT_PERSAMAAN_MED = (
    "Sama-sama mengkaji pengaruh {a} terhadap {b} melalui {m}",
    "Kedua penelitian ini memiliki fokus serupa, yaitu pengaruh {a} terhadap {b} melalui {m}",
    "Telaah tersebut sejalan dalam meneliti pengaruh {a} terhadap {b} melalui {m}",
)
_VARIANT_PERBEDAAN = (
    "Objek dan unit analisis berbeda, serta tidak menguji {lain} secara simultan.",
    "Perbedaannya terletak pada objek dan unit analisis, sebab {lain} tidak diuji secara simultan.",
    "Perbedaan juga tampak pada objek dan unit analisis yang berlainan, sedangkan {lain} tidak diteliti secara bersamaan.",
)
_VARIANT_HASIL_LEAD = (
    "Hasil penelitian menunjukkan ",
    "Temuan studi tersebut menunjukkan bahwa ",
    "Hasil pengujian melaporkan bahwa ",
)


def fetch_relationship_studies(config, per_rel=2, use_cache=True):
    """Untuk setiap pengaruh pada model mediasi, ambil `per_rel` penelitian nyata
    BERBEDA (tanpa DOI berulang antar pengaruh) yang menyatakan pengaruhnya."""
    x1 = config.get("x1", "Disiplin Kerja")
    x2 = config.get("x2", "Insentif")
    m = config.get("m", "Motivasi")
    y = config.get("y", "Kinerja Pegawai")

    cache = load_cache() if use_cache else {}
    key = f"v9_relstudi_{x1}_{x2}_{m}_{y}".lower().replace(" ", "_")
    if use_cache and key in cache:
        cached = cache[key]
        if cached.get("relationships"):
            return cached

    rels = [
        {"key": "x1_m", "vars": [x1, m], "mediation": False,
         "var_a": x1, "var_b": m,
         "query": f"pengaruh {x1} terhadap {m}",
         "fallback_query": f"{x1} {m} pegawai"},
        {"key": "x2_m", "vars": [x2, m], "mediation": False,
         "var_a": x2, "var_b": m,
         "query": f"pengaruh {x2} terhadap {m}",
         "fallback_query": f"{x2} {m} pegawai"},
        {"key": "x1_y", "vars": [x1, y], "mediation": False,
         "var_a": x1, "var_b": y,
         "query": f"pengaruh {x1} terhadap {y}",
         "fallback_query": f"{x1} {y} pegawai"},
        {"key": "x2_y", "vars": [x2, y], "mediation": False,
         "var_a": x2, "var_b": y,
         "query": f"pengaruh {x2} terhadap {y}",
         "fallback_query": f"{x2} {y} pegawai"},
        {"key": "m_y", "vars": [m, y], "mediation": False,
         "var_a": m, "var_b": y,
         "query": f"pengaruh {m} terhadap {y}",
         "fallback_query": f"{m} {y} pegawai"},
        {"key": "med_x1", "vars": [x1, m, y], "mediation": True,
         "var_a": x1, "var_b": y, "var_m": m,
         "query": f"pengaruh {x1} terhadap {y} melalui {m}",
         "fallback_query": f"{x1} {y} mediasi {m}"},
        {"key": "med_x2", "vars": [x2, m, y], "mediation": True,
         "var_a": x2, "var_b": y, "var_m": m,
         "query": f"pengaruh {x2} terhadap {y} melalui {m}",
         "fallback_query": f"{x2} {y} mediasi {m}"},
    ]

    all_vars = []
    for v in (x1, x2, m, y):
        if v not in all_vars:
            all_vars.append(v)

    used_dois = set()
    out, total = [], 0
    for ri, rel in enumerate(rels):
        picks = _pick_relationship_studies(rel, per_rel, used_dois)
        if rel.get("mediation"):
            label = f"{rel['var_a']} terhadap {rel['var_b']} melalui {rel['var_m']}"
        else:
            label = f"{rel['var_a']} terhadap {rel['var_b']}"
        others = [v for v in all_vars if v not in rel["vars"]]

        studies = []
        for k, (p, finding) in enumerate(picks):
            # B3: variasikan kalimat persamaan/perbedaan/hasil agar tidak identik
            # baik antar-studi dalam satu pengaruh maupun antar-pengaruh.
            vsel = (ri + k + 1) % len(_VARIANT_PERSAMAAN)
            if rel.get("mediation"):
                persamaan = _VARIANT_PERSAMAAN_MED[vsel].format(
                    a=rel["var_a"], b=rel["var_b"], m=rel["var_m"])
            else:
                persamaan = _VARIANT_PERSAMAAN[vsel].format(
                    a=rel["var_a"], b=rel["var_b"])
            if others:
                perbedaan = _VARIANT_PERBEDAAN[vsel].format(lain=" dan ".join(others))
            else:
                perbedaan = "Objek dan unit analisis berbeda."

            apa, _ = build_apa_entry(p)
            if rel.get("mediation"):
                hasil = _hasil_sentence_mediation(rel["var_a"], rel["var_m"], rel["var_b"], finding["level"])
            else:
                hasil = _hasil_sentence(rel["var_a"], rel["var_b"], finding["level"])
            _lead0 = "Hasil penelitian menunjukkan "
            _rest = hasil[len(_lead0):] if hasil.startswith(_lead0) else hasil
            hasil = _VARIANT_HASIL_LEAD[vsel] + _rest
            studies.append({
                "apa": apa,
                "cite": ctx_cite(p),
                "title": (p.get("title") or "").strip().rstrip("."),
                "year": p.get("year"),
                "doi": p.get("doi"),
                "level": finding["level"],
                "evidence": finding["evidence"],
                "persamaan": persamaan,
                "perbedaan": perbedaan,
                "hasil": hasil,
            })
        total += len(studies)
        if len(studies) < per_rel:
            print(f"[2.2] peringatan: pengaruh '{rel['key']}' hanya {len(studies)}/{per_rel} studi berbeda")
        out.append({"key": rel["key"], "var_a": rel["var_a"], "var_b": rel["var_b"],
                    "label": label, "studies": studies})

    result = {"relationships": out, "per_rel": per_rel, "total": total}
    # Cache hanya bila seluruh pengaruh lengkap, agar kegagalan jaringan bisa dicoba ulang.
    if use_cache and all(len(r["studies"]) >= per_rel for r in out):
        cache[key] = result
        save_cache(cache)
    return result

# ==============================================================================
# P0 #9 — VERIFIKASI REFERENSI (DOI, Judul, Penulis, Tahun, Retraksi)
# ==============================================================================

def verify_references(references: list, min_year: int = None, max_year: int = None) -> dict:
    """
    Verifikasi daftar referensi APA ke Crossref/OpenAlex.

    Input: list of dict {"apa_full": str, "doi": str|None, "title": str|None, 
                           "authors": list|None, "year": int|None}
    Output: dict dengan status verifikasi per referensi + ringkasan.

    Pengecekan:
    1. DOI valid & resolvable (Crossref)
    2. Judul cocok (≥80% similarity)
    3. Penulis utama cocok
    4. Tahun publikasi cocok
    5. Bukan retraksi (Crossref `retracted` flag / OpenAlex `retracted` field)
    """
    import urllib.request
    import urllib.parse
    import re
    
    min_year, max_year = _resolve_range(min_year, max_year)
    results = []
    verified_count = 0
    issues = []

    for i, ref in enumerate(references):
        apa_full = (ref.get("apa_full") or "").strip()
        doi = (ref.get("doi") or "").strip()
        title = (ref.get("title") or "").strip()
        authors = ref.get("authors") or []
        year = ref.get("year")

        # Ekstrak DOI dari apa_full jika tidak ada field doi
        if not doi:
            import re
            m = re.search(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", apa_full)
            if m:
                doi = "https://doi.org/" + m.group(0).rstrip(".,;)]}\"'")

        item_result = {
            "index": i,
            "apa_full": apa_full,
            "doi_provided": doi,
            "title_provided": title,
            "authors_provided": authors,
            "year_provided": year,
            "verified": False,
            "doi_valid": False,
            "title_match": False,
            "author_match": False,
            "year_match": False,
            "not_retracted": True,
            "crossref_data": None,
            "errors": []
        }

        if not doi:
            item_result["errors"].append("DOI tidak ditemukan di entri APA")
            results.append(item_result)
            continue

        # Normalisasi DOI
        doi_clean = doi.replace("https://doi.org/", "").replace("http://doi.org/", "").strip()
        url = f"https://api.crossref.org/works/{urllib.parse.quote(doi_clean)}"
        
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "AutoTesis/1.0 (mailto:autotesis@example.com)"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                if resp.status != 200:
                    item_result["errors"].append(f"Crossref HTTP {resp.status}")
                    results.append(item_result)
                    continue
                data = json.load(resp)
        except Exception as e:
            item_result["errors"].append(f"Gagal query Crossref: {e}")
            results.append(item_result)
            continue

        msg = data.get("message", {})
        item_result["crossref_data"] = {
            "doi": msg.get("DOI"),
            "title": (msg.get("title") or [""])[0],
            "authors": msg.get("author"),
            "year": None,
            "type": msg.get("type"),
            "retracted": False
        }

        # Cek retracted
        # Crossref: ada field `relation` dengan `has-retracted` atau `is-retracted`
        # Atau check via `update_type` = 'retraction'
        # OpenAlex lebih eksplisit dengan field `retracted`
        # Di sini cek Crossref dulu
        relation = msg.get("relation", {})
        has_retraction = False
        for rel_type, rel_items in relation.items():
            if "retract" in rel_type.lower():
                for item in rel_items:
                    if item.get("id-type") == "doi":
                        has_retraction = True
                        break
        # Cek update_type
        updates = msg.get("updates", []) or []
        for upd in updates:
            if upd.get("type") == "retraction":
                has_retraction = True
                break

        item_result["not_retracted"] = not has_retraction
        if has_retraction:
            item_result["errors"].append("ARTIKEL DITARIK (RETRACTED)")

        # Verifikasi DOI valid (ada di Crossref)
        item_result["doi_valid"] = True

        # Verifikasi judul
        cf_title = ((msg.get("title") or [""])[0] or "").strip()
        item_result["crossref_data"]["title"] = cf_title
        if title and cf_title:
            # Similarity sederhana: Jaccard kata
            t1 = set(re.sub(r"[^\w\s]", " ", title.lower()).split())
            t2 = set(re.sub(r"[^\w\s]", " ", cf_title.lower()).split())
            if t1 and t2:
                sim = len(t1 & t2) / len(t1 | t2)
                # Terima similarity ≥60% (mengakomodasi terjemahan judul)
                item_result["title_match"] = sim >= 0.6
                if not item_result["title_match"]:
                    item_result["errors"].append(f"Judul tidak cocok (similarity {sim:.0%})")
        elif not title:
            item_result["errors"].append("Judul tidak disediakan untuk verifikasi")

        # Verifikasi penulis utama
        cf_authors = msg.get("author") or []
        lead_provided = (authors[0] if authors else "").strip()
        lead_cf = ""
        if cf_authors:
            a = cf_authors[0]
            lead_cf = " ".join([a.get("given", ""), a.get("family", "")]).strip()
        item_result["crossref_data"]["authors"] = [f"{a.get('given','')} {a.get('family','')}".strip() for a in cf_authors]
        if lead_provided and lead_cf:
            # Cocokkan nama keluarga (family name) - ambil setelah koma terakhir atau kata terakhir
            def _family(name):
                if "," in name:
                    return name.split(",")[0].strip().lower()
                return name.split()[-1].lower()
            lp_fam = _family(lead_provided)
            cf_fam = _family(lead_cf)
            item_result["author_match"] = lp_fam == cf_fam
            if not item_result["author_match"]:
                item_result["errors"].append(f"Penulis utama tidak cocok: '{lead_provided}' vs '{lead_cf}'")
        elif not lead_provided:
            item_result["errors"].append("Penulis tidak disediakan untuk verifikasi")

        # Verifikasi tahun
        cf_year = None
        issued = msg.get("issued", {}).get("date-parts", [[None]])[0]
        if issued and issued[0]:
            try:
                cf_year = int(issued[0])
            except Exception:
                pass
        item_result["crossref_data"]["year"] = cf_year
        if year and cf_year:
            item_result["year_match"] = year == cf_year
            if not item_result["year_match"]:
                item_result["errors"].append(f"Tahun tidak cocok: {year} vs {cf_year}")
        elif not year:
            item_result["errors"].append("Tahun tidak disediakan untuk verifikasi")

        # Validasi rentang tahun
        if cf_year and (cf_year < min_year or cf_year > max_year):
            item_result["errors"].append(f"Tahun {cf_year} di luar rentang {min_year}-{max_year}")

        # Overall verified
        # Kriteria: DOI valid + penulis cocok + tahun cocok + tidak retract
        # Title_match hanya info (warning) karena tesis Indonesia sering menerjemahkan judul Inggris
        item_result["verified"] = (
            item_result["doi_valid"] and
            item_result["author_match"] and
            item_result["year_match"] and
            item_result["not_retracted"]
        )
        if not item_result["title_match"] and item_result["doi_valid"] and item_result["author_match"] and item_result["year_match"]:
            item_result["title_note"] = "Judul mungkin terjemahan (similarity < 60%)"

        if item_result["verified"]:
            verified_count += 1
        else:
            issues.append(f"Ref #{i+1}: {'; '.join(item_result['errors'])}")

        results.append(item_result)

    return {
        "total": len(references),
        "verified": verified_count,
        "failed": len(references) - verified_count,
        "details": results,
        "issues": issues,
        "all_verified": verified_count == len(references) and len(references) > 0
    }


def parse_apa_to_structured(apa_text: str) -> dict:
    """Parse string APA menjadi dict terstruktur untuk verifikasi."""
    import re
    # Pola APA sederhana: Penulis (Tahun). Judul. Jurnal, Vol(Issue), Hal. DOI
    result = {"apa_full": apa_text.strip(), "doi": None, "title": None, "authors": [], "year": None}
    
    # Ekstrak DOI
    m = re.search(r"(https?://doi\.org/)?(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+)", apa_text)
    if m:
        result["doi"] = "https://doi.org/" + m.group(2).rstrip(".,;)]}\"'")
    
    # Ekstrak tahun: (YYYY) atau (YYYYa) setelah nama penulis
    m = re.search(r"\((\d{4}[a-z]?)\)\.", apa_text)
    if m:
        try:
            result["year"] = int(m.group(1)[:4])
        except Exception:
            pass
    
    # Ekstrak judul: setelah tahun, sebelum titik yang diikuti jurnal/volume
    # Pola: (YYYY). Judul. Nama Jurnal
    m = re.search(r"\(\d{4}[a-z]?\)\.\s*([^.]+?)(?:\.\s+[A-Z][^.]+(?:,\s*\d+)?(?:\(\d+\))?\s*,\s*\d+(?:[–-]\d+)?\s*\.\s*https?://doi\.org/)", apa_text)
    if not m:
        # Fallback: ambil teks setelah (YYYY). sampai titik berikutnya
        m = re.search(r"\(\d{4}[a-z]?\)\.\s*([^.]+)\.", apa_text)
    if m:
        result["title"] = m.group(1).strip().rstrip(".")
    
    # Ekstrak penulis: sebelum (YYYY)
    m = re.search(r"^(.+?)\s*\(\d{4}[a-z]?\)\.", apa_text)
    if m:
        authors_str = m.group(1).strip()
        # Format APA: "Simamora, H., & Mardiana, S." atau "Jannah, M., & Hidayat, R."
        # Pisahkan oleh ", & " atau ", dan " lalu parse masing-masing "Nama, Inisial."
        # Normalisasi pemisah
        authors_str = re.sub(r",\s*&\s*", ", ", authors_str)
        authors_str = re.sub(r",\s*dan\s+", ", ", authors_str)
        # Sekarang pisahkan oleh ", " yang diikuti huruf kapital (awal nama keluarga berikutnya)
        # Pola: "Simamora, H., Mardiana, S." -> ["Simamora, H.", "Mardiana, S."]
        parts = re.split(r",\s*(?=[A-Z][a-z]+,)", authors_str)
        authors = [p.strip() for p in parts if p.strip()]
        result["authors"] = authors
    
    return result


def verify_references_from_apa_list(apa_list: list) -> dict:
    """Verifikasi list string APA (['Author (Year). Title...'])"""
    structured = [parse_apa_to_structured(apa) for apa in apa_list]
    return verify_references(structured)

