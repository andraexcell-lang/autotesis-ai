# backend/variable_tracer.py
"""
AGEN PENELUSUR & PENYIMPAN DIMENSI-INDIKATOR VIA LIVE ACADEMIC API (Tahap B).

Alur (sepenuhnya model/agen AI, sesuai kesepakatan):
  1. Cek pustaka permanen  -> sudah ada & lolos? pakai, tanpa biaya API.
  2. Cari di OpenAlex (artikel + buku), rentang 2022-2026, has_doi, Open Access.
  3. Unduh PDF kandidat, ekstrak teks (PyMuPDF).
  4. Gemini mengekstrak dimensi & indikator HANYA dari teks sumber itu,
     disertai kutipan literal.
  5. Verifikasi: SETIAP indikator wajib cocok di teks sumber; yang gagal dibuang.
  6. Lolos -> simpan PERMANEN ke pustaka (variable_library/). Gagal semua ->
     needs_review (tidak pernah mengarang).
  7. Unggahan pengguna juga diekstrak lewat jalur yang sama (trace_from_upload).

Tidak ada knowledge-base template sebagai isi; sumber tunggal per variabel.
"""

import datetime
import json
import os
import re
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed

from llm_ontology_researcher import (
    _call_gemini_json, _call_gemini_matrix, _validate_llm_ontology, last_matrix_model,
    _match_variable_topic, _dimensions_fit_topic,
)
from llm_title_parser import load_config, MODEL_CHAIN, EXTRACT_CHAIN, extract_chain, gemini_keys
import variable_library as VL
import doc_text_extract as DTE

BACKEND_DIR = os.path.dirname(os.path.abspath(__file__))
TMP_DIR = os.path.join(BACKEND_DIR, "_trace_tmp")

TRACE_TIMEOUT_PDF = 30
TRACE_MAX_CANDIDATES = 8
TRACE_MAX_CHARS = 4000       # potongan teks yang dikirim ke Gemini per kandidat (dari 15000 -> hemat 73%)
TRACE_MIN_INDICATORS = 6     # TARGET: 2-3 dimensi x 2 indikator (diupayakan lewat top-up)
TRACE_MIN_DIMENSIONS = 2     # Kebijakan generated_v1: 2-7 dimensi per variabel
# LANTAI KERAS (bukan target): sumber primer yang SECARA EKSPLISIT hanya
# mencantumkan 4 indikator (mis. Dessler 2020 untuk Training: Reaction,
# Learning, Behavior, Results) tetap SAH untuk PLS-SEM (minimum 3 indikator).
# Tanpa lantai ini sistem membuang sumber nyata hanya karena <6, lalu variabel
# berakhir kosong. Indikator TIDAK pernah ditambah tanpa kutipan literal.
TRACE_MIN_INDICATORS_ACCEPT = 4
TRACE_MIN_DIMENSIONS_ACCEPT = 2
TRACE_MAX_DIMENSIONS = 7     # batas FINAL setelah seleksi 2-7 dimensi
TRACE_SOURCE_MAX_DIMENSIONS = 20  # sumber boleh 2-20 dimensi; dipilih 2-7
# LAPIS 2 (fallback): bila semua kandidat gagal ambang Lapis 1, kandidat terbaik
# diterima lalu DICUKUPKAN dari SUMBER YANG SAMA dengan jendela teks diperbesar.
# Semua item tetap wajib kutipan LITERAL (anti-halusinasi; sesuai keputusan user).
TRACE_TOPUP_CHARS = 16000    # jendela teks Lapis 2 (Lapis 1 hanya 4000)
TRACE_TOPUP_TRIES = 3        # maksimum kandidat near-miss yang dicoba di Lapis 2
# Ambang LAPIS 2: setelah dicukupkan dari sumber yang sama, minimal 2 dimensi
# (beberapa konstruk nyata hanya punya 2 dimensi sah, mis. Lingkungan Kerja:
# fisik & non-fisik) dan tetap >=6 indikator. Hasil akhir tetap dirapikan
# (dedup dimensi kembar + batas indikator per dimensi) agar tidak menggelembung.
TRACE_MIN_DIMENSIONS_L2 = 2
TRACE_MAX_IND_PER_DIM = 4    # batas indikator per dimensi (desain: 2-4)
TRACE_DEDUP_JACCARD = 0.6    # ambang tumpang-tindih indikator untuk gabung dimensi kembar
# CATATAN (keputusan user, 15 Okt 2026):
#   Satu variabel = 3-6 dimensi; tiap dimensi = 2-4 indikator.
#   Ambang lama (>=3 dimensi & >=8 indikator) terlalu ketat: sumber NYATA
#   seperti skala NASA-TLX (6 indikator literal) gugur padahal sah.
#   Jumlah indikator FINAL sengaja TIDAK dikunci di sini — akan diatur pada
#   prompt penulisan Bab 2 sesuai kebutuhan tesis.
UA = "AutoTesisVariableTracer/1.0 (mailto:tesismagister@research.ac.id)"

# Cache ekstraksi per (DOI, var_name) -> hasil dimensi/indikator
# DOI dinormalisasi lowercase; var_name lowercase.
EXTRACTION_CACHE_FILE = os.path.join(BACKEND_DIR, "_extraction_cache.json")

# Cache SUMBER DIMENSI per-judul (thesis_key) -> entri dimensi terverifikasi.
# Menjamin regenerasi judul yang sama memakai sumber dimensi yang SAMA
# (deterministik, tanpa unduh/LLM ulang) sehingga tetap berbeda antar-judul.
DIM_SOURCE_CACHE_FILE = os.path.join(BACKEND_DIR, "_dimension_source_cache.json")

# Rate limiting untuk OpenAlex (menghindari HTTP 429 Too Many Requests)
SEARCH_MIN_INTERVAL = 1.5    # detik minimal antar request OpenAlex
SEARCH_MAX_RETRY = 4         # jumlah percobaan ulang saat kena 429
SEARCH_BACKOFF_BASE = 5.0    # detik; backoff eksponensial: 5, 10, 20, 40
_last_search_ts = 0.0        # timestamp request pencarian terakhir (modul-level)

# Lock modul: melindungi file cache saat beberapa variabel riset BERJALAN SERENTAK
# (paralelisasi 4 variabel pada build_full_thesis_ontology).
_CACHE_LOCK = threading.RLock()
_SEARCH_LOCK = threading.Lock()

# Batas waktu fase pencarian (semua sumber paralel). Profil cepat memastikan tiap
# sumber selesai jauh di bawah ini; angka ini hanya jaring pengaman bila ada yang
# tertahan lebih lama dari perkiraan.
SEARCH_WALL_MAX = 25.0
# Cukup kandidat == jumlah yang benar-benar akan dicoba -> tak perlu lagi menunggu
# sumber yang masih tertahan backoff (OpenAlex sering jadi yang paling lambat).
SEARCH_ENOUGH = TRACE_MAX_CANDIDATES
# Wajib menunggu minimal sekian detik sebelum melepas pekerjaan yang tersisa,
# supaya sumber yang SEHAT (biasanya selesai < 8 dtk) tetap sempat menyumbang
# hasilnya; yang bermasalah baru ditinggalkan setelah ambang ini.
SEARCH_MIN_WAIT = 8.0


# ============================================================================
# CACHE EKSTRAKSI PER (DOI, VAR_NAME)
# ============================================================================
def _load_extraction_cache():
    """Load cache ekstraksi dari file JSON."""
    with _CACHE_LOCK:
        if os.path.exists(EXTRACTION_CACHE_FILE):
            try:
                with open(EXTRACTION_CACHE_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (json.JSONDecodeError, OSError):
                return {}
    return {}


def _save_extraction_cache(cache):
    """Simpan cache ekstraksi ke file JSON."""
    with _CACHE_LOCK:
        try:
            tmp = EXTRACTION_CACHE_FILE + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(cache, f, indent=2, ensure_ascii=False)
            os.replace(tmp, EXTRACTION_CACHE_FILE)
        except OSError as e:
            print(f"[cache] Warning: gagal simpan extraction cache: {e}")


def _load_dim_source_cache():
    """Muat cache sumber dimensi per-judul."""
    with _CACHE_LOCK:
        if os.path.exists(DIM_SOURCE_CACHE_FILE):
            try:
                with open(DIM_SOURCE_CACHE_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (json.JSONDecodeError, OSError):
                return {}
    return {}


def _save_dim_source_cache(cache):
    """Simpan cache sumber dimensi per-judul."""
    with _CACHE_LOCK:
        try:
            tmp = DIM_SOURCE_CACHE_FILE + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(cache, f, indent=2, ensure_ascii=False)
            os.replace(tmp, DIM_SOURCE_CACHE_FILE)
        except OSError as e:
            print(f"[cache] Warning: gagal simpan dimension source cache: {e}")


def _cache_key(doi, var_name):
    """Normalisasi kunci cache: DOI lowercase + var_name lowercase."""
    if not doi:
        return None
    return f"{doi.strip().lower()}|{var_name.strip().lower()}"


# ---------------------------------------------------------------------------
# 1. Pencarian OpenAlex 2022-2026
# ---------------------------------------------------------------------------

def _search_openalex(query, kind="article", limit=6, fast=False):
    """
    Cari karya 2022-2026 ber-DOI & Open Access di OpenAlex.
    kind: "article" | "book". Mengembalikan list dict kandidat.

    fast=True dipakai saat pencarian multi-sumber: cukup 2 percobaan dengan
    backoff pendek, karena ada sumber cadangan -- lebih baik gagal cepat
    daripada menahan hasil dari Crossref/DOAJ yang sehat.
    """
    clean_q = re.sub(r"[^\w\s]", " ", query).strip()
    enc = urllib.parse.quote(clean_q)
    year_min, year_max = VL.active_year_range()
    max_retry = 3 if fast else SEARCH_MAX_RETRY
    backoff_base = 3.0 if fast else SEARCH_BACKOFF_BASE
    url = (
        f"https://api.openalex.org/works?search={enc}"
        f"&filter=has_doi:true,is_oa:true,type:{kind},"
        f"from_publication_date:{year_min}-01-01,to_publication_date:{year_max}-12-31"
        f"&per-page={limit}"
        "&select=id,doi,title,publication_year,type,primary_location,open_access,"
        "abstract_inverted_index,authorships,biblio"
        # mailto= sebagai query -> masuk "polite pool" OpenAlex (rate limit longgar)
        "&mailto=tesismagister@research.ac.id"
    )
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    candidates = []
    data = None
    global _last_search_ts
    for attempt in range(max_retry):
        # Jaga jeda minimal antar-request (rate limiting client-side).
        # Lock dipakai supaya panggilan paralel tidak menembak OpenAlex bersamaan.
        with _SEARCH_LOCK:
            wait = SEARCH_MIN_INTERVAL - (time.time() - _last_search_ts)
            if wait > 0:
                time.sleep(wait)
            _last_search_ts = time.time()
        try:
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            _last_search_ts = time.time()
            break
        except urllib.error.HTTPError as e:
            _last_search_ts = time.time()
            if e.code == 429 and attempt < max_retry - 1:
                backoff = backoff_base * (2 ** attempt)
                print(f"[tracer] OpenAlex 429 (rate limit) -> tunggu {backoff:.0f}s, coba lagi ({attempt+1}/{max_retry})")
                time.sleep(backoff)
                continue
            print(f"[tracer] pencarian OpenAlex gagal: {e}")
            break
        except Exception as e:
            print(f"[tracer] pencarian OpenAlex gagal: {e}")
            break
    if not data:
        return []
    try:
        for item in data.get("results", []):
            doi = item.get("doi")
            if not doi or not str(doi).startswith("http"):
                continue
            loc = item.get("primary_location") or {}
            src = (loc.get("source") or {}).get("display_name") or "Jurnal Ilmiah"
            pdf_url = None
            oa = item.get("open_access") or {}
            if isinstance(oa.get("oa_url"), str) and oa.get("oa_url"):
                pdf_url = oa["oa_url"]
            elif isinstance(loc.get("pdf_url"), str) and loc.get("pdf_url"):
                pdf_url = loc["pdf_url"]
            authors = []
            for a in item.get("authorships", []) or []:
                nm = (a.get("author") or {}).get("display_name")
                if nm:
                    authors.append(nm)
            biblio = item.get("biblio") or {}
            candidates.append({
                "doi": doi,
                "title": (item.get("title") or "Tanpa Judul").strip(),
                "year": item.get("publication_year") or 0,
                "journal": src,
                "kind": kind,
                "pdf_url": pdf_url,
                "authors": authors,
                "volume": biblio.get("volume"),
                "issue": biblio.get("issue"),
                "first_page": biblio.get("first_page"),
                "last_page": biblio.get("last_page"),
                "abstract": _abstract_to_text(item.get("abstract_inverted_index")),
            })
    except Exception as e:
        print(f"[tracer] pencarian OpenAlex gagal: {e}")
    return candidates


def _abstract_to_text(inverted):
    if not isinstance(inverted, dict) or not inverted:
        return ""
    try:
        pos = []
        for word, idxs in inverted.items():
            for i in idxs:
                pos.append((i, word))
        pos.sort()
        return " ".join(w for _, w in pos)
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# 1b. Pencarian MULTI-SUMBER (Crossref, Semantic Scholar) + Unpaywall
# ---------------------------------------------------------------------------
# Masalah yang ditemui saat uji langsung: SEMUA kandidat datang dari satu titik
# tunggal (OpenAlex). Begitu OpenAlex memberi 429 dengan Retry-After berjam-jam,
# seluruh telusur mandek dan seluruh variabel masuk needs_review.
#
# Solusinya: tambah sumber lain yang semuanya GRATIS dan TIDAK PERLU API KEY,
# dengan rate limit terpisah -- bila satu sumber diblokir, yang lain tetap jalan.
# Ketiganya dijalankan PARALEL sehingga sekaligus lebih cepat dan lebih banyak
# kandidat jurnal yang terkumpul.
#
# Hasil uji pengukuran (kueri "Employee Engagement dimensi indikator"):
#   OpenAlex saja  -> 4 kandidat, 20,5 dtk
#   Paralel 3 sumber -> 27 kandidat unik (OpenAlex 21 + Crossref 12 + S2 4)
#
# DOAJ sempat dicoba lalu DIHAPUS dari rangkaian: latensi terukur 24-64 dtk per
# kueri (kueri Indonesia "kinerja pegawai" = 63,6 dtk) sehingga selalu melewati
# batas pencarian dan hanya menahan hasil tanpa menambah kandidat berarti.
#
# Catatan penting: Crossref hanyalah registry DOI (jarang punya tautan PDF), jadi
# kandidat tanpa PDF diungkap lewat Unpaywall -> DOI jadi PDF Open Access.

CROSSREF_BASE = "https://api.crossref.org/works"
S2_BASE = "https://api.semanticscholar.org/graph/v1/paper/search"
DOAJ_BASE = "https://doaj.org/api/search/articles"
UNPAYWALL_BASE = "https://api.unpaywall.org/v2"
CONTACT_EMAIL = "tesismagister@research.ac.id"

# Jeda minimal antar-request, terpisah per sumber (kuota tiap API berbeda).
_THROTTLE = {"crossref": 0.6, "s2": 1.2, "unpaywall": 0.4}
_last_ts = {k: 0.0 for k in _THROTTLE}


def _throttle(source):
    """Jaga jeda minimal antar-request untuk satu sumber (aman dipanggil dari thread)."""
    global _last_ts
    wait = _THROTTLE.get(source, 1.0) - (time.time() - _last_ts.get(source, 0.0))
    if wait > 0:
        time.sleep(wait)
    _last_ts[source] = time.time()


def _http_json(url, source, timeout=15, max_retry=3, backoff_base=None):
    """GET JSON dengan throttle + backoff 429/5xx. Mengembalikan dict, atau None bila gagal."""
    backoff_base = SEARCH_BACKOFF_BASE if backoff_base is None else backoff_base
    for attempt in range(max_retry):
        _throttle(source)
        req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code in (429, 502, 503, 504) and attempt < max_retry - 1:
                backoff = backoff_base * (2 ** attempt)
                print(f"[tracer] {source} HTTP {e.code} -> tunggu {backoff:.0f}s ({attempt+1}/{max_retry})")
                time.sleep(backoff)
                continue
            print(f"[tracer] {source} dilewati (HTTP {e.code})")
            return None
        except Exception as e:
            print(f"[tracer] {source} gagal: {e}")
            return None
    return None


def _strip_jats(text):
    """Crossref mengirim abstrak sebagai XML JATS -> bersihkan menjadi teks biasa."""
    if not text:
        return ""
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", str(text))).strip()


def _search_crossref(query, kind="article", limit=4, fast=False):
    """
    Crossref: registry DOI terbesar di dunia (~150 juta karya). Bebas, tanpa API key.
    Kekuatannya di CAKUPAN (nyaris semua jurnal ber-DOI terdaftar di sini).
    """
    clean_q = re.sub(r"[^\w\s]", " ", query).strip()
    type_filter = "journal-article" if kind == "article" else "book"
    y_min, y_max = VL.active_year_range()
    url = (
        f"{CROSSREF_BASE}?query.bibliographic={urllib.parse.quote(clean_q)}"
        f"&filter=from-pub-date:{y_min}-01-01,until-pub-date:{y_max}-12-31,type:{type_filter}"
        f"&rows={limit}&mailto={CONTACT_EMAIL}"
    )
    data = _http_json(url, "crossref",
                      max_retry=2 if fast else 3,
                      backoff_base=2.0 if fast else None)
    if not data:
        return []
    out = []
    for item in (data.get("message") or {}).get("items", []) or []:
        doi = item.get("DOI")
        if not doi:
            continue
        try:
            year = int(((item.get("issued") or {}).get("date-parts") or [[0]])[0][0])
        except Exception:
            year = 0
        _ymin, _ymax = VL.active_year_range()
        if not (_ymin <= year <= _ymax):
            continue
        titles = item.get("title") or []
        authors = []
        for a in item.get("author") or []:
            nm = " ".join(x for x in (a.get("given"), a.get("family")) if x)
            if nm:
                authors.append(nm)
        pdf_url = None
        for lk in item.get("link") or []:
            ct = str(lk.get("content-type") or "")
            if "pdf" in ct.lower() and lk.get("URL"):
                pdf_url = lk["URL"]
                break
        page = str(item.get("page") or "")
        pages = page.split("-") if page else []
        out.append({
            "doi": f"https://doi.org/{doi}",
            "title": (titles[0] if titles else "Tanpa Judul").strip(),
            "year": year,
            "journal": (item.get("container-title") or ["Jurnal Ilmiah"])[0],
            "kind": kind,
            "pdf_url": pdf_url,
            "authors": authors,
            "volume": item.get("volume"),
            "issue": item.get("issue"),
            "first_page": pages[0] if pages else None,
            "last_page": pages[-1] if len(pages) > 1 else None,
            "abstract": _strip_jats(item.get("abstract")),
            "_source": "crossref",
        })
    return out


def _search_s2(query, kind="article", limit=4, fast=False):
    """
    Semantic Scholar: bebas, tanpa API key, dan MENYEDIAKAN TAUTAN PDF Open Access
    langsung (openAccessPdf) -- sangat membantu karena tahap verifikasi butuh PDF.
    """
    clean_q = re.sub(r"[^\w\s]", " ", query).strip()
    y_min, y_max = VL.active_year_range()
    fields = "title,year,abstract,venue,externalIds,openAccessPdf,authors,journal"
    url = (
        f"{S2_BASE}?query={urllib.parse.quote(clean_q)}"
        f"&year={y_min}-{y_max}&limit={limit}&fields={fields}"
    )
    data = _http_json(url, "s2",
                      max_retry=3 if fast else 3,
                      backoff_base=3.0 if fast else None)
    if not data:
        return []
    out = []
    for item in data.get("data", []) or []:
        doi = (item.get("externalIds") or {}).get("DOI")
        if not doi:
            continue
        year = item.get("year") or 0
        _ymin, _ymax = VL.active_year_range()
        if not (_ymin <= year <= _ymax):
            continue
        oa = item.get("openAccessPdf") or {}
        jr = item.get("journal") or {}
        out.append({
            "doi": f"https://doi.org/{doi}",
            "title": (item.get("title") or "Tanpa Judul").strip(),
            "year": year,
            "journal": item.get("venue") or jr.get("name") or "Jurnal Ilmiah",
            "kind": kind,
            "pdf_url": oa.get("url") or None,
            "authors": [a.get("name") for a in item.get("authors") or [] if a.get("name")],
            "volume": jr.get("volume"),
            "issue": jr.get("issue"),
            "first_page": None,
            "last_page": None,
            "abstract": item.get("abstract") or "",
            "_source": "semantic_scholar",
        })
    return out


def _search_doaj(query, kind="article", limit=4, fast=False):
    """
    DOAJ (Directory of Open Access Journals): sumber kandidat keempat.

    Keunggulan DOAJ dibanding Crossref: banyak hasilnya sudah MENYEDIAKAN PDF
    lengkap langsung, sehingga tahap ekstraksi/verifikasi tidak bergantung pada
    Unpaywall. Ini penting karena uji lapangan menunjukkan OpenAlex & Semantic
    Scholar konsisten membalas HTTP 429 (rate limit publik), sehingga Crossref
    sendirian menyisakan banyak kandidat tanpa teks.

    Rentang tahun DIFILTER DI PYTHON. Endpoint pencarian DOAJ ternyata BUKAN
    Lucene: menambahkan operator seperti "AND" atau "year:[2022 TO 2026]"
    membuat seluruh indeks (4 juta+ dokumen) dianggap cocok, jadi operator
    sengaja tidak dipakai -- hanya kata kunci polos.
    """
    if kind != "article":
        return []                      # DOAJ tidak memuat buku
    clean_q = re.sub(r"[^\w\s]", " ", query).strip()
    y_min, y_max = VL.active_year_range()
    if not clean_q:
        return []
    # DOAJ memakai AND sederhana atas SEMUA kata, jadi kueri panjang
    # ("Beban Kerja indikator pengukuran") sering 0 hasil. Coba bertahap:
    # kueri penuh -> tanpa kata teknis -> inti variabel (2 kata pertama).
    teknis = {"dimensi", "indikator", "pengukuran", "operasionalisasi",
              "dimension", "indicator", "measurement"}
    kata_tanpa_teknis = [w for w in clean_q.split() if w.lower() not in teknis]
    variants = [clean_q]
    if kata_tanpa_teknis and len(kata_tanpa_teknis) < len(clean_q.split()):
        variants.append(" ".join(kata_tanpa_teknis))
    if len(kata_tanpa_teknis) > 2:
        variants.append(" ".join(kata_tanpa_teknis[:2]))

    data = None
    for qv in variants:
        url = f"{DOAJ_BASE}/{urllib.parse.quote(qv)}?pageSize=10"
        data = _http_json(url, "doaj",
                          max_retry=2 if fast else 3,
                          backoff_base=2.0 if fast else None)
        if data and data.get("results"):
            break
    if not data:
        return []
    out = []
    for item in data.get("results", []) or []:
        bib = item.get("bibjson") or {}
        doi = None
        for ident in bib.get("identifier") or []:
            if str(ident.get("type", "")).lower() == "doi" and ident.get("id"):
                doi = str(ident["id"]).strip()
                break
        if not doi:
            continue                   # tanpa DOI -> tak bisa diverifikasi
        # Utamakan tautan berlabel "pdf"; baru fallback URL berakhiran .pdf
        pdf = None
        for link in bib.get("link") or []:
            if str(link.get("type", "")).lower() == "pdf" and link.get("url"):
                pdf = str(link["url"]).strip()
                break
        if not pdf:
            for link in bib.get("link") or []:
                u = str(link.get("url") or "")
                if u.lower().split("?")[0].endswith(".pdf"):
                    pdf = u
                    break
        try:
            year = int(bib.get("year") or 0)
        except (TypeError, ValueError):
            year = 0
        if year and not (y_min <= year <= y_max):
            continue
        jr = bib.get("journal") or {}
        out.append({
            "doi": f"https://doi.org/{doi}",
            "title": (bib.get("title") or "Tanpa Judul").strip(),
            "year": year,
            "journal": jr.get("title") or "Jurnal Ilmiah",
            "kind": kind,
            "pdf_url": pdf or None,
            "authors": [a.get("name") for a in bib.get("author") or [] if a.get("name")],
            "volume": jr.get("volume"),
            "issue": jr.get("issue"),
            "first_page": None,
            "last_page": None,
            "abstract": bib.get("abstract") or "",
            "_source": "doaj",
        })
    return out


def _resolve_pdf_unpaywall(cand):
    """
    Kandidat Crossref/S2 sering belum punya tautan PDF karena mereka hanyalah
    pendaftar metadata. Unpaywall menerjemahkan DOI -> PDF Open Access paling layak.
    Gratis, tanpa API key (cukup email). Mengisi cand["pdf_url"] bila berhasil.
    """
    if cand.get("pdf_url") or not cand.get("doi"):
        return cand
    doi = str(cand["doi"]).replace("https://doi.org/", "")
    url = f"{UNPAYWALL_BASE}/{urllib.parse.quote(doi)}?email={CONTACT_EMAIL}"
    data = _http_json(url, "unpaywall", timeout=12, max_retry=2)
    if not data:
        return cand
    locs = [data.get("best_oa_location")] + list(data.get("oa_locations") or [])
    for loc in locs:
        if isinstance(loc, dict) and loc.get("url_for_pdf"):
            cand["pdf_url"] = loc["url_for_pdf"]
            return cand
    for loc in locs:
        u = str((loc or {}).get("url") or "")
        if u.lower().split("?")[0].endswith(".pdf"):
            cand["pdf_url"] = u
            return cand
    return cand


def _run_search_job(source, query, kind, limit):
    """Satu pekerjaan pencarian (dipanggil dari thread). Aman gagal tanpa merusak yang lain.

    Selalu memakai profil CEPAT: di keadaan multi-sumber, sumber yang tertahan
    sebaiknya menyerah cepat supaya sumber sehat lain tidak ikut tertahan.
    """
    try:
        if source == "openalex":
            cands = _search_openalex(query, kind=kind, limit=limit, fast=True)
            for c in cands:
                c.setdefault("_source", "openalex")
            return source, cands
        if source == "crossref":
            return source, _search_crossref(query, kind=kind, limit=limit, fast=True)
        if source == "s2":
            return source, _search_s2(query, kind=kind, limit=limit, fast=True)
        if source == "doaj":
            return source, _search_doaj(query, kind=kind, limit=limit, fast=True)
    except Exception as e:
        print(f"[tracer] {source} gagal: {e}")
    return source, []


def _relevance(cand, seed):
    """
    Skor relevansi kandidat terhadap variabel yang dicari.

    Tanpa pengurutan ini, kandidat terbaru ber-DOI (mis. makalah "Green HRM" atau
    "Talent Management") bisa menghabiskan 8 slot percobaan padahal tak pernah
    memuat tabel dimensi/indikator variabel yang dicari -- dan ekstraksi otomatis
    tetap gagal walaupun LLM-nya sehat.
    """
    if not seed:
        return 0
    s = seed.lower().strip()
    judul = (cand.get("title") or "").lower()
    abstr = (cand.get("abstract") or "").lower()
    skor = 0
    if s and s in judul:
        skor += 3               # judul memuat nama variabel -> sinyal terkuat
    if s and s in abstr:
        skor += 1
    teks = judul + " " + abstr
    # Makalah pengukuran / kesiapan instrumen menyebut kata kunci ini
    if "dimensi" in teks or "dimension" in teks:
        skor += 1
    if "indikator" in teks or "indicator" in teks:
        skor += 1
    return skor


def _search_all(queries, kinds=("article", "book"), limit=4, seed="",
                wall_max=SEARCH_WALL_MAX):
    """
    Jalankan SEMUA sumber secara PARALEL -> gabung -> urutkan -> dedup DOI.
    Mengembalikan (list kandidat unik, statistik jumlah hasil per sumber).

    Urutan: (1) RELEVANSI tertinggi -- memuat nama variabel + kata dimensi/
    indikator; (2) sudah punya PDF sehingga langsung bisa diverifikasi; (3) tahun
    paling baru. Pengurutan relevansi penting karena hanya TRACE_MAX_CANDIDATES
    slot yang dicoba, jadi slot itu harus diisi makalah yang benar-benar mungkin
    memuat instrumen variabel tersebut.

    Pengaman kecepatan:
      1. batas waktu total (wall_max) -> sumber yang tertahan tak menahan hasil;
      2. begitu kandidat mencukupi dan SEARCH_MIN_WAIT terlampaui, pekerjaan yang
         masih berjalan dilepas -- telusur tetap cepat walaupun OpenAlex diblokir.
    """
    jobs = []
    for qi, q in enumerate(queries):
        for kind in kinds:
            jobs.append(("openalex", q, kind))
        jobs.append(("crossref", q, "article"))
        jobs.append(("doaj", q, "article"))
        if qi == 0:      # sumber cadangan cukup divalidasi lewat kueri utama
            jobs.append(("s2", q, "article"))

    stats, pooled = {}, []
    # Satu worker per pekerjaan -> tidak ada sumber yang harus mengantre di
    # belakang sumber lain yang sedang menahan backoff.
    ex = ThreadPoolExecutor(max_workers=len(jobs))
    fut_src = {ex.submit(_run_search_job, s, q, k, limit): s for s, q, k in jobs}
    started = time.time()
    try:
        for fut in as_completed(list(fut_src), timeout=wall_max):
            try:
                source, cands = fut.result()
            except Exception as e:
                print(f"[tracer] pekerjaan pencarian gagal: {e}")
                continue
            stats[source] = stats.get(source, 0) + len(cands)
            pooled.extend(cands)
            # Sudah cukup kandidat dan SEARCH_MIN_WAIT terlampaui -> berhenti
            # menunggu. Sumber apa pun yang masih tertahan (OpenAlex kena backoff,
            # sumber lambat) dilepas begitu saja: hasil sehat yang lain lebih
            # berharga daripada menunggu tambahan yang tak menentu.
            elapsed = time.time() - started
            if len(pooled) >= SEARCH_ENOUGH and elapsed >= SEARCH_MIN_WAIT:
                tertahan = sum(1 for f in fut_src if not f.done())
                if tertahan:
                    print(f"[tracer] {len(pooled)} kandidat terkumpul dalam {elapsed:.1f}s, "
                          f"{tertahan} pekerjaan masih tertahan -> pencarian dilepas")
                break
    except TimeoutError:
        print(f"[tracer] pencarian dihentikan pada batas {wall_max:.0f}s "
              f"({len(pooled)} kandidat dari {stats})")
    except Exception as e:
        print(f"[tracer] pencarian paralel gagal: {e}")
    try:
        ex.shutdown(wait=False, cancel_futures=True)
    except TypeError:          # Python < 3.9 tidak mengenal cancel_futures
        ex.shutdown(wait=False)

    pooled.sort(key=lambda c: (-_relevance(c, seed),
                               0 if c.get("pdf_url") else 1,
                               -(c.get("year") or 0)))
    seen, uniq = set(), []
    for c in pooled:
        doi = (c.get("doi") or "").lower()
        if not doi or doi in seen:
            continue
        seen.add(doi)
        uniq.append(c)
    return uniq, stats


# ---------------------------------------------------------------------------
# 2-3. Unduh & ekstrak teks kandidat
# ---------------------------------------------------------------------------

def _prepare_candidate(cand, idx):
    """Unduh PDF kandidat ke folder sementara, ekstrak teksnya."""
    if not cand.get("pdf_url"):
        return None
    os.makedirs(TMP_DIR, exist_ok=True)
    path = os.path.join(TMP_DIR, f"cand_{idx}.pdf")
    if not DTE.download_pdf(cand["pdf_url"], path, timeout=TRACE_TIMEOUT_PDF):
        return None
    text, _ = DTE.extract_any_text(path, max_pages=None)
    if not text or len(text.strip()) < 500:
        return None
    return text


def _fetch_text(idx, cand):
    """Satu kandidat: bila belum punya PDF, ungkap dulu lewat Unpaywall, lalu unduh."""
    try:
        if not cand.get("pdf_url"):
            _resolve_pdf_unpaywall(cand)
        return idx, _prepare_candidate(cand, idx)
    except Exception as e:
        print(f"[tracer] unduh kandidat {idx} gagal: {e}")
        return idx, None


def _download_all(candidates):
    """
    Unduh & ekstrak PDF semua kandidat secara PARALEL.
    Tahap ini I/O-bound (timeout 30 dtk per PDF) sehingga serial bisa memakan
    ber-menit; paralel memangkasnya menjadi sekitar waktu kandidat terlambat.
    Mengembalikan {index_kandidat: teks}.
    """
    texts = {}
    if not candidates:
        return texts
    os.makedirs(TMP_DIR, exist_ok=True)
    with ThreadPoolExecutor(max_workers=min(4, len(candidates))) as ex:
        futs = [ex.submit(_fetch_text, i, c) for i, c in enumerate(candidates)]
        for fut in futs:
            try:
                i, t = fut.result()
            except Exception as e:
                print(f"[tracer] unduh kandidat gagal: {e}")
                continue
            if t:
                texts[i] = t
    return texts


def _focus_window(text, max_chars=TRACE_MAX_CHARS, var_name=""):
    """
    Ambil jendela teks yang paling KAYA kata kunci dimensi/indikator.
    
    Strategi:
      1. Pisah teks jadi chunk 1000 char (overlap 200)
      2. Skor tiap chunk: hitung kemunculan kata kunci relevan
      3. Ambil top-N chunk dengan skor tertinggi TANPA OVERLAP, gabung hingga max_chars
      
    Kata kunci: nama variabel + istilah metodologi (id/en).
    """
    # Normalisasi spasi
    text = re.sub(r"[ \t]+", " ", text)
    text_lower = text.lower()
    
    # Kata kunci: variabel + istilah standar dimensi/indikator
    keywords = {"dimensi", "indikator", "dimension", "indicator", 
                "measurement", "scale", "item", "butir", "statement",
                "skala", "pengukuran", "operasionalisasi"}
    if var_name:
        # Tambahkan kata-kata dari nama variabel (pisah spasi)
        for w in var_name.lower().split():
            if len(w) > 2:
                keywords.add(w)
    
    # Bagi chunk 1000 char, overlap 200
    chunk_size = 1000
    overlap = 200
    chunks = []
    for start in range(0, len(text), chunk_size - overlap):
        end = min(start + chunk_size, len(text))
        chunk = text[start:end]
        if len(chunk.strip()) < 100:
            continue
        # Skor: hitung kemunculan kata kunci (case-insensitive)
        chunk_lower = chunk.lower()
        score = sum(chunk_lower.count(k) for k in keywords)
        # Bonus kecil kalau chunk awal (biasanya abstrak/pendahuluan)
        if start < 2000:
            score += 1
        chunks.append((score, start, end, chunk))
    
    # Urutkan skor descending
    chunks.sort(key=lambda x: x[0], reverse=True)
    
    # Pilih top chunk, MERGE overlapping regions
    selected_regions = []  # list of (start, end)
    total = 0
    
    for score, start, end, chunk in chunks:
        if total >= max_chars:
            break
            
        # Cek overlap dengan region yang sudah dipilih
        # Jika overlap, merge region (extend end)
        merged = False
        for i, (s_start, s_end) in enumerate(selected_regions):
            if not (end <= s_start or start >= s_end):  # ada overlap
                # Merge: extend region
                new_start = min(start, s_start)
                new_end = max(end, s_end)
                selected_regions[i] = (new_start, new_end)
                merged = True
                break
        
        if not merged:
            selected_regions.append((start, end))
        
        # Rebuild selected_regions: merge adjacent/overlapping regions
        selected_regions.sort(key=lambda x: x[0])
        merged_regions = []
        for r_start, r_end in selected_regions:
            if merged_regions and r_start <= merged_regions[-1][1] + 50:  # gap <= 50 char -> merge
                merged_regions[-1] = (merged_regions[-1][0], max(merged_regions[-1][1], r_end))
            else:
                merged_regions.append((r_start, r_end))
        selected_regions = merged_regions
        
        # Hitung total char dari merged regions
        total = sum(r_end - r_start for r_start, r_end in selected_regions)
        if total > max_chars:
            # Trim region terakhir
            excess = total - max_chars
            last_start, last_end = selected_regions[-1]
            selected_regions[-1] = (last_start, last_end - excess)
            total = max_chars
    
    # Ekstrak teks dari region yang dipilih
    selected_chunks = [text[r_start:r_end] for r_start, r_end in selected_regions]
    return " ... ".join(selected_chunks)


# ---------------------------------------------------------------------------
# 4. Ekstraksi Gemini (SATU sumber, kutipan literal)
# ---------------------------------------------------------------------------

PROMPT_EXTRACT = """Anda adalah ahli metodologi penelitian kuantitatif Manajemen. Dari TEKS SUMBER di bawah ini, ekstrak DIMENSI dan INDIKATOR untuk variabel "{var_name}" (peran: {role} dalam model X1, X2 -> M -> Y).

Petunjuk KERAS:
1. Hanya gunakan informasi yang BENAR-BENAR ADA di teks sumber. DILARANG menambah dimensi/indikator dari pengetahuan umum.
2. Setiap indikator adalah KUTIPAN LITERAL (persis kata-kata) yang ada di teks sumber, bukan parafrase.
3. Bila teks memuat tabel/pembahasan indikator variabel tersebut, gunakan itu sebagai dasar.
4. Hasilkan SELURUH dimensi yang benar-benar tertulis di sumber, maksimal {dmax}; tiap dimensi 2-4 indikator (minimal 2 per dimensi). Sistem akan memilih 2-7 dimensi paling relevan, jadi jangan memotong dimensi sumber.
5. Bila teks TIDAK memuat dimensi/indikator variabel itu, balas JSON: {{"dimensi": [], "alasan": "..."}}
6. HANYA balas JSON valid tanpa teks lain.

=== TEKS SUMBER (potongan) ===
{sumber}

=== KELUARAN JSON ===
{{
  "dimensi": [
    {{"nama": "Nama dimensi (bila ada di teks)", "indikator": ["kutipan literal 1", "kutipan literal 2"]}}
  ],
  "alasan": "opsional"
}}
"""


PROMPT_TOPUP = """Anda adalah ahli metodologi penelitian kuantitatif Manajemen. Dari SATU TEKS SUMBER di bawah ini, LENGKAPI dimensi dan indikator untuk variabel "{var_name}" (peran: {role} dalam model X1, X2 -> M -> Y).

Dimensi/indikator yang SUDAH ditemukan (jangan sekadar diulang; carilah TAMBAHANNYA bila teks memang memuatnya):
{existing}

Petunjuk KERAS:
1. Gunakan HANYA informasi yang BENAR-BENAR ADA di teks sumber. DILARANG menambah dari pengetahuan umum.
2. Setiap indikator adalah KUTIPAN LITERAL (persis kata-kata) dari teks sumber, bukan parafrase.
3. Kembalikan set LENGKAP: gabungkan yang sudah ada dengan tambahan, hingga seluruh dimensi sumber (maksimal {dmax}) dan SETIAP dimensi 2-4 indikator (minimal 2).
4. Bila teks memang TIDAK memuat cukup dimensi/indikator, kembalikan sebanyak yang benar-benar ada (jangan memaksa mengarang).
5. HANYA balas JSON valid tanpa teks lain.

=== TEKS SUMBER (jendela diperluas) ===
{sumber}

=== KELUARAN JSON ===
{{
  "dimensi": [
    {{"nama": "Nama dimensi (bila ada di teks)", "indikator": ["kutipan literal 1", "kutipan literal 2"]}}
  ],
  "grand_theory": "bila tertulis di teks, jika tidak kosongkan",
  "tokoh_ahli": "bila tertulis di teks, jika tidak kosongkan",
  "definisi": "bila tertulis di teks, jika tidak kosongkan"
}}
"""


def _extract_dims_from_text(var_name, role, source_text, config, doi=None, diag=None):
    """
    Ekstrak dimensi & indikator dari teks sumber dengan cache per (DOI, var_name).
    
    Jika DOI disediakan dan hasil sudah di-cache -> kembalikan cache (tanpa panggil LLM).
    diag (opsional): dict yang diisi alasan keberhasilan/kegagalan secara jujur.
    """
    if diag is not None:
        diag.clear()
        diag.update({"stage": "start", "doi": doi or "",
                     "text_len": len(source_text or "")})

    # 1. Cek cache jika DOI tersedia
    if doi:
        cache = _load_extraction_cache()
        key = _cache_key(doi, var_name)
        if key and key in cache:
            cached = cache[key]
            print(f"[cache] HIT: {key} -> pakai cache (hemat LLM)")
            if diag is not None:
                dims = cached.get("dimensi", []) or []
                diag.update({"stage": "cache_hit",
                             "valid_dims": len(dims),
                             "valid_inds": sum(len(d.get("indikator", [])) for d in dims),
                             "model": cached.get("_model", "")})
            return cached
    
    # 2. Ekstrak via LLM (matriks)
    full_text = _focus_window(source_text, var_name=var_name)
    prompt = PROMPT_EXTRACT.format(var_name=var_name, role=role, sumber=full_text,
                                   dmin=TRACE_MIN_DIMENSIONS,
                                   dmax=TRACE_SOURCE_MAX_DIMENSIONS)
    chain = extract_chain()   # #4 Nemotron didahulukan (hemat kuota Gemini)
    timeout = max(int(config.get("timeout_seconds", 12)) * 3, 60)
    keys = gemini_keys(config)
    try:
        raw = _call_gemini_matrix(keys, chain, prompt, timeout)
        if diag is not None:
            try:
                preview = json.dumps(raw, ensure_ascii=False)[:400]
            except Exception:
                preview = str(raw)[:400]
            diag.update({"stage": "model_response",
                         "model": last_matrix_model(),
                         "raw_type": type(raw).__name__,
                         "raw_preview": preview})
        cand = _validate_llm_ontology(raw, var_name=var_name)
        if cand and cand.get("dimensi"):
            cand["_model"] = last_matrix_model()   # provenance: nemotron/gemini
            if diag is not None:
                dims = cand.get("dimensi", []) or []
                diag.update({"stage": "validated",
                             "valid_dims": len(dims),
                             "valid_inds": sum(len(d.get("indikator", [])) for d in dims)})
            # 3. Simpan ke cache jika DOI tersedia
            if doi:
                cache = _load_extraction_cache()
                key = _cache_key(doi, var_name)
                if key:
                    cache[key] = cand
                    _save_extraction_cache(cache)
                    print(f"[cache] SAVE: {key}")
            return cand
        if diag is not None:
            dims = (cand or {}).get("dimensi", []) or []
            diag.update({"stage": "schema_rejected",
                         "valid_dims": len(dims),
                         "valid_inds": sum(len(d.get("indikator", [])) for d in dims)})
    except Exception as e:
        print(f"[tracer] ekstraksi matriks gagal: {type(e).__name__} {str(e)[:60]}")
        if diag is not None:
            diag.update({"stage": "model_error",
                         "error": f"{type(e).__name__}: {str(e)[:300]}"})
    return None



# ---------------------------------------------------------------------------
# #5 BATCH EXTRACTION — beberapa sumber dalam SATU panggilan LLM
# ---------------------------------------------------------------------------
# Sebelumnya: 1 paper = 1 panggilan LLM (8 kandidat = 8 panggilan).
# Sekarang  : 3 paper = 1 panggilan  -> panggilan turun ~60-70%, hemat kuota.
# JAMINAN ANTI-HALLUCINATION tetap sama: setiap kutipan indikator masih
# diverifikasi literal terhadap TEKS PAPER BERSANGKUTAN oleh pemanggil,
# jadi kalau model mencampur kutipan antar paper, hasilnya terbuang otomatis.
# ---------------------------------------------------------------------------

TRACE_BATCH_SIZE = 3       # jumlah sumber per panggilan LLM

PROMPT_EXTRACT_BATCH = """Anda adalah ahli metodologi penelitian kuantitatif Manajemen. Ekstrak DIMENSI dan INDIKATOR untuk variabel "{var_name}" (peran: {role} dalam model X1, X2 -> M -> Y) dari {n} SUMBER TERPISAH.

Petunjuk KERAS:
1. Tiap sumber punya nomor (SUMBER 0, SUMBER 1, ...). Untuk SUMBER bernomor i, hasilkan HANYA dari teks SUMBER i. DILARANG mencampur atau meminjam kutipan antar sumber.
2. Setiap indikator adalah KUTIPAN LITERAL (persis kata-kata) yang ada di teks sumber bersangkutan, bukan parafrase.
3. Bila satu sumber TIDAK memuat dimensi/indikator variabel itu, balas "dimensi": [] untuk nomor tersebut. Jangan mengarang.
4. Per sumber hasilkan SELURUH dimensi yang benar-benar tertulis (maksimal {dmax}); tiap dimensi 2-4 indikator (minimal 2 per dimensi). Sistem memilih 2-7 dimensi final, jadi jangan memotong dimensi sumber.
5. Sertakan juga definisi, grand_theory, dan tokoh_ahli yang benar-benar tertulis di sumber bersangkutan (kosongkan bila tidak ada).
6. HANYA balas JSON valid tanpa teks lain.

=== DAFTAR SUMBER ===
{sumber_list}

=== KELUARAN JSON ===
{{
  "hasil": [
    {{"id": 0, "dimensi": [{{"nama": "Nama dimensi (bila ada di teks)", "indikator": ["kutipan literal 1", "kutipan literal 2"]}}], "definisi": "", "grand_theory": "", "tokoh_ahli": ""}},
    {{"id": 1, "dimensi": []}}
  ]
}}
"""


def _parse_batch_items(raw):
    """Ubah JSON batch menjadi {id: item}. Bila model membalas format tunggal
    (mengabaikan struktur batch), kembalikan {0: objek} agar tetap terpakai."""
    if not isinstance(raw, dict):
        return {}
    items = {}
    hasil = raw.get("hasil")
    if isinstance(hasil, list):
        for i, it in enumerate(hasil):
            if isinstance(it, dict) and "dimensi" in it:
                try:
                    items[int(it.get("id", i))] = it
                except (TypeError, ValueError):
                    items[i] = it
    elif "dimensi" in raw:
        items[0] = raw
    return items


def _extract_dims_batch(var_name, role, papers, config):
    """
    #5 Ekstraksi BATCH: beberapa sumber dalam SATU panggilan LLM.

    papers : list of (idx, text) — idx adalah nomor urut kandidat asli.
    Mengembalikan dict {idx: cand | None} (cand tervalidasi _validate_llm_ontology).
    Verifikasi kutipan literal TETAP dilakukan pemanggil terhadap teks masing-masing.
    """
    out = {}
    if not papers:
        return out

    # Satu sumber saja -> pakai jalur tunggal yang sudah teruji.
    if len(papers) == 1:
        idx, text = papers[0]
        out[idx] = _extract_dims_from_text(var_name, role, text, config)
        return out

    # PENTING: prune tiap paper dulu (sama seperti jalur tunggal) supaya
    # 3 paper x 15.000 char tidak membanjiri konteks model.
    blocks = [f"=== SUMBER {idx} (potongan) ===\n"
              f"{_focus_window(text, var_name=var_name)}"
              for idx, text in papers]
    prompt = PROMPT_EXTRACT_BATCH.format(
        var_name=var_name, role=role, n=len(papers),
        sumber_list="\n\n".join(blocks),
        dmin=TRACE_MIN_DIMENSIONS, dmax=TRACE_SOURCE_MAX_DIMENSIONS)
    chain = extract_chain()   # #4 Nemotron didahulukan
    timeout = max(int(config.get("timeout_seconds", 12)) * 4, 90)
    keys = gemini_keys(config)

    try:
        raw = _call_gemini_matrix(keys, chain, prompt, timeout, max_tokens=3072)
    except Exception as e:
        # Batch gagal -> jangan buang kandidat: coba per-sumber seperti biasa.
        print(f"[tracer] batch {len(papers)} sumber gagal: {type(e).__name__} "
              f"{str(e)[:60]} -> fallback per-sumber")
        for idx, text in papers:
            out[idx] = _extract_dims_from_text(var_name, role, text, config)
        return out

    items = _parse_batch_items(raw)
    if not items:
        print("[tracer] batch: JSON batch kosong -> fallback per-sumber")
        for idx, text in papers:
            out[idx] = _extract_dims_from_text(var_name, role, text, config)
        return out

    for idx, _text in papers:
        item = items.get(idx)
        if item is None:
            out[idx] = None          # model tidak menjawab nomor ini
            continue
        try:
            # PENTING: item sudah berupa dict -> kirim LANGSUNG.
            # (JSON string ditolak _validate_llm_ontology -> selalu None.)
            cand = _validate_llm_ontology(item, var_name=var_name)
        except Exception:
            cand = None
        if cand:
            cand["_model"] = last_matrix_model()   # provenance: nemotron/gemini
        if cand and not cand.get("dimensi"):
            cand = None
        out[idx] = cand
    return out


# ---------------------------------------------------------------------------
# 5-6. Verifikasi & simpan
# ---------------------------------------------------------------------------

def _check_pass(clean_dims, report):
    n_ind = sum(len(d.get("indikator", [])) for d in clean_dims)
    if not clean_dims:
        return False
    # Konstruk datar: satu dimensi sah bila memuat >=6 indikator literal.
    if len(clean_dims) == 1:
        return (n_ind >= TRACE_MIN_INDICATORS
                and report.get("passed", 0) >= TRACE_MIN_INDICATORS)
    return (len(clean_dims) >= TRACE_MIN_DIMENSIONS
            and n_ind >= TRACE_MIN_INDICATORS
            and report.get("passed", 0) >= TRACE_MIN_INDICATORS)


def _check_accept(clean_dims, report):
    """Lantai keras: terima sumber eksplisit >=2 dimensi & >=4 indikator.

    Konstruk datar (1 dimensi) diterima bila memuat >=4 indikator literal.
    Dipakai HANYA setelah setiap indikator lolos verifikasi kutipan literal.
    Target 6 indikator tetap diupayakan lebih dulu lewat LAPIS 2; fungsi ini
    mencegah sumber nyata (mis. 4 indikator Dessler) dibuang percuma.
    """
    n_ind = sum(len(d.get("indikator", [])) for d in clean_dims)
    if not clean_dims:
        return False
    if len(clean_dims) == 1:
        return (n_ind >= TRACE_MIN_INDICATORS_ACCEPT
                and report.get("passed", 0) >= TRACE_MIN_INDICATORS_ACCEPT)
    return (len(clean_dims) >= TRACE_MIN_DIMENSIONS_ACCEPT
            and n_ind >= TRACE_MIN_INDICATORS_ACCEPT
            and report.get("passed", 0) >= TRACE_MIN_INDICATORS_ACCEPT)


def validate_ontology_complete(onto, min_indicators=4):
    """Validasi kelengkapan ontologi SEBELUM menulis Bab III-V.

    Kebijakan generated_v1:
      - setiap variabel 2-7 dimensi;
      - setiap dimensi 2-4 indikator operasional;
      - total minimal 4 indikator per variabel (2 dimensi x 2 indikator).
    """
    variables = (onto or {}).get("variables") or {}
    problems, summary = [], {}
    for role in ("x1", "x2", "m", "y"):
        var = variables.get(role) or {}
        name = var.get("nama_resmi") or var.get("input_name") or role.upper()
        dims = var.get("dimensi") or []
        n_ind = sum(len(d.get("indikator") or []) for d in dims)
        summary[role] = {"name": name, "dimensions": len(dims),
                         "indicators": n_ind}
        if not (2 <= len(dims) <= 7):
            problems.append(
                f"{role.upper()} {name}: {len(dims)} dimensi (wajib 2-7)")
        for d in dims:
            n = len(d.get("indikator") or [])
            if not (2 <= n <= 4):
                problems.append(
                    f"{role.upper()} {name} / {d.get('nama', '?')}: "
                    f"{n} indikator (wajib 2-4)")
        if n_ind < int(min_indicators):
            problems.append(
                f"{role.upper()} {name}: total {n_ind} indikator "
                f"(minimal {int(min_indicators)})")
    return (not problems), problems, summary


def _extract_doi(text):
    """Ambil DOI pertama dari teks berkas (untuk memakai cache ekstraksi).

    PDF sering memotong DOI di tengah kata karena pembungkusan baris,
    mis. '10.22441/indikat' + newline + 'or.v8i3.26506'. Gabungkan hanya
    baris yang diikuti huruf kecil agar tidak menelan kata berikutnya.
    """
    probe = re.sub(r"(?<=[A-Za-z0-9])\r?\n(?=[a-z])", "", (text or "")[:40000])
    m = re.search(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", probe)
    if not m:
        return ""
    return m.group(0).rstrip(".,;)]}\"'")


def _diagnose_extract_failure(diag):
    """Ubah diagnostik internal menjadi kalimat jujur untuk pengguna/log."""
    diag = diag or {}
    stage = diag.get("stage")
    if stage == "model_error":
        return ("Model gagal dipanggil: %s" % diag.get("error", "tanpa keterangan"))
    if stage == "schema_rejected":
        return ("Output model ditolak skema: %s dimensi valid, %s indikator "
                "(minimal 2 dimensi berindikator)." %
                (diag.get("valid_dims", 0), diag.get("valid_inds", 0)))
    if stage == "empty":
        return "Model mengembalikan JSON tanpa dimensi/indikator."
    if stage == "text_too_short":
        return "Teks berkas terlalu pendek untuk diekstrak (%s karakter)." % diag.get("text_len", 0)
    return "Model tidak mengembalikan hasil yang bisa dipakai."



def _check_pass_l2(clean_dims, report):
    """Ambang LAPIS 2 (sumber sama, jendela diperbesar): >=2 dimensi & >=6 indikator."""
    n_ind = sum(len(d.get("indikator", [])) for d in clean_dims)
    return (len(clean_dims) >= TRACE_MIN_DIMENSIONS_L2
            and n_ind >= TRACE_MIN_INDICATORS
            and report.get("passed", 0) >= TRACE_MIN_INDICATORS)


def _finalize_dims(dims, max_dims=TRACE_SOURCE_MAX_DIMENSIONS):
    """
    Rapikan dimensi hasil ekstraksi agar layak jadi instrumen:
      1. Gabungkan dimensi KEMBAR (nama sama / indikator tumpang-tindih tinggi,
         mis. 'Physical Work Environment' vs 'Lingkungan kerja fisik').
      2. Batasi jumlah dimensi SUMBER (<= 20) dan indikator per dimensi
         (<= TRACE_MAX_IND_PER_DIM). Seleksi final 2-7 dilakukan saat ontologi
         dibangun agar dimensi yang tidak dipakai dapat dicatat.
    Tidak ada item baru yang dibuat; hanya gabung/trim item yang sudah ada.
    """
    def _n(s):
        return re.sub(r"\s+", " ", str(s).strip().lower())

    merged, sets = [], []
    for d in dims or []:
        nama = str(d.get("nama", "")).strip()
        inds = [str(i).strip() for i in (d.get("indikator") or []) if str(i).strip()]
        key = {_n(i) for i in inds}
        target = None
        for j, prev in enumerate(merged):
            if _n(prev["nama"]) == _n(nama):
                target = j
                break
            pk = sets[j]
            if key and pk:
                union = len(key | pk)
                if union and (len(key & pk) / union) >= TRACE_DEDUP_JACCARD:
                    target = j
                    break
        if target is None:
            merged.append({"nama": nama, "indikator": inds})
            sets.append(set(key))
        else:
            prev, pk = merged[target], sets[target]
            for i in inds:
                ni = _n(i)
                if ni not in pk:
                    prev["indikator"].append(i)
                    pk.add(ni)
    out = []
    for d in merged[:max_dims]:
        inds = d["indikator"][:TRACE_MAX_IND_PER_DIM]
        if inds:
            out.append({"nama": d["nama"], "indikator": inds})
    return out


def _dims_fit_rumpun(var_name, dims):
    """
    True bila minimal satu nama dimensi menyentuh rumpun variabel (anti salah
    konstruk, mis. variabel Disiplin tetapi berdimensi Motivasi).
    Variabel yang tak dikenal peta rumpun -> dianggap lolos (None).
    """
    topic = _match_variable_topic(var_name)
    if not topic:
        return True
    return _dimensions_fit_topic(dims or [], topic)


def _make_library_entry(cand, clean, var_name, dims):
    """Bangun dict entri pustaka dari kandidat + dimensi terverifikasi."""
    entri_apa, _italic = _build_apa(cand)
    return {
        "variable": var_name,
        "synonyms": [var_name.lower()],
        "source": {
            "type": cand.get("_source") or "openalex",
            "label": _short_cite(cand),
            "apa": entri_apa,
            "doi": cand["doi"],
            "url": cand["pdf_url"] or cand["doi"],
            "year": cand["year"],
            "doc_kind": "artikel" if cand["kind"] == "article" else "buku",
            "journal_title": cand["title"],
        },
        "grand_theory": dims.get("grand_theory", ""),
        "tokoh_ahli": dims.get("tokoh_ahli", ""),
        "definisi": dims.get("definisi", ""),
        "dimensi": clean,
        "verified_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "engine": "llm:" + (dims.get("_model") or "unknown"),
    }


def _merge_dims(base_clean, add_dims):
    """
    Gabungkan dimensi Lapis 1 dengan hasil Lapis 2 (SUMBER YANG SAMA).
    Dimensi dipadankan by nama (case-insensitive); indikator didedup.
    Hasil akhir tetap diverifikasi ulang secara literal oleh pemanggil.
    """
    def _n(s):
        return re.sub(r"\s+", " ", str(s).strip().lower())

    out, index = [], {}
    for d in base_clean or []:
        nama = str(d.get("nama", "")).strip()
        key = _n(nama)
        node = {"nama": nama,
                "indikator": [str(i).strip() for i in (d.get("indikator") or []) if str(i).strip()]}
        index[key] = node
        out.append(node)
    for d in add_dims or []:
        nama = str(d.get("nama", "")).strip()
        if not nama:
            continue
        key = _n(nama)
        inds = [str(i).strip() for i in (d.get("indikator") or []) if str(i).strip()]
        if key in index:
            have = {_n(x) for x in index[key]["indikator"]}
            for x in inds:
                if _n(x) not in have:
                    index[key]["indikator"].append(x)
                    have.add(_n(x))
        else:
            node = {"nama": nama, "indikator": inds}
            index[key] = node
            out.append(node)
    return out


def _extract_topup(var_name, role, source_text, config, existing_clean):
    """
    LAPIS 2: minta model MELENGKAPI dimensi/indikator dari SUMBER YANG SAMA
    dengan jendela teks diperbesar. Semua item tetap wajib kutipan LITERAL
    (diverifikasi ulang oleh pemanggil) sehingga tidak ada karangan yang lolos.
    """
    existing = "\n".join(
        f"- {d.get('nama')}: {len(d.get('indikator') or [])} indikator"
        for d in (existing_clean or [])) or "- (belum ada)"
    big = _focus_window(source_text, max_chars=TRACE_TOPUP_CHARS, var_name=var_name)
    prompt = PROMPT_TOPUP.format(
        var_name=var_name, role=role, existing=existing, sumber=big,
        dmin=TRACE_MIN_DIMENSIONS, dmax=TRACE_SOURCE_MAX_DIMENSIONS)
    chain = extract_chain()
    timeout = max(int(config.get("timeout_seconds", 12)) * 4, 90)
    keys = gemini_keys(config)
    try:
        raw = _call_gemini_matrix(keys, chain, prompt, timeout, max_tokens=3072)
    except Exception as e:
        print(f"[tracer] topup matriks gagal: {type(e).__name__} {str(e)[:60]}")
        return None
    cand = _validate_llm_ontology(raw, var_name=var_name)
    if cand and cand.get("dimensi"):
        cand["_model"] = last_matrix_model()
        return cand
    return None


# ---------------------------------------------------------------------------
# 5b. DIMENSI DARI SUMBER MILIK JUDUL (bukan telusur OpenAlex sendiri)
# ---------------------------------------------------------------------------
def _ref_to_cand(ref):
    """Ubah referensi sumber milik judul (dari academic_connector) menjadi
    kandidat bergaya tracer agar bisa diunduh & diekstrak."""
    kind = "book" if (ref.get("type") == "buku") else "article"
    return {
        "doi": (ref.get("doi") or "").strip(),
        "title": ref.get("title", ""),
        "authors": ref.get("authors", []) or [],
        "year": ref.get("year"),
        "journal": ref.get("journal") or ref.get("publisher") or "",
        "volume": ref.get("volume", ""),
        "issue": ref.get("issue", ""),
        "first_page": "", "last_page": "",
        "kind": kind,
        "pdf_url": ref.get("pdf_url", "") or "",
        "_source": "title_pool",
    }


def extract_dims_from_refs(var_name, role, refs, judul="", instansi="",
                           max_try=6, cache_key=None):
    """Ekstrak dimensi & indikator untuk ``var_name`` dari SUMBER MILIK JUDUL.

    Berbeda dari trace_variable() yang menelusur OpenAlex sendiri, fungsi ini
    HANYA memakai daftar referensi yang sudah di-fetch untuk judul bersangkutan
    (mis. sources_y). Dengan begitu tiap judul memakai sumber dimensi yang
    BERBEDA (unik antar-judul), sedangkan dimensi/indikator tetap wajib
    terverifikasi sebagai kutipan LITERAL dari teks sumber.

    Mengembalikan dict:
        {"dimensi": [...], "jurnal_rujukan": {"sitasi", "entri"},
         "source_ref": {...}, "verified": {...}}
    atau None bila tak ada sumber yang lolos verifikasi.
    """
    if not refs:
        return None

    # 0. Cache per-judul (regenerasi judul yang sama -> sumber dimensi yang sama).
    if cache_key:
        cache = _load_dim_source_cache()
        hit = cache.get(cache_key)
        if hit:
            hit_doi = ((hit.get("source_ref") or {}).get("doi") or "").lower().strip()
            ref_dois = {(r.get("doi") or "").lower().strip() for r in refs}
            if hit_doi and hit_doi in ref_dois:
                print(f"[tracer] dimensi '{var_name}' dari cache judul (tanpa unduh/LLM)")
                return hit
            print(f"[tracer] cache dimensi judul '{var_name}' menunjuk sumber yang "
                  f"tak lagi ada di daftar -> riset ulang")

    config = load_config()
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return None

    tried = 0
    attempts = []
    for ref in refs:
        if tried >= max_try:
            break
        doi = (ref.get("doi") or "").strip()
        if not doi:
            continue
        tried += 1
        short = (ref.get("title") or "")[:48]
        try:
            cand = _ref_to_cand(ref)
            if not cand.get("pdf_url"):
                _resolve_pdf_unpaywall(cand)
            if not cand.get("pdf_url"):
                attempts.append(f"{short} (PDF OA tidak ada)")
                continue
            text = _prepare_candidate(cand, tried)
            if not text:
                attempts.append(f"{short} (teks gagal diunduh)")
                continue
            dims = _extract_dims_from_text(var_name, role, text, config, doi=doi)
            if not dims or not dims.get("dimensi"):
                attempts.append(f"{short} (tak memuat dimensi)")
                continue
            clean, vrep = VL.verify_indicators(dims.get("dimensi", []), text)
            if not _check_pass(clean, vrep):
                attempts.append(f"{short} (kutipan indikator tak cocok)")
                continue
            result = {
                "dimensi": clean,
                "definisi": dims.get("definisi", ""),
                "grand_theory": dims.get("grand_theory", ""),
                "tokoh_ahli": dims.get("tokoh_ahli", ""),
                "jurnal_rujukan": {
                    "sitasi": ref.get("apa_in_text", "") or "",
                    "entri": ref.get("apa_full", "") or "",
                },
                "source_ref": ref,
                "verified": vrep,
                "engine": "llm:" + (dims.get("_model") or "unknown"),
            }
            if cache_key:
                cache = _load_dim_source_cache()
                cache[cache_key] = result
                _save_dim_source_cache(cache)
            print(f"[tracer] dimensi '{var_name}' dari sumber judul: {short}")
            return result
        except Exception as e:
            attempts.append(f"{short} (error: {type(e).__name__})")

    print("[tracer] dimensi dari sumber judul GAGAL untuk "
          f"'{var_name}': " + " | ".join(attempts))
    return None


def trace_variable(var_name, role="X", judul="", instansi="", force=False):
    """
    Alur utama: pustaka -> pencarian live -> ekstraksi -> verifikasi -> simpan.
    Mengembalikan dict:
      {"status": "library" | "ok" | "needs_review" | "no_api",
       "entry": {...} | None, "report": {...}}
    """
    config = load_config()
    # Pakai daftar key; "api_key" tunggal tetap didukung untuk config lama.
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return {"status": "no_api", "entry": None, "report": {"note": "API nonaktif"}}

    # 1. Pustaka permanen
    if not force:
        existing = VL.load_entry(var_name)
        if existing and not VL.entry_needs_review(existing):
            return {"status": "library", "entry": existing,
                    "report": {"source": (existing.get("source") or {}).get("label", "")}}

    report = {"searched": [], "tried": 0, "dropped": []}
    queries = [
        f"{var_name} dimensi indikator",
        f"{var_name} kinerja pegawai indikator",
        f"indikator {var_name}",
    ]

    # Cari KE SEMUA SUMBER sekaligus (paralel): OpenAlex + Crossref +
    # Semantic Scholar + DOAJ. Bila satu sumber diblokir, sumber lain tetap
    # menyuplai kandidat, sehingga telusur tidak lagi mandek pada satu API.
    uniq, src_stats = _search_all(queries, limit=4, seed=var_name)
    report["sources"] = {k: v for k, v in src_stats.items() if v}
    report["searched"] = [c["title"][:70] for c in uniq[:8]]
    if src_stats:
        print(f"[tracer] hasil per sumber untuk '{var_name}': {report['sources']}")

    target = uniq[:TRACE_MAX_CANDIDATES]
    texts = _download_all(target)

    # Kandidat siap-proses (teks berhasil diunduh), tetap urut relevansi.
    ready = []
    for idx, cand in enumerate(target):
        report["tried"] += 1
        text = texts.get(idx)
        if text is None:
            report["dropped"].append(f"{cand['title'][:50]} (teks gagal diunduh)")
            continue
        ready.append((idx, cand, text))

    cache = _load_extraction_cache()   # snapshot; disimpan ulang tiap ada hasil baru
    partials = []   # calon LAPIS 2: punya dimensi terverifikasi tapi belum capai ambang

    # #5 Proses per BLOK (TRACE_BATCH_SIZE sumber per panggilan LLM) dan
    # berhenti segera setelah satu sumber lolos verifikasi.
    for start in range(0, len(ready), TRACE_BATCH_SIZE):
        chunk = ready[start:start + TRACE_BATCH_SIZE]

        # a) Ambil dari cache dulu (TANPA LLM) bila DOI+variabel pernah diekstrak.
        dims_map = {}
        pending = []        # [(idx, cand, text)] yang butuh panggilan LLM
        for idx, cand, text in chunk:
            doi = cand.get("doi")
            key = _cache_key(doi, var_name) if doi else None
            if key and key in cache:
                print(f"[cache] HIT: {key} -> pakai cache (hemat LLM)")
                dims_map[idx] = cache[key]
            else:
                pending.append((idx, cand, text))

        # b) Sisa kandidat: SATU panggilan LLM untuk seluruh blok.
        if pending:
            print(f"[tracer] batch '{var_name}': {len(pending)} sumber dalam 1 panggilan")
            new_dims = _extract_dims_batch(
                var_name, role, [(i, t) for i, _c, t in pending], config)
            for idx, cand, _text in pending:
                dims = new_dims.get(idx)
                dims_map[idx] = dims
                doi = cand.get("doi")
                key = _cache_key(doi, var_name) if doi else None
                if dims and key:
                    cache[key] = dims
                    print(f"[cache] SAVE: {key}")
            if any(new_dims.values()):
                _save_extraction_cache(cache)

        # c) Verifikasi kutipan literal terhadap TEKS paper masing-masing.
        for idx, cand, text in chunk:
            dims = dims_map.get(idx)
            if not dims:
                report["dropped"].append(f"{cand['title'][:50]} (tidak memuat dimensi/indikator)")
                continue
            clean, vrep = VL.verify_indicators(dims.get("dimensi", []), text)
            clean = _finalize_dims(clean)
            clean, vrep = VL.verify_indicators(clean, text)
            if not _dims_fit_rumpun(var_name, clean):
                report["dropped"].append(
                    f"{cand['title'][:50]} (salah rumpun '{_match_variable_topic(var_name)}')")
                continue
            if not _check_pass(clean, vrep):
                n_ind = sum(len(d.get("indikator", [])) for d in clean)
                # Sumber eksplisit dengan 4 indikator literal tetap sah untuk
                # PLS-SEM. Terima dengan catatan jujur, jangan buang.
                if _check_accept(clean, vrep):
                    entry = _make_library_entry(cand, clean, var_name, dims)
                    try:
                        path = VL.save_entry(entry)
                        entry["_file"] = path
                    except Exception as e:
                        report["dropped"].append(f"gagal simpan (eksplisit): {e}")
                        continue
                    report["tier"] = "1-explicit"
                    report["note"] = (
                        f"Sumber eksplisit memuat {len(clean)} dimensi/{n_ind} "
                        f"indikator (target {TRACE_MIN_INDICATORS}); diterima karena "
                        "semua indikator lolos verifikasi kutipan literal.")
                    return {"status": "ok", "entry": entry,
                            "report": {**report, "verified": vrep}}
                report["dropped"].append(
                    f"{cand['title'][:50]} (di bawah ambang: {len(clean)}dim/{n_ind}ind)")
                if clean:
                    partials.append({"cand": cand, "text": text,
                                     "clean": clean, "vrep": vrep, "dims": dims})
                continue
            entry = _make_library_entry(cand, clean, var_name, dims)
            try:
                path = VL.save_entry(entry)
                entry["_file"] = path
            except Exception as e:
                report["dropped"].append(f"gagal simpan: {e}")
                continue
            report["tier"] = "1"
            return {"status": "ok", "entry": entry,
                    "report": {**report, "verified": vrep}}

    # ------------------------------------------------------------------
    # LAPIS 2 (fallback): tak ada kandidat yang capai ambang Lapis 1.
    # Terima kandidat TERBAIK lalu CUKUPKAN kekurangannya DARI SUMBER YANG
    # SAMA dengan jendela teks diperbesar. Semua item wajib kutipan LITERAL
    # (diverifikasi ulang) -> tetap anti-halusinasi.
    # ------------------------------------------------------------------
    if partials:
        def _score(p):
            n_i = sum(len(d.get("indikator", [])) for d in p["clean"])
            return (len(p["clean"]) / float(TRACE_MIN_DIMENSIONS)
                    + n_i / float(TRACE_MIN_INDICATORS))
        partials.sort(key=_score, reverse=True)
        for p in partials[:TRACE_TOPUP_TRIES]:
            cand, text, base = p["cand"], p["text"], p["clean"]
            n_dim = len(base)
            n_ind = sum(len(d.get("indikator", [])) for d in base)
            print(f"[tracer] LAPIS 2 '{var_name}': cukupkan dari "
                  f"'{cand.get('title', '')[:48]}' (mulai {n_dim}dim/{n_ind}ind)")
            up = _extract_topup(var_name, role, text, config, base)
            if not up or not up.get("dimensi"):
                continue
            merged = _merge_dims(base, up.get("dimensi", []))
            merged, _ = VL.verify_indicators(merged, text)
            merged = _finalize_dims(merged)
            clean2, vrep2 = VL.verify_indicators(merged, text)
            n2 = sum(len(d.get("indikator", [])) for d in clean2)
            print(f"[tracer] LAPIS 2 '{var_name}' hasil: "
                  f"{len(clean2)}dim/{n2}ind (passed={vrep2.get('passed')})")
            if not _dims_fit_rumpun(var_name, clean2):
                print(f"[tracer] LAPIS 2 '{var_name}': salah rumpun "
                      f"'{_match_variable_topic(var_name)}' -> dilewati")
                continue
            if not _check_pass_l2(clean2, vrep2):
                continue
            meta = {
                "grand_theory": up.get("grand_theory") or p["dims"].get("grand_theory", ""),
                "tokoh_ahli": up.get("tokoh_ahli") or p["dims"].get("tokoh_ahli", ""),
                "definisi": up.get("definisi") or p["dims"].get("definisi", ""),
                "_model": up.get("_model") or p["dims"].get("_model"),
            }
            entry = _make_library_entry(cand, clean2, var_name, meta)
            try:
                path = VL.save_entry(entry)
                entry["_file"] = path
            except Exception as e:
                report["dropped"].append(f"gagal simpan (lapis 2): {e}")
                continue
            report["tier"] = "2"
            report["topped_up_from"] = (cand.get("title") or "")[:70]
            return {"status": "ok", "entry": entry,
                    "report": {**report, "verified": vrep2}}

    return {"status": "needs_review", "entry": None,
            "report": {**report,
                       "note": "Tidak ada sumber dalam rentang tahun sitasi aktif yang lolos verifikasi kutipan. "
                               "Silakan unggah jurnal/buku khusus untuk variabel ini."}}


def trace_from_upload(file_path, var_name, role="X", judul="", instansi="", display_name=None):
    """
    Jalur unggahan pengguna: ekstrak teks dari berkas -> LLM -> verifikasi -> simpan.
    Sumber tahun bebas (keputusan disepakati: unggahan bebas tahun).

    Perbaikan otonomi:
      - DOI di berkas dipakai untuk cache (tidak lagi memaksa doi=None).
      - Bila indikator < target 6, dicoba top-up dari BERKAS YANG SAMA (LAPIS 2).
      - Sumber eksplisit >=4 indikator tetap diterima (lantai keras PLS-SEM).
      - Kegagalan dilaporkan dengan alasan sebenarnya (stage/counts), bukan
        pesan generik "model tidak menemukan".
    """
    config = load_config()
    # Pakai daftar key; "api_key" tunggal tetap didukung untuk config lama.
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return {"status": "no_api", "entry": None, "report": {"note": "API nonaktif"}}

    text, kind = DTE.extract_any_text(file_path)
    if not text or len(text.strip()) < 300:
        diag = {"stage": "text_too_short", "text_len": len((text or "").strip())}
        return {"status": "needs_review", "entry": None,
                "report": {"note": _diagnose_extract_failure(diag), "diag": diag}}

    # DOI (bila ada di berkas) -> pakai cache; unggahan tanpa DOI tetap jalan.
    doi = _extract_doi(text)
    diag = {}
    dims = _extract_dims_from_text(var_name, role, text, config, doi=doi or None, diag=diag)
    if not dims:
        note = _diagnose_extract_failure(diag)
        return {"status": "needs_review", "entry": None,
                "report": {"note": note, "diag": diag}}

    clean, vrep = VL.verify_indicators(dims.get("dimensi", []), text)
    clean = _finalize_dims(clean)

    # LAPIS 2: cukupkan dari sumber yang sama bila masih di bawah target.
    if not _check_pass(clean, vrep) and clean:
        print(f"[upload] '{var_name}': {len(clean)}dim/"
              f"{sum(len(d.get('indikator', [])) for d in clean)}ind "
              "-> coba top-up dari berkas yang sama")
        try:
            up = _extract_topup(var_name, role, text, config, clean)
        except Exception as e:
            up = None
            diag["topup_error"] = f"{type(e).__name__}: {str(e)[:200]}"
        if up and up.get("dimensi"):
            merged = _merge_dims(clean, up.get("dimensi", []))
            merged, _ = VL.verify_indicators(merged, text)
            merged = _finalize_dims(merged)
            clean, vrep = VL.verify_indicators(merged, text)
            for k in ("grand_theory", "tokoh_ahli", "definisi"):
                if up.get(k) and not dims.get(k):
                    dims[k] = up[k]

    n_dim = len(clean)
    n_ind = sum(len(d.get("indikator", [])) for d in clean)
    passed = int(vrep.get("passed", 0) or 0)

    if not _check_accept(clean, vrep):
        note = (
            f"Terkumpul {n_dim} dimensi/{n_ind} indikator setelah verifikasi; "
            f"minimal {TRACE_MIN_DIMENSIONS_ACCEPT} dimensi/"
            f"{TRACE_MIN_INDICATORS_ACCEPT} indikator. "
            f"Kutipan lolos verifikasi: {passed}. "
            "Tidak ada indikator yang ditambahkan tanpa kutipan literal."
        )
        if diag.get("topup_error"):
            note += f" Top-up gagal: {diag['topup_error']}"
        return {"status": "needs_review", "entry": None,
                "report": {"note": note, "verified": vrep, "diag": diag}}

    label = display_name or os.path.basename(file_path)
    entry = {
        "variable": var_name,
        "synonyms": [var_name.lower()],
        "source": {
            "type": "upload",
            "label": label,
            "apa": label,
            "doi": doi,
            "url": ("https://doi.org/" + doi) if doi else "",
            "year": datetime.datetime.now().year,
            "doc_kind": "unggahan pengguna",
            "journal_title": label,
        },
        "grand_theory": dims.get("grand_theory", ""),
        "tokoh_ahli": dims.get("tokoh_ahli", ""),
        "definisi": dims.get("definisi", ""),
        "dimensi": clean,
        "verified_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "engine": "llm:" + (dims.get("_model") or "unknown"),
    }
    try:
        path = VL.save_entry(entry)
        entry["_file"] = path
    except Exception as e:
        return {"status": "needs_review", "entry": None,
                "report": {"note": f"gagal simpan: {e}", "diag": diag}}
    note = None
    if n_ind < TRACE_MIN_INDICATORS:
        note = (f"Sumber eksplisit memuat {n_dim} dimensi/{n_ind} indikator "
                f"(target {TRACE_MIN_INDICATORS}). Diterima karena semua indikator "
                "lolos verifikasi kutipan literal.")
    return {"status": "ok", "entry": entry,
            "report": {"verified": vrep, "source": label,
                       "indicator_count": n_ind, "target": TRACE_MIN_INDICATORS,
                       "note": note, "diag": diag}}



# ---------------------------------------------------------------------------
# Pembantu sitasi / APA
# ---------------------------------------------------------------------------

def _short_cite(cand):
    """Bentuk sitasi singkat APA: Penulis (tahun); Penulis & Penulis (tahun); dst."""
    names = []
    for a in cand.get("authors", []) or []:
        parts = a.strip().split()
        if parts:
            names.append(parts[-1])
    year = cand.get("year", 2023)
    if not names:
        return f"Peneliti ({year})"
    if len(names) == 1:
        return f"{names[0]} ({year})"
    if len(names) == 2:
        return f"{names[0]} & {names[1]} ({year})"
    return f"{names[0]} et al. ({year})"


def _build_apa(cand):
    """Entri APA lengkap memakai metadata OpenAlex asli (penulis, volume, halaman)."""
    from academic_connector import format_authors_apa
    authors_apa = format_authors_apa(cand.get("authors", []))
    title = cand.get("title", "").strip().rstrip(".")
    src = cand.get("journal", "Jurnal Ilmiah")
    year = cand.get("year", 2023)
    doi = cand.get("doi", "")
    vol, iss = cand.get("volume"), cand.get("issue")
    fp, lp = cand.get("first_page"), cand.get("last_page")
    vol_str = ""
    if vol and iss:
        vol_str = f", {vol}({iss})"
    elif vol:
        vol_str = f", {vol}"
    elif iss:
        vol_str = f", ({iss})"
    page_str = f", {fp}-{lp}" if fp and lp else (f", {fp}" if fp else "")
    entri = f"{authors_apa} ({year}). {title}. {src}{vol_str}{page_str}. {doi}"
    return entri, src


if __name__ == "__main__":
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    var = sys.argv[1] if len(sys.argv) > 1 else "Kinerja Pegawai"
    role = sys.argv[2] if len(sys.argv) > 2 else "Y"
    res = trace_variable(var, role=role,
                         judul=f"Pengaruh X terhadap {var}",
                         instansi="Instansi Pemerintah Daerah")
    print("\n=== HASIL TRACE ===")
    print("status:", res["status"])
    print("laporan:", json.dumps(res.get("report", {}), ensure_ascii=False, indent=2)[:1500])
    if res.get("entry"):
        e = res["entry"]
        print("\nSumber:", e["source"]["label"], "|", e["source"]["doi"])
        print("Indikator lolos:", sum(len(d["indikator"]) for d in e["dimensi"]))
        for d in e["dimensi"]:
            print(f"- {d['nama']}: {len(d['indikator'])} indikator")