# -*- coding: utf-8 -*-
"""
Korpus 95 Referensi Pendukung (references/02_pendukung) — DIMENSI & INDIKATOR ONLY.

Perjanjian dengan pengguna (diekskusi sesuai arahan):
  - 95 PDF pendukung dipakai HANYA sebagai KORPUS dimensi & indikator.
  - PDF ini TIDAK pernah menjadi sumber kutipan bab (daftar pustaka "Sumber
    Rujukan & Kutipan" tetap dibangun dari rentang tahun sitasi yang
    dikonfigurasi di Menu 1 — default 20 tahun, mis. 2006-2026).
  - Bila dimensi/indikator TIDAK ditemukan di korpus -> pencarian ONLINE
    (OpenAlex/Crossref) WAJIB dijalankan sebagai fallback.

Alur ekstraksi sama dengan jalur unggahan (variable_tracer._extract_dims_from_text)
tetapi TANPA efek samping menyimpan ke variable_library: hasil diverifikasi literal
terhadap teks PDF (variable_library.verify_indicators), dirapikan (_finalize_dims),
dan diterima hanya bila lolos lantai keras PLS-SEM (>=2 dimensi & >=4 indikator
kutipan literal yang lolos verifikasi).
"""

import json
import os
import re

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PENDUKUNG_FILE = os.path.join(CURRENT_DIR, "references", "pendukung.json")
PDF_DIR = os.path.join(CURRENT_DIR, "references", "02_pendukung")
_CORPUS_CACHE_FILE = os.path.join(CURRENT_DIR, "cache_corpus_dims.json")

_REGISTRY = None
_CORPUS_CACHE = None


# ---------------------------------------------------------------------------
# Registry korpus
# ---------------------------------------------------------------------------

def load_corpus_registry():
    """Entri pendukung.json (95 referensi pendukung yang sudah diimpor)."""
    global _REGISTRY
    if _REGISTRY is None:
        try:
            with open(PENDUKUNG_FILE, "r", encoding="utf-8") as f:
                _REGISTRY = (json.load(f) or {}).get("entries", []) or []
        except Exception as e:
            print(f"[corpus] gagal memuat pendukung.json: {e}")
            _REGISTRY = []
    return _REGISTRY


def _tokens(s):
    return set(re.findall(r"[a-z0-9]+", str(s or "").lower()))


def _fuzzy_score(a, b):
    ta, tb = _tokens(a), _tokens(b)
    if not ta or not tb:
        return 0.0
    return 2.0 * len(ta & tb) / (len(ta) + len(tb))


def find_corpus_options(var_name, limit=3):
    """Kandidat PDF korpus yang membahas variabel tsb (fuzzy match, maks. limit)."""
    registry = load_corpus_registry()
    scored = []
    cl = (var_name or "").strip().lower()
    for e in registry:
        v = (e.get("variable") or "").strip()
        if not v:
            continue
        s = _fuzzy_score(cl, v)
        vl = v.lower()
        if cl and (cl == vl):
            s += 1.0
        elif cl and (cl in vl or vl in cl):
            s += 0.5
        if s >= 0.30:
            scored.append((s, e))
    scored.sort(key=lambda x: -x[0])
    return [e for _, e in scored[:limit]]


# ---------------------------------------------------------------------------
# Cache hasil ekstraksi korpus (agar klik "Telusuri" berulang tidak boros kuota)
# ---------------------------------------------------------------------------

def _load_corpus_cache():
    global _CORPUS_CACHE
    if _CORPUS_CACHE is None:
        try:
            with open(_CORPUS_CACHE_FILE, "r", encoding="utf-8") as f:
                _CORPUS_CACHE = json.load(f)
        except Exception:
            _CORPUS_CACHE = {}
    return _CORPUS_CACHE


def _save_corpus_cache():
    try:
        with open(_CORPUS_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(_CORPUS_CACHE, f, ensure_ascii=False, indent=1)
    except Exception as e:
        print(f"[corpus] gagal simpan cache: {e}")


def _pdf_text(url_path):
    """Baca teks PDF korpus (url_path = /references/02_pendukung/<file>.pdf)."""
    fname = os.path.basename(str(url_path or ""))
    full = os.path.join(PDF_DIR, fname)
    if not fname or not os.path.exists(full):
        return "", "file_missing"
    try:
        import doc_text_extract as DTE
        text, _kind = DTE.extract_any_text(full)
        return text or "", "ok"
    except Exception as e:
        return "", f"extract_error:{e}"


def _counts(dims):
    dims = dims or []
    return len(dims), sum(len((d.get("indikator") or []) or []) for d in dims)


# ---------------------------------------------------------------------------
# Ekstraksi dimensi & indikator dari SATU PDF korpus
# ---------------------------------------------------------------------------

def build_corpus_option(entry, var_name, role="X1", instansi="Instansi Terkait",
                        title=""):
    """Ekstrak dims/indikator dari satu PDF korpus -> dict source-option.

    Mengembalikan dict dengan kunci: kind, key, label, apa, year, journal,
    doc_kind, url, download_kind, passed, note, dims, inds, onto (opsional),
    dims_blueprint (opsional), sumber.
    """
    url_path = entry.get("url") or ""
    fname = os.path.basename(str(url_path or ""))
    label = entry.get("label") or entry.get("apa") or (fname or "PDF Pendukung")
    apa = entry.get("apa") or label
    year = entry.get("year")
    journal = entry.get("journal_title") or ""
    doc_kind = entry.get("doc_kind") or ""
    key = f"corpus|{fname}"
    base = {
        "kind": "corpus",
        "key": key,
        "label": label,
        "apa": apa,
        "year": year,
        "journal": journal,
        "doc_kind": doc_kind,
        "url": url_path,
        "download_kind": "pdf",
        "passed": False,
        "note": "",
        "dims": 0,
        "inds": 0,
        "dims_blueprint": [],
        "sumber": {
            "label": label,
            "apa": apa,
            "doi": entry.get("doi") or "",
            "url": url_path,
            "journal": journal,
            "download_kind": "pdf",
        },
    }
    if not url_path or not fname:
        base["note"] = "PDF korpus tidak memiliki berkas yang tersaji."
        return base

    cache_key = f"{fname}|{str(var_name or '').strip().lower()}"
    cache = _load_corpus_cache()
    if cache_key in cache:
        try:
            base.update(cache[cache_key])
        except Exception:
            pass
        return base

    text, status = _pdf_text(url_path)
    if status != "ok" or not text or len(text.strip()) < 800:
        base["note"] = ("Teks PDF tidak terbaca / terlalu pendek untuk diekstrak"
                        if status == "ok" else status)
        return base

    try:
        from llm_title_parser import load_config, gemini_keys
        import variable_tracer as VT
        import variable_library as VL
    except Exception as e:
        base["note"] = f"Modul ekstraksi tidak tersedia: {e}"
        return base

    config = load_config()
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        base["note"] = ("API nonaktif — korpus belum bisa diekstrak. "
                        "Pilih Hasil Telusur Online atau nyalakan API.")
        return base

    dims_raw = VT._extract_dims_from_text(
        var_name, role, text, config,
        doi=entry.get("doi") or None)
    if not dims_raw or not dims_raw.get("dimensi"):
        base["note"] = ("Korpus ini tidak memuat dimensi/indikator eksplisit "
                        "untuk variabel tsb.")
        return base

    clean, report = VL.verify_indicators(dims_raw.get("dimensi") or [], text)
    final = VT._finalize_dims(clean)
    ok = VT._check_accept(final, report) or VT._check_pass_l2(final, report)
    n_dim, n_ind = _counts(final)
    base["dims"] = n_dim
    base["inds"] = n_ind
    base["passed"] = ok
    if not ok:
        base["note"] = (f"Korpus memuat {n_dim} dimensi / {n_ind} indikator "
                        "terverifikasi — di bawah lantai PLS-SEM (min. 2 dimensi "
                        "& 4 indikator).")
        cache[cache_key] = {"dims": n_dim, "inds": n_ind, "passed": False,
                            "note": base["note"]}
        _save_corpus_cache()
        return base

    onto = {
        "nama_resmi": entry.get("variable") or var_name,
        "input_name": var_name,
        "tokoh_ahli": label,
        "grand_theory": dims_raw.get("grand_theory") or "Manajemen Sumber Daya Manusia",
        "definisi": dims_raw.get("definisi") or "",
        "dimensi": final,
        "engine": "corpus-pdf",
        "jurnal_rujukan": {
            "label": label, "apa": apa, "year": year,
            "doi": entry.get("doi") or "", "url": url_path,
            "doc_kind": doc_kind,
        },
    }
    base["onto"] = onto
    base["note"] = "Dimensi & indikator diverifikasi literal dari PDF korpus."
    cache[cache_key] = {k: base[k] for k in
                        ("dims", "inds", "passed", "note", "onto")}
    _save_corpus_cache()
    return base


def corpus_to_dims_blueprint(onto):
    """Ubah onto-like (dimensi: [{nama, indikator}]) -> bentuk dimensi blueprint S1."""
    dims = (onto or {}).get("dimensi") or []
    out = []
    for d in dims[:3]:
        nama = d.get("nama") or "Dimensi"
        inds = [str(i) for i in (d.get("indikator") or []) if str(i)][:3]
        out.append({
            "nama": nama,
            "desc": f"{nama}: Mengukur efektivitas dan keteraturan pelaksanaan {nama.lower()}.",
            "indikator": inds,
            "jumlah_butir": len(inds),
        })
    return out
