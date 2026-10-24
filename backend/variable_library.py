# backend/variable_library.py
"""
PUSTAKA VARIABEL PERMANEN (Tahap A).

Penyimpanan dimensi & indikator yang bersumber dari SATU jurnal/buku nyata
ber-DOI (atau berkas unggahan pengguna), dipakai ulang lintas judul tesis.

Prinsip:
  1. Satu variabel  ->  SATU sumber aktif (keputusan sadar pengguna bila mengganti).
  2. Setiap indikator menyimpan KUTIPAN ASLI dari sumber + bukti kecocokan teks.
  3. Tanpa sumber yang lolos verifikasi -> variabel ditandai needs_review
     (TIDAK diisi karangan model).
  4. File per variabel di backend/variable_library/<slug>.json -> permanen,
     tidak terpengaruh saklar use_ontology_cache.

Struktur entri (dict):
{
  "variable": "Kinerja Pegawai",
  "synonyms": ["kinerja pegawai", "employee performance"],
  "source": {
    "type": "openalex" | "upload",
    "label": "Ulya & Rahmah (2025)",
    "apa": "Ulya, N., & Rahmah, S. (2025). .... DOI:10.xxxx",
    "doi": "https://doi.org/10.xxxx",
    "url": "https://...",
    "year": 2025,
    "doc_kind": "artikel" | "buku" | "unggahan pengguna",
    "file": "nama_berkas.pdf"        # untuk unggahan
  },
  "dimensi": [
    {"nama": "...", "indikator": ["...", "..."]}
  ],
  "verified_at": "2026-10-14T...",
  "engine": "gemini:gemini-3.1-flash-lite"
}

Aturan verifikasi (disetujui pengguna):
  - SETIAP indikator wajib punya kutipan yang cocok di teks sumber
    (toleransi normalisasi tanda baca). Yang gagal -> Dibuang.
  - Sumber 2022-2026 untuk hasil pencarian; unggahan pengguna bebas tahun.
"""

import datetime
import json
import os
import re
import unicodedata

BACKEND_DIR = os.path.dirname(os.path.abspath(__file__))
LIBRARY_DIR = os.path.join(BACKEND_DIR, "variable_library")
UPLOAD_DIR = os.path.join(BACKEND_DIR, "uploaded_sources")

# Batas tahun pencarian live (disetujui: 2022-2026)
SEARCH_YEAR_MIN = 2022
SEARCH_YEAR_MAX = 2026

# Padanan istilah yang lazim dalam judul tesis Indonesia/Inggris. Ini HANYA
# menyamakan NAMA variabel saat mencari entri pustaka; tidak menambah
# indikator dan tidak mengubah isi sumber.
_CANON = {
    "karyawan": "pegawai", "karyawannya": "pegawai", "karyawan2": "pegawai",
    "employee": "pegawai", "employees": "pegawai", "staff": "pegawai",
    "karier": "karir", "career": "karir",
    "pelatihan": "training", "diklat": "training",
    "performance": "kinerja",
    "competency": "kompetensi", "competence": "kompetensi",
    "development": "pengembangan",
}
_STOP = {"dan", "di", "pada", "the", "of", "and", "a", "an", "pt", "persero",
         "cabang", "tbk"}


def _slug(name):
    """Nama variabel -> nama berkas aman (huruf kecil, spasi jadi _)."""
    s = unicodedata.normalize("NFKD", str(name or "")).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-zA-Z0-9]+", "_", s).strip("_").lower()
    return s or "unknown_variable"


def _tokens(name):
    """Token kanonik untuk pencocokan nama variabel (bukan untuk indikator)."""
    s = unicodedata.normalize("NFKD", str(name or "")).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-zA-Z0-9]+", " ", s).lower()
    out = set()
    for t in s.split():
        t = _CANON.get(t, t)
        if t not in _STOP and len(t) > 1:
            out.add(t)
    return out


def _name_score(target, candidate):
    a, b = _tokens(target), _tokens(candidate)
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def library_path(variable):
    return os.path.join(LIBRARY_DIR, f"{_slug(variable)}.json")


def _read_entry(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            entry = json.load(f)
    except Exception:
        return None
    if not isinstance(entry, dict) or not entry.get("dimensi"):
        return None
    if not isinstance(entry.get("source"), dict) or not entry["source"].get("label"):
        return None
    return entry


def load_entry(variable):
    """Muat entri variabel; cocokkan sinonim bila nama berkas persis tidak ada.

    Contoh nyata: judul memakai "Kinerja Karyawan", pustaka menyimpan
    "Kinerja Pegawai". Tanpa pencocokan ini variabel dianggap belum diriset
    lalu sistem mencari ulang ke API (lambat/rapuh) padahal entri sudah ada.
    """
    path = library_path(variable)
    if os.path.exists(path):
        return _read_entry(path)

    target = str(variable or "").strip()
    if not target or not os.path.isdir(LIBRARY_DIR):
        return None
    best, best_score = None, 0.0
    for fn in os.listdir(LIBRARY_DIR):
        if not fn.endswith(".json"):
            continue
        entry = _read_entry(os.path.join(LIBRARY_DIR, fn))
        if not entry:
            continue
        names = [entry.get("variable", "")] + list(entry.get("synonyms") or [])
        score = max((_name_score(target, n) for n in names), default=0.0)
        if score > best_score:
            best, best_score = entry, score
    # Ambang 0.5: "kinerja karyawan" cocok dengan "Kinerja" (0.5) dan
    # "Kinerja Pegawai" (1.0 setelah kanonik karyawan=pegawai).
    return best if best_score >= 0.5 else None


def save_entry(entry):
    """Simpan entri ke pustaka permanen. Mengembalikan path berkas."""
    if not entry or not entry.get("variable") or not entry.get("dimensi"):
        raise ValueError("Entri pustaka tidak lengkap (variable/dimensi wajib)")
    os.makedirs(LIBRARY_DIR, exist_ok=True)
    entry = dict(entry)
    entry["saved_at"] = datetime.datetime.now().isoformat(timespec="seconds")
    path = library_path(entry["variable"])
    with open(path, "w", encoding="utf-8") as f:
        json.dump(entry, f, indent=2, ensure_ascii=False)
    return path


def delete_entry(variable):
    path = library_path(variable)
    if os.path.exists(path):
        os.remove(path)
        return True
    return False


def list_library():
    """Daftar ringkas semua variabel tersimpan: [{variable, source, dims, indicators}]."""
    if not os.path.isdir(LIBRARY_DIR):
        return []
    out = []
    for fn in sorted(os.listdir(LIBRARY_DIR)):
        if not fn.endswith(".json"):
            continue
        try:
            with open(os.path.join(LIBRARY_DIR, fn), "r", encoding="utf-8") as f:
                e = json.load(f)
            n_ind = sum(len(d.get("indikator", [])) for d in e.get("dimensi", []))
            out.append({
                "variable": e.get("variable", fn),
                "source_label": (e.get("source") or {}).get("label", ""),
                "source_year": (e.get("source") or {}).get("year"),
                "source_type": (e.get("source") or {}).get("type", ""),
                "dims": len(e.get("dimensi", [])),
                "indicators": n_ind,
                "verified_at": e.get("verified_at", e.get("saved_at", "")),
            })
        except Exception:
            continue
    return out


def indicator_count(entry):
    if not entry:
        return 0
    return sum(len(d.get("indikator", [])) for d in entry.get("dimensi", []))


def verify_indicators(dimensi, source_text):
    """
    Verifikasi KUTIPIAN tiap indikator terhadap teks sumber (Tahap A, aturan 1:
    SETIAP indikator wajib cocok). Indikator tanpa bukti dibuang.

    Mengembalikan (dimensi_bersih, laporan) di mana laporan =
      {"checked": n, "passed": n, "dropped": [indikator...]}
    """
    from doc_text_extract import quote_in_source
    clean = []
    dropped = []
    checked = passed = 0
    for d in dimensi or []:
        nama = str(d.get("nama", "")).strip()
        kept = []
        for ind in d.get("indikator", []) or []:
            text = ind if isinstance(ind, str) else str(ind)
            text = text.strip()
            if not text:
                continue
            checked += 1
            if quote_in_source(text, source_text):
                kept.append(text)
                passed += 1
            else:
                dropped.append(text)
        if nama and kept:
            clean.append({"nama": nama, "indikator": kept})
    return clean, {"checked": checked, "passed": passed, "dropped": dropped}


def _flat_promotable(dims):
    """True bila 1 dimensi berisi 2-7 butir pendek yang layak jadi dimensi.

    Contoh: Training Dessler (Reaction/Learning/Behavior/Results). Butir
    panjang/definitif tidak boleh dipromosikan menjadi dimensi.
    """
    if len(dims) != 1:
        return False
    inds = [str(x).strip() for x in (dims[0].get("indikator") or []) if str(x).strip()]
    if not (2 <= len(inds) <= 7):
        return False
    import re as _re
    for s in inds:
        if len(s.split()) > 6 or len(s) > 70:
            return False
        if _re.search(r"\b(adalah|merupakan|mencakup|mengacu|didefinisikan|"
                      r"diartikan|meliputi|terdiri\s+dari)\b", s, _re.I):
            return False
    return True


def entry_needs_review(entry):
    """True bila entri tidak layak dipakai otomatis (kosong / tanpa sumber).

    Kebijakan generated_v1: 2-7 dimensi, atau 1 dimensi datar yang butirnya
    layak dipromosikan menjadi dimensi. Indikator boleh belum ada karena akan
    dikembangkan dari dimensi pada tahap sanitasi.
    """
    if not entry:
        return True
    dims = entry.get("dimensi") or []
    if not dims:
        return True
    if len(dims) == 1:
        if not _flat_promotable(dims):
            return True
    elif len(dims) < 2:
        return True
    # Sumber dengan 8-20 dimensi TIDAK ditolak. Tahap sanitasi memilih 2-7
    # dimensi yang paling relevan dan mencatat sisanya sebagai tidak dipakai.
    src = entry.get("source") or {}
    if not src.get("label"):
        return True
    # Hasil pencarian live wajib berada di rentang 2022-2026
    if src.get("type") == "openalex":
        year = src.get("year") or 0
        if not (SEARCH_YEAR_MIN <= year <= SEARCH_YEAR_MAX):
            return True
    return False


def to_ontology_stub(entry, role, input_name):
    """
    Ubah entri pustaka menjadi dict sestruktur riset ontologi sehingga bisa
    langsung dipakai blueprint (kisi-kisi, kuesioner, fenomena Bab I).
    Pernyataan kuesioner dibuat terpisah oleh agen kuesioner bila belum ada.
    """
    try:
        import indicator_layers as IL
    except Exception:
        IL = None
    dims = []
    for d in entry.get("dimensi", []):
        inds = d.get("indikator", [])
        op = list(d.get("indikator_operasional") or inds)
        teo = list(d.get("indikator_teoretis") or [])
        if IL and (not teo or len(teo) < 2):
            teo = IL.bangun_indikator_teoretis(op, d.get("nama", ""))
        dims.append({
            "nama": d.get("nama", ""),
            "indikator": op,
            "indikator_operasional": op,
            "indikator_teoretis": teo,
            "pernyataan": list(d.get("pernyataan", []) or [""] * len(op)),
        })
    src = entry.get("source") or {}
    return {
        "role": role,
        "input_name": input_name or entry.get("variable", ""),
        "nama_resmi": input_name or entry.get("variable", ""),
        "grand_theory": entry.get("grand_theory", ""),
        "tokoh_ahli": entry.get("tokoh_ahli", ""),
        "definisi": entry.get("definisi", ""),
        "dimensi": dims,
        "jurnal_rujukan": {"sitasi": src.get("label", ""), "entri": src.get("apa", "")},
        "skala_pengukuran": "Skala Likert 1-5 (1: Sangat Tidak Setuju s/d 5: Sangat Setuju)",
        "engine": f"library:{src.get('type', 'unknown')}:{src.get('year', '')}",
        "needs_review": False,
        "source_note": f"Sumber: {src.get('label', '')}",
    }
