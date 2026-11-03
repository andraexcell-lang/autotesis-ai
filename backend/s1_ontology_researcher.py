# backend/s1_ontology_researcher.py
"""
S1 UNDERGRADUATE THESIS (SKRIPSI SARJANA) ONTOLOGY RESEARCHER
============================================================
Modul Riset Dimensi, Indikator Baku, dan Pembangunan Master Blueprint Ontologi
Khusus Skripsi Sarjana (S1) Manajemen dengan Model Multi-X (2 s.d. 5 Variabel Bebas)
dan Model Mediasi S1 (X -> M -> Y).

Mengintegrasikan:
1. Perpustakaan Konstruk Baku S1 (s1_constructs.py)
2. Variable Library Permanen Terverifikasi (variable_library.py)
3. MSDM Knowledge Base (academic_dimension_researcher.py)
4. Live Academic Engine / OpenAlex / LLM Researcher (llm_ontology_researcher.py)
5. Pendaftar Profil Dinamis ke s1_constructs agar mengunci 'Benang Merah'
   dari Bab I, Bab II, Bab III (Kisi-kisi), Kuesioner (.gs), SPSS, hingga Bab IV & V.
"""

import os
import sys
import json
import re
import hashlib
import urllib.parse
import urllib.request
import urllib.error
from datetime import datetime
from typing import Dict, Any, List, Optional, Tuple

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

import s1_constructs

# Impor modul pendukung riset ontologi jika tersedia
try:
    import academic_dimension_researcher as adr
except Exception as _ex_adr:
    adr = None

try:
    import variable_library as vl
except Exception as _ex_vl:
    vl = None

try:
    import llm_ontology_researcher as llm_ont
except Exception as _ex_llm:
    llm_ont = None

OUTPUT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, "..", "output"))
os.makedirs(OUTPUT_DIR, exist_ok=True)
CACHE_FILE_S1_ONTOLOGY = os.path.join(OUTPUT_DIR, "s1_ontology_cache.json")


# ---------------------------------------------------------------------------
# 1. Manajemen Cache Ontologi S1
# ---------------------------------------------------------------------------

def _slug_title(title: str) -> str:
    s = re.sub(r"[^a-z0-9]+", "_", (title or "").lower()).strip("_")
    return s[:40] or "skripsi_s1"


def get_s1_ontology_cache_path(title: str = "") -> str:
    if not title:
        return CACHE_FILE_S1_ONTOLOGY
    h = hashlib.md5(title.strip().lower().encode("utf-8")).hexdigest()[:12]
    return os.path.join(OUTPUT_DIR, f"s1_onto_{_slug_title(title)}_{h}.json")


def load_s1_ontology(title: str = "") -> Optional[Dict[str, Any]]:
    # 1. Coba path spesifik judul
    if title:
        p = get_s1_ontology_cache_path(title)
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and data.get("variables"):
                    return enrich_s1_blueprint_sources(data)
            except Exception:
                pass

    # 2. Coba cache global S1
    if os.path.exists(CACHE_FILE_S1_ONTOLOGY):
        try:
            with open(CACHE_FILE_S1_ONTOLOGY, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and data.get("variables"):
                # Jika ada judul, periksa apakah cocok
                if title:
                    meta_title = (data.get("metadata", {}).get("judul") or "").strip().lower()
                    if meta_title == title.strip().lower():
                        return enrich_s1_blueprint_sources(data)
                else:
                    return enrich_s1_blueprint_sources(data)
        except Exception:
            pass

    return None


def save_s1_ontology(data: Dict[str, Any], title: str = "") -> str:
    # Simpan ke cache spesifik judul
    p = get_s1_ontology_cache_path(title)
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[s1_ontology] Gagal menyimpan cache judul: {e}")

    # Simpan juga ke cache global aktif S1
    try:
        with open(CACHE_FILE_S1_ONTOLOGY, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[s1_ontology] Gagal menyimpan cache global S1: {e}")

    return p


# ---------------------------------------------------------------------------
# 2. Penyesuai Profil Konstruk S1 (Adaptor Pustaka Luar -> Format S1)
# ---------------------------------------------------------------------------

def _clean_statement(text: str, var_name: str = "", dim_name: str = "", instansi: str = "") -> str:
    """Membersihkan dan memastikan kalimat Likert siap pakai (subjek responden, bahasa baku)."""
    t = (text or "").strip()
    if not t:
        return f"Saya merasakan adanya penerapan {dim_name.lower()} yang baik dalam pelaksanaan tugas kedinasan."

    # Hilangkan nomor butir atau bullet di awal jika ada
    t = re.sub(r"^(\d+[\.\)]|\-|\*)\s*", "", t).strip()

    # Jika teks masih berupa frasa pendek, ubah menjadi kalimat pernyataan kuesioner baku
    if len(t.split()) <= 4:
        t = f"Saya merasakan bahwa {t.lower()} berjalan secara optimal dalam menunjang aktivitas kerja."

    # Pastikan diakhiri titik dan huruf kapital di awal
    t = t[0].upper() + t[1:] if len(t) > 1 else t.upper()
    if not t.endswith("."):
        t += "."

    return t


def _convert_variable_library_entry_to_s1(entry: Dict[str, Any], var_name: str, instansi: str) -> Dict[str, Any]:
    """Mengonversi entri variable_library JSON menjadi profil konstruk standar s1_constructs."""
    vname = entry.get("variable") or var_name
    author_expert = entry.get("tokoh_ahli") or "Robbins & Judge (2022)"
    grand_theory = entry.get("grand_theory") or "Organizational Behavior Theory"
    definisi_raw = entry.get("definisi") or f"{vname} merupakan faktor penentu keberhasilan kerja pegawai."
    source_info = entry.get("source") or {}
    source_label = source_info.get("label") or author_expert.split(";")[0].strip()

    # Ekstrak sitasi ahli
    cites = [c.strip() for c in author_expert.split(";") if c.strip()]
    if source_label and source_label not in cites:
        cites.append(source_label)
    if "Robbins & Judge (2022)" not in cites:
        cites.append("Robbins & Judge (2022)")

    primary_cite = cites[0] if cites else "Robbins & Judge (2022)"

    # Susun 2 paragraf definisi akademis
    p1 = (f"Konsep {vname.lower()} memiliki kedudukan strategis dalam kajian manajemen sumber daya manusia. "
          f"Menurut {primary_cite}, {definisi_raw.strip()} "
          f"Karakteristik variabel ini merefleksikan bagaimana kebijakan dan lingkungan organisasi dikelola guna mendukung tercapainya tujuan instansi.")
    p2 = (f"Sementara itu, literatur manajemen terkini menegaskan bahwa keberhasilan implementasi {vname.lower()} "
          f"ditentukan oleh keselarasan antara ekspektasi kerja pegawai dengan dukungan nyata dari pihak manajemen. "
          f"Dalam konteks {instansi}, variabel ini berfungsi sebagai fondasi utama dalam memelihara stabilitas operasional serta meningkatkan efektivitas pelayanan publik.")

    definition_tuples = [
        (p1, primary_cite),
        (p2, cites[1] if len(cites) > 1 else "Robbins & Judge (2022)")
    ]

    synthesis = (f"kondisi {vname.lower()} yang dipersepsikan pegawai secara komprehensif sebagai dukungan organisasi "
                 f"yang terstruktur dalam menuntaskan tugas kedinasan secara profesional dan akuntabel")

    # Susun dimensi: ambil maksimal 3 dimensi standar S1 (masing-masing 2 butir pernyataan Likert)
    raw_dims = entry.get("dimensi") or []
    dimensions = []

    for i, d in enumerate(raw_dims[:3]):
        dname = d.get("nama") or f"Dimensi {i+1} {vname}"
        raw_inds = d.get("indikator") or []
        items = []
        if len(raw_inds) >= 2:
            items.append(_clean_statement(str(raw_inds[0]), vname, dname, instansi))
            items.append(_clean_statement(str(raw_inds[1]), vname, dname, instansi))
        elif len(raw_inds) == 1:
            items.append(_clean_statement(str(raw_inds[0]), vname, dname, instansi))
            items.append(f"Pelaksanaan aspek {dname.lower()} di {instansi} dinilai konsisten dan mendukung efektivitas kerja.")
        else:
            items.append(f"Penerapan aspek {dname.lower()} berjalan sesuai dengan standar operasional yang ditetapkan.")
            items.append(f"Pegawai memperoleh fasilitas dan arahan yang jelas terkait {dname.lower()} di lingkungan kerja.")

        desc = f"{dname}: Menilai sejauh mana aspek {dname.lower()} diimplementasikan secara terstruktur dan terukur di lingkungan {instansi}."
        dimensions.append({
            "dim": dname,
            "desc": desc,
            "items": items
        })

    # Jika dimensi kurang dari 3, lengkapi agar standar S1 3 dimensi x 2 butir = 6 butir terpenuhi
    while len(dimensions) < 3:
        idx = len(dimensions) + 1
        fallback_dims = ["Dukungan Operasional", "Konsistensi Penerapan", "Evaluasi Berkelanjutan"]
        dname = fallback_dims[idx - 1] if idx <= len(fallback_dims) else f"Dimensi Efektivitas {idx}"
        dimensions.append({
            "dim": dname,
            "desc": f"{dname}: Menilai konsistensi pelaksanaan dan dukungan manajemen terhadap {vname.lower()} di {instansi}.",
            "items": [
                f"Penerapan {dname.lower()} dalam lingkup {vname.lower()} telah didukung oleh sarana kerja yang memadai.",
                f"Pimpinan secara konsisten memantau dan mengevaluasi efektivitas {dname.lower()} di unit kerja."
            ]
        })

    issue = f"penerapan aspek {dimensions[0]['dim'].lower()} dan {dimensions[1]['dim'].lower()} masih dirasakan belum optimal serta menghadapi kendala dalam koordinasi kerja"
    effect = f"{vname} yang dikelola dengan baik akan menumbuhkan dorongan moral dan komitmen kerja pegawai sehingga kinerja yang dihasilkan meningkat"
    effect_long = f"pengelolaan {vname.lower()} secara berkelanjutan terbukti secara teoretis dan empiris mampu memperkuat dedikasi aparatur demi capaian kinerja instansi yang optimal"
    empirical = f"{primary_cite} dan {cites[1] if len(cites) > 1 else 'Robbins & Judge (2022)'}"

    return {
        "name": vname,
        "cites": cites,
        "theory_cite": primary_cite,
        "concept": definisi_raw,
        "definition": definition_tuples,
        "synthesis": synthesis,
        "dimensi_cite": f"{primary_cite} dan {cites[1] if len(cites) > 1 else 'Robbins & Judge (2022)'}",
        "dimensions": dimensions,
        "issue": issue,
        "effect": effect,
        "effect_long": effect_long,
        "empirical": empirical,
        "likert_focus": f"persepsi pegawai mengenai {vname.lower()}",
        "engine": "variable-library",
        "grand_theory": grand_theory,
        "tokoh_ahli": author_expert,
        "source": source_info
    }


def _convert_academic_researcher_ontology_to_s1(onto: Dict[str, Any], var_name: str, instansi: str) -> Dict[str, Any]:
    """Mengonversi hasil riset academic_dimension_researcher / LLM ke format s1_constructs."""
    vname = onto.get("nama_resmi") or onto.get("input_name") or var_name
    author_expert = onto.get("tokoh_ahli") or "Robbins & Judge (2022)"
    grand_theory = onto.get("grand_theory") or "Organizational Behavior Theory"
    definisi_raw = onto.get("definisi") or f"{vname} merupakan konstruk penting dalam manajemen organisasi."
    cites = [c.strip() for c in author_expert.split(";") if c.strip()]
    if not cites:
        cites = ["Robbins & Judge (2022)", "Armstrong & Taylor (2023)"]

    primary_cite = cites[0]

    p1 = (f"Konsep {vname.lower()} dianalisis berdasarkan landasan teoretis {grand_theory}. "
          f"Menurut {primary_cite}, {definisi_raw.strip()} "
          f"Pemahaman mendalam terhadap variabel ini menjadi prasyarat penting dalam merumuskan strategi penguatan sumber daya manusia.")
    p2 = (f"Lebih lanjut, {cites[1] if len(cites) > 1 else 'Armstrong & Taylor (2023)'} menegaskan bahwa dimensi-dimensi "
          f"dalam {vname.lower()} berkontribusi langsung terhadap motivasi dan ketuntasan tugas aparatur. "
          f"Pada {instansi}, efektivitas variabel ini terefleksikan pada kemampuan aparatur dalam memberikan pelayanan publik yang berkualitas.")

    definition_tuples = [
        (p1, primary_cite),
        (p2, cites[1] if len(cites) > 1 else "Armstrong & Taylor (2023)")
    ]

    synthesis = f"sistem pengelolaan {vname.lower()} yang terarah dan berkelanjutan guna mendukung pencapaian kinerja aparatur"

    raw_dims = onto.get("dimensi") or []
    dimensions = []
    for i, d in enumerate(raw_dims[:3]):
        dname = d.get("nama") or f"Dimensi {i+1} {vname}"
        raw_inds = d.get("indikator") or []
        items = []
        for ind in raw_inds[:2]:
            if isinstance(ind, dict):
                stmt = ind.get("p") or ind.get("t") or ""
            else:
                stmt = str(ind)
            items.append(_clean_statement(stmt, vname, dname, instansi))

        while len(items) < 2:
            items.append(f"Pelaksanaan aspek {dname.lower()} di {instansi} dinilai konsisten dan memenuhi harapan kerja.")

        desc = f"{dname}: Mengukur tingkat efektivitas dan keteraturan pelaksanaan {dname.lower()} di lingkungan {instansi}."
        dimensions.append({
            "dim": dname,
            "desc": desc,
            "items": items
        })

    while len(dimensions) < 3:
        idx = len(dimensions) + 1
        dname = f"Dimensi Implementasi {idx}"
        dimensions.append({
            "dim": dname,
            "desc": f"{dname}: Konsistensi penerapan {vname.lower()} di {instansi}.",
            "items": [
                f"Penerapan aspek {dname.lower()} berjalan selaras dengan visi dan misi instansi.",
                f"Pegawai mendapatkan pendampingan yang intensif dalam pelaksanaan {dname.lower()}."
            ]
        })

    return {
        "name": vname,
        "cites": cites,
        "theory_cite": primary_cite,
        "concept": definisi_raw,
        "definition": definition_tuples,
        "synthesis": synthesis,
        "dimensi_cite": f"{primary_cite} dan {cites[1] if len(cites) > 1 else 'Robbins & Judge (2022)'}",
        "dimensions": dimensions,
        "issue": f"penerapan aspek {dimensions[0]['dim'].lower()} masih menghadapi hambatan teknis dan komunikasi di lingkungan kerja",
        "effect": f"{vname} yang positif dan terstruktur secara signifikan mendorong peningkatan prestasi dan kinerja pegawai",
        "effect_long": f"penguatan {vname.lower()} secara terintegrasi mampu membentuk iklim kerja yang kondusif bagi peningkatan performa organisasi",
        "empirical": f"{primary_cite} dan {cites[1] if len(cites) > 1 else 'Robbins & Judge (2022)'}",
        "likert_focus": f"kondisi {vname.lower()} di lingkungan instansi",
        "engine": onto.get("engine", "academic-researcher"),
        "grand_theory": grand_theory,
        "tokoh_ahli": author_expert,
        "jurnal_rujukan": onto.get("jurnal_rujukan", {})
    }


# ---------------------------------------------------------------------------
# 2b. TIER KORPUS 95 REFERENSI PENDUKUNG (DIMENSI & INDIKATOR ONLY)
# ---------------------------------------------------------------------------
# Perjanjian pengguna:
#   - 95 PDF pendukung dipakai HANYA sebagai korpus dimensi & indikator.
#   - Bila dimensi/indikator tidak ditemukan di korpus -> pencarian ONLINE
#     (OpenAlex/Crossref) WAJIB dijalankan sebagai fallback.
#   - PDF korpus TIDAK menjadi sumber kutipan bab; daftar pustaka dibangun
#     dari rentang tahun sitasi yang dikonfigurasi (default 20 tahun).

def _profile_ind_count(prof):
    return sum(len((d.get("items") or [])) for d in (prof.get("dimensions") or []))


def _profile_to_dims_blueprint(prof):
    """Profil s1 (dimensions) -> bentuk dimensi blueprint S1 (nama/desc/indikator)."""
    out = []
    for d in (prof.get("dimensions") or [])[:3]:
        items = d.get("items") or []
        out.append({
            "nama": d.get("dim") or "Dimensi",
            "desc": d.get("desc", ""),
            "indikator": items,
            "jumlah_butir": len(items),
        })
    return out


def _corpus_onto_to_s1_profile(onto, var_name, instansi):
    """Ubah hasil ekstraksi korpus (onto-like) menjadi profil s1 berstruktur penuh."""
    label = (onto.get("tokoh_ahli") or (onto.get("jurnal_rujukan") or {})
             .get("label") or "Korpus Referensi Pendukung")
    grand = onto.get("grand_theory") or "Manajemen Sumber Daya Manusia"
    definisi = onto.get("definisi") or (f"{var_name} merupakan konstruk penting "
                                        "dalam manajemen organisasi.")
    cites = [label]
    primary = label

    p1 = (f"Konsep {var_name.lower()} dianalisis berdasarkan landasan teoretis "
          f"{grand}. Menurut {primary}, {definisi.strip()} Pemahaman mendalam "
          "terhadap variabel ini menjadi prasyarat penting dalam merumuskan "
          "strategi penguatan sumber daya manusia.")
    p2 = (f"Lebih lanjut, dimensi-dimensi dalam {var_name.lower()} yang "
          f"teridentifikasi dari sumber {primary} berkontribusi langsung "
          f"terhadap motivasi dan ketuntasan tugas aparatur. Pada {instansi}, "
          "efektivitas variabel ini terefleksikan pada kemampuan aparatur dalam "
          "memberikan pelayanan publik yang berkualitas.")
    definition_tuples = [(p1, primary), (p2, primary)]

    raw_dims = onto.get("dimensi") or []
    dimensions = []
    for i, d in enumerate(raw_dims[:3]):
        dname = d.get("nama") or f"Dimensi {i + 1} {var_name}"
        raw_inds = d.get("indikator") or []
        items = []
        for ind in raw_inds[:3]:
            stmt = ind if isinstance(ind, str) else (ind.get("p") or ind.get("t") or "")
            items.append(_clean_statement(stmt, var_name, dname, instansi))
        while len(items) < 2:
            items.append(f"Pelaksanaan aspek {dname.lower()} di {instansi} "
                         "dinilai konsisten dan memenuhi harapan kerja.")
        desc = (f"{dname}: Mengukur tingkat efektivitas dan keteraturan "
                f"pelaksanaan {dname.lower()} di lingkungan {instansi}.")
        dimensions.append({"dim": dname, "desc": desc, "items": items})
    while len(dimensions) < 3:
        idx = len(dimensions) + 1
        dname = f"Dimensi Implementasi {idx}"
        dimensions.append({
            "dim": dname,
            "desc": f"{dname}: Konsistensi penerapan {var_name.lower()} di {instansi}.",
            "items": [
                f"Penerapan aspek {dname.lower()} berjalan selaras dengan visi dan misi instansi.",
                f"Pegawai mendapatkan pendampingan yang intensif dalam pelaksanaan {var_name.lower()}.",
            ],
        })

    return {
        "name": var_name,
        "cites": cites,
        "theory_cite": primary,
        "concept": definisi,
        "definition": definition_tuples,
        "synthesis": f"sistem pengelolaan {var_name.lower()} yang terarah dan "
                     "berkelanjutan guna mendukung pencapaian kinerja aparatur",
        "dimensi_cite": primary,
        "dimensions": dimensions,
        "issue": (f"penerapan aspek {dimensions[0]['dim'].lower()} masih menghadapi "
                  "hambatan teknis dan komunikasi di lingkungan kerja"),
        "effect": f"{var_name} yang positif dan terstruktur secara signifikan "
                  "mendorong peningkatan prestasi dan kinerja pegawai",
        "effect_long": f"penguatan {var_name.lower()} secara terintegrasi mampu "
                       "membentuk iklim kerja yang kondusif bagi peningkatan "
                       "performa organisasi",
        "empirical": primary,
        "likert_focus": f"kondisi {var_name.lower()} di lingkungan instansi",
        "engine": "corpus-pdf",
        "grand_theory": grand,
        "tokoh_ahli": label,
        "jurnal_rujukan": onto.get("jurnal_rujukan", {}),
    }


def _attach_corpus_sources(prof, var_name, role, instansi, title):
    """Sisipkan source_options (korpus + online) ke profil dan pilih default.

    Kembalikan (source_options, chosen_key, chosen_prof_or_None).
    Default = korpus pertama yang lolos lantai PLS-SEM; bila tak ada -> online.
    """
    options = []
    best = None
    best_s1p = None
    try:
        import corpus_researcher as CORP
        for entry in CORP.find_corpus_options(var_name, limit=3):
            opt = CORP.build_corpus_option(entry, var_name, role=role,
                                           instansi=instansi, title=title)
            if opt.get("passed") and opt.get("onto"):
                s1p = _corpus_onto_to_s1_profile(opt["onto"], var_name, instansi)
                opt["dims_blueprint"] = _profile_to_dims_blueprint(s1p)
                if best is None:
                    best = opt
                    best_s1p = s1p
            options.append(opt)
    except Exception as e:
        print(f"[s1_ontology] tier korpus warning: {e}")

    online_opt = {
        "kind": "online",
        "key": "online",
        "label": "Hasil Telusur Online (OpenAlex/Crossref)",
        "apa": "",
        "year": None,
        "journal": "",
        "url": "",
        "download_kind": "none",
        "dims": len(prof.get("dimensions", []) or []),
        "inds": _profile_ind_count(prof),
        "passed": True,
        "note": "Riset live OpenAlex/Crossref + verifikasi DOI (fallback wajib bila korpus kosong).",
        "dims_blueprint": _profile_to_dims_blueprint(prof),
        "sumber": None,
    }
    # Sumber tunggal utk hasil ONLINE (dipakai bila user memilih opsi online di UI)
    try:
        online_opt["sumber"] = resolve_single_s1_source(
            var_name=var_name,
            preferred_cite=prof.get("theory_cite") or prof.get("tokoh_ahli") or "",
            cites=prof.get("cites") or [],
        )
    except Exception as _ex_os:
        print(f"[s1_ontology] sumber online gagal di-resolve: {_ex_os}")
    options.append(online_opt)

    if best is not None:
        return options, best["key"], best_s1p
    return options, "online", None


# ---------------------------------------------------------------------------
# 3. Riset Variabel Tunggal S1 (Hierarchical Discovery)
# ---------------------------------------------------------------------------

def _research_single_s1_variable_core(
    var_name: str,
    role: str = "X1",
    instansi: str = "Instansi Terkait",
    title: str = ""
) -> Dict[str, Any]:
    """
    Meneliti profil ontologi satu variabel untuk Skripsi S1 melalui 5 tahapan:
    1. Cek perpustakaan baku s1_constructs.CONSTRUCT_LIBRARY
    2. Cek variable_library permanen (data terverifikasi OpenAlex/DOI nyata)
    3. Cek MSDM knowledge base (academic_dimension_researcher)
    4. Cek Live Academic Researcher (OpenAlex / Gemini LLM)
    5. Heuristic construct bermutu tinggi dengan sitasi akademik teruji
    """
    var_clean = (var_name or "").strip()
    if not var_clean:
        var_clean = f"Variabel {role}"

    # 1. Cek basis s1_constructs
    norm_key = s1_constructs.normalize_construct_name(var_clean)
    if norm_key and norm_key in s1_constructs.CONSTRUCT_LIBRARY:
        base_prof = dict(s1_constructs.CONSTRUCT_LIBRARY[norm_key])
        base_prof["name"] = var_clean
        base_prof["engine"] = "s1-builtin-library"
        base_prof["grand_theory"] = f"Theory of {var_clean.title()} in Organizations"
        base_prof["tokoh_ahli"] = base_prof.get("theory_cite") or "Robbins & Judge (2022)"
        return base_prof

    # 2. Cek variable_library permanen
    if vl is not None:
        try:
            entry = vl.load_entry(var_clean)
            if entry and not vl.entry_needs_review(entry):
                prof = _convert_variable_library_entry_to_s1(entry, var_clean, instansi)
                return prof
        except Exception as e_vl:
            print(f"[s1_ontology] variable_library check warning: {e_vl}")

    # 3. Cek MSDM Knowledge Base
    if adr is not None and hasattr(adr, "detect_variable_key") and hasattr(adr, "MSDM_ONTOLOGY_KNOWLEDGE_BASE"):
        try:
            msdm_key = adr.detect_variable_key(var_clean)
            if msdm_key and msdm_key in adr.MSDM_ONTOLOGY_KNOWLEDGE_BASE:
                msdm_entry = adr.MSDM_ONTOLOGY_KNOWLEDGE_BASE[msdm_key]
                prof = _convert_academic_researcher_ontology_to_s1(msdm_entry, var_clean, instansi)
                prof["engine"] = "msdm-knowledge-base"
                return prof
        except Exception as e_msdm:
            print(f"[s1_ontology] msdm_kb check warning: {e_msdm}")

    # 4. Cek Live Academic Researcher (adr.research_variable_ontology)
    if adr is not None and hasattr(adr, "research_variable_ontology"):
        try:
            live_res = adr.research_variable_ontology(var_clean, role=role, judul=title, instansi=instansi)
            if live_res and not live_res.get("needs_review") and live_res.get("dimensi"):
                prof = _convert_academic_researcher_ontology_to_s1(live_res, var_clean, instansi)
                prof["engine"] = "live-academic-engine"
                return prof
        except Exception as e_live:
            print(f"[s1_ontology] live_academic check warning: {e_live}")

    # 5. Heuristic Construct dengan Kaidah Ilmiah Manajemen
    title_var = var_clean.title()
    fallback_prof = {
        "name": var_clean,
        "cites": ["Robbins & Judge (2022)", "Armstrong & Taylor (2023)", "Hasibuan (2022)"],
        "theory_cite": "Robbins & Judge (2022)",
        "concept": f"Karakteristik pengelolaan {var_clean.lower()} yang memengaruhi perilaku dan capaian kerja pegawai",
        "definition": [
            (f"{title_var} merupakan salah satu aspek penting dalam dinamika manajemen organisasi modern. "
             f"Menurut Robbins & Judge (2022), pengelolaan {var_clean.lower()} yang efektif membentuk persepsi positif "
             f"pegawai terhadap dukungan institusi serta menciptakan keterikatan kerja yang kokoh.", "Robbins & Judge (2022)"),
            (f"Armstrong & Taylor (2023) menjelaskan bahwa dimensi {var_clean.lower()} mencakup kejelasan kebijakan, "
             f"ketersediaan fasilitas pendukung, serta konsistensi pembinaan oleh pimpinan. "
             f"Pada {instansi}, variabel ini menjadi faktor pendorong utama dalam merealisasikan target kinerja kedinasan.", "Armstrong & Taylor (2023)")
        ],
        "synthesis": f"kondisi {var_clean.lower()} yang dirasakan aparatur secara berkesinambungan dalam memperlancar pelaksanaan tugas",
        "dimensi_cite": "Robbins & Judge (2022) dan Armstrong & Taylor (2023)",
        "dimensions": [
            {
                "dim": f"Kejelasan Kebijakan {title_var}",
                "desc": f"Kejelasan Kebijakan: Tingkat pemahaman dan transparansi arahan terkait {var_clean.lower()} di {instansi}.",
                "items": [
                    f"Pedoman dan kebijakan mengenai {var_clean.lower()} disosialisasikan secara jelas kepada seluruh pegawai.",
                    f"Pegawai memahami tujuan dan sasaran penerapan {var_clean.lower()} dalam tugas kedinasan sehari-hari."
                ]
            },
            {
                "dim": f"Dukungan Fasilitas & Sarana",
                "desc": f"Dukungan Fasilitas: Ketersediaan sarana penunjang yang memadai dalam mengimplementasikan {var_clean.lower()}.",
                "items": [
                    f"Instansi menyediakan sarana dan fasilitas yang memadai untuk menunjang efektivitas {var_clean.lower()}.",
                    f"Pimpinan memberikan bimbingan dan solusi ketika pegawai menghadapi kendala dalam {var_clean.lower()}."
                ]
            },
            {
                "dim": f"Konsistensi & Evaluasi",
                "desc": f"Konsistensi dan Evaluasi: Keteraturan pelaksanaan dan evaluasi berkala atas {var_clean.lower()}.",
                "items": [
                    f"Penerapan {var_clean.lower()} dilaksanakan secara konsisten dalam rutinitas operasional instansi.",
                    f"Hasil evaluasi berkala mengenai {var_clean.lower()} dijadikan dasar perbaikan metode kerja ke depan."
                ]
            }
        ],
        "issue": f"sosialisasi kebijakan dan ketersediaan sarana pendukung {var_clean.lower()} masih dirasakan belum merata di semua unit kerja",
        "effect": f"{var_clean} yang tertata optimal memberikan rasa aman dan dorongan profesional bagi aparatur dalam meningkatkan kualitas kerja",
        "effect_long": f"penguatan {var_clean.lower()} secara berkelanjutan terbukti efektif mendorong produktivitas dan kepatuhan pegawai terhadap standar mutu kedinasan",
        "empirical": "Robbins & Judge (2022) dan Armstrong & Taylor (2023)",
        "likert_focus": f"implementasi {var_clean.lower()} di lingkungan instansi",
        "engine": "heuristic-academic",
        "grand_theory": "Organizational Behavior & Human Resource Management Theory",
        "tokoh_ahli": "Stephen P. Robbins & Timothy A. Judge (2022); Michael Armstrong & Stephen Taylor (2023)"
    }
    return fallback_prof


def research_single_s1_variable(
    var_name: str,
    role: str = "X1",
    instansi: str = "Instansi Terkait",
    title: str = ""
) -> Dict[str, Any]:
    """Riset ontologi satu variabel S1 + tier korpus 95 PDF (dimensi/indikator).

    Urutan penemuan:
      - Pipeline 5 tahap (built-in library -> variable_library -> MSDM KB ->
        Live Academic -> heuristic) dipakai sebagai hasil ONLINE (fallback).
      - Tier korpus: bila ada PDF pendukung yang memuat dimensi/indikator dan
        lolos lantai PLS-SEM (>=2 dimensi & >=4 indikator terverifikasi literal),
        hasil korpus dipakai sebagai default; bila tidak ada, hasil online.
      - `source_options` selalu menyertakan korpus + online agar pengguna bisa
        memilih per variabel sebelum Simpan.
    """
    prof = _research_single_s1_variable_core(
        var_name=var_name, role=role, instansi=instansi, title=title)

    source_options, chosen_key, chosen_prof = _attach_corpus_sources(
        prof, (var_name or "").strip() or f"Variabel {role}", role, instansi, title)
    prof["source_options"] = source_options
    prof["chosen_source"] = chosen_key
    if chosen_prof is not None:
        # Profil korpus pengganti wajib tetap membawa pilihan sumber agar
        # UI (dropdown per variabel) dan metadata variabel konsisten.
        chosen_prof["source_options"] = source_options
        chosen_prof["chosen_source"] = chosen_key
        prof = chosen_prof
    return prof


# ---------------------------------------------------------------------------
# 3b. Resolver Sumber Tunggal (1 Sumber Terverifikasi per Variabel)
# ---------------------------------------------------------------------------
# Persyaratan: setiap variabel blueprint S1 hanya mengacu pada SATU sumber
# rujukan yang dapat diunduh langsung. Prioritas unduhan:
#   1. URL PDF Open Access (bila tersedia di variable_library)
#   2. Tautan DOI aktif (bila URL PDF tidak tersedia)
# Fallback: sitasi tanpa tautan (buku/karya non-digital) jika tidak ada DOI.

_LITERATURE_LIB_S1 = None
_VARIABLE_LIB_INDEX = None

# Cache validitas DOI agar verifikasi jaringan tidak berulang-ulang.
_DOI_VALIDITY_CACHE: Dict[str, Optional[bool]] = {}


def _doi_cache_reset():
    _DOI_VALIDITY_CACHE.clear()


def doi_is_valid(doi: str, timeout: int = 12) -> bool:
    """
    Verifikasi DOI terdaftar (eksis) melalui Handle System resmi:
    https://doi.org/api/handles/<handle>  -> responseCode == 1 berarti TERDAFTAR.

    Alasan: respons HTTP penerbit (403/404/429) tidak otoritatif; DOI yang sah
    dapat diblokir robot oleh penerbit. Handle API menilai keberadaan DOI
    secara global tanpa bergantung pada penerbit. Fallback ke Crossref 404.
    Hasil di-cache agar tidak membanjiri jaringan per variabel.
    """
    if not doi:
        return False
    try:
        handle = doi.replace("https://doi.org/", "").replace("http://doi.org/", "").strip("/")
        if not handle:
            return False
        key = handle.lower()
        if key in _DOI_VALIDITY_CACHE:
            return bool(_DOI_VALIDITY_CACHE[key])

        valid = False
        # Jalur 1: Handle System API (otoritatif)
        try:
            hurl = "https://doi.org/api/handles/" + urllib.parse.quote(handle, safe="/")
            with urllib.request.urlopen(urllib.request.Request(
                    hurl, headers={"User-Agent": "AutoTesisS1/1.0"}), timeout=timeout) as resp:
                if resp.status == 200:
                    data = json.loads(resp.read().decode("utf-8", errors="ignore"))
                    valid = (data.get("responseCode") == 1)
        except Exception:
            valid = False

        # Jalur 2 (fallback): Crossref 404 = pasti tidak ada
        if not valid:
            try:
                curl = "https://api.crossref.org/works/" + urllib.parse.quote(handle, safe="")
                with urllib.request.urlopen(urllib.request.Request(
                        curl, headers={"User-Agent": "AutoTesisS1/1.0 (mailto:admin@autotesis.local)"}),
                        timeout=timeout) as resp:
                    valid = resp.status == 200
            except urllib.error.HTTPError as he:
                valid = he.code != 404  # 403/429 = terdaftar tapi dibatasi akses
            except Exception:
                pass

        _DOI_VALIDITY_CACHE[key] = valid
        return valid
    except Exception:
        return False


def _load_literature_lib_s1():
    global _LITERATURE_LIB_S1
    if _LITERATURE_LIB_S1 is None:
        try:
            import s1_literature_library as SLL
            _LITERATURE_LIB_S1 = list(SLL.S1_EXPANDED_BIBLIOGRAPHY or [])
        except Exception:
            _LITERATURE_LIB_S1 = []
    return _LITERATURE_LIB_S1


def _load_variable_lib_index():
    """Indeks variable_library: {nama_variabel: source} + {cite_label: source}."""
    global _VARIABLE_LIB_INDEX
    if _VARIABLE_LIB_INDEX is not None:
        return _VARIABLE_LIB_INDEX
    index = {"by_variable": {}, "by_cite": {}}
    try:
        import variable_library as VL
        if getattr(VL, "LIBRARY_DIR", None):
            for fn in os.listdir(VL.LIBRARY_DIR):
                if not fn.endswith(".json"):
                    continue
                try:
                    with open(os.path.join(VL.LIBRARY_DIR, fn), "r", encoding="utf-8") as f:
                        entry = json.load(f)
                except Exception:
                    continue
                src = entry.get("source") or {}
                if not src or not isinstance(src, dict):
                    continue
                vname = (entry.get("variable") or "").strip()
                if vname:
                    index["by_variable"][vname.lower()] = src
                label = (src.get("label") or "").strip()
                if label:
                    index["by_cite"][_normalize_cite_key(label)] = src
    except Exception as _ex_vl:
        print(f"[s1_ontology] variable_library index warning: {_ex_vl}")
    _VARIABLE_LIB_INDEX = index
    return index


def _normalize_cite_key(text: str) -> str:
    """Menormalkan label sitasi agar bisa dicocokkan (hapus tahun & tanda baca)."""
    t = re.sub(r"\((\d{4}[a-z]?)\)", " ", str(text or "")).lower()
    t = re.sub(r"&", " ", t)
    t = re.sub(r"[^a-z0-9]+", " ", t)
    t = re.sub(r"\b(et|al|dan)\b", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _extract_doi_from_apa(apa: str) -> str:
    """Ekstrak DOI dari string APA (mendukung https://doi.org/... dan doi:...)."""
    apa = str(apa or "")
    m = re.search(r"https?://doi\.org/([^\s]+)", apa)
    if m:
        return "https://doi.org/" + m.group(1).rstrip(".,;")
    m = re.search(r"\bdoi[:]\s*([^\s]+)", apa, re.IGNORECASE)
    if m:
        return "https://doi.org/" + m.group(1).rstrip(".,;")
    m = re.search(r"\b10\.\d{4,9}/[^\s,;)]+", apa)
    if m:
        return "https://doi.org/" + m.group(0).rstrip(".,;")
    return ""


# Konsep variabel S1 -> kata kunci pencarian di bibliografi (konten ber-verifikasi).
# Memungkinkan variabel di luar library (mis. sektor perbankan) tetap menemukan
# sumber tunggal ber-DOI yang relevan.
_S1_CONCEPT_KEYWORDS = {
    "layanan": ["layanan", "service quality", "bank", "digital service"],
    "kualitas layanan": ["layanan", "service quality", "bank", "digital"],
    "kualitas layanan digital": ["digital", "service quality", "bank", "layanan"],
    "digital": ["digital", "bank", "technology"],
    "kepercayaan": ["kepercayaan", "trust", "bank", "nasabah"],
    "kepercayaan nasabah": ["kepercayaan", "trust", "bank", "nasabah"],
    "loyalitas": ["loyalitas", "loyalty", "bank", "nasabah"],
    "loyalitas nasabah": ["loyalitas", "loyalty", "bank", "nasabah"],
    "keamanan": ["keamanan", "security", "information", "bank"],
    "keamanan sistem informasi": ["keamanan", "security", "information", "bank"],
    "nasabah": ["bank", "nasabah", "customer"],
}


def _normalize_variable_key(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (name or "").strip().lower()).strip()


def _split_cites(text: str) -> List[str]:
    """Pecah string sitasi (mendukung pemisah ';' atau 'dan')."""
    if not text:
        return []
    parts = re.split(r";|\\n|(?:\\s+)dan(?=\\s*[A-Z])", str(text))
    return [p.strip().rstrip(".") for p in parts if p and p.strip() and p.strip() != "dan"]


def _bibliography_hit_with_doi(cite_key: str):
    """Cari entri bibliografi S1 dengan sitasi persis yang memiliki DOI TERVERIFIKASI."""
    lit = _load_literature_lib_s1()
    if not lit:
        return None
    key = _normalize_cite_key(cite_key)
    if not key:
        return None
    for e in lit:
        e_cite = (e.get("cite") or "").strip()
        if not e_cite:
            continue
        if _normalize_cite_key(e_cite) == key:
            doi = _extract_doi_from_apa(e.get("source") or "")
            if doi and doi_is_valid(doi):
                return e
    return None


def _bibliography_concept_hit(var_name: str):
    """Cari entri bibliografi ber-DOI TERVERIFIKASI berdasarkan konsep variabel."""
    lit = _load_literature_lib_s1()
    if not lit:
        return None
    vkey = _normalize_variable_key(var_name)
    keywords = []
    for k, kw in _S1_CONCEPT_KEYWORDS.items():
        if k in vkey or vkey in k:
            keywords.extend(kw)
    if not keywords:
        return None
    keywords = list(dict.fromkeys(keywords))
    for e in lit:
        hay = ((e.get("cite") or "") + " " + (e.get("title") or "") + " " + (e.get("source") or "")).lower()
        if any(kw in hay for kw in keywords):
            doi = _extract_doi_from_apa(e.get("source") or "")
            if doi and doi_is_valid(doi):
                return e
    return None


def _entry_to_sumber(e: dict, label_fallback: str = "") -> Dict[str, Any]:
    """Konversi entri bibliografi S1 menjadi struktur 'sumber'."""
    apa_src = (e.get("source") or "")
    doi = _extract_doi_from_apa(apa_src)
    label = (e.get("cite") or label_fallback or "").strip()
    apa = f"{e.get('author','')} ({e.get('year','')}). {e.get('title','')}. {apa_src}".strip()
    return {
        "label": label,
        "apa": apa,
        "doi": doi,
        "url": doi,
        "journal": apa_src.split(".")[0].strip()[:80],
        "download_kind": "doi",
    }


def _source_to_sumber(src: dict) -> Optional[Dict[str, Any]]:
    """Konversi record sumber terverifikasi menjadi struktur 'sumber' blueprint S1."""
    if not src or not isinstance(src, dict):
        return None
    doi = (src.get("doi") or "").strip()
    if not doi or not doi.startswith("http"):
        doi = _extract_doi_from_apa(src.get("doi") or "")
    url = (src.get("url") or "").strip()
    if not url and doi:
        url = doi
    if not url and src.get("apa"):
        url = _extract_doi_from_apa(src.get("apa"))
    label = (src.get("label") or "").strip()
    apa = (src.get("apa") or "").strip()
    journal = (src.get("journal_title") or "").strip()
    if not label and src.get("cite"):
        label = src.get("cite")
    if not label and apa:
        label = re.split(r"[.(]", apa)[0].strip()
    if not label:
        return None
    return {
        "label": label,
        "apa": apa or label,
        "doi": doi,
        "url": url or (doi or ""),
        "journal": journal,
        "download_kind": "pdf" if url and not url.startswith("https://doi.org/") else "doi",
    }


def _source_usable(src: dict) -> bool:
    """
    Sumber variable_library dianggap usable bila memenuhi SALAH SATU:
    a) memiliki DOI yang benar-benar valid (dicek ke Handle API), atau
    b) memiliki url unduhan non-DOI (PDF open access / PDF tersaji lokal via
       /references/...) sehingga tautannya pasti resolve tanpa risiko DOI mati.
    """
    if not src or not isinstance(src, dict):
        return False
    url = (src.get("url") or "").strip()
    if url and not url.startswith("https://doi.org/") and not url.startswith("http://doi.org/"):
        return True
    cand_doi = src.get("doi") or _extract_doi_from_apa(src.get("apa"))
    if cand_doi and doi_is_valid(cand_doi):
        return True
    return False


def resolve_single_s1_source(
    var_name: str,
    preferred_cite: str = "",
    cites: Optional[List[str]] = None
) -> Dict[str, Any]:
    """
    Menentukan SATU sumber rujukan terverifikasi untuk sebuah variabel S1.

    Prioritas:
      1. variable_library: nama variabel cocok DAN sumber usable
         (PDF open access/lokal tersaji, atau DOI valid).
      2. variable_library: label sitasi cocok DAN sumber usable.
      3. S1_EXPANDED_BIBLIOGRAPHY: sitasi cocok DAN mengandung DOI.
      4. preferred_cite: sitasi terpilih tanpa tautan (buku/non-digital).
    """
    var_key = (var_name or "").strip().lower()
    cite_list = [c for c in (cites or []) if c and str(c).strip()]
    preferred = (preferred_cite or "").strip()
    if preferred and preferred not in cite_list:
        cite_list.insert(0, preferred)

    vindex = _load_variable_lib_index()

    # 1. variable_library berdasarkan nama variabel (DOI valid / PDF tersaji)
    src = vindex["by_variable"].get(var_key)
    if src and _source_usable(src):
        return _source_to_sumber(src)

    # 2. variable_library berdasarkan label sitasi (dari cites yang tersedia)
    if cite_list:
        for c in cite_list:
            cand = vindex["by_cite"].get(_normalize_cite_key(c))
            if cand and _source_usable(cand):
                return _source_to_sumber(cand)

    # 2b. Pencarian sitasi di bibliografi ber-DOI.
    #     Dukungan tokoh_ahli ganda: "A (2022); B (2023); C (2022)" dipecah dulu.
    for c in cite_list:
        for part in _split_cites(c):
            hit = _bibliography_hit_with_doi(part)
            if hit:
                return _entry_to_sumber(hit, label_fallback=part)

    # 3. Bibliography S1 yang memiliki DOI
    lit = _load_literature_lib_s1()
    if lit:
        preferred_key = _normalize_cite_key(preferred) if preferred else ""
        # 3a. Cari sitasi yang sama persis dengan preferred_cite (prioritas tinggi)
        for e in lit:
            e_cite = (e.get("cite") or "").strip()
            if preferred_key and _normalize_cite_key(e_cite) == preferred_key:
                doi = _extract_doi_from_apa(e.get("source") or "")
                if doi and doi_is_valid(doi):
                    return {
                        "label": e_cite or preferred,
                        "apa": f"{e.get('author','')} ({e.get('year','')}). {e.get('title','')}. {e.get('source','')}".strip(),
                        "doi": doi,
                        "url": doi,
                        "journal": (e.get("source") or "").split(".")[0].strip()[:80],
                        "download_kind": "doi",
                    }
        # 3b. Cari sitasi mana pun yang cocok dengan salah satu cites
        for c in cite_list:
            c_key = _normalize_cite_key(c)
            if not c_key:
                continue
            for e in lit:
                e_cite = (e.get("cite") or "").strip()
                if e_cite and _normalize_cite_key(e_cite) == c_key:
                    doi = _extract_doi_from_apa(e.get("source") or "")
                    if doi and doi_is_valid(doi):
                        return {
                            "label": e_cite,
                            "apa": f"{e.get('author','')} ({e.get('year','')}). {e.get('title','')}. {e.get('source','')}".strip(),
                            "doi": doi,
                            "url": doi,
                            "journal": (e.get("source") or "").split(".")[0].strip()[:80],
                            "download_kind": "doi",
                        }

    # 3c. Pencarian konsep variabel di bibliografi ber-DOI (sektor perbankan/digital dll.)
    concept_hit = _bibliography_concept_hit(var_name)
    if concept_hit:
        return _entry_to_sumber(concept_hit, label_fallback=var_name)

    # 4. Fallback: sitasi terpilih tanpa tautan unduhan (buku/karya cetak)
    fallback_cite = preferred or (cite_list[0] if cite_list else "Robbins & Judge (2022)")
    return {
        "label": fallback_cite,
        "apa": fallback_cite,
        "doi": "",
        "url": "",
        "journal": "",
        "download_kind": "none",
    }


def _s1_var_to_profile_for_options(v: Dict[str, Any]) -> Dict[str, Any]:
    """Sintesis profil pseudo dari variabel blueprint tersimpan untuk
    membangun ulang source_options bila field tsb tidak ada (cache lama)."""
    dims = v.get("dimensi") or []
    dimensions = []
    for d in dims:
        if not isinstance(d, dict):
            continue
        dimensions.append({
            "dim": d.get("nama") or "Dimensi",
            "desc": d.get("desc", ""),
            "items": d.get("indikator") or [],
        })
    return {
        "name": v.get("nama_resmi") or v.get("input_name") or "",
        "cites": v.get("cites") or [],
        "theory_cite": v.get("theory_cite") or "",
        "tokoh_ahli": v.get("tokoh_ahli") or "",
        "dimensions": dimensions,
        "engine": v.get("engine", "s1-engine"),
    }


def enrich_s1_blueprint_sources(blueprint: Dict[str, Any]) -> Dict[str, Any]:
    """Melengkapi setiap variabel blueprint S1 dengan field 'sumber' (1 sumber tunggal)."""
    if not blueprint or not isinstance(blueprint, dict):
        return blueprint
    vars_map = blueprint.get("variables") or {}
    meta = blueprint.get("metadata") or {}
    instansi = meta.get("instansi") or "Instansi Terkait"
    title = meta.get("judul") or blueprint.get("title") or ""
    for code, v in vars_map.items():
        if not isinstance(v, dict):
            continue
        # Heal: cache lama tanpa source_options (profil korpus pengganti yang
        # kehilangan field saat research versi awal) dibangun ulang agar
        # dropdown per-variabel tetap tersedia di UI.
        if not v.get("source_options"):
            try:
                pseudo = _s1_var_to_profile_for_options(v)
                opts, chosen_key, _ = _attach_corpus_sources(
                    pseudo,
                    v.get("input_name") or v.get("nama_resmi") or code,
                    v.get("role") or code.upper(),
                    instansi,
                    title,
                )
                v["source_options"] = opts
                if not v.get("chosen_source"):
                    v["chosen_source"] = chosen_key
            except Exception as _ex_heal:
                print(f"[s1_ontology] heal source_options {code} gagal: {_ex_heal}")
        if not v.get("sumber"):
            # Prioritas 0: sumber korpus terpilih (PDF lokal references/02_pendukung)
            chosen = v.get("chosen_source") or ""
            if chosen and chosen.startswith("corpus|"):
                for opt in (v.get("source_options") or []):
                    if opt.get("key") == chosen and opt.get("sumber"):
                        v["sumber"] = dict(opt["sumber"])
                        break
            if not v.get("sumber"):
                cites_list = v.get("cites") or []
                # Blueprint dari preset menyimpan daftar ahli pada 'tokoh_ahli' (dipisah ';')
                prefer = v.get("theory_cite") or v.get("tokoh_ahli") or ""
                src = resolve_single_s1_source(
                    var_name=v.get("input_name") or v.get("nama_resmi") or code,
                    preferred_cite=prefer,
                    cites=cites_list,
                )
                v["sumber"] = src
    return blueprint


# ---------------------------------------------------------------------------
# 3c. SUMBER RUJUKAN & KUTIPAN (Daftar Pustaka 20-Tahun, dikonfigurasi)
# ---------------------------------------------------------------------------
# Perjanjian pengguna:
#   - SEMUA kutipan bab wajib berada dalam rentang tahun yang dipilih di
#     Menu 1 (DEFAULT: 20 TAHUN KEBELAKANG, mis. 2006-2026).
#   - 95 PDF pendukung TIDAK masuk daftar ini (korpus dimensi/indikator only).
#   - Bila variabel belum punya sumber dalam rentang -> telusur live OpenAlex
#     ber-DOI dipakai sebagai cadangan.
_S1_CITE_STAPLES = [
    {"label": "Hair, Hult, Ringle & Sarstedt (2022)",
     "apa": "Hair, J. F., Hult, G. T. M., Ringle, C. M., & Sarstedt, M. (2022). A Primer on Partial Least Squares Structural Equation Modeling (PLS-SEM) (3rd ed.). SAGE Publications.",
     "year": 2022},
    {"label": "Sugiyono (2019)",
     "apa": "Sugiyono. (2019). Metode Penelitian Kuantitatif, Kualitatif, dan R&D. Alfabeta.",
     "year": 2019},
    {"label": "Sekaran & Bougie (2021)",
     "apa": "Sekaran, U., & Bougie, R. (2021). Research Methods for Business: A Skill-Building Approach (8th ed.). Wiley.",
     "year": 2021},
]


def _year_from_apa(apa):
    m = re.search(r"\((\d{4}[a-z]?)\)", str(apa or ""))
    return int(m.group(1)) if m else 0


def _citation_candidates_for(var_name):
    """Kandidat sitasi lokal (variable_library + bibliografi S1) untuk satu variabel."""
    cands = []
    try:
        vindex = _load_variable_lib_index()
        src = vindex["by_variable"].get((var_name or "").strip().lower())
        if src:
            cands.append(dict(src))
    except Exception:
        pass
    try:
        lit = _load_literature_lib_s1()
        vkey = _normalize_variable_key(var_name)
        keywords = []
        for k, kw in _S1_CONCEPT_KEYWORDS.items():
            if k in vkey or vkey in k:
                keywords.extend(kw)
        keywords = list(dict.fromkeys(keywords))
        for e in lit or []:
            hay = ((e.get("cite") or "") + " " + (e.get("title") or "") + " "
                   + (e.get("source") or "")).lower()
            if keywords and any(kw in hay for kw in keywords):
                cands.append({"label": e.get("cite", ""), "apa": e.get("source", ""),
                              "year": _year_from_apa(e.get("source"))})
    except Exception:
        pass
    return cands


def _topic_pool_candidates(var_name):
    """Kandidat dari pool jurnal terverifikasi academic_connector (di-fetch by var)."""
    out = []
    try:
        import academic_connector as _ac
        topic = _ac.detect_variable_topic(var_name)
        pool = _ac.TOPIC_JOURNAL_POOLS.get(topic) or []
        for p in pool:
            year = p.get("year")
            if not year:
                continue
            label = _ac.make_in_text_citation(p.get("authors") or [], year)
            apa, _ = _ac.build_apa_entry(p)
            out.append({"label": label, "apa": apa, "year": year, "doi": p.get("doi"),
                        "url": p.get("doi"), "journal": p.get("journal", "")})
    except Exception as e:
        print(f"[s1_citeref] pool topic warning: {e}")
    return out


def _live_citation_fallback(var_name, min_year, max_year, limit=2):
    """Cadangan wajib: telusur live OpenAlex/Crossref dalam rentang tahun aktif."""
    out = []
    try:
        import academic_connector as _ac
        for query in (f"pengaruh {var_name} terhadap kinerja pegawai",
                      f"{var_name} kinerja pegawai", var_name):
            papers = _ac.search_openalex_verified(query, limit=4,
                                                  min_year=min_year)
            for paper in papers:
                year = paper.get("year") or 0
                if not (min_year <= year <= max_year):
                    continue
                doi = paper.get("doi") or ""
                apa, _ = _ac.build_apa_entry(paper)
                label = _ac.make_in_text_citation(paper.get("authors") or [], year)
                out.append({"label": label, "apa": apa, "year": year,
                            "doi": doi, "url": doi,
                            "journal": paper.get("journal", "")})
                if len(out) >= limit:
                    break
            if out:
                break
    except Exception as e:
        print(f"[s1_citeref] live fallback warning: {e}")
    return out


# ---------------------------------------------------------------------------
# Heuristik tag Sumber Rujukan (asal & indeks) utk filter "Pengaturan sitasi"
# ---------------------------------------------------------------------------
_INTERNATIONAL_JOURNAL_RE = re.compile(
    r"(international\s+journal|\bjournal of\b|\bmanagement science\b|"
    r"\bstrategic management\b|\bhuman resource management\b|\bscopus\b|"
    r"elsevier|springer|\bwiley\b|emerald|sage publications|"
    r"taylor\s*&\s*francis|routledge|\bmdpi\b|\bieee\b|"
    r"academy of management|oxford university|cambridge university|"
    r"\bdoi\.org/10\.)", re.I)

_SINTA_JOURNAL_RE = re.compile(
    r"(\bjurnal\b|e-jurnal|jurnal ilmiah|akreditasi sinta|\bsinta\b|"
    r"terakreditasi nasional|stie\b|stia\b|stim\b|stiami\b|uin\b|"
    r"universitas\b|politeknik\b|\bipb\b|\bugm\b|\bunair\b|\bundip\b|"
    r"\bunpad\b|\bitb\b|\bums\b|\buii\b|alfabeta|salemba|erlangga|"
    r"rajawali\b)", re.I)


def _tag_citation_entry(cand):
    """Tag asal & indeks per entri kandidat (best-effort heuristik nama jurnal).

    Returns (origin_tag, index_tag) dgn nilai:
      origin : "internasional" | "indonesia" | "umum"
      index  : "scopus" | "sinta" | "umum"
    """
    hay = " ".join(str(cand.get(k) or "") for k in
                   ("journal", "apa", "title", "source", "label", "authors"))
    if _INTERNATIONAL_JOURNAL_RE.search(hay):
        return ("internasional", "scopus")
    if _SINTA_JOURNAL_RE.search(hay):
        return ("indonesia", "sinta")
    return ("umum", "umum")


def build_s1_citation_references(config: Dict[str, Any]) -> Dict[str, Any]:
    """Bangun daftar 'Sumber Rujukan & Kutipan' utk seluruh variabel S1.

    Pengaturan sitasi lengkap (rentang tahun, asal, gaya, bahasa, sumber)
    diambil dari academic_connector.get_cite_settings(). Setiap variabel
    mendapat maks. 3 sitasi dalam rentang; bila kosong -> telusur live OpenAlex.
    Filter asal & sumber bersifat best-effort berbasis heuristik nama jurnal.
    """
    try:
        import academic_connector as _ac
        min_y, max_y = _ac.get_cite_year_range()
    except Exception:
        min_y, max_y = 2006, 2026

    cite_settings = {}
    try:
        import academic_connector as _ac2
        cite_settings = _ac2.get_cite_settings() or {}
    except Exception:
        cite_settings = {}
    if cite_settings.get("min_year") is not None:
        min_y = int(cite_settings["min_year"])
    if cite_settings.get("max_year") is not None:
        max_y = int(cite_settings["max_year"])
    origin_sel = cite_settings.get("origin") or "all"
    style_sel = cite_settings.get("style") or "apa7"
    lang_sel = cite_settings.get("language") or "id"
    src_sel = [s for s in (cite_settings.get("source_types") or [])
               if s in ("umum", "sinta", "scopus")]
    # "Umum" berarti "sumber mana pun" -> tanpa filter indeks bila dicentang.
    allow_any_src = ("umum" in src_sel) or (not src_sel)

    def _passes_cite_filter(cand):
        ot, it = _tag_citation_entry(cand)
        if origin_sel == "indonesia" and ot != "indonesia":
            return False
        if origin_sel == "internasional" and ot != "internasional":
            return False
        if not allow_any_src and it not in src_sel:
            return False
        return True

    judul = config.get("judul") or config.get("title") or ""
    instansi = config.get("instansi") or "Instansi Terkait"
    x_names = [x for x in (config.get("x_names") or []) if x and str(x).strip()][:5]
    m_name = (config.get("m_name") or config.get("m") or "").strip()
    y_name = (config.get("y_name") or config.get("y") or "Kinerja Pegawai").strip()
    model_type = config.get("model_type") or "regression"

    roles = [(f"X{i + 1}", str(x).strip()) for i, x in enumerate(x_names)]
    if model_type == "mediation" and m_name:
        roles.append(("M", m_name))
    roles.append(("Y", y_name))

    seen = set()
    groups = []
    total = 0
    for role, vname in roles:
        if not vname:
            continue
        hits = _citation_candidates_for(vname) + _topic_pool_candidates(vname)
        picked = []
        for cand in hits:
            y = cand.get("year") or _year_from_apa(cand.get("apa"))
            if not y or not (min_y <= y <= max_y):
                continue
            label = cand.get("label") or cand.get("cite") or ""
            apa = cand.get("apa") or label
            if not label or not _passes_cite_filter(cand):
                continue
            key = (label.strip().lower(), y)
            if key in seen:
                continue
            seen.add(key)
            ot, it = _tag_citation_entry(cand)
            picked.append({
                "label": label,
                "apa": apa,
                "year": y,
                "doi": cand.get("doi") or "",
                "url": cand.get("url") or cand.get("doi") or "",
                "journal": cand.get("journal") or "",
                "source_kind": cand.get("source_kind", "terverifikasi"),
                "origin_tag": ot,
                "index_tag": it,
            })
            if len(picked) >= 3:
                break
        if not picked:
            live = _live_citation_fallback(vname, min_y, max_y)
            for lr in live:
                if not _passes_cite_filter(lr):
                    continue
                key = ((lr.get("label") or "").strip().lower(), lr.get("year") or 0)
                if key in seen:
                    continue
                seen.add(key)
                lr["source_kind"] = "live-openalex"
                ot, it = _tag_citation_entry(lr)
                lr["origin_tag"] = ot
                lr["index_tag"] = it
                picked.append(lr)
                if len(picked) >= 2:
                    break
        total += len(picked)
        groups.append({"role": role, "variable": vname, "entries": picked})

    # Pelengkap metodologi (dalam rentang) — sekali, bukan per variabel.
    staples = []
    for s in _S1_CITE_STAPLES:
        y = s.get("year") or 0
        if min_y <= y <= max_y and _passes_cite_filter(s):
            key = (s["label"].lower(), y)
            if key not in seen:
                seen.add(key)
                s2 = dict(s)
                s2["source_kind"] = "metodologi"
                s2["doi"] = ""
                s2["url"] = ""
                s2["journal"] = ""
                ot, it = _tag_citation_entry(s2)
                s2["origin_tag"] = ot
                s2["index_tag"] = it
                staples.append(s2)
    total += len(staples)

    return {
        "status": "success",
        "judul": judul,
        "instansi": instansi,
        "min_year": min_y,
        "max_year": max_y,
        "total": total,
        "groups": groups,
        "staples": staples,
        "settings": {
            "year_mode": cite_settings.get("year_mode") or "all",
            "min_year": cite_settings.get("min_year"),
            "max_year": cite_settings.get("max_year"),
            "origin": origin_sel,
            "style": style_sel,
            "language": lang_sel,
            "source_types": src_sel,
            "source_files": cite_settings.get("source_files") or [],
        },
    }



def apply_s1_blueprint(blueprint: Dict[str, Any], title: str = "") -> bool:
    """Terapkan blueprint (yang mungkin sudah di-swap sumbernya per variabel).

    Dipanggil saat 💾 Simpan Konfigurasi: blueprint disimpan ke cache server
    (sumber kebenaran penulisan bab) dan tiap variabel didaftarkan ulang ke
    memori dinamis s1_constructs agar bab membaca dimensi/indikator pilihan.
    """
    if not blueprint or not isinstance(blueprint, dict):
        return False
    try:
        vars_map = blueprint.get("variables") or {}
        for _vk, vd in vars_map.items():
            if not isinstance(vd, dict):
                continue
            vname = vd.get("input_name") or vd.get("nama_resmi") or ""
            if vname:
                s1_constructs.register_dynamic_construct(vname, vd)
        if not title:
            title = (blueprint.get("metadata") or {}).get("judul") or ""
        save_s1_ontology(blueprint, title)
        print(f"[s1_ontology] Blueprint diterapkan & disimpan utk '{title[:50]}' "
              f"({len(vars_map)} variabel).")
        return True
    except Exception as e:
        print(f"[s1_ontology] apply_s1_blueprint gagal: {e}")
        return False


# ---------------------------------------------------------------------------
# 4. Master Engine Riset Blueprint Ontologi Skripsi S1
# ---------------------------------------------------------------------------

def research_s1_ontology(
    title: str,
    x_vars: List[str],
    m_var: Optional[str] = None,
    y_var: str = "Kinerja Pegawai",
    model_type: str = "regression",
    instansi: str = "Instansi Terkait",
    sektor: str = "",
    force_refresh: bool = False
) -> Dict[str, Any]:
    """
    Riset Komprehensif Dimensi & Indikator Seluruh Variabel Skripsi S1 (Multi-X 2..5).
    Menghasilkan Blueprint Ontologi S1 yang mengunci Benang Merah dari Bab I s/d Lampiran.
    """
    title_clean = (title or "").strip()
    instansi_clean = (instansi or "Instansi Terkait").strip()
    y_clean = (y_var or "Kinerja Pegawai").strip()
    x_clean_list = [x.strip() for x in (x_vars or []) if x and x.strip()]
    if not x_clean_list:
        x_clean_list = ["Kepemimpinan Transformasional", "Budaya Organisasi"]

    # Cek cache bila tidak force refresh
    if not force_refresh:
        cached = load_s1_ontology(title_clean)
        if cached:
            # Periksa apakah variabel bebasnya cocok
            vars_cached = cached.get("variables", {})
            cached_x_names = [v.get("input_name", "") for k, v in vars_cached.items() if k.lower().startswith("x")]
            if cached_x_names == x_clean_list and vars_cached.get("y", {}).get("input_name") == y_clean:
                print(f"[s1_ontology] Menggunakan cache Blueprint Ontologi S1 untuk '{title_clean[:50]}...'")
                # Pastikan setiap variabel membawa sumber tunggal terverifikasi
                enrich_s1_blueprint_sources(cached)
                # Daftarkan kembali ke memori dinamis s1_constructs
                for vk, vdata in vars_cached.items():
                    s1_constructs.register_dynamic_construct(vdata.get("input_name", ""), vdata)
                return cached

    # Siapkan daftar spesifikasi variabel (role_code, var_name, role_label)
    specs: List[Tuple[str, str, str]] = []
    for i, x_name in enumerate(x_clean_list):
        code = f"x{i+1}"
        role_label = f"X{i+1}"
        specs.append((code, x_name, role_label))

    if model_type == "mediation" and m_var and m_var.strip():
        specs.append(("m", m_var.strip(), "M"))

    specs.append(("y", y_clean, "Y"))

    variables_result: Dict[str, Any] = {}
    kisi_kisi_list: List[Dict[str, Any]] = []
    tabel_1_1_params: List[str] = []
    tabel_1_2_issues: List[str] = []
    total_indikator = 0
    total_dimensi = 0

    for code, vname, role_label in specs:
        prof = research_single_s1_variable(
            var_name=vname,
            role=role_label,
            instansi=instansi_clean,
            title=title_clean
        )

        # Daftarkan ke s1_constructs secara dinamis
        s1_constructs.register_dynamic_construct(vname, prof)

        # Susun struktur ontologi untuk variabel ini
        dims_clean = []
        item_counter = 1
        for d in prof.get("dimensions", [])[:3]:
            dim_name = d.get("dim") or "Dimensi"
            dim_items = d.get("items") or []
            total_dimensi += 1
            dims_clean.append({
                "nama": dim_name,
                "desc": d.get("desc", ""),
                "indikator": dim_items,
                "jumlah_butir": len(dim_items)
            })

            for it_text in dim_items:
                item_code = f"{role_label}_{item_counter}"
                kisi_kisi_list.append({
                    "var_code": role_label,
                    "var_name": vname,
                    "dim_name": dim_name,
                    "item_code": item_code,
                    "statement": it_text
                })
                item_counter += 1
                total_indikator += 1

                # Masukkan indikator Y ke parameter evaluasi pra-survei Bab I
                if role_label == "Y":
                    tabel_1_1_params.append(f"{dim_name}: {it_text[:65]}...")

        if role_label.startswith("X") or role_label == "M":
            tabel_1_2_issues.append(f"{role_label} ({vname}): {prof.get('issue', '')}")

        variables_result[code] = {
            "role": role_label,
            "input_name": vname,
            "nama_resmi": prof.get("name") or vname,
            "grand_theory": prof.get("grand_theory", "Organizational Theory"),
            "tokoh_ahli": prof.get("tokoh_ahli") or prof.get("theory_cite", "Robbins & Judge (2022)"),
            "definisi": prof.get("definition", [("", "")])[0][0] if prof.get("definition") else "",
            "synthesis": prof.get("synthesis", ""),
            "dimensi": dims_clean,
            "cites": prof.get("cites", []),
            "issue": prof.get("issue", ""),
            "effect": prof.get("effect", ""),
            "skala_pengukuran": "Skala Likert 1–5",
            "engine": prof.get("engine", "s1-engine"),
            "source_options": prof.get("source_options", []),
            "chosen_source": prof.get("chosen_source", "online"),
        }

    blueprint_s1 = {
        "status": "success",
        "metadata": {
            "judul": title_clean,
            "instansi": instansi_clean,
            "model_type": model_type,
            "n_x": len(x_clean_list),
            "total_dimensi": total_dimensi,
            "total_indikator": total_indikator,
            "skala": "Skala Likert 1-5 (1: Sangat Tidak Setuju s/d 5: Sangat Setuju)",
            "timestamp": datetime.now().isoformat(),
            "jenjang": "S1 (Sarjana)",
            "needs_review": []
        },
        "variables": variables_result,
        "anchors": {
            "tabel_1_1_params": tabel_1_1_params,
            "tabel_1_2_issues": tabel_1_2_issues,
            "kisi_kisi": kisi_kisi_list
        }
    }

    # Kuasai blueprint final dengan sumber tunggal per variabel
    blueprint_s1 = enrich_s1_blueprint_sources(blueprint_s1)

    # Simpan ke cache
    save_s1_ontology(blueprint_s1, title_clean)
    return blueprint_s1
