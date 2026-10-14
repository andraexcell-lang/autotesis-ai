# backend/llm_background_writer.py
"""
AGEN PENULIS LATAR BELAKANG MASALAH (Sub Bab 1.1) — 100% dikerjakan model AI (Gemini).

Kerangka prompt (dari pengguna, sudah diperbaiki):
  1. Alur umum -> khusus: NASIONAL -> pengertian & teori Y -> faktor Y -> teori M -> teori X1 -> teori X2
  2. Paragraf profil objek penelitian, dihubungkan dengan fenomena Y
  3. Paragraf fenomena belum optimalnya kinerja (Y) + data tabel pendukung
  4. Paragraf fenomena M (dimensi & indikator)
  5. Paragraf fenomena X1 (dimensi & indikator)
  6. Paragraf fenomena X2 (dimensi & indikator)
  7. Dua paragraf research gap (membandingkan penelitian terdahulu)
  8. Satu paragraf penutup (penetapan judul)

Jaminan anti-halusinasi:
  - Sitasi HANYA dari daftar sitasi terverifikasi DOI (OpenAlex) yang dikirim engine.
  - Angka tabel HANYA dari engine (AI dilarang mengarang angka kinerja instansi).
  - Validasi: semua paragraf ada, tiap paragraf wajar panjangnya, tidak ada sitasi di luar daftar.
  - Bila gagal -> None (engine memakai template cadangan).
"""

import re
import time

from llm_title_parser import load_config, MODEL_CHAIN, gemini_keys
from llm_ontology_researcher import _call_gemini_matrix

BG_TIMEOUT = 90
BG_MAX_TOKENS = 4096

REQUIRED_KEYS = (
    "nasional",        # 1. pembuka nasional
    "teori_y",         # 1b. pengertian & teori Y + faktor-faktor Y
    "teori_m",         # teori M
    "teori_x1",        # teori X1
    "teori_x2",        # teori X2
    "profil_objek",    # 2. profil objek + hubungan ke fenomena Y
    "fenomena_y",      # 3. fenomena belum optimalnya Y (+ tabel)
    "fenomena_m",      # 4. fenomena M
    "fenomena_x1",     # 5. fenomena X1
    "fenomena_x2",     # 6. fenomena X2
    "gap_1",           # 7a. research gap (perbandingan penelitian terdahulu)
    "gap_2",           # 7b. research gap (lanjutan, arah mediasi)
    "penutup",         # 8. penutup & penetapan judul
)

PROMPT_BG = """Anda adalah penulis BAB I tesis Magister Manajemen (S2) Indonesia. Tulis SUB BAB 1.1 LATAR BELAKANG MASALAH secara utuh dan natural, dengan alur umum ke khusus.

=== JUDUL PENELITIAN ===
{judul}

=== KONTEKS OBJEK PENELITIAN ===
Instansi/objek : {instansi}
Sektor         : {sektor}
Peran strategis: {peran}
Isu lapangan   : {isu}

=== VARIABEL & ONTOLOGI HASIL RISET (wajib dijadikan dasar; jangan menyimpang) ===
{ontologi_block}

=== DATA TABEL PENDUKUNG (ANGKA NYATA DARI SISTEM — JANGAN UBAH, JANGAN TAMBAH ANGKA BARU) ===
Tabel 1.1 — capaian {y} (tahun, target, realisasi, gap):
{tabel1_block}
Tabel 1.2 — isu operasional (ringkasan): {tabel2_ringkas}

=== DAFTAR SITASI TERVERIFIKASI (satu-satunya sitasi yang BOLEH dipakai, format APA in-text) ===
{sitasi_block}

=== STRUKTUR KELUARAN: JSON dengan 13 kunci, masing-masing satu paragraf ===
"nasional"     : Pembuka NASIONAL: urgensi isu {sektor} dan tata kelola/aparatur di tingkat nasional. Sitasi: {c_macro}; {c_sdm}
"teori_y"      : Pengertian {y} menurut para ahli (sebut nama teori & tokoh yang tepat untuk {y}), lalu faktor-faktor yang memengaruhi {y}. Sitasi: {c_kinerja}
"teori_m"      : Pengertian {m} dan teorinya yang tepat, hubungan {m} dengan {y}. Sitasi: {c_m1}
"teori_x1"     : Pengertian {x1} dan teorinya yang tepat, hubungan {x1} dengan {y}. Sitasi: {c_x11}
"teori_x2"     : Pengertian {x2} dan teorinya yang tepat, hubungan {x2} dengan {y}. Sitasi: {c_x21}
"profil_objek" : Profil {instansi} (peran & fungsi), dan hubungkan langsung dengan fenomena {y}. TANPA sitasi.
"fenomena_y"   : Fenomena BELUM OPTIMALNYA {y} di {instansi}: jelaskan dimensi & indikator {y}, rujuk angka Tabel 1.1 secara eksplisit (tahun & gap) DAN gunakan persis angka dari tabel tanpa mengubahnya, serta sebut isu operasional Tabel 1.2. Sitasi: {c_compl}
"fenomena_m"   : Fenomena {m} di {instansi} berdasarkan dimensi & indikator {m}, dan dampaknya pada {y}. TANPA sitasi.
"fenomena_x1"  : Fenomena {x1} di {instansi} berdasarkan dimensi & indikator {x1}, dan dampaknya pada {y}. TANPA sitasi.
"fenomena_x2"  : Fenomena {x2} di {instansi} berdasarkan dimensi & indikator {x2}, dan dampaknya pada {y}. TANPA sitasi.
"gap_1"        : RESEARCH GAP 1 — bandingkan penelitian terdahulu tentang pengaruh {x1} dan {x2} terhadap {y}. Tunjukkan KONTRADIKSI atau celah hasil. Sitasi: {c_x12}; {c_x22}
"gap_2"        : RESEARCH GAP 2 — bandingkan penelitian terdahulu tentang {m} sebagai mediasi. Tunjukkan belum adanya kajian pada objek {instansi}. Sitasi: {c_med1}; {c_med2}
"penutup"      : Kesimpulan urgensi penelitian, posisikan penelitian sebagai konfirmatori, lalu tetapkan judul penelitian dalam tanda kutip persis seperti JUDUL di atas. TANPA sitasi.

=== ATURAN KERAS ===
1. Bahasa Indonesia akademik formal. Tiap paragraf ±5 kalimat (sekitar 5 baris), mengalir, tidak kaku.
2. SITASI: gunakan HANYA sitasi dari DAFTAR SITASI TERVERIFIKASI, format APA persis seperti tertulis. DILARANG menambah, mengubah, atau mengarang nama penulis, tahun, atau jurnal.
3. ANGKA: hanya angka dari DATA TABEL PENDUKUNG. DILARANG mengarang angka, persentase, atau hasil statistik baru.
4. DILARANG mengklaim hasil penelitian ini sudah terbukti. Gunakan "diduga", "berpotensi", "perlu dikonfirmasi".
5. DILARANG mengulang kalimat pembuka yang sama antar paragraf. Setiap paragraf membuka dengan kalimat berbeda.
6. DILARANG menyebut "Tabel 1.1" atau "Tabel 1.2" kecuali di "fenomena_y".
7. Dimensi & indikator yang disebut harus berasal dari ONTOLOGI di atas, bukan karangan.
8. HANYA balas JSON valid berisi 13 kunci di atas, tanpa teks lain.
9. TEORI harus TEPAT dan RELEVAN untuk variabelnya, sesuai dengan 'Definisi' variabel di ONTOLOGI. DILARANG memakai teori motivasi untuk variabel kinerja/produktivitas, dan sebaliknya. Setiap variabel (Y, M, X1, X2) harus memakai teori yang sesuai variabelnya masing-masing; jangan menyalin teori antar variabel.
10. SETIAP SITASI HANYA BOLEH MUNCUL SATU KALI di seluruh naskah. Dilarang mengulang sitasi yang sama di paragraf lain.
11. Gunakan HANYA sitasi dengan tahun 2022–2026 dari DAFTAR SITASI TERVERIFIKASI. Sitasi tahun di luar rentang itu DILARANG, kecuali teori klasik yang tercantum di ONTOLOGI.
12. Angka pada "fenomena_y" harus IDENTIK dengan DATA TABEL PENDUKUNG; dilarang membulatkan ulang atau mengubah angka.
"""


def _fmt_citation(acad, key):
    return acad.get(key, "")


def _ontologi_block(onto):
    lines = []
    for key, label in (("x1", "X1"), ("x2", "X2"), ("m", "M"), ("y", "Y")):
        v = ((onto or {}).get("variables") or {}).get(key) or {}
        if not v:
            continue
        dims = []
        for d in v.get("dimensi", []):
            inds = d.get("indikator", [])
            dims.append(f"{d.get('nama', '-')} (indikator: {'; '.join(inds)})")
        lines.append(
            f"- {label} {v.get('nama_resmi', '-')}\n"
            f"  Teori   : {v.get('grand_theory', '-')}\n"
            f"  Definisi: {v.get('definisi', '-')}\n"
            f"  Dimensi : {' | '.join(dims) if dims else '(belum diriset)'}"
        )
    return "\n".join(lines) if lines else "(ontologi belum tersedia)"


def _tabel1_block(ctx):
    rows = ctx.get("iku_data") or []
    headers = ctx.get("iku_headers") or []
    out = [" | ".join(headers)] if headers else []
    for r in rows:
        out.append(" | ".join(str(c) for c in r))
    return "\n".join(out) if out else "(data tabel 1.1 tidak tersedia)"


def _sitasi_block(acad):
    keys = [
        ("cite_macro", "Makro (nasional)"), ("cite_sdm", "SDM sektor publik"),
        ("cite_kinerja", "Kinerja Y"), ("cite_m_1", "Variabel M"),
        ("cite_x1_1", "Variabel X1 #1"), ("cite_x1_2", "Variabel X1 #2"),
        ("cite_x2_1", "Variabel X2 #1"), ("cite_x2_2", "Variabel X2 #2"),
        ("cite_compl", "Kepatuhan semu"), ("cite_med_1", "Mediasi #1"),
        ("cite_med_2", "Mediasi #2"),
    ]
    return "\n".join(f"- {label}: {acad.get(k, '')}" for k, label in keys if acad.get(k))


def _tokoh_sah(onto):
    """Kumpulkan nama tokoh/teori dari ontologi agar sitasi klasik (mis. Herzberg) tetap sah."""
    names = []
    for v in ((onto or {}).get("variables") or {}).values():
        blob = f"{v.get('grand_theory', '')} {v.get('tokoh_ahli', '')}"
        names += re.findall(r"([A-Z][A-Za-z\-]+)\s*[&,]?\s*(?:et al\.?)?\s*\(?(?:19|20)\d{2}\)?", blob)
    return names


def _validate_bg(raw, acad, judul, extra_sah=None):
    """Validasi ketat: 13 paragraf, panjang wajar, penulis sitasi hanya dari daftar sah, judul tercantum."""
    if isinstance(raw, list) and raw and isinstance(raw[0], dict):
        raw = raw[0]
    if not isinstance(raw, dict):
        return None
    out = {}
    for k in REQUIRED_KEYS:
        val = raw.get(k)
        if not isinstance(val, str):
            return None
        text = val.strip()
        if len(text) < 180:
            return None
        out[k] = text

    # Anti-halusinasi: setiap nama penulis dalam kurung sitasi harus berasal dari
    # daftar sah (sitasi OpenAlex terverifikasi + tokoh/teori dari ontologi).
    sah_text = " ".join(str(v) for v in acad.values() if isinstance(v, str))
    sah_text += " " + " ".join(str(x) for x in (extra_sah or []))
    sah_low = sah_text.lower()
    unknown = []
    for grup in re.findall(r"\(([^()]*?\d{4}[a-z]?[^()]*)\)", " ".join(out.values())):
        for bagian in grup.split(";"):
            m = re.match(r"\s*([A-Z][\w\-]+)", bagian.strip())
            if m and m.group(1).lower() not in sah_low:
                unknown.append(bagian.strip())
    if unknown:
        print(f"[background] penulis di luar daftar sah: {unknown[:3]}")
        return None

    if judul and judul.strip().lower() not in out["penutup"].lower():
        print("[background] judul tidak tercantum di penutup -> tolak")
        return None

    # Anti-gaya (bukan penolakan): catat bila sitasi jurnal OpenAlex terpakai >1 kali.
    # Sitasi teori klasik (Herzberg, Locke, dst.) boleh berulang secara sah.
    from collections import Counter
    journal_cites = {str(v).lower() for k, v in acad.items()
                     if isinstance(v, str) and v.startswith("(")}
    semu = " ".join(out.values()).lower()
    kembar = [c for c in journal_cites if semu.count(c) > 1]
    if kembar:
        print(f"[background] catatan: sitasi jurnal dipakai >1x (gaya, tidak menolak): {kembar[:3]}")
    return out


def write_latar_belakang(cfg, onto, ctx, acad):
    """
    Menulis seluruh sub bab 1.1 via Gemini. Mengembalikan dict paragraf, atau None bila gagal.
    """
    config = load_config()
    # Pakai daftar key; "api_key" tunggal tetap didukung untuk config lama.
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return None

    y = cfg.get("y", "Kinerja Pegawai")
    prompt = PROMPT_BG.format(
        judul=cfg.get("judul", ""),
        instansi=cfg.get("instansi", ""),
        sektor=ctx.get("sektor", cfg.get("sektor", "")),
        peran=ctx.get("peran_strategis", ""),
        isu=ctx.get("isu_lapangan", ""),
        ontologi_block=_ontologi_block(onto),
        y=y,
        m=cfg.get("m", "M"),
        x1=cfg.get("x1", "X1"),
        x2=cfg.get("x2", "X2"),
        tabel1_block=_tabel1_block(ctx),
        tabel2_ringkas=ctx.get("tabel2_nama", "-"),
        sitasi_block=_sitasi_block(acad),
        c_macro=_fmt_citation(acad, "cite_macro"),
        c_sdm=_fmt_citation(acad, "cite_sdm"),
        c_kinerja=_fmt_citation(acad, "cite_kinerja"),
        c_m1=_fmt_citation(acad, "cite_m_1"),
        c_x11=_fmt_citation(acad, "cite_x1_1"),
        c_x12=_fmt_citation(acad, "cite_x1_2"),
        c_x21=_fmt_citation(acad, "cite_x2_1"),
        c_x22=_fmt_citation(acad, "cite_x2_2"),
        c_compl=_fmt_citation(acad, "cite_compl"),
        c_med1=_fmt_citation(acad, "cite_med_1"),
        c_med2=_fmt_citation(acad, "cite_med_2"),
    )

    chain = [m for m in config.get("models", []) if m] or MODEL_CHAIN
    timeout = max(int(config.get("timeout_seconds", BG_TIMEOUT)), BG_TIMEOUT)
    # Matriks: habiskan SEMUA key per model sebelum turun ke model berikutnya.
    keys = gemini_keys(config)
    try:
        raw = _call_gemini_matrix(keys, chain, prompt, timeout, max_tokens=BG_MAX_TOKENS)
        result = _validate_bg(raw, acad, cfg.get("judul", ""), _tokoh_sah(onto))
        if result:
            print(f"[background] Latar belakang 1.1 ditulis (matriks)")
            return result
        print(f"[background] keluaran tidak valid -> gagal")
        return None
    except Exception as e:
        print(f"[background] matriks gagal: {type(e).__name__} {str(e)[:80]}")
        return None
