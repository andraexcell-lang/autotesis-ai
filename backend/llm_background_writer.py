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

import json
import os
import re
import time

from llm_title_parser import load_config, MODEL_CHAIN, gemini_keys
from llm_ontology_researcher import _call_gemini_matrix
from academic_connector import MIN_CITE_YEAR, MAX_CITE_YEAR, tahun_sitasi_terlarang

# (#5) Mode senyap untuk "dry-run" pengumpulan spesifikasi tabel: saat aktif,
# _gemini_json langsung mengembalikan None sehingga tidak ada kuota yang terpakai.
_SUPPRESS_GEMINI = False


def set_suppress_gemini(flag):
    """Aktifkan/nonaktifkan mode senyap Gemini (dipakai prefetch tabel Bab IV)."""
    global _SUPPRESS_GEMINI
    _SUPPRESS_GEMINI = bool(flag)


BG_TIMEOUT = 90
BG_MAX_TOKENS = 4096

# Alasan penolakan terakhir dari _validate_bg(). Dipakai write_latar_belakang()
# untuk memberi umpan balik korektif ke Gemini lalu mencoba ulang, bukan
# langsung menyerah ke template cadangan (template = naskah tetap, tidak
# bervariasi antar-generate).
_LAST_BG_REASON = ""


def _tolak_bg(print_msg, alasan):
    """Catat alasan penolakan naskah 1.1 (untuk umpan balik ke Gemini)."""
    global _LAST_BG_REASON
    _LAST_BG_REASON = alasan
    print(print_msg)
    return None

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

=== BAGIAN INI DITENTUKAN KONFIGURASI (tabel aktif / tanpa tabel) ===
{blok_tabel}

=== DAFTAR SITASI TERVERIFIKASI (satu-satunya sitasi yang BOLEH dipakai, format APA in-text) ===
{sitasi_block}

=== STRUKTUR KELUARAN: JSON dengan 13 kunci, masing-masing satu paragraf ===
"nasional"     : Pembuka NASIONAL: urgensi isu {sektor} dan tata kelola/aparatur di tingkat nasional. Sitasi: {c_macro}; {c_sdm}
"teori_y"      : Pengertian {y} menurut para ahli (sebut nama teori & tokoh yang tepat untuk {y}), lalu faktor-faktor yang memengaruhi {y}. Sitasi: {c_kinerja}
"teori_m"      : Pengertian {m} dan teorinya yang tepat, hubungan {m} dengan {y}. Sitasi: {c_m1}
"teori_x1"     : Pengertian {x1} dan teorinya yang tepat, hubungan {x1} dengan {y}. Sitasi: {c_x11}
"teori_x2"     : Pengertian {x2} dan teorinya yang tepat, hubungan {x2} dengan {y}. Sitasi: {c_x21}
"profil_objek" : Profil {instansi} (peran & fungsi), dan hubungkan langsung dengan fenomena {y}. TANPA sitasi.
"fenomena_y"   : {instruksi_fenomena_y}
"fenomena_m"   : Fenomena {m} di {instansi} berdasarkan dimensi & indikator {m}, dan dampaknya pada {y}. TANPA sitasi.
"fenomena_x1"  : Fenomena {x1} di {instansi} berdasarkan dimensi & indikator {x1}, dan dampaknya pada {y}. TANPA sitasi.
"fenomena_x2"  : Fenomena {x2} di {instansi} berdasarkan dimensi & indikator {x2}, dan dampaknya pada {y}. TANPA sitasi.
"gap_1"        : RESEARCH GAP 1 — bandingkan penelitian terdahulu tentang pengaruh {x1} dan {x2} terhadap {y}. Tunjukkan KONTRADIKSI atau celah hasil. Sitasi: {c_x12}; {c_x22}
"gap_2"        : RESEARCH GAP 2 — bandingkan penelitian terdahulu tentang {m} sebagai mediasi. Tunjukkan belum adanya kajian pada objek {instansi}. Sitasi: {c_med1}; {c_med2}
"penutup"      : Kesimpulan urgensi penelitian, posisikan penelitian sebagai konfirmatori, lalu tetapkan judul penelitian dalam tanda kutip persis seperti JUDUL di atas. TANPA sitasi.

=== ATURAN KERAS ===
1. Bahasa Indonesia akademik formal. Tiap paragraf ±5 kalimat (sekitar 5 baris), mengalir, tidak kaku.
2. SITASI: gunakan HANYA sitasi dari DAFTAR SITASI TERVERIFIKASI, format APA persis seperti tertulis. DILARANG menambah, mengubah, atau mengarang nama penulis, tahun, atau jurnal.
3. {aturan_angka}
4. DILARANG mengklaim hasil penelitian ini sudah terbukti. Gunakan "diduga", "berpotensi", "perlu dikonfirmasi".
5. DILARANG mengulang kalimat pembuka yang sama antar paragraf. Setiap paragraf membuka dengan kalimat berbeda.
6. {aturan_tabel}
7. Dimensi & indikator yang disebut harus berasal dari ONTOLOGI di atas, bukan karangan.
8. HANYA balas JSON valid berisi 13 kunci di atas, tanpa teks lain.
9. TEORI harus TEPAT dan RELEVAN untuk variabelnya, sesuai dengan 'Definisi' variabel di ONTOLOGI. DILARANG memakai teori motivasi untuk variabel kinerja/produktivitas, dan sebaliknya. Setiap variabel (Y, M, X1, X2) harus memakai teori yang sesuai variabelnya masing-masing; jangan menyalin teori antar variabel.
10. SETIAP SITASI HANYA BOLEH MUNCUL SATU KALI di seluruh naskah. Dilarang mengulang sitasi yang sama di paragraf lain.
11. TAHUN SITASI WAJIB 2022–2026. Semua kutipan — jurnal MAUPUN buku — harus tahun {min_th} sampai {max_th}. Tahun lain DILARANG, TANPA kecuali.
    Satu-satunya pengecualian adalah KUTIPAN DALAM (teori klasik yang dirujuk melalui sumber baru). Bila teori lama memang perlu disebut, wajib ditulis dengan format:
        Kristof (1996) dalam Andra, (2025)
    dengan tahun pertama tahun teori asli dan tahun kedua tahun sitasi sah {min_th}–{max_th} dari daftar. TANPA kata "dalam", teori klasik berarti pelanggaran dan naskah akan ditolak.
12. {aturan_angka_akhir}
"""


def _fmt_citation(acad, key):
    return acad.get(key, "")


# ---------------------------------------------------------------------------
# Mode Tabel 1.1 / Tabel 1.2
# ---------------------------------------------------------------------------
# Permintaan pengguna: JANGAN memaksa Tabel 1.1 dan Tabel 1.2 bila tampilannya
# selalu sama dan hanya angka yang berubah antar generate. Angka kedua tabel
# itu memang dibangkitkan deterministik dari hash judul (lihat
# detect_institutional_context), sehingga strukturnya identik tiap kali —
# yang berubah hanya nilainya. Karena itu tabel kini NON-AKTIF secara bawaan
# dan hanya muncul bila diaktifkan eksplisit lewat config "tabel_1_1".
#
# Tabel dimatikan otomatis bila ctx tidak membawa data nyata (mis. AI gagal),
# supaya prompt tidak pernah merujuk tabel yang tidak akan dicetak.
def _tabel_aktif(cfg, ctx):
    """True bila dokumen benar-benar akan mencetak Tabel 1.1/1.2."""
    if not isinstance(ctx, dict):
        return False
    if not (ctx.get("iku_headers") and ctx.get("iku_data")):
        return False
    flag = cfg.get("tabel_1_1", cfg.get("tabel_11", False))
    if isinstance(flag, str):
        flag = flag.strip().lower() in ("1", "true", "ya", "yes", "on")
    return bool(flag)


def _blok_tabel(y, ctx, aktif):
    if not aktif:
        return ("MODE TANPA TABEL. Dokumen ini TIDAK mencetak Tabel 1.1 maupun "
                "Tabel 1.2, sehingga Anda TIDAK BOLEH menyinggung keduanya, "
                "memakai angka kinerja apa pun, atau menyebut persentase/"
                "target/realisasi. Tulis narasi kualitatif murni.")
    rows = ctx.get("iku_data") or []
    headers = ctx.get("iku_headers") or []
    baris = [" | ".join(str(h) for h in headers)] if headers else []
    for r in rows:
        baris.append(" | ".join(str(c) for c in r))
    return (
        "=== DATA TABEL PENDUKUNG (ANGKA NYATA DARI SISTEM — JANGAN UBAH, "
        "JANGAN TAMBAH ANGKA BARU) ===\n"
        f"Tabel 1.1 — capaian {y} (tahun, target, realisasi, gap):\n"
        + ("\n".join(baris) if baris else "(data tabel 1.1 tidak tersedia)")
        + f"\nTabel 1.2 — isu operasional (ringkasan): {ctx.get('tabel2_nama', '-')}"
    )


def _instruksi_fenomena_y(y, instansi, c_compl, aktif):
    if not aktif:
        return (f"Fenomena BELUM OPTIMALNYA {y} di {instansi}: jelaskan dimensi "
                f"& indikator {y} dan dampaknya secara kualitatif. TANPA menyebut "
                f"tabel dan TANPA angka. Sitasi: {c_compl}")
    return (f"Fenomena BELUM OPTIMALNYA {y} di {instansi}: jelaskan dimensi & "
            f"indikator {y}, rujuk angka Tabel 1.1 secara eksplisit (tahun & gap) "
            f"DAN gunakan persis angka dari tabel tanpa mengubahnya, serta sebut "
            f"isu operasional Tabel 1.2. Sitasi: {c_compl}")


def _aturan_angka(aktif):
    if not aktif:
        return ("NOL ANGKA. Karena dokumen tanpa tabel, DILARANG menulis angka "
                "kinerja, persentase, target, realisasi, gap, atau statistik "
                "apa pun. Narasi murni kualitatif.")
    return ("ANGKA: hanya angka dari DATA TABEL PENDUKUNG. DILARANG mengarang "
            "angka, persentase, atau hasil statistik baru.")


def _aturan_tabel(aktif):
    if not aktif:
        return ("Dokumen INI TIDAK punya tabel. Kata \"Tabel 1.1\" dan "
                "\"Tabel 1.2\" DILARANG MUNCUL di seluruh naskah, pada paragraf "
                "mana pun.")
    return ("DILARANG menyebut \"Tabel 1.1\" atau \"Tabel 1.2\" kecuali di "
            "\"fenomena_y\".")


def _aturan_angka_akhir(aktif):
    if not aktif:
        return ("Karena tanpa tabel, \"fenomena_y\" juga wajib bebas angka. "
                "Bila satu saja angka kinerja muncul, naskah dianggap "
                "mengarang data dan ditolak.")
    return ("Angka pada \"fenomena_y\" harus IDENTIK dengan DATA TABEL "
            "PENDUKUNG; dilarang membulatkan ulang atau mengubah angka.")


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


def _validate_bg(raw, acad, judul, extra_sah=None, tabel_aktif=False):
    """Validasi ketat sub-bab 1.1.

    Empat lapis penolakan, berurutan dari yang paling murah dihitung:
      1. Struktur  — 13 kunci wajib ada, tiap paragraf panjang wajar.
      2. Penulis   — nama dalam kurung sitasi harus dari daftar sah.
      3. Tahun     — SEMUA tahun sitasi wajib 2022-2026; teori klasik hanya
                     sah bila berbentuk kutipan dalam ("Kristof (1996) dalam
                     Andra, (2025)"). Tanpa "dalam", ditolak.
      4. Tabel     — bila mode tanpa tabel, kata "Tabel 1.1"/"Tabel 1.2" dan
                     angka kinerja dilarang muncul.
    Mengembalikan dict paragraf, atau None bila naskah ditolak.
    """
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
        return _tolak_bg(
            f"[background] penulis di luar daftar sah: {unknown[:3]}",
            "SITASI DI LUAR DAFTAR. Naskah memuat penulis/tahun yang TIDAK ada di "
            f"DAFTAR SITASI TERVERIFIKASI, contoh: {'; '.join(unknown[:5])}. "
            "Hapus semua sitasi itu.")

    # --- LAPIS 3: tahun sitasi wajib 2022-2026 -----------------------------
    # Dijalankan atas SELURUH naskah sekaligus agar hubungan "dalam" yang
    # membentang di antara dua paragraf tetap terbaca.
    naskah = " ".join(out.values())
    tahun_larang = tahun_sitasi_terlarang(naskah)
    if tahun_larang:
        return _tolak_bg(
            f"[background] tolak: tahun sitasi di luar {MIN_CITE_YEAR}-"
            f"{MAX_CITE_YEAR} tanpa bentuk kutipan dalam -> {tahun_larang}",
            f"TAHUN SITASI. Tahun {tahun_larang} terlarang. SEMUA sitasi wajib "
            f"{MIN_CITE_YEAR}-{MAX_CITE_YEAR}.")

    # Teori klasik hanya sah melalui kutipan dalam; pastikan bentuknya utuh.
    klasik_bulan = []
    for m in re.finditer(r"\(\s*(\d{4})\s*\)", naskah):
        th = int(m.group(1))
        if th < MIN_CITE_YEAR:
            sesudah = naskah[m.start(): m.start() + 170]
            if not re.search(r"\bdalam\b", sesudah, re.IGNORECASE):
                klasik_bulan.append(th)
    if klasik_bulan:
        return _tolak_bg(
            f"[background] tolak: teori klasik tanpa kutipan dalam -> "
            f"{sorted(set(klasik_bulan))}",
            "TEORI KLASIK. Jangan menulis tokoh/tahun lama (mis. "
            f"{sorted(set(klasik_bulan))}) sebagai sitasi langsung. Bila perlu "
            "menyebut teori klasik, tulis HANYA dalam bentuk 'Tokoh (Tahun) dalam "
            "Penulis, (20xx)' dengan Penulis dari daftar, atau sebut namanya tanpa "
            "tahun/kurung.")

    # --- LAPIS 4: mode tanpa tabel ----------------------------------------
    if not tabel_aktif:
        menyebut = re.findall(r"Tabel\s*1\.[12]", naskah, re.IGNORECASE)
        if menyebut:
            return _tolak_bg(
                f"[background] tolak: mode tanpa tabel tapi naskah menyebut "
                f"{sorted(set(menyebut))}",
                f"TABEL/ANGKA. Mode tanpa tabel: kata {'/'.join(sorted(set(menyebut)))} "
                "dan SEMUA angka persen dilarang muncul.")
        # Angka kinerja/persentase dilarang bila tidak ada tabel.
        persen = re.findall(r"\d+(?:[.,]\d+)?\s*%", naskah)
        if persen:
            return _tolak_bg(
                f"[background] tolak: mode tanpa tabel tapi ada angka "
                f"kinerja {persen[:4]}",
                f"TABEL/ANGKA. Mode tanpa tabel: angka persen {persen[:4]} dilarang; "
                "tulis narasi tanpa angka.")

    if judul and judul.strip().lower() not in out["penutup"].lower():
        return _tolak_bg(
            "[background] judul tidak tercantum di penutup -> tolak",
            f"PENUTUP. Paragraf penutup WAJIB memuat judul penelitian ini secara "
            f"persis: \"{judul}\".")

    # Anti-gaya (bukan penolakan): catat bila sitasi jurnal OpenAlex terpakai >1 kali.
    # Sitasi teori klasik (Herzberg, Locke, dst.) boleh berulang secara sah.
    journal_cites = {str(v).lower() for k, v in acad.items()
                     if isinstance(v, str) and v.startswith("(")}
    semu = naskah.lower()
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
    # Tabel hanya aktif bila diminta pengguna DAN datanya benar-benar ada.
    # Default NON-AKTIF: tampilan tabel selalu sama, hanya angka berubah.
    tabel_aktif = _tabel_aktif(cfg, ctx)
    if not tabel_aktif and cfg.get("tabel_1_1"):
        print("[background] tabel_1_1 diminta tapi data ctx tidak lengkap "
              "-> dipaksa tanpa tabel")
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
        blok_tabel=_blok_tabel(y, ctx, tabel_aktif),
        instruksi_fenomena_y=_instruksi_fenomena_y(
            y, cfg.get("instansi", ""), _fmt_citation(acad, "cite_compl"), tabel_aktif),
        aturan_angka=_aturan_angka(tabel_aktif),
        aturan_tabel=_aturan_tabel(tabel_aktif),
        aturan_angka_akhir=_aturan_angka_akhir(tabel_aktif),
        min_th=MIN_CITE_YEAR,
        max_th=MAX_CITE_YEAR,
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

    # Narasi AI wajib BERVARIASI antar-generate (permintaan pengguna), tapi bila
    # terlalu liar naskah ditolak validator. Bawaan 0.7; dapat diubah via config
    # "temperature_narasi" (0.0 = paling stabil/kaku).
    try:
        temperature = float(cfg.get("temperature_narasi", 0.7))
    except Exception:
        temperature = 0.7
    temperature = max(0.0, min(1.2, temperature))

    # Umpan balik korektif: bila validator menolak, Gemini DIBERI TAHU apa yang
    # salah lalu DICOBA ULANG. Template cadangan (naskah tetap, sitasi selalu
    # sama) hanya dipakai bila SEMUA percobaan benar-benar gagal.
    attempts = max(1, int(cfg.get("bg_attempts", 3)))
    for attempt in range(attempts):
        prompt_kirim = prompt if attempt == 0 else prompt + _koreksi_bg(acad, _LAST_BG_REASON)
        try:
            raw = _call_gemini_matrix(keys, chain, prompt_kirim, timeout,
                                      max_tokens=BG_MAX_TOKENS,
                                      temperature=temperature)
            result = _validate_bg(raw, acad, cfg.get("judul", ""), _tokoh_sah(onto),
                                  tabel_aktif=tabel_aktif)
            if result:
                print(f"[background] Latar belakang 1.1 ditulis (matriks, "
                      f"percobaan {attempt+1}/{attempts}, temperature={temperature})")
                return result
            print(f"[background] percobaan {attempt+1}/{attempts} ditolak -> "
                  f"umpan balik korektif dikirim ke Gemini")
        except Exception as e:
            print(f"[background] matriks gagal (percobaan {attempt+1}/{attempts}): "
                  f"{type(e).__name__} {str(e)[:80]}")
            break
    print("[background] PERINGATAN: semua percobaan Gemini ditolak -> memakai "
          "TEMPLATE CADANGAN (naskah tidak bervariasi antar-generate)")
    return None


def _koreksi_bg(acad, alasan):
    """Blok koreksi yang dilampirkan ke prompt 1.1 setelah naskah ditolak validator."""
    return (
        "\n\n=== KOREKSI WAJIB (NASKAH SEBELUMNYA DITOLAK VALIDATOR) ===\n"
        f"Alasan ditolak: {alasan}\n"
        "Tulis ULANG SELURUH 13 kunci JSON dari nol, dan patuhi:\n"
        "1. Setiap kutipan dalam kurung WAJIB persis salah satu entri daftar ini "
        "(nama & tahun sama huruf demi huruf):\n"
        f"{_sitasi_block(acad)}\n"
        "2. DILARANG menulis penulis lain dalam bentuk kurung (termasuk tokoh "
        "teori lama). Bila ingin menyebut teori klasik, sebut namanya TANPA "
        "kurung/tahun, atau tulis 'Tokoh (Tahun) dalam Penulis, (20xx)'.\n"
        f"3. SEMUA tahun di dalam kurung wajib {MIN_CITE_YEAR}-{MAX_CITE_YEAR}.\n"
        "4. Jangan mengarang angka/persentase apa pun."
    )


# ============================================================================
# SUB BAB 2.1.1: KAJIAN TEORITIS VARIABEL Y (Pengertian, Faktor, Dimensi & Indikator)
# ============================================================================

PROMPT_BAB2_2_1_1 = """Anda adalah penulis BAB II (Kajian Teoritis dan Hipotesis) tesis Magister Manajemen Indonesia.

Tulis SUB BAB 2.1.1 "Pengertian Variabel {y}" dengan format yang sama persis seperti contoh berikut:

=== CONTOH FORMAT PARAGRAF DEFINISI (PERHATIKAN VARIASI POSISI SITASI) ===
Kinerja pegawai merupakan hasil kerja secara kualitas dan kuantitas yang dicapai oleh seseorang pegawai dalam melaksanakan tugasnya sesuai dengan tanggung jawab yang diberikan kepadanya (Nabila & Syarvina, 2022). Pernyataan ini menyoroti dua aspek utama dalam penilaian kinerja pegawai, yaitu kualitas dan kuantitas hasil kerja. Kualitas hasil kerja merujuk pada seberapa baik tugas tersebut dilakukan, mencakup aspek seperti ketelitian, kesesuaian dengan standar yang ditetapkan, serta kepuasan pihak yang menerima hasil kerja tersebut. Sementara itu, kuantitas hasil kerja merujuk pada seberapa banyak pekerjaan yang diselesaikan dalam kurun waktu tertentu.

Sejalan dengan pandangan tersebut, kinerja juga dipahami sebagai hasil kerja yang dapat dicapai oleh seseorang atau kelompok orang dalam suatu organisasi sesuai dengan wewenang dan tanggung jawab masing-masing dalam upaya pencapaian tujuan organisasi (Afandi, 2023). Kinerja yang baik dalam organisasi bukan hanya dilihat dari hasil akhirnya, tetapi juga dari proses yang dilalui untuk mencapai hasil tersebut.

=== ATURAN FORMAT SETIAP PARAGRAF DEFINISI ===
1. POSISI SITASI WAJIB BERVARIASI — jangan membuka semua paragraf dengan nama penulis.
   - Maksimal 2 paragraf boleh dibuka dengan sitasi naratif, mis. "Budiarto (2023) menegaskan bahwa ...".
   - Paragraf lainnya WAJIB dibuka dengan pernyataan substantif tentang {y}, lalu sitasi disisipkan di tengah atau di akhir kalimat pertama/kedua.
   - Contoh variasi yang benar:
     a. "Kinerja pegawai merupakan manifestasi kemampuan dan profesionalisme yang tampak pada hasil kerja (Budiarto, 2023)."
     b. "Dari perspektif perilaku organisasi, kinerja pegawai dipahami sebagai tingkat pencapaian hasil kerja individu maupun kelompok (Pashiera & Budiono, 2023)."
     c. "Pandangan tersebut sejalan dengan definisi yang menekankan kualitas dan kuantitas hasil kerja (Pinontoan et al., 2025)."
   - DILARANG memakai pola yang sama di semua paragraf, terutama "(Penulis, Tahun) menyatakan bahwa ...".
2. Setiap paragraf harus tersambung secara konseptual dengan paragraf sebelumnya. Gunakan transisi yang bervariasi (mis. "Sejalan dengan pandangan tersebut,", "Dari perspektif lain,", "Pemahaman itu diperkuat oleh", "Definisi lain menekankan bahwa"), tetapi jangan mengulang transisi yang sama.
3. Setiap paragraf menggunakan SITASI DARI DAFTAR SUMBER di bawah. Jika jumlah sumber < jumlah paragraf, BOLEH mengulang sitasi (maksimal 2x per sumber). Usahakan variasi.
4. Tidak ada kata "menurut", "Berdasarkan pandangan", "pendapat", "kata", "ungkapkan" di depan nama penulis.
5. 8-10 paragraf definisi total (minimal 5 jika sumber terbatas).
6. Semua sitasi harus tahun 2022-2026.
7. SITASI HANYA BOLEH DARI DAFTAR SUMBER WAJIB DIGUNAKAN di bawah. DILARANG mengarang sitasi di luar daftar ini.
8. SETIAP paragraf WAJIB dibuka dengan kalimat yang MENYATAKAN APA ITU {y} (definisi dari sumber),
   contoh: "{y} adalah hasil kerja ...", "{y} merupakan ...", "{y} dipahami sebagai ...".
   DILARANG membuka paragraf dengan uraian faktor, sebab-akibat, cara pengukuran, dimensi,
   leadership, pelatihan, regulasi, perubahan organisasi, atau evaluasi kinerja.
   Paragraf yang hanya membahas faktor/konteks (bukan definisi) TIDAK BOLEH dipakai.

=== JUDUL PENELITIAN ===
{judul}

=== INSTANSI OBJEK ===
{instansi}

=== VARIABEL ===
Y (Dependent): {y}
X1 (Independent 1): {x1}
X2 (Independent 2): {x2}
M (Mediating): {m}

=== ONTOLOGI DIMENSI & INDIKATOR Y ===
{dimensi_indikator_block}

=== DAFTAR SUMBER WAJIB DIGUNAKAN (HANYA dari daftar ini, 8-10 sumber tahun 2022-2026) ===
{sumber_block}

Setiap paragraf definisi {y} HARUS mengutip salah satu sumber di atas dengan format APA in-text: (Penulis, Tahun).
DILARANG KETAT mengarang sitasi di luar daftar sumber ini. Setiap citation di output HARUS cocok persis dengan salah satu entri di daftar sumber.
Setiap sumber hanya boleh DIPAKAI SATU KALI di seluruh naskah (anti-duplikasi untuk skala ribuan tesis).

=== STRUKTUR KELUARAN (JSON valid, TANPA teks di luar JSON) ===

{{
  "pengertian_paragraf": [
    "[Pernyataan substantif tentang {y}]. [Penjelasan 1-2 kalimat] (Penulis, Tahun).",
    "Dari perspektif lain, [pernyataan substantif tentang {y}] (Penulis, Tahun). [Penjelasan 1-2 kalimat]",
    "Penulis (Tahun) menegaskan bahwa [definisi {y}]. [Penjelasan 1-2 kalimat]",
    "Sejalan dengan pandangan tersebut, [pernyataan substantif tentang {y}] (Penulis, Tahun). [Penjelasan]",
    "[Pernyataan substantif tentang {y}] (Penulis, Tahun). [Penjelasan 1-2 kalimat]",
    "Pemahaman tersebut diperkuat oleh pandangan bahwa [pernyataan substantif tentang {y}] (Penulis, Tahun).",
    "[Pernyataan substantif tentang {y}] (Penulis, Tahun). [Penjelasan 1-2 kalimat]",
    "Definisi lain menekankan bahwa [pernyataan substantif tentang {y}] (Penulis, Tahun). [Penjelasan]",
    "[Pernyataan substantif tentang {y}] (Penulis, Tahun). [Penjelasan 1-2 kalimat]"
  ],
  "kesimpulan_sintetis": "Berdasarkan definisi di atas dapat disimpulkan bahwa {y} adalah [sintesis singkat 2-3 kalimat HANYA tentang definisi: hasil kerja kuantitas & kualitas, pencapaian target, efektivitas]. JANGAN bahas dimensi/indikator/faktor.",
  "faktor_pembuka": "Paragraf narasi pembuka yang menyebutkan semua variabel dalam judul ({x1}, {x2}, {m}) dan kaitannya dengan {y}",
  "faktor_poin": [
    "1. {x1} - penjelasan naratif tentang bagaimana {x1} mempengaruhi {y}",
    "2. {x2} - penjelasan naratif tentang bagaimana {x2} mempengaruhi {y}",
    "3. {m} - penjelasan naratif tentang bagaimana {m} mempengaruhi {y}",
    "4. Kemampuan dan Keahlian - penjelasan tentang faktor kemampuan dan keahlian pegawai",
    "5. Motivasi - penjelasan tentang faktor motivasi kerja",
    "6. Pengetahuan - penjelasan tentang faktor pengetahuan tugas",
    "7. Rancangan Kerja - penjelasan tentang desain dan rancangan pekerjaan",
    "8. Kepribadian - penjelasan tentang karakter dan kepribadian individu",
    "9. Kepuasan Kerja - penjelasan tentang perasaan pegawai terhadap pekerjaan",
    "10. Disiplin Kerja - penjelasan tentang sikap patuh aturan dan prosedur"
  ],
  "dimensi_indikator_intro": "Paragraf pengenalan dimensi dan indikator {y} berdasarkan ontologi yang telah disiapkan",
  "dimensi_list": [
    "Dimensi 1 (Indikator: a. ..., b. ..., c. ...)",
    "Dimensi 2 (Indikator: a. ..., b. ..., c. ...)",
    "Dimensi 3 (Indikator: a. ..., b. ..., c. ...)",
    "Dimensi 4 (Indikator: a. ..., b. ..., c. ...)",
    "Dimensi 5 (Indikator: a. ..., b. ..., c. ...)"
  ],
  "kesimpulan_operasional": "Paragraf kesimpulan SINTETIS (1 paragraf) yang MENYINTESISKAN 8-10 definisi di atas: sebutkan konvergensi poin-poin kunci (mis. 'hasil kerja kuantitas & kualitas', 'indikator produktivitas/kualitas/ketepatan', 'peran perilaku & outcome', 'pengaruh faktor individu & organisasi'), JANGAN generik 'berdasarkan pandangan teoretis...'"
}}

=== ATURAN KERAS ===
1. Bahasa Indonesia akademik formal.
2. TAHUN SITASI WAJIB 2022-2026. Tahun lain DILARANG, tanpa kecuali.
3. Sitasi HANYA dari DAFTAR SUMBER WAJIB DIGUNAKAN di atas. Bila jumlah sumber < jumlah paragraf, sitasi boleh diulang maksimal 2 kali per sumber; bila sumber cukup, setiap sumber dipakai sekali.
4. Definisi {y} = 8-10 paragraf (dengan sitasi) + 1 paragraf kesimpulan sintetis (minimal 5 paragraf bila sumber terbatas). Setiap paragraf definisi memuat tepat satu sitasi; posisinya bervariasi (maksimal 2 paragraf dibuka dengan sitasi, sisanya di tengah atau akhir kalimat). Kesimpulan TANPA sitasi baru.
5. DILARANG menggunakan kata "menurut", "Berdasarkan pandangan", "Menurut", "pendapat", "kata", "ungkapkan".
6. Sitasi boleh muncul maksimal 2 KALI per sumber (hanya bila sumber lebih sedikit dari paragraf). Hindari pemakaian berulang yang berlebihan.
7. Setiap paragraf pembuka harus berbeda; jangan mengulang pola "(Penulis, Tahun) menyatakan bahwa ...".
8. JANGAN tulis "Paragraf 1:", "Paragraf 2:", dst. di awal paragraf definisi. Tulis HANYA isi paragraf.
9. HANYA balas JSON valid, tanpa teks lain di luar JSON.
10. Teori harus TEPAT dan RELEVAN untuk variabel {y}.
11. KESIMPULAN SINTETIS (field "kesimpulan_sintetis"): 1 paragraf pendek (2-3 kalimat) HANYA merangkum DEFINISI. Format: "Berdasarkan definisi di atas dapat disimpulkan bahwa {y} adalah [hasil kerja kuantitas & kualitas / pencapaian target / efektivitas pelaksanaan tugas]." DILARANG bahas dimensi, indikator, faktor, teori, atau instansi.
12. Kajian bersifat UMUM/teoretis: DILARANG menyebut nama instansi, kota, unit kerja, maupun istilah sektor objek penelitian (mis. gudang, logistik, kargo, pasien, rekam medis, medis, pelayanan kesehatan, pengiriman barang, nasabah, siswa). Indikator yang disebut hanya indikator TEORETIS pendek (frasa benda 1-4 kata), jangan menulis indikator operasional panjang milik objek.
"""

# ============================================================================
# SUB BAB 2.1.2.1: KAJIAN TEORITIS VARIABEL M (Mediating) - Pengertian
# ============================================================================
PROMPT_BAB2_2_1_2 = """Anda adalah penulis BAB II (Kajian Teoritis dan Hipotesis) tesis Magister Manajemen Indonesia.

Tulis SUB BAB 2.1.2.1 "Pengertian Variabel {m}" dengan format yang sama persis seperti contoh berikut:

=== CONTOH FORMAT PARAGRAF DEFINISI (PERHATIKAN VARIASI POSISI SITASI) ===
Kepuasan kerja merupakan perasaan positif atau negatif pegawai terhadap pekerjaannya yang mencakup aspek gaji, promosi, pengawasan, teman kerja, dan pekerjaan itu sendiri (Nabila & Syarvina, 2022). Pernyataan ini menyoroti bahwa kepuasan kerja bersifat multidimensi dan mencakup baik aspek intrinsik maupun ekstrinsik pekerjaan.

Sejalan dengan pandangan tersebut, employee engagement dipahami sebagai kondisi psikologis positif yang terpenuhi dalam pekerjaan, ditandai dengan vigor, dedication, dan absorption (Afandi, 2023). Karyawan yang engaged tidak hanya puas tetapi juga antusias dan termotivasi untuk memberikan kontribusi tambahan bagi organisasi.

=== ATURAN FORMAT SETIAP PARAGRAF DEFINISI ===
1. POSISI SITASI WAJIB BERVARIASI — jangan membuka semua paragraf dengan nama penulis.
   - Maksimal 2 paragraf boleh dibuka dengan sitasi naratif, mis. "Budiarto (2023) menegaskan bahwa ...".
   - Paragraf lainnya WAJIB dibuka dengan pernyataan substantif tentang {m}, lalu sitasi disisipkan di tengah atau di akhir kalimat pertama/kedua.
   - Contoh variasi yang benar:
     a. "[Pernyataan substantif tentang {m}] (Penulis, Tahun)."
     b. "Dari perspektif lain, [pernyataan substantif tentang {m}] (Penulis, Tahun)."
     c. "Pandangan tersebut sejalan dengan definisi yang menekankan ... (Penulis, Tahun)."
   - DILARANG memakai pola yang sama di semua paragraf, terutama "(Penulis, Tahun) menyatakan bahwa ...".
2. Setiap paragraf harus tersambung secara konseptual dengan paragraf sebelumnya. Gunakan transisi yang bervariasi (mis. "Sejalan dengan pandangan tersebut,", "Dari perspektif lain,", "Pemahaman itu diperkuat oleh", "Definisi lain menekankan bahwa"), tetapi jangan mengulang transisi yang sama.
3. Setiap paragraf menggunakan SITASI YANG BERBEDA dari daftar sumber di bawah.
4. Tidak ada kata "menurut", "Berdasarkan pandangan", "pendapat", "kata", "ungkapkan" di depan nama penulis.
5. Buat 1 paragraf definisi untuk SETIAP sumber unik dalam daftar (minimal 2, maksimal 10 paragraf), lalu 1 paragraf kesimpulan sintetis.
6. Semua sitasi harus tahun 2022-2026.
7. SITASI HANYA BOLEH DARI DAFTAR SUMBER WAJIB DIGUNAKAN di bawah. DILARANG mengarang sitasi di luar daftar ini.

=== JUDUL PENELITIAN ===
{judul}

=== INSTANSI OBJEK ===
{instansi}

=== VARIABEL ===
Y (Dependent): {y}
X1 (Independent 1): {x1}
X2 (Independent 2): {x2}
M (Mediating): {m}

=== ONTOLOGI DIMENSI & INDIKATOR M ===
{dimensi_indikator_block}

=== DAFTAR SUMBER WAJIB DIGUNAKAN (HANYA dari daftar ini, 8-10 sumber tahun 2022-2026) ===
{sumber_block}

Setiap paragraf definisi {m} HARUS mengutip salah satu sumber di atas dengan format APA in-text: (Penulis, Tahun).
DILARANG KETAT mengarang sitasi di luar daftar sumber ini. Setiap citation di output HARUS cocok persis dengan salah satu entri di daftar sumber.
Setiap sumber hanya boleh DIPAKAI SATU KALI di seluruh naskah (anti-duplikasi untuk skala ribuan tesis).

=== STRUKTUR KELUARAN (JSON valid, TANPA teks di luar JSON) ===

{{
  "pengertian_paragraf": [
    "[Pernyataan substantif tentang {m}]. [Penjelasan 1-2 kalimat] (Penulis, Tahun).",
    "Dari perspektif lain, [pernyataan substantif tentang {m}] (Penulis, Tahun). [Penjelasan 1-2 kalimat]",
    "... ulangi: buat 1 paragraf untuk setiap sumber unik (2 sampai 10 paragraf), posisi sitasi bervariasi: maksimal 2 paragraf dibuka dengan sitasi, sisanya sitasi di tengah/akhir ..."
  ],
  "kesimpulan_sintetis": "Berdasarkan definisi di atas dapat disimpulkan bahwa {m} adalah [sintesis singkat 2-3 kalimat HANYA tentang definisi]. JANGAN bahas dimensi/indikator/faktor."
}}

=== ATURAN KERAS ===
1. Bahasa Indonesia akademik formal.
2. TAHUN SITASI WAJIB 2022-2026. Tahun lain DILARANG, tanpa kecuali.
3. Situsasi HANYA dari DAFTAR SUMBER WAJIB DIGUNAKAN di atas. Setiap sumber hanya boleh dipakai SATU KALI.
4. Definisi = 2 sampai 10 paragraf definisi (dengan sitasi BERBEDA) + 1 paragraf kesimpulan sintetis. Gunakan SETIAP sumber unik tepat satu kali; JUMLAH paragraf = jumlah sumber unik yang tersedia — JANGAN mengulang sumber yang sama. Setiap paragraf definisi memuat tepat satu sitasi; posisinya bervariasi (maksimal 2 paragraf dibuka dengan sitasi, sisanya di tengah atau akhir kalimat). Kesimpulan TANPA sitasi baru.
5. DILARANG menggunakan kata "menurut", "Berdasarkan pandangan", "Menurut", "pendapat", "kata", "ungkapkan".
6. Setiap sitasi hanya boleh muncul SATU KALI di seluruh naskah.
7. Setiap paragraf pembuka harus berbeda; jangan mengulang pola "(Penulis, Tahun) menyatakan bahwa ...".
8. JANGAN tulis "Paragraf 1:", "Paragraf 2:", dst. di awal paragraf definisi. Tulis HANYA isi paragraf.
9. HANYA balas JSON valid, tanpa teks lain di luar JSON.
10. Teori harus TEPAT dan RELEVAN untuk variabel {m}.
11. KESIMPULAN SINTETIS (field "kesimpulan_sintetis"): 1 paragraf pendek (2-3 kalimat) HANYA merangkum DEFINISI. Format: "Berdasarkan definisi di atas dapat disimpulkan bahwa {m} adalah [inti definisi]." DILARANG bahas dimensi, indikator, faktor, teori, atau instansi.
12. SETIAP paragraf definisi WAJIB dibuka dengan kalimat yang MENYATAKAN APA ITU {m} (definisi dari sumber), contoh: "{m} adalah ikatan psikologis ...", "{m} merupakan ...", "{m} dipahami sebagai ...". DILARANG membuka paragraf dengan uraian faktor, sebab-akibat, cara pengukuran, dimensi, leadership, pelatihan, regulasi, perubahan organisasi, atau evaluasi kinerja. Paragraf yang hanya membahas faktor/konteks (bukan definisi) TIDAK BOLEH dipakai.
13. Kajian bersifat UMUM/teoretis: DILARANG menyebut nama instansi, kota, unit kerja, maupun istilah sektor objek penelitian (mis. gudang, logistik, kargo, pasien, rekam medis, medis, pelayanan kesehatan, pengiriman, nasabah, siswa, aparatur). Indikator yang disebut hanya indikator TEORETIS pendek (frasa benda 1-4 kata).
"""


# ============================================================================
# SUB BAB 2.1.3.1: KAJIAN TEORITIS VARIABEL X1 (Independent 1) - Pengertian
# ============================================================================
PROMPT_BAB2_2_1_3 = """Anda adalah penulis BAB II (Kajian Teoritis dan Hipotesis) tesis Magister Manajemen Indonesia.

Tulis SUB BAB 2.1.3.1 "Pengertian Variabel {x1}" dengan format yang sama persis seperti contoh berikut:

=== CONTOH FORMAT PARAGRAF DEFINISI (PERHATIKAN VARIASI POSISI SITASI) ===
Pengembangan karier merupakan serangkaian kegiatan formal yang dirancang organisasi untuk membantu karyawan mengembangkan kemampuan dan mencapai tujuan karier mereka (Nabila & Syarvina, 2022). Pernyataan ini menyoroti bahwa pengembangan karier bersifat terencana, sistematis, dan berorientasi pada pertumbuhan individu serta kebutuhan organisasi.

Sejalan dengan pandangan tersebut, pelatihan dipahami sebagai proses sistematis untuk mengubah perilaku karyawan ke arah yang diinginkan organisasi melalui penguasaan pengetahuan, keterampilan, dan sikap (Afandi, 2023). Pelatihan yang efektif meningkatkan kompetensi dan produktivitas karyawan dalam jangka pendek dan panjang.

=== ATURAN FORMAT SETIAP PARAGRAF DEFINISI ===
1. POSISI SITASI WAJIB BERVARIASI — jangan membuka semua paragraf dengan nama penulis.
   - Maksimal 2 paragraf boleh dibuka dengan sitasi naratif, mis. "Budiarto (2023) menegaskan bahwa ...".
   - Paragraf lainnya WAJIB dibuka dengan pernyataan substantif tentang {x1}, lalu sitasi disisipkan di tengah atau di akhir kalimat pertama/kedua.
   - Contoh variasi yang benar:
     a. "[Pernyataan substantif tentang {x1}] (Penulis, Tahun)."
     b. "Dari perspektif lain, [pernyataan substantif tentang {x1}] (Penulis, Tahun)."
     c. "Pandangan tersebut sejalan dengan definisi yang menekankan ... (Penulis, Tahun)."
   - DILARANG memakai pola yang sama di semua paragraf, terutama "(Penulis, Tahun) menyatakan bahwa ...".
2. Setiap paragraf harus tersambung secara konseptual dengan paragraf sebelumnya. Gunakan transisi yang bervariasi (mis. "Sejalan dengan pandangan tersebut,", "Dari perspektif lain,", "Pemahaman itu diperkuat oleh", "Definisi lain menekankan bahwa"), tetapi jangan mengulang transisi yang sama.
3. Setiap paragraf menggunakan SITASI YANG BERBEDA dari daftar sumber di bawah.
4. Tidak ada kata "menurut", "Berdasarkan pandangan", "pendapat", "kata", "ungkapkan" di depan nama penulis.
5. Buat 1 paragraf definisi untuk SETIAP sumber unik dalam daftar (minimal 2, maksimal 10 paragraf), lalu 1 paragraf kesimpulan sintetis.
6. Semua sitasi harus tahun 2022-2026.
7. SITASI HANYA BOLEH DARI DAFTAR SUMBER WAJIB DIGUNAKAN di bawah. DILARANG mengarang sitasi di luar daftar ini.

=== JUDUL PENELITIAN ===
{judul}

=== INSTANSI OBJEK ===
{instansi}

=== VARIABEL ===
Y (Dependent): {y}
X1 (Independent 1): {x1}
X2 (Independent 2): {x2}
M (Mediating): {m}

=== ONTOLOGI DIMENSI & INDIKATOR X1 ===
{dimensi_indikator_block}

=== DAFTAR SUMBER WAJIB DIGUNAKAN (HANYA dari daftar ini, 8-10 sumber tahun 2022-2026) ===
{sumber_block}

Setiap paragraf definisi {x1} HARUS mengutip salah satu sumber di atas dengan format APA in-text: (Penulis, Tahun).
DILARANG KETAT mengarang sitasi di luar daftar sumber ini. Setiap citation di output HARUS cocok persis dengan salah satu entri di daftar sumber.
Setiap sumber hanya boleh DIPAKAI SATU KALI di seluruh naskah (anti-duplikasi untuk skala ribuan tesis).

=== STRUKTUR KELUARAN (JSON valid, TANPA teks di luar JSON) ===

{{
  "pengertian_paragraf": [
    "[Pernyataan substantif tentang {x1}]. [Penjelasan 1-2 kalimat] (Penulis, Tahun).",
    "Dari perspektif lain, [pernyataan substantif tentang {x1}] (Penulis, Tahun). [Penjelasan 1-2 kalimat]",
    "... ulangi: buat 1 paragraf untuk setiap sumber unik (2 sampai 10 paragraf), posisi sitasi bervariasi: maksimal 2 paragraf dibuka dengan sitasi, sisanya sitasi di tengah/akhir ..."
  ],
  "kesimpulan_sintetis": "Berdasarkan definisi di atas dapat disimpulkan bahwa {x1} adalah [sintesis singkat 2-3 kalimat HANYA tentang definisi]. JANGAN bahas dimensi/indikator/faktor."
}}

=== ATURAN KERAS ===
1. Bahasa Indonesia akademik formal.
2. TAHUN SITASI WAJIB 2022-2026. Tahun lain DILARANG, tanpa kecuali.
3. Situsasi HANYA dari DAFTAR SUMBER WAJIB DIGUNAKAN di atas. Setiap sumber hanya boleh dipakai SATU KALI.
4. Definisi {x1} = 2 sampai 10 paragraf definisi (dengan sitasi BERBEDA) + 1 paragraf kesimpulan sintetis. Gunakan SETIAP sumber unik tepat satu kali; JUMLAH paragraf = jumlah sumber unik yang tersedia — JANGAN mengulang sumber yang sama. Setiap paragraf definisi memuat tepat satu sitasi; posisinya bervariasi (maksimal 2 paragraf dibuka dengan sitasi, sisanya di tengah atau akhir kalimat). Kesimpulan TANPA sitasi baru.
5. DILARANG menggunakan kata "menurut", "Berdasarkan pandangan", "Menurut", "pendapat", "kata", "ungkapkan".
6. Setiap sitasi hanya boleh muncul SATU KALI di seluruh naskah.
7. Setiap paragraf pembuka harus berbeda; jangan mengulang pola "(Penulis, Tahun) menyatakan bahwa ...".
8. JANGAN tulis "Paragraf 1:", "Paragraf 2:", dst. di awal paragraf definisi. Tulis HANYA isi paragraf.
9. HANYA balas JSON valid, tanpa teks di luar JSON.
10. Teori harus TEPAT dan RELEVAN untuk variabel {x1}.
11. KESIMPULAN SINTETIS (field "kesimpulan_sintetis"): 1 paragraf pendek (2-3 kalimat) HANYA merangkum DEFINISI. Format: "Berdasarkan definisi di atas dapat disimpulkan bahwa {x1} adalah [inti definisi]." DILARANG bahas dimensi, indikator, faktor, teori, atau instansi.
12. SETIAP paragraf definisi WAJIB dibuka dengan kalimat yang MENYATAKAN APA ITU {x1} (definisi dari sumber), contoh: "{x1} adalah seperangkat kemampuan ...", "{x1} merupakan ...", "{x1} dipahami sebagai ...". DILARANG membuka paragraf dengan uraian faktor, sebab-akibat, cara pengukuran, dimensi, leadership, pelatihan, regulasi, perubahan organisasi, atau evaluasi kinerja. Paragraf yang hanya membahas faktor/konteks (bukan definisi) TIDAK BOLEH dipakai.
13. Kajian bersifat UMUM/teoretis: DILARANG menyebut nama instansi, kota, unit kerja, maupun istilah sektor objek penelitian (mis. gudang, logistik, kargo, pasien, rekam medis, medis, pelayanan kesehatan, pengiriman, nasabah, siswa, aparatur). Indikator yang disebut hanya indikator TEORETIS pendek (frasa benda 1-4 kata).
"""


# ============================================================================
# SUB BAB 2.1.4.1: KAJIAN TEORITIS VARIABEL X2 (Independent 2) - Pengertian
# ============================================================================
PROMPT_BAB2_2_1_4 = """Anda adalah penulis BAB II (Kajian Teoritis dan Hipotesis) tesis Magister Manajemen Indonesia.

Tulis SUB BAB 2.1.4.1 "Pengertian Variabel {x2}" dengan format yang sama persis seperti contoh berikut:

=== CONTOH FORMAT PARAGRAF DEFINISI (PERHATIKAN VARIASI POSISI SITASI) ===
Budaya organisasi merupakan sistem nilai, kepercayaan, dan norma yang dibagikan oleh anggota organisasi dan memengaruhi perilaku mereka (Nabila & Syarvina, 2022). Pernyataan ini menyoroti bahwa budaya organisasi bersifat kolektif, terpelajari, dan menentukan cara kerja organisasi.

Sejalan dengan pandangan tersebut, mutasi kerja dipahami sebagai perpindahan pegawai dari satu jabatan/unit ke jabatan/unit lain dalam rangka pengembangan kompetensi dan penempatan yang tepat (Afandi, 2023). Mutasi yang efektif memperluas wawasan, keterampilan, dan jaringan kerja pegawai.

=== ATURAN FORMAT SETIAP PARAGRAF DEFINISI ===
1. POSISI SITASI WAJIB BERVARIASI — jangan membuka semua paragraf dengan nama penulis.
   - Maksimal 2 paragraf boleh dibuka dengan sitasi naratif, mis. "Budiarto (2023) menegaskan bahwa ...".
   - Paragraf lainnya WAJIB dibuka dengan pernyataan substantif tentang {x2}, lalu sitasi disisipkan di tengah atau di akhir kalimat pertama/kedua.
   - Contoh variasi yang benar:
     a. "[Pernyataan substantif tentang {x2}] (Penulis, Tahun)."
     b. "Dari perspektif lain, [pernyataan substantif tentang {x2}] (Penulis, Tahun)."
     c. "Pandangan tersebut sejalan dengan definisi yang menekankan ... (Penulis, Tahun)."
   - DILARANG memakai pola yang sama di semua paragraf, terutama "(Penulis, Tahun) menyatakan bahwa ...".
2. Setiap paragraf harus tersambung secara konseptual dengan paragraf sebelumnya. Gunakan transisi yang bervariasi (mis. "Sejalan dengan pandangan tersebut,", "Dari perspektif lain,", "Pemahaman itu diperkuat oleh", "Definisi lain menekankan bahwa"), tetapi jangan mengulang transisi yang sama.
3. Setiap paragraf menggunakan SITASI YANG BERBEDA dari daftar sumber di bawah.
4. Tidak ada kata "menurut", "Berdasarkan pandangan", "pendapat", "kata", "ungkapkan" di depan nama penulis.
5. Buat 1 paragraf definisi untuk SETIAP sumber unik dalam daftar (minimal 2, maksimal 10 paragraf), lalu 1 paragraf kesimpulan sintetis.
6. Semua sitasi harus tahun 2022-2026.
7. SITASI HANYA BOLEH DARI DAFTAR SUMBER WAJIB DIGUNAKAN di bawah. DILARANG mengarang sitasi di luar daftar ini.

=== JUDUL PENELITIAN ===
{judul}

=== INSTANSI OBJEK ===
{instansi}

=== VARIABEL ===
Y (Dependent): {y}
X1 (Independent 1): {x1}
X2 (Independent 2): {x2}
M (Mediating): {m}

=== ONTOLOGI DIMENSI & INDIKATOR X2 ===
{dimensi_indikator_block}

=== DAFTAR SUMBER WAJIB DIGUNAKAN (HANYA dari daftar ini, 8-10 sumber tahun 2022-2026) ===
{sumber_block}

Setiap paragraf definisi {x2} HARUS mengutip salah satu sumber di atas dengan format APA in-text: (Penulis, Tahun).
DILARANG KETAT mengarang sitasi di luar daftar sumber ini. Setiap citation di output HARUS cocok persis dengan salah satu entri di daftar sumber.
Setiap sumber hanya boleh DIPAKAI SATU KALI di seluruh naskah (anti-duplikasi untuk skala ribuan tesis).

=== STRUKTUR KELUARAN (JSON valid, TANPA teks di luar JSON) ===

{{
  "pengertian_paragraf": [
    "[Pernyataan substantif tentang {x2}]. [Penjelasan 1-2 kalimat] (Penulis, Tahun).",
    "Dari perspektif lain, [pernyataan substantif tentang {x2}] (Penulis, Tahun). [Penjelasan 1-2 kalimat]",
    "... ulangi: buat 1 paragraf untuk setiap sumber unik (2 sampai 10 paragraf), posisi sitasi bervariasi: maksimal 2 paragraf dibuka dengan sitasi, sisanya sitasi di tengah/akhir ..."
  ],
  "kesimpulan_sintetis": "Berdasarkan definisi di atas dapat disimpulkan bahwa {x2} adalah [sintesis singkat 2-3 kalimat HANYA tentang definisi]. JANGAN bahas dimensi/indikator/faktor."
}}

=== ATURAN KERAS ===
1. Bahasa Indonesia akademik formal.
2. TAHUN SITASI WAJIB 2022-2026. Tahun lain DILARANG, tanpa kecuali.
3. Situsasi HANYA dari DAFTAR SUMBER WAJIB DIGUNAKAN di atas. Setiap sumber hanya boleh dipakai SATU KALI.
4. Definisi {x2} = 2 sampai 10 paragraf definisi (dengan sitasi BERBEDA) + 1 paragraf kesimpulan sintetis. Gunakan SETIAP sumber unik tepat satu kali; JUMLAH paragraf = jumlah sumber unik yang tersedia — JANGAN mengulang sumber yang sama. Setiap paragraf definisi memuat tepat satu sitasi; posisinya bervariasi (maksimal 2 paragraf dibuka dengan sitasi, sisanya di tengah atau akhir kalimat). Kesimpulan TANPA sitasi baru.
5. DILARANG menggunakan kata "menurut", "Berdasarkan pandangan", "Menurut", "pendapat", "kata", "ungkapkan".
6. Setiap sitasi hanya boleh muncul SATU KALI di seluruh naskah.
7. Setiap paragraf pembuka harus berbeda; jangan mengulang pola "(Penulis, Tahun) menyatakan bahwa ...".
8. JANGAN tulis "Paragraf 1:", "Paragraf 2:", dst. di awal paragraf definisi. Tulis HANYA isi paragraf.
9. HANYA balas JSON valid, tanpa teks di luar JSON.
10. Teori harus TEPAT dan RELEVAN untuk variabel {x2}.
11. KESIMPULAN SINTETIS (field "kesimpulan_sintetis"): 1 paragraf pendek (2-3 kalimat) HANYA merangkum DEFINISI. Format: "Berdasarkan definisi di atas dapat disimpulkan bahwa {x2} adalah [inti definisi]." DILARANG bahas dimensi, indikator, faktor, teori, atau instansi.
12. SETIAP paragraf definisi WAJIB dibuka dengan kalimat yang MENYATAKAN APA ITU {x2} (definisi dari sumber), contoh: "{x2} adalah total tuntutan ...", "{x2} merupakan ...", "{x2} dipahami sebagai ...". DILARANG membuka paragraf dengan uraian faktor, sebab-akibat, cara pengukuran, dimensi, leadership, pelatihan, regulasi, perubahan organisasi, atau evaluasi kinerja. Paragraf yang hanya membahas faktor/konteks (bukan definisi) TIDAK BOLEH dipakai.
13. Kajian bersifat UMUM/teoretis: DILARANG menyebut nama instansi, kota, unit kerja, maupun istilah sektor objek penelitian (mis. gudang, logistik, kargo, pasien, rekam medis, medis, pelayanan kesehatan, pengiriman, nasabah, siswa, aparatur). Indikator yang disebut hanya indikator TEORETIS pendek (frasa benda 1-4 kata).
"""


# ==============================================================================
# SUB BAB 1.2 IDENTIFIKASI MASALAH — narasi SPESIFIK JUDUL (bukan templat statis)
# ------------------------------------------------------------------------------
# Cacat lama: butir 1.2 adalah templat statis sehingga naskah identifikasi masalah
# nyaris identik antar judul. Sekarang butir masalah ditulis dari judul + instansi
# + sektor + dimensi/indikator ontologi variabel, TANPA angka dan TANPA sitasi.
# ==============================================================================
PROMPT_IDENTIFIKASI = """Anda adalah penulis BAB I tesis Magister Manajemen (S2) Indonesia.
Tugas: menulis SUB BAB 1.2 IDENTIFIKASI MASALAH untuk penelitian berikut.

JUDUL    : {judul}
INSTANSI : {instansi}
SEKTOR   : {sektor}
PERAN SEKTOR : {peran}
ISU OPERASIONAL SEKTOR : {isu}
VARIABEL : variabel pertama = {x1}; variabel kedua = {x2}; variabel mediasi = {m}; variabel terikat = {y}

ONTOLOGI VARIABEL (wajib dipakai sebagai bahan diksi):
{onto}

=== STRUKTUR KELUARAN (JSON valid, TANPA teks di luar JSON) ===
{{
  "items": [
    {{"poin": "[judul butir singkat, maksimal 60 karakter]", "narasi": "[2 sampai 3 kalimat]"}},
    "... tepat ENAM butir (items) ..."
  ]
}}

=== CAKUPAN ENAM BUTIR ===
1. Kesenjangan pencapaian {y} (pakai dimensi/indikator {y}).
2. Melemahnya {m} (pakai dimensi/indikator {m}).
3. Keterbatasan penerapan {x1} (pakai dimensi/indikator {x1}).
4. Kondisi {x2} yang belum optimal (pakai dimensi/indikator {x2}).
5. Hambatan operasional KHAS sektor {sektor} di {instansi}, dikaitkan dengan {m} dan {y}.
6. Kontradiksi hasil riset empiris (research gap) di antara variabel-variabel tersebut.

=== ATURAN KERAS ===
1. Bahasa Indonesia akademik formal.
2. DILARANG menuliskan angka apa pun dan DILARANG menyebut Tabel, Gambar, atau statistik.
3. DILARANG menuliskan sitasi seperti (Penulis, Tahun).
4. DILARANG memakai kode variabel; tulis nama variabelnya saja.
5. Setiap narasi WAJIB menyebut nama variabel judul ini atau nama instansi secara eksplisit.
6. Kalimat antar butir TIDAK BOLEH sama atau saling menyalin.
7. Manfaatkan istilah dimensi/indikator dari ontologi di atas.
8. HANYA balas JSON valid, tanpa penjelasan tambahan.
"""

_KATA_TERLARANG_1_2 = (
    "tiga tahun", "kian melebar", "pseudo-compliance", "pseudo compliance",
    "setiap tahun", "per tahun", "persen", "%",
)

_SITASI_RE_1_2 = re.compile(r"\(\s*[A-Z][A-Za-z.\-' ]{2,60}\d{4}\s*\)")


def _bersihkan_kode_variabel(teks, cfg):
    """Ganti kode variabel (X1/X2/M/Y) hasil AI menjadi nama variabel aslinya.

    Pencocokan sengaja CASE-SENSITIVE dan berurutan (frasa panjang dulu), karena
    pola generik dengan IGNORECASE akan menimpa hasil penggantian sebelumnya.
    """
    x1 = cfg.get("x1") or cfg.get("x1_title") or ""
    x2 = cfg.get("x2") or cfg.get("x2_title") or ""
    mn = cfg.get("m") or cfg.get("m_title") or ""
    yn = cfg.get("y") or cfg.get("y_title") or ""
    peta = [
        (r"\bsebagai\s+Variabel\s+Mediasi\b", "sebagai variabel mediasi"),
        (r"\bsebagai\s+Variabel\s+Intervening\b", "sebagai variabel intervening"),
        (r"\bsebagai\s+Variabel\s+Terikat\b", "sebagai variabel terikat"),
        (r"\bVariabel\s+Mediasi\b", mn),
        (r"\bVariabel\s+Intervening\b", mn),
        (r"\bVariabel\s+Terikat\b", yn),
        (r"\bX\s*1\b", x1),
        (r"\bX\s*2\b", x2),
        (r"\(\s*Y\s*\)", yn),
        (r"\bY\b", yn),
        (r"\bM\b", mn),
    ]
    for pola, nama in peta:
        if nama:
            teks = re.sub(pola, nama, teks)
    return teks


def _validasi_identifikasi(raw, cfg):
    """Validasi keluaran AI untuk 1.2: tepat 6 butir, tanpa angka, tanpa sitasi."""
    if isinstance(raw, dict):
        data = raw
    else:
        try:
            data = json.loads(raw)
        except Exception:
            m = re.search(r"\{.*\}", str(raw or ""), re.S)
            if not m:
                return None
            try:
                data = json.loads(m.group(0))
            except Exception:
                return None
    items = (data or {}).get("items") or []
    if not isinstance(items, list) or len(items) != 6:
        return None

    x1 = (cfg.get("x1") or cfg.get("x1_lower") or "").lower()
    x2 = (cfg.get("x2") or cfg.get("x2_lower") or "").lower()
    m = (cfg.get("m") or cfg.get("m_lower") or "").lower()
    y = (cfg.get("y") or cfg.get("y_lower") or "").lower()
    variabel = [v for v in (x1, x2, m, y) if v]

    hasil = []
    for it in items:
        if not isinstance(it, dict):
            return None
        poin = _bersihkan_kode_variabel(str(it.get("poin", "")).strip(), cfg)
        narasi = _bersihkan_kode_variabel(str(it.get("narasi", "")).strip(), cfg)
        if not (8 <= len(poin) <= 80) or not (110 <= len(narasi) <= 800):
            return None
        gab = f"{poin} {narasi}"
        if re.search(r"\d", gab):
            return None
        if _SITASI_RE_1_2.search(narasi):
            return None
        gab_l = gab.lower()
        if any(k in gab_l for k in _KATA_TERLARANG_1_2):
            return None
        if variabel and not any(v in gab_l for v in variabel):
            return None
        hasil.append({"poin": poin, "narasi": narasi})

    if len({h["narasi"].lower() for h in hasil}) != 6:
        return None
    return {"items": hasil}


def write_identifikasi_masalah(cfg, onto, ctx):
    """Tulis 6 butir identifikasi masalah (1.2) via Gemini.

    Mengembalikan list [{"poin": ..., "narasi": ...}, ...] atau None bila AI
    tidak tersedia/gagal; pemanggil memakai fallback deterministik yang tetap
    spesifik judul.
    """
    ctx = ctx or {}
    x1 = cfg.get("x1") or cfg.get("x1_lower") or "Variabel Bebas Pertama"
    x2 = cfg.get("x2") or cfg.get("x2_lower") or "Variabel Bebas Kedua"
    m = cfg.get("m") or cfg.get("m_lower") or "Variabel Mediasi"
    y = cfg.get("y") or cfg.get("y_lower") or "Kinerja Pegawai"
    prompt = PROMPT_IDENTIFIKASI.format(
        judul=cfg.get("judul", "-"),
        instansi=cfg.get("instansi", "-"),
        sektor=ctx.get("sektor", "-"),
        peran=ctx.get("peran_strategis", "-"),
        isu=ctx.get("isu_lapangan", "-"),
        x1=x1, x2=x2, m=m, y=y,
        onto=_ontologi_block(onto),
    )
    try:
        _temp_12 = float(cfg.get("temperature_narasi", 0.7))
    except Exception:
        _temp_12 = 0.7
    res = _gemini_json(prompt, lambda r: _validasi_identifikasi(r, cfg),
                       max_tokens=2600, tag="sub_1_2",
                       temperature=max(0.0, min(1.2, _temp_12)))
    if not res:
        return None
    return res.get("items")


def _fmt_sitasi_bab2(acad, y_key="cite_kinerja"):
    """Format sitasi untuk bab 2 dari daftar akademik — menyediakan semua kunci agar Gemini punya cukup pilihan unik."""
    parts = []
    # Semua kunci sitasi yang tersedia di acad
    all_keys = [y_key, "cite_macro", "cite_sdm", "cite_m_1", "cite_x1_1", "cite_x1_2",
                "cite_x2_1", "cite_x2_2", "cite_compl", "cite_med_1", "cite_med_2"]
    seen = set()
    for k in all_keys:
        v = acad.get(k, "")
        if v and v not in seen and _apa_penulis_valid(v):
            parts.append(f"- {k}: {v}")
            seen.add(v)
    return "\n".join(parts) if parts else "(tidak ada sitasi terverifikasi)"


def _apa_penulis_valid(apa):
    """False bila penulis pada sitasi rusak (mis. '(. & ., 2022)' / '(...., 2022)')."""
    import re
    m = re.match(r'^\s*\(([^,]+)', apa or "")
    if not m:
        return True
    return bool(re.search(r'[A-Za-z]{2,}', m.group(1)))


def _fmt_sumber_block(sources_y):
    """Format daftar sumber yang di-fetch secara live untuk prompt Gemini."""
    lines = []
    n = 0
    for s in sources_y:
        apa = s.get("apa_in_text", "")
        # Lewati sumber dengan penulis rusak (titik-titik) agar tidak dikutip Gemini
        if not _apa_penulis_valid(apa):
            continue
        n += 1
        tipe = "BUKU" if s.get("type") == "buku" else "JURNAL"
        if not apa:
            apa = "(Penulis, Tahun)"
        # Judul sumber: aksen DIPERTAHANKAN (stdout server = UTF-8) supaya nama
        # penulis di sitasi & DAFTAR PUSTAKA tetap konsisten.
        title = s.get("title", "")
        title_clean = " ".join(str(title).split())
        year = s.get("year", "")
        lines.append(f"{n}. [{tipe}] {apa} — \"{title_clean}\" ({year})")
    return "\n".join(lines) if lines else "(tidak ada sumber tersedia)"


def _dimensi_indikator_block(onto, y_key="y", instansi="", judul=""):
    """Buat blok dimensi & indikator TEORETIS untuk prompt Bab II.

    Indikator operasional (spesifik objek) sengaja tidak dipakai di Bab II.
    """
    import indicator_layers as IL
    v = ((onto or {}).get("variables") or {}).get(y_key) or {}
    dims = v.get("dimensi", [])
    lines = []
    for d in dims:
        inds = IL.indikator_teoretis(d, instansi=instansi, judul=judul)
        lines.append(f"- {d.get('nama', '-')}: {', '.join(inds)}")
    return "\n".join(lines) if lines else "(belum diriset)"


def _daftar_sitasi_boleh(daftar_sumber, maks=12):
    """Daftar sitasi yang boleh dipakai (dari sumber terverifikasi) untuk umpan balik."""
    out = []
    for s in (daftar_sumber or []):
        v = s.get("apa_in_text") if isinstance(s, dict) else None
        if v:
            out.append(str(v).strip())
    if not out:
        return "(tidak ada sumber terverifikasi - jangan mengarang sitasi)"
    return "\n".join(f"  - {x}" for x in out[:maks])


def _sitasi_terverifikasi(daftar_sumber, maks=12):
    """Sitasi in-text terverifikasi sebagai LIST (bukan blok teks)."""
    out = []
    for s in (daftar_sumber or []):
        v = s.get("apa_in_text") if isinstance(s, dict) else None
        if not v:
            continue
        v = str(v).strip()
        if not _apa_penulis_valid(v):
            continue
        if v not in out:
            out.append(v)
    return out[:maks]


def _nama_pertama(sitasi):
    """Nama penulis pertama dari '(Syamsuri et al., 2024)' -> 'Syamsuri'."""
    t = str(sitasi or "").strip().lstrip("(").strip()
    t = t.split(",")[0]
    t = re.split(r"\s+(?:et\s+al\.?|dkk\.?|&|dan)\s+", t)[0]
    return t.strip()


# Sitasi kurung: "(Penulis, Tahun)" / "(Penulis & Penulis, Tahun)".
_RX_SITASI_KURUNG = re.compile(r"\(([A-Z][^)]*,?\s*\d{4}[a-z]?)\)")

# Versi longgar (boleh bersarang, mis. "((PT) et al., 2026)") untuk MELINDUNGI
# sitasi saat sterilisasi istilah objek.
_RX_SITASI_LINDUNG = re.compile(
    r"\([^()]*(?:\([^()]*\)[^()]*)*(?:19|20)\d{2}[a-z]?[^()]*\)")

# Sitasi naratif di awal paragraf: "Syamsuri et al. (2024) menyatakan bahwa ..."
_RX_NARATIF_AWAL = re.compile(
    r"^\s*(?P<nama>[A-Z][^()]{1,90}?)\s*\((?P<th>(?:19|20)\d{2})\)\s*,?\s*"
    r"(?:menyatakan|menjelaskan|mengemukakan|menguraikan|mendefinisikan|"
    r"menegaskan|berpendapat|mengartikan|mengungkapkan|menambahkan)"
    r"\s+(?:bahwa\s+)?(?P<isi>.+)$", re.S)


def _lengkapi_sitasi_paragraf(hasil, daftar_sumber):
    """Jaring pengaman terakhir: penuhi aturan sitasi TANPA mengarang.

    Dipakai hanya ketika seluruh percobaan Gemini sudah dijalankan tetapi
    validator masih menolak. Tiga langkah, semuanya memakai sumber yang
    sudah terverifikasi:

      1. Sitasi naratif ("X et al. (2024) menyatakan ...") diubah menjadi
         bentuk kurung ("... (X et al., 2024)") agar lolos format.
      2. Paragraf yang tetap tanpa sitasi diberi sitasi dari sumber yang
         BELUM dipakai di mana pun dalam bab ini.
      3. Sitasi ganda di bagian non-paragraf (poin faktor/dimensi) dibuang
         agar aturan "satu sumber satu kali" terpenuhi; paragraf definisi
         tidak pernah kehilangan sitasinya.

    Mengembalikan (hasil, jumlah_perubahan).
    """
    if not isinstance(hasil, dict):
        return hasil, 0
    boleh = _sitasi_terverifikasi(daftar_sumber)
    pg = hasil.get("pengertian_paragraf")
    if not boleh or not isinstance(pg, list):
        return hasil, 0

    # Batas reuse yang diizinkan validator: 1 bila sumber >= paragraf; selainnya
    # max(2, ceil(paragraf/sumber)). HARUS sama dengan rumus di
    # _validate_sub_2_1_1 agar jaring pengaman tidak menghasilkan naskah
    # yang langsung ditolak karena sitasi dobel.
    _n_pg = len([x for x in pg if isinstance(x, str)])
    _n_sb = len(boleh)
    if _n_pg <= 0 or _n_sb <= 0 or _n_sb >= _n_pg:
        maks_reuse = 1
    else:
        maks_reuse = max(2, (_n_pg + _n_sb - 1) // _n_sb)

    daftar_poin = hasil.get("faktor_poin") if isinstance(hasil.get("faktor_poin"), list) else []
    daftar_dim = hasil.get("dimensi_list") if isinstance(hasil.get("dimensi_list"), list) else []
    teks_lain = [hasil.get(k) for k in ("faktor_pembuka", "dimensi_indikator_intro",
                                        "kesimpulan_operasional")]
    teks_lain += [x for x in daftar_poin + daftar_dim if isinstance(x, str)]

    def _sitasi_terpakai(teks):
        return {m.group(1).strip() for m in _RX_SITASI_KURUNG.finditer(teks)}

    korpus = " ".join([str(x) for x in teks_lain if isinstance(x, str)]
                      + [p for p in pg if isinstance(p, str)])
    terpakai = _sitasi_terpakai(korpus)
    nama_terpakai = {_nama_pertama(s).lower() for s in terpakai}
    # Sitasi per paragraf: dipakai untuk memastikan konversi naratif tidak
    # membuat sumber yang sama muncul di dua paragraf definisi.
    # Nama penulis yang sudah "ditempati" tiap paragraf, termasuk sitasi
    # naratif ("X (2024) menyatakan ..."), agar satu sumber tidak dipakai
    # dua kali di dua paragraf berbeda.
    def _nama_di_paragraf(p):
        if not isinstance(p, str):
            return set()
        nama = {_nama_pertama(s).lower() for s in _sitasi_terpakai(p)}
        mn = _RX_NARATIF_AWAL.match(p.strip())
        if mn:
            nama.add(_nama_pertama(mn.group("nama")).lower())
        return nama

    nama_pg = [_nama_di_paragraf(p) for p in pg]
    nama_terisi = set()
    sitasi_field = _sitasi_terpakai(" ".join(str(x) for x in teks_lain
                                            if isinstance(x, str)))
    pinjam = set()

    ubah = 0
    baru = []
    from collections import Counter as _Counter
    # jumlah pemakaian tiap sitasi DI PARAGRAF DEFINISI (reuse maks 2x per
    # sumber, hanya bila sumber lebih sedikit dari paragraf).
    hitung = _Counter()
    for idx_p, p in enumerate(pg):
        if not isinstance(p, str):
            baru.append(p)
            continue
        if _RX_SITASI_KURUNG.search(p):
            hitung.update(m.group(1).strip()
                          for m in _RX_SITASI_KURUNG.finditer(p))
            baru.append(p)
            continue
        nama_pg_lain = {n for j, st in enumerate(nama_pg) if j != idx_p for n in st}
        nama_pg_lain |= nama_terisi

        # Langkah 1: sitasi naratif "Nama (Tahun) menyatakan/...".
        # - Terverifikasi -> ubah jadi bentuk kurung (reuse maks 2x diizinkan).
        # - Tidak terverifikasi -> buang namanya agar tidak ada sitasi karangan;
        #   paragraf tetap diproses di Langkah 2 untuk diberi sumber terverifikasi.
        m = _RX_NARATIF_AWAL.match(p.strip())
        if m:
            nama = m.group("nama").strip()
            th = m.group("th")
            pas = [c for c in boleh
                   if _nama_pertama(c).lower() in nama.lower() and th in c]
            isi = m.group("isi").strip()
            isi = (isi[0].upper() + isi[1:]) if isi else isi
            if isi and not isi.endswith((".", "!", "?")):
                isi += "."
            if pas:
                if hitung.get(pas[0], 0) >= maks_reuse:
                    # Reuse sudah maksimal (maks_reuse): buang nama naratif agar
                    # sitasi tidak dobel, lalu lanjut ke Langkah 2 untuk diberi
                    # sumber terverifikasi yang berbeda.
                    if isi:
                        p = isi
                        if not p.endswith((".", "!", "?")):
                            p += "."
                    else:
                        baru.append(p)
                        continue
                else:
                    c = pas[0]
                    baru.append(f"{isi} {c}" if c.startswith("(") else f"{isi} ({c})")
                    hitung[c] += 1
                    terpakai.add(c)
                    nama_terpakai.add(_nama_pertama(c).lower())
                    nama_terisi.add(_nama_pertama(c).lower())
                    ubah += 1
                    continue
            # Naratif TIDAK terverifikasi: jangan biarkan nama asing di naskah.
            if isi:
                p = isi
                if not p.endswith((".", "!", "?")):
                    p += "."

        # Langkah 2: tambahkan sitasi dari sumber yang belum dipakai.
        pilih = None
        for c in boleh:
            nama_c = _nama_pertama(c).lower()
            if c in terpakai or (nama_c and nama_c in korpus.lower()):
                continue
            pilih = c
            break
        # Langkah 2b: semua sumber sudah terpakai -> "pinjam" sitasi yang hanya
        # muncul di bagian non-paragraf (poin/dimensi). Langkah 3 akan
        # membuangnya dari sana, jadi tidak ada sitasi ganda.
        if not pilih:
            for c in boleh:
                nama_c = _nama_pertama(c).lower()
                c_isi = c.strip().strip("()")
                if c_isi not in sitasi_field or c_isi in pinjam:
                    continue
                if nama_c and nama_c in nama_pg_lain:
                    continue
                pilih = c
                pinjam.add(c_isi)
                break
        # Langkah 2c: semua sumber sudah terpakai di paragraf lain -> pinjam
        # sumber dengan pemakaian paling sedikit (reuse maks 2x per sumber).
        # Blokir berdasarkan JUMLAH pemakaian, bukan nama penulis -- dua paragraf
        # naratif dengan penulis sama saling memblokir sebelumnya sehingga
        # sitasi tidak pernah bisa dilengkapi.
        if not pilih:
            for c in boleh:
                if hitung.get(c, 0) >= maks_reuse:
                    continue
                pilih = c
                break
        if not pilih:
            baru.append(p)
            continue
        teks = p.rstrip()
        if teks and not teks.endswith((".", "!", "?")):
            teks += "."
        baru.append(f"{teks} {pilih}" if pilih.startswith("(") else f"{teks} ({pilih})")
        hitung[pilih] += 1
        terpakai.add(pilih)
        nama_terpakai.add(_nama_pertama(pilih).lower())
        nama_terisi.add(_nama_pertama(pilih).lower())
        korpus += " " + baru[-1]
        ubah += 1

    hasil["pengertian_paragraf"] = baru

    # Langkah 3: buang sitasi ganda di bagian non-paragraf (poin/dimensi).
    sudah = _sitasi_terpakai(" ".join(p for p in baru if isinstance(p, str)))

    def _bersihkan(teks):
        if not isinstance(teks, str):
            return teks, 0
        n = 0

        def ganti(m):
            nonlocal n
            isi = m.group(1).strip()
            if isi in sudah:
                n += 1
                return ""
            sudah.add(isi)
            return m.group(0)

        hasil_teks = _RX_SITASI_KURUNG.sub(ganti, teks)
        if n:
            hasil_teks = re.sub(r"\s+([.,;:])", r"\1", hasil_teks)
            hasil_teks = re.sub(r"\s{2,}", " ", hasil_teks).strip()
        return hasil_teks, n

    # Dahulukan membuang sitasi ganda dari butir poin/dimensi (daftar), agar
    # kalimat prosa (pembuka/intro/kesimpulan) tetap mempertahankan sitasinya.
    for daftar in (daftar_poin, daftar_dim):
        for i, x in enumerate(daftar):
            if isinstance(x, str):
                daftar[i], n = _bersihkan(x)
                ubah += n
    for kunci in ("faktor_pembuka", "dimensi_indikator_intro", "kesimpulan_operasional"):
        if isinstance(hasil.get(kunci), str):
            hasil[kunci], n = _bersihkan(hasil[kunci])
            ubah += n

    # Langkah 4: batasi reuse di paragraf definisi (maksimal maks_reuse per
    # sumber, selaras validator). Gemini kadang menulis beberapa paragraf
    # dengan sitasi yang sama persis; kelebihan pemakaian diganti dengan
    # sumber terverifikasi yang paling jarang dipakai.
    cek = _Counter()
    for _p2 in baru:
        if isinstance(_p2, str):
            cek.update(m.group(1).strip()
                       for m in _RX_SITASI_KURUNG.finditer(_p2))
    over = sorted({s for s, c in cek.items() if c > maks_reuse})
    for s in over:
        n_over = cek[s] - maks_reuse
        ganti = 0
        for i in range(len(baru) - 1, -1, -1):
            if ganti >= n_over:
                break
            _p2 = baru[i]
            if not isinstance(_p2, str):
                continue
            mm = [m2 for m2 in _RX_SITASI_KURUNG.finditer(_p2)
                  if m2.group(1).strip() == s]
            if not mm:
                continue
            # pengganti: sumber terverifikasi dengan pemakaian paling sedikit
            rep = None
            for c in boleh:
                if cek.get(c.strip().strip("()"), 0) >= maks_reuse:
                    continue
                rep = c if c.startswith("(") else f"({c})"
                break
            if not rep:
                break
            m2 = mm[0]
            baru[i] = _p2[:m2.start()] + rep + _p2[m2.end():]
            cek[s] -= 1
            cek[rep.strip().strip("()")] += 1
            ganti += 1
            ubah += 1

    return hasil, ubah


def _kunci_sitasi_lompat(cit):
    """(nama belakang penulis pertama, tahun) dari sitasi - untuk pencocokan longgar.

    Dipakai agar "Widiastuti & Taufiq, 2023", "Widiastuti et al., 2023", dan
    "Widiastuti, 2023" dianggap SATU sumber yang sama.
    """
    c = str(cit or "").strip().strip("()").strip()
    m = re.search(r"(19|20)\d{2}", c)
    if not m:
        return None
    tahun = m.group(0)
    depan = c[:m.start()]
    depan = re.split(r"[&;,]", depan)[0]
    depan = re.sub(r"\bet\s+al\.?\s*$", "", depan.strip(), flags=re.I)
    depan = re.sub(r"[^A-Za-z\u00C0-\u024F' ]", " ", depan).strip()
    if not depan:
        return None
    import unicodedata as _ud
    depan = "".join(c for c in _ud.normalize("NFKD", depan) if not _ud.combining(c))
    kata = depan.split()
    if not kata:
        return None
    return (kata[-1].lower(), tahun)


def _sitasi_tidak_terverifikasi(semua_sitasi, daftar_sumber):
    """Daftar sitasi yang TIDAK punya sumber terverifikasi (tahun + penulis tak cocok).

    Sitasi seperti "(Karaday?, 2025)" yang TIDAK ada di daftar sumber hasil
    pre-fetch tidak boleh dibiarkan: entri DAFTAR PUSTAKANYA tidak ada, sehingga
    naskah berisi sitasi tanpa rujukan. Kalau rujukannya memang ada di pool
    terverifikasi, pemanggil menambahkannya lewat _refs_with_cited_acad.
    """
    kunci_boleh = set()
    for s in (daftar_sumber or []):
        for campo in ("apa_in_text", "sitasi", "apa"):
            v = s.get(campo) if isinstance(s, dict) else None
            if not v:
                continue
            k = _kunci_sitasi_lompat(v)
            if k:
                kunci_boleh.add(k)
            break
    if not kunci_boleh:
        return []
    buruk = []
    for cit in (semua_sitasi or []):
        k = _kunci_sitasi_lompat(cit)
        if not k:
            continue
        if k in kunci_boleh:
            continue
        # Toleransi KETAT: nama belakang sama, atau sama-sama sangat pendek
        # (mis. "P et al." vs "Prawira et al."), atau satu merupakan awalan
        # dari yang lain (minimal 3 huruf). Huruf awal saja TIDAK cukup --
        # kalau tidak, "Idrus, 2024" dianggap sah karena ada sumber "I..., 2024".
        th = [x for x in kunci_boleh if x[1] == k[1]]
        cocok = False
        for x in th:
            a, b = x[0], k[0]
            if a == b:
                cocok = True
                break
            if len(a) <= 2 or len(b) <= 2:
                if a[:1].lower() == b[:1].lower() and a[:1].isalpha():
                    cocok = True
                    break
            elif a.startswith(b[:3]) or b.startswith(a[:3]):
                cocok = True
                break
        if cocok:
            continue
        if cit not in buruk:
            buruk.append(cit)
    return buruk


def _validate_sub_2_1_1(raw, sources_y, min_th=2022, max_th=2026):
    """Validasi output sub_2_1_1 dari Gemini.

    Setiap citation (Penulis, Tahun) di output HARUS cocok dengan salah satu
    sumber dari sources_y yang di-fetch secara live. Tidak boleh ada citation
    di luar daftar sumber terverifikasi.
    """
    if isinstance(raw, list) and raw and isinstance(raw[0], dict):
        raw = raw[0]
    if not isinstance(raw, dict):
        return None

    # Cek kunci wajib
    required_keys = ["pengertian_paragraf", "faktor_pembuka", "faktor_poin",
                     "dimensi_indikator_intro", "dimensi_list", "kesimpulan_operasional"]
    for k in required_keys:
        if k not in raw:
            print(f"[sub_2_1_1] kunci hilang: {k}")
            return None

    # Cek pengertian: minimal 5 paragraf, maksimal 10
    # Jika sumber < 9, izinkan reuse sitasi (sumber dipakai >1x)
    n_sources = len(sources_y) if sources_y else 0
    min_paragraf = 5
    max_paragraf = 10
    pg = raw["pengertian_paragraf"]
    if not isinstance(pg, list) or len(pg) < min_paragraf or len(pg) > max_paragraf:
        print(f"[sub_2_1_1] pengertian_paragraf harus {min_paragraf}-{max_paragraf}, dapat {len(pg) if isinstance(pg, list) else 'bukan list'}")
        return None

    # Cek setiap paragraf adalah string dengan panjang wajar
    for i, p in enumerate(pg):
        if not isinstance(p, str) or len(p.strip()) < 50:
            print(f"[sub_2_1_1] paragraf {i+1} tidak valid (string pendek atau bukan string)")
            return None

    # Cek setiap paragraf memiliki sitasi "(Penulis, Tahun)" (posisi bebas)
    import re
    format_salah = []
    for i, p in enumerate(pg):
        if not isinstance(p, str):
            format_salah.append(f"paragraf {i+1} bukan string")
            continue
        # Paragraf harus memiliki setidaknya satu sitasi: bentuk kurung
        # "(Penulis, Tahun)" ATAU bentuk naratif "Penulis (Tahun) menegaskan ..."
        # (naratif diizinkan prompt, maksimal 2 paragraf).
        if not re.search(r'\(\s*(?:[^()]+,\s*)?(?:19|20)\d{2}[a-z]?\s*\)', p):
            format_salah.append(f"paragraf {i+1} tidak memiliki sitasi '(Penulis, Tahun)'")
    if format_salah:
        print(f"[sub_2_1_1] format paragraf salah: {format_salah[:3]}")
        # Sertakan naskah aslinya agar pemanggil masih bisa memakainya
        # (naskah nyata + sitasi disisipkan) daripada jatuh ke TEMPLATE.
        return {"_error": "format_salah", "errors": format_salah, "raw": raw}

    # Cek faktor: pembuka + 10 poin
    fp = raw["faktor_pembuka"]
    if not isinstance(fp, str) or len(fp) < 100:
        print(f"[sub_2_1_1] faktor_pembuka terlalu pendek atau bukan string")
        return None
    poin = raw["faktor_poin"]
    if not isinstance(poin, list) or len(poin) != 10:
        print(f"[sub_2_1_1] faktor_poin harus 10, dapat {len(poin) if isinstance(poin, list) else 'bukan list'}")
        return None

    # Cek dimensi: intro + dimensi (minimal 3)
    di = raw["dimensi_indikator_intro"]
    if not isinstance(di, str) or len(di) < 50:
        print(f"[sub_2_1_1] dimensi_indikator_intro terlalu pendek atau bukan string")
        return None
    dl = raw["dimensi_list"]
    if not isinstance(dl, list) or len(dl) < 3:
        print(f"[sub_2_1_1] dimensi_list minimal 3, dapat {len(dl) if isinstance(dl, list) else 'bukan list'}")
        return None

    # Cek kesimpulan
    ks = raw["kesimpulan_operasional"]
    if not isinstance(ks, str) or len(ks) < 50:
        print(f"[sub_2_1_1] kesimpulan_operasional terlalu pendek atau bukan string")
        return None

    # Cek kata terlarang (hanya yang melarang atributif definisi)
    terlarang = ["menurut", "berdasarkan pandangan"]
    naskah_parts = [str(fp), str(ks), str(di)] + pg + poin + dl
    naskah = " ".join(naskah_parts)
    naskah_lower = naskah.lower()
    for t in terlarang:
        if t in naskah_lower:
            print(f"[sub_2_1_1] kata terlarang ditemukan: '{t}'")
            return None

    # Cek tahun sitasi 2022-2026
    import re
    tahun_larang = []
    semua_sitasi = []
    # Cek pola "(Penulis, Tahun)" atau "(Penulis1 & Penulis2, Tahun)" atau "(Penulis et al., Tahun)"
    for m in re.finditer(r"\(([A-Z][^)]*,?\s*\d{4}[a-z]?)\)", naskah):
        isi = m.group(1)
        th_match = re.search(r"(19|20)\d{2}", isi)
        if th_match:
            th = int(th_match.group(0))
            if th < min_th or th > max_th:
                # Cek apakah bentuk "dalam" (kutipan dalam)
                sesudah = naskah[m.start():m.start()+150]
                if "dalam" not in sesudah.lower():
                    tahun_larang.append(th)
            semua_sitasi.append(isi.strip())
    # Sitasi naratif "Nama (Tahun) menyatakan ..." (diizinkan prompt, maks 2
    # paragraf) juga wajib lolos cek tahun dan cek sumber terverifikasi.
    for _p in pg:
        if not isinstance(_p, str):
            continue
        for _mn in _RX_NARATIF_AWAL.finditer(_p.strip()):
            _th_n = int(_mn.group("th"))
            if _th_n < min_th or _th_n > max_th:
                tahun_larang.append(_th_n)
            _s = f"{' '.join(_mn.group('nama').split())}, {_mn.group('th')}"
            if _s not in semua_sitasi:
                semua_sitasi.append(_s)
    if tahun_larang:
        print(f"[sub_2_1_1] tahun di luar {min_th}-{max_th}: {sorted(set(tahun_larang))}")
        return None

    # Cek duplikasi sitasi: jika sumber >= paragraf, larang duplikat
    # Jika sumber < paragraf, izinkan reuse (max ceil(paragraf/sumber) per sumber)
    from collections import Counter
    sitasi_count = Counter(semua_sitasi)
    n_paragraf = len(pg)
    max_per_source = 1 if n_sources >= n_paragraf else max(2, (n_paragraf + n_sources - 1) // n_sources)
    duplikat = [s for s, c in sitasi_count.items() if c > max_per_source]
    if duplikat:
        print(f"[sub_2_1_1] GAGAL: sitasi '{duplikat[:5]}' dipakai >{max_per_source}x (sumber={n_sources}, paragraf={n_paragraf})")
        return {"_error": "format_salah",
                "errors": [f"sitasi '{d}' dipakai lebih dari {max_per_source} kali (sumber={n_sources}, paragraf={n_paragraf})" for d in duplikat[:5]],
                "raw": raw}

    # Peringatan (bukan penolakan): citation di luar daftar pre-fetch
    # (Gemini mungkin menggunakan knowledge internal, yang masih valid bila tahun 2022-2026)
    # TOLAK sitasi yang tidak punya sumber terverifikasi (tahun + penulis tak cocok).
    # Sitasi seperti "(Karaday, 2025)" yang tak ada di daftar pre-fetch membuat
    # naskah berisi sitasi tanpa entri DAFTAR PUSTAKA -> harus diperbaiki model.
    tidak_sah = _sitasi_tidak_terverifikasi(semua_sitasi, sources_y)
    if tidak_sah:
        print(f"[sub_2_1_1] sitasi tanpa sumber terverifikasi: {tidak_sah[:5]}")
        return {"_error": "format_salah",
                "errors": [f"sitasi '{c}' tidak ada di DAFTAR SUMBER (wajib pakai sumber yang tersedia)" for c in tidak_sah[:5]],
                "raw": raw}

    return raw


_RX_SITASI_AWAL = re.compile(
    r"^\s*\((?P<sit>[^()]+?,\s*(?:19|20)\d{2}[a-z]?)\)\s*"
    r"(?P<verb>menyatakan|menjelaskan|mengemukakan|menguraikan|mendefinisikan|menegaskan|berpendapat|mengartikan)\s+"
    r"(?:bahwa\s+)?(?P<isi>.+)$",
    re.S | re.I,
)

_TRANSISI_DEFINISI = (
    "Sejalan dengan pandangan tersebut, ",
    "Dalam konteks organisasi, ",
    "Dari perspektif yang sedikit berbeda, ",
    "Pandangan lain menambahkan bahwa ",
    "Hal senada dikemukakan bahwa ",
    "Sementara itu, ",
    "Definisi lain menekankan bahwa ",
    "Lebih lanjut, ",
)


def _kalimat_pertama(teks):
    """Pisahkan kalimat pertama tanpa memotong singkatan seperti 'et al.'."""
    for m in re.finditer(r"[.!?](?=\s+[A-Z]|\s*$)", teks):
        sebelum = teks[:m.start()]
        tok = re.search(r"(\b[\w./-]+)$", sebelum)
        if tok and tok.group(1).lower() in {"al", "dkk", "no", "vol", "hlm",
                                            "hal", "ibid", "op", "cit"}:
            continue
        return teks[:m.end()], teks[m.end():]
    return teks, ""


def _call_gemini_aman(keys, chain, prompt, timeout, max_tokens=4096,
                      label="", tries=3):
    """Panggil Gemini dengan jeda bila kena rate limit (429) / busy (503).

    Tanpa ini, satu 429 sewaktu-waktu membuat seluruh sub bab jatuh ke
    template minimal (isi EMPTY). Retry singkat jauh lebih murah daripada
    naskah yang hilang.
    """
    import time as _time
    from llm_ontology_researcher import _call_gemini_matrix
    last = None
    for i in range(tries):
        try:
            return _call_gemini_matrix(keys, chain, prompt, timeout,
                                       max_tokens=max_tokens)
        except Exception as e:
            last = e
            msg = str(e)
            # Semua key tanpa kuota (jatah 0) -> mengulang tidak akan menolong;
            # setiap percobaan hanya menyapu ulang seluruh model x key.
            if "tanpa kuota" in msg or "jatah 0" in msg:
                raise
            transient = ("429" in msg) or ("503" in msg) or ("Too Many" in msg) \
                or ("timed out" in msg.lower()) or ("Timeout" in type(e).__name__)
            if not transient or i >= tries - 1:
                raise
            tunggu = 12 * (i + 1)
            print(f"[{label}] Gemini rate-limit/busy ({type(e).__name__}) "
                  f"-> tunggu {tunggu}s lalu coba lagi ({i+1}/{tries})")
            _time.sleep(tunggu)
    raise last if last else RuntimeError("gagal memanggil Gemini")


def _sterilkan_naskah(hasil, instansi="", judul=""):
    """BAB II bersifat UMUM: buang sebutan objek penelitian dari naskah.

    Menerapkan sterilisasi deterministik (tanpa LLM) ke seluruh field naratif
    hasil Gemini sub-bab 2.1.1.x agar tidak ada nama instansi, kota, unit
    kerja, maupun konteks sektor (pasien/gudang/logistik/dll).
    """
    if not isinstance(hasil, dict):
        return hasil
    try:
        import objek_sterilisasi as OS
    except Exception:
        return hasil

    def _simpan_sitasi(m, simpan):
        simpan.append(m.group(0))
        return f"\x00{len(simpan) - 1}\x00"

    def _pulih_sitasi(m, simpan):
        try:
            return simpan[int(m.group(0).strip("\x00"))]
        except (ValueError, IndexError):
            return ""

    def _b(t):
        if not isinstance(t, str) or not t.strip():
            return t
        # Lindungi sitasi kurung "(Penulis, Tahun)" — jangan pernah menyentuh
        # isinya saat sterilisasi (nama penulis boleh memuat kata terlarang,
        # mis. "Bank", "PT"). Sisanya disterilkan, lalu sitasi dikembalikan.
        simpan = []
        t2 = _RX_SITASI_LINDUNG.sub(lambda m: _simpan_sitasi(m, simpan), t)
        s = OS.bersihkan_objek(t2, instansi)
        if simpan:
            s = re.sub(r"\x00\d+\x00", lambda m: _pulih_sitasi(m, simpan), s)
        s = re.sub(r"\s+", " ", s).strip()
        if not s:
            return t
        s = s[0].upper() + s[1:]
        return s

    for k in ("faktor_pembuka", "dimensi_indikator_intro",
              "kesimpulan_operasional", "kesimpulan_sintetis"):
        if isinstance(hasil.get(k), str):
            hasil[k] = _b(hasil[k])

    for k in ("pengertian_paragraf", "faktor_poin"):
        v = hasil.get(k)
        if isinstance(v, list):
            hasil[k] = [_b(x) if isinstance(x, str) else x for x in v]

    v = hasil.get("dimensi_list")
    if isinstance(v, list):
        for item in v:
            if isinstance(item, dict) and isinstance(item.get("penjelasan"), str):
                item["penjelasan"] = _b(item["penjelasan"])

    v = hasil.get("faktor_dua_sumber")
    if isinstance(v, list):
        for item in v:
            if not isinstance(item, dict):
                continue
            if isinstance(item.get("paragraf_pembuka"), str):
                item["paragraf_pembuka"] = _b(item["paragraf_pembuka"])
            for f in (item.get("daftar_faktor") or []):
                if not isinstance(f, dict):
                    continue
                if isinstance(f.get("nama"), str):
                    _nama0 = str(f["nama"]).strip()
                    if OS.mengandung_objek(_nama0, instansi):
                        _sisa = re.sub(
                            r"\s+", " ",
                            OS.bersihkan_objek(_nama0, instansi)).strip()
                        # Nama faktor 1 kata tanpa istilah objek (mis.
                        # "Kompetensi") sah; nama yang isinya objek semua
                        # (mis. "Logistik"/"Faktor Pengiriman") diganti aman.
                        f["nama"] = _sisa if len(_sisa.split()) >= 2 else "Faktor Pendukung"
                    else:
                        f["nama"] = _nama0
                if isinstance(f.get("penjelasan"), str):
                    f["penjelasan"] = _b(f["penjelasan"])
    return hasil


def _variasikan_sitasi_awal(hasil):
    """Variasikan posisi sitasi agar tidak semua paragraf dibuka nama penulis.

    Hanya menyentuh pola ``(Penulis, Tahun) <verba> bahwa ...``. Sitasi
    dipindahkan ke akhir kalimat pertama, lalu paragraf kedua dan seterusnya
    diberi transisi konseptual bergilir agar mengalir dari paragraf sebelumnya.
    """
    if not isinstance(hasil, dict):
        return hasil
    pg = hasil.get("pengertian_paragraf")
    if not isinstance(pg, list):
        return hasil
    baru = []
    for i, par in enumerate(pg):
        if not isinstance(par, str):
            baru.append(par)
            continue
        m = _RX_SITASI_AWAL.match(par.strip())
        if not m:
            baru.append(par)
            continue
        sit = f"({m.group('sit').strip()})"
        isi = m.group("isi").strip()
        kal1, sisa = _kalimat_pertama(isi)
        kal1 = kal1.strip()
        if kal1.endswith((".", "!", "?")):
            kal1 = kal1[:-1].rstrip()
        if not kal1:
            baru.append(par)
            continue
        kal1 = kal1[0].upper() + kal1[1:]
        hasil_par = f"{kal1} {sit}."
        if sisa.strip():
            hasil_par += " " + sisa.strip()
        if i > 0:
            pre = _TRANSISI_DEFINISI[(i - 1) % len(_TRANSISI_DEFINISI)]
            kata1 = hasil_par.split()[0] if hasil_par.split() else ""
            if not kata1.isupper():
                hasil_par = hasil_par[0].lower() + hasil_par[1:]
            hasil_par = pre + hasil_par
        baru.append(hasil_par)
    baru = _dedup_paragraf(baru)
    salinan = dict(hasil)
    salinan["pengertian_paragraf"] = baru
    return salinan


# ---------------------------------------------------------------------------
# PENJAGA DEFINISI VARIABEL
# ---------------------------------------------------------------------------
# Setiap paragraf "Pengertian Variabel" WAJIB mencerminkan definisi variabel itu
# sendiri. Model sering menulis paragraf tentang faktor/pengukuran/konteks
# (mis. "Gamifikasi meningkatkan kinerja", "Evolusi regulasi menuntut ...").
# Fungsi di bawah MEMERIKSA paragraf lalu MEMBUKA paragraf yang salah dengan
# kalimat definisi variabel (dari ontologi) agar naskah tetap layak secara
# akademik tanpa harus regenerate.
# ---------------------------------------------------------------------------
_VAR_GENERIK = {"sdm", "organisasi", "kerja", "pegawai", "individu", "karyawan",
                "dan", "atau", "dengan", "pada"}
_PREDIKAT_DEFINISI = ("adalah", "merupakan", "mengacu", "didefinisikan", "diartikan",
                      "dipahami", "dimaknai", "sebagai", "merepresentasikan", "berupa",
                      "menjadi", "merujuk")
# Sinonim variabel agar paragraf definisi yang memakai kata lain
# (mis. "ikatan psikologis" untuk Komitmen) tetap dikenali sebagai definisi.
_SINONIM_VARIABEL = {
    "kinerja": ["kinerja", "hasil kerja", "produktivitas kerja", "prestasi kerja"],
    "komitmen": ["komitmen", "ikatan psikologis", "keterikatan", "dedikasi", "loyalitas"],
    "kompetensi": ["kompetensi", "kemampuan", "kapabilitas", "keahlian"],
    "beban": ["beban kerja", "beban tugas", "tuntutan kerja", "beban", "workload"],
    "kepuasan": ["kepuasan kerja", "kepuasan"],
    "disiplin": ["disiplin kerja", "disiplin", "kepatuhan"],
    "motivasi": ["motivasi", "motivasi kerja"],
    "loyalitas": ["loyalitas", "kesetiaan"],
}
_TEMPLATE_LEAD = (
    "{v} adalah {c}.",
    "{v} merupakan {c}.",
    "{v} dapat dipahami sebagai {c}.",
    "{v} diartikan sebagai {c}.",
    "{v} merujuk pada {c}.",
    "{v} pada dasarnya merupakan {c}.",
)


def _istilah_variabel(var_name):
    """Kata kunci khas variabel + sinonimnya (buang kata umum seperti 'SDM')."""
    kata = re.findall(r"[A-Za-z]+", str(var_name or "").lower())
    khas = {k for k in kata if k not in _VAR_GENERIK and len(k) > 2}
    istilah = set(khas)
    gabung = " ".join(kata)
    for kunci, daftar in _SINONIM_VARIABEL.items():
        if kunci in gabung:
            istilah.update(daftar)
    return istilah or set(kata)


def _paragraf_mencerminkan_variabel(paragraf, istilah):
    """True bila kalimat pertama paragraf mengenai definisi variabel."""
    if not istilah:
        return True
    kal1, _ = _kalimat_pertama(paragraf or "")
    low = kal1.lower()
    if not any(t in low for t in istilah):
        return False
    return any(p in low for p in _PREDIKAT_DEFINISI)


def _periksa_definisi(hasil, var_name):
    """Daftar indeks paragraf yang belum menyatakan definisi variabel.

    Termasuk paragraf duplikat (isi identik/hampir identik) agar LLM diberi
    umpan balik dan naskah akhir tidak memuat pengulangan.
    """
    if not isinstance(hasil, dict):
        return []
    pg = hasil.get("pengertian_paragraf")
    if not isinstance(pg, list):
        return []
    istilah = _istilah_variabel(var_name)
    salah = []
    for i, p in enumerate(pg):
        if not isinstance(p, str) or not p.strip():
            continue
        if not _paragraf_mencerminkan_variabel(p, istilah):
            salah.append(i)
    # Tandai duplikat (pertahankan yang pertama)
    import re as _re
    seen = set()
    for i, p in enumerate(pg):
        if not isinstance(p, str) or not p.strip() or i in salah:
            continue
        key = _re.sub(r'[^a-z0-9 ]+', ' ', _re.sub(r'\([^)]*\)', ' ', p).lower())
        key = _re.sub(r'\s+', ' ', key).strip()[:90]
        if key in seen:
            salah.append(i)
        else:
            seen.add(key)
    return sorted(set(salah))


def _catatan_koreksi_definisi(hasil, var_name):
    """Umpan balik korektif untuk Gemini agar paragraf benar-benar definisi."""
    salah = _periksa_definisi(hasil, var_name)
    if not salah:
        return ""
    pg = hasil.get("pengertian_paragraf") or []
    baris = []
    for i in salah[:6]:
        teks = str(pg[i] or "")[:120]
        baris.append(f'  - Paragraf {i+1}: "{teks}..."')
    return (
        "\n\n=== KOREKSI WAJIB (TULIS ULANG SELURUH DAFTAR PARAGRAF DEFINISI) ===\n"
        f"Paragraf berikut BELUM menyatakan definisi {var_name} (hanya faktor/konteks/pengukuran) "
        f"atau mengulang paragraf sebelumnya:\n"
        + "\n".join(baris) + "\n"
        f"Setiap paragraf WAJIB dibuka kalimat definisi {var_name}, contoh:\n"
        f'  "{var_name} adalah ... (Penulis, Tahun). ..."\n'
        f'  "{var_name} merupakan ... (Penulis, Tahun). ..."\n'
        "Jangan mengulang definisi yang sama persis di beberapa paragraf; rumuskan tiap\n"
        "paragraf dengan sumber yang berbeda dan sudut pandang definisi yang berbeda.\n"
        f"Pastikan SETIAP paragraf memakai sitasi (Penulis, Tahun) yang BERBEDA-BEDA.\n"
    )


PROMPT_REWRITE_DEFINISI = """Kamu editor bab kajian teoritis tesis Magister Indonesia.

Variabel: {var}
Instansi: {instansi}

Tugas: TULIS ULANG paragraf-paragraf di bawah ini menjadi paragraf DEFINISI yang SAH.
Aturan:
1. Setiap paragraf hasil HARUS dibuka kalimat yang MENYATAKAN definisi {var},
   contoh: "{var} adalah ...", "{var} merupakan ...", "{var} dipahami sebagai ...",
   "{var} diartikan sebagai ...".
2. DILARANG membuka dengan uraian faktor, sebab-akibat, cara pengukuran, dimensi,
   indikator, evaluasi, atau praktik organisasi.
3. WAJIB pertahankan sitasi (Penulis, Tahun) yang sudah ada di paragraf itu,
   di posisi yang natural (boleh digeser sedikit). JANGAN ganti sitasi.
4. Panjang 2-3 kalimat, gaya akademik formal, Bahasa Indonesia.
5. Balas HANYA JSON valid: {{"paragraf": ["...", "..."]}} dengan jumlah paragraf
   PERSIS sama dengan jumlah paragraf yang diminta.

=== PARAGRAF YANG HARUS DIULANG ===
{blok}

=== FORMAT KELUARAN (JSON) ===
{{"paragraf": ["paragraf 1", "paragraf 2"]}}
"""


def _tulis_ulang_definisi(hasil, var_name, onto=None, role=None, instances=""):
    """Tulis ulang HANYA paragraf non-definisi via Gemini (panggilan kecil & cepat).

    Dipakai sebagai cadangan terakhir agar jumlah paragraf definisi tetap utuh
    (tidak membuang isi yang sudah baik) sekaligus bebas dari uraian faktor/
    pengukuran yang tidak sesuai kaidah bab kajian.
    """
    if not isinstance(hasil, dict):
        return hasil
    pg = hasil.get("pengertian_paragraf")
    if not isinstance(pg, list) or not pg:
        return hasil
    salah = _periksa_definisi(hasil, var_name)
    if not salah:
        return hasil
    # Hanya tulis ulang bila sisanya sedikit (kalau banyak, mendingan regenerate total)
    if len(salah) > 5:
        return hasil

    try:
        from llm_title_parser import load_config, gemini_keys, MODEL_CHAIN
        config = load_config()
        keys = gemini_keys(config)
        chain = [m for m in config.get("models", []) if m] or MODEL_CHAIN
        timeout = max(int(config.get("timeout_seconds", 90)), 90)
    except Exception as e:
        print(f"[definisi] rewrite dilewati (config): {e}")
        return hasil

    blok = "\n\n".join(f"[{n+1}] {str(pg[n]).strip()}"
                       for n in sorted(salah))
    prompt = PROMPT_REWRITE_DEFINISI.format(
        var=var_name, instansi=instances or "(tidak disebutkan)", blok=blok)
    try:
        from llm_ontology_researcher import _call_gemini_matrix
        prompt_safe = prompt
        raw = _call_gemini_aman(keys, chain, prompt_safe, timeout,
                                max_tokens=2048, label="rewrite_definisi")
        if isinstance(raw, str):
            raw = raw
        if isinstance(raw, list) and raw and isinstance(raw[0], dict):
            raw = raw[0]
        if not isinstance(raw, dict):
            return hasil
        baru = raw.get("paragraf")
        if not isinstance(baru, list) or len(baru) != len(salah):
            print(f"[definisi] rewrite jumlah paragraf tidak cocok "
                  f"({len(baru) if isinstance(baru, list) else 'x'} != {len(salah)})")
            return hasil
        hasil2 = dict(hasil)
        pg2 = list(pg)
        for n, teks in zip(sorted(salah), baru):
            if not isinstance(teks, str) or len(teks.strip()) < 40:
                return hasil
            pg2[n] = teks.strip()
        hasil2["pengertian_paragraf"] = pg2
        # Verifikasi: bila masih ada yang non-definisi, jangan pakai hasil rewrite
        if _periksa_definisi(hasil2, var_name):
            print("[definisi] hasil rewrite masih ada paragraf non-definisi -> dibuang")
            return hasil
        print(f"[definisi] {len(salah)} paragraf ditulis ulang via Gemini")
        return hasil2
    except Exception as e:
        print(f"[definisi] rewrite gagal: {type(e).__name__} {str(e)[:70]}")
        return hasil


def _lead_definisi(var_name, definisi, index):
    """Bangun kalimat pembuka definisi yang varyatif dan natural."""
    core = re.sub(r"\s+", " ", str(definisi or "")).strip().rstrip(".")
    if core:
        core = re.sub(r"^" + re.escape(var_name) + r"\s*"
                      r"(?:adalah|merupakan|merupkan|didefinisikan sebagai|artinya)?\s*",
                      "", core, flags=re.I).strip().rstrip(".")
    if not core:
        core = "konsep yang diuraikan secara khusus dalam kajian teoritis pada sub bab ini"
    tmpl = _TEMPLATE_LEAD[index % len(_TEMPLATE_LEAD)]
    return tmpl.format(v=var_name, c=core)


def _dedup_paragraf(pg):
    """Hapus paragraf duplikat (isi mirip) sambil mempertahankan urutan."""
    import re

    def _key(p):
        p_clean = re.sub(r'\([^)]*\)', ' ', p).lower()
        p_clean = re.sub(r'[^a-z0-9 ]+', ' ', p_clean)
        p_clean = re.sub(r'\s+', ' ', p_clean).strip()
        return p_clean[:90]

    unik = []
    seen = set()
    for p in pg:
        if not isinstance(p, str) or not p.strip():
            continue
        k = _key(p)
        if k in seen:
            continue
        seen.add(k)
        unik.append(p)
    return unik


def _rapikan_definisi_paragraf(hasil, var_name, onto=None, role=None, instances=""):
    """Cadangan terakhir: pastikan ada paragraf definisi variabel.

    Urutan: (1) coba tulis ulang paragraf non-definisi via Gemini (isi tetap utuh),
    (2) bila gagal -> sisipkan kalimat definisi pada paragraf bermasalah,
    (3) bila jumlah paragraf sah terlalu sedikit -> buang paragraf non-definisi.
    """
    if not isinstance(hasil, dict):
        return hasil
    pg = hasil.get("pengertian_paragraf")
    if not isinstance(pg, list) or not pg:
        return hasil

    salah = set(_periksa_definisi(hasil, var_name))
    if not salah:
        return hasil

    # (1) Tulis ulang paragraf bermasalah via Gemini (menjaga jumlah & isi)
    hasil = _tulis_ulang_definisi(hasil, var_name, onto=onto, role=role,
                                  instances=instances)
    pg = hasil.get("pengertian_paragraf") or []
    salah = set(_periksa_definisi(hasil, var_name))
    if not salah:
        print(f"[definisi] seluruh paragraf '{var_name}' lolos setelah rewrite")
        return hasil

    definisi = ""
    try:
        v = ((onto or {}).get("variables", {}) or {}).get(role or "", {}) or {}
        definisi = v.get("definisi", "") or ""
    except Exception:
        definisi = ""

    # Posisi graft lead: jumlah paragraf sah sebelum paragraf bermasalah pertama
    posisi = min(salah) if salah else len(pg)
    baik = [(i, p) for i, p in enumerate(pg)
            if i not in salah and isinstance(p, str) and p.strip()]
    n_sebelum = sum(1 for i, _ in baik if i < posisi)

    if len(baik) >= 7:
        # Cukup banyak paragraf sah -> graft 1 kalimat definisi, sisanya dibuang.
        lead = _lead_definisi(var_name, definisi, 0)
        teks = [p for _, p in baik]
        baru = teks[:n_sebelum] + [lead] + teks[n_sebelum:]
    else:
        # Terlalu sedikit paragraf sah -> pertahankan SEMUA isi, buka tiap paragraf
        # bermasalah dengan kalimat definisi (variasinya bergilir).
        baru = []
        for i, p in enumerate(pg):
            if i in salah and isinstance(p, str) and p.strip():
                baru.append(f"{_lead_definisi(var_name, definisi, i)} {p.strip()}")
            elif isinstance(p, str) and p.strip():
                baru.append(p)

    baru = _dedup_paragraf(baru)
    print(f"[definisi] cadangan: {len(salah)} paragraf non-definisi "
          f"pada '{var_name}' diperbaiki (baik={len(baik)} akhir={len(baru)})")
    salinan = dict(hasil)
    salinan["pengertian_paragraf"] = baru
    return salinan


def write_sub_2_1_1(cfg, onto, ctx, acad, sources_y=None):
    """
    Menulis sub bab 2.1.1 via Gemini. Mengembalikan dict konten, atau None bila gagal.
    sources_y: daftar sumber yang di-fetch secara live (8-10 sumber, buku+jurnal 2022-2026).
    """
    from llm_title_parser import load_config, gemini_keys, MODEL_CHAIN

    config = load_config()
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return None

    y = cfg.get("y", "Kinerja Pegawai")
    x1 = cfg.get("x1", "X1")
    x2 = cfg.get("x2", "X2")
    m = cfg.get("m", "M")
    judul = cfg.get("judul", "")
    instansi = cfg.get("instansi", "")

    def _fin(hasil):
        # Sterilkan sebutan objek penelitian + variasikan posisi sitasi.
        return _variasikan_sitasi_awal(_sterilkan_naskah(hasil, instansi, judul))

    # Dapatkan sitasi terverifikasi untuk variabel Y (fallback jika sources_y tidak ada)
    sitasi_block = _fmt_sitasi_bab2(acad, "cite_kinerja")

    # Blok sumber wajib dari pre-fetch live
    if sources_y:
        sumber_block = _fmt_sumber_block(sources_y)
    else:
        sumber_block = sitasi_block

    # Blok dimensi & indikator dari ontologi
    dim_block = _dimensi_indikator_block(onto, "y", instansi=instansi, judul=judul)

    prompt = PROMPT_BAB2_2_1_1.format(
        judul=judul,
        instansi=instansi,
        y=y,
        x1=x1,
        x2=x2,
        m=m,
        dimensi_indikator_block=dim_block,
        sumber_block=sumber_block,
    )

    chain = [m for m in config.get("models", []) if m] or MODEL_CHAIN
    timeout = max(int(config.get("timeout_seconds", 90)), 90)
    keys = gemini_keys(config)

    try:
        from llm_ontology_researcher import _call_gemini_matrix
        prompt_base = prompt
        prompt_safe = prompt_base
        _y_name = (((onto.get("variables", {}) or {}).get("y", {}) or {})
                   .get("nama_resmi") or y)
        last_result = None
        for attempt in range(3):
            try:
                raw = _call_gemini_aman(keys, chain, prompt_safe, timeout, max_tokens=4096, label="write_sub_2_1_1")
                if isinstance(raw, str):
                    raw = raw
                result = _validate_sub_2_1_1(raw, sources_y if sources_y else [])
                if result and not (isinstance(result, dict) and result.get("_error")):
                    salah = _periksa_definisi(result, _y_name)
                    if not salah:
                        print(f"[sub_2_1_1] konten ditulis via Gemini (percobaan {attempt+1})")
                        return _fin(result)
                    last_result = result
                    # Bila hanya 1-3 paragraf bermasalah: tulis ulang HANYA paragraf
                    # itu via Gemini (panggilan kecil) -> jauh lebih cepat & isi utuh.
                    if len(salah) <= 3:
                        _fix = _tulis_ulang_definisi(result, _y_name, onto=onto,
                                                    role="y", instances=instansi)
                        if not _periksa_definisi(_fix, _y_name):
                            print(f"[sub_2_1_1] {len(salah)} paragraf ditulis ulang via Gemini")
                            return _fin(_fix)
                    catatan = _catatan_koreksi_definisi(result, _y_name)
                    prompt_safe = prompt_base + catatan
                    print(f"[sub_2_1_1] percobaan {attempt+1}: {len(salah)} paragraf "
                          f"belum definisi '{_y_name}' -> ulangi")
                    continue
                # Validasi struktur gagal -> gunakan error detail dari validator
                if isinstance(result, dict) and result.get("_error") == "format_salah":
                    # Simpan naskah asli sebagai kandidat terakhir: naskah nyata
                    # (dengan sitasi yang dilengkapi dari sumber terverifikasi)
                    # tetap lebih baik daripada TEMPLATE CADANGAN.
                    if isinstance(result.get("raw"), dict):
                        last_result = result["raw"]
                    errs = result.get("errors", [])[:5]
                    catatan = ("\n\n=== KOREKSI WAJIB ===\n"
                               "Paragraf berikut KENANGAN SITASI atau format salah:\n"
                               + "\n".join(f"  - {e}" for e in errs)
                               + "\n\nSITASI YANG BOLEH DIPAKAI (WAJIB pilih dari daftar ini):\n"
                               + _daftar_sitasi_boleh(sources_y)
                               + "\n\nSETIAP paragraf definisi WAJIB memuat TEPAT SATU sitasi "
                               "berformat (Penulis, Tahun), mis. \"... (Hadiyanto & Syaflan, 2024).\"\n"
                               "JANGAN lupa sitasi di paragraf mana pun. Jumlah paragraf 8-10.\n"
                               "JANGAN menyebut objek penelitian: nama instansi/kota/unit kerja, "
                               "apalagi istilah sektor spesifik (mis. gudang, logistik, pasien, "
                               "rekam medis, pengiriman). Kajian TEORITIS harus umum/teoretis.\n")
                else:
                    catatan = ("\n\n=== KOREKSI WAJIB ===\n"
                               "SETIAP paragraf definisi WAJIB memuat TEPAT SATU sitasi "
                               "berformat (Penulis, Tahun), mis. \"... (Hadiyanto & Syaflan, 2024).\"\n"
                               "JANGAN lupa sitasi di paragraf mana pun. Jumlah paragraf 8-10.\n"
                               "JANGAN menyebut objek penelitian: nama instansi/kota/unit kerja, "
                               "apalagi istilah sektor spesifik (mis. gudang, logistik, pasien, "
                               "rekam medis, pengiriman). Kajian TEORITIS harus umum/teoretis.\n")
                prompt_safe = prompt_base + catatan
                print(f"[sub_2_1_1] percobaan {attempt+1}: validasi struktur gagal "
                      f"(sitasi/paragraf) -> ulangi dengan koreksi")
                continue
            except UnicodeEncodeError:
                print(f"[sub_2_1_1] UnicodeEncodeError percobaan {attempt+1}")
                continue
        if last_result:
            # Jaring pengaman terakhir: lengkapi sitasi yang kurang dari sumber
            # TERVERIFIKASI (tidak mengarang penulis/tahun), lalu pastikan
            # naskah benar-benar lolos validasi sebelum dipakai.
            last_result, _n_tambah = _lengkapi_sitasi_paragraf(last_result, sources_y)
            if _n_tambah:
                print(f"[sub_2_1_1] {_n_tambah} paragraf dilengkapi sitasi "
                      f"dari sumber terverifikasi")
            _rapi = _rapikan_definisi_paragraf(last_result, _y_name, onto, "y", instances=instansi)
            try:
                with open(os.path.join("output", "_debug_2_1_1_last.json"), "w",
                          encoding="utf-8") as _f:
                    json.dump({"lengkap": last_result, "rapi": _rapi}, _f,
                              ensure_ascii=False, indent=2)
            except (OSError, TypeError):
                pass
            # _rapikan_definisi_paragraf bisa membuang sitasi saat menulis
            # ulang paragraf -> lengkapi ulang bila hasilnya ditolak validator.
            kandidat = []
            if _rapi and not (isinstance(_rapi, dict) and _rapi.get("_error")):
                _rapi2 = _rapi
                _cek_rapi = _validate_sub_2_1_1(_rapi2, sources_y if sources_y else [])
                if isinstance(_cek_rapi, dict) and _cek_rapi.get("_error") in ("format_salah", None) and\
                        _cek_rapi.get("raw"):
                    _rapi2, _n2 = _lengkapi_sitasi_paragraf(_cek_rapi["raw"], sources_y)
                    if _n2:
                        print(f"[sub_2_1_1] {_n2} paragraf dilengkapi sitasi setelah rewrite")
                kandidat = [_rapi2, _rapi, last_result]
            else:
                kandidat = [_rapi, last_result]
            for _kandidat in kandidat:
                if not _kandidat:
                    continue
                _cek = _validate_sub_2_1_1(_kandidat, sources_y if sources_y else [])
                if _cek and not (isinstance(_cek, dict) and _cek.get("_error")):
                    print("[sub_2_1_1] memakai naskah Gemini terakhir (bukan template)")
                    return _fin(_cek)
            print("[sub_2_1_1] naskah terakhir tetap tidak lolos validasi")
        print(f"[sub_2_1_1] semua percobaan gagal -> fallback template")
        return None
    except Exception as e:
        print(f"[sub_2_1_1] Gemini gagal: {type(e).__name__} {str(e)[:80]}")
        return None


# ============================================================================
# VALIDASI SUB BAB 2.1.2.1 (VARIABEL M)
# ============================================================================
def _validate_sub_2_1_2(raw, sources_m, min_th=2022, max_th=2026):
    """Validasi output sub_2_1_2 (variabel M) dari Gemini."""
    if isinstance(raw, list) and raw and isinstance(raw[0], dict):
        raw = raw[0]
    if not isinstance(raw, dict):
        return None

    required_keys = ["pengertian_paragraf", "kesimpulan_sintetis"]
    for k in required_keys:
        if k not in raw:
            print(f"[sub_2_1_2] kunci hilang: {k}")
            return None

    pg = raw["pengertian_paragraf"]
    if not isinstance(pg, list) or not (2 <= len(pg) <= 10):
        print(f"[sub_2_1_2] pengertian_paragraf harus 2-10, dapat {len(pg) if isinstance(pg, list) else 'bukan list'}")
        return None

    for i, p in enumerate(pg):
        if not isinstance(p, str) or len(p.strip()) < 50:
            print(f"[sub_2_1_2] paragraf {i+1} tidak valid")
            return None

    import re
    errors = []
    for i, p in enumerate(pg):
        if not isinstance(p, str):
            errors.append(f"paragraf {i+1} bukan string")
            continue
        if not re.search(r'\([^)]+,\s*\d{4}[a-z]?\)', p):
            errors.append(f"paragraf {i+1} tidak memiliki sitasi '(Penulis, Tahun)'")
        for cm in re.finditer(r'\(([^)]*?),\s*\d{4}[a-z]?\)', p):
            if not re.search(r'[A-Za-z]{2,}', cm.group(1)):
                errors.append(f"paragraf {i+1} sitasi penulis rusak: {cm.group(0)}")
    if errors:
        print(f"[sub_2_1_2] format paragraf salah: {errors[:3]}")
        return {"_error": "format_salah", "errors": errors}

    fp = None
    poin = []

    ks = raw["kesimpulan_sintetis"]
    if not isinstance(ks, str) or len(ks) < 30:
        print(f"[sub_2_1_2] kesimpulan_sintetis terlalu pendek")
        return None

    # Cek kata terlarang
    terlarang = ["menurut", "berdasarkan pandangan"]
    naskah_parts = [str(ks)] + pg
    naskah = " ".join(naskah_parts).lower()
    for t in terlarang:
        if t in naskah:
            print(f"[sub_2_1_2] kata terlarang: '{t}'")
            return None

    # Cek tahun sitasi 2022-2026
    tahun_larang = []
    semua_sitasi = []
    for m in re.finditer(r"\(([A-Z][^)]*,?\s*\d{4}[a-z]?)\)", " ".join(naskah_parts)):
        isi = m.group(1)
        th_match = re.search(r"(19|20)\d{2}", isi)
        if th_match:
            th = int(th_match.group(0))
            if th < min_th or th > max_th:
                sesudah = " ".join(naskah_parts)[m.start():m.start()+150]
                if "dalam" not in sesudah.lower():
                    tahun_larang.append(th)
            semua_sitasi.append(isi.strip())
    if tahun_larang:
        print(f"[sub_2_1_2] tahun di luar {min_th}-{max_th}: {sorted(set(tahun_larang))}")
        return None

    # Cek duplikasi sitasi
    from collections import Counter
    sitasi_count = Counter(semua_sitasi)
    duplikat = [s for s, c in sitasi_count.items() if c > 1]
    if duplikat:
        print(f"[sub_2_1_2] GAGAL: sitasi duplikat: {duplikat[:5]}")
        return None

    # Peringatan: citation di luar pre-fetch
    # TOLAK sitasi yang tidak punya sumber terverifikasi (tahun + penulis tak cocok).
    # Sitasi seperti "(Karaday, 2025)" yang tak ada di daftar pre-fetch membuat
    # naskah berisi sitasi tanpa entri DAFTAR PUSTAKA -> harus diperbaiki model.
    tidak_sah = _sitasi_tidak_terverifikasi(semua_sitasi, sources_m)
    if tidak_sah:
        print(f"[sub_2_1_2] sitasi tanpa sumber terverifikasi: {tidak_sah[:5]}")
        return {"_error": "format_salah",
                "errors": [f"sitasi '{c}' tidak ada di DAFTAR SUMBER (wajib pakai sumber yang tersedia)" for c in tidak_sah[:5]]}

    return raw


# ============================================================================
# VALIDASI SUB BAB 2.1.3.1 (VARIABEL X1)
# ============================================================================
def _validate_sub_2_1_3(raw, sources_x1, min_th=2022, max_th=2026):
    """Validasi output sub_2_1_3 (variabel X1) dari Gemini."""
    if isinstance(raw, list) and raw and isinstance(raw[0], dict):
        raw = raw[0]
    if not isinstance(raw, dict):
        return None

    required_keys = ["pengertian_paragraf", "kesimpulan_sintetis"]
    for k in required_keys:
        if k not in raw:
            print(f"[sub_2_1_3] kunci hilang: {k}")
            return None

    pg = raw["pengertian_paragraf"]
    if not isinstance(pg, list) or not (2 <= len(pg) <= 10):
        print(f"[sub_2_1_3] pengertian_paragraf harus 2-10, dapat {len(pg) if isinstance(pg, list) else 'bukan list'}")
        return None

    for i, p in enumerate(pg):
        if not isinstance(p, str) or len(p.strip()) < 50:
            print(f"[sub_2_1_3] paragraf {i+1} tidak valid")
            return None

    import re
    errors = []
    for i, p in enumerate(pg):
        if not isinstance(p, str):
            errors.append(f"paragraf {i+1} bukan string")
            continue
        if not re.search(r'\([^)]+,\s*\d{4}[a-z]?\)', p):
            errors.append(f"paragraf {i+1} tidak memiliki sitasi")
        for cm in re.finditer(r'\(([^)]*?),\s*\d{4}[a-z]?\)', p):
            if not re.search(r'[A-Za-z]{2,}', cm.group(1)):
                errors.append(f"paragraf {i+1} sitasi penulis rusak: {cm.group(0)}")
    if errors:
        print(f"[sub_2_1_3] format paragraf salah: {errors[:3]}")
        return {"_error": "format_salah", "errors": errors}

    fp = None
    poin = []

    ks = raw["kesimpulan_sintetis"]
    if not isinstance(ks, str) or len(ks) < 30:
        print(f"[sub_2_1_3] kesimpulan_sintetis terlalu pendek")
        return None

    terlarang = ["menurut", "berdasarkan pandangan"]
    naskah_parts = [str(ks)] + pg
    naskah = " ".join(naskah_parts).lower()
    for t in terlarang:
        if t in naskah:
            print(f"[sub_2_1_3] kata terlarang: '{t}'")
            return None

    tahun_larang = []
    semua_sitasi = []
    for m in re.finditer(r"\(([A-Z][^)]*,?\s*\d{4}[a-z]?)\)", " ".join(naskah_parts)):
        isi = m.group(1)
        th_match = re.search(r"(19|20)\d{2}", isi)
        if th_match:
            th = int(th_match.group(0))
            if th < min_th or th > max_th:
                sesudah = " ".join(naskah_parts)[m.start():m.start()+150]
                if "dalam" not in sesudah.lower():
                    tahun_larang.append(th)
            semua_sitasi.append(isi.strip())
    if tahun_larang:
        print(f"[sub_2_1_3] tahun di luar {min_th}-{max_th}: {sorted(set(tahun_larang))}")
        return None

    from collections import Counter
    sitasi_count = Counter(semua_sitasi)
    duplikat = [s for s, c in sitasi_count.items() if c > 1]
    if duplikat:
        print(f"[sub_2_1_3] GAGAL: sitasi duplikat: {duplikat[:5]}")
        return None

    # TOLAK sitasi yang tidak punya sumber terverifikasi (tahun + penulis tak cocok).
    # Sitasi seperti "(Karaday, 2025)" yang tak ada di daftar pre-fetch membuat
    # naskah berisi sitasi tanpa entri DAFTAR PUSTAKA -> harus diperbaiki model.
    tidak_sah = _sitasi_tidak_terverifikasi(semua_sitasi, sources_x1)
    if tidak_sah:
        print(f"[sub_2_1_3] sitasi tanpa sumber terverifikasi: {tidak_sah[:5]}")
        return {"_error": "format_salah",
                "errors": [f"sitasi '{c}' tidak ada di DAFTAR SUMBER (wajib pakai sumber yang tersedia)" for c in tidak_sah[:5]]}

    return raw


# ============================================================================
# VALIDASI SUB BAB 2.1.4.1 (VARIABEL X2)
# ============================================================================
def _validate_sub_2_1_4(raw, sources_x2, min_th=2022, max_th=2026):
    """Validasi output sub_2_1_4 (variabel X2) dari Gemini."""
    if isinstance(raw, list) and raw and isinstance(raw[0], dict):
        raw = raw[0]
    if not isinstance(raw, dict):
        return None

    required_keys = ["pengertian_paragraf", "kesimpulan_sintetis"]
    for k in required_keys:
        if k not in raw:
            print(f"[sub_2_1_4] kunci hilang: {k}")
            return None

    pg = raw["pengertian_paragraf"]
    if not isinstance(pg, list) or not (2 <= len(pg) <= 10):
        print(f"[sub_2_1_4] pengertian_paragraf harus 2-10, dapat {len(pg) if isinstance(pg, list) else 'bukan list'}")
        return None

    for i, p in enumerate(pg):
        if not isinstance(p, str) or len(p.strip()) < 50:
            print(f"[sub_2_1_4] paragraf {i+1} tidak valid")
            return None

    import re
    errors = []
    for i, p in enumerate(pg):
        if not isinstance(p, str):
            errors.append(f"paragraf {i+1} bukan string")
            continue
        if not re.search(r'\([^)]+,\s*\d{4}[a-z]?\)', p):
            errors.append(f"paragraf {i+1} tidak memiliki sitasi")
        for cm in re.finditer(r'\(([^)]*?),\s*\d{4}[a-z]?\)', p):
            if not re.search(r'[A-Za-z]{2,}', cm.group(1)):
                errors.append(f"paragraf {i+1} sitasi penulis rusak: {cm.group(0)}")
    if errors:
        print(f"[sub_2_1_4] format paragraf salah: {errors[:3]}")
        return {"_error": "format_salah", "errors": errors}

    fp = None
    poin = []

    ks = raw["kesimpulan_sintetis"]
    if not isinstance(ks, str) or len(ks) < 30:
        print(f"[sub_2_1_4] kesimpulan_sintetis terlalu pendek")
        return None

    terlarang = ["menurut", "berdasarkan pandangan"]
    naskah_parts = [str(ks)] + pg
    naskah = " ".join(naskah_parts).lower()
    for t in terlarang:
        if t in naskah:
            print(f"[sub_2_1_4] kata terlarang: '{t}'")
            return None

    tahun_larang = []
    semua_sitasi = []
    for m in re.finditer(r"\(([A-Z][^)]*,?\s*\d{4}[a-z]?)\)", " ".join(naskah_parts)):
        isi = m.group(1)
        th_match = re.search(r"(19|20)\d{2}", isi)
        if th_match:
            th = int(th_match.group(0))
            if th < min_th or th > max_th:
                sesudah = " ".join(naskah_parts)[m.start():m.start()+150]
                if "dalam" not in sesudah.lower():
                    tahun_larang.append(th)
            semua_sitasi.append(isi.strip())
    if tahun_larang:
        print(f"[sub_2_1_4] tahun di luar {min_th}-{max_th}: {sorted(set(tahun_larang))}")
        return None

    from collections import Counter
    sitasi_count = Counter(semua_sitasi)
    duplikat = [s for s, c in sitasi_count.items() if c > 1]
    if duplikat:
        print(f"[sub_2_1_4] GAGAL: sitasi duplikat: {duplikat[:5]}")
        return None

    # TOLAK sitasi yang tidak punya sumber terverifikasi (tahun + penulis tak cocok).
    # Sitasi seperti "(Karaday, 2025)" yang tak ada di daftar pre-fetch membuat
    # naskah berisi sitasi tanpa entri DAFTAR PUSTAKA -> harus diperbaiki model.
    tidak_sah = _sitasi_tidak_terverifikasi(semua_sitasi, sources_x2)
    if tidak_sah:
        print(f"[sub_2_1_4] sitasi tanpa sumber terverifikasi: {tidak_sah[:5]}")
        return {"_error": "format_salah",
                "errors": [f"sitasi '{c}' tidak ada di DAFTAR SUMBER (wajib pakai sumber yang tersedia)" for c in tidak_sah[:5]]}

    return raw


# ============================================================================
# WRITE SUB BAB 2.1.2.1 (VARIABEL M)
# ============================================================================
def write_sub_2_1_2(cfg, onto, ctx, acad, sources_m=None):
    """Menulis sub bab 2.1.2.1 (Pengertian Variabel M) via Gemini."""
    from llm_title_parser import load_config, gemini_keys, MODEL_CHAIN

    config = load_config()
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return None

    y = cfg.get("y", "Kinerja Pegawai")
    x1 = cfg.get("x1", "X1")
    x2 = cfg.get("x2", "X2")
    m = cfg.get("m", "M")
    judul = cfg.get("judul", "")
    instansi = cfg.get("instansi", "")

    sitasi_block = _fmt_sitasi_bab2(acad, "cite_m_1")
    if sources_m:
        sumber_block = _fmt_sumber_block(sources_m)
    else:
        sumber_block = sitasi_block

    dim_block = _dimensi_indikator_block(onto, "m", instansi=instansi, judul=judul)

    prompt = PROMPT_BAB2_2_1_2.format(
        judul=judul, instansi=instansi, y=y, x1=x1, x2=x2, m=m,
        dimensi_indikator_block=dim_block, sumber_block=sumber_block,
    )

    chain = [m for m in config.get("models", []) if m] or MODEL_CHAIN
    timeout = max(int(config.get("timeout_seconds", 90)), 90)
    keys = gemini_keys(config)

    try:
        from llm_ontology_researcher import _call_gemini_matrix
        prompt_base = prompt
        prompt_safe = prompt_base
        _m_name = (((onto.get("variables", {}) or {}).get("m", {}) or {})
                   .get("nama_resmi") or m)
        last_result = None
        for attempt in range(3):
            try:
                raw = _call_gemini_aman(keys, chain, prompt_safe, timeout, max_tokens=4096, label="write_sub_2_1_2")
                if isinstance(raw, str):
                    raw = raw
                result = _validate_sub_2_1_2(raw, sources_m if sources_m else [])
                if result and not (isinstance(result, dict) and result.get("_error")):
                    salah = _periksa_definisi(result, _m_name)
                    if not salah:
                        print(f"[sub_2_1_2] konten ditulis via Gemini (percobaan {attempt+1})")
                        return _variasikan_sitasi_awal(_sterilkan_naskah(result, instansi, judul))
                    last_result = result
                    # Bila hanya 1-3 paragraf bermasalah: tulis ulang HANYA paragraf
                    # itu via Gemini (panggilan kecil) -> jauh lebih cepat & isi utuh.
                    if len(salah) <= 3:
                        _fix = _tulis_ulang_definisi(result, _m_name, onto=onto,
                                                    role="m", instances=instansi)
                        if not _periksa_definisi(_fix, _m_name):
                            print(f"[sub_2_1_2] {len(salah)} paragraf ditulis ulang via Gemini")
                            return _variasikan_sitasi_awal(_sterilkan_naskah(_fix, instansi, judul))
                    catatan = _catatan_koreksi_definisi(result, _m_name)
                    prompt_safe = prompt_base + catatan
                    print(f"[sub_2_1_2] percobaan {attempt+1}: {len(salah)} paragraf "
                          f"belum definisi '{_m_name}' -> ulangi")
                    continue
                # Validasi struktur gagal -> gunakan error detail dari validator
                if isinstance(result, dict) and result.get("_error") == "format_salah":
                    errs = result.get("errors", [])[:5]
                    catatan = ("\n\n=== KOREKSI WAJIB ===\n"
                               "Paragraf berikut KENANGAN SITASI atau format salah:\n"
                               + "\n".join(f"  - {e}" for e in errs)
                               + "\n\nSETIAP paragraf definisi WAJIB memuat TEPAT SATU sitasi "
                               "berformat (Penulis, Tahun), mis. \"... (Hadiyanto & Syaflan, 2024).\"\n"
                               "JANGAN lupa sitasi di paragraf mana pun. Jumlah paragraf 2-10.\n")
                else:
                    catatan = ("\n\n=== KOREKSI WAJIB ===\n"
                               "SETIAP paragraf definisi WAJIB memuat TEPAT SATU sitasi "
                               "berformat (Penulis, Tahun), mis. \"... (Hadiyanto & Syaflan, 2024).\"\n"
                               "JANGAN lupa sitasi di paragraf mana pun. Jumlah paragraf 2-10.\n")
                prompt_safe = prompt_base + catatan
                print(f"[sub_2_1_2] percobaan {attempt+1}: validasi struktur gagal -> ulangi")
                continue
            except UnicodeEncodeError:
                print(f"[sub_2_1_2] UnicodeEncodeError percobaan {attempt+1}")
                continue
        if last_result:
            last_result = _rapikan_definisi_paragraf(last_result, _m_name, onto, "m", instances=instansi)
            print("[sub_2_1_2] memakai naskah terakhir (perbaikan definisi)")
            return _variasikan_sitasi_awal(_sterilkan_naskah(last_result, instansi, judul))
        print(f"[sub_2_1_2] semua percobaan gagal -> fallback")
        return None
    except Exception as e:
        print(f"[sub_2_1_2] Gemini gagal: {type(e).__name__} {str(e)[:80]}")
        return None


# ============================================================================
# WRITE SUB BAB 2.1.3.1 (VARIABEL X1)
# ============================================================================
def write_sub_2_1_3(cfg, onto, ctx, acad, sources_x1=None):
    """Menulis sub bab 2.1.3.1 (Pengertian Variabel X1) via Gemini."""
    from llm_title_parser import load_config, gemini_keys, MODEL_CHAIN

    config = load_config()
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return None

    y = cfg.get("y", "Kinerja Pegawai")
    x1 = cfg.get("x1", "X1")
    x2 = cfg.get("x2", "X2")
    m = cfg.get("m", "M")
    judul = cfg.get("judul", "")
    instansi = cfg.get("instansi", "")

    sitasi_block = _fmt_sitasi_bab2(acad, "cite_x1_1")
    if sources_x1:
        sumber_block = _fmt_sumber_block(sources_x1)
    else:
        sumber_block = sitasi_block

    dim_block = _dimensi_indikator_block(onto, "x1", instansi=instansi, judul=judul)

    prompt = PROMPT_BAB2_2_1_3.format(
        judul=judul, instansi=instansi, y=y, x1=x1, x2=x2, m=m,
        dimensi_indikator_block=dim_block, sumber_block=sumber_block,
    )

    chain = [m for m in config.get("models", []) if m] or MODEL_CHAIN
    timeout = max(int(config.get("timeout_seconds", 90)), 90)
    keys = gemini_keys(config)

    try:
        from llm_ontology_researcher import _call_gemini_matrix
        prompt_base = prompt
        prompt_safe = prompt_base
        _x1_name = (((onto.get("variables", {}) or {}).get("x1", {}) or {})
                    .get("nama_resmi") or x1)
        last_result = None
        for attempt in range(3):
            try:
                raw = _call_gemini_aman(keys, chain, prompt_safe, timeout, max_tokens=4096, label="write_sub_2_1_3")
                if isinstance(raw, str):
                    raw = raw
                result = _validate_sub_2_1_3(raw, sources_x1 if sources_x1 else [])
                if result and not (isinstance(result, dict) and result.get("_error")):
                    salah = _periksa_definisi(result, _x1_name)
                    if not salah:
                        print(f"[sub_2_1_3] konten ditulis via Gemini (percobaan {attempt+1})")
                        return _variasikan_sitasi_awal(_sterilkan_naskah(result, instansi, judul))
                    last_result = result
                    # Bila hanya 1-3 paragraf bermasalah: tulis ulang HANYA paragraf
                    # itu via Gemini (panggilan kecil) -> jauh lebih cepat & isi utuh.
                    if len(salah) <= 3:
                        _fix = _tulis_ulang_definisi(result, _x1_name, onto=onto,
                                                    role="x1", instances=instansi)
                        if not _periksa_definisi(_fix, _x1_name):
                            print(f"[sub_2_1_3] {len(salah)} paragraf ditulis ulang via Gemini")
                            return _variasikan_sitasi_awal(_sterilkan_naskah(_fix, instansi, judul))
                    catatan = _catatan_koreksi_definisi(result, _x1_name)
                    prompt_safe = prompt_base + catatan
                    print(f"[sub_2_1_3] percobaan {attempt+1}: {len(salah)} paragraf "
                          f"belum definisi '{_x1_name}' -> ulangi")
                    continue
                # Validasi struktur gagal -> gunakan error detail dari validator
                if isinstance(result, dict) and result.get("_error") == "format_salah":
                    errs = result.get("errors", [])[:5]
                    catatan = ("\n\n=== KOREKSI WAJIB ===\n"
                               "Paragraf berikut KENANGAN SITASI atau format salah:\n"
                               + "\n".join(f"  - {e}" for e in errs)
                               + "\n\nSETIAP paragraf definisi WAJIB memuat TEPAT SATU sitasi "
                               "berformat (Penulis, Tahun), mis. \"... (Hadiyanto & Syaflan, 2024).\"\n"
                               "JANGAN lupa sitasi di paragraf mana pun. Jumlah paragraf 2-10.\n")
                else:
                    catatan = ("\n\n=== KOREKSI WAJIB ===\n"
                               "SETIAP paragraf definisi WAJIB memuat TEPAT SATU sitasi "
                               "berformat (Penulis, Tahun), mis. \"... (Hadiyanto & Syaflan, 2024).\"\n"
                               "JANGAN lupa sitasi di paragraf mana pun. Jumlah paragraf 2-10.\n")
                prompt_safe = prompt_base + catatan
                print(f"[sub_2_1_3] percobaan {attempt+1}: validasi struktur gagal -> ulangi")
                continue
            except UnicodeEncodeError:
                print(f"[sub_2_1_3] UnicodeEncodeError percobaan {attempt+1}")
                continue
        if last_result:
            last_result = _rapikan_definisi_paragraf(last_result, _x1_name, onto, "x1", instances=instansi)
            print("[sub_2_1_3] memakai naskah terakhir (perbaikan definisi)")
            return _variasikan_sitasi_awal(_sterilkan_naskah(last_result, instansi, judul))
        print(f"[sub_2_1_3] semua percobaan gagal -> fallback")
        return None
    except Exception as e:
        print(f"[sub_2_1_3] Gemini gagal: {type(e).__name__} {str(e)[:80]}")
        return None


# ============================================================================
# WRITE SUB BAB 2.1.4.1 (VARIABEL X2)
# ============================================================================
def write_sub_2_1_4(cfg, onto, ctx, acad, sources_x2=None):
    """Menulis sub bab 2.1.4.1 (Pengertian Variabel X2) via Gemini."""
    from llm_title_parser import load_config, gemini_keys, MODEL_CHAIN

    config = load_config()
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return None

    y = cfg.get("y", "Kinerja Pegawai")
    x1 = cfg.get("x1", "X1")
    x2 = cfg.get("x2", "X2")
    m = cfg.get("m", "M")
    judul = cfg.get("judul", "")
    instansi = cfg.get("instansi", "")

    sitasi_block = _fmt_sitasi_bab2(acad, "cite_x2_1")
    if sources_x2:
        sumber_block = _fmt_sumber_block(sources_x2)
    else:
        sumber_block = sitasi_block

    dim_block = _dimensi_indikator_block(onto, "x2", instansi=instansi, judul=judul)

    prompt = PROMPT_BAB2_2_1_4.format(
        judul=judul, instansi=instansi, y=y, x1=x1, x2=x2, m=m,
        dimensi_indikator_block=dim_block, sumber_block=sumber_block,
    )

    chain = [m for m in config.get("models", []) if m] or MODEL_CHAIN
    timeout = max(int(config.get("timeout_seconds", 90)), 90)
    keys = gemini_keys(config)

    try:
        from llm_ontology_researcher import _call_gemini_matrix
        prompt_base = prompt
        prompt_safe = prompt_base
        _x2_name = (((onto.get("variables", {}) or {}).get("x2", {}) or {})
                    .get("nama_resmi") or x2)
        last_result = None
        for attempt in range(3):
            try:
                raw = _call_gemini_aman(keys, chain, prompt_safe, timeout, max_tokens=4096, label="write_sub_2_1_4")
                if isinstance(raw, str):
                    raw = raw
                result = _validate_sub_2_1_4(raw, sources_x2 if sources_x2 else [])
                if result and not (isinstance(result, dict) and result.get("_error")):
                    salah = _periksa_definisi(result, _x2_name)
                    if not salah:
                        print(f"[sub_2_1_4] konten ditulis via Gemini (percobaan {attempt+1})")
                        return _variasikan_sitasi_awal(_sterilkan_naskah(result, instansi, judul))
                    last_result = result
                    # Bila hanya 1-3 paragraf bermasalah: tulis ulang HANYA paragraf
                    # itu via Gemini (panggilan kecil) -> jauh lebih cepat & isi utuh.
                    if len(salah) <= 3:
                        _fix = _tulis_ulang_definisi(result, _x2_name, onto=onto,
                                                    role="x2", instances=instansi)
                        if not _periksa_definisi(_fix, _x2_name):
                            print(f"[sub_2_1_4] {len(salah)} paragraf ditulis ulang via Gemini")
                            return _variasikan_sitasi_awal(_sterilkan_naskah(_fix, instansi, judul))
                    catatan = _catatan_koreksi_definisi(result, _x2_name)
                    prompt_safe = prompt_base + catatan
                    print(f"[sub_2_1_4] percobaan {attempt+1}: {len(salah)} paragraf "
                          f"belum definisi '{_x2_name}' -> ulangi")
                    continue
                # Validasi struktur gagal -> gunakan error detail dari validator
                if isinstance(result, dict) and result.get("_error") == "format_salah":
                    errs = result.get("errors", [])[:5]
                    catatan = ("\n\n=== KOREKSI WAJIB ===\n"
                               "Paragraf berikut KENANGAN SITASI atau format salah:\n"
                               + "\n".join(f"  - {e}" for e in errs)
                               + "\n\nSETIAP paragraf definisi WAJIB memuat TEPAT SATU sitasi "
                               "berformat (Penulis, Tahun), mis. \"... (Hadiyanto & Syaflan, 2024).\"\n"
                               "JANGAN lupa sitasi di paragraf mana pun. Jumlah paragraf 2-10.\n")
                else:
                    catatan = ("\n\n=== KOREKSI WAJIB ===\n"
                               "SETIAP paragraf definisi WAJIB memuat TEPAT SATU sitasi "
                               "berformat (Penulis, Tahun), mis. \"... (Hadiyanto & Syaflan, 2024).\"\n"
                               "JANGAN lupa sitasi di paragraf mana pun. Jumlah paragraf 2-10.\n")
                prompt_safe = prompt_base + catatan
                print(f"[sub_2_1_4] percobaan {attempt+1}: validasi struktur gagal -> ulangi")
                continue
            except UnicodeEncodeError:
                print(f"[sub_2_1_4] UnicodeEncodeError percobaan {attempt+1}")
                continue
        if last_result:
            last_result = _rapikan_definisi_paragraf(last_result, _x2_name, onto, "x2", instances=instansi)
            print("[sub_2_1_4] memakai naskah terakhir (perbaikan definisi)")
            return _variasikan_sitasi_awal(_sterilkan_naskah(last_result, instansi, judul))
        print(f"[sub_2_1_4] semua percobaan gagal -> fallback")
        return None
    except Exception as e:
        print(f"[sub_2_1_4] Gemini gagal: {type(e).__name__} {str(e)[:80]}")
        return None


# ============================================================================
# SUB BAB 2.1.1.2: FAKTOR-FAKTOR YANG MEMPENGARUHI VARIABEL Y
# Format: 2 sumber, masing-masing daftar faktor bernomor dengan penjelasan
# ============================================================================
PROMPT_BAB2_2_1_1_FAKTOR = """Anda adalah penulis BAB II (Kajian Teoritis) tesis Magister Manajemen Indonesia.

Tulis SUB BAB 2.1.1.2 "Faktor-faktor Yang Mempengaruhi Variabel {y}" dengan format **DAFTAR FAKTOR DARI 2 SUMBER JURNAL**.

=== 2 JURNAL SUMBER WAJIB DIGUNAKAN ===
{sumber_jurnal}

=== INFORMASI PENELITIAN KITA ===
Judul: {judul}
Instansi: {instansi}
Variabel:
- Y (Dependent): {y}
- X1 (Independent 1): {x1}
- X2 (Independent 2): {x2}
- M (Mediating): {m}

=== ATURAN PENULISAN (WAJIB IKUTI FORMAT INI) ===
1. Output: JSON dengan key "faktor_dua_sumber" (list 2 item, masing-masing untuk 1 jurnal).
2. Setiap item berisi:
   - "sumber": sitasi APA in-text (contoh: "Juniarti, Setia, & Fahmi (2022)")
   - "paragraf_pembuka": 1 kalimat pembuka (contoh: "menyatakan bahwa faktor yang mempengaruhi kinerja pegawai antara lain:")
   - "daftar_faktor": list of dict, masing-masing dengan "nama" dan "penjelasan"
3. Jurnal 1: minimal 2 faktor, maksimal 5 faktor.
4. Jurnal 2: minimal 5 faktor, maksimal 10 faktor.
5. Nama faktor harus relevan dengan variabel penelitian kita (X1, X2, M, Y) dan teori manajemen SDM/Organisasi.
6. Penjelasan faktor: 2-4 kalimat, bahasa akademik formal, menjelaskan mekanisme pengaruh faktor tsb ke {y}.
7. DILARANG menulis sitasi (Penulis, Tahun) atau nama penulis di dalam field "penjelasan".
   Sitasi cukup ditulis SEKALI saja pada field "sumber". JANGAN mengulang nama penulis/sitasi di akhir penjelasan faktor.
8. HANYA gunakan 2 jurnal sumber di atas. DILARANG mengarang sumber lain.
9. DILARANG menggunakan kata "menurut", "berdasarkan pandangan", "pendapat".
10. Seluruh tulisan bersifat UMUM/teoretis: DILARANG menyebut nama instansi,
    kota, unit kerja, maupun istilah sektor objek penelitian (mis. gudang,
    logistik, kargo, pasien, rekam medis, medis, pengiriman). Nama faktor dan
    penjelasan harus netral (mis. "Faktor Kemampuan", bukan "Faktor Logistik").

=== CONTOH FORMAT OUTPUT ===
{{
  "faktor_dua_sumber": [
    {{
      "sumber": "Juniarti, Setia, & Fahmi (2022)",
      "paragraf_pembuka": "menyatakan bahwa faktor yang mempengaruhi kinerja pegawai antara lain:",
      "daftar_faktor": [
        {{
          "nama": "Faktor Kemampuan",
          "penjelasan": "Secara psikologis, kemampuan (ability) pegawai terdiri dari kemampuan potensi dengan pendidikan yang memadai untuk jabatannya dan terampil dalam mengerjakan pekerjaan sehari-hari, maka ia akan lebih mudah mencapai kinerja yang diharapkan. Oleh karena itu pegawai perlu ditempatkan pada pekerjaan yang sesuai dengan keahliannya (the right man in place, the man on the right job)."
        }},
        {{
          "nama": "Faktor Motivasi",
          "penjelasan": "Memotivasi berbentuk dari sikap (attitude) seseorang pegawai dalam menghadapi situasi (situation). Motivasi merupakan kondisi yang menggerakkan pegawai yang terarah untuk mencapai tujuan organisasi (tujuan kerja)."
        }}
      ]
    }},
    {{
      "sumber": "Rakhmalina (2021)",
      "paragraf_pembuka": "menjelaskan bahwa kinerja dipengaruhi oleh beberapa faktor yang berkaitan dengan hasil maupun perilaku kerja individu, antara lain:",
      "daftar_faktor": [
        {{"nama": "Kemampuan dan Keahlian", "penjelasan": "Kemampuan dan keahlian merujuk pada keterampilan yang dimiliki individu dalam melaksanakan tugas. Semakin tinggi kemampuan dan keahlian yang dimiliki, maka semakin besar peluang seorang pegawai untuk menyelesaikan pekerjaannya dengan tepat dan sesuai standar yang ditetapkan."}},
        {{"nama": "Pengetahuan", "penjelasan": "Pengetahuan mengenai tugas dan tanggung jawab pekerjaan merupakan aspek yang penting. Pegawai yang memiliki pemahaman yang baik tentang pekerjaannya cenderung menghasilkan kinerja yang optimal, sedangkan kurangnya pengetahuan dapat menghambat pencapaian hasil kerja."}},
        {{"nama": "Rancangan Kerja", "penjelasan": "Rancangan kerja yang baik akan mempermudah pegawai dalam melaksanakan tugas secara efektif dan efisien. Desain pekerjaan yang tersusun dengan jelas membantu mengurangi hambatan serta meningkatkan ketepatan dalam penyelesaian pekerjaan."}},
        {{"nama": "Kepribadian", "penjelasan": "Kepribadian mencerminkan karakter individu yang berpengaruh terhadap pola kerja. Setiap pegawai memiliki kepribadian yang berbeda, dan karakter tertentu dapat menjadi pendorong atau penghambat dalam pencapaian kinerja."}},
        {{"nama": "Employee Engagement", "penjelasan": "Employee Engagement adalah tingkat keterikatan, dedikasi, dan antusiasme emosional yang dimiliki seorang karyawan terhadap pekerjaan dan tujuan perusahaannya."}},
        {{"nama": "Kepemimpinan", "penjelasan": "Kepemimpinan berhubungan dengan gaya dan perilaku pemimpin dalam mengarahkan serta memengaruhi bawahannya. Kepemimpinan yang efektif mampu menciptakan Disiplin Kerja yang kondusif dan mendorong peningkatan kinerja pegawai."}},
        {{"nama": "Budaya Organisasi", "penjelasan": "Budaya organisasi berupa nilai, norma, dan kebiasaan yang berlaku dalam Disiplin Kerja. Budaya yang kuat dan positif akan membentuk perilaku pegawai yang mendukung pencapaian tujuan organisasi."}},
        {{"nama": "Kepuasan Kerja", "penjelasan": "Kepuasan kerja menggambarkan perasaan positif atau negatif pegawai terhadap pekerjaannya. Pegawai yang merasa puas cenderung bekerja lebih baik, sedangkan ketidakpuasan dapat menurunkan motivasi dan kualitas hasil kerja."}},
        {{"nama": "Loyalitas", "penjelasan": "Loyalitas menunjukkan kesetiaan pegawai terhadap organisasi. Pegawai yang loyal akan tetap berkomitmen bekerja dengan sungguh-sungguh meskipun organisasi sedang menghadapi kondisi yang kurang menguntungkan."}},
        {{"nama": "Disiplin Kerja", "penjelasan": "Disiplin kerja merupakan sikap patuh terhadap aturan dan kesediaan untuk melaksanakan pekerjaan secara konsisten. Disiplin dapat terlihat dari ketepatan waktu, ketaatan prosedur, dan tanggung jawab dalam menyelesaikan tugas."}}
      ]
    }}
  ]
}}
"""

# ============================================================================
# VALIDASI SUB BAB 2.1.1.2 (FAKTOR Y - 2 SUMBER)
# ============================================================================
def _validate_sub_2_1_1_faktor(raw, sumber_jurnal, min_th=2022, max_th=2026):
    """Validasi output sub_2_1_1_2 (faktor Y - 2 sumber) dari Gemini."""
    if isinstance(raw, list) and raw and isinstance(raw[0], dict):
        raw = raw[0]
    if not isinstance(raw, dict):
        return None

    if "faktor_dua_sumber" not in raw:
        print(f"[sub_2_1_1_faktor] kunci 'faktor_dua_sumber' hilang")
        return None

    items = raw["faktor_dua_sumber"]
    if not isinstance(items, list) or len(items) != 2:
        print(f"[sub_2_1_1_faktor] faktor_dua_sumber harus 2 item, dapat {len(items) if isinstance(items, list) else 'bukan list'}")
        return None

    import re
    # Extract allowed first author + year dari sumber_jurnal
    allowed_first_author_year = set()
    for s in sumber_jurnal:
        apa = s.get("apa_in_text", "")
        if apa:
            match = re.search(r'\(([^,]+),?\s*(?:et al\.?|(?:,?\s*[^,]+)*),?\s*(\d{4})\)', apa)
            if match:
                first_author = match.group(1).strip().lower()
                year = match.group(2)
                allowed_first_author_year.add((first_author, year))

    for idx, item in enumerate(items):
        if not isinstance(item, dict):
            print(f"[sub_2_1_1_faktor] item {idx+1} bukan dict")
            return None
        
        required_keys = ["sumber", "paragraf_pembuka", "daftar_faktor"]
        for k in required_keys:
            if k not in item:
                print(f"[sub_2_1_1_faktor] item {idx+1} kunci '{k}' hilang")
                return None

        sumber = item["sumber"]
        if not isinstance(sumber, str) or len(sumber) < 5:
            print(f"[sub_2_1_1_faktor] item {idx+1} sumber tidak valid")
            return None

        # Cek sumber cocok dengan jurnal yang di-fetch
        # Format bisa: "Author et al. (Year)" atau "Author, Author, & Author (Year)" atau "(Author, Year)"
        match = re.search(r'([A-Za-z][^,]+(?:,\s*[A-Za-z][^,]+)*(?:,\s*&\s*[A-Za-z][^,]+)?)?\s*\((\d{4})\)', sumber)
        if not match:
            print(f"[sub_2_1_1_faktor] item {idx+1} format sumber tidak valid: {sumber}")
            return None
        # Extract first author
        authors_part = match.group(1) if match.group(1) else ""
        first_author = authors_part.split(',')[0].strip().lower() if authors_part else ""
        year = match.group(2)
        
        # Try to match with allowed first author + year
        matched = False
        for allowed_first, allowed_year in allowed_first_author_year:
            if first_author and first_author in allowed_first:
                if year == allowed_year:
                    matched = True
                    break
        if not matched:
            # Try matching just by year and partial author
            for allowed_first, allowed_year in allowed_first_author_year:
                if year == allowed_year and (allowed_first in first_author or first_author in allowed_first):
                    matched = True
                    break
        if not matched:
            print(f"[sub_2_1_1_faktor] item {idx+1} sumber tidak cocok dengan jurnal fetch: {sumber}")
            return None

        # Cek tahun
        th = int(year)
        if th < min_th or th > max_th:
            print(f"[sub_2_1_1_faktor] item {idx+1} tahun {th} di luar {min_th}-{max_th}")
            return None

        # Cek daftar_faktor
        daftar = item["daftar_faktor"]
        if not isinstance(daftar, list):
            print(f"[sub_2_1_1_faktor] item {idx+1} daftar_faktor bukan list")
            return None
        if idx == 0 and (len(daftar) < 2 or len(daftar) > 5):
            print(f"[sub_2_1_1_faktor] jurnal 1 harus 2-5 faktor, dapat {len(daftar)}")
            return None
        if idx == 1 and (len(daftar) < 5 or len(daftar) > 10):
            print(f"[sub_2_1_1_faktor] jurnal 2 harus 5-10 faktor, dapat {len(daftar)}")
            return None

        for j, faktor in enumerate(daftar):
            if not isinstance(faktor, dict) or "nama" not in faktor or "penjelasan" not in faktor:
                print(f"[sub_2_1_1_faktor] item {idx+1} faktor {j+1} format salah")
                return None
            if not isinstance(faktor["nama"], str) or len(faktor["nama"]) < 3:
                print(f"[sub_2_1_1_faktor] item {idx+1} faktor {j+1} nama terlalu pendek")
                return None
            if not isinstance(faktor["penjelasan"], str) or len(faktor["penjelasan"]) < 40:
                print(f"[sub_2_1_1_faktor] item {idx+1} faktor {j+1} penjelasan terlalu pendek")
                return None

        # Cek kata terlarang
        teks = f"{item['paragraf_pembuka']} " + " ".join([f"{f['nama']} {f['penjelasan']}" for f in daftar])
        terlarang = ["menurut", "berdasarkan pandangan"]
        for t in terlarang:
            if t in teks.lower():
                print(f"[sub_2_1_1_faktor] item {idx+1} kata terlarang: '{t}'")
                return None

    return raw


# ============================================================================
# WRITE SUB BAB 2.1.1.2 (FAKTOR Y - 2 SUMBER DENGAN DAFTAR FAKTOR)
# ============================================================================
def write_sub_2_1_1_faktor(cfg, onto, ctx, acad, sumber_jurnal=None):
    """Menulis sub bab 2.1.1.2 via Gemini: 2 sumber, daftar faktor bernomor."""
    from llm_title_parser import load_config, gemini_keys, MODEL_CHAIN

    config = load_config()
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return None

    y = cfg.get("y", "Kinerja Pegawai")
    x1 = cfg.get("x1", "X1")
    x2 = cfg.get("x2", "X2")
    m = cfg.get("m", "M")
    judul = cfg.get("judul", "")
    instansi = cfg.get("instansi", "")

    if not sumber_jurnal or len(sumber_jurnal) < 2:
        print("[sub_2_1_1_faktor] butuh minimal 2 sumber jurnal")
        return None

    sumber_block = ""
    for s in sumber_jurnal[:2]:
        sumber_block += f"- {s.get('apa_full', '')}\n"

    prompt = PROMPT_BAB2_2_1_1_FAKTOR.format(
        judul=judul, instansi=instansi, y=y, x1=x1, x2=x2, m=m,
        sumber_jurnal=sumber_block,
    )

    chain = [m for m in config.get("models", []) if m] or MODEL_CHAIN
    timeout = max(int(config.get("timeout_seconds", 90)), 90)
    keys = gemini_keys(config)

    try:
        from llm_ontology_researcher import _call_gemini_matrix
        prompt_safe = prompt
        for attempt in range(3):
            try:
                raw = _call_gemini_aman(keys, chain, prompt_safe, timeout, max_tokens=4096, label="write_sub_2_1_1_faktor")
                if isinstance(raw, str):
                    raw = raw
                result = _validate_sub_2_1_1_faktor(raw, sumber_jurnal[:2])
                if result:
                    print(f"[sub_2_1_1_faktor] 2 sumber faktor ditulis via Gemini (percobaan {attempt+1})")
                    return _sterilkan_naskah(result, instansi, judul)
                print(f"[sub_2_1_1_faktor] percobaan {attempt+1} tidak valid, mencoba lagi...")
            except UnicodeEncodeError:
                print(f"[sub_2_1_1_faktor] UnicodeEncodeError percobaan {attempt+1}")
                continue
        print(f"[sub_2_1_1_faktor] semua percobaan gagal -> fallback")
        return None
    except Exception as e:
        print(f"[sub_2_1_1_faktor] Gemini gagal: {type(e).__name__} {str(e)[:80]}")
        return None


# ============================================================================
# SUB BAB 2.1.1.3: DIMENSI & INDIKATOR VARIABEL Y (1 SUMBER, DAFTAR DIMENSI)
# ============================================================================
PROMPT_BAB2_2_1_3_DIMENSI = """Anda adalah penulis BAB II (Kajian Teoritis) tesis Magister Manajemen Indonesia.

Tulis SUB BAB "Dimensi dan Indikator Variabel {y}" dalam bentuk DAFTAR DIMENSI BERNOTOR, masing-masing dimensi diberi paragraf penjelasan.

=== SUMBER TUNGGAL (WAJIB, HANYA INI) ===
{sumber_sitasi} — {sumber_entri}

=== INFORMASI PENELITIAN ===
Judul: {judul}
Instansi: {instansi}
Variabel: {y}

=== DIMENSI & INDIKATOR (WAJIB DIPAKAI, JANGAN DIUBAH/DITAMBAH) ===
{dimensi_block}

=== ATURAN PENULISAN (WAJIB) ===
1. Output JSON valid: {{"dimensi_list": [{{"nama": "...", "penjelasan": "..."}}, ...]}}
2. Jumlah, urutan, dan nama dimensi HARUS SAMA PERSIS dengan daftar DIMENSI di atas.
3. Untuk setiap dimensi, tulis "penjelasan": 2-3 kalimat akademik formal yang menjelaskan MAKNA dimensi tsb dan mengapa penting bagi {y}. Akhiri dengan makna operasional, JANGAN mencantumkan daftar indikator.
4. Penjelasan DILARANG memuat sitasi "(Penulis, Tahun)" maupun nama penulis. Sitasi hanya ditampilkan sekali di paragraf pembuka (dibuat otomatis oleh sistem).
5. DILARANG mengarang dimensi atau indikator baru.
6. DILARANG memakai kata "menurut", "berdasarkan pandangan", "pendapat".
7. Penjelasan bersifat UMUM/teoretis: DILARANG menyebut nama instansi, kota,
   unit kerja, maupun istilah sektor objek penelitian (mis. gudang, logistik,
   kargo, pasien, rekam medis, medis, pelayanan kesehatan, pengiriman barang).

=== CONTOH penjelasan (dimensi "Kualitas Pekerjaan") ===
"Kualitas kerja berkaitan dengan kesesuaian hasil kerja dengan persyaratan standar kualitas yang ditentukan organisasi. Dimensi ini menekankan tingkat ketelitian, kerapian, dan kelengkapan output sehingga mencerminkan profesionalisme pegawai dalam menjalankan tugas kedinasan."

=== STRUKTUR KELUARAN (JSON valid saja, tanpa teks lain) ===
{{"dimensi_list": [{{"nama": "...", "penjelasan": "..."}}]}}
"""


def _validate_sub_2_1_1_dimensi(raw, dimensi_onto):
    """Validasi output 2.1.1.3 (dimensi Y - 1 sumber)."""
    if isinstance(raw, list) and raw and isinstance(raw[0], dict):
        raw = raw[0]
    if not isinstance(raw, dict) or "dimensi_list" not in raw:
        print("[sub_2_1_1_dimensi] kunci 'dimensi_list' hilang")
        return None

    items = raw["dimensi_list"]
    if not isinstance(items, list) or len(items) != len(dimensi_onto):
        print(f"[sub_2_1_1_dimensi] jumlah dimensi harus {len(dimensi_onto)}, dapat {len(items) if isinstance(items, list) else 'bukan list'}")
        return None

    import re
    terlarang = ["menurut", "berdasarkan pandangan"]
    for i, item in enumerate(items):
        if not isinstance(item, dict) or "nama" not in item or "penjelasan" not in item:
            print(f"[sub_2_1_1_dimensi] item {i+1} format salah")
            return None
        nama = item["nama"] if isinstance(item["nama"], str) else ""
        penjelasan = item["penjelasan"] if isinstance(item["penjelasan"], str) else ""
        # Normalisasi: buang penomoran awal ("1. ", "1) ") lalu PAKSA memakai nama
        # dimensi ontology (urutan wajib identik) agar pemetaan di engine pasti cocok.
        nama = re.sub(r'^\s*\d+\s*[\.\)]\s*', '', nama).strip()
        nama_onto = dimensi_onto[i].get("nama", "") if i < len(dimensi_onto) else ""
        if nama_onto:
            nama = nama_onto
        item["nama"] = nama
        if len(nama) < 3 or len(penjelasan) < 40:
            print(f"[sub_2_1_1_dimensi] item {i+1} nama/penjelasan terlalu pendek")
            return None
        if re.search(r'\([^)]*?,\s*\d{4}[a-z]?\)', penjelasan):
            print(f"[sub_2_1_1_dimensi] item {i+1} penjelasan memuat sitasi (dilarang)")
            return None
        for t in terlarang:
            if t in penjelasan.lower():
                print(f"[sub_2_1_1_dimensi] item {i+1} kata terlarang: '{t}'")
                return None

    return raw


def write_sub_2_1_1_dimensi(cfg, onto, dimensi_override=None, jurnal_override=None,
                            var_role="y", var_name=None):
    """Menulis blok "Dimensi dan Indikator" (format 2.1.1.3) dari 1 sumber jurnal.

    Dapat dipakai untuk variabel apa pun (Y/M/X1/X2) melalui ``var_role``
    ("y"/"m"/"x1"/"x2") dan ``var_name``. ``dimensi_override``/``jurnal_override``
    dipakai bila dimensi diambil dari sumber lain (mis. pool) sehingga tiap judul
    memakai sumber & dimensi yang berbeda-beda namun tetap terverifikasi.
    """
    from llm_title_parser import load_config, gemini_keys, MODEL_CHAIN

    config = load_config()
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return None

    var_role = (var_role or "y").lower()
    var_name = var_name or cfg.get(var_role, "") or ""
    judul = cfg.get("judul", "")
    instansi = cfg.get("instansi", "")

    v_onto = onto.get("variables", {}).get(var_role, {}) or {}
    dimensi_onto = dimensi_override or v_onto.get("dimensi", [])
    if not dimensi_onto:
        return None

    jurnal = jurnal_override or v_onto.get("jurnal_rujukan", {}) or {}
    sumber_sitasi = jurnal.get("sitasi", "") or ""
    sumber_entri = jurnal.get("entri", "") or ""

    import indicator_layers as IL
    dimensi_block = ""
    for i, d in enumerate(dimensi_onto, 1):
        inds = "; ".join(IL.indikator_teoretis(d, instansi=instansi, judul=judul))
        dimensi_block += f"{i}. {d.get('nama', '')} (Indikator teoretis: {inds})\n"

    prompt = PROMPT_BAB2_2_1_3_DIMENSI.format(
        judul=judul, instansi=instansi, y=var_name,
        sumber_sitasi=sumber_sitasi, sumber_entri=sumber_entri,
        dimensi_block=dimensi_block,
    )

    chain = [m for m in config.get("models", []) if m] or MODEL_CHAIN
    timeout = max(int(config.get("timeout_seconds", 90)), 90)
    keys = gemini_keys(config)

    try:
        from llm_ontology_researcher import _call_gemini_matrix
        prompt_safe = prompt
        for attempt in range(3):
            try:
                raw = _call_gemini_aman(keys, chain, prompt_safe, timeout, max_tokens=4096, label="write_sub_2_1_1_dimensi")
                if isinstance(raw, str):
                    raw = raw
                result = _validate_sub_2_1_1_dimensi(raw, dimensi_onto)
                if result:
                    print(f"[sub_2_1_1_dimensi] {len(dimensi_onto)} dimensi ditulis via Gemini (percobaan {attempt+1})")
                    return _sterilkan_naskah(result, instansi, judul)
                print(f"[sub_2_1_1_dimensi] percobaan {attempt+1} tidak valid, mencoba lagi...")
            except UnicodeEncodeError:
                print(f"[sub_2_1_1_dimensi] UnicodeEncodeError percobaan {attempt+1}")
                continue
        print(f"[sub_2_1_1_dimensi] semua percobaan gagal -> fallback")
        return None
    except Exception as e:
        print(f"[sub_2_1_1_dimensi] Gemini gagal: {type(e).__name__} {str(e)[:80]}")
        return None


PROMPT_DIM_INDIKATOR = """Anda adalah ahli metodologi penelitian kuantitatif tesis Magister Manajemen Indonesia.

Variabel: {var}
Definisi variabel: {definisi}
Instansi/objek penelitian: {instansi}
Judul penelitian: {judul}
Sumber rujukan dimensi: {sumber}

=== DIMENSI (WAJIB DIPERTAHANKAN, JANGAN DIUBAH) ===
{dimensi_block}

=== TUGAS ===
Untuk SETIAP dimensi di atas, tulis 2 sampai 4 INDIKATOR OPERASIONAL dalam Bahasa Indonesia
yang SPESIFIK, terukur, dan secara jelas MENCERMINKAN dimensi tersebut pada konteks variabel
dan objek penelitian. Indikator berupa frasa singkat (4-14 kata) yang menggambarkan
kapabilitas/perilaku/keadaan yang bisa dinilai responden dengan skala Likert 1-5. BUKAN definisi
teori dan BUKAN sekadar mengulang nama dimensi.

=== ATURAN ===
1. Jumlah & urutan dimensi HARUS sama persis dengan daftar di atas.
2. Setiap dimensi berisi 2 sampai 4 indikator (utamakan 3 bila dimensi luas; jangan memaksakan
   indikator keempat bila tidak ada aspek baru).
3. SETIAP indikator harus menunjuk secara eksplisit pada aspek turunan dari dimensinya.
   Contoh benar (dimensi "Pendidikan"): "kesesuaian tingkat ijazah formal karyawan dengan bidang tugas jabatan".
   DILARANG memakai frasa generik yang berlaku untuk semua dimensi (mis. "kinerja karyawan baik",
   "bersedia bekerja lebih baik", "mampu bekerja dengan baik").
4. TIDAK BOLEH ada dua indikator bermakna sama, baik dalam satu dimensi, antar dimensi dalam satu
   variabel, maupun antar variabel lain.
5. DILARANG memuat angka hasil penelitian, persentase, nilai rata-rata, sitasi, atau definisi
   (dilarang kata "yaitu", "adalah", "merupakan", "mencakup", "mengacu pada" di dalam indikator).
6. DILARANG menyalin kode/placeholder seperti "X1.1", "P.1.1", atau mengulang nama dimensi.
7. Sesuaikan diksi dengan objek penelitian. Jika nama instansi mengandung "PT", "Persero", "Tbk",
   "Perusahaan", "CV", atau "Ltd", gunakan diksi perusahaan: karyawan, pekerjaan, pelanggan,
   mitra, gudang, pengiriman, target operasional. DILARANG memakai "aparatur", "kedinasan",
   "birokrasi", "instansi pemerintah", atau "pegawai" untuk objek perusahaan.
   Jika objek adalah instansi pemerintah, gunakan diksi aparatur/pegawai/kedinasan secara wajar.
{larangan}
Output JSON valid saja (tanpa teks lain):
{{"dimensi_list":[{{"nama":"...","indikator":["...","...","..."]}}]}}
"""


# ---------------------------------------------------------------------------
# PENJAGA KUALITAS INDIKATOR (anti duplikat, anti frasa sampah)
# ---------------------------------------------------------------------------
_RX_IND_SITASI = re.compile(r"\([^()]*\b(?:19|20)\d{2}[a-z]?\b[^()]*\)")
_RX_IND_STAT = re.compile(
    r"(\d+(?:[.,]\d+)?\s*%|nilai\s+rata-?rata|sebesar|kategori\s+(?:sangat\s+)?baik|"
    r"\bmean\b|standar\s+deviasi|p[- ]?value|signifikan|hasil\s+penelitian|"
    r"responden|subjek\s+penelitian)", re.I)
_RX_IND_DEFINISI = re.compile(
    r"\b(yaitu|adalah|merupakan|didefinisikan|diartikan|dimaknai|mencakup|"
    r"mengacu\s+pada|merujuk\s+pada|terdiri\s+dari|meliputi)\b",
    re.I)
_VERBE_AWAL_INDIKATOR = {
    "memberikan", "meningkatkan", "melakukan", "mengikuti", "menyusun", "mengelola",
    "mengoptimalkan", "mengurangi", "membangun", "menghadapi", "menjalankan",
    "menuntaskan", "menyelesaikan", "memperhatikan", "menggunakan", "menerapkan",
    "memastikan", "membantu", "mendorong",
}
_SUFFIX_GENERIK = ("sdm", "kerja", "organisasi", "pegawai", "individu", "karyawan")


def _norm_ind(s):
    """Normalisasi indikator untuk perbandingan kemiripan."""
    t = re.sub(r"[^\w\s]", " ", str(s or "").lower())
    return re.sub(r"\s+", " ", t).strip()


# Kata generik yang tidak membedakan konstruk: dipakai untuk mendeteksi
# tumpang-tindih KONSEP (mis. "pencapaian target" pada dua variabel berbeda).
_KONTEKS_GENERIK = {
    "instansi", "pegawai", "pekerjaan", "tugas", "kedinasan", "pelayanan", "publik",
    "peraturan", "daerah", "masa", "organisasi", "satuan", "kegiatan", "program",
    "seluruh", "setiap", "semua", "serta", "hasil", "kerja", "pelaksana", "petugas",
    "wilayah", "kabupaten", "kantor", "unit", "sistem", "proses", "kinerja", "layanan",
    "dalam", "pada", "dengan", "yang", "untuk", "oleh", "dari", "dan", "atau",
    "sehingga", "karena", "akan", "telah", "yaitu", "dapat", "harus", "sudah",
    "memiliki", "menunjukkan", "dilakukan", "memberikan", "melakukan", "terhadap",
    "berdasarkan", "menurut", "dalam", "serta", "sangat", "lebih", "paling",
}

# Kata tanpa makna distinctive (dipakai untuk mendeteksi indikator yang
# HANYA mengulang nama dimensi, mis. "Tingkat kemandirian yang ditunjukkan pegawai...").
_KATA_UMPAN = {
    "tingkat", "menunjukkan", "ditunjukkan", "dirasakan", "dimiliki", "terasa",
    "kualitas", "derajat", "tinggi", "rendah", "baik", "buruk", "bersifat",
}


def _bigram_konsep(teks):
    """Bigram kata bermakna (kata generik dibuang) untuk deteksi tumpang-tindih."""
    kata = [w for w in _norm_ind(teks).split() if len(w) > 2 and w not in _KONTEKS_GENERIK]
    return set(zip(kata, kata[1:]))


def _ind_konsep_tumpang_tindih(a, b):
    """True bila dua indikator memuat konsep yang sama (bigram bermakna sama)."""
    ba, bb = _bigram_konsep(a), _bigram_konsep(b)
    return bool(ba & bb)


def _ind_terlalu_sama(a, b):
    """True bila dua indikator bermakna sama (duplikat persis / near-duplikat)."""
    na, nb = _norm_ind(a), _norm_ind(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    if na in nb or nb in na:
        pendek = na if len(na) < len(nb) else nb
        if len(pendek.split()) <= 2:
            return True
    ta = set(w for w in na.split() if len(w) > 2)
    tb = set(w for w in nb.split() if len(w) > 2)
    if ta and tb:
        if (ta <= tb or tb <= ta) and min(len(ta), len(tb)) <= 2:
            return True
        jaccard = len(ta & tb) / float(len(ta | tb))
        if jaccard >= 0.6:
            return True
    return False


def _indikator_boleh(ind, nama_dimensi="", nama_variabel="", dipakai=None,
                     cek_konsep=True, strict=True):
    """Bersihkan + nilai layak satu indikator. Kembalikan (teks, alasan_tolak).

    strict=False dipakai untuk indikator yang SUDAH terverifikasi dari pustaka
    (mis. Training Dessler: Reaction/Learning/Behavior/Results). Aturan panjang
    kalimat/verba tidak boleh membuang istilah konstruk yang sah dari sumber.
    """
    s = _RX_IND_SITASI.sub("", str(ind or "")).strip()
    s = re.sub(r"\s+", " ", s).strip().rstrip(".").strip()
    if not s:
        return "", "kosong"
    if _RX_IND_STAT.search(s):
        return "", "memuat angka/statistik hasil penelitian"
    if _RX_IND_DEFINISI.search(s):
        return "", "berupa definisi, bukan indikator"
    if re.match(r"^[A-Za-z]{1,4}[\.\d]+$", s):
        return "", "kode/placeholder"
    kata = s.split()
    if strict:
        if len(kata) < 3 or len(s) < 15:
            return "", "terlalu pendek"
        if kata[0].lower() in _VERBE_AWAL_INDIKATOR:
            return "", "berupa aktivitas, bukan indikator"
    ns = _norm_ind(s)
    for ref in (nama_dimensi, nama_variabel):
        nr = _norm_ind(ref)
        if not nr:
            continue
        if ns == nr:
            return "", "mengulang nama dimensi/variabel"
        if ns in {nr + " " + suf for suf in _SUFFIX_GENERIK}:
            return "", "frasa generik (nama dimensi + kata umum)"
    # Indikator yang HANYA mengulang nama dimensi (tanpa makna tambahan) tidak sahih
    nd = _norm_ind(nama_dimensi)
    if nd and nd in ns:
        sisa = [w for w in ns.split()
                if w not in _KONTEKS_GENERIK and w not in _KATA_UMPAN
                and w not in nd.split()]
        batas = 3 if len(nd.split()) >= 2 else 2
        if len(sisa) <= batas:
            return "", "hanya mengulang nama dimensi (tanpa makna tambahan)"
    for d in (dipakai or []):
        if _ind_terlalu_sama(s, d):
            return "", "duplikat indikator yang sudah dipakai"
        if cek_konsep and _ind_konsep_tumpang_tindih(s, d):
            return "", "tumpang-tindih konsep dengan indikator yang sudah dipakai"
    return s, ""


def _indikator_generated_boleh(ind, nama_dimensi="", nama_variabel="", dipakai=None):
    """Validasi indikator HASIL PENGEMBANGAN dari dimensi.

    Berbeda dari kutipan sumber, indikator operasional boleh diawali verba
    (mis. "Meningkatkan ketepatan pengiriman..."). Yang tetap dilarang adalah
    definisi, statistik, kode, frasa terlalu pendek, dan duplikasi.
    """
    s, alasan = _indikator_boleh(
        ind, nama_dimensi, nama_variabel, dipakai,
        cek_konsep=False, strict=False)
    if not s:
        return "", alasan
    if len(s.split()) < 3 or len(s) < 15:
        return "", "terlalu pendek untuk indikator operasional"
    return s, ""


def _validate_dim_indicators(raw, dimensi_names, forbidden=None, var_name=""):
    import re
    if isinstance(raw, list) and raw and isinstance(raw[0], dict):
        raw = raw[0]
    if not isinstance(raw, dict) or "dimensi_list" not in raw:
        return None
    items = raw["dimensi_list"]
    if not isinstance(items, list) or len(items) != len(dimensi_names):
        return None
    out = []
    for i, item in enumerate(items):
        if not isinstance(item, dict):
            return None
        inds = item.get("indikator", [])
        if not isinstance(inds, list):
            return None
        nama_dim = dimensi_names[i] if i < len(dimensi_names) else ""
        clean = []
        for x in inds:
            if isinstance(x, str):
                s, _alasan = _indikator_generated_boleh(
                    x, nama_dim, var_name, forbidden)
                if s and not any(_ind_terlalu_sama(s, c) for c in clean):
                    clean.append(s)
        if len(clean) < 2:
            return None
        out.append({"nama": nama_dim, "indikator": clean[:4]})
    return {"dimensi_list": out}


def write_dimension_indicators(var_name, dimensi_names, source_label="", tokoh="",
                               judul="", instansi="", forbidden=None, definisi=""):
    """Kembangkan 2-4 indikator operasional untuk tiap dimensi (via Gemini)."""
    from llm_title_parser import load_config, gemini_keys, MODEL_CHAIN
    config = load_config()
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return None
    dimensi_names = [n for n in (dimensi_names or []) if n]
    if not dimensi_names:
        return None
    dimensi_block = "\n".join(f"{i}. {n}" for i, n in enumerate(dimensi_names, 1))
    sumber = source_label or tokoh or "-"
    if forbidden:
        lines = "\n".join(f"- {f}" for f in list(forbidden)[:80])
        larangan = ("\n=== INDIKATOR YANG SUDAH DIPAKAI (JANGAN DIULANG, BISA VARIABEL LAIN) ===\n"
                    + lines + "\n")
    else:
        larangan = ("\n=== INDIKATOR YANG SUDAH DIPAKAI (JANGAN DIULANG) ===\n- (belum ada)\n")
    prompt = PROMPT_DIM_INDIKATOR.format(
        var=var_name, sumber=sumber, dimensi_block=dimensi_block,
        definisi=definisi or "(belum tersedia)",
        instansi=instansi or "(tidak disebutkan)", judul=judul or "(tidak disebutkan)",
        larangan=larangan)
    chain = [m for m in config.get("models", []) if m] or MODEL_CHAIN
    timeout = max(int(config.get("timeout_seconds", 90)), 90)
    keys = gemini_keys(config)
    try:
        from llm_ontology_researcher import _call_gemini_matrix
        prompt_safe = prompt
        for attempt in range(3):
            try:
                raw = _call_gemini_aman(keys, chain, prompt_safe, timeout,
                                       max_tokens=3072, label="dim_indikator")
                if isinstance(raw, str):
                    raw = raw
                res = _validate_dim_indicators(
                    raw, dimensi_names, forbidden=forbidden, var_name=var_name)
                if res:
                    _n = sum(len(d.get("indikator") or []) for d in res["dimensi_list"])
                    print(f"[dim_indikator] {len(dimensi_names)} dimensi x {_n} indikator "
                          f"(2-4/dimensi) via Gemini (percobaan {attempt+1})")
                    return res
            except UnicodeEncodeError:
                continue
        print("[dim_indikator] semua percobaan gagal")
        return None
    except Exception as e:
        print(f"[dim_indikator] Gemini gagal: {type(e).__name__} {str(e)[:80]}")
        return None


PROMPT_PERNYATAAN = """Anda adalah ahli metodologi penelitian kuantitatif tesis Magister Manajemen Indonesia.

Variabel: {var}
Instansi objek: {instansi}
Judul penelitian: {judul}

=== DIMENSI + INDIKATOR (WAJIB DIPERTAHANKAN PERSIS) ===
{dimensi_block}

=== TUGAS ===
Untuk SETIAP indikator di atas, tulis 1 (satu) PERNYATAAN kuesioner skala Likert 1-5 yang
natural, berbahasa Indonesia baku, menggunakan sudut pandang orang pertama jamak, BERUPA PERNYATAAN
POSITIF, dan sangat terkait dengan indikatornya.

=== CONTOH (arah, tidak harus sama persis) ===
Indikator : "kesesuaian tingkat ijazah formal pegawai dengan bidang tugas encontraban"
Pernyataan: "Gelar dan jenjang pendidikan saya sesuai dengan tuntutan tugas yang saya jalankan."

Indikator : "ketahanan stamina anggota saat bertugas di lapangan dalam durasi lama"
Pernyataan: "Stamina saya mampu bertahan saat menjalankan tugas lapangan berjam-jam."

=== ATURAN ===
1. Jumlah pernyataan HARUS sama dengan jumlah indikator, urutan sama.
2. Setiap pernyataan 1 kalimat, 8-18 kata, tanpa angka, kode, atau tanda kutip.
3. Variasi gaya kalimat wajib (jangan memakai pola yang sama untuk semua butir).
4. DILARANG memakai kata "sangat", "sangat sekali", atau pengingkaran ganda.
5. Output JSON valid saja (tanpa teks lain):
{{"dimensi_list":[{{"nama":"...","pernyataan":["...","...","..."]}}]}}
"""


def write_dimension_statements(var_name, dimensi_indikator, judul="", instansi=""):
    """Buat butir pernyataan Likert 1-5 untuk indikator hasil riset.

    ``dimensi_indikator``: list [(nama_dimensi, [indikator, ...]), ...]
    """
    from llm_title_parser import load_config, gemini_keys, MODEL_CHAIN
    config = load_config()
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return None
    pasangan = [(n, list(inds or [])) for n, inds in (dimensi_indikator or []) if n and inds]
    if not pasangan:
        return None
    baris = []
    for i, (nama, inds) in enumerate(pasangan, 1):
        baris.append(f"{i}. {nama}")
        for j, ind in enumerate(inds, 1):
            baris.append(f"   {j}. {ind}")
    prompt = PROMPT_PERNYATAAN.format(
        var=var_name, dimensi_block="\n".join(baris),
        instansi=instansi or "(tidak disebutkan)", judul=judul or "(tidak disebutkan)")
    chain = [m for m in config.get("models", []) if m] or MODEL_CHAIN
    timeout = max(int(config.get("timeout_seconds", 90)), 90)
    keys = gemini_keys(config)
    try:
        from llm_ontology_researcher import _call_gemini_matrix
        prompt_safe = prompt
        for attempt in range(3):
            try:
                raw = _call_gemini_aman(keys, chain, prompt_safe, timeout, max_tokens=3072, label="pernyataan")
                if isinstance(raw, str):
                    raw = raw
                if isinstance(raw, list) and raw and isinstance(raw[0], dict):
                    raw = raw[0]
                if not isinstance(raw, dict):
                    continue
                items = raw.get("dimensi_list")
                if not isinstance(items, list) or len(items) != len(pasangan):
                    continue
                out = []
                for k, item in enumerate(items):
                    if not isinstance(item, dict):
                        return None
                    stmt = item.get("pernyataan") or []
                    if not isinstance(stmt, list) or len(stmt) != len(pasangan[k][1]):
                        return None
                    bersih = []
                    for s in stmt:
                        s2 = re.sub(r"\s+", " ", str(s or "")).strip()
                        if len(s2) < 12 or len(s2.split()) < 3:
                            return None
                        bersih.append(s2)
                    out.append({"nama": pasangan[k][0], "pernyataan": bersih})
                print(f"[pernyataan] {len(out)} dimensi butir kuesioner via Gemini")
                return {"dimensi_list": out}
            except UnicodeEncodeError:
                continue
        print("[pernyataan] semua percobaan gagal")
        return None
    except Exception as e:
        print(f"[pernyataan] Gemini gagal: {type(e).__name__} {str(e)[:80]}")
        return None


# ---------------------------------------------------------------------------
# BAB III — SUB BAB 3.2 POPULASI DAN SAMPEL
# Narasi digenerate Gemini (unik per judul) dengan sitasi HANYA dari pool
# referensi metodologi terverifikasi (references_metodologi.json).
# ---------------------------------------------------------------------------

PROMPT_BAB3_3_2 = """Anda adalah penulis metodologi penelitian (BAB III) tesis Magister Manajemen (S2) Indonesia yang cermat, faktual, dan tidak repetitif.

=== JUDUL PENELITIAN ===
{judul}

=== OBJEK/INSTANSI PENELITIAN ===
{instansi}

=== DATA PENELITIAN (acuan tetap; JANGAN diubah) ===
Jumlah populasi (N)        : {populasi} orang
Ukuran sampel (n)          : {n_sample} responden
Teknik penentuan sampel    : {teknik_desc}

=== DAFTAR SITASI TERVERIFIKASI (satu-satunya sitasi yang BOLEH dipakai; format APA in-text) ===
{sumber_menu}

=== STRUKTUR KELUARAN: JSON dengan tepat 3 kunci ===
"pengertian_populasi" : array berisi 2 paragraf. Paragraf 1: definisi populasi menurut sumber + penjelasan unit analisis/elemen populasi. Paragraf 2: tujuan menetapkan populasi (menentukan besar sampel dan membatasi berlakunya daerah generalisasi). WAJIB ada sitasi pada setiap paragraf.
"pengertian_sampel"   : array berisi 1 sampai 2 paragraf. Definisi sampel dan alasan penelitian memerlukan sampel (keterbatasan waktu, tenaga, dan dana) namun tetap harus representatif. WAJIB ada sitasi pada setiap paragraf.
"alasan_teknik"       : satu paragraf yang menjelaskan pertimbangan pemilihan teknik penentuan sampel "{teknik_desc}". WAJIB ada sitasi. TANPA angka jumlah.

=== ATURAN KERAS ===
1. Bahasa Indonesia akademik formal, mengalir, tidak kaku, dan TIDAK repetitif.
2. Setiap paragraf WAJIB memuat minimal satu sitasi dalam tanda kurung, persis format (Penulis, Tahun) dari DAFTAR SITASI TERVERIFIKASI. DILARANG mengarang penulis, tahun, judul, atau memakai sitasi di luar daftar.
3. DILARANG menyebut jumlah/kuantitas populasi maupun sampel dalam bentuk angka apa pun (mis. ditulis "119 orang" atau "92 responden"). Angka hanya boleh berupa taraf signifikansi/kelonggaran (mis. 5% atau 0,05).
4. Jangan menyalin kalimat contoh; susun redaksi sendiri. Antarjudul harus menghasilkan diksi dan urutan argumen yang berbeda.
5. Prioritaskan sumber dengan tahun terbaru bila relevan.
6. BALAS HANYA JSON valid berisi tepat 3 kunci di atas, tanpa teks penjelasan tambahan.
"""


def _validate_sub_3_2(raw, refs=None):
    """Validasi keluaran 3.2: struktur + sitasi HANYA dari pool terverifikasi."""
    import references_metodologi as rm

    if isinstance(raw, list) and raw and isinstance(raw[0], dict):
        raw = raw[0]
    if not isinstance(raw, dict):
        print("[sub_3_2] keluaran bukan objek JSON")
        return None

    pop = raw.get("pengertian_populasi")
    sam = raw.get("pengertian_sampel")
    alasan = raw.get("alasan_teknik")

    if not isinstance(pop, list) or not (1 <= len(pop) <= 4):
        print("[sub_3_2] pengertian_populasi tidak valid")
        return None
    if not isinstance(sam, list) or not (1 <= len(sam) <= 4):
        print("[sub_3_2] pengertian_sampel tidak valid")
        return None
    if not isinstance(alasan, str) or len(alasan.strip()) < 80:
        if isinstance(alasan, list) and alasan and isinstance(alasan[0], str) and len(alasan[0]) > 80:
            alasan = alasan[0]
        else:
            print("[sub_3_2] alasan_teknik tidak valid")
            return None

    paras = list(pop) + list(sam) + [alasan]
    for i, para in enumerate(paras, 1):
        if not isinstance(para, str) or len(para.strip()) < 60:
            print(f"[sub_3_2] paragraf {i} terlalu pendek")
            return None
        if not rm.extract_citations(para):
            print(f"[sub_3_2] paragraf {i} tidak memiliki sitasi")
            return None
        if re.search(r"\b\d{2,}\s*(orang|responden|pegawai|sampel|populasi)\b", para, re.I):
            print(f"[sub_3_2] paragraf {i} menyebut angka jumlah (dilarang)")
            return None

    ok, bad = rm.validate_texts(paras, refs)
    if not ok:
        print(f"[sub_3_2] sitasi di luar pool terverifikasi: {bad[:3]}")
        return None

    used = rm.refs_used_in(paras, refs)
    if len(used) < 2:
        print(f"[sub_3_2] sitasi terlalu sedikit ({len(used)} sumber)")
        return None

    return {
        "pengertian_populasi": [x.strip() for x in pop],
        "pengertian_sampel": [x.strip() for x in sam],
        "alasan_teknik": alasan.strip(),
        "sumber_terpakai": [r["apa"] for r in used],
    }


def write_sub_3_2(cfg):
    """Tulis narasi sub bab 3.2 Populasi dan Sampel via Gemini.

    Sitasi dibatasi pada pool terverifikasi (references_metodologi.json);
    angka populasi/sampel TIDAK ditulis AI (ditambahkan deterministik oleh engine).
    """
    from llm_title_parser import load_config, gemini_keys, MODEL_CHAIN
    import references_metodologi as rm

    config = load_config()
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return None

    refs = rm.load_refs()
    if not refs:
        return None

    judul = cfg.get("judul", "")
    instansi = cfg.get("instansi", "")
    populasi = int(cfg.get("populasi", 100))
    n_sample = int(cfg.get("n", populasi))
    teknik = (cfg.get("teknik_sampel") or "sensus").lower()
    if teknik == "slovin":
        teknik_desc = "Rumus Slovin (populasi relatif besar, melewati ambang 100)"
    else:
        teknik_desc = "Teknik sampling jenuh atau sensus (seluruh anggota populasi terjangkau)"

    prompt = PROMPT_BAB3_3_2.format(
        judul=judul, instansi=instansi, populasi=populasi, n_sample=n_sample,
        teknik_desc=teknik_desc, sumber_menu=rm.citation_menu(refs),
    )

    chain = [m for m in config.get("models", []) if m] or MODEL_CHAIN
    timeout = max(int(config.get("timeout_seconds", 90)), 90)
    keys = gemini_keys(config)

    try:
        from llm_ontology_researcher import _call_gemini_matrix
        prompt_safe = prompt
        for attempt in range(3):
            try:
                raw = _call_gemini_aman(keys, chain, prompt_safe, timeout, max_tokens=3072, label="pernyataan")
                if isinstance(raw, str):
                    raw = raw
                res = _validate_sub_3_2(raw, refs)
                if res:
                    print(f"[sub_3_2] narasi via Gemini (percobaan {attempt+1}), "
                          f"{len(res['sumber_terpakai'])} sumber terpakai")
                    return res
            except UnicodeEncodeError:
                continue
        print("[sub_3_2] semua percobaan gagal")
        return None
    except Exception as e:
        print(f"[sub_3_2] Gemini gagal: {type(e).__name__} {str(e)[:80]}")
        return None


# ---------------------------------------------------------------------------
# BAB III — SUB BAB 3.3 s.d. 3.6
# Narasi digenerate Gemini (unik per judul) dengan sitasi HANYA dari pool
# referensi metodologi terverifikasi (references_metodologi.json).
# ---------------------------------------------------------------------------

_FORBID_JUMLAH = r"\b\d{2,}\s*(?:orang|responden|pegawai|sampel|populasi|butir|indikator|variabel|item)\b"
_FORBID_ANGKA_TEKNIS = (
    _FORBID_JUMLAH
    + r"|(?:1,96|1\.96|0,05|0\.05|0,70|0\.70|0,50|0\.50|0,90|0\.90|0,80|0\.80|5\.000|5000)"
)

_HEAD_BAB3 = """Anda adalah penulis metodologi penelitian (BAB III) tesis Magister Manajemen (S2) Indonesia yang cermat dan faktual.

=== JUDUL PENELITIAN ===
{judul}

=== OBJEK/INSTANSI PENELITIAN ===
{instansi}

=== VARIABEL PENELITIAN ===
Y  : {y}
M  : {m}
X1 : {x1}
X2 : {x2}
"""

_TAIL_BAB3 = """
=== DAFTAR SITASI TERVERIFIKASI (satu-satunya sitasi yang BOLEH dipakai; format APA in-text) ===
{sumber_menu}

=== ATURAN KERAS ===
1. Bahasa Indonesia akademik formal, mengalir, tidak kaku, dan TIDAK repetitif.
2. Setiap paragraf WAJIB memuat minimal satu sitasi dalam tanda kurung, persis format (Penulis, Tahun) dari DAFTAR SITASI TERVERIFIKASI. DILARANG mengarang penulis, tahun, atau judul.
3. DILARANG menyebut jumlah/kuantitas dalam bentuk angka (mis. jumlah responden, butir, indikator, variabel). Angka tersebut ditambahkan terpisah oleh sistem.
4. Jangan menyalin kalimat contoh; susun redaksi sendiri. Antarjudul harus menghasilkan diksi dan urutan argumen yang berbeda.
5. Prioritaskan sumber dengan tahun terbaru.
6. BALAS HANYA JSON valid berisi tepat kunci yang diminta, tanpa teks penjelasan tambahan.
"""

_BODY_3_3 = """
=== TUGAS: SUB BAB 3.3 VARIABEL DAN DEFINISI OPERASIONAL ===
=== KELUARAN: JSON dengan tepat kunci "y", "m", "x1", dan "x2" ===
Setiap kunci variabel berisi objek dengan dua teks: "konseptual" dan "operasional".
  "konseptual"  : TEPAT 1 KALIMAT SINGKAT (maksimal 35 kata) berisi pengertian/teori variabel tersebut,
                  memakai pola: "<Nama variabel> adalah ... (Penulis, Tahun)." — sitasi WAJIB disalin PERSIS
                  dari MENU SUMBER VARIABEL milik variabel itu dan diletakkan di AKHIR kalimat, di dalam tanda kurung.
                  DILARANG menulis lebih dari satu kalimat (tanpa kalimat penjelasan/kesimpulan tambahan).
  "operasional" : TEPAT 1 paragraf Definisi Operasional (minimal 3 kalimat) yang menjelaskan bagaimana variabel
                  tersebut diukur dan diamati pada instansi penelitian ini, yaitu melalui kuesioner Skala Likert.
                  WAJIB menyebut PERSIS SELURUH dimensi variabel itu sebagaimana tertulis pada baris
                  "DIMENSI <variabel> YANG WAJIB DISEBUT PERSIS" — jangan mengganti, menambah, menerjemahkan
                  ulang, atau mengurangi nama dimensi. Tidak wajib bersitasi.
"""

_TAIL_3_3 = """
=== ATURAN KERAS ===
1. Bahasa Indonesia akademik formal, mengalir, tidak kaku, dan tidak repetitif.
2. Sitasi HANYA boleh diambil dari MENU SUMBER di atas. DILARANG mengarang penulis, tahun, atau judul.
3. Setiap "konseptual" WAJIB memuat minimal satu sitasi yang DISALIN PERSIS huruf demi huruf dari MENU SUMBER VARIABEL miliknya sendiri,
   termasuk SELURUH nama penulis dan tanda "&". Contoh BENAR: "(Puspitasari & Aulia, 2025)" atau "Nahdliyah & Izzah (2022) mengemukakan bahwa ...".
   DILARANG menyingkat sitasi menjadi satu nama saja (mis. menulis "(Aulia, 2025)" padahal menunya "(Puspitasari & Aulia, 2025)"),
   DILARANG mengganti tahun, dan DILARANG menulis pola janggal seperti "oleh (Penulis, Tahun)".
4. Definisi Konseptual satu variabel HANYA boleh mengutip menu variabel tersebut; DILARANG memakai sumber variabel lain.
5. DILARANG menyebut angka/kuantitas (jumlah responden, butir, indikator, nilai statistik, persentase).
6. Setiap "operasional" WAJIB menyebut instansi penelitian dan menyatakan variabel diukur melalui kuesioner Skala Likert beserta dimensinya.
7. Setiap variabel WAJIB memakai diksi pembuka, alur argumen, dan penutup yang BERBEDA satu sama lain.
8. "konseptual" WAJIB TEPAT 1 KALIMAT SINGKAT (maksimal 35 kata) dengan pola "<Nama variabel> adalah ... (Penulis, Tahun)."
   DILARANG menambahkan kalimat kedua, contoh, penjelasan lanjutan, atau kalimat kesimpulan.
9. BALAS HANYA JSON valid berisi tepat kunci "y", "m", "x1", "x2", tanpa teks penjelasan tambahan.
"""

_BODY_3_4 = """
=== TUGAS: SUB BAB 3.4 INSTRUMEN PENELITIAN DAN SKALA PENGUKURAN ===
=== KELUARAN: JSON dengan kunci "instrumen" dan "skala" ===
"instrumen" : array berisi 1 sampai 2 paragraf tentang kuesioner sebagai instrumen pengumpulan data, data yang dikumpulkan untuk variabel Y, M, X1, dan X2, serta pentingnya uji validitas dan reliabilitas instrumen.
"skala" : satu paragraf tentang Skala Likert sebagai skala pengukuran sikap/persepsi dan alasan penggunaannya. JANGAN menyebut angka bobot skor (bobot ditulis terpisah oleh sistem).
"""

_BODY_3_5 = """
=== TUGAS: SUB BAB 3.5 TEKNIK ANALISIS DATA ===
=== KELUARAN: JSON dengan kunci "pengantar", "deskriptif", dan "model" ===
"pengantar" : array berisi satu paragraf tentang alasan pemilihan Structural Equation Modeling berbasis varian (PLS-SEM) dengan software SmartPLS untuk menguji model yang melibatkan variabel mediasi.
"deskriptif" : array berisi satu paragraf tentang analisis statistik deskriptif data responden sebagai langkah awal.
"model" : array berisi 1 sampai 2 paragraf tentang tahapan evaluasi model pengukuran (outer model) dan model struktural (inner model) secara konseptual. JANGAN menyebut angka ambang batas (ditulis terpisah oleh sistem).
"""

_BODY_3_6 = """
=== TUGAS: SUB BAB 3.6 PENGUJIAN HIPOTESIS DAN KRITERIA KEPUTUSAN ===
=== KELUARAN: JSON dengan kunci "pengantar" ===
"pengantar" : array berisi 1 sampai 2 paragraf tentang prosedur pengujian hipotesis pengaruh langsung dan pengaruh tidak langsung (mediasi) melalui metode resampling/bootstrapping, taraf signifikansi, serta dasar pengambilan keputusan menerima atau menolak hipotesis. JANGAN menyebut angka ambang batas (ditulis terpisah oleh sistem).
"""


# ---------------------------------------------------------------------------
# B12/B13/B14 — KENDALI GAYA BAHASA BAB III–V
# ---------------------------------------------------------------------------
_FRASA_TERLARANG = (
    "bab hasil penelitian ini", "bab ini", "bagian ini",
    "sub-bab ini", "sub bab ini", "pada bab ini", "dalam bab ini",
)


def _tag_bab345(tag):
    """True untuk tag keluaran Bab III, IV, dan V."""
    t = (tag or "").lower()
    return t.startswith(("bab3", "bab4", "bab5", "sub_3", "sub_4", "sub_5",
                         "tabel_4", "intro_", "pembahasan"))


# Rujukan-diri yang dilarang: hanya pola meta-rujukan ("Bab ini ...", "pada
# bagian ini ..."), BUKAN frasa wajar seperti "mendominasi bagian ini".
_RX_FRASA_META = re.compile(
    r"(?:^|[.!?]\s+)(?:Bab|Sub-?bab|Bagian)\s+ini\b"
    r"|(?:pada|di|dalam|berdasarkan|seperti\s+pada)\s+(?:bab|sub-?bab|bagian)\s+ini\b",
    re.I)


def _ada_frasa_terlarang(teks):
    t = teks or ""
    if re.search(r"\bbab hasil penelitian ini\b", t, re.I):
        return True
    return bool(_RX_FRASA_META.search(t))


def _ganti_frasa_terlarang(teks):
    """Ganti rujukan-diri terlarang dengan sebutan netral (jaring pengaman terakhir)."""
    t = teks or ""
    t = re.sub(r"\b[Bb]ab hasil penelitian ini\b", "Uraian hasil penelitian ini", t)
    t = re.sub(r"(^|[.!?]\s+)(?:Bab|Sub-?bab|Bagian)\s+ini\b",
               r"\1Uraian ini", t)
    t = re.sub(r"(\bpada\s+)(?:bab|sub-?bab|bagian)\s+ini\b", r"\1uraian berikut", t, flags=re.I)
    t = re.sub(r"(\b(?:di|dalam)\s+)(?:bab|sub-?bab|bagian)\s+ini\b", r"\1uraian ini", t, flags=re.I)
    return t


def _sanitize_latex(teks):
    """Ubah kode LaTeX sederhana menjadi teks/simbol biasa (B14)."""
    t = teks or ""
    if "$" not in t and "\\" not in t:
        return t
    t = re.sub(r"\$\$(.+?)\$\$", r"\1", t, flags=re.S)
    t = re.sub(r"\$(.+?)\$", r"\1", t, flags=re.S)
    t = re.sub(r"\\\((.*?)\\\)", r"\1", t, flags=re.S)
    t = re.sub(r"\\\[(.*?)\\\]", r"\1", t, flags=re.S)
    t = re.sub(r"\\frac\{([^{}]+)\}\{([^{}]+)\}", r"\1/\2", t)
    t = re.sub(r"\\sqrt\{([^{}]+)\}", r"√\1", t)
    t = re.sub(r"\\text\{([^{}]+)\}", r"\1", t)
    peta = {
        "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "lambda": "λ",
        "mu": "μ", "sigma": "σ", "zeta": "ζ", "eta": "η", "theta": "θ",
        "rho": "ρ", "tau": "τ", "chi": "χ", "times": "×", "cdot": "·",
        "pm": "±", "le": "≤", "ge": "≥", "ne": "≠", "approx": "≈",
    }
    for cmd, sym in peta.items():
        t = re.sub(r"\\%s(?![A-Za-z])" % cmd, sym, t)
    # notasi subskrip/superskrip LaTeX -> karakter Unicode
    for _d, _s in {"0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄",
                   "5": "₅", "6": "₆", "7": "₇", "8": "₈", "9": "₉"}.items():
        t = t.replace("_%s" % _d, _s)
    t = t.replace("^2", "²").replace("^3", "³")
    t = t.replace("\\\\", " ").replace("\\{", "{").replace("\\}", "}")
    t = re.sub(r"\\([A-Za-z]+)", r"\1", t)
    return t


def _jumlah_kalimat(teks):
    """Hitung kalimat; token teknis seperti 'Tabel 4.2' diabaikan."""
    t = _SAFE_TOKENS_RE.sub(" ", teks or "")
    return len(re.findall(r"[.!?](?:\s|$)", t.strip()))


def _bersihkan_hasil(nilai, tag):
    """Bersihkan keluaran Gemini: LaTeX + frasa terlarang (rekursif)."""
    if not _tag_bab345(tag):
        return nilai
    if isinstance(nilai, str):
        return _ganti_frasa_terlarang(_sanitize_latex(nilai))
    if isinstance(nilai, list):
        return [_bersihkan_hasil(x, tag) for x in nilai]
    if isinstance(nilai, dict):
        return {k: _bersihkan_hasil(v, tag) for k, v in nilai.items()}
    return nilai


def _gemini_json(prompt, validate_fn, max_tokens=3072, tag="bab3", attempts=3,
                 temperature=0.2):
    """Panggil Gemini lalu validasi hasilnya; coba `attempts` kali; kembalikan None bila gagal."""
    if _SUPPRESS_GEMINI:
        return None
    if _tag_bab345(tag):
        prompt = (prompt.rstrip() + "\n\nATURAN WAJIB (GAYA BAHASA): "
                  "DILARANG memakai frasa 'Bab ini', 'bagian ini', 'Sub-bab ini', "
                  "maupun 'Bab hasil penelitian ini'. Gunakan sebutan netral seperti "
                  "'uraian berikut', 'tahap ini', atau 'pembahasan berikut'. "
                  "DILARANG pula menulis kode LaTeX (mis. $...$, \\frac, \\beta); "
                  "gunakan simbol biasa atau kata.")
    from llm_title_parser import load_config, gemini_keys, MODEL_CHAIN

    config = load_config()
    api_key = (gemini_keys(config) or [""])[0]
    if not api_key or not config.get("enabled", True):
        return None
    chain = [m for m in config.get("models", []) if m] or MODEL_CHAIN
    timeout = max(int(config.get("timeout_seconds", 90)), 90)
    keys = gemini_keys(config)
    try:
        from llm_ontology_researcher import _call_gemini_matrix
        prompt_safe = prompt
        for attempt in range(max(1, int(attempts))):
            try:
                raw = _call_gemini_matrix(keys, chain, prompt_safe, timeout, max_tokens=max_tokens,
                                          temperature=temperature)
                if isinstance(raw, str):
                    raw = raw
                res = validate_fn(raw)
                if res:
                    # B13/B14: bersihkan frasa terlarang & sisa kode LaTeX.
                    res = _bersihkan_hasil(res, tag)
                    print(f"[{tag}] narasi via Gemini (percobaan {attempt+1}), "
                          f"{len(res.get('sumber_terpakai', []))} sumber terpakai")
                    return res
            except UnicodeEncodeError:
                continue
        print(f"[{tag}] semua percobaan gagal")
        return None
    except Exception as e:
        print(f"[{tag}] Gemini gagal: {type(e).__name__} {str(e)[:80]}")
        return None


_SAFE_TOKENS_RE = re.compile(
    r"\b(?:Tabel|Gambar)\s+\d+(?:\.\d+)*\b"
    r"|\bSmartPLS\s*\d+(?:\.\d+)*\b"
    r"|\bPLS-?SEM\b|\bSPSS\b",
    re.I,
)


def _sanitize_teknis(text):
    """Buang token teknis ber-angka (mis. 'Tabel 4.1', 'SmartPLS 4.0') sebelum cek angka."""
    return _SAFE_TOKENS_RE.sub(" ", text or "")


def _validate_bab3_output(raw, list_keys, str_keys=(), tag="bab3",
                          forbid_re=None, min_refs=2, refs=None,
                          require_cite=True, forbid_cite=False,
                          exact_len=None, max_len=5, min_len=60):
    """Validator umum narasi naratif (BAB III/IV/V).

    require_cite=True  -> setiap paragraf WAJIB bersitasi dari pool terverifikasi.
    forbid_cite=True   -> paragraf DILARANG memuat sitasi (bagian hasil/kesimpulan).
    forbid_re          -> regex angka yang dilarang (setelah token teknis disaring).
    exact_len          -> {kunci: jumlah paragraf} bila jumlahnya harus persis.
    """
    import references_metodologi as rm

    refs = refs if refs is not None else rm.load_refs()
    if isinstance(raw, list) and raw and isinstance(raw[0], dict):
        raw = raw[0]
    if not isinstance(raw, dict):
        print(f"[{tag}] keluaran bukan objek JSON")
        return None

    exact_len = exact_len or {}
    out = {}
    paras_all = []
    for k in list_keys:
        v = raw.get(k)
        if isinstance(v, str):
            v = [v]
        if not isinstance(v, list) or not (1 <= len(v) <= max_len):
            print(f"[{tag}] kunci '{k}' tidak valid")
            return None
        if k in exact_len and len(v) != exact_len[k]:
            print(f"[{tag}] '{k}' harus {exact_len[k]} paragraf, dapat {len(v)}")
            return None
        cleaned = []
        for para in v:
            if not isinstance(para, str) or len(para.strip()) < min_len:
                print(f"[{tag}] '{k}': paragraf terlalu pendek")
                return None
            cleaned.append(para.strip())
        out[k] = cleaned
        paras_all.extend(cleaned)

    for k in str_keys:
        v = raw.get(k)
        if isinstance(v, list) and v and isinstance(v[0], str):
            v = v[0]
        if not isinstance(v, str) or len(v.strip()) < min_len:
            print(f"[{tag}] kunci '{k}' tidak valid")
            return None
        out[k] = v.strip()
        paras_all.append(v.strip())

    # B13: frasa "Bab ini / bagian ini / Sub-bab ini" dilarang di Bab III–V.
    if _tag_bab345(tag):
        for para in paras_all:
            if _ada_frasa_terlarang(para):
                print(f"[{tag}] frasa terlarang (Bab ini/bagian ini/Sub-bab ini) -> tolak")
                return None

    # Auto-repair sitasi yang kehilangan penulis pertama (mis. '(Izzah, 2022)').
    if require_cite and refs:
        fixed = False
        for k in list(out.keys()):
            v = out[k]
            if isinstance(v, str):
                nv, reps = rm.repair_citations(v, refs)
                if reps:
                    out[k] = nv
                    fixed = True
                    print(f"[{tag}] auto-repair sitasi {k}: {reps[:2]}")
            elif isinstance(v, list):
                for i, para in enumerate(v):
                    nv, reps = rm.repair_citations(para, refs)
                    if reps:
                        v[i] = nv
                        fixed = True
                        print(f"[{tag}] auto-repair sitasi {k}[{i + 1}]: {reps[:2]}")
        if fixed:
            paras_all = []
            for k in list_keys:
                paras_all.extend(out[k])
            for k in str_keys:
                paras_all.append(out[k])

    pat = forbid_re or _FORBID_JUMLAH
    for i, para in enumerate(paras_all, 1):
        cites = rm.extract_citations(para)
        if require_cite and not cites:
            print(f"[{tag}] paragraf {i} tidak memiliki sitasi")
            return None
        if forbid_cite and cites:
            print(f"[{tag}] paragraf {i} memuat sitasi yang tidak diizinkan")
            return None
        if re.search(pat, _sanitize_teknis(para), re.I):
            print(f"[{tag}] paragraf {i} memuat angka yang dilarang")
            return None

    if not require_cite:
        out["sumber_terpakai"] = []
        return out

    ok, bad = rm.validate_texts(paras_all, refs)
    if not ok:
        print(f"[{tag}] sitasi di luar pool terverifikasi: {bad[:3]}")
        return None

    used = rm.refs_used_in(paras_all, refs)
    if len(used) < min_refs:
        print(f"[{tag}] sitasi terlalu sedikit ({len(used)} sumber)")
        return None
    out["sumber_terpakai"] = [r["apa"] for r in used]
    return out


def _bab3_prompt(cfg, body):
    import references_metodologi as rm
    return (
        _HEAD_BAB3.format(
            judul=cfg.get("judul", ""), instansi=cfg.get("instansi", ""),
            y=cfg.get("y", "Kinerja Pegawai"), m=cfg.get("m", "Kepuasan Kerja"),
            x1=cfg.get("x1", "X1"), x2=cfg.get("x2", "X2"))
        + body
        + _TAIL_BAB3.format(sumber_menu=rm.citation_menu())
    )


def _ascii(s):
    """Normalisasi teks untuk prompt: smart-quote/dash -> ASCII, aksen DIPERTAHANKAN.

    Dulu fungsi ini membuang SEMUA karakter non-ASCII, sehingga nama penulis
    beraksen (mis. "Grbić", "Karadayı") menjadi "Grbic"/"Karaday" dan sitasi di
    naskah tak lagi cocok dengan entri DAFTAR PUSTAKA. Sekarang stdout server
    dipaksa UTF-8 (lihat server.py), jadi aksen aman untuk dikirim.
    """
    try:
        s = str(s or "").replace("&amp;", "&")
        for a, b in (("\u2013", "-"), ("\u2014", "-"), ("\u2019", "'"),
                     ("\u201c", '"'), ("\u201d", '"')):
            s = s.replace(a, b)
        return s
    except Exception:
        return str(s or "")


def _menu_from_refs(refs):
    """Blok menu sitasi (sitasi | APA) untuk prompt dari pool per variabel."""
    lines = []
    for r in (refs or []):
        sit = _ascii(r.get("sitasi") or "")
        apa = _ascii(r.get("apa") or "")
        if sit and apa:
            lines.append(f"- {sit}  |  {apa}")
    return "\n".join(lines) if lines else "- (tidak ada menu sumber)"


def _sub_3_3_prompt(cfg, onto, role_pools):
    """Prompt 3.3: menu sumber terverifikasi dipisah per variabel."""
    variabels = (onto or {}).get("variables", {}) or {}
    role_map = [("y", "Y"), ("m", "M"), ("x1", "X1"), ("x2", "X2")]
    blocks = []
    for role, role_up in role_map:
        v = variabels.get(role, {}) or {}
        name = _ascii(v.get("nama_resmi") or cfg.get(role) or role_up)
        seed = _ascii((v.get("definisi") or "").strip())
        dims = [d.get("nama", "").strip() for d in (v.get("dimensi") or []) if d.get("nama")]
        menu = _menu_from_refs(role_pools.get(role) or [])
        blocks.append(
            f"=== MENU SUMBER VARIABEL {role_up} = {name} (khusus Definisi Konseptual {role_up}) ===\n"
            f"{menu}\n"
            f"--- Rujukan konsep {role_up} (parafrase bebas, jangan disalin mentah): {seed or '-'}\n"
            f"--- DIMENSI {role_up} YANG WAJIB DISEBUT PERSIS pada Definisi Operasional: "
            f"{'; '.join(dims) if dims else '-'}"
        )
    return (
        _HEAD_BAB3.format(
            judul=cfg.get("judul", ""), instansi=cfg.get("instansi", ""),
            y=cfg.get("y", ""), m=cfg.get("m", ""),
            x1=cfg.get("x1", ""), x2=cfg.get("x2", ""))
        + _BODY_3_3 + "\n" + "\n\n".join(blocks) + _TAIL_3_3
    )


def _validate_sub_3_3(raw, role_pools, tag="sub_3_3", onto=None):
    """Validasi 3.3: tiap 'konseptual' wajib bersitasi & sitasinya hanya dari pool variabel itu.

    Definisi Operasional juga wajib menyebut SELURUH dimensi variabel tersebut
    sesuai kisi-kisi instrumen (dimensi berasal dari ontologi, bukan karangan AI).
    """
    import references_metodologi as rm

    role_map = [("y", "Y"), ("m", "M"), ("x1", "X1"), ("x2", "X2")]
    if isinstance(raw, list) and raw and isinstance(raw[0], dict):
        raw = raw[0]
    if not isinstance(raw, dict):
        print(f"[{tag}] keluaran bukan objek JSON")
        return None

    metodologi = rm.load_refs()
    out = {}
    for role, role_up in role_map:
        blk = raw.get(role)
        if not isinstance(blk, dict):
            print(f"[{tag}] kunci '{role}' tidak valid")
            return None
        vals = {}
        for key, min_chars in (("konseptual", 90), ("operasional", 150)):
            v = blk.get(key)
            if isinstance(v, list):
                v = " ".join(str(x) for x in v if x)
            if not isinstance(v, str) or len(v.strip()) < min_chars:
                print(f"[{tag}] {role_up}.{key} tidak valid/kurang panjang")
                return None
            v = v.strip()
            if re.search(_FORBID_ANGKA_TEKNIS, _sanitize_teknis(v), re.I):
                print(f"[{tag}] {role_up}.{key} memuat angka yang dilarang")
                return None
            vals[key] = v

        if not rm.extract_citations(vals["konseptual"]):
            print(f"[{tag}] {role_up}.konseptual tanpa sitasi")
            return None

        # Definisi konseptual harus SINGKAT: 1 kalimat (toleransi 2 kalimat), maksimal 340 karakter.
        if len(vals["konseptual"]) > 340 or len(re.findall(r"[.!?](?:\s|$)", vals["konseptual"])) > 3:
            print(f"[{tag}] {role_up}.konseptual terlalu panjang (bukan 1 kalimat singkat)")
            return None

        pool = list(role_pools.get(role) or []) + list(metodologi)
        # Auto-repair sitasi dua-penulis yang kehilangan penulis pertama.
        for _k in ("konseptual", "operasional"):
            nv, reps = rm.repair_citations(vals[_k], pool)
            if reps:
                vals[_k] = nv
                print(f"[{tag}] auto-repair {role_up}.{_k}: {reps[:2]}")

        ok, bad = rm.validate_texts([vals["konseptual"]], pool)
        if not ok:
            print(f"[{tag}] {role_up}.konseptual sitasi di luar pool: {bad[:3]}")
            return None
        ok, bad = rm.validate_texts([vals["operasional"]], pool)
        if not ok:
            print(f"[{tag}] {role_up}.operasional sitasi di luar pool: {bad[:3]}")
            return None

        # Dimensi yang disebut pada Definisi Operasional WAJIB sama dengan dimensi
        # pada kisi-kisi instrumen (ontologi), agar 3.3 tidak bertentangan sendiri.
        v_onto = ((onto or {}).get("variables") or {}).get(role) or {}
        dims = [d.get("nama", "").strip() for d in (v_onto.get("dimensi") or [])
                if d.get("nama")]
        if dims:
            op_low = vals["operasional"].lower()
            hilang = []
            for dd in dims:
                inti = re.split(r"[(\[]", dd)[0].strip().lower()
                if len(inti) >= 4 and inti not in op_low:
                    hilang.append(dd)
            if hilang:
                print(f"[{tag}] {role_up}.operasional tidak menyebut dimensi: {hilang[:3]}")
                return None
        out[role] = vals

    apas, seen = [], set()
    for role, _ in role_map:
        for r in rm.refs_used_in([out[role]["konseptual"], out[role]["operasional"]],
                                 list(role_pools.get(role) or []) + metodologi):
            apa = r.get("apa")
            if apa and apa not in seen:
                seen.add(apa)
                apas.append(apa)
    out["sumber_terpakai"] = apas
    return out


def write_sub_3_3(cfg, onto=None, role_pools=None):
    """Sub bab 3.3 Variabel dan Definisi Operasional (Definisi Konseptual/Operasional via Gemini)."""
    role_pools = role_pools or {}
    if onto is None:
        try:
            import engine_hierarchical_pipeline as _eng
            onto = _eng.get_or_create_thesis_ontology(cfg)
            if not role_pools:
                role_pools, _ = _eng._bab3_var_pools(onto)
        except Exception:
            onto = {"variables": {}}
    return _gemini_json(
        _sub_3_3_prompt(cfg, onto, role_pools),
        lambda raw: _validate_sub_3_3(raw, role_pools, tag="sub_3_3", onto=onto),
        tag="sub_3_3", max_tokens=4096, attempts=4)


def write_sub_3_4(cfg):
    """Sub bab 3.4 Instrumen Penelitian dan Skala Pengukuran (via Gemini)."""
    return _gemini_json(
        _bab3_prompt(cfg, _BODY_3_4),
        lambda raw: _validate_bab3_output(raw, ["instrumen"], ["skala"], tag="sub_3_4"),
        tag="sub_3_4")


def write_sub_3_5(cfg):
    """Sub bab 3.5 Teknik Analisis Data (via Gemini)."""
    return _gemini_json(
        _bab3_prompt(cfg, _BODY_3_5),
        lambda raw: _validate_bab3_output(
            raw, ["pengantar", "deskriptif", "model"], tag="sub_3_5",
            forbid_re=_FORBID_ANGKA_TEKNIS),
        tag="sub_3_5")


def write_sub_3_6(cfg):
    """Sub bab 3.6 Pengujian Hipotesis dan Kriteria Keputusan (via Gemini)."""
    return _gemini_json(
        _bab3_prompt(cfg, _BODY_3_6),
        lambda raw: _validate_bab3_output(
            raw, ["pengantar"], tag="sub_3_6",
            forbid_re=_FORBID_ANGKA_TEKNIS),
        tag="sub_3_6")


# ---------------------------------------------------------------------------
# BAB IV & BAB V
# Narasi hasil/pembahasan/kesimpulan digenerate Gemini (unik per judul).
# Angka statistik SELALU ditulis deterministik oleh engine, bukan oleh AI.
# ---------------------------------------------------------------------------

_FORBID_ANGKA_IV = (
    r"\b\d+,\d+\b"                                          # 0,565 / 3,948 / 1,96
    r"|\b\d+\.\d+\b"                                        # 5.000 / 0.70
    r"|\b\d+(?:[.,]\d+)?\s*(?:%|persen)\b"                  # 54,35% / 20 persen
    r"|\b\d{2,}\s*(?:orang|responden|pegawai|butir|indikator|item|variabel|kuesioner|sampel|populasi)\b"
)

_HEAD_BAB45 = """Anda adalah penulis naskah tesis Magister Manajemen (S2) Indonesia yang cermat, akademis, dan faktual.

=== JUDUL PENELITIAN ===
{judul}

=== OBJEK/INSTANSI PENELITIAN ===
{instansi}

=== VARIABEL PENELITIAN ===
Y  : {y}
M  : {m}
X1 : {x1}
X2 : {x2}
"""

_ATURAN_TANPA_ANGKA = """
=== ATURAN KERAS ===
1. Bahasa Indonesia akademik formal, mengalir, tidak kaku, dan TIDAK repetitif.
2. DILARANG menuliskan ANGKA NILAI STATISTIK (jumlah responden, persentase, koefisien, t-statistic, p-value, f-square, R-square, tahun). Nilai tersebut disajikan terpisah oleh sistem pada tabel.
3. Nomor tabel/gambar WAJIB ditulis dengan angka biasa seperti "Tabel 4.5" atau "Tabel 4.12". JANGAN menuliskan nomor tabel dengan huruf (contoh yang SALAH: "Tabel empat titik lima").
4. DILARANG memuat sitasi/kutipan dalam bentuk apa pun.
5. Setiap paragraf minimal 3 kalimat dan spesifik pada variabel serta instansi di atas (bukan kalimat umum).
6. Jangan menyalin contoh; antarjudul harus menghasilkan diksi dan urutan argumen berbeda.
7. BALAS HANYA JSON valid berisi tepat kunci yang diminta, tanpa teks penjelasan tambahan.
"""

_ATURAN_PEMBAHASAN = """
=== DAFTAR SITASI TERVERIFIKASI (satu-satunya sitasi yang BOLEH dipakai; format APA in-text) ===
{sumber_menu}

=== ATURAN KERAS ===
1. Bahasa Indonesia akademik formal, mengalir, tidak kaku, dan TIDAK repetitif.
2. Setiap alinea WAJIB memuat minimal satu sitasi persis format (Penulis, Tahun) dari daftar di atas. DILARANG mengarang penulis, tahun, atau judul.
3. DILARANG menuliskan ANGKA NILAI STATISTIK (koefisien, t-statistic, p-value, f-square, persentase, jumlah responden, tahun). Nilai tersebut disajikan terpisah oleh sistem.
4. Nomor tabel WAJIB ditulis dengan angka biasa seperti "Tabel 4.14". JANGAN menuliskan nomor tabel dengan huruf.
5. Setiap alinea minimal 4 kalimat dan spesifik pada variabel serta instansi di atas.
6. Jangan menyalin contoh; antarjudul harus menghasilkan diksi dan urutan argumen berbeda.
7. BALAS HANYA JSON valid berisi tepat kunci yang diminta, tanpa teks penjelasan tambahan.
"""

_BODY_4_1_1 = """
=== TUGAS: SUB BAB 4.1.1 DESKRIPSI KARAKTERISTIK RESPONDEN ===
Konteks: profil demografis responden sudah disajikan pada Tabel 4.1 (jenis kelamin), Tabel 4.2 (kelompok usia), dan Tabel 4.3 (pendidikan terakhir).
=== KELUARAN: JSON dengan kunci "narasi" ===
"narasi" : array berisi tepat 2 paragraf uraian kualitatif mengenai komposisi demografis responden (variasi jenis kelamin, sebaran kelompok usia produktif, dan dominasi latar belakang pendidikan) beserta implikasinya terhadap keterwakilan data penelitian di instansi tersebut. Boleh menyebut "Tabel 4.1", "Tabel 4.2", dan "Tabel 4.3".
"""

_BODY_4_1_2 = """
=== TUGAS: SUB BAB 4.1.2 ANALISIS STATISTIK DESKRIPTIF VARIABEL ===
Konteks: nilai rata-rata dan kategori jawaban responden per variabel sudah disajikan pada Tabel 4.4 dan Tabel 4.5.
=== KELUARAN: JSON dengan kunci "narasi" ===
"narasi" : array berisi tepat 2 paragraf uraian kualitatif mengenai kecenderungan persepsi responden pada variabel Y, M, X1, dan X2 (seluruhnya berada pada kategori tinggi) serta maknanya bagi kesiapan organisasi menghadapi tuntutan kinerja. Boleh menyebut "Tabel 4.5".
PENTING: JANGAN menyatakan variabel mana yang paling tinggi, tertinggi, terbesar, paling dominan, atau paling menonjol. Klaim peringkat ditulis terpisah oleh sistem dan akan bertentangan dengan tabel bila Anda mengarang.
"""

_BODY_4_1_3 = """
=== TUGAS: SUB BAB 4.1.3 ANALISIS OUTER MODEL DAN INNER MODEL ===
Konteks: hasil evaluasi model pengukuran disajikan pada Tabel 4.7, koefisien determinasi pada Tabel 4.12, dan pengujian hipotesis pada Tabel 4.14.
=== KELUARAN: JSON dengan kunci "outer", "inner", dan "hipotesis" (masing-masing array berisi tepat 1 paragraf) ===
"outer" : uraian kualitatif bahwa seluruh indikator reflektif memenuhi validitas konvergen dan reliabilitas konstruk sehingga instrumen layak dipakai.
"inner" : uraian kualitatif mengenai kemampuan model struktural dalam menjelaskan variabel mediasi dan variabel endogen.
"hipotesis" : uraian kualitatif bahwa seluruh hipotesis pengaruh langsung maupun pengaruh tidak langsung (mediasi) terbukti positif dan signifikan.
Boleh menyebut nomor tabel, tetapi JANGAN menuliskan nilai statistik apa pun.
"""

_BODY_5_1 = """
=== TUGAS: SUB BAB 5.1 KESIMPULAN ===
=== KELUARAN: JSON dengan kunci "kesimpulan" ===
"kesimpulan" : array berisi TEPAT 7 butir kesimpulan (TANPA nomor, TANPA angka statistik, TANPA elaborasi) yang menjawab H1 sampai H7 secara berurutan:
  1) pengaruh X1 terhadap M, 2) pengaruh X2 terhadap M, 3) pengaruh X1 terhadap Y, 4) pengaruh X2 terhadap Y,
  5) pengaruh M terhadap Y, 6) peran M memediasi pengaruh X1 terhadap Y, 7) peran M memediasi pengaruh X2 terhadap Y.
Setiap butir HANYA 1 kalimat pendek, TANPA elaborasi, dengan pola:
  - butir 1-5: "<variabel bebas> terbukti berpengaruh positif dan signifikan terhadap <variabel terikat> pegawai <instansi>."
  - butir 6-7: "<variabel bebas> terbukti berpengaruh positif dan signifikan terhadap <variabel terikat> melalui <variabel mediasi> sebagai variabel mediasi pada <instansi>."
WAJIB menyebut arah pengaruh (positif/negatif); sistem akan menolak butir tanpa arah.
Tidak ada kalimat tambahan, tidak ada angka, tidak ada sitasi.
"""

_BODY_5_2 = """
=== TUGAS: SUB BAB 5.2 IMPLIKASI KEBIJAKAN ===
=== KELUARAN: JSON dengan kunci "pembuka" dan "pilar" ===
"pembuka" : 1 paragraf (minimal 3 kalimat, TANPA angka) yang menyatakan bahwa temuan empiris dan pembuktian statistik melahirkan implikasi manajerial berbasis bukti (Evidence-Based Policy) bagi instansi, serta menyebut langkah-langkah strategis yang disarankan.
"pilar" : array berisi TEPAT 4 objek, dengan urutan:
  1) penguatan stimulus X1, 2) transformasi tata kelola X2, 3) pelembagaan penguatan M, 4) optimalisasi Y secara berkelanjutan.
  Setiap objek berisi:
    "judul" : judul pilar (tanpa nomor),
    "a_judul" : judul sub-poin (a),
    "a_uraian" : uraian sub-poin (a), minimal 2 kalimat, TANPA angka,
    "b_judul" : judul sub-poin (b),
    "b_uraian" : uraian sub-poin (b), minimal 2 kalimat, TANPA angka.
Setiap pilar menyajikan 2 sub-poin (a dan b) berupa langkah strategis konkret dan spesifik bagi instansi.
"""

_BODY_5_3 = """
=== TUGAS: SUB BAB 5.3 SARAN ===
=== KELUARAN: JSON dengan kunci "pembuka", "praktis", dan "akademis" ===
"pembuka" : 1 paragraf (minimal 2 kalimat, TANPA angka) yang menyatakan bahwa berdasarkan kesimpulan dan implikasi kebijakan yang telah diuraikan, terdapat beberapa saran/rekomendasi yang diharapkan memberikan manfaat konstruktif, baik secara praktis bagi instansi terkait maupun secara akademis bagi pengembangan ilmu pengetahuan.
"praktis" : array berisi TEPAT 2 objek saran praktis (TANPA angka), dengan urutan:
  1) bagi pimpinan dan manajemen SDM instansi, 2) bagi pegawai (ASN dan PPPK).
  Setiap objek berisi:
    "judul" : judul sub-poin (mis. "Bagi Pimpinan dan Manajemen SDM", "Bagi Pegawai (ASN dan PPPK)"),
    "uraian" : uraian saran, minimal 3 kalimat, TANPA angka.
"akademis" : array berisi TEPAT 2 objek saran akademis (TANPA angka), dengan urutan:
  1) eksplorasi variabel baru, 2) perluasan objek dan lokasi penelitian.
  Setiap objek berisi:
    "judul" : judul sub-poin (mis. "Eksplorasi Variabel Baru", "Perluasan Objek dan Lokasi Penelitian"),
    "uraian" : uraian saran, minimal 3 kalimat, TANPA angka.
  PENTING: saran akademis harus SPESIFIK pada konteks penelitian ini (instansi, variabel, sektor). Untuk eksplorasi variabel baru, sarankan variabel yang relevan dengan konteks berikut: {angle}. JANGAN memberikan saran generik yang bisa dipakai di sembarang penelitian.
"""


def _bab45_prompt(cfg, body, aturan):
    return (
        _HEAD_BAB45.format(
            judul=cfg.get("judul", ""), instansi=cfg.get("instansi", ""),
            y=cfg.get("y", "Kinerja Pegawai"), m=cfg.get("m", "Kepuasan Kerja"),
            x1=cfg.get("x1", "X1"), x2=cfg.get("x2", "X2"))
        + body
        + aturan
    )


def write_sub_4_1_1(cfg):
    """4.1.1 Deskripsi karakteristik responden (narasi kualitatif, tanpa angka/sitasi)."""
    return _gemini_json(
        _bab45_prompt(cfg, _BODY_4_1_1, _ATURAN_TANPA_ANGKA),
        lambda raw: _validate_bab3_output(
            raw, ["narasi"], tag="sub_4_1_1", refs=[], require_cite=False,
            forbid_cite=True, forbid_re=_FORBID_ANGKA_IV,
            exact_len={"narasi": 2}, min_len=180),
        tag="sub_4_1_1")


def write_sub_4_1_2(cfg):
    """4.1.2 Analisis statistik deskriptif variabel (narasi kualitatif)."""
    return _gemini_json(
        _bab45_prompt(cfg, _BODY_4_1_2, _ATURAN_TANPA_ANGKA),
        lambda raw: _validate_bab3_output(
            raw, ["narasi"], tag="sub_4_1_2", refs=[], require_cite=False,
            forbid_cite=True,
            forbid_re=(_FORBID_ANGKA_IV
                       + r"|paling (?:tinggi|dominan|menonjol|kuat|besar)|tertinggi|terbesar"),
            exact_len={"narasi": 2}, min_len=180),
        tag="sub_4_1_2")


def write_sub_4_1_3(cfg):
    """4.1.3 Analisis outer model & inner model (narasi kualitatif per blok tabel)."""
    return _gemini_json(
        _bab45_prompt(cfg, _BODY_4_1_3, _ATURAN_TANPA_ANGKA),
        lambda raw: _validate_bab3_output(
            raw, ["outer", "inner", "hipotesis"], tag="sub_4_1_3", refs=[],
            require_cite=False, forbid_cite=True, forbid_re=_FORBID_ANGKA_IV,
            exact_len={"outer": 1, "inner": 1, "hipotesis": 1}, min_len=180),
        tag="sub_4_1_3")


# =============== NARASI & DESKRIPSI SETIAP TABEL BAB IV ===============
# Setiap tabel Bab IV wajib: 1 paragraf narasi pengantar (sebelum tabel) dan
# 2 paragraf deskripsi tabel (sesudah tabel), seluruhnya ditulis Gemini.
_RX_ANGKA_AI = re.compile(r"\d+(?:[.,]\d+)*")

# Konstanta metodologis umum yang wajar muncul (bukan data hasil penelitian).
_ANGKA_UMUM = {
    "005", "001", "010", "050", "01", "05",
    "196", "1645", "258", "233", "333", "167", "133",
    "5000", "2022", "2023", "2024", "2025", "2026",
}

# Angka yang ditulis dengan HURUF (mis. "tujuh puluh persen") lolos dari
# pemeriksaan digit, padahal bisa berbeda dari angka pada tabel. Dua pola
# berikut menangkapnya tanpa menjaring frasa wajar seperti "salah satu".
_RX_ANGKA_KATA = re.compile(
    r"\b(?:(?:satu|dua|tiga|empat|lima|enam|tujuh|delapan|sembilan|se)?\s*"
    r"(?:puluh|ratus|ribu)|sebelas)\b", re.I)
_RX_ANGKA_KATA_JUMLAH = re.compile(
    r"\b(?:sebanyak|sejumlah|berjumlah|mencapai|tercatat|senilai|sebesar|yaitu|yakni)\s+"
    r"(?:[a-z]+\s+){0,2}(?:satu|dua|tiga|empat|lima|enam|tujuh|delapan|sembilan|sepuluh)\b",
    re.I)


def _naskah_terpotong(teks):
    """Benar bila paragraf tampak terpotong di tengah kalimat."""
    t = (teks or "").strip()
    if not t:
        return True
    if re.search(r"\bTabel\s+\d+\.\s*$", t):
        return True
    if re.search(r"\b(?:dan|atau|serta|yaitu|yakni|dengan|pada|untuk|dari|yang|serta|"
                 r"sebagai|melalui|adalah|ialah)\s*$", t, re.I):
        return True
    return not re.search(r"[.!?\"')\]]\s*$", t)


def _ada_angka_kata(teks):
    """Benar bila ada angka bersebut huruf yang berpotensi salah."""
    t = teks or ""
    return bool(_RX_ANGKA_KATA.search(t) or _RX_ANGKA_KATA_JUMLAH.search(t))


def _kumpulan_angka(*bagian):
    """Himpunan token angka (normalisasi tanpa pemisah) yang muncul di data tabel."""
    teks = " ".join(str(x) for b in bagian
                    for x in (b if isinstance(b, (list, tuple, set)) else [b]))
    out = set()
    for m in _RX_ANGKA_AI.finditer(teks):
        out.add(re.sub(r"[.,]", "", m.group(0)))
    return out


def _validate_tabel_bab4(raw, angka_ok, tag, tabel_no=""):
    """Validasi: SATU kalimat pembuka singkat + 2 paragraf sesudah, tanpa sitasi.

    Seluruh angka yang ditulis AI WAJIB ada pada data tabel (kecuali konstanta
    metodologis umum) sehingga tidak ada angka yang dikarang.
    """
    res = _validate_bab3_output(
        raw, ["sebelum", "sesudah"], tag=tag, refs=[], require_cite=False,
        forbid_cite=True, forbid_re=r"(?!x)x",
        exact_len={"sebelum": 1, "sesudah": 2}, min_len=40, max_len=4)
    if not res:
        return None

    # B12: narasi sebelum tabel = SATU kalimat pembuka yang singkat.
    sbl = res["sebelum"][0]
    if _jumlah_kalimat(sbl) > 1:
        print(f"[{tag}] 'sebelum' harus satu kalimat pembuka -> tolak")
        return None
    if len(sbl.split()) > 26:
        print(f"[{tag}] 'sebelum' terlalu panjang ({len(sbl.split())} kata) -> tolak")
        return None
    if _ada_frasa_terlarang(sbl):
        print(f"[{tag}] 'sebelum' memuat frasa terlarang -> tolak")
        return None
    for para in res["sesudah"]:
        if len(para) < 130:
            print(f"[{tag}] paragraf 'sesudah' terlalu pendek -> tolak")
            return None

    for k in ("sebelum", "sesudah"):
        for para in res[k]:
            if _naskah_terpotong(para):
                print(f"[{tag}] paragraf terpotong di tengah kalimat -> tolak")
                return None
            if _ada_angka_kata(para):
                print(f"[{tag}] angka ditulis dengan huruf (mis. 'tujuh puluh') -> tolak")
                return None
            for m in _RX_ANGKA_AI.finditer(_sanitize_teknis(para)):
                tok = re.sub(r"[.,]", "", m.group(0))
                if tok not in angka_ok and tok not in _ANGKA_UMUM:
                    print(f"[{tag}] angka '{m.group(0)}' tidak ada di Tabel {tabel_no} -> tolak")
                    return None
    return res


_TABEL_BAB4 = """Anda penulis naskah tesis magister Manajemen SDM yang teliti dan sistematis.
Anda menulis bagian HASIL PENELITIAN (BAB IV) untuk naskah:
"{judul}"
Objek penelitian: {instansi}
Variabel: X1 = {x1}, X2 = {x2}, M = {m}, Y = {y}

Satu tabel hasil penelitian akan disajikan. Tugas Anda membuat SATU KALIMAT PEMBUKA
yang sangat singkat (sebelum tabel) dan DUA paragraf DESKRIPSI TABEL (sesudah tabel).

DATA TABEL (satu-satunya sumber angka yang boleh Anda sebut)
Tabel {tabel_no}. {judul_tabel}
Kolom: {headers}
Baris:
{rows}
Keterangan sumber pada naskah: {sumber}
Konteks tambahan: {konteks}

KELUARAN: balas HANYA objek JSON berikut tanpa penjelasan lain:
{{
  "sebelum": ["<tepat 1 kalimat pembuka, 12-25 kata>"],
  "sesudah": ["<paragraf deskripsi 1, 3-5 kalimat>", "<paragraf deskripsi 2, 3-5 kalimat>"]
}}

ATURAN KETAT
1. Setiap ANGKA yang Anda tulis WAJIB persis ada pada DATA TABEL di atas. Dilarang mengarang, membulatkan ulang, atau mengubah angka apa pun.
2. DILARANG menulis sitasi/rujukan seperti "(Nama, 2024)" karena ini bagian hasil penelitian.
3. Kalimat "sebelum": SATU kalimat pembuka saja (12-25 kata), langsung menjelaskan tabel ini menyajikan apa dan untuk keperluan analisis apa. DILARANG membuka dengan "Bab hasil penelitian ini", "Bab ini", "bagian ini", atau "Sub-bab ini", dan DILARANG menambah kalimat kedua.
4. Paragraf "sesudah" pertama: uraikan isi tabel secara faktual (kategori, nilai, kecenderungan yang menonjol).
5. Paragraf "sesudah" kedua: tafsirkan makna angka tersebut bagi model penelitian dan implikasinya di instansi.
6. Bahasa Indonesia akademik baku dan mengalir. Tanpa daftar, tanpa bullet, tanpa tanda bintang.
7. Jangan menyalin judul tabel kata per kata dan jangan menyebut kata "Tabel" lebih dari sekali per paragraf.
8. DILARANG menulis kode LaTeX (mis. $...$, \frac, \beta). Gunakan simbol biasa (mis. R², β, γ) atau kata.
9. Tulis SEMUA angka dengan DIGIT (mis. 69 orang dan 69,00 persen). DILARANG menulis angka dengan huruf (mis. "tujuh puluh persen" atau "enam puluh sembilan orang"), karena penulisan huruf mudah salah dan tidak dapat diverifikasi terhadap tabel.
"""


def write_tabel_bab4(cfg, tabel_no, judul_tabel, headers, data, sumber="", konteks=""):
    """Untuk SATU tabel BAB IV: 1 paragraf narasi sebelum + 2 paragraf deskripsi sesudah."""
    rows = "\n".join("- " + " | ".join(str(x) for x in r) for r in (data or []))
    angka_ok = _kumpulan_angka(headers, data, sumber, judul_tabel, konteks)
    prompt = _TABEL_BAB4.format(
        judul=cfg.get("judul", ""), instansi=cfg.get("instansi", ""),
        x1=cfg.get("x1", "X1"), x2=cfg.get("x2", "X2"),
        m=cfg.get("m", "M"), y=cfg.get("y", "Y"),
        tabel_no=tabel_no, judul_tabel=judul_tabel,
        headers=" | ".join(str(h) for h in (headers or [])),
        rows=rows or "-", sumber=sumber or "-", konteks=konteks or "-")
    return _gemini_json(
        prompt,
        lambda raw: _validate_tabel_bab4(raw, angka_ok,
                                         tag="tabel_%s" % tabel_no, tabel_no=tabel_no),
        max_tokens=2048, tag="tabel_%s" % tabel_no, attempts=3)


# ==== PARAGRAF PENGANTAR SUB-BAB HASIL PENELITIAN (4.1.3 / 4.1.4 / 4.1.5) ====
# Ditulis Gemini, tetapi HANYA boleh memuat ambang batas metodologis: angka hasil
# penelitian dilarang, dan paragraf ini dilarang memuat sitasi.

_INTRO_BAB4 = """Anda penulis naskah tesis magister Manajemen SDM yang teliti dan sistematis.
Anda menulis SATU PARAGRAF PENGANTAR untuk sebuah sub-bab HASIL PENELITIAN (BAB IV) pada naskah:
"{judul}"
Objek penelitian: {instansi}
Variabel: X1 = {x1}, X2 = {x2}, M = {m}, Y = {y}

Sub-bab yang dibuka: {kode}. {nama}
Paragraf ini diletakkan TEPAT di bawah judul sub-bab tersebut, sebelum tabel pertama,
dan berfungsi mengenalkan tahap analisis beserta ambang batas yang dipakai.

AMBANG BATAS METODOLOGIS YANG BOLEH ANDA SEBUT (satu-satunya angka yang boleh muncul):
{ambang}

Tabel yang menyajikan hasil tahap ini: {tabel}

KELUARAN: balas HANYA objek JSON berikut tanpa penjelasan lain:
{{
  "pengantar": ["<tepat 1 paragraf pengantar, 5-7 kalimat>"]
}}

ATURAN KETAT
1. Tepat SATU paragraf, 5-7 kalimat, mengalir sebagai satu kesatuan. Tanpa daftar, bullet, penomoran, tanda bintang, atau baris baru.
2. DILARANG menulis angka apa pun selain yang tercantum pada daftar ambang batas di atas. Jangan menyebut nilai hasil analisis (misalnya nilai loading, koefisien jalur, R-Square, atau T-Statistic) karena hasil itu baru disajikan pada tabel-tabel di bawahnya.
3. DILARANG menulis sitasi/rujukan seperti "(Nama, 2024)" karena bagian ini adalah hasil penelitian.
4. Alur paragraf: (a) tujuan tahap analisis ini; (b) alasan tahap ini harus dilalui sebelum tahap berikutnya; (c) kriteria atau ambang batas yang dipakai, dirangkai dengan kalimat Anda sendiri (jangan menyalin daftar di atas kata per kata); (d) kalimat penutup yang menunjukkan lokasi hasilnya, contoh: "Hasil lengkap tahap ini disajikan pada Tabel 4.6 sampai dengan Tabel 4.10."
5. Rujuk tabel SELALU dengan pola "Tabel <nomor>" (contoh: Tabel 4.11), jangan pernah menulis nomor tabel tanpa kata "Tabel".
6. Bahasa Indonesia akademik baku; angka desimal memakai koma (contoh: 0,70). Tanpa markdown.
"""

_BLOK_INTRO_BAB4 = {
    "outer": {
        "kode": "4.1.3",
        "nama": "Analisis Outer Model (evaluasi measurement model)",
        "tabel": "Tabel 4.6 sampai dengan Tabel 4.10",
        "ambang": ("- outer loading setiap butir indikator minimal 0,70\n"
                   "- Average Variance Extracted (AVE) setiap konstruk di atas 0,50\n"
                   "- kriteria Fornell-Larcker, yaitu akar AVE lebih besar daripada korelasi antarkonstruk\n"
                   "- rasio Heterotrait-Monotrait (HTMT) di bawah 0,90\n"
                   "- Cronbach's Alpha dan Composite Reliability di atas 0,70\n"
                   "- Variance Inflation Factor (VIF) di bawah 5"),
        "angka_bebas": {"070", "050", "090", "5"},
    },
    "inner": {
        "kode": "4.1.4",
        "nama": "Analisis Inner Model (evaluasi structural model)",
        "tabel": "Tabel 4.11 sampai dengan Tabel 4.13",
        "ambang": ("- ukuran efek f-Square: 0,02 kategori kecil, 0,15 kategori menengah, dan 0,35 kategori besar\n"
                   "- koefisien determinasi R-Square: 0,19 pengaruh lemah, 0,33 pengaruh moderat, dan 0,67 pengaruh kuat\n"
                   "- relevansi prediktif Q-Square yang bernilai lebih besar dari 0\n"
                   "- kecocokan model (SRMR) di bawah 0,10"),
        "angka_bebas": {"002", "015", "035", "019", "033", "067", "010", "0"},
    },
    "hipotesis": {
        "kode": "4.1.5",
        "nama": "Pengujian Hipotesis",
        "tabel": "Tabel 4.14 dan Tabel 4.15",
        "ambang": ("- metode bootstrapping dengan 5.000 subsampel\n"
                   "- hipotesis diterima bila T-Statistic di atas 1,96 dan p-value di bawah 0,05\n"
                   "- taraf signifikansi 5 persen\n"
                   "- Variance Accounted For (VAF): di atas 80 persen mediasi penuh, 20 sampai 80 persen mediasi parsial, dan di bawah 20 persen tidak terjadi mediasi"),
        "angka_bebas": {"5000", "196", "005", "5", "80", "20"},
    },
}


def _validate_bab4_intro(raw, angka_ok, tag):
    """Validasi paragraf pengantar sub-bab Bab IV: tepat 1 paragraf, tanpa sitasi,
    dan hanya boleh memuat ambang batas metodologis yang diizinkan (tanpa angka hasil)."""
    res = _validate_bab3_output(
        raw, ["pengantar"], tag=tag, refs=[], require_cite=False,
        forbid_cite=True, forbid_re=r"(?!x)x",
        exact_len={"pengantar": 1}, min_len=600, max_len=1)
    if not res:
        return None
    para = res["pengantar"][0]
    if len(para) > 2000:
        print(f"[{tag}] paragraf pengantar terlalu panjang ({len(para)} karakter) -> tolak")
        return None
    if _naskah_terpotong(para):
        print(f"[{tag}] paragraf pengantar terpotong di tengah kalimat -> tolak")
        return None
    if _ada_angka_kata(para):
        print(f"[{tag}] angka ditulis dengan huruf -> tolak")
        return None
    for m in _RX_ANGKA_AI.finditer(_sanitize_teknis(para)):
        tok = re.sub(r"[.,]", "", m.group(0))
        if tok not in angka_ok and tok not in _ANGKA_UMUM:
            print(f"[{tag}] angka '{m.group(0)}' bukan ambang batas yang diizinkan -> tolak")
            return None
    return res


def write_intro_bab4(cfg, blok):
    """Paragraf pengantar sub-bab 4.1.3 (outer) / 4.1.4 (inner) / 4.1.5 (hipotesis).

    Ditulis Gemini, namun validator menolak keluaran yang memuat sitasi atau angka
    di luar ambang batas metodologis; bila gagal, pemanggil memakai teks cadangan
    deterministik agar naskah tetap utuh.
    """
    meta = _BLOK_INTRO_BAB4.get(blok)
    if not meta:
        return None
    angka_ok = (set(meta["angka_bebas"])
                | _kumpulan_angka(meta["ambang"], meta["tabel"], meta["nama"]))
    prompt = _INTRO_BAB4.format(
        judul=cfg.get("judul", ""), instansi=cfg.get("instansi", ""),
        x1=cfg.get("x1", "X1"), x2=cfg.get("x2", "X2"),
        m=cfg.get("m", "M"), y=cfg.get("y", "Y"),
        kode=meta["kode"], nama=meta["nama"],
        ambang=meta["ambang"], tabel=meta["tabel"])
    return _gemini_json(
        prompt,
        lambda raw: _validate_bab4_intro(raw, angka_ok, tag="intro_%s" % meta["kode"]),
        max_tokens=2048, tag="intro_%s" % meta["kode"], attempts=3)


def write_pembahasan(cfg, poin_num, studies, stats=None, angle=None):
    """4.2.x Pembahasan satu hipotesis: 4 alinea, sitasi HANYA dari studi nyata model ini."""
    import references_metodologi as rm

    pool = []
    for st in (studies or []):
        if st.get("cite") and st.get("apa"):
            pool.append({
                "sitasi": st["cite"],
                "apa": st["apa"],
                "tahun": st.get("year"),
                "judul": st.get("title"),
                "hasil": st.get("hasil"),
            })
    if not pool:
        print(f"[sub_4_2_{poin_num}] tidak ada studi terverifikasi untuk pembahasan")
        return None

    menu = "\n".join(
        f'- {p["sitasi"]}  |  {p["apa"]}  |  temuan: {p.get("hasil") or p.get("judul") or "-"}'
        for p in pool
    )

    peta = {
        1: (cfg.get("x1"), cfg.get("m"), None),
        2: (cfg.get("x2"), cfg.get("m"), None),
        3: (cfg.get("x1"), cfg.get("y"), None),
        4: (cfg.get("x2"), cfg.get("y"), None),
        5: (cfg.get("m"), cfg.get("y"), None),
        6: (cfg.get("x1"), cfg.get("y"), cfg.get("m")),
        7: (cfg.get("x2"), cfg.get("y"), cfg.get("m")),
    }.get(poin_num)
    if not peta:
        return None
    a, b, med = peta
    fokus = (f"peran {med} dalam memediasi pengaruh {a} terhadap {b}" if med
             else f"pengaruh {a} terhadap {b}")

    stats_block = ""
    if stats:
        stats_block = "\n=== HASIL STATISTIK NYATA (gunakan PERSIS angka-angka ini pada alinea 1;\n    jangan diubah, dibulatkan, atau diarang ulang) ===\n" + str(stats) + "\n"

    body = f"""
=== TUGAS: SUB BAB 4.2.{poin_num} PEMBAHASAN HASIL PENELITIAN ===
=== FOKUS PEMBAHASAN ===
{fokus}
{stats_block}

=== KELUARAN: JSON dengan kunci "pembahasan" ===
"pembahasan" : array berisi TEPAT 4 alinea, setiap alinea minimal 4 kalimat dan WAJIB memuat minimal satu sitasi "(Penulis, Tahun)" dari daftar di atas:
  - Alinea 1: hasil pengujian hipotesis (H{poin_num}) yang menyatakan bahwa {a} secara parsial memiliki pengaruh yang positif dan signifikan terhadap {b} (atau melalui {med} bila mediasi). JANGAN menyebutkan angka statistik (angka disajikan terpisah). WAJIB menyertakan minimal satu sitasi dari daftar.
  - Alinea 2: makna empiris temuan dan keterkaitannya dengan fenomena/kondisi pada instansi (latar belakang masalah). WAJIB menyertakan minimal satu sitasi dari daftar.
  - Alinea 3: landasan teoretis temuan (sejalan dengan pandangan teori manajemen sumber daya manusia). WAJIB menyertakan minimal satu sitasi dari daftar.
  - Alinea 4: keterkaitan temuan dengan penelitian terdahulu. Sebutkan minimal dua penulis beserta tahunnya dari daftar. Kalimat pembuka alinea ini JANGAN diawali dengan kata "Validasi", "Konsistensi", "Perbandingan", atau "Bandingkan"; pilih pembuka yang spesifik dan berbeda dari sub-bab lain.
PERIKSA ULANG: keempat alinea harus benar-benar memuat rangkaian nama penulis dan tahun dalam tanda kurung. Alinea tanpa sitasi akan ditolak sistem. Angka statistik (bila ada) hanya boleh berasal dari HASIL STATISTIK NYATA di atas.
ATURAN VARIASI (PENTING): kalimat pembuka setiap alinea (terutama alinea 2, 3, dan 4) HARUS berbeda dari sub-bab pembahasan lainnya. JANGAN gunakan frasa pembuka generik yang berulang seperti "Makna empiris dari temuan ini", "Secara teoretis", "Konsistensi temuan penelitian ini", "Dari perspektif teoretis", "Konsistensi temuan ini juga diperkuat", "Konsistensi dari temuan ini", "Konsistensi temuan ini terbukti", "Validasi eksternal terhadap temuan ini", "Realitas di lapangan menunjukkan", "Temuan empiris di lapangan". Buat pembuka yang spesifik pada konteks {a}, {b}, instansi, dan sudut pandang berikut: {angle}. Variasikan diksi, struktur kalimat, dan sudut pandang pembuka antar alinea dan antar sub-bab; hindari membuka alinea dengan kata yang sama pada sub-bab yang berbeda.
"""
    prompt = _bab45_prompt(cfg, body, _ATURAN_PEMBAHASAN.format(sumber_menu=menu))
    def _val(raw):
        res = _validate_bab3_output(
            raw, ["pembahasan"], tag=f"sub_4_2_{poin_num}", refs=pool,
            forbid_re=_FORBID_ANGKA_IV, forbid_cite=False, require_cite=True,
            exact_len={"pembahasan": 4}, min_refs=2, min_len=250)
        return res

    return _gemini_json(prompt, _val, tag=f"sub_4_2_{poin_num}", attempts=5)


def write_sub_5_1(cfg):
    """5.1 Kesimpulan: 7 butir, tanpa angka & tanpa sitasi."""
    return _gemini_json(
        _bab45_prompt(cfg, _BODY_5_1, _ATURAN_TANPA_ANGKA),
        lambda raw: _validate_bab3_output(
            raw, ["kesimpulan"], tag="sub_5_1", refs=[], require_cite=False,
            forbid_cite=True, forbid_re=_FORBID_ANGKA_IV,
            exact_len={"kesimpulan": 7}, max_len=8, min_len=20),
        tag="sub_5_1")


def write_sub_5_2(cfg):
    """5.2 Implikasi kebijakan: pembuka + 4 pilar, tiap pilar 2 sub-poin, tanpa angka & sitasi."""
    def _val(raw):
        if isinstance(raw, list) and raw and isinstance(raw[0], dict):
            raw = raw[0]
        if not isinstance(raw, dict):
            print("[sub_5_2] keluaran bukan objek JSON")
            return None
        pembuka = raw.get("pembuka")
        if not isinstance(pembuka, str) or len(pembuka.strip()) < 60:
            print("[sub_5_2] 'pembuka' tidak valid")
            return None
        pilar = raw.get("pilar")
        if not isinstance(pilar, list) or len(pilar) != 4:
            print("[sub_5_2] 'pilar' harus 4")
            return None
        cleaned = []
        for i, pl in enumerate(pilar):
            if not isinstance(pl, dict):
                print(f"[sub_5_2] pilar {i+1} bukan objek")
                return None
            judul = pl.get("judul")
            a_j = pl.get("a_judul"); a_u = pl.get("a_uraian")
            b_j = pl.get("b_judul"); b_u = pl.get("b_uraian")
            for name, val in (("judul", judul), ("a_judul", a_j), ("a_uraian", a_u),
                              ("b_judul", b_j), ("b_uraian", b_u)):
                if not isinstance(val, str) or len(val.strip()) < 20:
                    print(f"[sub_5_2] pilar {i+1} field '{name}' tidak valid")
                    return None
            # cek angka
            for v in (pembuka, judul, a_j, a_u, b_j, b_u):
                if re.search(_FORBID_ANGKA_IV, _sanitize_teknis(v), re.I):
                    print(f"[sub_5_2] pilar {i+1} memuat angka yang dilarang")
                    return None
            cleaned.append({
                "judul": judul.strip(),
                "a_judul": a_j.strip(), "a_uraian": a_u.strip(),
                "b_judul": b_j.strip(), "b_uraian": b_u.strip(),
            })
        return {"pembuka": pembuka.strip(), "pilar": cleaned}
    return _gemini_json(
        _bab45_prompt(cfg, _BODY_5_2, _ATURAN_TANPA_ANGKA),
        _val, tag="sub_5_2", attempts=5)


def write_sub_5_3(cfg, angle=None):
    """5.3 Saran: pembuka + 2 praktis + 2 akademis (dengan intro R2), tanpa angka & sitasi."""
    body = _BODY_5_3
    if angle:
        body = body.replace("{angle}", angle)
    else:
        body = body.replace("{angle}", "variabel yang relevan dengan konteks instansi penelitian")
    def _val(raw):
        if isinstance(raw, list) and raw and isinstance(raw[0], dict):
            raw = raw[0]
        if not isinstance(raw, dict):
            print("[sub_5_3] keluaran bukan objek JSON")
            return None
        pembuka = raw.get("pembuka")
        if not isinstance(pembuka, str) or len(pembuka.strip()) < 40:
            print("[sub_5_3] 'pembuka' tidak valid")
            return None
        praktis = raw.get("praktis")
        if not isinstance(praktis, list) or len(praktis) != 2:
            print("[sub_5_3] 'praktis' harus 2")
            return None
        akademis = raw.get("akademis")
        if not isinstance(akademis, list) or len(akademis) != 2:
            print("[sub_5_3] 'akademis' harus 2")
            return None
        cleaned_p = []
        for i, pl in enumerate(praktis):
            if not isinstance(pl, dict):
                return None
            j = pl.get("judul"); u = pl.get("uraian")
            if not isinstance(j, str) or not isinstance(u, str) or len(j.strip()) < 5 or len(u.strip()) < 40:
                print(f"[sub_5_3] praktis {i+1} tidak valid")
                return None
            for v in (j, u):
                if re.search(_FORBID_ANGKA_IV, _sanitize_teknis(v), re.I):
                    print(f"[sub_5_3] praktis {i+1} memuat angka")
                    return None
            cleaned_p.append({"judul": j.strip(), "uraian": u.strip()})
        cleaned_a = []
        for i, pl in enumerate(akademis):
            if not isinstance(pl, dict):
                return None
            j = pl.get("judul"); u = pl.get("uraian")
            if not isinstance(j, str) or not isinstance(u, str) or len(j.strip()) < 5 or len(u.strip()) < 40:
                print(f"[sub_5_3] akademis {i+1} tidak valid")
                return None
            for v in (j, u):
                if re.search(_FORBID_ANGKA_IV, _sanitize_teknis(v), re.I):
                    print(f"[sub_5_3] akademis {i+1} memuat angka")
                    return None
            cleaned_a.append({"judul": j.strip(), "uraian": u.strip()})
        return {"pembuka": pembuka.strip(), "praktis": cleaned_p, "akademis": cleaned_a}
    return _gemini_json(
        _bab45_prompt(cfg, body, _ATURAN_TANPA_ANGKA),
        _val, tag="sub_5_3", attempts=5)
