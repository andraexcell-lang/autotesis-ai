# backend/llm_title_parser.py
"""
LLM TITLE PARSER - Integrasi Google Gemini untuk /api/parse-title.

Arsitektur HYBRID berlapis:
  1. Google Gemini (utama)  -> pemahaman semantik penuh, mengurai judul bentuk apa pun.
  2. Regex Rule-Based       -> fallback otomatis saat API offline / kuota habis.

Pemilihan model (diuji 14 Okt 2026 pada akun pengguna):
  - gemini-4-argon & gemini-3.1-pro-preview : HTTP 429 (kuota tidak tersedia).
  - gemini-2.5-pro                          : HTTP 404 (tidak tersedia untuk akun baru).
  - gemini-2.5-flash                        : AKURAT, ~1.6 detik  -> UTAMA.
  - gemini-3.1-flash-lite                   : AKURAT, ~1.4 detik  -> CADANGAN 1.
  - gemini-3-flash-preview                  : lambat & terpotong  -> cadangan 2 (opsional).

API key dibaca dari file backend/llm_config.json (lihat llm_config.example.json).
Jangan hardcode API key di file ini.
"""

import json
import os
import ssl
import urllib.request
import urllib.error

BACKEND_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BACKEND_DIR, "llm_config.json")

GEMINI_BASE = "https://generativelanguage.googleapis.com/v1beta/models"

# Urutan model: yang paling pintar & tersedia lebih dulu.
#
# CATATAN (uji 14 Okt 2026 malam): "gemini-2.5-flash" dan "gemini-2.5-flash-lite"
# SUDAH DIHAPUS Google -- membalas 404 "no longer available" pada SEMUA key.
# Memakainya hanya membuang satu percobaan penuh di setiap rantai (termasuk
# jeda backoff), padahal hasilnya pasti gagal. Daftar di bawah semuanya teruji
# memberi HTTP 200 atau 503 sementara, tak ada yang 404.
#
# Urutan juga disusun dari pengukuran nyata: "gemini-flash-lite-latest" dan
# "gemini-3.5-flash-lite" terbukti 200 dan menang pada uji ekstraksi, sedangkan
# "gemini-3.1-flash-lite" saat itu sering 503 sehingga diletakkan belakangan
# (dengan pemutus sirkuit kesibukan agar tak membebankan jeda backoff).
# URUTAN BARU (diuji 21 Okt 2026, 1 key sehat, 16 model dicek langsung):
#   Model lama -> dipindah ke belakang karena ternyata SLOWEST di kelasnya:
#     "gemini-flash-lite-latest" 12,4 dtk (tapi dipakai karena paling stabil)
#     "gemini-3.5-flash-lite"    15,4 dtk
#   Model baru -> paling cepat &output JSON valid:
#     "gemini-3.1-flash-lite"     6,1 dtk  <- dipakai sebagai UTAMA
#     "gemini-3.8-flash"         11,5 dtk
#     "gemini-3.5-flash"         22,5 dtk
#   Catatan: "gemini-4-argon", "gemini-3.1-pro-preview", "gemini-pro-latest" TIDAK
#   bisa diakses (butuh allowlist preview); "gemini-2.5-flash"/"-lite"/"-pro"
#   sudah dihapus Google (HTTP 404) -> tidak dicantumkan agar tak buang percobaan.
#
# CATATAN PENTING - chain PEMULIS harus PATUH format sitasi:
#   "gemini-3.1-flash-lite" memang lebih cepat (6,1 dtk vs 12,4 dtk) tapi UJI
#   NYATA: model ini SERING MELEWATI sitasi di sebagian paragraf -> validasi
#   gagal 3x berturut-turut -> sub bab jatuh ke template minimal. Karena itu
#   model ini TIDAK dipakai sebagai primer untuk penulisan naskah.
#   "gemini-flash-lite-latest" terbukti paling patuh (sitasi lengkap, JSON
#   valid) -> tetap PRIMER.
# Model yang TIDAK bisa diakses (catat saja, jangan masuk chain):
#   "gemini-4-argon", "gemini-3.1-pro-preview", "gemini-pro-latest" -> butuh allowlist.
#   "gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-2.5-pro" -> HTTP 404 (dihapus).
MODEL_CHAIN = [
    "gemini-flash-lite-latest",
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
    "gemini-3.8-flash",
    "gemini-3.5-flash",
    "gemini-flash-latest",
    "nvidia/nemotron-3-ultra-550b-a55b",
]

# ============================================================================
# #4 EXTRACT_CHAIN — Nemotron DIDAHULUKAN untuk tugas EKSTRAKSI
# ============================================================================
# Alasan (dipilih, bukan asal):
#   * Nemotron = GRATIS (tidak menggerus kuota Gemini 15 key).
#   * Uji 14 Okt: ekstraksi Indonesia 5/5 kutipan literal COCOK.
#   * Latensi ~1 dtk vs 4.5 dtk jeda rate-limit Gemini.
#   * Bila NIM sibuk/503 -> pemutus sirkuit menandai sibuk 60 dtk lalu
#     TURUN ke model Gemini secara otomatis (tidak memblokir).
# Dipakai untuk: ekstraksi dimensi/indikator per paper, ontologi, teori.
# TIDAK dipakai untuk: parse judul & tulis Bab 1.1 (kualitas prosa kritis,
# hanya 1 panggilan -> lebih baik memakai Gemini).
# ============================================================================
EXTRACT_CHAIN = [
    "nvidia/nemotron-3-ultra-550b-a55b",   # GRATIS, prioritas #1
    "gemini-3.1-flash-lite",               # 6,1 dtk (tercepat & JSON valid)
    "gemini-3.8-flash",
    "gemini-flash-lite-latest",
    "gemini-3.5-flash-lite",
    "gemini-flash-latest",
]


def extract_chain():
    """Rantai model untuk tugas EKSTRAKSI (Nemotron didahulukan — #4).

    Sengaja TIDAK memakai cfg["models"]: config umum tetap diurutkan untuk
    kualitas prosa (tulis bab / parse judul), sedangkan ekstraksi butuh
    banyak panggilan -> hemat kuota Gemini.

    Override lewat env EXTRACT_MODELS (JSON array ATAU dipisah koma):
      EXTRACT_MODELS='["nvidia/nemotron-3-ultra-550b-a55b","gemini-flash-lite-latest"]'
    """
    raw = str(os.environ.get("EXTRACT_MODELS") or "").strip()
    if raw:
        if raw.startswith("["):
            try:
                parsed = json.loads(raw)
                out = [str(m).strip() for m in parsed if str(m).strip()]
                if out:
                    return out
            except (json.JSONDecodeError, TypeError):
                pass
        else:
            out = [m.strip() for m in raw.split(",") if m.strip()]
            if out:
                return out
    return list(EXTRACT_CHAIN)

DEFAULT_TIMEOUT = 12  # detik per panggilan model

PROMPT_TEMPLATE = """Anda adalah parser akademik untuk judul tesis Magister Manajemen (S2) Indonesia.

Ekstrak komponen struktural dari judul tesis berikut:
{title}

Aturan:
1. Variabel X1 dan X2 adalah dua variabel bebas (independen).
2. Variabel M adalah variabel mediasi/intervening (bila ada kata "melalui", "dimediasi oleh", "sebagai variabel intervening/mediator", atau setelah "terhadap M serta dampaknya pada Y").
3. Variabel Y adalah variabel terikat (dependen).
4. instansi adalah objek penelitian (dinas/badan/kantor/perusahaan/sekolah, lengkap dengan kabupaten/kota/provinsi bila ada).
5. Koreksi ejaan/typo pada istilah akademik (contoh: "organizatinal" -> "organizational"), tetapi JANGAN mengubah nama instansi (mis. "Ogan Komering Ilir" harus tetap apa adanya).
6. Jika sebuah komponen tidak ditemukan, isi dengan string kosong "".

HANYA balas dengan JSON valid, tanpa penjelasan lain, dengan format persis:
{{"x1": "", "x2": "", "m": "", "y": "", "instansi": ""}}"""


def load_config():
    """Muat konfigurasi API key dari llm_config.json.

    Di cloud (Render/Fly.io) file llm_config.json TIDAK ikut ter-deploy
    (dilindungi .gitignore), jadi nilai dibaca dari Environment Variables:

      GEMINI_API_KEYS  -> "api_keys"  (JSON array ATAU dipisah koma)
      GEMINI_API_KEY   -> "api_key"   (key tunggal, legacy)
      NVIDIA_API_KEY   -> "nvidia_api_key"
      LLM_MODELS       -> "models"    (JSON array ATAU dipisah koma)
      LLM_ENABLED      -> "enabled"   ("true"/"false")
      LLM_TIMEOUT      -> "timeout_seconds"
    """
    cfg = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except (json.JSONDecodeError, OSError):
            cfg = {}
    if not isinstance(cfg, dict):
        cfg = {}

    def _split(val):
        """Terima JSON array atau string dipisah koma."""
        val = str(val or "").strip()
        if not val:
            return None
        if val.startswith("["):
            try:
                parsed = json.loads(val)
                return [str(x).strip() for x in parsed if str(x).strip()]
            except (json.JSONDecodeError, TypeError):
                pass
        return [p.strip() for p in val.split(",") if p.strip()]

    env_keys = _split(os.environ.get("GEMINI_API_KEYS"))
    if env_keys:
        cfg["api_keys"] = env_keys
    env_one = str(os.environ.get("GEMINI_API_KEY") or "").strip()
    if env_one:
        cfg["api_key"] = env_one
    env_nv = str(os.environ.get("NVIDIA_API_KEY") or "").strip()
    if env_nv:
        cfg["nvidia_api_key"] = env_nv
    env_models = _split(os.environ.get("LLM_MODELS"))
    if env_models:
        cfg["models"] = env_models
    if "LLM_ENABLED" in os.environ:
        cfg["enabled"] = str(os.environ["LLM_ENABLED"]).lower() in ("1", "true", "yes", "on")
    if str(os.environ.get("LLM_TIMEOUT") or "").strip().isdigit():
        cfg["timeout_seconds"] = int(os.environ["LLM_TIMEOUT"])
    return cfg


def gemini_keys(cfg):
    """Daftar API key Gemini yang sah, untuk ROTASI.

    Membaca "api_keys" (daftar) lalu "api_key" (tunggal, untuk config lama),
    membuang yang kosong atau masih berupa placeholder "GANTI_...".

    PENTING: rotasi ini untuk KETAHANAN, BUKAN untuk melipatgandakan kuota.
    Uji 14 Okt 2026 membuktikan 14 key memberi 429 kuota IDENTIK pada dua model
    berbeda -> kuotanya berbagi satu proyek Google. Jadi rotasi berguna ketika
    sebuah key DICABUT/ditolak (HTTP 401/403), bukan saat kuota habis.
    """
    keys = []
    for k in cfg.get("api_keys") or []:
        k = str(k or "").strip()
        if k and not k.startswith("GANTI_") and k not in keys:
            keys.append(k)
    k1 = str(cfg.get("api_key") or "").strip()
    if k1 and not k1.startswith("GANTI_") and k1 not in keys:
        keys.insert(0, k1)
    return keys


def _next_key(current):
    """Key berikutnya dari "api_keys" (melingkar). None bila tak ada cadangan."""
    keys = gemini_keys(load_config())
    if not keys:
        return None
    if current not in keys:
        return keys[0]
    return keys[(keys.index(current) + 1) % len(keys)]


def _call_gemini(api_key, model, title, timeout):
    """Panggil satu model. Mengembalikan dict hasil atau melempar Exception.

    Model berawalan "nvidia/" diarahkan ke endpoint NVIDIA (fallback); sisanya
    ke Gemini. Import dilakukan secara malas untuk menghindari siklus impor.

    Bila sebuah key DITOLAK (HTTP 401/403), otomatis berpindah ke key berikutnya
    dari daftar "api_keys" -- sama seperti perilaku _call_gemini_raw di
    llm_ontology_researcher, agar parser judul ikut kebal bila ada key dicabut.
    """
    prompt = PROMPT_TEMPLATE.format(title=title)
    if str(model).startswith("nvidia/"):
        from llm_ontology_researcher import _call_nvidia_json
        return _call_nvidia_json(model, prompt, timeout, max_tokens=400)

    payload = {
        "contents": [{
            "parts": [{"text": prompt}]
        }],
        "generationConfig": {
            "responseMimeType": "application/json",
            "temperature": 0.1,
            "maxOutputTokens": 400,
        },
    }
    ctx = ssl.create_default_context()
    sisa_rotasi = 3          # batas berpindah key akibat 401/403
    while True:
        # URL dibentuk di dalam loop agar rotasi key dapat mengganti kuncinya.
        url = f"{GEMINI_BASE}/{model}:generateContent?key={api_key}"
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                body = json.loads(resp.read().decode("utf-8"))
            break
        except urllib.error.HTTPError as e:
            if e.code in (401, 403) and sisa_rotasi > 0:
                lain = _next_key(api_key)
                if lain and lain != api_key:
                    sisa_rotasi -= 1
                    api_key = lain
                    continue
            raise

    text = body["candidates"][0]["content"]["parts"][0]["text"]
    return json.loads(text)


def parse_title_llm(raw_title):
    """
    Parse judul memakai Gemini, berantai antar model.
    Mengembalikan dict {"x1","x2","m","y","instansi"} ATAU None bila semua model gagal.
    """
    if not raw_title or not raw_title.strip():
        return None

    cfg = load_config()
    # Pakai daftar key (bukan hanya "api_key" tunggal), agar config lama tetap
    # jalan dan config dengan "api_keys" ikut terpakai.
    api_key = (gemini_keys(cfg) or [""])[0]
    if not api_key:
        return None

    enabled = cfg.get("enabled", True)
    if not enabled:
        return None

    chain = [m for m in cfg.get("models", []) if m] or MODEL_CHAIN
    timeout = int(cfg.get("timeout_seconds", DEFAULT_TIMEOUT))

    for model in chain:
        try:
            result = _call_gemini(api_key, model, raw_title, timeout)
            parsed = {
                "x1": (result.get("x1") or "").strip(),
                "x2": (result.get("x2") or "").strip(),
                "m": (result.get("m") or "").strip(),
                "y": (result.get("y") or "").strip(),
                "instansi": (result.get("instansi") or "").strip(),
            }
            # Validasi minimal: minimal Y atau X1 harus terisi agar dianggap sukses
            if parsed["y"] or parsed["x1"]:
                parsed["_engine"] = f"gemini:{model}"
                return parsed
        except Exception:
            # Model ini gagal (kuota/timeout/format) -> coba model berikutnya
            continue

    return None


if __name__ == "__main__":
    # Uji cepat: python llm_title_parser.py
    test_title = ("Kecerdasan Emosional Dan Kompetensi Terhadap Kinerja Pegawai "
                  "Dinas Pendidikan Kabupaten Ogan Komering Ilir "
                  "Melalui Organizatinal Citizenship Behavior")
    res = parse_title_llm(test_title)
    if res:
        print("HASIL LLM:", json.dumps(res, ensure_ascii=False, indent=2))
    else:
        print("LLM tidak tersedia (cek llm_config.json / koneksi) -> fallback regex aktif")
