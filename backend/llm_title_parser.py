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
MODEL_CHAIN = [
    "gemini-flash-lite-latest",
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
    "gemini-flash-latest",
    "nvidia/nemotron-3-ultra-550b-a55b",
]

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
    """Muat konfigurasi API key dari llm_config.json."""
    if not os.path.exists(CONFIG_PATH):
        return {}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


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
