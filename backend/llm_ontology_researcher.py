# backend/llm_ontology_researcher.py
"""
AGEN RISET ONTOLOGI BERBASIS GOOGLE GEMINI.

Tugas agen (sesuai roadmap pengembangan AutoTesis AI):
  1. Menerima nama variabel penelitian (X1, X2, M, Y) apa pun bentuknya.
  2. Meriset Grand Theory, tokoh ahli, definisi operasional, dimensi, dan
     indikator Likert dari pengetahuan Gemini (menggantikan dict hardcode 12 variabel).
  3. Memverifikasi & melengkapi jurnal rujukan ber-DOI NYATA via OpenAlex.
  4. Menyimpan hasil ke cache lokal agar hemat kuota API.

Jaminan anti-halusinasi:
  - Dimensi/indikator dihasilkan LLM (terlatih pada ribuan jurnal), namun
    jurnal rujukan HANYA diambil dari OpenAlex (DOI nyata) atau pool lokal
    terverifikasi. LLM dilarang mengarang DOI.

Arsitektur hybrid berlapis:
  Lapis 1: Gemini  -> riset dimensi/indikator apapun variabelnya.
  Lapis 2: Knowledge base lokal (12 variabel MSDM) -> cadangan offline.
"""

import json
import os
import re
import ssl
import time
import urllib.error
import urllib.request

from llm_title_parser import load_config, MODEL_CHAIN, GEMINI_BASE, gemini_keys

BACKEND_DIR = os.path.dirname(os.path.abspath(__file__))
ONTO_CACHE_FILE = os.path.join(BACKEND_DIR, "cache_llm_ontology.json")

ONTO_TIMEOUT = 25  # detik; prompt ontologi lebih panjang dari parsing judul

# Rate limiting Gemini (kuota per-menit; mencegah 429/503 saat burst kandidat)
GEMINI_MIN_INTERVAL = 4.5   # detik minimal antar panggilan (~13 RPM, aman di bawah 15 RPM)
GEMINI_MAX_RETRY = 1        # satu percobaan ulang INTERNAL saat 429/503
GEMINI_BACKOFF = 6.0        # detik; jeda tunggal sebelum percobaan ulang
GEMINI_KEY_ROTATE = 3       # batas berpindah API key akibat key DITOLAK (401/403)
_last_gemini_ts = 0.0       # timestamp panggilan Gemini terakhir (modul-level)

# --- Fallback: NVIDIA NIM (endpoint kompatibel OpenAI) -------------------
# Dipakai BILA Gemini gagal (mis. kuota free-tier habis). Model NVIDIA dikenali
# lewat awalan "nvidia/" pada nama model di llm_config.json.
NVIDIA_BASE = "https://integrate.api.nvidia.com/v1"
NVIDIA_MIN_INTERVAL = 1.5   # detik minimal antar panggilan NVIDIA
_last_nvidia_ts = 0.0

# Circuit breaker kuota Gemini: saat 429 berbunyi "exceeded your current quota",
# kuota habis berjam-jam -> jangan buang waktu backoff, lompat cepat ke fallback.
# Dicatat PER MODEL, karena tiap model punya kuota sendiri: kuota habis pada satu
# model tidak boleh memblokir model lain yang masih sehat.
_gemini_quota_dead_until = {}   # {model: epoch saat boleh dicoba lagi}
GEMINI_QUOTA_DEAD_SEC = 600.0   # anggap mati 10 menit; reset asli bisa lebih lama

# Pemutus sirkuit KESIBUKAN (terpisah dari kuota): HTTP 503 berarti server Gemini
# sedang kewalahan. Bila satu percobaan ulang pun tetap 503, model itu layak
# dilewati sejenak -- tanpa ini, model 503-menahun membebankan jeda backoff
# (~22 dtk) pada SETIAP panggilan sepanjang rantai, memperlambat semua telusur.
_gemini_busy_until = {}          # {model: epoch saat boleh dicoba lagi}
GEMINI_BUSY_DEAD_SEC = 120.0     # 2 menit; 503 umumnya pulih cepat

PROMPT_TEMPLATE = """Anda adalah ahli metodologi penelitian kuantitatif Magister Manajemen (S2) Indonesia yang memahami standar penulisan tesis kelas Yurahadi/Tridinanti.

Konteks tesis:
Judul  : {judul}
Variabel: {var_name} (peran: {role} dalam model mediasi X1, X2 -> M -> Y)
Objek  : {instansi}

Tugas: Keluarkan ontologi akademik untuk variabel tersebut dalam format JSON berikut:

{{
  "grand_theory": "Teori induk paling relevan lengkap dengan tokoh & tahun, contoh: Goal-Setting Theory (Locke & Latham, 1990)",
  "tokoh_ahli": "2-4 kutipan ahli kunci dengan tahun, NAMA LENGKAP (Nama Depan + Nama Belakang), contoh: Anwar Prabu Mangkunegara (2022); Kasmir (2023)",
  "definisi": "Definisi operasional 1-2 kalimat gaya kajian pustaka tesis",
  "dimensi": [
    {{
      "nama": "Nama Dimensi 1",
      "indikator": [
        {{"t": "indikator yang bisa dinilai responden", "p": "Pernyataan kuesioner Likert 1-5 untuk indikator ini"}},
        {{"t": "indikator kedua", "p": "Pernyataan kuesioner untuk indikator kedua"}}
      ]
    }},
    {{"nama": "Nama Dimensi 2", "indikator": [{{"t": "...", "p": "..."}}, {{"t": "...", "p": "..."}}]}}
  ]
}}

Aturan KERAS:
1. Hasilkan 3-4 dimensi; setiap dimensi 2-3 indikator (total indikator 8-12 butir).
2. "t" (teks indikator) berupa FENOMENA perilaku/keadaan yang bisa dinilai responden dengan Likert 1-5 (bukan definisi teori).
3. "p" (pernyataan kuesioner) adalah KALIMAT PERNYATAAN LIKERT penuh, subjek PERSETUJUAN responden, bersifat netral/positif, berbahasa Indonesia baku, dimulai dengan "Saya ...", siap dipakai langsung di angket. Jangan mengulang kata "mengoptimalkan". Contoh baik: "Saya mampu menyelesaikan pekerjaan sesuai target yang ditetapkan instansi."
4. Gaya akademik Indonesia, istilah baku, tanpa jargon Inggris kecuali memang baku (contoh: Organizational Citizenship Behavior).
5. Jika variabel adalah variabel mediasi (M), dimensi harus menggambarkan masalah operasional yang menjembatani X ke Y.
6. JIKA benar-benar bukan variabel penelitian manajemen SDM yang lazim, tetap berikan ontologi paling mendekati berdasarkan teori yang ada.
7. JANGAN mengarang nama jurnal atau DOI. Jurnal rujukan akan diverifikasi terpisah oleh sistem.
8. HANYA balas JSON valid tanpa teks lain."""


# --- PETA RUMPUN VARIABEL (anti-halusinasi semantik) -------------------------
# Masalah nyata yang pernah terjadi: Gemini meriset "Disiplin Kerja" tetapi
# mengembalikan dimensi Motivasi (Kebutuhan Berprestasi/Afiliasi/Kekuasaan).
# Tanpa pemeriksaan ini, kesalahan konstruk masuk diam-diam ke Bab II/III.
# Aturan: minimal satu nama dimensi variabel harus mengandung kata kunci
# rumpunnya sendiri. Hasil pelanggaran DITOLAK -> model berikutnya -> KB lokal.
TOPIC_KEYWORDS = {
    "disiplin": ["disiplin", "ketaatan", "kepatuhan", "tata tertib", "sop",
                 "prosedur", "kode etik", "peraturan", "jam kerja", "inventaris", "absensi"],
    "kepuasan_kerja": ["kepuasan", "pekerjaan", "kompensasi", "penghargaan", "gaji",
                       "lingkungan", "rekan", "pengembangan", "karier", "promosi",
                       "hubungan", "beban", "kemajuan", "tuntutan"],
    "ocb": ["altru", "conscientious", "kesadaran", "sportivitas", "civic",
            "kewargaan", "inisiatif", "partisipasi", "relawan"],
    "motivasi": ["motivasi", "kebutuhan", "prestasi", "afiliasi", "kekuasaan",
                 "penghargaan", "dorongan", "insentif", "harapan", "tujuan", "goal"],
    "kinerja": ["kinerja", "kualitas", "kuantitas", "produktivitas", "ketepatan",
                "target", "waktu", "tugas", "efektivitas"],
    "promosi_jabatan": ["promosi", "kenaikan", "pangkat", "jabatan", "jenjang",
                        "prosedural", "karier", "kesesuaian", "peluang"],
    "mutasi_kerja": ["mutasi", "rotasi", "penempatan", "alih tugas", "perpindahan",
                     "penyegaran", "penugasan"],
    "pengembangan_karier": ["karier", "karir", "bimbingan", "perencanaan", "mentor",
                            "promosi", "pengembangan", "pelatihan", "kompetensi"],
    "budaya_organisasi": ["budaya", "inovasi", "risiko", "orientasi", "stabilitas",
                          "tim", "kekeluargaan", "birokrasi", "hasil", "detail"],
    "engagement_komitmen": ["vigor", "dedication", "absorption", "semangat", "dedikasi",
                            "keterikatan", "komitmen", "afektif", "normatif", "berkelanjutan"],
    "kecerdasan_emosional": ["kesadaran diri", "pengaturan diri", "empati",
                             "keterampilan sosial", "emosi", "motivasi diri"],
    "beban_kerja": ["beban", "tuntutan", "mental", "fisik", "stres", "tekanan",
                    "frustrasi", "waktu"],
    "pelatihan": ["materi", "metode", "instruktur", "pelatih", "fasilitas", "sarana",
                  "pelatihan", "diklat", "simulasi", "praktik"],
}

# Urut dari paling spesifik ke paling umum (dicek berurutan, yang pertama menang)
_VAR_TOPIC_MAP = [
    ("kepuasan_kerja", ["kepuasan"]),
    ("ocb", ["citizenship", "ocb", "kewargaan"]),
    ("promosi_jabatan", ["promosi", "kenaikan pangkat"]),
    ("mutasi_kerja", ["mutasi", "rotasi", "tour of duty"]),
    ("pengembangan_karier", ["karier", "karir", "career"]),
    ("budaya_organisasi", ["budaya", "culture"]),
    ("disiplin", ["disiplin", "tata tertib", "kepatuhan"]),
    ("engagement_komitmen", ["engagement", "keterikatan", "komitmen"]),
    ("kecerdasan_emosional", ["emosional", "emotional", "kecerdasan emosi"]),
    ("beban_kerja", ["beban", "workload", "stres kerja"]),
    ("pelatihan", ["pelatihan", "training", "diklat", "workshop"]),
    ("motivasi", ["motivasi", "motivation"]),
    ("kinerja", ["kinerja", "performance", "prestasi kerja", "produktivitas", "productivity", "efektivitas kerja"]),
]


def _match_variable_topic(var_name):
    """Petakan nama variabel ke rumpun KB. None = variabel bebas (lolos tanpa cek)."""
    v = (var_name or "").lower().strip()
    for topic, kws in _VAR_TOPIC_MAP:
        if any(k in v for k in kws):
            return topic
    return None


def _dimensions_fit_topic(dims, topic):
    """
    True bila minimal satu nama dimensi menyentuh kata kunci rumpun variabel.
    Tanpa topic / tanpa keyword -> dianggap lolos (variabel di luar KB).
    """
    kws = TOPIC_KEYWORDS.get(topic)
    if not kws or not dims:
        return True
    for d in dims:
        blob = str(d.get("nama", "")).lower()
        if any(k in blob for k in kws):
            return True
    return False


def _next_gemini_key(current):
    """Key berikutnya dari daftar "api_keys" di config (melingkar dari `current`).

    Dipakai hanya saat key DITOLAK (HTTP 401/403). Mengembalikan None bila tak
    ada key cadangan, sehingga pemanggil tetap melempar error aslinya.
    """
    try:
        keys = gemini_keys(load_config())
    except Exception:
        return None
    if not keys:
        return None
    if current not in keys:
        return keys[0]
    return keys[(keys.index(current) + 1) % len(keys)]


def _call_gemini_raw(api_key, model, prompt, timeout, max_tokens=2048):
    """Panggil Gemini (generateContent) dengan instruksi JSON mode.

    Menerapkan rate limiting client-side (jeda minimal antar-panggilan) serta
    backoff internal untuk HTTP 429/503, sehingga burst kandidat tidak langsung
    membanjiri kuota per-menit API key.

    Circuit breaker kuota: bila 429 berbunyi "exceeded your current quota", kuota
    habis selama berjam-jam -> langsung melempar (tanpa membuang waktu backoff)
    agar pemanggil dapat segera jatuh ke fallback NVIDIA.
    Melempar Exception bila gagal.

    ROTASI KEY (ketahanan): bila sebuah key DITOLAK (HTTP 401/403 -- dicabut atau
    salah proyek), otomatis berpindah ke key berikutnya dari "api_keys" tanpa
    membuang percobaan backoff. Catatan jujur: rotasi ini TIDAK melipatgandakan
    kuota -- uji 14 key (14 Okt 2026) membuktikan semuanya memberi 429 kuota
    IDENTIK pada dua model berbeda, artinya kuotanya berbagi satu proyek Google.
    Rotasi hanya menolong kalau satu key dicabut di antara banyak key sehat.
    """
    global _last_gemini_ts
    # Lewati CEPAT hanya model yang kuotanya terbukti habis (per model, bukan global)
    if time.time() < _gemini_quota_dead_until.get(model, 0.0):
        raise RuntimeError(f"kuota Gemini '{model}' masih habis -> langsung ke fallback")
    # Lewati CEPAT model yang baru saja terbukti kewalahan (503 beruntun)
    if time.time() < _gemini_busy_until.get(model, 0.0):
        raise RuntimeError(f"model '{model}' terbukti sibuk (503) -> langsung ke model berikutnya")
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "responseMimeType": "application/json",
            "temperature": 0.2,
            "maxOutputTokens": max_tokens,
        },
    }
    ctx = ssl.create_default_context()
    last_err = None
    sisa_rotasi = GEMINI_KEY_ROTATE   # sisa jatah berpindah key akibat 401/403
    for attempt in range(GEMINI_MAX_RETRY + 1):
        # Jaga jeda minimal antar-panggilan Gemini (rate limiting client-side)
        wait = GEMINI_MIN_INTERVAL - (time.time() - _last_gemini_ts)
        if wait > 0:
            time.sleep(wait)
        # URL dibentuk di sini (bukan sebelum loop) agar rotasi key dapat
        # mengganti kunci tanpa harus menyusun ulang panggilan dari awal.
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
            _last_gemini_ts = time.time()
            text = body["candidates"][0]["content"]["parts"][0]["text"]
            return json.loads(text)
        except urllib.error.HTTPError as e:
            _last_gemini_ts = time.time()
            last_err = e
            try:
                detail = e.read().decode("utf-8", "ignore")
            except Exception:
                detail = ""
            # Kuota harian habis: tandai model ini mati & lempar segera (retry percuma)
            if e.code == 429 and "exceeded your current quota" in detail.lower():
                _gemini_quota_dead_until[model] = time.time() + GEMINI_QUOTA_DEAD_SEC
                print(f"[gemini] kuota '{model}' habis -> lewati model ini {GEMINI_QUOTA_DEAD_SEC:.0f}s")
                raise
            # Key DITOLAK (dicabut / salah proyek) -> rotasi ke key berikutnya.
            # Sengaja TIDAK memakai jatah backoff 429/503: penyebabnya berbeda,
            # dan berpindah key seharusnya segera dicoba, bukan menunggu.
            if e.code in (401, 403) and sisa_rotasi > 0:
                lain = _next_gemini_key(api_key)
                if lain and lain != api_key:
                    sisa_rotasi -= 1
                    api_key = lain
                    print(f"[gemini] key ditolak (HTTP {e.code}) -> rotasi ke key "
                          f"cadangan ({sisa_rotasi} sisa)")
                    continue
            # 503 (server kewalahan) yang tetap gagal SETELAH dicoba ulang ->
            # tandai model sibuk sejenak, supaya panggilan berikutnya langsung
            # pindah model tanpa membuang waktu backoff pada model yang sama.
            if e.code == 503 and attempt >= GEMINI_MAX_RETRY:
                _gemini_busy_until[model] = time.time() + GEMINI_BUSY_DEAD_SEC
                print(f"[gemini] {model} sibuk (503) -> lewati model ini {GEMINI_BUSY_DEAD_SEC:.0f}s")
            # 429 (rate) & 503 (sibuk) layak dicoba ulang; error lain langsung lempar
            if e.code in (429, 503) and attempt < GEMINI_MAX_RETRY:
                backoff = GEMINI_BACKOFF * (2 ** attempt)
                print(f"[gemini] {model} HTTP {e.code} -> tunggu {backoff:.0f}s, ulang ({attempt+1}/{GEMINI_MAX_RETRY})")
                time.sleep(backoff)
                continue
            raise
    # Tercapai jika seluruh percobaan internal melempar 429/503
    raise last_err


def _extract_json(text):
    """Ambil objek JSON pertama dari teks LLM yang TIDAK memakai JSON mode.

    Menangani code-fence ```json ... ```, teks prolog di sekitarnya, serta
    struktur bersarang. Melempar ValueError bila tak ada JSON valid.
    """
    if not text or not str(text).strip():
        raise ValueError("output LLM kosong")
    s = str(text).strip()
    m = re.search(r"```(?:json)?\s*(.*?)```", s, re.DOTALL | re.IGNORECASE)
    if m:
        s = m.group(1).strip()
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        pass
    # Petakan kurung buka pertama yang seimbang (menghormati string & escape)
    for opener, closer in (("{", "}"), ("[", "]")):
        start = s.find(opener)
        if start == -1:
            continue
        depth, in_str, esc = 0, False, False
        for i in range(start, len(s)):
            ch = s[i]
            if in_str:
                if esc:
                    esc = False
                elif ch == "\\":
                    esc = True
                elif ch == '"':
                    in_str = False
                continue
            if ch == '"':
                in_str = True
            elif ch == opener:
                depth += 1
            elif ch == closer:
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(s[start:i + 1])
                    except json.JSONDecodeError:
                        break
    raise ValueError("tidak ditemukan JSON valid pada output LLM")


def _call_nvidia_json(model, prompt, timeout, max_tokens=2048):
    """Panggil model NVIDIA via endpoint NIM (kompatibel OpenAI).

    API key diambil dari config ("nvidia_api_key"). `enable_thinking: false`
    dipakai agar model menjawab langsung dengan isi, bukan jejak penalaran,
    sehingga JSON ekstraksi tidak tercampur reasoning. Melempar Exception bila gagal.
    """
    global _last_nvidia_ts
    cfg = load_config()
    api_key = str(cfg.get("nvidia_api_key") or "").strip()
    if not api_key:
        raise RuntimeError("nvidia_api_key belum diisi di llm_config.json")
    url = f"{NVIDIA_BASE}/chat/completions"
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.2,
        "max_tokens": max_tokens,
        "chat_template_kwargs": {"enable_thinking": False},
    }
    ctx = ssl.create_default_context()
    last_err = None
    for attempt in range(GEMINI_MAX_RETRY + 1):
        wait = NVIDIA_MIN_INTERVAL - (time.time() - _last_nvidia_ts)
        if wait > 0:
            time.sleep(wait)
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                body = json.loads(resp.read().decode("utf-8"))
            _last_nvidia_ts = time.time()
            text = body["choices"][0]["message"]["content"]
            try:
                return _extract_json(text)
            except ValueError as je:
                last_err = je
                if attempt < GEMINI_MAX_RETRY:
                    print(f"[nvidia] {model} output bukan JSON -> ulang ({attempt+1}/{GEMINI_MAX_RETRY})")
                    continue
                raise
        except urllib.error.HTTPError as e:
            _last_nvidia_ts = time.time()
            last_err = e
            if e.code in (429, 503) and attempt < GEMINI_MAX_RETRY:
                backoff = GEMINI_BACKOFF * (2 ** attempt)
                print(f"[nvidia] {model} HTTP {e.code} -> tunggu {backoff:.0f}s, ulang ({attempt+1}/{GEMINI_MAX_RETRY})")
                time.sleep(backoff)
                continue
            raise
    raise last_err


def _call_llm_json(api_key, model, prompt, timeout, max_tokens=2048):
    """Dispatcher LLM satu pintu untuk seluruh pemanggil (tracer, teori, Bab I).

    Model berawalan "nvidia/" -> endpoint NVIDIA (fallback); selainnya -> Gemini.
    """
    if str(model).startswith("nvidia/"):
        return _call_nvidia_json(model, prompt, timeout, max_tokens)
    return _call_gemini_raw(api_key, model, prompt, timeout, max_tokens)


# ============================================================================
# MATRIX DISPATCHER — strategi "habiskan SEMUA key pada model, baru pindah model"
# ============================================================================
# Filosofi (permintaan user): untuk model terpintar (mis. 3.8-flash),
# coba SEMUA 15 key berurutan sampai SEMUA kena 429 kuota.
# Hanya jika SEMUA key habis pada model itu, baru turun ke model berikutnya.
# Ini memaksimalkan pemakaian model terpintar di seluruh pool key.
# ============================================================================

# Lacak key yang sudah kena 429 kuota PER MODEL: {model: {key_index: epoch}}
_gemini_quota_dead_matrix = {}   # {model: {key_idx: epoch_until}}
GEMINI_QUOTA_DEAD_SEC = 600.0    # 10 menit; reset asli bisa lebih lama

# Lacak key yang kena 503 (sibuk) PER MODEL: {model: {key_idx: epoch}}
_gemini_busy_matrix = {}         # {model: {key_idx: epoch_until}}
GEMINI_BUSY_DEAD_SEC = 120.0     # 2 menit


def _call_gemini_matrix(keys, models, prompt, timeout, max_tokens=2048):
    """
    STRATEGI MATRIKS: untuk setiap model (urut prioritas), coba SEMUA key.
    
    Alur:
      for model in models:
          for key_idx, key in enumerate(keys):
              if key+model kena 429 kuota -> tandai mati, coba key berikutnya
              if key+model kena 503 (setelah retry) -> tandai sibuk, coba key berikutnya
              if 401/403 -> rotasi key (existing behavior)
              if 200 -> KEMBALIKAN HASIL
          # semua key habis untuk model ini -> lanjut ke model berikutnya
      raise last_error
    """
    global _last_gemini_ts, _gemini_quota_dead_matrix, _gemini_busy_matrix
    from llm_title_parser import load_config, GEMINI_BASE
    import ssl
    
    ctx = ssl.create_default_context()
    last_err = None
    
    for model in models:
        if str(model).startswith("nvidia/"):
            # NVIDIA tidak punya banyak key -> gunakan dispatcher lama
            return _call_nvidia_json(model, prompt, timeout, max_tokens)
        
        # Lewati model jika SEMUA key sudah diketahui mati kuota (per model)
        dead_keys = _gemini_quota_dead_matrix.get(model, {})
        if len(dead_keys) >= len(keys):
            # semua key diketahui habis kuota pada model ini -> lewati model
            print(f"[matrix] model '{model}': SEMUA {len(keys)} key kena kuota -> lewati model")
            continue
        
        print(f"[matrix] === MODEL: {model} (mencoba {len(keys)} key) ===", flush=True)
        
        for key_idx, api_key in enumerate(keys):
            # Cek apakah key ini sudah diketahui habis kuota pada model ini
            if _gemini_quota_dead_matrix.get(model, {}).get(key_idx, 0) > time.time():
                print(f"[matrix]   key #{key_idx+1} dilewati (kuota masi habis untuk {model})")
                continue
            if _gemini_busy_matrix.get(model, {}).get(key_idx, 0) > time.time():
                print(f"[matrix]   key #{key_idx+1} dilewati (sibuk 503 untuk {model})")
                continue
            
            print(f"[matrix]   mencoba key #{key_idx+1}...", flush=True)
            
            payload = {
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {
                    "responseMimeType": "application/json",
                    "temperature": 0.2,
                    "maxOutputTokens": max_tokens,
                },
            }
            
            last_err = None
            for attempt in range(GEMINI_MAX_RETRY + 1):
                # Rate limit global (antara panggilan apa pun)
                wait = GEMINI_MIN_INTERVAL - (time.time() - _last_gemini_ts)
                if wait > 0:
                    time.sleep(wait)
                
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
                    _last_gemini_ts = time.time()
                    text = body["candidates"][0]["content"]["parts"][0]["text"]
                    print(f"[matrix]   [OK] SUKSES key #{key_idx+1} pada {model}", flush=True)
                    return json.loads(text)
                except urllib.error.HTTPError as e:
                    _last_gemini_ts = time.time()
                    last_err = e
                    try:
                        detail = e.read().decode("utf-8", "ignore")
                    except Exception:
                        detail = ""
                    
                    # 429 QUOTA: tandai key+model ini mati
                    if e.code == 429 and "exceeded your current quota" in detail.lower():
                        _gemini_quota_dead_matrix.setdefault(model, {})[key_idx] = time.time() + GEMINI_QUOTA_DEAD_SEC
                        print(f"[matrix]   [X] key #{key_idx+1} KUOTA HABIS ({model}) -> tandai mati {GEMINI_QUOTA_DEAD_SEC:.0f}s, coba key berikutnya")
                        break  # keluar dari retry loop, lanjut ke key berikutnya
                    
                    # 401/403: key dicabut/salah -> rotasi ke key berikutnya (tanpa backoff)
                    if e.code in (401, 403):
                        print(f"[matrix]   [X] key #{key_idx+1} DITOLAK (HTTP {e.code}) -> coba key berikutnya")
                        break
                    
                    # 503/429 rate: coba ulang dengan backoff
                    if e.code in (429, 503) and attempt < GEMINI_MAX_RETRY:
                        backoff = GEMINI_BACKOFF * (2 ** attempt)
                        print(f"[matrix]   ⏳ {model} key #{key_idx+1} HTTP {e.code} -> tunggu {backoff:.0f}s, ulang ({attempt+1}/{GEMINI_MAX_RETRY})")
                        time.sleep(backoff)
                        continue
                    
                    # 503 yang TIDAK sembuh setelah retry -> tandai sibuk, coba key berikutnya
                    if e.code == 503 and attempt >= GEMINI_MAX_RETRY:
                        _gemini_busy_matrix.setdefault(model, {})[key_idx] = time.time() + GEMINI_BUSY_DEAD_SEC
                        print(f"[matrix]   [X] key #{key_idx+1} SIBUK 503 ({model}) -> tandai sibuk {GEMINI_BUSY_DEAD_SEC:.0f}s, coba key berikutnya")
                        break
                    
                    # Error lain -> lempar
                    raise
            # end retry loop
        # end key loop
        print(f"[matrix]   SEMUA {len(keys)} key habis untuk {model} -> turun ke model berikutnya", flush=True)
    # end model loop
    
    # Semua model + key habis
    raise last_err or RuntimeError("Semua model + key habis")


# Alias lama untuk kompatibilitas mundur (single key + single model)
_call_gemini_json = _call_llm_json


def _load_ont_cache():
    if os.path.exists(ONTO_CACHE_FILE):
        try:
            with open(ONTO_CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def _save_ont_cache(cache):
    try:
        with open(ONTO_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, indent=2, ensure_ascii=False)
    except OSError as e:
        print(f"Warning: gagal simpan cache ontologi LLM: {e}")


def _validate_llm_ontology(raw, var_name=""):
    """
    Validasi & rapikan hasil JSON dari LLM.
    Menangani dua bentuk keluaran model:
      - objek dict (normal)
      - list berisi beberapa entri (model kadang mengembalikan semua variabel sekaligus)
    """
    # Normalisasi: bila berupa list, pilih entri yang cocok dengan var_name, atau entri pertama
    if isinstance(raw, list):
        if not raw:
            return None
        chosen = None
        if var_name:
            vn = var_name.lower()
            for item in raw:
                if isinstance(item, dict) and vn in str(item.get("variabel", "")).lower():
                    chosen = item
                    break
        if chosen is None:
            chosen = raw[0] if isinstance(raw[0], dict) else None
        raw = chosen
    if not isinstance(raw, dict):
        return None
    dims = raw.get("dimensi")
    if not isinstance(dims, list) or not dims:
        return None
    clean_dims = []
    for d in dims:
        if not isinstance(d, dict):
            continue
        nama = str(d.get("nama", "")).strip()
        inds = d.get("indikator")
        if not nama or not isinstance(inds, list):
            continue
        clean_inds = []    # teks indikator (string) - bentuk lama dipertahankan
        clean_stmts = []   # pernyataan kuesioner paralel ("" bila tak ada)
        for i in inds:
            if isinstance(i, dict):
                teks = str(i.get("t", "")).strip()
                if teks:
                    clean_inds.append(teks)
                    clean_stmts.append(str(i.get("p", "")).strip())
            else:
                teks = str(i).strip()
                if teks:
                    clean_inds.append(teks)
                    clean_stmts.append("")
        if clean_inds:
            clean_dims.append({"nama": nama, "indikator": clean_inds, "pernyataan": clean_stmts})
    if len(clean_dims) < 2:
        return None
    return {
        "grand_theory": str(raw.get("grand_theory", "")).strip(),
        "tokoh_ahli": str(raw.get("tokoh_ahli", "")).strip(),
        "definisi": str(raw.get("definisi", "")).strip(),
        "dimensi": clean_dims,
    }


def _verify_journal(var_name):
    """
    Cari 1 jurnal rujukan ber-DOI NYATA untuk variabel via OpenAlex.
    Fallback: pool jurnal terverifikasi di academic_connector.
    Mengembalikan dict {sitasi, entri} ATAU None.
    """
    try:
        import academic_connector
    except ImportError:
        return None

    # Kueri gabungan agar relevan dengan konteks tesis kinerja pegawai
    for query in (f"pengaruh {var_name} terhadap kinerja pegawai",
                  f"{var_name} kinerja pegawai"):
        papers = academic_connector.search_openalex_verified(query, limit=3)
        if papers:
            paper = papers[0]
            entri, _ = academic_connector.build_apa_entry(paper)
            sitasi = academic_connector.make_in_text_citation(
                paper.get("authors", []), paper.get("year", 2023))
            return {"sitasi": sitasi, "entri": entri}

    # Fallback: pool jurnal lokal terverifikasi
    try:
        topic = academic_connector.detect_variable_topic(var_name)
        pool = academic_connector.TOPIC_JOURNAL_POOLS.get(topic)
        if pool:
            paper = pool[0]
            entri, _ = academic_connector.build_apa_entry(paper)
            sitasi = academic_connector.make_in_text_citation(
                paper.get("authors", []), paper.get("year", 2023))
            return {"sitasi": sitasi, "entri": entri}
    except Exception:
        pass
    return None


def _fill_theory(entry, var_name, role, judul, instansi, api_key, chain, timeout):
    """
    Lengkapi grand_theory / tokoh_ahli / definisi bila belum ada di pustaka.
    Teori bersifat naratif (bukan alat ukur), jadi boleh dari Gemini.
    Dimensi & indikator dari hasil prompt ini SENGAJA DIABAIKAN.
    Hasil disimpan kembali ke pustaka agar tidak diulang untuk judul berikutnya.
    """
    if entry.get("grand_theory") and entry.get("definisi"):
        return entry
    prompt = PROMPT_TEMPLATE.format(
        judul=judul or "(judul tesis)", var_name=var_name, role=role,
        instansi=instansi or "Instansi Pemerintah Daerah")
    # Matriks: habiskan SEMUA key per model.
    keys = gemini_keys(load_config())
    try:
        raw = _call_gemini_matrix(keys, chain, prompt, timeout)
    except Exception as e:
        print(f"[theory] matriks gagal: {type(e).__name__} {str(e)[:60]}")
        return entry
    if isinstance(raw, list):
        raw = raw[0] if raw and isinstance(raw[0], dict) else {}
    if not isinstance(raw, dict):
        return entry
    entry["grand_theory"] = str(raw.get("grand_theory", "")).strip()
    raw_tokoh = str(raw.get("tokoh_ahli", "")).strip()
    # Validasi tokoh ahli melalui OpenAlex (anti-halusinasi)
    try:
        import tokoh_verifier
        valid_tokoh, proofs, dropped = tokoh_verifier.verify_tokoh(raw_tokoh)
        entry["tokoh_ahli"] = valid_tokoh if valid_tokoh else raw_tokoh
        entry["tokoh_proofs"] = proofs
        if dropped:
            print(f"[tokoh_verifier] Tokoh dibuang (tidak terverifikasi OpenAlex): {dropped}")
    except Exception as e:
        print(f"[tokoh_verifier] Warning: {e}")
        entry["tokoh_ahli"] = raw_tokoh
    entry["definisi"] = str(raw.get("definisi", "")).strip()
    if entry["grand_theory"] and entry["definisi"]:
        import variable_library as VL
        VL.save_entry(entry)
    return entry


def research_variable_llm(var_name, role="X", judul="", instansi=""):
    """
    Riset ontologi satu variabel (MODE TERVERIFIKASI).
      1. Pustaka permanen (variable_library) bila sudah ada & layak -> dipakai ulang.
      2. Telusur live OpenAlex 2022-2026 + ekstraksi + verifikasi kutipan (variable_tracer).
      3. Gagal semua -> None (pemanggil menandai needs_review; TIDAK mengarang).
    Mengembalikan dict sestruktur research_variable_ontology() ATAU None.
    """
    if not var_name or not var_name.strip():
        return None

    cfg = load_config()
    # Pakai daftar key; "api_key" tunggal tetap didukung untuk config lama.
    api_key = (gemini_keys(cfg) or [""])[0]
    if not api_key or not cfg.get("enabled", True):
        return None

    import variable_library as VL
    import variable_tracer

    chain = [m for m in cfg.get("models", []) if m] or MODEL_CHAIN
    timeout = max(int(cfg.get("timeout_seconds", ONTO_TIMEOUT)), ONTO_TIMEOUT)

    entry = VL.load_entry(var_name)
    if entry is None or VL.entry_needs_review(entry):
        res = variable_tracer.trace_variable(var_name, role=role, judul=judul, instansi=instansi)
        entry = res.get("entry")
        if not entry:
            print(f"[ontology] '{var_name}' tidak punya sumber terverifikasi: {res.get('report', {}).get('note', '')}")
            return None

    entry = _fill_theory(entry, var_name, role, judul, instansi, api_key, chain, timeout)
    stub = VL.to_ontology_stub(entry, role, var_name)
    if not stub.get("grand_theory"):
        stub["needs_review"] = True
    return stub


def _legacy_free_research_variable_llm(var_name, role="X", judul="", instansi=""):
    """
    [TIDAK DIPAKAI] Versi lama: dimensi/indikator dikarang Gemini dari memori.
    Dipertahankan hanya sebagai referensi; dilarang dipanggil kembali.
    """
    if not var_name or not var_name.strip():
        return None

    cfg = load_config()
    # Pakai daftar key; "api_key" tunggal tetap didukung untuk config lama.
    api_key = (gemini_keys(cfg) or [""])[0]
    if not api_key or not cfg.get("enabled", True):
        return None

    # Cek cache HANYA bila flag use_ontology_cache aktif.
    # Default: OFF (agen AI selalu meriset dari nol agar hasil ter-ground judul).
    cache = {}
    use_cache = cfg.get("use_ontology_cache", False)
    cache_key = var_name.lower().strip()
    if use_cache:
        cache = _load_ont_cache()
        if cache_key in cache:
            cached = dict(cache[cache_key])
            topic = _match_variable_topic(var_name)
            if topic and not _dimensions_fit_topic(cached.get("dimensi", []), topic):
                print(f"[ontology] cache '{var_name}' salah konstruk rumpun '{topic}' -> dihapus, riset ulang")
                cache.pop(cache_key, None)
                _save_ont_cache(cache)
            else:
                cached["role"] = role
                return cached

    chain = [m for m in cfg.get("models", []) if m] or MODEL_CHAIN
    # Ontologi = prompt panjang -> timeout lebih besar dari config parsing judul
    timeout = max(int(cfg.get("timeout_seconds", ONTO_TIMEOUT)), ONTO_TIMEOUT)
    prompt = PROMPT_TEMPLATE.format(
        judul=judul or "(judul tesis)",
        var_name=var_name,
        role=role,
        instansi=instansi or "Instansi Pemerintah Daerah")

    topic = _match_variable_topic(var_name)
    # Matriks: habiskan SEMUA key per model. Validasi topik di dalam loop validasi.
    keys = gemini_keys(load_config())
    try:
        raw = _call_gemini_matrix(keys, chain, prompt, timeout)
    except Exception as e:
        print(f"[ontology] matriks gagal: {type(e).__name__} {str(e)[:60]}")
        return None
    candidate = _validate_llm_ontology(raw, var_name=var_name)
    if not candidate:
        return None
    # Tolak hasil yang salah rumpun (mis. Disiplin -> dimensi motivasi)
    if not _dimensions_fit_topic(candidate["dimensi"], topic):
        print(f"[ontology] hasil matriks salah rumpun untuk '{var_name}' ({topic}) -> gagal")
        return None
    result = candidate
    used_model = "matriks"

    if not result:
        return None

    # Verifikasi jurnal rujukan dengan DOI NYATA (bukan hasil karangan LLM)
    journal = _verify_journal(var_name) or {
        "sitasi": "(Tanpa Rujukan Terverifikasi, 2024)",
        "entri": "Rujukan akan diverifikasi otomatis melalui database OpenAlex ber-DOI.",
    }
    result["jurnal_rujukan"] = journal
    result["nama_resmi"] = var_name
    result["skala_pengukuran"] = "Skala Likert 1–5 (1: Sangat Tidak Setuju s/d 5: Sangat Setuju)"
    result["engine"] = f"gemini:{used_model}"

    # Simpan ke cache (tanpa field role, agar reusable antar-peran) bila flag aktif
    if use_cache:
        to_cache = {k: v for k, v in result.items() if k != "role"}
        cache[cache_key] = to_cache
        _save_ont_cache(cache)

    out = dict(result)
    out["role"] = role
    return out


if __name__ == "__main__":
    # Uji cepat: python llm_ontology_researcher.py
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    for test_var, test_role in [("Kompetensi", "X2"),
                                ("Organizational Citizenship Behavior", "M")]:
        res = research_variable_llm(
            test_var, role=test_role,
            judul="Kecerdasan Emosional dan Kompetensi terhadap Kinerja Pegawai Dinas Pendidikan Kabupaten OKI melalui OCB",
            instansi="Dinas Pendidikan Kabupaten OKI")
        if res:
            print(f"\n=== {test_var} ({test_role}) [{res['engine']}] ===")
            print(f"Teori : {res['grand_theory']}")
            print(f"Ahli  : {res['tokoh_ahli']}")
            for d in res["dimensi"]:
                print(f"  - {d['nama']}: {'; '.join(d['indikator'])}")
            print(f"Jurnal: {res['jurnal_rujukan']['entri'][:120]}...")
        else:
            print(f"\n{test_var}: LLM gagal -> fallback knowledge base")
