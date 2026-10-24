# backend/s3_title_parser.py
"""
S3 DOCTORAL TITLE PARSER & AUTO-DETECTOR
========================================
Modul Pemilah & Penganalisis Semantik Judul Disertasi Doktor (S3).
Mendukung Model Hibrida Tingkat Lanjut:
  - Multi-Variabel Bebas (X1, X2, X3, ... Xn)
  - Variabel Mediasi (M)
  - Variabel Moderasi (Z)
  - Variabel Terikat (Y)
  - Deteksi Otomatis Jenis Model (moderated_mediation, moderation, mediation)
  - Objek & Instansi Penelitian

Arsitektur:
  1. Lapis 1: Gemini LLM dengan prompt khusus doktoral.
  2. Lapis 2: Robust Regex Semantic Fallback bila API offline/kuota habis.
"""

import re
import json
import ssl
import urllib.request
import urllib.error
from typing import Dict, Any, List, Optional

try:
    import llm_title_parser
except ImportError:
    llm_title_parser = None


DOCTORAL_PROMPT_TEMPLATE = """Anda adalah parser semantik ahli untuk judul Disertasi Doktor (S3) Manajemen di Indonesia.

Analisis dan ekstrak komponen struktural dari judul riset doktoral berikut:
"{title}"

Aturan Khusus Doktoral:
1. "model_type":
   - "moderated_mediation" : jika judul mengandung kata mediasi/intervening DAN kata moderasi/dimoderasi/interaksi.
   - "moderation" : jika judul mengandung kata moderasi/dimoderasi/interaksi, TETAPI tidak ada mediasi.
   - "mediation" : jika judul mengandung kata mediasi/intervening/melalui, TETAPI tidak ada moderasi.
   - Default jika tidak ada: "moderated_mediation".
2. "x_vars": array string berisi SEMUA variabel bebas (X1, X2, X3, dst). Bisa 2, 3, 4 atau lebih.
3. "m_vars": array string berisi variabel mediasi / intervening (jika ada, misal setelah kata "melalui", "dimediasi oleh", dsb). Jika tidak ada, isi array kosong [].
4. "z_vars": array string berisi variabel moderasi (jika ada, misal setelah kata "dengan ... sebagai moderasi", "dimoderasi oleh", dsb). Jika tidak ada, isi array kosong [].
5. "y_var": string variabel terikat (dependen), misal setelah kata "terhadap".
6. "instansi": string nama objek penelitian / instansi / organisasi / wilayah lengkap.
7. Bersihkan kata tugas akademik seperti "sebagai variabel intervening", "variabel moderasi", dll dari nama variabel.

HANYA balas dengan JSON valid tanpa penjelasan tambahan:
{{
  "model_type": "moderated_mediation",
  "x_vars": ["Variabel X1", "Variabel X2"],
  "m_vars": ["Variabel Mediasi"],
  "z_vars": ["Variabel Moderasi"],
  "y_var": "Variabel Terikat",
  "instansi": "Nama Instansi"
}}"""


def clean_academic_phrase(text: str) -> str:
    """Membersihkan imbuhan kata tugas dari nama variabel."""
    if not text:
        return ""
    t = text.strip()
    # Hilangkan kata awalan/akhiran umum
    patterns = [
        r"^(pengaruh|analisis|studi tentang|kajian)\s+",
        r"\s+(sebagai\s+variabel\s+(intervening|mediasi|moderasi|independen|dependen)).*$",
        r"\s+(sebagai\s+(mediasi|moderasi|intervening)).*$",
        r"^(dan|serta|,)\s+",
        r"\s+(dan|serta)$"
    ]
    for p in patterns:
        t = re.sub(p, "", t, flags=re.IGNORECASE).strip()
    # Hapus spasi ganda
    t = re.sub(r"\s+", " ", t)
    return t.title() if (t.isupper() or t.islower()) else t


def parse_title_regex_doctoral(title: str) -> Dict[str, Any]:
    """
    Regex fallback rule-based yang tangguh untuk membedah judul disertasi S3.
    """
    raw = (title or "").strip()
    if not raw:
        return {
            "model_type": "moderated_mediation",
            "x_vars": ["Kepemimpinan Transformasional", "Budaya Organisasi"],
            "m_vars": ["Motivasi Kerja"],
            "z_vars": ["Etos Kerja"],
            "y_var": "Kinerja Pegawai",
            "instansi": "Instansi Pemerintah Daerah",
            "engine": "regex-fallback"
        }

    # 1. Deteksi Instansi (pada / di / studi pada / studi empiris di)
    instansi = "Instansi Pemerintah Daerah"
    m_inst = re.search(r"\b(?:pada|di|studi\s+pada|studi\s+kasus\s+di|studi\s+empiris\s+pada)\s+([^\.]+?)(?:\s+tahun\s+\d{4})?$", raw, re.IGNORECASE)
    body_text = raw
    if m_inst:
        instansi = m_inst.group(1).strip()
        body_text = raw[:m_inst.start()].strip()

    # 2. Deteksi Moderasi (dengan ... sebagai variabel moderasi / dimoderasi oleh ...)
    z_vars = []
    m_mod = re.search(r"\b(?:dengan|dimoderasi\s+oleh)\s+([^,]+?)\s+(?:sebagai\s+(?:variabel\s+)?moderasi|sebagai\s+moderator|memoderasi)", body_text, re.IGNORECASE)
    if not m_mod:
        m_mod = re.search(r"\b(?:dengan\s+moderasi|variabel\s+moderasi)\s+([^,]+?)(?:\s+(?:pada|di|melalui)|$)", body_text, re.IGNORECASE)
    
    if m_mod:
        z_name = clean_academic_phrase(m_mod.group(1))
        if z_name:
            z_vars.append(z_name)
        # Hapus bagian moderasi dari body_text agar tidak mencemari variabel lain
        body_text = body_text[:m_mod.start()] + " " + body_text[m_mod.end():]

    # 3. Deteksi Mediasi (melalui ... sebagai variabel intervening/mediasi)
    m_vars = []
    m_med = re.search(r"\b(?:melalui|dimediasi\s+oleh)\s+([^,]+?)(?:\s+sebagai\s+(?:variabel\s+)?(?:intervening|mediasi)|\s+dengan|\s+pada|$)", body_text, re.IGNORECASE)
    if m_med:
        m_name = clean_academic_phrase(m_med.group(1))
        if m_name:
            m_vars.append(m_name)
        body_text = body_text[:m_med.start()] + " " + body_text[m_med.end():]

    # 4. Deteksi Variabel Terikat (Y) (terhadap ...)
    y_var = "Kinerja Pegawai"
    m_y = re.search(r"\b(?:terhadap|dampaknya\s+pada)\s+([^,]+?)(?:\s+melalui|\s+dengan|\s+sebagai|$)", body_text, re.IGNORECASE)
    x_part = body_text
    if m_y:
        y_name = clean_academic_phrase(m_y.group(1))
        if y_name:
            y_var = y_name
        x_part = body_text[:m_y.start()].strip()

    # 5. Deteksi Variabel Bebas (X1, X2, X3, ... Xn) dari x_part
    x_part = re.sub(r"^(?:pengaruh|analisis\s+pengaruh|studi\s+tentang)\s+", "", x_part, flags=re.IGNORECASE).strip()
    
    # Split dengan koma atau kata 'dan' / 'serta'
    # contoh: "A, B, dan C" atau "A dan B"
    raw_x_tokens = re.split(r",\s*|\s+(?:dan|serta)\s+", x_part, flags=re.IGNORECASE)
    x_vars = []
    for tok in raw_x_tokens:
        cleaned = clean_academic_phrase(tok)
        if cleaned and len(cleaned) > 2 and cleaned.lower() not in ("pengaruh", "analisis", "hubungan"):
            x_vars.append(cleaned)

    if not x_vars:
        x_vars = ["Kompetensi Pegawai", "Kepemimpinan Transformasional"]

    # 6. Tentukan Model Type
    has_m = len(m_vars) > 0
    has_z = len(z_vars) > 0

    if has_m and has_z:
        model_type = "moderated_mediation"
    elif has_z and not has_m:
        model_type = "moderation"
    elif has_m and not has_z:
        model_type = "mediation"
    else:
        # Default jika tidak terdeteksi mediasi/moderasi di judul
        model_type = "moderated_mediation"
        m_vars = ["Kepuasan Kerja"]
        z_vars = ["Budaya Organisasi"]

    return {
        "model_type": model_type,
        "x_vars": x_vars,
        "m_vars": m_vars,
        "z_vars": z_vars,
        "y_var": y_var,
        "instansi": instansi.title() if instansi.isupper() else instansi,
        "engine": "regex-fallback"
    }


def parse_title_doctoral(raw_title: str) -> Dict[str, Any]:
    """
    Parser Judul Doktoral Utama: Coba via Gemini LLM terlebih dahulu,
    lalu fallback ke regex parser tangguh.
    """
    raw_title = (raw_title or "").strip()
    if not raw_title:
        return parse_title_regex_doctoral("")

    if llm_title_parser is not None:
        try:
            cfg = llm_title_parser.load_config()
            keys = llm_title_parser.gemini_keys(cfg)
            if keys and keys[0]:
                api_key = keys[0]
                models = cfg.get("models") or [
                    "gemini-3.1-flash-lite",
                    "gemini-flash-lite-latest",
                    "gemini-3.8-flash"
                ]
                prompt = DOCTORAL_PROMPT_TEMPLATE.format(title=raw_title)
                payload = {
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generationConfig": {
                        "responseMimeType": "application/json",
                        "temperature": 0.1,
                        "maxOutputTokens": 600,
                    },
                }
                ctx = ssl.create_default_context()
                for model in models:
                    if str(model).startswith("nvidia/"):
                        continue
                    try:
                        url = f"{llm_title_parser.GEMINI_BASE}/{model}:generateContent?key={api_key}"
                        req = urllib.request.Request(
                            url,
                            data=json.dumps(payload).encode("utf-8"),
                            headers={"Content-Type": "application/json"},
                            method="POST",
                        )
                        with urllib.request.urlopen(req, timeout=10, context=ctx) as resp:
                            res_body = json.loads(resp.read().decode("utf-8"))
                        text = res_body["candidates"][0]["content"]["parts"][0]["text"]
                        parsed_llm = json.loads(text)
                        
                        # Validasi output LLM
                        if parsed_llm.get("x_vars") and parsed_llm.get("y_var"):
                            parsed_llm["engine"] = f"gemini ({model})"
                            # Clean up
                            parsed_llm["x_vars"] = [clean_academic_phrase(x) for x in parsed_llm["x_vars"] if x]
                            parsed_llm["m_vars"] = [clean_academic_phrase(m) for m in parsed_llm.get("m_vars", []) if m]
                            parsed_llm["z_vars"] = [clean_academic_phrase(z) for z in parsed_llm.get("z_vars", []) if z]
                            parsed_llm["y_var"] = clean_academic_phrase(parsed_llm["y_var"])
                            parsed_llm["instansi"] = parsed_llm.get("instansi", "").strip() or "Instansi Terkait"
                            return parsed_llm
                    except Exception as e:
                        # Try next model in chain
                        continue
        except Exception:
            pass

    # Fallback to Regex Parser
    return parse_title_regex_doctoral(raw_title)


if __name__ == "__main__":
    test_title_1 = (
        "PENGARUH TRANSFORMATIONAL LEADERSHIP, IKLIM ORGANISASI, DAN SELF-EFFICACY "
        "TERHADAP KINERJA PEGAWAI MELALUI MOTIVASI KERJA SEBAGAI VARIABEL INTERVENING "
        "DENGAN BUDAYA ORGANISASI SEBAGAI VARIABEL MODERASI PADA PEMERINTAH DAERAH KABUPATEN OGAN KOMERING ILIR"
    )
    print("=== TEST PARSER DOKTORAL 1 (Mediasi Termoderasi 3X) ===")
    res1 = parse_title_regex_doctoral(test_title_1)
    print(json.dumps(res1, indent=2, ensure_ascii=False))

    test_title_2 = (
        "PENGARUH BEBAN KERJA DAN DISIPLIN KERJA TERHADAP PRODUKTIVITAS KERJA "
        "DENGAN KOMPENSASI SEBAGAI VARIABEL MODERASI PADA PT SEMEN BATURAJA"
    )
    print("\n=== TEST PARSER DOKTORAL 2 (Moderasi Murni 2X) ===")
    res2 = parse_title_regex_doctoral(test_title_2)
    print(json.dumps(res2, indent=2, ensure_ascii=False))
