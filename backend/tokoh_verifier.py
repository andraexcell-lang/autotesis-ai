# backend/tokoh_verifier.py
"""
VALIDASI TOKOH AHLI / PENELITI AKADEMIK (ANTI-HALUSINASI).

Memverifikasi nama-nama tokoh/ahli yang dihasilkan LLM terhadap repositori OpenAlex Authors.
Strategi pencocokan (diperketat agar "Robbins" tidak salah ke ahli saraf):
  1. Cari author di OpenAlex dengan NAMA LENGKAP (depan + belakang).
  2. Cocokkan: nama belakang author SESUAI, dan minimal satu kata nama depan
     (selain kata sambung) juga ada di nama author tersebut.
  3. Bila Gemini hanya memberi nama belakang (tanpa depan), pencocokan hanya
     memakai nama belakang -> hasilnya TANDAI sebagai "lemah" (bukan bukti kuat).
Nama yang terbukti akademisi (works_count >= 1) dipertahankan; nama fiktif dibuang.
"""

import json
import re
import urllib.parse
import urllib.request

UA = "AutoTesisTokohVerifier/1.0 (mailto:tesismagister@research.ac.id)"

# Kata sambung umum pada nama (bukan penanda nama depan)
_STOP = {"dan", "bin", "binti", "al", "ibnu", "putra", "putri", "van", "de", "di"}


def parse_entries(tokoh_text):
    """
    Ekstrak (nama_lengkap, tahun, label_asli) dari teks:
    'Anwar Prabu Mangkunegara (2021); Stephen P. Robbins & Timothy A. Judge (2019)'
    """
    if not tokoh_text:
        return []
    out = []
    chunks = re.split(r"[;\n]+", str(tokoh_text))
    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk:
            continue
        m = re.search(r"^(.+?)\s*\(\s*((?:19|20)\d{2})\s*\)", chunk)
        if m:
            out.append((m.group(1).strip(), int(m.group(2)), chunk))
        else:
            out.append((chunk, 0, chunk))
    return out


def _split_names(name_field):
    """
    Pecah satu field nama berisi beberapa tokoh ('A. B. Robbins & T. A. Judge')
    menjadi daftar nama kandidat.
    """
    parts = re.split(r"\s*&\s*|\s+dan\s+", str(name_field))
    return [p.strip() for p in parts if p.strip()]


def _significant_tokens(full_name):
    """Kata-kata nama selain kata sambung dan inisial satu huruf."""
    toks = re.findall(r"[A-Za-z\u00C0-\u024F'\-]+", str(full_name))
    return [t for t in toks if len(t) > 1 and t.lower() not in _STOP]


def _author_matches(cand_tokens, author_display):
    """
    True bila nama belakang cocok dan minimal satu kata nama depan ikut cocok.
    """
    if not cand_tokens:
        return False
    author_tokens = [t.lower() for t in _significant_tokens(author_display)]
    cand_lower = [t.lower() for t in cand_tokens]
    last = cand_lower[-1]
    if last not in author_tokens:
        return False
    if len(cand_lower) == 1:
        return True          # hanya nama belakang -> cocok tapi lemah
    firsts = cand_lower[:-1]
    return any(f in author_tokens for f in firsts)


def _check_author_openalex(candidates, timeout=8):
    """
    candidates: daftar nama kandidat dari satu entri ('Stephen P. Robbins', 'Timothy A. Judge').
    Cari author yang cocok di OpenAlex untuk salah satu kandidat.
    Mengembalikan (info_author, kuat?) atau (None, False).
    """
    best_weak = None
    for name in candidates:
        tokens = _significant_tokens(name)
        if not tokens:
            continue
        # Cari dengan nama lengkap (lebih spesifik), fallback ke nama belakang
        query = " ".join(tokens[-2:]) if len(tokens) >= 2 else tokens[-1]
        enc = urllib.parse.quote(query)
        url = f"https://api.openalex.org/authors?search={enc}&per-page=5&mailto=tesismagister@research.ac.id"
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except Exception:
            continue
        for r in (data.get("results") or []):
            disp = r.get("display_name", "")
            works = r.get("works_count", 0)
            if works < 1 or not disp:
                continue
            if _author_matches(tokens, disp):
                info = {
                    "queried_name": name,
                    "display_name": disp,
                    "works_count": works,
                    "institution": (r.get("last_known_institutions") or [{}])[0].get("display_name", "")
                                    if r.get("last_known_institutions") else "",
                    "openalex_id": r.get("id"),
                    "strong": len(tokens) >= 2,
                }
                if info["strong"]:
                    return info, True
                if best_weak is None:
                    best_weak = info
    if best_weak:
        return best_weak, False
    return None, False


def verify_tokoh(tokoh_text):
    """
    Memverifikasi daftar tokoh ahli.
    Mengembalikan: (tokoh_terverifikasi_str, daftar_bukti, dibuang)
    """
    kept = []
    proofs = []
    dropped = []

    for raw_name, tahun, label in parse_entries(tokoh_text):
        candidates = _split_names(raw_name)
        info, strong = _check_author_openalex(candidates)
        if info:
            kept.append(label)
            proofs.append({
                "label": label,
                "verified_scholar": info["display_name"],
                "works_count": info["works_count"],
                "institution": info["institution"],
                "openalex_id": info["openalex_id"],
                "match_strength": "kuat" if strong else "lemah (hanya nama belakang)",
            })
        else:
            dropped.append(label)

    return "; ".join(kept), proofs, dropped


if __name__ == "__main__":
    tests = [
        "Anwar Prabu Mangkunegara (2021); Stephen P. Robbins & Timothy A. Judge (2019); Sedarmayanti (2020)",
        "TokohFiktifAsalAsalan (2024); Zylokrapov999 (2022)",
    ]
    for t in tests:
        print("=" * 70)
        print("UJI:", t)
        valid, proofs, dropped = verify_tokoh(t)
        print("Lolos :", valid)
        for p in proofs:
            print(f"  - {p['label']} -> '{p['verified_scholar']}' ({p['works_count']} karya, {p['match_strength']})")
        print("Dibuang:", dropped)
