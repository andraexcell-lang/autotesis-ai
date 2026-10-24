# backend/references_metodologi.py
"""
Pool referensi metodologi TERVERIFIKASI untuk BAB III (3.2 Populasi & Sampel,
3.3 Variabel & Definisi Operasional, 3.4 Instrumen & Skala Pengukuran,
3.5 Teknik Analisis Data, dan 3.6 Pengujian Hipotesis).

Tujuan (permintaan pengguna):
  - Sitasi HARUS benar-benar ada (buku/jurnal nyata, terverifikasi online/offline).
  - Hanya buku/edisi terbaru (tidak ada buku 2019 ke bawah).

Daftar nyata tersimpan di `references_metodologi.json`. Modul ini menyediakan:
  - load_refs()           : baca seluruh entri terverifikasi
  - citation_menu()       : daftar sitasi untuk dimasukkan ke prompt LLM
  - extract_citations()   : ambil sitasi '(Penulis, Tahun)' dari teks
  - refs_used_in()        : entri pool yang benar-benar dikutip di teks
  - validate_texts()      : pastikan TIDAK ada sitasi di luar pool (strict)
"""

import json
import os
import re

_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "references_metodologi.json")

_YEAR_RE = re.compile(r"((?:19|20)\d{2})")
_CITE_PAREN_RE = re.compile(r"\(([^()]*?(?:19|20)\d{2}[a-z]?)\)")
_NAME_RE = re.compile(r"([A-Z][A-Za-z\u00C0-\u024F'\-]+)")

# Kelompok penulis gaya naratif: "Roring & Walangitan", "A dan B", "A and B".
# Sengaja TIDAK memakai koma sebagai pemisah agar tidak salah menyerap kata
# bergaris besar sebelumnya (mis. "Misalnya, Sugiyono" -> hanya "Sugiyono").
_AUTH_TOKEN = r"[A-Z][A-Za-z\u00C0-\u024F'\-]+"
_AUTH_SEP = r"\s*(?:&|\bdan\b|\band\b)\s*"
_GROUP_END_RE = re.compile(rf"({_AUTH_TOKEN}(?:{_AUTH_SEP}{_AUTH_TOKEN})*)\s*$")
_NARR_GROUP_RE = re.compile(rf"({_AUTH_TOKEN}(?:{_AUTH_SEP}{_AUTH_TOKEN})*)\s*\(((?:19|20)\d{{2}})[a-z]?\)")
_NAME_STOP = {
    "misalnya", "menurut", "oleh", "dalam", "pada", "dan", "and", "serta",
    "seperti", "hasil", "penelitian", "kajian", "studi", "teori", "konsep",
    "adapun", "berdasarkan", "sebagaimana", "sedangkan", "namun",
}


def _first_author_of(group):
    """Nama penulis PERTAMA dari sekelompok nama (mengabaikan kata pengantar)."""
    names = _NAME_RE.findall(group or "")
    while names and names[0].lower() in _NAME_STOP:
        names = names[1:]
    return names[0] if names else None


def load_refs():
    try:
        with open(_PATH, encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"[references_metodologi] gagal memuat pool: {e}")
        return []
    return [r for r in (data.get("sumber") or []) if r.get("sitasi") and r.get("apa")]


def _key(cit):
    """Kunci identitas sitasi: (nama-keluarga-pertama huruf kecil, tahun)."""
    if not cit:
        return None
    ym = _YEAR_RE.search(cit)
    if not ym:
        return None
    names = _NAME_RE.findall(cit)
    if not names:
        return None
    return (names[0].lower(), ym.group(1))


def extract_citations(text):
    """Ambil semua sitasi '(Penulis, Tahun)' + gaya naratif 'Penulis (Tahun)'."""
    if not isinstance(text, str):
        return []
    out = []
    for m in _CITE_PAREN_RE.finditer(text):
        inner = m.group(1)
        parts = [x.strip() for x in inner.split(";") if x.strip()]
        if re.search(r"[A-Za-z]{2,}", inner):
            out.extend(parts)
        else:
            # gaya naratif: "Menurut Sugiyono (2022)" / "Roring & Walangitan (2023)"
            ym = _YEAR_RE.search(inner)
            if not ym:
                continue
            seg = text[: m.start()]
            mg = _GROUP_END_RE.search(seg)
            if mg:
                first = _first_author_of(mg.group(1))
                if first:
                    out.append(f"{first}, {ym.group(1)}")
                    continue
            # fallback (mis. "Smith et al."): ambil nama kapital terakhir
            names = _NAME_RE.findall(seg)
            if names:
                out.append(f"{names[-1]}, {ym.group(1)}")
    return out


def refs_used_in(texts, refs=None):
    """Entri pool yang benar-benar dikutip pada teks (terurut tahun terbaru dulu)."""
    refs = refs if refs is not None else load_refs()
    keys = set()
    for t in (texts or []):
        for cit in extract_citations(t):
            k = _key(cit)
            if k:
                keys.add(k)
    used = [r for r in refs if _key(r.get("sitasi")) in keys]
    used.sort(key=lambda r: -int(r.get("tahun") or 0))
    return used


def validate_texts(texts, refs=None):
    """Pastikan setiap sitasi berada di dalam pool. -> (ok: bool, bad: list[str])."""
    refs = refs if refs is not None else load_refs()
    allowed = {_key(r.get("sitasi")) for r in refs}
    allowed.discard(None)
    bad = []
    for t in (texts or []):
        for cit in extract_citations(t):
            if _key(cit) not in allowed:
                bad.append(cit)
    return (len(bad) == 0, bad)


def _intext_names(sitasi):
    """Ambil rangkaian nama penulis (tanpa tahun) dari string sitasi pool."""
    s = (sitasi or "").strip()
    ym = _YEAR_RE.search(s)
    if not ym:
        return ""
    pre = s[: ym.start()].strip().rstrip("(,").strip()
    pre = pre.strip("()").strip().rstrip(",").strip()
    return pre


def surname_alias_index(refs):
    """Indeks alias (nama-keluarga-bukan-pertama, tahun) -> (key kanonik, nama in-text).

    Untuk memulihkan sitasi yang kehilangan penulis pertama, mis.
    '(Izzah, 2022)' -> '(Nahdliyah & Izzah, 2022)'. Alias yang ambigu (menunjuk
    lebih dari satu entri pool) sengaja DIBUANG agar tidak salah perbaikan.
    """
    alias = {}
    ambiguous = set()
    for r in refs or []:
        sit = r.get("sitasi")
        names = _NAME_RE.findall(sit or "")
        ym = _YEAR_RE.search(sit or "")
        if len(names) < 2 or not ym:
            continue
        year = ym.group(1)
        ckey = (names[0].lower(), year)
        intext = _intext_names(sit) or " & ".join(names)
        for nm in names[1:]:
            k = (nm.lower(), year)
            if k == ckey:
                continue
            if k in alias and alias[k][0] != ckey:
                ambiguous.add(k)
            else:
                alias[k] = (ckey, intext, year)
    for k in ambiguous:
        alias.pop(k, None)
    return alias


def repair_citations(text, refs):
    """Perbaiki sitasi yang menjatuhkan penulis pertama BILA alias unik di pool.

    Menangani dua bentuk: '(Izzah, 2022)' dan gaya naratif 'Izzah (2022)'.
    Sitasi yang sudah valid / tidak dikenali TIDAK diubah (biar validator tetap
    menolaknya). Mengembalikan (teks_baru, daftar_perbaikan).
    """
    if not isinstance(text, str) or not text:
        return text, []
    refs = refs or []
    allowed = {_key(r.get("sitasi")) for r in refs}
    allowed.discard(None)
    alias = surname_alias_index(refs)
    if not alias:
        return text, []
    repairs = []

    # 1) Sitasi di dalam tanda kurung.
    out, last = [], 0
    for m in _CITE_PAREN_RE.finditer(text):
        out.append(text[last:m.start()])
        parts = [x.strip() for x in m.group(1).split(";") if x.strip()]
        new_parts, changed = [], False
        for part in parts:
            k = _key(part)
            if k in allowed or k not in alias:
                new_parts.append(part)
                continue
            _ckey, intext, year = alias[k]
            new_parts.append(f"{intext}, {year}")
            repairs.append((part, new_parts[-1]))
            changed = True
        out.append("(" + "; ".join(new_parts) + ")" if changed else m.group(0))
        last = m.end()
    out.append(text[last:])
    new_text = "".join(out)

    # 2) Sitasi gaya naratif "Izzah (2022)".
    def _repl(m):
        group, year = m.group(1), m.group(2)
        first = _first_author_of(group)
        if not first:
            return m.group(0)
        k = (first.lower(), year)
        if k in allowed or k not in alias:
            return m.group(0)
        _ckey, intext, _yy = alias[k]
        new = f"{intext} ({year})"
        repairs.append((f"{group} ({year})", new))
        return new
    new_text = _NARR_GROUP_RE.sub(_repl, new_text)

    return new_text, repairs


def citation_menu(refs=None):
    """Blok teks daftar sitasi (tahun terbaru dulu) untuk dimasukkan ke prompt."""
    refs = refs if refs is not None else load_refs()
    refs = sorted(refs, key=lambda r: -int(r.get("tahun") or 0))
    lines = []
    for r in refs:
        topik = ", ".join(r.get("topik") or [])
        lines.append(f'- {r["sitasi"]}  |  {r["apa"]}  |  cocok: {topik}')
    return "\n".join(lines)


if __name__ == "__main__":
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    refs = load_refs()
    print(f"total refs: {len(refs)}")
    demo = "(Sugiyono, 2022). Menurut Hair et al. (2022) sampel minimal. (Penulis Palsu, 2021)."
    ok, bad = validate_texts([demo])
    print("valid:", ok, "| bad:", bad)
    print("used:", [r["id"] for r in refs_used_in([demo])])
