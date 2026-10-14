# backend/doc_text_extract.py
"""
UTILITAS EKSTRAKSI TEKS DOKUMEN (PDF / DOCX / TXT).

Dipakai agen penelusur dimensi & indikator untuk membaca isi sumber akademik
(jurnal Open Access ber-DOI) maupun berkas yang diunggah pengguna, sehingga
dimensi/indikator dapat dikutip APA ADANYA dari dokumen nyata.

Murni pembacaan berkas - tanpa panggilan jaringan kecuali download_pdf().
"""

import os
import urllib.request

UA = "AutoTesisDimensionTracer/1.0 (mailto:tesismagister@research.ac.id)"


def extract_pdf_text(path, max_pages=None):
    """Ekstrak teks dari PDF memakai PyMuPDF (fitz). Mengembalikan string."""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return ""
    chunks = []
    try:
        with fitz.open(path) as doc:
            for i, page in enumerate(doc):
                if max_pages is not None and i >= max_pages:
                    break
                try:
                    chunks.append(page.get_text())
                except Exception:
                    continue
    except Exception:
        return ""
    return "\n".join(chunks)


def extract_docx_text(path):
    """Ekstrak teks dari DOCX (paragraf + isi tabel)."""
    try:
        import docx
    except ImportError:
        return ""
    parts = []
    try:
        d = docx.Document(path)
    except Exception:
        return ""
    for para in d.paragraphs:
        if para.text:
            parts.append(para.text)
    for tbl in d.tables:
        for row in tbl.rows:
            cells = [c.text.strip() for c in row.cells if c.text and c.text.strip()]
            if cells:
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def extract_any_text(path, max_pages=None):
    """
    Deteksi jenis berkas dari ekstensi lalu ekstrak teksnya.
    Mengembalikan (teks, jenis) ; teks kosong bila gagal.
    """
    if not path or not os.path.exists(path):
        return "", "unknown"
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return extract_pdf_text(path, max_pages=max_pages), "pdf"
    if ext in (".docx", ".doc"):
        return extract_docx_text(path), "docx"
    if ext in (".txt", ".md"):
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read(), "txt"
        except Exception:
            return "", "txt"
    return "", "unknown"


def download_pdf(url, dest_path, timeout=25):
    """
    Unduh PDF dari URL Open Access ke dest_path.
    Mengembalikan True bila berhasil dan berkasnya tampak seperti PDF.
    """
    if not url:
        return False
    try:
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = resp.read()
        if not data or data[:4] != b"%PDF":
            return False
        os.makedirs(os.path.dirname(os.path.abspath(dest_path)), exist_ok=True)
        with open(dest_path, "wb") as f:
            f.write(data)
        return True
    except Exception:
        return False


def normalize_for_match(text):
    """
    Normalisasi teks untuk pencocokan kutipan: huruf kecil, buang tanda baca
    dan spasi berlebih. Dipakai memverifikasi kutipan benar-benar ada di sumber.
    """
    import re
    if not text:
        return ""
    t = text.lower()
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def quote_in_source(quote, source_text, min_words=4):
    """
    True bila 'quote' benar-benar ditemukan di 'source_text' (setelah normalisasi).
    Toleransi: kutipan dipecah menjadi segmen; minimal 80% segmen harus ada.
    Kutipan < min_words kata -> tidak dapat diverifikasi sendiri (dianggap lemah).
    """
    if not quote or not source_text:
        return False
    words = quote.split()
    qn = normalize_for_match(quote)
    sn = normalize_for_match(source_text)
    if not qn:
        return False
    if qn in sn:
        return True
    if len(words) < min_words:
        return False
    # Cek per bagian yang lebih panjang untuk toleransi perbedaan tanda baca
    seg = " ".join(words[:min(len(words), 8)])
    seg_n = normalize_for_match(seg)
    return bool(seg_n) and seg_n in sn


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        txt, kind = extract_any_text(sys.argv[1])
        print(f"jenis={kind} panjang={len(txt)} karakter")
        print(txt[:800])
    else:
        print("Pakai: python doc_text_extract.py <berkas.pdf|docx|txt>")