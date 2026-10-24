# backend/indicator_layers.py
"""
Dua lapisan indikator + seleksi dimensi otomatis yang transparan.

Kebijakan:
  - Dimensi sumber boleh 2-20. Sistem memilih 2-7 dimensi yang paling relevan
    dan mencatat dimensi yang tidak dipakai beserta alasannya.
  - Bab II memakai ``indikator_teoretis``: generik, tidak memuat nama/objek/
    sektor penelitian.
  - Bab III/IV/Lampiran memakai ``indikator_operasional`` (alias ``indikator``):
    spesifik pada dimensi dan konteks objek penelitian.

Modul ini sengaja tidak memanggil LLM. Pemetaan teoretis dilakukan secara
deterministik dari indikator operasional agar dua lapisan selalu paralel dan
tidak menambah risiko halusinasi.
"""

import re

import objek_sterilisasi as OS

MAX_DIMENSI_FINAL = 7
MIN_DIMENSI_FINAL = 2
MAX_DIMENSI_SUMBER = 20

# Panjang maksimal indikator teoretis (Bab II) dalam kata.
MAX_KATA_INDIKATOR_TEORETIS = 4

# Istilah yang menunjukkan konteks objek/sektor. Lapisan teoretis Bab II tidak
# boleh memuatnya; lapisan operasional justru membutuhkannya.
_OBJEK_TERLARANG = re.compile(
    r"\b(?:PT|Persero|Tbk|Cabang|Palembang|Bhanda|Ghara|Reksa|"
    r"perusahaan|gudang|pergudangan|logistik|kargo|mitra)\b", re.I)

# Frasa panjang diganti lebih dulu agar hasilnya tidak menggantung.
_GANTI_FRASA = [
    (r"\bBhanda\s+Ghara\s+Reksa\b", ""),
    (r"\bPT\s+[A-Z][\w&.\-]*(?:\s+[A-Z][\w&.\-]*){0,7}\b", ""),
    (r"\b(?:Persero|Tbk|Cabang|Palembang)\b", ""),
    (r"\bdi\s+(?:area\s+)?(?:gudang|pergudangan|logistik)\b", ""),
    (r"\bdi\s+lingkungan\s+perusahaan\b", ""),
    (r"\bdi\s+perusahaan\b", ""),
    (r"\bpada\s+perusahaan\b", ""),
    (r"\bmilik\s+mitra(?:\s+bisnis)?(?:\s+perusahaan)?\b", ""),
    (r"\bmitra\s+bisnis\b", ""),
    (r"\bvolume\s+tugas\s+penanganan\s+logistik\s+gudang\b", "volume tugas"),
    (r"\btugas\s+bongkar\s+muat\s+barang\b", "tugas"),
    (r"\baktivitas\s+bongkar\s+muat\b", "aktivitas kerja"),
    (r"\bbongkar\s+muat\b", ""),
    (r"\bpelacakan\s+pengiriman(?:\s+\w+)?\b", "pelacakan pekerjaan"),
    (r"\bdokumen\s+administrasi\s+pengiriman\b", "dokumen administrasi"),
    (r"\bpengiriman\s+dokumen\b", "penyusunan dokumen"),
    (r"\blonjakan\s+volume\s+pengiriman(?:\s+\w+)?\b", "lonjakan volume pekerjaan"),
    (r"\bvolume\s+pengiriman(?:\s+\w+)?\b", "volume pekerjaan"),
    (r"\bpekerjaan\s+barang\b", "pekerjaan"),
    (r"\bdi\s+sesuai\b", "sesuai"),
    (r"\bpenanganan\s+komplain\b", "penanganan keluhan"),
    (r"\baset\s+barang\b", "aset organisasi"),
    (r"\blayanan\s+gudang\b", "layanan"),
    (r"\boperasional\s+gudang\b", "operasional"),
    (r"\boperasional\s+logistik\b", "operasional"),
    (r"\bpekerjaan\s+logistik\b", "pekerjaan"),
    (r"\btugas\s+logistik\b", "tugas"),
    (r"\bmanajemen\s+logistik\b", "manajemen"),
    (r"\bmanajemen\s+pergudangan\b", "manajemen"),
    (r"\bkeselamatan\s+kerja\s+logistik\b", "keselamatan kerja"),
    (r"\bpenataan\s+barang\s+gudang\b", "penataan barang"),
    (r"\bpenanganan\s+logistik\s+gudang\b", "penanganan pekerjaan"),
    (r"\bpenanganan\s+logistik\b", "penanganan pekerjaan"),
    (r"\bmenangani\s+operasional\b", "menangani pekerjaan"),
    (r"\bkargo\b", "barang"),
]

_KATA_UJUNG = re.compile(
    r"\s+\b(?:di|pada|dalam|untuk|dengan|dari|terhadap|bagi|oleh|sebagai)\b\s*$",
    re.I)
_PREPOSISI_MENGGANTUNG = re.compile(
    r"\s+\b(?:di|pada|dalam|untuk|dengan|dari|terhadap|bagi|oleh)\b\s*(?=[,.;]|$)",
    re.I)


def _kapital(teks):
    t = str(teks or "").strip()
    for i, ch in enumerate(t):
        if ch.isalpha():
            return t[:i] + ch.upper() + t[i + 1:]
    return t


def _rapikan(teks):
    t = re.sub(r"\s+", " ", str(teks or "")).strip()
    t = re.sub(r"\s+([,.;:])", r"\1", t)
    t = t.strip(" ,;:-")
    t = _KATA_UJUNG.sub("", t)
    t = _PREPOSISI_MENGGANTUNG.sub("", t)
    t = re.sub(r"\s{2,}", " ", t).strip(" ,;:-")
    return t


def generikkan_indikator(teks, instansi=""):
    """Buang penanda objek/sektor dari satu indikator operasional.

    Contoh:
      "Penguasaan regulasi standar prosedur operasional logistik gudang
       perusahaan" -> "Penguasaan regulasi standar prosedur operasional"
    """
    t = str(teks or "").strip()
    if not t:
        return ""
    for pola, baru in _GANTI_FRASA:
        t = re.sub(pola, baru, t, flags=re.I)
    t = OS.bersihkan_objek(t, instansi)
    return _kapital(_rapikan(t))


def _fallback_teoretis(nama_dimensi, geser=0, instansi=""):
    dim = generikkan_indikator(nama_dimensi, instansi) or "aspek kerja"
    dim = re.sub(r"^(?:dimensi|aspek)\s+", "", dim, flags=re.I).strip()
    kandidat = [
        f"Tingkat {dim} yang ditunjukkan dalam bekerja",
        f"Kesesuaian {dim} dengan tuntutan tugas",
        f"Konsistensi {dim} pada pelaksanaan pekerjaan",
        f"Kemampuan menerapkan {dim} dalam tugas",
        f"Penguasaan {dim} dalam pelaksanaan tugas",
    ]
    hasil = kandidat[int(geser) % len(kandidat)]
    # Ringkas menjadi frasa pendek (mis. "Tingkat Kuantitas Hasil").
    pendek = OS.pendekkan_indikator(hasil, instansi)
    return pendek or hasil


def _norm(teks):
    return re.sub(r"[^\w\s]", " ", str(teks or "").lower()).split()


def bangun_indikator_teoretis(operasional, nama_dimensi="", instansi="", judul=""):
    """Bangun 2-4 indikator teoretis generik dari indikator operasional.

    Urutan dipertahankan sebisa mungkin. Bila hasil generik kembar, dipakai
    fallback berbasis nama dimensi agar jumlah tetap 2-4 dan tidak ada duplikat.
    """
    ops = [str(x).strip() for x in (operasional or []) if str(x).strip()]
    target = max(MIN_DIMENSI_FINAL, min(4, len(ops))) if ops else 0
    if not target:
        return []

    hasil, dipakai = [], set()
    for i, op in enumerate(ops):
        # Lapisan Bab II = frasa pendek 1-4 kata, generik, tanpa objek.
        g = OS.pendekkan_indikator(op, instansi)
        if (not g or OS.mengandung_objek(g, instansi)
                or " ".join(_norm(g)) in dipakai):
            g = _fallback_teoretis(nama_dimensi, i, instansi)
        if " ".join(_norm(g)) in dipakai:
            g = ""
        if g:
            hasil.append(_kapital(g))
            dipakai.add(" ".join(_norm(g)))

    geser = 0
    while len(hasil) < target and geser < 12:
        g = _fallback_teoretis(nama_dimensi, geser, instansi)
        if " ".join(_norm(g)) not in dipakai:
            hasil.append(_kapital(g))
            dipakai.add(" ".join(_norm(g)))
        geser += 1
    return hasil[:4]


def indikator_teoretis(d, instansi="", judul=""):
    """Ambil lapisan teoretis; bila belum ada, turunkan dari operasional.

    Nilai tersimpan (cache blueprint) tetap diringkas ulang agar selalu
    berupa frasa pendek 1-4 kata tanpa istilah objek/sektor.
    """
    if not isinstance(d, dict):
        return []
    teo = [str(x).strip() for x in (d.get("indikator_teoretis") or []) if str(x).strip()]
    _derajat = getattr(OS, "_DERAJAT_AWAL", {"minimalnya", "tingginya", "rendahnya", "banyaknya", "kurangnya", "adanya"})
    _prep_ujung = getattr(OS, "_PREP_UJUNG", set())
    _ada_cacat = any((len(x.split()) == 1 and (x.lower() in _derajat or x.lower().endswith("nya"))) or
                     (x.split() and x.split()[-1].lower() in _prep_ujung) for x in teo)
    if teo and not OS.mengandung_objek(" ".join(teo), instansi) and not _ada_cacat:
        pendek, dipakai = [], set()
        for x in teo:
            g = OS.pendekkan_indikator(x, instansi)
            k = " ".join(_norm(g)) if g else ""
            if g and k not in dipakai:
                pendek.append(g)
                dipakai.add(k)
        if pendek:
            return pendek[:4]
    return bangun_indikator_teoretis(
        d.get("indikator_operasional") or d.get("indikator") or [],
        d.get("nama", ""), instansi=instansi, judul=judul)


def indikator_operasional(d):
    if not isinstance(d, dict):
        return []
    return list(d.get("indikator_operasional") or d.get("indikator") or [])


# ---------------------------------------------------------------------------
# SELEKSI DIMENSI 2-7
# ---------------------------------------------------------------------------
_STOP = {
    "dan", "atau", "di", "pada", "dalam", "untuk", "dengan", "dari", "ke",
    "yang", "sebagai", "terhadap", "oleh", "bagi", "serta", "the", "of", "and",
    "a", "an", "to", "in", "on", "for", "with", "by", "as", "pt", "persero",
    "tbk", "cabang", "perusahaan", "karyawan", "pegawai", "kerja", "pekerjaan",
}
_CANON = {
    "karyawan": "pegawai", "employee": "pegawai", "staff": "pegawai",
    "training": "pelatihan", "career": "karir", "karier": "karir",
    "performance": "kinerja", "competency": "kompetensi",
    "competence": "kompetensi", "development": "pengembangan",
}


def _tokens(teks):
    s = re.sub(r"[^a-zA-Z0-9]+", " ", str(teks or "")).lower()
    out = set()
    for t in s.split():
        t = _CANON.get(t, t)
        if len(t) > 2 and t not in _STOP:
            out.add(t)
    return out


def _jaccard(a, b):
    if not a or not b:
        return 0.0
    return len(a & b) / float(len(a | b))


def _alasan_pilih(nama, overlap, idx, n):
    if overlap >= 0.25:
        return ("Dipilih karena istilah dimensi paling dekat dengan variabel/"
                "definisi dan konteks penelitian.")
    if idx == 0:
        return "Dipilih karena merupakan dimensi inti yang disebut paling awal dalam sumber."
    return ("Dipilih karena berada di urutan awal sumber dan memberi cakupan "
            "teoretis yang berbeda.")


def select_dimensions(dims, var_name="", definisi="", instansi="", judul="",
                      max_dims=MAX_DIMENSI_FINAL):
    """Pilih 2-7 dimensi dari seluruh dimensi sumber dan catat yang dibuang.

    Mengembalikan ``(dimensi_dipilih, metadata)``. Bila jumlah sumber <= batas,
    semua dipakai. Bila > batas, skor menggabungkan kedekatan istilah dengan
    variabel/definisi, posisi awal dalam sumber, dan perbedaan cakupan.
    """
    dims = [d for d in (dims or []) if isinstance(d, dict) and str(d.get("nama", "")).strip()]
    n = len(dims)
    if n <= max_dims:
        dipilih = []
        for i, d in enumerate(dims):
            alasan = ("Dipakai seluruhnya: jumlah dimensi sumber berada dalam "
                      "batas 2-7." if n >= MIN_DIMENSI_FINAL else
                      "Dipakai karena merupakan satu-satunya dimensi sumber.")
            d["urutan_sumber"] = i + 1
            d["skor_pemilihan"] = 1.0
            d["alasan_pemilihan"] = alasan
            dipilih.append({"nama": d["nama"], "urutan_sumber": i + 1,
                            "skor": 1.0, "alasan": alasan})
        return dims, {
            "kebijakan": "2-7 dimensi dipilih otomatis dari seluruh dimensi sumber",
            "sumber_total": n,
            "dipilih_total": n,
            "batas_maksimum": max_dims,
            "dipilih": dipilih,
            "tidak_dipakai": [],
        }

    query = _tokens(" ".join([var_name or "", definisi or "", instansi or "", judul or ""]))
    kandidat = []
    for i, d in enumerate(dims):
        nama = str(d.get("nama", "")).strip()
        isi = " ".join([nama] + [str(x) for x in (d.get("indikator") or [])[:2]])
        tok = _tokens(isi)
        overlap = (len(tok & query) / float(len(tok))) if tok else 0.0
        sentral = 1.0 - 0.45 * (i / float(max(1, n - 1)))
        skor = 0.55 * sentral + 0.45 * overlap
        if re.match(r"^(?:lain|lainnya|umum|general|other|misc)\b", nama, re.I):
            skor -= 0.60
        kandidat.append({"idx": i, "d": d, "nama": nama, "tok": tok,
                         "overlap": overlap, "skor": skor, "skor_dasar": skor})

    # Greedy: setiap langkah memilih skor efektif tertinggi setelah penalti
    # tumpang-tindih dihitung terhadap dimensi yang SUDAH dipilih.
    sisa = list(kandidat)
    terpilih = []
    while sisa and len(terpilih) < max_dims:
        for k in sisa:
            tumpang = max((_jaccard(k["tok"], s["tok"]) for s in terpilih),
                          default=0.0)
            k["tumpang"] = tumpang
            k["skor"] = k["skor_dasar"] - (0.35 if tumpang >= 0.85 else 0.0)
        terbaik = max(sisa, key=lambda x: (x["skor"], -x["idx"]))
        terpilih.append(terbaik)
        sisa.remove(terbaik)

    terpilih.sort(key=lambda x: x["idx"])
    dipakai_idx = {k["idx"] for k in terpilih}
    if len(terpilih) < MIN_DIMENSI_FINAL:
        for k in kandidat:
            if k["idx"] not in dipakai_idx:
                terpilih.append(k)
                dipakai_idx.add(k["idx"])
            if len(terpilih) >= MIN_DIMENSI_FINAL:
                break

    terpilih.sort(key=lambda x: x["idx"])
    dipilih, tidak = [], []
    for k in terpilih:
        d = k["d"]
        alasan = _alasan_pilih(k["nama"], k["overlap"], k["idx"], n)
        d["urutan_sumber"] = k["idx"] + 1
        d["skor_pemilihan"] = round(float(k["skor"]), 3)
        d["alasan_pemilihan"] = alasan
        dipilih.append({"nama": k["nama"], "urutan_sumber": k["idx"] + 1,
                        "skor": round(float(k["skor"]), 3), "alasan": alasan})
    for k in kandidat:
        if k["idx"] in dipakai_idx:
            continue
        if k.get("tumpang", 0.0) >= 0.85:
            alasan = ("Tidak dipakai karena tumpang-tindih makna dengan dimensi "
                      "yang sudah dipilih.")
        elif k["overlap"] < 0.10:
            alasan = ("Tidak dipakai karena keterkaitannya dengan variabel/"
                      "definisi penelitian lebih rendah.")
        else:
            alasan = ("Tidak dipakai karena skor relevansi lebih rendah dan "
                      "batas maksimum 7 dimensi.")
        tidak.append({"nama": k["nama"], "urutan_sumber": k["idx"] + 1,
                      "skor": round(float(k["skor"]), 3), "alasan": alasan})

    return [k["d"] for k in terpilih], {
        "kebijakan": "2-7 dimensi dipilih otomatis dari seluruh dimensi sumber",
        "sumber_total": n,
        "dipilih_total": len(terpilih),
        "batas_maksimum": max_dims,
        "dipilih": dipilih,
        "tidak_dipakai": tidak,
    }


def pastikan_layer(data, judul="", instansi="", max_dims=MAX_DIMENSI_FINAL):
    """Lengkapi lapisan teoretis/operasional + seleksi dimensi pada blueprint.

    Tidak memanggil LLM dan tidak mengubah indikator operasional. Mengembalikan
    ``True`` bila ada perubahan yang perlu disimpan.
    """
    if not isinstance(data, dict):
        return False
    variables = data.get("variables") or {}
    if not variables:
        return False
    berubah = False
    total_teo = 0
    for v in variables.values():
        if not isinstance(v, dict):
            continue
        dims = [d for d in (v.get("dimensi") or []) if isinstance(d, dict)]
        if len(dims) > max_dims:
            dims, seleksi = select_dimensions(
                dims, var_name=v.get("nama_resmi") or v.get("input_name") or "",
                definisi=v.get("definisi", ""), instansi=instansi, judul=judul,
                max_dims=max_dims)
            v["dimensi"] = dims
            v["dimension_selection"] = seleksi
            berubah = True
        elif not v.get("dimension_selection"):
            dims, seleksi = select_dimensions(
                dims, var_name=v.get("nama_resmi") or v.get("input_name") or "",
                definisi=v.get("definisi", ""), instansi=instansi, judul=judul,
                max_dims=max_dims)
            v["dimensi"] = dims
            v["dimension_selection"] = seleksi
            berubah = True

        for d in v.get("dimensi") or []:
            op = [str(x).strip() for x in
                  (d.get("indikator_operasional") or d.get("indikator") or [])
                  if str(x).strip()]
            if op and d.get("indikator_operasional") != op:
                d["indikator_operasional"] = op
                berubah = True
            if op and d.get("indikator") != op:
                d["indikator"] = op
                berubah = True
            teo = [str(x).strip() for x in (d.get("indikator_teoretis") or [])
                   if str(x).strip()]
            _derajat = getattr(OS, "_DERAJAT_AWAL", {"minimalnya", "tingginya", "rendahnya", "banyaknya", "kurangnya", "adanya"})
            _prep_ujung = getattr(OS, "_PREP_UJUNG", set())
            _ada_cacat = any((len(x.split()) == 1 and (x.lower() in _derajat or x.lower().endswith("nya"))) or
                             (x.split() and x.split()[-1].lower() in _prep_ujung) for x in teo)
            if (len(teo) < MIN_DIMENSI_FINAL
                    or OS.mengandung_objek(" ".join(teo), instansi)
                    or _ada_cacat
                    or any(len(x.split()) > MAX_KATA_INDIKATOR_TEORETIS
                           for x in teo)):
                teo = bangun_indikator_teoretis(
                    op, d.get("nama", ""), instansi=instansi, judul=judul)
                if teo != (d.get("indikator_teoretis") or []):
                    d["indikator_teoretis"] = teo
                    berubah = True
            total_teo += len(d.get("indikator_teoretis") or [])
        v["indicator_policy"] = "generated_v1"

    meta = data.setdefault("metadata", {})
    if meta.get("total_indikator_teoretis") != total_teo:
        meta["total_indikator_teoretis"] = total_teo
        berubah = True
    if meta.get("indicator_layers") != {
            "teoretis": "Bab II (generik, tanpa nama/objek/sektor penelitian)",
            "operasional": "Bab III/IV/Lampiran (spesifik pada konteks objek)"}:
        meta["indicator_layers"] = {
            "teoretis": "Bab II (generik, tanpa nama/objek/sektor penelitian)",
            "operasional": "Bab III/IV/Lampiran (spesifik pada konteks objek)"}
        berubah = True
    if not meta.get("dimension_selection_policy"):
        meta["dimension_selection_policy"] = (
            "Dimensi sumber 2-20 diterima; 2-7 dipilih otomatis dan dimensi "
            "yang tidak dipakai dicatat beserta alasannya.")
        berubah = True
    return berubah
