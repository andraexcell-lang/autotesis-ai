# backend/semantic_title_parser.py
"""
SEMANTIC HYBRID TITLE PARSER
Mesin Pemilah Semantik Judul Penelitian Tesis Magister (S2).

Mampu mengurai berbagai ragam variasi judul penelitian kuantitatif model mediasi:
- X1 (Variabel Bebas 1)
- X2 (Variabel Bebas 2)
- M  (Variabel Intervening / Mediasi)
- Y  (Variabel Terikat / Dependen)
- Objek / Instansi Penelitian
- Sektor Kelembagaan (Pemerintahan, Finansial, Kesehatan, Riset & Inovasi, dll.)
"""

import re

def clean_variable_text(text):
    if not text:
        return ""
    s = text.strip()
    # Hapus imbuhan awal
    s = re.sub(r'^(?:variabel|faktor|aspek|dimensi|pengaruh|peran)\s+', '', s, flags=re.IGNORECASE)
    # Hapus imbuhan akhir seperti "sebagai variabel intervening/mediasi"
    s = re.sub(r'\s+sebagai\s+(?:variabel\s+)?(?:intervening|mediasi|mediator)$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s+yang\s+dimediasi\s+oleh$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s+melalui$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s+dengan$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'^(?:dan|serta)\s+', '', s, flags=re.IGNORECASE)
    # Bersihkan spasi ganda
    s = re.sub(r'\s+', ' ', s).strip()
    # Koreksi salah ketik umum pada istilah akademik (dalam frasa)
    typo_fix = {
        "organizatinal": "organizational",
        "organizasional": "organizational",
    }
    low = s.lower()
    for wrong, right in typo_fix.items():
        if wrong in low:
            # Ganti secara case-insensitive pada posisi yang sama
            idx = low.find(wrong)
            s = s[:idx] + right + s[idx + len(wrong):]
            low = s.lower()
    return format_title_case_id(s)

def format_title_case_id(text):
    if not text:
        return ""
    words = text.strip().split()
    lowercase_words = {"dan", "yang", "di", "ke", "dari", "pada", "terhadap", "melalui", "sebagai", "oleh", "untuk", "dalam", "serta", "dengan"}
    known_acronyms = {"OCB", "SPBE", "BPKAD", "BRIDA", "BPPD", "BAPENDA", "DPMPTSP", "PTSP", "DPRD", "RSUD", "PUPR", "DLH", "OPD", "SKPD", "PALI", "OKI", "BPD", "ASN", "PNS", "SLA", "OSS", "PT", "CV", "SD", "SMP", "SMA", "SMK", "IKU", "SPM", "PAD", "WTP", "LKPD", "SPJ"}
    result = []
    for i, w in enumerate(words):
        lw = w.lower()
        clean_w = re.sub(r'[^a-zA-Z0-9]', '', w).upper()
        if clean_w in known_acronyms:
            prefix = "(" if w.startswith("(") else ""
            suffix = ")" if w.endswith(")") else ""
            result.append(f"{prefix}{clean_w}{suffix}")
        elif i > 0 and lw in lowercase_words:
            result.append(lw)
        else:
            result.append(w.capitalize())
    return " ".join(result)

def format_proper_instansi(inst):
    if not inst:
        return "Instansi Pemerintah Daerah"
    # Bersihkan tanda baca tepi
    s = inst.strip(" .,;:-()")
    # Tangani akronim umum agar tetap kapital
    acronyms = {
        "BPKAD", "BRIDA", "BPPD", "BAPENDA", "DPMPTSP", "PTSP", "DPRD", "RSUD",
        "PUPR", "DLH", "OPD", "SKPD", "PALI", "OKI", "BPD", "ASN", "PNS", "SLA",
        "OSS", "PT", "CV", "SD", "SMP", "SMA", "SMK", "PPKBPPPA", "DPPKBPPPA",
        "DP3AP2KB", "DP3A", "BKKBN", "BKPSDM", "BAPPEDA", "DISDIK", "DINKES",
        "DINSOS", "DISNAKER", "DISPORA", "DISHUB", "KOMINFO", "BPSDM", "KPP", "PLN"
    }
    words = s.split()
    res = []
    for w in words:
        clean_w = re.sub(r'[^a-zA-Z0-9]', '', w).upper()
        if clean_w in acronyms:
            # Pertahankan tanda kurung bila ada
            prefix = "(" if w.startswith("(") else ""
            suffix = ")" if w.endswith(")") else ""
            res.append(f"{prefix}{clean_w}{suffix}")
        elif w.lower() in {"dan", "di", "pada", "ke", "dari", "kota", "kabupaten", "provinsi"}:
            if w.lower() in {"kota", "kabupaten", "provinsi"}:
                res.append(w.capitalize())
            else:
                res.append(w.lower())
        else:
            res.append(w.capitalize())
    return " ".join(res)

def _kw_hit(corpus, kws):
    """Cocokkan kata kunci sektor; akronim pendek wajib berbatas kata.

    Bug lama: ``"SD" in "SDM"`` membuat judul ber-``Kompetensi SDM``
    terdeteksi sebagai sektor pendidikan.
    """
    for k in kws:
        kk = (k or "").strip()
        if not kk:
            continue
        if len(kk) <= 4 and kk.isalnum() and kk.upper() == kk:
            if re.search(rf"(?<![A-Za-z0-9]){re.escape(kk)}(?![A-Za-z0-9])", corpus):
                return True
        elif kk in corpus:
            return True
    return False


def detect_institutional_sector(instansi, judul):
    corpus = f"{instansi} {judul}".upper()
    if any(k in corpus for k in ["BPKAD", "KEUANGAN DAN ASET", "PERBENDAHARAAN", "KAS DAERAH", "PENGELOLAAN KEUANGAN"]):
        return "Pengelolaan Keuangan dan Aset Daerah"
    elif any(k in corpus for k in ["BRIDA", "BALITBANGDA", "RISET DAN INOVASI", "LITBANG", "KELITBANGAN"]):
        return "Penelitian, Pengembangan, dan Inovasi Daerah"
    elif any(k in corpus for k in ["BPPD", "BAPENDA", "PENDAPATAN DAERAH", "PAJAK", "RETRIBUSI", "SAMSAT"]):
        return "Pengelolaan Pendapatan Asli Daerah (PAD) dan Perpajakan"
    elif any(k in corpus for k in ["BANK", "PERBANKAN", "BPD", "PARAMESWARA", "KREDIT", "NASABAH", "BPR"]):
        return "Industri Perbankan dan Jasa Finansial Daerah"
    elif any(k in corpus for k in ["DPMPTSP", "PTSP", "PERIZINAN", "PENANAMAN MODAL", "INVESTASI", "OSS"]):
        return "Pelayanan Publik Terpadu dan Penanaman Modal Daerah"
    elif any(k in corpus for k in ["RUMAH SAKIT", "RSUD", "KESEHATAN", "PUSKESMAS", "MEDIS", "KLINIK"]):
        return "Pelayanan Kesehatan dan Keselamatan Pasien"
    elif any(k in corpus for k in ["DPRD", "SEKRETARIAT DEWAN", "SEKWAN"]):
        return "Penyelenggaraan Administrasi dan Fasilitasi Tugas Legislatif Daerah"
    elif _kw_hit(corpus, ["SATPOL", "PAMONG PRAJA", "KETERTIBAN UMUM", "TRANTIBUM",
                          "PENEGAKAN PERDA", "PERATURAN DAERAH"]):
        return "Penyelenggaraan Ketenteraman dan Ketertiban Umum serta Penegakan Peraturan Daerah"
    elif _kw_hit(corpus, ["PENDIDIKAN", "SEKOLAH", "GURU", "SMK", "SMA", "SMP", "SD", "MADRASAH"]):
        return "Penyelenggaraan Pendidikan dan Peningkatan Mutu Pembelajaran"
    elif any(k in corpus for k in ["UNIVERSITAS", "DOSEN", "FAKULTAS", "POLITEKNIK", "INSTITUT", "KAMPUS"]):
        return "Penyelenggaraan Pendidikan Tinggi dan Tri Dharma Perguruan Tinggi"
    elif any(k in corpus for k in ["PERKIMTAN", "PERKIM", "PERUMAHAN", "KAWASAN PERMUKIMAN", "PERTANAHAN"]):
        return "Penyelenggaraan Urusan Perumahan, Kawasan Permukiman, Pertanahan, dan Infrastruktur Publik"
    elif any(k in corpus for k in ["PT ", "CV ", "PERUSAHAAN", "PERSEROAN", "INDUSTRI", "MANUFAKTUR", "RETAIL", "STORE"]):
        return "Dunia Usaha Industri Komersial dan Korporasi"
    else:
        return "Penyelenggaraan Tata Kelola Pemerintahan dan Pelayanan Publik"


def parse_academic_title(raw_title):
    """
    Mengurai judul tesis ke dalam komponen struktural:
    x1, x2, m, y, instansi, sektor.
    Mendukung lebih dari 10 pola gramatikal judul tesis Indonesia baku.
    """
    if not raw_title:
        return {
            "x1": "Motivasi Kerja",
            "x2": "Disiplin Kerja",
            "m": "Employee Engagement",
            "y": "Kinerja Pegawai",
            "instansi": "Instansi Pemerintah Daerah",
            "sektor": "Penyelenggaraan Tata Kelola Pemerintahan dan Pelayanan Publik"
        }

    t = re.sub(r'\s+', ' ', raw_title.strip())

    x1, x2, m, y, instansi = "", "", "", "", ""
    core_title = t

    # 1. Deteksi instansi di tengah judul jika diikuti kata sambung mediasi (MELALUI, SEBAGAI, DENGAN, YANG DIMEDIASI)
    # Contoh: "... Terhadap Kinerja Pegawai Dinas Perkimtan Kota Palembang Dengan Motivasi Kerja Sebagai Variabel Intervening"
    mid_inst = re.search(
        r'(?:\b(?:PADA|DI)\s+)?\b((?:DINAS|BADAN|KANTOR|SEKRETARIAT|INSPEKTORAT|RSUD|RUMAH SAKIT|BPKAD|BRIDA|BPPD|BAPENDA|DPMPTSP|DPRD|PUSKESMAS|BKPSDM|BAPPEDA)\s+.+?)(?=\s+(?:MELALUI|SEBAGAI|DENGAN|YANG\s+DIMEDIASI)\b)',
        core_title, re.IGNORECASE)
    if mid_inst:
        instansi = mid_inst.group(1).strip()
        core_title = (core_title[:mid_inst.start()] + " " + core_title[mid_inst.end():]).strip()
        core_title = re.sub(r'\s+', ' ', core_title)
    else:
        # 2. Pisahkan lokus / instansi jika berada di akhir judul dengan penanda lokasi
        lokus_patterns = [
            r'(?:STUDI KASUS PADA|STUDI EMPIRIS PADA|STUDI PADA)\s+(.+)$',
            r'\b(?:PADA|DI)\s+(?!PENGARUH\b|HUBUNGAN\b|ANALISIS\b|VARIABEL\b)((?:BADAN|DINAS|KANTOR|BPKAD|BRIDA|BPPD|BAPENDA|DPMPTSP|SEKRETARIAT|BANK|RSUD|RUMAH SAKIT|PT|CV|SMA|SMP|SD|UNIVERSITAS|KECAMATAN|KABUPATEN|KOTA|INSPEKTORAT|PUSKESMAS|BKPSDM|BAPPEDA).+)$',
            r'\b(?:PADA|DI)\s+(?!PENGARUH\b|HUBUNGAN\b|ANALISIS\b|VARIABEL\b)([A-Z0-9\s\(\)\.\,\-]+)$',
            r'\b((?:BADAN|DINAS|KANTOR|BPKAD|BRIDA|BPPD|BAPENDA|DPMPTSP|SEKRETARIAT|BANK|RSUD|RUMAH SAKIT|PT|CV|SMA|SMP|SD|UNIVERSITAS|KECAMATAN|KABUPATEN|KOTA|INSPEKTORAT|PUSKESMAS|BKPSDM|BAPPEDA)\s+[A-Z0-9\s\(\)\.\,\-]+)$'
        ]

        for lp in lokus_patterns:
            m_lokus = re.search(lp, core_title, re.IGNORECASE)
            if m_lokus:
                instansi = m_lokus.group(1).strip()
                core_title = core_title[:m_lokus.start()].strip()
                break

    # POLA 0 (BARU): JUDUL TANPA KATA "PENGARUH"/"PERAN"/"ANALISIS" DI AWAL
    # Contoh: "Kecerdasan Emosional Dan Kompetensi Terhadap Kinerja Pegawai Melalui OCB"
    if not re.match(r'^\s*(?:PENGARUH|PERAN|ANALISIS)\b', core_title, re.IGNORECASE):
        p0 = re.search(r'^(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+?)\s+MELALUI\s+(.+)$', core_title, re.IGNORECASE)
        if p0:
            x1, x2, y, m = p0.group(1).strip(), p0.group(2).strip(), p0.group(3).strip(), p0.group(4).strip()
        else:
            p0 = re.search(r'^(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+?)\s+YANG\s+DIMEDIASI\s+OLEH\s+(.+)$', core_title, re.IGNORECASE)
            if p0:
                x1, x2, y, m = p0.group(1).strip(), p0.group(2).strip(), p0.group(3).strip(), p0.group(4).strip()
            else:
                p0 = re.search(r'^(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+?)\s+DENGAN\s+(.+?)\s+SEBAGAI\s+(?:VARIABEL\s+)?(?:INTERVENING|MEDIASI|MEDIATOR)\s*$', core_title, re.IGNORECASE)
                if p0:
                    x1, x2, y, m = p0.group(1).strip(), p0.group(2).strip(), p0.group(3).strip(), p0.group(4).strip()
                else:
                    p0 = re.search(r'^(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+)$', core_title, re.IGNORECASE)
                    if p0:
                        x1, x2, y = p0.group(1).strip(), p0.group(2).strip(), p0.group(3).strip()

    # POLA 1: PERAN M DALAM MEMEDIASI PENGARUH X1 DAN X2 TERHADAP Y (atau PERAN MEDIASI M PADA PENGARUH X1 DAN X2 TERHADAP Y)
    p1 = re.search(r'PERAN\s+(?:MEDIASI\s+)?(.+?)\s+(?:DALAM\s+MEMEDIASI\s+PENGARUH|PADA\s+PENGARUH)\s+(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+)$', core_title, re.IGNORECASE)
    if p1:
        m = p1.group(1).strip()
        x1 = p1.group(2).strip()
        x2 = p1.group(3).strip()
        y = p1.group(4).strip()

    # POLA 2: PENGARUH X1 DAN X2 TERHADAP M SERTA DAMPAKNYA PADA/TERHADAP Y
    elif re.search(r'SERTA\s+DAMPAKNYA\s+(?:PADA|TERHADAP)', core_title, re.IGNORECASE):
        p2 = re.search(r'PENGARUH\s+(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+?)\s+SERTA\s+DAMPAKNYA\s+(?:PADA|TERHADAP)\s+(.+)$', core_title, re.IGNORECASE)
        if p2:
            x1 = p2.group(1).strip()
            x2 = p2.group(2).strip()
            m = p2.group(3).strip()
            y = p2.group(4).strip()

    # POLA 3: PENGARUH X1 DAN X2 TERHADAP Y MELALUI M
    elif re.search(r'TERHADAP\s+(.+?)\s+MELALUI\s+(.+)$', core_title, re.IGNORECASE):
        p3 = re.search(r'PENGARUH\s+(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+?)\s+MELALUI\s+(.+)$', core_title, re.IGNORECASE)
        if p3:
            x1 = p3.group(1).strip()
            x2 = p3.group(2).strip()
            y = p3.group(3).strip()
            m = p3.group(4).strip()

    # POLA 4: PENGARUH X1 DAN X2 TERHADAP Y YANG DIMEDIASI OLEH M
    elif re.search(r'TERHADAP\s+(.+?)\s+YANG\s+DIMEDIASI\s+OLEH\s+(.+)$', core_title, re.IGNORECASE):
        p4 = re.search(r'PENGARUH\s+(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+?)\s+YANG\s+DIMEDIASI\s+OLEH\s+(.+)$', core_title, re.IGNORECASE)
        if p4:
            x1 = p4.group(1).strip()
            x2 = p4.group(2).strip()
            y = p4.group(3).strip()
            m = p4.group(4).strip()

    # POLA 5: PENGARUH X1 DAN X2 TERHADAP Y DENGAN M SEBAGAI VARIABEL INTERVENING/MEDIASI
    elif re.search(r'DENGAN\s+(.+?)\s+SEBAGAI\s+(?:VARIABEL\s+)?(?:INTERVENING|MEDIASI|MEDIATOR)', core_title, re.IGNORECASE):
        p5 = re.search(r'PENGARUH\s+(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+?)\s+DENGAN\s+(.+?)\s+SEBAGAI', core_title, re.IGNORECASE)
        if p5:
            x1 = p5.group(1).strip()
            x2 = p5.group(2).strip()
            y = p5.group(3).strip()
            m = p5.group(4).strip()

    # POLA 6: ANALISIS MEDIASI M PADA PENGARUH X1 DAN X2 TERHADAP Y
    elif re.search(r'ANALISIS\s+MEDIASI\s+(.+?)\s+PADA\s+PENGARUH\s+(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+)$', core_title, re.IGNORECASE):
        p6 = re.search(r'ANALISIS\s+MEDIASI\s+(.+?)\s+PADA\s+PENGARUH\s+(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+)$', core_title, re.IGNORECASE)
        if p6:
            m = p6.group(1).strip()
            x1 = p6.group(2).strip()
            x2 = p6.group(3).strip()
            y = p6.group(4).strip()

    # POLA FALLBACK UMUM: PENGARUH X1, X2 TERHADAP Y
    elif re.search(r'PENGARUH\s+(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+)$', core_title, re.IGNORECASE):
        p_fb = re.search(r'PENGARUH\s+(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+)$', core_title, re.IGNORECASE)
        if p_fb:
            x1 = p_fb.group(1).strip()
            x2 = p_fb.group(2).strip()
            y = p_fb.group(3).strip()
            m = "Employee Engagement"

    # Jika variabel masih belum terisi, terapkan default yang aman
    x1_clean = clean_variable_text(x1) if x1 else "Motivasi Kerja"
    x2_clean = clean_variable_text(x2) if x2 else "Disiplin Kerja"
    m_clean = clean_variable_text(m) if m else "Kepuasan Kerja"
    y_clean = clean_variable_text(y) if y else "Kinerja Pegawai"
    instansi_clean = format_proper_instansi(instansi) if instansi else "Instansi Pemerintah Daerah"

    sektor = detect_institutional_sector(instansi_clean, raw_title)

    return {
        "x1": x1_clean,
        "x2": x2_clean,
        "m": m_clean,
        "y": y_clean,
        "instansi": instansi_clean,
        "sektor": sektor,
        "raw_title": raw_title
    }

if __name__ == "__main__":
    test_titles = [
        "PENGARUH ORGANIZATIONAL CITIZENSHIP BEHAVIOR DAN MOTIVASI KERJA TERHADAP KINERJA PEGAWAI MELALUI KEPUASAN KERJA SEBAGAI VARIABEL INTERVENING PADA BPKAD KABUPATEN OKI",
        "PENGARUH PROMOSI JABATAN DAN MUTASI KERJA TERHADAP KINERJA PEGAWAI MELALUI MOTIVASI KERJA SEBAGAI VARIABEL INTERVENING PADA BRIDA KABUPATEN OKI",
        "PENGARUH PENGEMBANGAN KARIER DAN BUDAYA ORGANISASI TERHADAP KINERJA PEGAWAI MELALUI KEPUASAN KERJA SEBAGAI VARIABEL INTERVENING PADA BPPD KABUPATEN OKI",
        "PERAN EMPLOYEE ENGAGEMENT DALAM MEMEDIASI PENGARUH MOTIVASI KERJA DAN DISIPLIN KERJA TERHADAP KINERJA PEGAWAI DI DPMPTSP PROVINSI SUMATERA SELATAN",
        "PENGARUH KECERDASAN EMOSIONAL DAN MOTIVASI KERJA TERHADAP ETOS KERJA SERTA DAMPAKNYA PADA KINERJA PEGAWAI DI BPKAD KABUPATEN PALI"
    ]
    for idx, t in enumerate(test_titles, 1):
        parsed = parse_academic_title(t)
        print(f"\n[{idx}] Judul: {t[:60]}...")
        print(f"     X1: {parsed['x1']} | X2: {parsed['x2']} | M: {parsed['m']} | Y: {parsed['y']}")
        print(f"     Objek: {parsed['instansi']} | Sektor: {parsed['sektor']}")
