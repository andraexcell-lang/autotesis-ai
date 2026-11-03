# backend/build_full_thesis_docx.py
"""
Script Pembangun Naskah Tesis Lengkap dan Dokumen Per-Bab:
ANDRE JULIANSYA (NPM: 255441003) - Universitas Tridinanti Palembang 2026.
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

import build_andre_thesis

OUTPUT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

import re
from docx.enum.section import WD_SECTION_START

# Angka Romawi HARUS tetap kapital saat kapitalisasi teks. Tanpa ini,
# str.title()/str.capitalize() mengubah "II" menjadi "Ii", mis.
# "Kantor Kesyahbandaran dan Otoritas Pelabuhan Kelas II Palembang"
# -> "... Kelas Ii Palembang".
ROMAN = {"I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
         "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII",
         "XIX", "XX", "XXI", "XXII", "XXIII", "XXIV"}


_CAPTION_RE = re.compile(r"^(Tabel|Gambar)\s+(\d+)\.\s*(\d+)", re.I)


def _rapikan_caption(text):
    """Seragamkan label tabel/gambar dan baris sumber.

    Caption resmi menjadi ``Tabel 3.1`` + baris baru judul. Kalimat naratif
    seperti ``Tabel 1.1 memperlihatkan ...`` tidak diubah karena kata
    setelah nomor berhuruf kecil.
    """
    m = _CAPTION_RE.match(text)
    if m:
        rest = text[m.end():]
        first = rest.lstrip()[:1]
        is_caption = rest.startswith(".") or "\n" in rest or (first.isalpha() and first.isupper())
        if is_caption:
            body = rest.lstrip(" .\r\n\t")
            text = f"{m.group(1)} {m.group(2)}.{m.group(3)}" + (f"\n{body}" if body else "")
    if text.startswith("Sumber"):
        text = re.sub(r"^Sumber\s*:\s*", "Sumber: ", text)
        text = re.sub(r",\s*\((20\d{2})\)", r" (\1)", text)
    return text


def _judul_rapi(judul):
    """Kapitalisasi judul tanpa merusak angka Romawi.

    ``str.title()`` mengubah "Kelas II" menjadi "Kelas Ii". Fungsi ini
    menulis Title Case per kata, tetapi angka Romawi (II, III, IV, ...)
    dibiarkan kapital seluruhnya.
    """
    if not judul:
        return judul
    hasil = []
    for w in str(judul).split():
        core = re.sub(r"[^A-Za-z0-9]", "", w).upper()
        if core in ROMAN:
            hasil.append(w.upper())
        else:
            hasil.append(w.title())
    return " ".join(hasil)


def format_proper_instansi(name):
    """
    Mengubah nama instansi yang ALL CAPS menjadi format baku EYD (Title Case)
    dengan tetap mempertahankan akronim dan singkatan resmi huruf kapital.
    Contoh: 'SEKRETARIAT DPRD KOTA PALEMBANG' -> 'Sekretariat DPRD Kota Palembang'
            'BPKAD KABUPATEN PALI' -> 'BPKAD Kabupaten PALI'
    """
    if not name:
        return "Instansi Pemerintah Daerah"
    UPPER_ACRONYMS = {
        "DPRD", "DPR", "BPKAD", "DPMPTSP", "RSUD", "PALI", "BAPPEDA", "BKPSDM", 
        "BAPENDA", "DINSOS", "DINKES", "DISDIK", "SDM", "ASN", "PNS", "TNI", 
        "POLRI", "BUMN", "BUMD", "SPM", "IKU", "SAKIP", "LAKIP", "OPD", "UPTD", 
        "BLUD", "PT", "CV", "FEB", "MM", "SPJ", "LPJ", "BOSP", "IGD", "RME", 
        "BPJS", "SNP", "SOP", "SUMSEL", "BPSDM", "KPP", "PLN", "UPDK", "BPS",
        "PPKBPPPA", "DPPKBPPPA", "DP3AP2KB", "DP3A", "BKKBN", "DISNAKER", "DISPORA", "DISHUB", "KOMINFO"
    }
    LOWER_WORDS = {"dan", "di", "ke", "dari", "pada", "untuk", "terhadap", "dengan", "yang", "sebagai"}

    words = name.strip().split()
    formatted = []
    for i, w in enumerate(words):
        w_clean = re.sub(r'[^a-zA-Z0-9]', '', w).upper()
        if w_clean in ROMAN:
            prefix = re.match(r'^[^a-zA-Z0-9]+', w)
            suffix = re.search(r'[^a-zA-Z0-9]+$', w)
            pre_str = prefix.group(0) if prefix else ""
            suf_str = suffix.group(0) if suffix else ""
            formatted.append(f"{pre_str}{w_clean}{suf_str}")
        elif w_clean in UPPER_ACRONYMS:
            prefix = re.match(r'^[^a-zA-Z0-9]+', w)
            suffix = re.search(r'[^a-zA-Z0-9]+$', w)
            pre_str = prefix.group(0) if prefix else ""
            suf_str = suffix.group(0) if suffix else ""
            formatted.append(f"{pre_str}{w_clean}{suf_str}")
        elif i > 0 and w.lower() in LOWER_WORDS:
            formatted.append(w.lower())
        else:
            formatted.append(w.capitalize())
    return " ".join(formatted)

# ==============================================================================
# KAMUS GLOBAL DAN ENGINE PEMIRING OTOMATIS ISTILAH ASING (EYD V & STANDAR TESIS)
# ==============================================================================

_BASE_AUTO_ITALIC_TERMS = [
    # 1. Rujukan Latin & Sitasi Akademik
    r"et al\.",
    r"et al",
    r"ibid\.",
    r"ibid",
    r"op\.cit\.",
    r"op\. cit\.",
    r"loc\.cit\.",
    r"loc\. cit\.",
    r"ceteris paribus",
    r"vice versa",
    r"status quo",
    r"ad hoc",
    r"de facto",
    r"de jure",
    r"per se",
    r"a priori",
    r"a posteriori",
    r"inter alia",

    # 2. Metodologi, Analisis Statistik & PLS-SEM
    r"partial least squares structural equation modeling",
    r"partial least squares - structural equation modeling",
    r"structural equation modeling",
    r"partial least squares",
    r"covariance-based structural equation modeling",
    r"pls-sem",
    r"sem-pls",
    r"cb-sem",
    r"smartpls 4",
    r"smartpls 3",
    r"smartpls",
    r"warp-pls",
    r"warppls",
    r"outer model",
    r"inner model",
    r"measurement model",
    r"structural model",
    r"outer loading",
    r"outer loadings",
    r"cross loading",
    r"cross loadings",
    r"indicator reliability",
    r"internal consistency reliability",
    r"composite reliability",
    r"average variance extracted",
    r"convergent validity",
    r"discriminant validity",
    r"fornell-larcker criterion",
    r"fornell-larcker",
    r"heterotrait-monotrait ratio",
    r"htmt ratio",
    r"collinearity statistics",
    r"variance inflation factor",
    r"path coefficient",
    r"path coefficients",
    r"coefficient of determination",
    r"effect size",
    r"f-square",
    r"q-square",
    r"r-square",
    r"r-squared",
    r"f-squared",
    r"q-squared",
    r"adjusted r-square",
    r"predictive relevance",
    r"blindfolding",
    r"bootstrapping",
    r"goodness of fit",
    r"standardized root mean square residual",
    r"two-tailed test",
    r"one-tailed test",
    r"two-tailed",
    r"one-tailed",
    r"p-value",
    r"p-values",
    r"t-statistic",
    r"t-statistics",
    r"t-count",
    r"t-table",
    r"f-count",
    r"f-table",
    r"purposive sampling",
    r"saturated sampling",
    r"total sampling",
    r"census sampling",
    r"snowball sampling",
    r"stratified random sampling",
    r"simple random sampling",
    r"cluster sampling",
    r"accidental sampling",
    r"convenience sampling",
    r"non-probability sampling",
    r"probability sampling",
    r"pilot study",
    r"pre-test",
    r"post-test",
    r"common method bias",
    r"harman's single factor test",
    r"confirmatory factor analysis",
    r"exploratory factor analysis",
    r"cronbach's alpha",
    r"cronbach alpha",
    r"item-total correlation",
    r"multicollinearity",
    r"homoscedasticity",
    r"heteroscedasticity",

    # 3. Teori Manajemen & Konsep Riset
    r"social exchange theory",
    r"two-factor theory",
    r"human capital theory",
    r"job demands-resources model",
    r"job demands-resources",
    r"conservation of resources theory",
    r"conservation of resources",
    r"self-determination theory",
    r"leader-member exchange",
    r"transformational leadership theory",
    r"transactional leadership theory",
    r"servant leadership theory",
    r"path-goal theory",
    r"contingency theory",
    r"agency theory",
    r"stewardship theory",
    r"resource-based view",
    r"institutional theory",
    r"grand theory",
    r"middle range theory",
    r"applied theory",
    r"evidence-based policy",
    r"evidence-based practice",
    r"evidence-based management",
    r"research gap",
    r"performance gap",
    r"phenomenon gap",
    r"theoretical gap",
    r"intervening",
    r"moderating",
    r"moderation",
    r"pseudo-compliance",
    r"quiet quitting",
    r"discretionary effort",
    r"person-organization fit",
    r"person-job fit",
    r"person-environment fit",
    r"psychological safety",
    r"psychological contract",
    r"psychological empowerment",

    # 4. Variabel, Dimensi & Indikator (SDM, Perilaku Organisasi & Manajemen)
    r"employee engagement",
    r"work-life balance",
    r"work life balance",
    r"work-family conflict",
    r"work family conflict",
    r"job satisfaction",
    r"turnover intention",
    r"turnover intentions",
    r"organizational citizenship behavior",
    r"self-efficacy",
    r"self efficacy",
    r"perceived organizational support",
    r"job performance",
    r"employee performance",
    r"task performance",
    r"contextual performance",
    r"burnout",
    r"counterproductive work behavior",
    r"work engagement",
    r"job crafting",
    r"job stress",
    r"work stress",
    r"workplace well-being",
    r"job burnout",
    r"emotional exhaustion",
    r"depersonalization",
    r"reduced personal accomplishment",
    r"vigor",
    r"dedication",
    r"absorption",
    r"idealized influence",
    r"inspirational motivation",
    r"intellectual stimulation",
    r"individual consideration",
    r"contingent reward",
    r"management by exception active",
    r"management by exception passive",
    r"management by exception",
    r"laissez-faire",
    r"servant leadership",
    r"digital leadership",
    r"authentic leadership",
    r"transformational leadership",
    r"transactional leadership",
    r"ethical leadership",
    r"participative leadership",
    r"directive leadership",
    r"affective commitment",
    r"normative commitment",
    r"continuance commitment",
    r"job embeddedness",
    r"organizational commitment",
    r"organizational culture",
    r"organizational climate",
    r"emotional knowledge",
    r"emotional expression",
    r"emotional regulation",
    r"altruism",
    r"sportsmanship",
    r"courtesy",
    r"civic virtue",
    r"conscientiousness",
    r"tangibles",
    r"reliability",
    r"responsiveness",
    r"assurance",
    r"empathy",
    r"goal attainment",
    r"career development",
    r"work environment",
    r"work motivation",
    r"work discipline",
    r"emotional intelligence",
    r"workload",
    r"work-life conflict",
    r"work-life enrichment",

    # 5. Layanan Medis & Manajemen Rumah Sakit
    r"clinical pathway",
    r"emergency response time",
    r"patient safety",
    r"bed occupancy rate",
    r"average length of stay",
    r"hospital acquired infections",
    r"nursing performance",
    r"quality of care",
    r"patient satisfaction",

    # 6. Istilah Korporat, Bisnis & Tata Kelola
    r"key performance indicators",
    r"key performance indicator",
    r"service level agreement",
    r"standard operating procedure",
    r"human resource development",
    r"talent management",
    r"human capital management",
    r"human capital",
    r"on the job training",
    r"off the job training",
    r"on-the-job training",
    r"off-the-job training",
    r"whistleblowing system",
    r"good corporate governance",
    r"corporate social responsibility",
    r"focus group discussion",
    r"desk research",
    r"in-depth interview"
]

_DYNAMIC_AUTO_ITALIC_TERMS = set()
_COMPILED_AUTO_ITALIC_REGEX = None


def _build_auto_italic_regex():
    """Mengompilasi regex global untuk deteksi markdown & istilah asing."""
    all_terms = list(_BASE_AUTO_ITALIC_TERMS) + list(_DYNAMIC_AUTO_ITALIC_TERMS)
    seen = set()
    unique_terms = []
    for t in all_terms:
        low = t.lower()
        if low not in seen:
            seen.add(low)
            unique_terms.append(t)
    unique_terms.sort(key=len, reverse=True)

    parts = []
    for t in unique_terms:
        esc = t if ("\\" in t) else re.escape(t)
        if esc.endswith(r"\."):
            parts.append(r"\b" + esc)
        else:
            parts.append(r"\b" + esc + r"\b")

    pattern_str = (
        r"(\*\*\*([^*]+)\*\*\*)"
        r"|(\*\*([^*]+)\*\*)"
        r"|(\*([^*]+)\*)"
        r"|" + "|".join(f"({p})" for p in parts)
    )
    return re.compile(pattern_str, re.IGNORECASE)


def _get_auto_italic_regex():
    global _COMPILED_AUTO_ITALIC_REGEX
    if _COMPILED_AUTO_ITALIC_REGEX is None:
        _COMPILED_AUTO_ITALIC_REGEX = _build_auto_italic_regex()
    return _COMPILED_AUTO_ITALIC_REGEX


def register_auto_italic_terms(terms):
    """Mendaftarkan istilah bahasa asing baru secara dinamis."""
    global _DYNAMIC_AUTO_ITALIC_TERMS, _COMPILED_AUTO_ITALIC_REGEX
    if isinstance(terms, str):
        terms = [terms]
    changed = False
    for t in (terms or []):
        t_clean = str(t).strip()
        if t_clean and t_clean.lower() not in _DYNAMIC_AUTO_ITALIC_TERMS:
            _DYNAMIC_AUTO_ITALIC_TERMS.add(t_clean)
            changed = True
    if changed:
        _COMPILED_AUTO_ITALIC_REGEX = _build_auto_italic_regex()


def add_rich_runs_to_para(para, text, bold_default=False, italic_default=False, font_size=12):
    """
    Menambahkan run ke paragraf dengan mem-parse teks berhuruf miring (italic).
    Menghilangkan tanda asterisk (*) mentah atau tanda kurung (*...*)
    dan memastikan istilah asing dicetak miring (italic = True) di Microsoft Word
    sesuai aturan baku EYD Edisi V dan Pedoman Penulisan Tesis Pascasarjana.
    """
    if text is None:
        return
    text_str = str(text)
    if not text_str:
        return

    # Bersihkan bentuk kurung berasterisk seperti (*kata*) atau (*kata) atau (kata*)
    normalized = re.sub(r"\(\*+([^*]+)\*+\)", r"(\1)", text_str)
    normalized = re.sub(r"\(\*+([^*]+)\)", r"(\1)", normalized)
    normalized = re.sub(r"\(([^*]+)\*+\)", r"(\1)", normalized)

    token_pattern = _get_auto_italic_regex()

    last_idx = 0
    for match in token_pattern.finditer(normalized):
        start, end = match.span()
        # Teks sebelum match
        if start > last_idx:
            pre_text = normalized[last_idx:start]
            r_pre = para.add_run(pre_text)
            r_pre.bold = bold_default
            r_pre.italic = italic_default
            r_pre.font.name = "Times New Roman"
            r_pre.font.size = Pt(font_size)
            r_pre.font.color.rgb = RGBColor(0, 0, 0)

        # Evaluasi jenis match
        if match.group(1):  # ***bold italic***
            matched_str = match.group(2)
            run_bold = True
            run_italic = True
        elif match.group(3):  # **bold**
            matched_str = match.group(4)
            run_bold = True
            run_italic = italic_default
        elif match.group(5):  # *italic*
            matched_str = match.group(6)
            run_bold = bold_default
            run_italic = True
        else:  # Istilah asing yang cocok otomatis
            matched_str = match.group(0)
            run_bold = bold_default
            run_italic = True

        r_match = para.add_run(matched_str)
        r_match.bold = run_bold
        r_match.italic = run_italic
        r_match.font.name = "Times New Roman"
        r_match.font.size = Pt(font_size)
        r_match.font.color.rgb = RGBColor(0, 0, 0)

        last_idx = end

    # Sisa teks setelah match terakhir
    if last_idx < len(normalized):
        post_text = normalized[last_idx:]
        r_post = para.add_run(post_text)
        r_post.bold = bold_default
        r_post.italic = italic_default
        r_post.font.name = "Times New Roman"
        r_post.font.size = Pt(font_size)
        r_post.font.color.rgb = RGBColor(0, 0, 0)

def setup_page(doc, continuous=False):
    for s in doc.sections:
        s.page_width = Cm(21.0)      # A4
        s.page_height = Cm(29.7)
        s.top_margin = Cm(4.0)       # 4 cm
        s.left_margin = Cm(4.0)      # 4 cm
        s.bottom_margin = Cm(3.0)    # 3 cm
        s.right_margin = Cm(3.0)     # 3 cm
        if continuous:
            s.start_type = WD_SECTION_START.CONTINUOUS

HANG_IN = 0.5   # posisi teks setelah nomor/huruf: sejajar dengan awal teks isi (0,5")

# Pola awalan penomoran/penghurufan pada judul sub-bab maupun butir daftar.
# Sengaja ketat agar paragraf isi (mis. "1.000 responden", "2.5 persen") tidak ikut tertangkap.
_RX_PREFIX_CAP = re.compile(
    r"^("
    r"\d+(?:\.\d+)*\."            # 1.    1.1.    2.1.3.
    r"|\d+(?:\.\d+){2,}"          # 3.2.1  (judul tanpa titik penutup)
    r"|\d+\)"                     # 1)
    r"|[a-e]\)\."                 # a).    b).
    r"|[a-e][.)]"                 # a.    a)
    r"|[A-Z]\."                   # A.
    r"|Lampiran\s+\d+"            # Lampiran 1
    r")(\s+)(.*)$", re.S)

# Nomor bertingkat 4 ke atas (mis. 2.1.1.1.). Nomor ini lebarnya sudah mencapai
# 0,5" sehingga tab stop 0,5" terlampaui dan Word melompat ke tab berikutnya (1,0").
_RX_NOMOR_PANJANG = re.compile(r"^\d+\.\d+\.\d+\.\d+")


def _hang_number(para, indent_in=HANG_IN):
    """
    Indentasi menggantung (hanging indent) untuk SEMUA paragraf bernomor/berhuruf:
    - nomor/huruf menempel di tepi margin kiri,
    - judul/teksnya mulai pada posisi 0,5" (sejajar dengan awal teks isi),
    - baris lanjutan sejajar dengan baris pertama, bukan dengan nomor.
    """
    text = (para.text or "").strip()
    if not text or re.search(r"\.{5,}", text):
        return para                 # bukan paragraf bernomor atau berisi titik-titik isian
    if not _RX_PREFIX_CAP.match(text):
        return para

    runs = list(para.runs)
    if runs:
        m = _RX_PREFIX_CAP.match(runs[0].text)
        if not m:
            return para
        # Pemisah nomor/huruf:
        # - nomor 4 tingkat ke atas (2.1.1.1., dst.) memakai SATU SPASI, karena
        #   lebarnya sudah melewati tab stop 0,5" sehingga TAB akan melompat ke
        #   1,0" dan jarak nomor-judul jadi tidak konsisten antarsub-bab;
        # - nomor pendek memakai TAB agar judul tepat mulai pada posisi indentasi.
        nomor = m.group(1)
        pemisah = " " if _RX_NOMOR_PANJANG.match(nomor) else "\t"
        runs[0].text = nomor + pemisah + m.group(3)
        if not m.group(3) and len(runs) > 1:
            runs[1].text = runs[1].text.lstrip()

    pf = para.paragraph_format
    pf.left_indent = Inches(indent_in)
    pf.first_line_indent = Inches(-indent_in)
    try:
        pf.tab_stops.clear_all()
    except Exception:
        pass
    pf.tab_stops.add_tab_stop(Inches(indent_in))
    return para


def p(doc, text="", bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0.5, space_before=0, space_after=6, hanging=0, font_size=12, line_spacing=2.0):
    if text:
        text = _rapikan_caption(str(text))
        text = _rapikan_pola_sitasi(str(text))
    para = doc.add_paragraph()
    para.alignment = align
    para.paragraph_format.line_spacing = line_spacing
    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after = Pt(space_after)
    if hanging > 0:
        para.paragraph_format.left_indent = Inches(hanging)
        para.paragraph_format.first_line_indent = Inches(-hanging)
    elif indent > 0:
        para.paragraph_format.first_line_indent = Inches(indent)
    
    if text:
        add_rich_runs_to_para(para, text, bold_default=bold, italic_default=italic, font_size=font_size)
        _hang_number(para)   # judul/butir bernomor otomatis bergaya menggantung
    return para

def _toc_line(doc, title, page):
    """Satu baris Daftar Isi/Tabel/Gambar/Lampiran.

    Memakai TAB dengan titik pengarah (dot leader) rata kanan pada tepi margin kanan
    (14,0 cm) sehingga nomor halaman selalu lurus di kanan dan judul panjang tidak
    terpotong menjadi baris ". . ." menggantung. Indentasi dibuat dari spasi awal judul
    (2 spasi = 0,5 cm) agar struktur bertingkat tetap terbaca.
    """
    teks = str(title)
    n_spasi = len(teks) - len(teks.lstrip(" "))
    teks = teks.lstrip(" ")
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = para.paragraph_format
    pf.line_spacing = 1.0
    pf.space_before = Pt(0)
    pf.space_after = Pt(1)
    if n_spasi:
        pf.left_indent = Cm(0.25 * n_spasi)
    pf.tab_stops.clear_all()
    pf.tab_stops.add_tab_stop(Cm(14.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    run = para.add_run("%s\t%s" % (teks, page))
    run.font.name = "Times New Roman"
    run.font.size = Pt(11)
    return para


def _set_outline_level(para, level):
    """Tandai paragraf sebagai JUDUL BERTINGKAT (outline level) untuk Daftar Isi otomatis.

    level 0 = BAB / halaman utama (BAB I, DAFTAR PUSTAKA, LAMPIRAN, ABSTRAK, ...)
    level 1 = sub-bab (1.1., 2.1., 4.1., 4.2., ...)
    level 2 = sub-sub-bab (2.1.1., 3.2.1., 4.1.3., 4.2.1., ...)

    Hanya menandai tingkat outline — TAMPILAN paragraf tidak berubah sama sekali.
    Field TOC di halaman DAFTAR ISI memakai tanda ini untuk menyusun daftar + nomor
    halaman secara otomatis (dihitung Word, bukan estimasi).
    """
    if para is None:
        return para
    try:
        pPr = para._p.get_or_add_pPr()
        el = pPr.find(qn('w:outlineLvl'))
        if el is None:
            el = parse_xml(r'<w:outlineLvl %s w:val="%d"/>' % (nsdecls('w'), int(level)))
            rPr = pPr.find(qn('w:rPr'))
            if rPr is not None:
                rPr.addprevious(el)      # urutan skema XML: outlineLvl sebelum rPr
            else:
                pPr.append(el)
        else:
            el.set(qn('w:val'), str(int(level)))
    except Exception as ex:
        print(f"[outline] gagal menandai judul: {ex}")
    return para


def _toc_field(doc, switch=r'TOC \o "1-3" \h \z \u', charformat=False):
    """Sisipkan FIELD DAFTAR ISI otomatis (Word yang menghitung nomor halaman).

    Dipakai menggantikan daftar isi statis: Word mengumpulkan semua judul bertingkat
    (outline level) pada dokumen beserta nomor halamannya, dan memperbaruinya sendiri
    setiap kali field diperbarui (klik kanan > Update Field, atau otomatis saat dibuka
    karena dokumen menyalakan w:updateFields).

    ``charformat=True`` (DAFTAR TABEL/GAMBAR/LAMPIRAN) menambahkan switch
    ``\\* CHARFORMAT``: format huruf pertama kode field (TNR 12, TANPA tebal) dipakai
    untuk seluruh hasil daftar, sehingga tetap tidak tebal walau field diperbarui.
    Gaya judul/caption di dalam bab tidak terpengaruh.
    """
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = para.paragraph_format
    pf.line_spacing = 1.0
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)

    run = para.add_run()
    if charformat:
        # Inilah "huruf pertama field name" yang diambil switch \* CHARFORMAT.
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.font.bold = False
    run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w')))
    instr = parse_xml(r'<w:instrText %s xml:space="preserve"/>' % nsdecls('w'))
    instr.text = " " + switch + " "
    run._r.append(instr)
    run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w')))

    # Teks sementara — otomatis HILANG saat Word memperbarui field.
    kosong = para.add_run("(Daftar Isi otomatis. Klik kanan di sini lalu pilih "
                          "\u201cUpdate Field\u201d untuk menghitung nomor halaman.)")
    kosong.italic = True
    kosong.font.size = Pt(11)

    end_run = para.add_run()
    end_run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w')))
    return para


def _set_update_fields(doc):
    """Nyalakan 'perbarui field saat dokumen dibuka' -> Daftar Isi otomatis terhitung."""
    try:
        settings = doc.settings.element
    except Exception:
        return
    try:
        el = settings.find(qn('w:updateFields'))
        if el is None:
            el = parse_xml(r'<w:updateFields %s w:val="true"/>' % nsdecls('w'))
            settings.append(el)
        else:
            el.set(qn('w:val'), 'true')
    except Exception as ex:
        print(f"[updateFields] gagal menyalakan: {ex}")


def add_page_number_footer(section, numfmt="decimal", start=None, font_size=11):
    """Sisipkan nomor halaman (field PAGE) di footer section + atur formatnya.

    ``numfmt`` : "decimal" (1, 2, 3) atau "lowerRoman" (i, ii, iii).
    ``start``  : nomor halaman awal section (None = lanjut dari section sebelumnya).
    """
    try:
        sectPr = section._sectPr
        el = sectPr.find(qn('w:pgNumType'))
        if el is None:
            el = parse_xml(r'<w:pgNumType %s/>' % nsdecls('w'))
            sectPr.append(el)
        el.set(qn('w:fmt'), numfmt)
        if start is not None:
            el.set(qn('w:start'), str(int(start)))
        elif el.get(qn('w:start')) is not None:
            del el.attrib[qn('w:start')]
    except Exception as ex:
        print(f"[pgNum] gagal mengatur format nomor halaman: {ex}")

    try:
        footer = section.footer
        footer.is_linked_to_previous = False
        para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        for r in list(para.runs):
            r._r.getparent().remove(r._r)
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        run.font.name = "Times New Roman"
        run.font.size = Pt(font_size)
        run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w')))
        instr = parse_xml(r'<w:instrText %s xml:space="preserve"/>' % nsdecls('w'))
        instr.text = " PAGE "
        run._r.append(instr)
        run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w')))
        mid = para.add_run("1")
        mid.font.name = "Times New Roman"
        mid.font.size = Pt(font_size)
        end = para.add_run()
        end._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w')))
    except Exception as ex:
        print(f"[pgNum] gagal membuat footer nomor halaman: {ex}")
    return section


def apply_section_margins(section):
    """Terapkan margin baku disertasi (4-4-3-3 cm) pada section baru."""
    try:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(4.0)
        section.left_margin = Cm(4.0)
        section.bottom_margin = Cm(3.0)
        section.right_margin = Cm(3.0)
    except Exception:
        pass
    return section


# ============ GAYA PARAGRAF CAPTION (untuk DAFTAR TABEL/GAMBAR/LAMPIRAN OTOMATIS) ============
# Word bisa menyusun daftar otomatis dari GAYA paragraf (switch TOC "\t").
# Tiga gaya ini HANYA menandai caption; tampilannya tetap dari format langsung
# (bold, rata tengah, ukuran font) sehingga tidak ada perubahan visual.

STYLE_CAP_TABEL = "Keterangan Tabel"
STYLE_CAP_GAMBAR = "Keterangan Gambar"
STYLE_CAP_LAMPIRAN = "Keterangan Lampiran"

_RX_CAP_TABEL = re.compile(r"^Tabel\s+(?:\d+|[A-Z])\s*\.\s*\d+\b")
_RX_CAP_GAMBAR = re.compile(r"^Gambar\s+(?:\d+|[A-Z])\s*\.\s*\d+\b")
_RX_CAP_LAMPIRAN = re.compile(r"^Lampiran\s+\d+\b")


def ensure_caption_styles(doc):
    """Sediakan gaya paragraf caption (dibuat sekali per dokumen)."""
    from docx.enum.style import WD_STYLE_TYPE
    for nm in (STYLE_CAP_TABEL, STYLE_CAP_GAMBAR, STYLE_CAP_LAMPIRAN):
        try:
            doc.styles[nm]
        except KeyError:
            st = doc.styles.add_style(nm, WD_STYLE_TYPE.PARAGRAPH)
            try:
                st.base_style = None
            except Exception:
                pass
    return doc


# Tingkat OUTLINE untuk caption. DAFTAR ISI memakai \o "1-3", sehingga caption
# TIDAK ikut ke DAFTAR ISI, tetapi dikumpulkan oleh daftar khususnya sendiri:
#   DAFTAR TABEL    -> \o "4-4"   (tingkat outline 4 -> w:outlineLvl val="3")
#   DAFTAR GAMBAR   -> \o "5-5"   (tingkat outline 5 -> w:outlineLvl val="4")
#   DAFTAR LAMPIRAN -> \o "6-6"   (tingkat outline 6 -> w:outlineLvl val="5")
OUTLINE_CAP_TABEL = 3
OUTLINE_CAP_GAMBAR = 4
OUTLINE_CAP_LAMPIRAN = 5

# Baris daftar LAMA (statis) memakai TAB + nomor halaman di ujung -> jangan ditandai.
_RX_BARIS_DAFTAR_LAMA = re.compile(r"\t\s*(\d+|[ivxlcdm]+)\.?\s*$", re.I)

# Baris pertama caption berformat HANYA nomor tabel/gambar (judulnya di baris berikutnya).
_RX_CAP_LABEL_ONLY = re.compile(r"^(?:Tabel|Gambar)\s+(?:\d+|[A-Z])\s*\.\s*\d+$")


def _layak_caption(t):
    """Pastikan paragraf memang CAPTION, bukan kalimat naskah yang menyebut nomor tabel.

    Caption selalu berbentuk: baris pertama HANYA "Tabel 4. 6" (judul menyusul pada
    baris berikutnya), atau judul pendek satu baris. Kalimat naskah seperti
    "Tabel 4.8 menampilkan perbandingan ..." tidak lolos karena panjang.
    """
    baris1 = re.split(r"[\v\n]", t)[0].strip()
    if _RX_CAP_LABEL_ONLY.match(baris1):
        return True
    return len(t) <= 140


def mark_caption_styles(doc):
    """Tandai caption Tabel/Gambar/Lampiran agar Word menyusun DAFTAR TABEL /
    DAFTAR GAMBAR / DAFTAR LAMPIRAN secara OTOMATIS (dihitung Word, bukan estimasi).

    Caption dikenali dari POSISINYA (bukan sekadar teks), supaya kalimat naskah
    yang menyebut nomor tabel tidak ikut tertandai:
      * caption TABEL  = paragraf tepat SEBELUM tabel;
      * caption GAMBAR = paragraf tepat SESUDAH gambar;
      * judul LAMPIRAN = paragraf berjudul "Lampiran N ..." yang punya tingkat outline.

    Dua tanda dipasang — TAMPILAN paragraf tidak berubah karena format langsung
    (bold, rata tengah, ukuran font) tetap menang:
      1. GAYA paragraf khusus ("Keterangan Tabel" / "Keterangan Gambar" /
         "Keterangan Lampiran") — memudahkan pengenalan & penyuntingan di Word;
      2. TINGKAT OUTLINE 4/5/6 — inilah yang dipakai field TOC
         (`\\o "4-4" \\u` untuk DAFTAR TABEL, `\\o "5-5" \\u` untuk GAMBAR,
         `\\o "6-6" \\u` untuk LAMPIRAN). Berada di luar rentang "1-3" milik
         DAFTAR ISI, sehingga caption tidak muncul dobel di DAFTAR ISI.
    """
    ensure_caption_styles(doc)
    from docx.text.paragraph import Paragraph

    jml = {"tabel": 0, "gambar": 0, "lampiran": 0}

    def _pasang(para, nm, lvl, kunci):
        try:
            para.style = doc.styles[nm]
        except Exception:
            pass
        _set_outline_level(para, lvl)
        jml[kunci] += 1

    anak = list(doc.element.body.iterchildren())
    par_sebelum = None      # paragraf isi terakhir (kandidat caption TABEL)
    for i, child in enumerate(anak):
        tag = child.tag.split("}")[-1]
        if tag == "tbl":
            if par_sebelum is not None:
                t = (par_sebelum.text or "").strip()
                if _RX_CAP_TABEL.match(t) and not _RX_BARIS_DAFTAR_LAMA.search(t) and _layak_caption(t):
                    _pasang(par_sebelum, STYLE_CAP_TABEL, OUTLINE_CAP_TABEL, "tabel")
            par_sebelum = None
            continue
        if tag != "p":
            continue
        par = Paragraph(child, doc)
        t = (par.text or "").strip()
        if child.findall(".//" + qn('w:drawing')) or child.findall(".//" + qn('w:pict')):
            # Paragraf bergambar -> caption GAMBAR adalah paragraf SESUDAHNYA.
            par_sebelum = None
            if i + 1 < len(anak) and anak[i + 1].tag.split("}")[-1] == "p":
                par_next = Paragraph(anak[i + 1], doc)
                tn = (par_next.text or "").strip()
                if _RX_CAP_GAMBAR.match(tn) and _layak_caption(tn):
                    _pasang(par_next, STYLE_CAP_GAMBAR, OUTLINE_CAP_GAMBAR, "gambar")
            continue
        if t:
            par_sebelum = par
        # Caption GAMBAR yang tidak memuat objek gambar tersemat (mis. bagan/kerangka
        # berpikir yang digambar manual oleh peneliti): dikenali dari baris pertama
        # yang HANYA memuat nomor gambar ("Gambar 2.1"), sehingga kalimat naskah tidak
        # ikut tertandai.
        if _RX_CAP_GAMBAR.match(t):
            _baris1 = re.split(r"[\v\n]", t)[0].strip()
            if _RX_CAP_LABEL_ONLY.match(_baris1):
                _pasang(par, STYLE_CAP_GAMBAR, OUTLINE_CAP_GAMBAR, "gambar")
        # Judul LAMPIRAN (mis. h2 "Lampiran 1  Kuesioner Penelitian")
        if _RX_CAP_LAMPIRAN.match(t):
            pPr = par._p.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:outlineLvl')) is not None:
                _pasang(par, STYLE_CAP_LAMPIRAN, OUTLINE_CAP_LAMPIRAN, "lampiran")
    return jml


def p_numbered(doc, num_str, title_str="", text_str="", space_after=3, line_spacing=2.0):
    """
    Penomoran standar ilmiah (Hanging Indent) 2 spasi:
    - Angka (1., 2., dst.) di tepi kiri margin (tidak menjorok).
    - Teks baris pertama mulai di posisi 0.5 inci melalui tab stop,
      sejajar dengan awal teks isi dan judul sub-bab.
    - Baris kedua dan seterusnya sejajar rata kiri dengan awal baris pertama (0.5 inci).
    """
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para.paragraph_format.line_spacing = line_spacing
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(space_after)
    
    indent_width = Inches(HANG_IN)
    para.paragraph_format.left_indent = indent_width
    para.paragraph_format.first_line_indent = -indent_width
    para.paragraph_format.tab_stops.add_tab_stop(indent_width)

    # 1. Angka di margin kiri + TAB (\t) agar baris pertama melompat tepat ke posisi indentasi baris kedua
    r_num = para.add_run(f"{num_str.strip()}\t")
    r_num.font.name = "Times New Roman"
    r_num.font.size = Pt(12)
    r_num.bold = False

    # 2. Judul poin (jika ada) — TIDAK bold: hanya judul BAB & sub-bab yang bold
    if title_str:
        clean_title = title_str.strip().replace("(*", "(").replace("*)", ")").replace("*", "")
        add_rich_runs_to_para(para, f"{clean_title}: ", bold_default=False, font_size=12)

    # 3. Isi uraian dengan dukungan kata asing miring
    if text_str:
        add_rich_runs_to_para(para, text_str.strip(), font_size=12)
    return para

def h1(doc, text):
    # Outline level 0 -> masuk Daftar Isi otomatis sebagai tingkat utama (BAB).
    return _set_outline_level(
        p(doc, text, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_before=14, space_after=6, line_spacing=2.0), 0)

def h2(doc, text):
    # Outline level 1 -> sub-bab (1.1., 4.2., ...).
    return _set_outline_level(
        p(doc, text, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_before=12, space_after=4, line_spacing=2.0), 1)

def h3(doc, text):
    # Judul sub-bab: BOLD saja, TIDAK miring. Istilah bahasa Inggris di dalamnya
    # tetap otomatis miring lewat add_rich_runs_to_para().
    # Outline level 2 -> sub-sub-bab (2.1.1., 3.2.1., 4.1.3., 4.2.1., ...).
    return _set_outline_level(
        p(doc, text, bold=True, italic=False, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_before=8, space_after=2, line_spacing=2.0), 2)

def apply_full_borders_table(table):
    """
    Tabel Standar Ilmiah:
    - Tampilkan SEMUA GARIS (semua batas: top, bottom, left, right, insideH, insideV)
    - Paskan ke jendela (100% margin lebar / 5000 pct) & paskan ke konten (autofit layout)
    - Isi tabel: Font Times New Roman 12 pt, 1 Spasi (Single Spacing)
    """
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    tblPr = table._tbl.tblPr

    # 1. Paskan ke jendela (100% lebar margin / 5000 pct)
    existing_tblW = tblPr.find(qn('w:tblW'))
    if existing_tblW is not None:
        tblPr.remove(existing_tblW)
    tblPr.append(parse_xml(r'<w:tblW %s w:w="5000" w:type="pct"/>' % nsdecls('w')))

    # 2. Paskan ke konten (autofit layout)
    existing_tblLayout = tblPr.find(qn('w:tblLayout'))
    if existing_tblLayout is not None:
        tblPr.remove(existing_tblLayout)
    tblPr.append(parse_xml(r'<w:tblLayout %s w:type="autofit"/>' % nsdecls('w')))

    # 3. Tampilkan semua garis (Semua Batas / All Borders)
    existing_tblBorders = tblPr.find(qn('w:tblBorders'))
    if existing_tblBorders is not None:
        tblPr.remove(existing_tblBorders)
    tblBorders = parse_xml(
        r'<w:tblBorders %s>'
        r'  <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        r'  <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        r'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        r'  <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        r'  <w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        r'  <w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(tblBorders)

    # 4. Standar Isi Tabel: Font Times New Roman 12 pt, 1 Spasi (Single Spacing)
    for row_idx, row in enumerate(table.rows):
        is_header = (row_idx == 0)
        trPr = row._tr.get_or_add_trPr()
        if is_header:
            trPr.append(parse_xml(r'<w:tblHeader %s/>' % nsdecls('w')))
        trPr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))

        for cell in row.cells:
            for para in cell.paragraphs:
                para.paragraph_format.line_spacing = 1.0
                para.paragraph_format.space_before = Pt(2)
                para.paragraph_format.space_after = Pt(2)
                for run in para.runs:
                    run.font.name = "Times New Roman"
                    if is_header:
                        run.bold = True

# Alias kompatibilitas
apply_open_table = apply_full_borders_table

def apply_fixed_widths(table, widths_cm):
    """Kunci lebar kolom tabel (cm) agar proporsional seperti contoh tesis.

    Menerima dua bentuk: angka cm (mis. 3.2) ATAU objek docx.shared.Length
    (mis. Cm(3.2) / Inches(1)). Tanpa normalisasi ini, Length (EMU) akan
    dibaca sebagai cm sehingga lebar kolom membesar ~360.000x dan Word
    "memperbaiki" tabel menjadi puluhan kolom saat disimpan.
    """
    def _as_cm(w):
        return float(w.cm) if hasattr(w, "cm") else float(w)

    widths_cm = [_as_cm(w) for w in widths_cm]

    # Jaga agar total lebar kolom tidak melebihi lebar blok teks halaman.
    try:
        _doc = table.part.document
        _sec = _doc.sections[0]
        _max_cm = _sec.page_width.cm - _sec.left_margin.cm - _sec.right_margin.cm
        _total = sum(widths_cm)
        if _total > _max_cm + 0.01 and _total > 0:
            _f = _max_cm / _total
            widths_cm = [w * _f for w in widths_cm]
    except Exception:
        pass

    tblPr = table._tbl.tblPr
    for el in tblPr.findall(qn('w:tblLayout')):
        tblPr.remove(el)
    tblPr.append(parse_xml(r'<w:tblLayout %s w:type="fixed"/>' % nsdecls('w')))
    table.autofit = False
    grid = table._tbl.find(qn('w:tblGrid'))
    if grid is not None:
        cols = grid.findall(qn('w:gridCol'))
        for gc, w in zip(cols, widths_cm):
            gc.set(qn('w:w'), str(int(Cm(w).twips)))
    for row in table.rows:
        for cell, w in zip(row.cells, widths_cm):
            cell.width = Cm(w)

def add_table_data(doc, title, headers, data, source="", alignments=None, widths=None,
                   cell_font_size=12, body_align_left=False, autofit_contents=False,
                   header_rows=1):
    if title:
        # Judul Tabel CENTER Rata Tengah, 1 Spasi (Bukan 2 Spasi)
        p(doc, title, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_before=12, space_after=4, line_spacing=1.0)
    table = doc.add_table(rows=len(data) + 1, cols=len(headers))
    header_row = table.rows[0]
    for c_idx, h in enumerate(headers):
        cell = header_row.cells[c_idx]
        para = cell.paragraphs[0]
        for r in list(para.runs):
            r._r.getparent().remove(r._r)
        add_rich_runs_to_para(para, str(h if h is not None else ""), bold_default=True, font_size=cell_font_size)
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r_idx, row_data in enumerate(data):
        row = table.rows[r_idx + 1]
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            para = cell.paragraphs[0]
            for r in list(para.runs):
                r._r.getparent().remove(r._r)
            add_rich_runs_to_para(para, str(val if val is not None else ""), bold_default=False, font_size=cell_font_size)
            if alignments and c_idx < len(alignments):
                para.alignment = alignments[c_idx]
            else:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_full_borders_table(table)
    # Warna biru pada baris header (semua tabel).
    set_header_shading(table, header_rows)
    # Ukuran font isi tabel (default 12; mis. Tabel 2.1 memakai 10 pt)
    if cell_font_size and float(cell_font_size) != 12.0:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(cell_font_size)
    if widths:
        try:
            apply_fixed_widths(table, widths)
        except Exception:
            pass

    # Khusus tabel Bab I-Bab III & tabel kuesioner Lampiran:
    # isi sel rata KIRI (header tetap tengah) + "Paskan Konten" (AutoFit Contents).
    if body_align_left:
        align_table_body(table)
    if autofit_contents:
        set_autofit_contents(table)
    
    # Di bawah tabel: tulisan "Sumber: ..."
    # Font 12, TIDAK HURUF MIRING (kecuali istilah bahasa Inggris), 1 SPASI
    if source:
        source_clean = source.strip()
        if not source_clean.lower().startswith("sumber"):
            source_clean = f"Sumber: {source_clean}"
        p(doc, source_clean, italic=False, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_before=3, space_after=8, font_size=12, line_spacing=1.0)
    else:
        p(doc, "", indent=0, space_after=4, line_spacing=1.0)


# ============ UTILITAS PENULISAN (sentence case) & TABEL KUESIONER ============

# Gemini kerap menulis "Sembiring et al., (2022)" (koma sebelum kurung).
# Gaya APA yang benar: "Sembiring et al. (2022)".
_RX_SITASI_KOMA = re.compile(r"\b(et al\.|dkk\.)\s*,\s*\(")


def _rapikan_pola_sitasi(teks):
    """Rapikan pola sitasi yang menyalahi gaya APA tanpa mengubah isi."""
    return _RX_SITASI_KOMA.sub(r"\1 (", teks)


# Kata umum (Indonesia/Inggris) yang tidak perlu kapital di tengah judul.
_ID_KECIL = {
    "dan", "atau", "yang", "untuk", "pada", "dalam", "dengan", "terhadap", "melalui",
    "sebagai", "dari", "ke", "di", "oleh", "serta", "juga", "tidak", "telah", "akan",
    "dapat", "adalah", "ini", "itu", "para", "suatu", "sebuah", "agar", "karena",
    "sehingga", "namun", "tetapi", "bila", "jika", "maka", "secara", "antara", "antar",
    "bagi", "tentang", "mengenai", "saat", "ketika", "hingga", "sampai", "baik", "maupun",
    "yaitu", "yakni", "guna", "demi", "tanpa", "atas", "bawah", "pun",
    "pengaruh", "hubungan", "peran", "analisis", "studi", "kajian", "penelitian",
    "efek", "dampak", "kinerja", "pegawai", "karyawan", "kerja", "kepuasan", "motivasi",
    "disiplin", "lingkungan", "organisasi", "budaya", "gaya", "kepemimpinan", "komitmen",
    "kompetensi", "pelatihan", "pengembangan", "karir", "karier", "mutasi", "promosi",
    "jabatan", "beban", "stres", "stress", "kreativitas", "kerjasama", "sama", "tim",
    "komunikasi", "produktivitas", "kualitas", "pelayanan", "publik", "aparatur",
    "pemerintah", "daerah", "instansi", "mediasi", "intervening", "moderasi",
    "variabel", "variable", "peningkatan", "penerapan", "implementasi", "evaluasi",
    "pengukuran", "tingkat", "tinggi", "rendah", "sedang", "besar", "kecil", "kuat",
    "lemah", "positif", "signifikan", "memediasi", "mempengaruhi", "memengaruhi",
    "berpengaruh", "ditingkatkan", "kasus", "perbandingan", "komparasi", "persamaan",
    "perbedaan", "hasil", "menunjukkan", "membuktikan", "bahwa", "parsial", "simultan",
    "langsung", "kunci", "faktor", "penentu", "dimensi", "indikator", "kisi",
    "instrumen", "butir", "pernyataan", "kuesioner", "responden", "sampel", "populasi",
    "the", "of", "and", "or", "in", "on", "at", "for", "with", "by", "from", "to",
    "a", "an", "as", "is", "are", "was", "were", "be", "this", "that", "study",
    "analysis", "effect", "effects", "impact", "role", "mediating", "mediation",
    "influence", "toward", "towards", "review", "systematic", "literature",
    "empirical", "evidence", "based", "case", "self", "efficacy", "work",
    "environment", "employee", "engagement", "performance", "job", "satisfaction",
    "organizational", "behavior", "behaviour", "commitment", "leadership", "style",
    "training", "development", "career", "promotion", "transfer", "discipline",
    "quality", "service", "public", "sector", "government", "civil", "servant",
    "into", "under", "among", "between",
    # kata umum Inggris (judul artikel/academic)
    "influence", "influences", "influencing", "impact", "impacts", "effect", "effects",
    "culture", "motivation", "motivations", "discipline", "satisfaction", "commitment",
    "engagement", "performance", "productivity", "quality", "service", "services",
    "leadership", "style", "styles", "behavior", "behaviour", "organizational",
    "organization", "organisational", "human", "resource", "resources", "management",
    "managerial", "business", "firm", "firms", "company", "companies", "industry",
    "industries", "sector", "sectors", "private", "governance", "policy", "policies",
    "employee", "employees", "worker", "workers", "staff", "personnel", "training",
    "development", "career", "promotion", "transfer", "compensation", "reward",
    "rewards", "benefits", "workload", "burnout", "turnover", "intention", "retention",
    "loyalty", "trust", "communication", "collaboration", "teamwork", "innovation",
    "creativity", "competency", "competence", "skills", "ability", "knowledge",
    "attitude", "perception", "perceived", "relationship", "relationships",
    "association", "correlation", "determinant", "determinants", "factors", "factor",
    "variables", "model", "models", "framework", "theory", "theoretical",
    "review", "literature", "meta", "surveys", "survey", "questionnaire",
    "respondents", "sample", "population", "data", "method", "methods", "methodology",
    "analyses", "results", "findings", "discussion", "conclusion", "conclusions",
    "implications", "limitations", "recommendation", "recommendations", "approach",
    "approaches", "introduction", "overview", "evidence", "based", "using", "case",
    "moderating", "moderation", "mediating", "intervening", "direct", "indirect",
    "partial", "full", "significant", "significance", "positive", "negative",
    "public", "government", "governance", "civil", "servant", "servants",
    # kata umum Indonesia tambahan
    "tambahan", "efektivitas", "efisiensi", "pencapaian", "target", "sasaran",
    "tujuan", "masalah", "fenomena", "kontribusi", "negatif", "dimediasi",
    "dipengaruhi", "upaya", "usaha", "cara", "fungsi", "tugas", "tanggung", "jawab",
    "setiap", "seluruh", "semua", "beberapa", "tersebut", "proses", "bentuk",
    "jenis", "perlu", "harus", "bisa", "belum", "bukan", "hanya", "saja", "lebih",
    "paling", "sangat", "cukup", "kurang", "seperti", "yaitu", "misalnya",
}

# Kata yang dipertahankan kapital (nama diri / tempat / institusi).
_PENULISAN_PROPER = {
    "indonesia", "puskesmas", "kabupaten", "kecamatan", "kelurahan", "provinsi",
    "kota", "desa", "universitas", "sekolah", "bank", "badan", "dinas", "kantor",
    "rumah", "sakit", "sumatera", "jawa", "kalimantan", "sulawesi", "papua", "bali",
    "jakarta", "palembang", "bandung", "surabaya", "medan", "semarang", "makassar",
    "ogan", "komering", "ilir", "utara", "selatan", "timur", "barat", "tengah",
    "sekretariat", "kementerian", "direktorat", "pajak",
}


def smart_sentence_case(text, proper_extra=()):
    """Ubah judul ber-Title Case / HURUF KAPITAL SEMUA menjadi sentence case.

    - Kata umum (``_ID_KECIL``) ditulis huruf kecil.
    - Judul yang seluruhnya kapital: seluruh kata menjadi huruf kecil.
    - Nama diri/tempat dan akronim (mis. PLS-SEM, TQM) dipertahankan.
    - Huruf pertama judul selalu kapital.
    """
    if not text:
        return text
    t = re.sub(r"[ \t]+", " ", str(text).strip())
    proper = {str(w).lower() for w in proper_extra if str(w).strip()} | _PENULISAN_PROPER
    semua_kapital = (t.upper() == t and re.search(r"[A-Z]{3,}", t) is not None)
    hasil, sudah_awal = [], False
    for tok in re.split(r"(\s+)", t):
        if not tok.strip():
            hasil.append(tok)
            continue
        core = re.sub(r"^[^\w]+|[^\w]+$", "", tok)  # Unicode-aware (Ö, Ç, Ć, ...)
        if not core:
            hasil.append(tok)
            continue
        low = core.lower()
        if low in proper:
            baru = low[:1].upper() + low[1:]
        elif semua_kapital:
            baru = low
        elif core.isupper() and 2 <= len(core) <= 8:
            baru = core
        elif low in _ID_KECIL:
            baru = low
        else:
            baru = core
        if not sudah_awal and re.search(r"[A-Za-z]", baru):
            baru = baru[:1].upper() + baru[1:]
            sudah_awal = True
        hasil.append(tok.replace(core, baru))
    return "".join(hasil)


def set_autofit_contents(table, window=True):
    """Word: "Paskan Konten" + "Paskan Jendela".

    - ``w:tblLayout`` = autofit  → lebar kolom mengikuti isi sel (Paskan Konten),
    - ``w:tblW`` = pct 5000      → lebar tabel 100% area teks/halaman (Paskan Jendela),
    - lebar sel yang dikunci sebelumnya dibuang agar Word menghitung ulang.
    """
    tblPr = table._tbl.tblPr
    for el in tblPr.findall(qn('w:tblW')):
        tblPr.remove(el)
    if window:
        tblPr.append(parse_xml(r'<w:tblW %s w:w="5000" w:type="pct"/>' % nsdecls('w')))
    else:
        tblPr.append(parse_xml(r'<w:tblW %s w:w="0" w:type="auto"/>' % nsdecls('w')))
    for el in tblPr.findall(qn('w:tblLayout')):
        tblPr.remove(el)
    tblPr.append(parse_xml(r'<w:tblLayout %s w:type="autofit"/>' % nsdecls('w')))
    table.autofit = True
    # Buang lebar tetap pada tiap sel (w:tcW) dan samakan kisi awal (w:gridCol).
    grid = table._tbl.find(qn('w:tblGrid'))
    if grid is not None:
        cols = grid.findall(qn('w:gridCol'))
        if cols:
            w_each = int(round(Cm(14.0).twips / float(len(cols))))
            for gc in cols:
                gc.set(qn('w:w'), str(w_each))
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for el in tcPr.findall(qn('w:tcW')):
                tcPr.remove(el)
    return table


# Judul kolom yang isinya TETAP rata tengah (selain baris header).
_KEEP_CENTER_HEADERS = {"no", "nomor", "no item pertanyaan", "no. item pertanyaan"}

# Warna latar baris header tabel (biru muda) — dipakai semua tabel.
HEADER_FILL = "D9E2F3"


def set_header_shading(table, header_rows=1, fill=HEADER_FILL):
    """Beri warna latar biru pada baris header tabel (``header_rows`` baris teratas)."""
    for row in list(table.rows)[:max(1, int(header_rows))]:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for el in tcPr.findall(qn('w:shd')):
                tcPr.remove(el)
            tcPr.append(parse_xml(
                r'<w:shd %s w:val="clear" w:color="auto" w:fill="%s"/>' % (nsdecls('w'), fill)))
    return table


def align_table_body(table, header_rows=1, body_align=WD_ALIGN_PARAGRAPH.LEFT):
    """Isi sel tabel rata KIRI; baris header tetap rata TENGAH.

    Kolom "No" (nomor baris) dan "No. Item Pertanyaan" tetap rata TENGAH.
    ``header_rows`` = jumlah baris header (tabel kuesioner/jadwal memakai 2 baris).
    ``body_align`` dapat diganti (mis. RIGHT) bila diperlukan.
    """
    rows = list(table.rows)
    if len(rows) <= header_rows:
        return table
    keep = set()
    for i, cell in enumerate(rows[0].cells):
        h = re.sub(r"\s+", " ", (cell.text or "").strip().lower())
        if h in _KEEP_CENTER_HEADERS:
            keep.add(i)
    for row in rows[:header_rows]:
        for cell in row.cells:
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for row in rows[header_rows:]:
        for ci, cell in enumerate(row.cells):
            mau = WD_ALIGN_PARAGRAPH.CENTER if ci in keep else body_align
            for para in cell.paragraphs:
                para.alignment = mau
    return table


# Nama lama (isi rata KANAN) dipertahankan agar skrip lama tetap berjalan.
def align_table_body_right(table, header_rows=1):
    return align_table_body(table, header_rows=header_rows,
                            body_align=WD_ALIGN_PARAGRAPH.RIGHT)


def rapikan_tabel_dokumen(doc):
    """Seragamkan tampilan SEMUA tabel dokumen (dipakai jalur lama: thesis_bab*.py & engine_docx.py).

    Aturan — sama dengan dokumen hasil jalur pipeline:
      * setiap tabel di dalam BAB I..BAB V / Lampiran: baris header berwarna BIRU + rata TENGAH;
      * tabel BAB I, BAB II, BAB III serta tabel berkepala "PERNYATAAN"/"BULAN"
        (jadwal penelitian & kuesioner Lampiran): lebar 100% area teks + AutoFit
        ("Paskan Konten" + "Paskan Jendela") dan isi sel rata KIRI
        (kolom "No" / "No. Item Pertanyaan" tetap rata TENGAH);
      * tabel Lampiran lain (mis. panduan skala Likert): lebar dipaskan, isi tetap rata tengah;
      * tabel bagian awal (sebelum BAB I) dibiarkan apa adanya.
    """
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    bab = None
    jumlah = 0
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            baris1 = ((Paragraph(child, doc).text or "").strip().split("\n")[0]).strip().upper()
            if baris1.startswith("BAB ") and len(baris1) <= 8:
                bab = baris1
            elif baris1.startswith("LAMPIRAN") and len(baris1) <= 14:
                bab = "LAMPIRAN"
            elif baris1.startswith("DAFTAR PUSTAKA"):
                bab = "DAFTAR PUSTAKA"
            continue
        if tag != "tbl" or bab is None:
            continue

        table = Table(child, doc)
        kepala = " ".join((c.text or "").strip().upper() for c in table.rows[0].cells)
        kuesioner = ("PERNYATAAN" in kepala) or ("BULAN" in kepala)
        hdr = 2 if kuesioner else 1
        if len(table.rows) <= hdr:
            hdr = 1

        set_header_shading(table, hdr)
        rata_kiri = bab in ("BAB I", "BAB II", "BAB III") or kuesioner
        if rata_kiri or bab == "LAMPIRAN":
            set_autofit_contents(table, window=True)
        if rata_kiri:
            align_table_body(table, hdr)
        else:
            for row in list(table.rows)[:hdr]:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        jumlah += 1

    # Tandai caption tabel/gambar/lampiran dengan gaya khusus supaya Word dapat
    # menyusun DAFTAR TABEL / GAMBAR / LAMPIRAN otomatis (field TOC "\t").
    try:
        mark_caption_styles(doc)
    except Exception as ex:
        print(f"[caption] gagal menandai gaya caption: {ex}")
    return jumlah


def _set_table_grid_fixed(table, widths_cm):
    """Kunci lebar kolom tabel (cm) dengan skema *fixed layout*."""
    tblPr = table._tbl.tblPr
    for el in tblPr.findall(qn('w:tblLayout')):
        tblPr.remove(el)
    tblPr.append(parse_xml(r'<w:tblLayout %s w:type="fixed"/>' % nsdecls('w')))
    table.autofit = False
    grid = table._tbl.find(qn('w:tblGrid'))
    if grid is not None:
        for gc, w in zip(grid.findall(qn('w:gridCol')), widths_cm):
            gc.set(qn('w:w'), str(int(Cm(w).twips)))


def add_likert_table(doc, statements, font_size=12, widths_cm=None):
    """Tabel kuesioner lampiran: NO | PERNYATAAN | PILIHAN (SS S KS TS STS).

    Header "PILIHAN" digabung melintang lima kolom; sel NO dan PERNYATAAN
    digabung tegak dua baris; header diberi latar biru muda; kolom jawaban
    dibiarkan kosong agar dapat dicentang responden.
    """
    choices = ["SS", "S", "KS", "TS", "STS"]
    stmts = list(statements or [])
    # Catatan: lebarnya TIDAK dikunci lagi. Tabel memakai "Paskan Konten"
    # (AutoFit Contents) sehingga Word menghitung lebar kolom dari isi sel.
    table = doc.add_table(rows=len(stmts) + 2, cols=7)
    r0, r1 = table.rows[0], table.rows[1]

    c_no = r0.cells[0].merge(r1.cells[0])
    c_no.text = "NO"
    c_pn = r0.cells[1].merge(r1.cells[1])
    c_pn.text = "PERNYATAAN"
    c_pl = r0.cells[2].merge(r0.cells[6])
    c_pl.text = "PILIHAN"
    for j, ch in enumerate(choices):
        r1.cells[2 + j].text = ch
    for i, s in enumerate(stmts, 1):
        row = table.rows[i + 1]
        row.cells[0].text = f"{i}."
        p_stmt = row.cells[1].paragraphs[0]
        for r in list(p_stmt.runs):
            r._r.getparent().remove(r._r)
        add_rich_runs_to_para(p_stmt, str(s if s is not None else ""), bold_default=False, font_size=font_size)
        for j in range(5):
            row.cells[2 + j].text = ""

    apply_full_borders_table(table)

    # Header: bold, rata tengah, latar biru muda (HEADER_FILL)
    set_header_shading(table, 2)
    for ri in (0, 1):
        for cell in table.rows[ri].cells:
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.bold = True
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(font_size)

    # Isi tabel: rata KIRI (header 2 baris di atas tetap rata tengah).
    # Kolom NO tetap rata tengah.
    for ri in range(2, len(table.rows)):
        row = table.rows[ri]
        row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        for j in range(2, 7):
            row.cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        for cell in row.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.bold = False
                    run.font.size = Pt(font_size)

    # "Paskan Konten" (AutoFit Contents): lebar kolom mengikuti isi sel.
    set_autofit_contents(table)
    return table

# ==================== BAGIAN AWAL ====================
_EN_VAR = {
    "kompensasi": "compensation",
    "kompensasi kerja": "compensation",
    "stres kerja": "work stress",
    "kecerdasan emosional": "emotional intelligence",
    "kepuasan kerja": "job satisfaction",
    "kinerja pegawai": "employee performance",
    "kinerja karyawan": "employee performance",
    "kinerja": "performance",
    "beban kerja": "workload",
    "komitmen organisasi": "organizational commitment",
    "motivasi kerja": "work motivation",
    "motivasi": "work motivation",
    "disiplin kerja": "work discipline",
    "disiplin": "discipline",
    "gaya kepemimpinan": "leadership style",
    "budaya organisasi": "organizational culture",
    "lingkungan kerja": "work environment",
    "pengembangan karier": "career development",
    "pengembangan karir": "career development",
    "kepemimpinan": "leadership",
    "insentif": "incentive",
    "tunjangan kerja": "work benefits",
}


def _en_var(nama):
    return _EN_VAR.get((nama or "").strip().lower(), nama)


def _nid(x, n=3):
    try:
        return f"{float(x):.{n}f}".replace(".", ",")
    except Exception:
        return str(x)


def _abstrak_otomatis(config):
    """Paragraf abstrak (ID & EN) dari angka analisis nyata, atau None.

    Sengaja deterministik: setiap angka diambil dari entri analisis PLS-SEM yang
    sama dengan Bab IV, sehingga abstrak tidak mungkin memuat angka karangan.
    """
    try:
        import engine_hierarchical_pipeline as ehp
        cfg = ehp.resolve_thesis_config(config)
        entry = ehp._thesis_analysis_entry(cfg)
    except Exception as e:
        print(f"[abstrak] dilewati: {e}")
        return None
    if not entry or not entry.get("metrics"):
        return None
    met = entry["metrics"]
    jalur = {(p.get("from"), p.get("to")): p for p in (met.get("paths") or [])}
    ind = {p.get("from"): p for p in (met.get("indirect") or [])}
    r2 = {r["construct"]: r["r2"] for r in (met.get("r2") or [])}
    srmr = met.get("srmr")

    x1, x2 = cfg["x1_title"], cfg["x2_title"]
    m, y = cfg["m_title"], cfg["y_title"]
    instansi = cfg["instansi"]
    n = entry.get("n") or cfg.get("n") or 0
    populasi = cfg.get("populasi") or entry.get("populasi") or 0
    teknik = cfg.get("teknik_sampel") or "sampling jenuh"

    def arah(pp):
        return "positif" if float((pp or {}).get("coef", 0)) >= 0 else "negatif"

    def sig(pp):
        return "signifikan" if pp and float(pp.get("p", 1)) < 0.05 else "tidak signifikan"

    def stat(pp, label=""):
        if not pp:
            return ""
        p_txt = "p < 0,05" if float(pp.get("p", 1)) < 0.05 else f"p = {_nid(pp.get('p'), 4)}"
        isi = f"(β = {_nid(pp['coef'])}; t = {_nid(pp['t'])}; {p_txt})"
        return f"{label} {isi}" if label else isi

    def vaf_jenis(pp):
        try:
            v = float(pp["vaf"]) * 100
        except Exception:
            return "tidak langsung"
        return "penuh" if v >= 80 else ("parsial" if v >= 20 else "lemah")

    def vaf_txt(pp):
        try:
            return f"{_nid(float(pp['vaf']) * 100, 1)} persen"
        except Exception:
            return "-"

    x1m, x2m = jalur.get(("X1", "M")), jalur.get(("X2", "M"))
    x1y, x2y = jalur.get(("X1", "Y")), jalur.get(("X2", "Y"))
    my = jalur.get(("M", "Y"))
    i1, i2 = ind.get("X1"), ind.get("X2")
    if not all((x1m, x2m, x1y, x2y, my)):
        return None

    _yp = y.lower() if "pegawai" in y.lower() else f"{y.lower()} pegawai"

    def _kap(s):
        return (s[:1].upper() + s[1:]) if s else s

    _hasil = (f"Hasil penelitian menunjukkan bahwa {x1.lower()} berpengaruh {arah(x1m)} dan "
              f"{sig(x1m)} terhadap {m.lower()} {stat(x1m)} dan terhadap {y.lower()} "
              f"{stat(x1y)}. {_kap(x2.lower())} juga berpengaruh {arah(x2m)} dan {sig(x2m)} "
              f"terhadap {m.lower()} {stat(x2m)} dan {y.lower()} {stat(x2y)}, sedangkan "
              f"{m.lower()} berpengaruh {arah(my)} dan {sig(my)} terhadap {y.lower()} {stat(my)}.")
    if i1 and i2:
        _hasil += (f" {_kap(m.lower())} terbukti memediasi secara {vaf_jenis(i1)} pengaruh "
                   f"{x1.lower()} (β = {_nid(i1['coef'])}; VAF = {vaf_txt(i1)}) dan "
                   f"{x2.lower()} (β = {_nid(i2['coef'])}; VAF = {vaf_txt(i2)}) terhadap "
                   f"{y.lower()}.")
    if r2.get("M") is not None and r2.get("Y") is not None:
        _hasil += (f" Nilai R-square tercatat sebesar {_nid(r2.get('M'))} untuk {m.lower()} "
                   f"dan {_nid(r2.get('Y'))} untuk {y.lower()}")
        _hasil += (f", dengan SRMR {_nid(srmr)} yang menunjukkan kecocokan model memadai."
                   if srmr is not None else ".")

    if populasi and int(populasi) > int(n):
        par1_id = (f"Penelitian ini bertujuan menganalisis pengaruh {x1.lower()} dan {x2.lower()} "
                   f"terhadap {m.lower()} serta dampaknya pada {y.lower()} di {instansi}. Penelitian "
                   f"menggunakan pendekatan kuantitatif dengan metode survei melalui kuesioner berskala "
                   f"Likert 1 sampai 5 yang disusun dari indikator setiap variabel. Populasi penelitian "
                   f"berjumlah {populasi} pegawai dengan penentuan sampel sebanyak {n} responden "
                   f"menggunakan teknik {teknik}. Analisis data dilakukan dengan "
                   f"Structural Equation Modeling berbasis Partial Least Squares (PLS-SEM) menggunakan "
                   f"perangkat lunak SmartPLS, yang meliputi evaluasi model pengukuran, evaluasi model "
                   f"struktural, dan pengujian hipotesis melalui bootstrapping 5.000 subsampel.")
    else:
        par1_id = (f"Penelitian ini bertujuan menganalisis pengaruh {x1.lower()} dan {x2.lower()} "
                   f"terhadap {m.lower()} serta dampaknya pada {y.lower()} di {instansi}. Penelitian "
                   f"menggunakan pendekatan kuantitatif dengan metode survei melalui kuesioner berskala "
                   f"Likert 1 sampai 5 yang disusun dari indikator setiap variabel. Populasi penelitian "
                   f"berjumlah {n} pegawai dan seluruhnya dijadikan sampel melalui teknik {teknik}, "
                   f"sehingga data yang dianalisis mencakup {n} responden. Analisis data dilakukan dengan "
                   f"Structural Equation Modeling berbasis Partial Least Squares (PLS-SEM) menggunakan "
                   f"perangkat lunak SmartPLS, yang meliputi evaluasi model pengukuran, evaluasi model "
                   f"struktural, dan pengujian hipotesis melalui bootstrapping 5.000 subsampel.")

    id_par = [
        par1_id,
        _hasil,
        (f"Dengan demikian, {x1.lower()} dan {x2.lower()} terbukti menjadi determinan penting "
         f"{m.lower()} yang pada akhirnya membentuk {_yp}. Pengelolaan "
         f"{x1.lower()} yang adaptif serta pengembangan {x2.lower()} perlu diarahkan pada "
         f"penguatan {m.lower()} agar {_yp} meningkat secara berkelanjutan."),
    ]

    ex1, ex2 = _en_var(x1), _en_var(x2)
    em, ey = _en_var(m), _en_var(y)
    _eyp = ey.lower() if "employee" in ey.lower() else f"{ey.lower()} of employees"
    if "slovin" in (teknik or "").lower():
        _teknik_en = "the Slovin formula"
    elif "sensus" in (teknik or "").lower() or "jenuh" in (teknik or "").lower():
        _teknik_en = "a census (saturated sampling)"
    elif "purposive" in (teknik or "").lower():
        _teknik_en = "purposive sampling"
    else:
        _teknik_en = teknik

    def _nen(x, n=3):
        try:
            return f"{float(x):.{n}f}"
        except Exception:
            return str(x)

    def estat(pp):
        if not pp:
            return ""
        p_txt = "p < 0.05" if float(pp.get("p", 1)) < 0.05 else f"p = {_nen(pp.get('p'), 4)}"
        return f"(β = {_nen(pp['coef'])}; t = {_nen(pp['t'])}; {p_txt})"

    def esig(pp):
        return "significantly" if pp and float(pp.get("p", 1)) < 0.05 else "not significantly"

    def earah(pp):
        return "positively" if float((pp or {}).get("coef", 0)) >= 0 else "negatively"

    _hasil_en = (f"The results show that {ex1.lower()} {earah(x1m)} and {esig(x1m)} affects "
                 f"{em.lower()} {estat(x1m)} and {ey.lower()} {estat(x1y)}. {ex2.capitalize()} "
                 f"also {earah(x2m)} and {esig(x2m)} affects {em.lower()} {estat(x2m)} and "
                 f"{ey.lower()} {estat(x2y)}, while {em.lower()} {earah(my)} and "
                 f"{esig(my)} affects {ey.lower()} {estat(my)}.")
    if i1 and i2:
        _hasil_en += (f" {em.capitalize()} partially mediates the effect of {ex1.lower()} "
                      f"(β = {_nen(i1['coef'])}; VAF = {_nen(float(i1['vaf']) * 100, 1)} percent) "
                      f"and {ex2.lower()} (β = {_nen(i2['coef'])}; "
                      f"VAF = {_nen(float(i2['vaf']) * 100, 1)} percent) on {ey.lower()}.")
    if r2.get("M") is not None and r2.get("Y") is not None:
        _hasil_en += (f" The R-square values are {_nen(r2.get('M'))} for {em.lower()} and "
                      f"{_nen(r2.get('Y'))} for {ey.lower()}")
        _hasil_en += (f", with an SRMR of {_nen(srmr)}, indicating an adequate model fit."
                      if srmr is not None else ".")

    if populasi and int(populasi) > int(n):
        par1_en = (f"This study aims to analyze the effect of {ex1.lower()} and {ex2.lower()} on "
                   f"{em.lower()} and its impact on {ey.lower()} at {instansi}. The research employed a "
                   f"quantitative approach with a survey method using a five-point Likert scale "
                   f"questionnaire developed from the indicators of each variable. The study population "
                   f"consisted of {populasi} employees, with a sample size of {n} respondents determined "
                   f"through {_teknik_en}. Data analysis was performed using Partial "
                   f"Least Squares Structural Equation Modeling (PLS-SEM) in SmartPLS, covering the "
                   f"measurement model, the structural model, and hypothesis testing through 5,000 "
                   f"bootstrap subsamples.")
    else:
        par1_en = (f"This study aims to analyze the effect of {ex1.lower()} and {ex2.lower()} on "
                   f"{em.lower()} and its impact on {ey.lower()} at {instansi}. The research employed a "
                   f"quantitative approach with a survey method using a five-point Likert scale "
                   f"questionnaire developed from the indicators of each variable. The population "
                   f"consisted of {n} employees, all of whom were selected through {_teknik_en}, so the "
                   f"analyzed data covered {n} respondents. Data analysis was performed using Partial "
                   f"Least Squares Structural Equation Modeling (PLS-SEM) in SmartPLS, covering the "
                   f"measurement model, the structural model, and hypothesis testing through 5,000 "
                   f"bootstrap subsamples.")

    en_par = [
        par1_en,
        _hasil_en,
        (f"Therefore, {ex1.lower()} and {ex2.lower()} are important determinants of "
         f"{em.lower()} that ultimately shape {ey.lower()}. Adaptive management of "
         f"{ex1.lower()} and continuous development of {ex2.lower()} should be directed at "
         f"strengthening {em.lower()} so that {_eyp} improves sustainably."),
    ]

    return {
        "id": id_par,
        "en": en_par,
        "kata_kunci": f"{x1}, {x2}, {m}, {y}, PLS-SEM, Mediasi",
        "keywords": f"{ex1.title()}, {ex2.title()}, {em.title()}, {ey.title()}, PLS-SEM, Mediation",
    }


LOGO_UNIV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "assets", "logo_universitas_tridinanti.png")


def _sisip_logo_cover(doc, lebar_in=2.5):
    """Menyisipkan lambang universitas pada sampul (rata tengah).

    Bila berkas logo tidak tersedia, paragraf kosong tetap dibuat agar tata
    letak tidak bergeser dan proses generate tidak pernah gagal karena gambar.
    """
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(0)
    par.paragraph_format.space_after = Pt(0)
    par.paragraph_format.first_line_indent = Inches(0)
    if os.path.exists(LOGO_UNIV_PATH):
        try:
            par.add_run().add_picture(LOGO_UNIV_PATH, width=Inches(lebar_in))
        except Exception as _ex:
            print(f"[cover] logo universitas gagal disisipkan: {_ex}")
    else:
        print(f"[cover] berkas logo tidak ditemukan: {LOGO_UNIV_PATH}")
    return par


def _baris_identitas_cover(doc, label, nilai):
    """Baris identitas penyusun pada sampul dengan kolom titik dua sejajar.

    Dibuat rata kiri berindentasi (tab) agar titik dua lurus satu sama lain,
    sesuai pedoman sampul tesis.
    """
    para = p(doc, f"{label}\t: {nilai}", bold=True,
             align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=2, line_spacing=1.0)
    pf = para.paragraph_format
    pf.left_indent = Cm(2.0)
    pf.tab_stops.clear_all()
    pf.tab_stops.add_tab_stop(Cm(4.6))
    return para


def _baris_kolom(doc, teks, tabs_cm=(), bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
                 space_before=0, space_after=0, font_size=12, line_spacing=1.0,
                 left_cm=None, first_cm=None):
    """Paragraf dengan TAB STOP eksplisit untuk blok berkolom yang lurus.

    Dibangun tanpa ``p()`` agar tidak dikenai penomoran menggantung otomatis
    (_hang_number) yang dapat merusak posisi tab stop.
    """
    para = doc.add_paragraph()
    para.alignment = align
    pf = para.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.first_line_indent = Inches(0)
    if left_cm is not None:
        pf.left_indent = Cm(left_cm)
    if first_cm is not None:
        pf.first_line_indent = Cm(first_cm)
    if tabs_cm:
        pf.tab_stops.clear_all()
        for _tc in tabs_cm:
            pf.tab_stops.add_tab_stop(Cm(_tc))
    if teks:
        add_rich_runs_to_para(para, teks, bold_default=bold, font_size=font_size)
    return para


def _identitas_hanging(doc, label, nilai, tab_cm=2.75):
    """Baris identitas (Nama/NPM/Konsentrasi/Judul Tesis) dengan titik dua lurus.

    Memakai indentasi menggantung: label di tepi kiri, isi mulai pada kolom
    `tab_cm` dan baris lanjutan judul tetap sejajar dengan kolom tersebut.
    """
    return _baris_kolom(doc, f"{label}\t:  {nilai}", tabs_cm=(tab_cm,), bold=True,
                        align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2, line_spacing=1.0,
                        left_cm=tab_cm, first_cm=-tab_cm)


def _lembar_pers_pembimbing(doc, judul, penulis, nim, konsentrasi):
    """Lembar Persetujuan Komisi Pembimbing — tata letak mengikuti contoh resmi
    Tridinanti: identitas dengan titik dua lurus (tab), blok pembimbing berkolom,
    blok Dekan/Kaprodi dua kolom. Tanpa hentian halaman di akhir."""
    h1(doc, "LEMBAR PERSETUJUAN KOMISI PEMBIMBING")
    p(doc, "Tesis ini Telah Disetujui Untuk Diujikan", bold=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_before=6, space_after=10,
      font_size=12, line_spacing=1.0)
    p(doc, judul, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
      space_after=10, font_size=14, line_spacing=1.3)
    _identitas_hanging(doc, "Nama", penulis)
    _identitas_hanging(doc, "NPM", nim)
    _identitas_hanging(doc, "Konsentrasi", konsentrasi)
    _identitas_hanging(doc, "Judul Tesis", _judul_rapi(judul))
    p(doc, "Menyetujui :", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
      space_before=10, space_after=8, font_size=12, line_spacing=1.0)
    _baris_kolom(doc, "Pembimbing Tesis", space_after=10)
    # Pembimbing I
    _baris_kolom(doc, "Tanggal ...................Pembimbing I\t: ...................................................",
                 tabs_cm=(7.0,), space_after=0)
    _baris_kolom(doc, "\tDr. Sari Sakarina, SE.,MM.,CHRM", bold=False, tabs_cm=(7.0,), space_after=0)
    _baris_kolom(doc, "\tNIDN : 0214038501", bold=False, tabs_cm=(7.0,), space_after=16)
    # Pembimbing II
    _baris_kolom(doc, "Tanggal ...................Pembimbing II\t: ...................................................",
                 tabs_cm=(7.0,), space_after=0)
    _baris_kolom(doc, "\tDr. Yuni Adinda Putri, S.E., M.Si", bold=False, tabs_cm=(7.0,), space_after=0)
    _baris_kolom(doc, "\tNIDN : 0207069401", bold=False, tabs_cm=(7.0,), space_after=18)
    p(doc, "Mengetahui :", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
      space_after=8, font_size=12, line_spacing=1.0)
    _baris_kolom(doc, "Dekan  FEB\tKaprodi MM", bold=False, tabs_cm=(7.8,), space_after=4)
    _baris_kolom(doc, "Dr. Msy. Mikial, SE., M.Si., Ak.,CA.,CSRS\tDr. Sari Sakarina, SE.,MM.,CHRM",
                 bold=False, tabs_cm=(7.8,), space_after=0)
    _baris_kolom(doc, "NIDN : 0205026401\tNIDN : 0214038501",
                 bold=False, tabs_cm=(7.8,), space_after=26)
    _baris_kolom(doc, "Tanggal ...........................\tTanggal ...........................",
                 tabs_cm=(7.8,), space_after=0)


def _lembar_pers_penguji(doc, judul, penulis, nim, konsentrasi, tahun):
    """Lembar Persetujuan Komisi Penguji Tesis — tata letak mengikuti contoh
    resmi Tridinanti. Tanpa hentian halaman di akhir."""
    h1(doc, "LEMBAR PERSETUJUAN KOMISI PENGUJI TESIS")
    p(doc, judul, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
      space_before=6, space_after=8, font_size=14, line_spacing=1.3)
    p(doc, "Dipertahankan di depan Komisi Penguji Tesis\nProgram Studi Magister Manajemen Fakultas Ekonomi dan Bisnis\nUniversitas Tridinanti",
      bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=10,
      font_size=12, line_spacing=1.15)
    _identitas_hanging(doc, "Nama", penulis)
    _identitas_hanging(doc, "NPM", nim)
    _identitas_hanging(doc, "Konsentrasi", konsentrasi)
    p(doc, "Menyetujui :", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
      space_before=10, space_after=12, font_size=12, line_spacing=1.0)
    # Ketua Penguji
    _baris_kolom(doc, "Ketua Penguji :\tTanggal\tTanda Tangan", tabs_cm=(7.0, 11.0), space_after=0)
    _baris_kolom(doc, "\t.......................\t.......................",
                 bold=False, tabs_cm=(7.0, 11.0), space_after=4)
    _baris_kolom(doc, "NIDN :", bold=False, tabs_cm=(7.0, 11.0), space_after=16)
    # Anggota Penguji
    _baris_kolom(doc, "Anggota Penguji :", tabs_cm=(7.0, 11.0), space_after=8)
    _baris_kolom(doc, "1. Dr. Sari Sakarina, SE.,MM. CHRM\t.......................\t.......................",
                 bold=False, tabs_cm=(7.0, 11.0), space_after=0)
    _baris_kolom(doc, "NIDN : 0214038501", bold=False, tabs_cm=(7.0, 11.0), space_after=12)
    _baris_kolom(doc, "2. Dr. Yuni Adinda Putri, S.E., M.Si\t.......................\t.......................",
                 bold=False, tabs_cm=(7.0, 11.0), space_after=0)
    _baris_kolom(doc, "NIDN : 0207069401", bold=False, tabs_cm=(7.0, 11.0), space_after=18)
    p(doc, "Mengetahui :", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
      space_after=8, font_size=12, line_spacing=1.0)
    _baris_kolom(doc, f"Palembang, {tahun}\tPalembang, {tahun}",
                 bold=False, tabs_cm=(7.8,), space_after=4)
    _baris_kolom(doc, "Dekan  FEB\tKaprodi MM", bold=False, tabs_cm=(7.8,), space_after=4)
    _baris_kolom(doc, "Dr. Msy. Mikial, SE., M.Si., Ak.,CA.,CSRS\tDr. Sari Sakarina, SE.,MM.,CHRM",
                 bold=False, tabs_cm=(7.8,), space_after=0)
    _baris_kolom(doc, "NIDN : 0205026401\tNIDN : 0214038501",
                 bold=False, tabs_cm=(7.8,), space_after=0)


def _abstrak_id(doc, data_abstrak=None):
    """Halaman ABSTRAK Indonesia.
    Bila analisis Bab IV sudah tersedia, memuat rangkuman tujuan, metode,
    angka koefisien jalur PLS-SEM, mediasi, R-square, dan kata kunci.
    """
    h1(doc, "ABSTRAK")
    if data_abstrak and data_abstrak.get("id"):
        for par in data_abstrak["id"]:
            p(doc, par, indent=0.5, space_before=4, space_after=6, line_spacing=1.15)
        if data_abstrak.get("kata_kunci"):
            p(doc, f"Kata Kunci: {data_abstrak['kata_kunci']}", bold=False, italic=False,
              indent=0, space_before=8, space_after=10, line_spacing=1.15)
    else:
        p(doc, "(Silakan isi abstrak di sini setelah Bab IV dan Bab V selesai — memuat tujuan, "
               "metode, hasil analisis beserta angka, dan kesimpulan.)",
          italic=True, indent=0.5, space_before=6, space_after=10, line_spacing=1.5)


def _abstrak_en(doc, data_abstrak=None):
    """Halaman ABSTRACT Inggris.
    Versi bahasa Inggris resmi dari abstrak penelitian (seluruh teks dicetak miring).
    """
    h1(doc, "ABSTRACT")
    if data_abstrak and data_abstrak.get("en"):
        for par in data_abstrak["en"]:
            p(doc, par, italic=True, indent=0.5, space_before=4, space_after=6, line_spacing=1.15)
        if data_abstrak.get("keywords"):
            p(doc, f"Keywords: {data_abstrak['keywords']}", bold=False, italic=True,
              indent=0, space_before=8, space_after=10, line_spacing=1.15)
    else:
        p(doc, "(Please write the abstract here after Chapter IV and Chapter V are completed — "
               "objective, method, analysis results with figures, and conclusions.)",
          italic=True, indent=0.5, space_before=6, space_after=10, line_spacing=1.5)


def _surat_pernyataan(doc, penulis, nim, judul):
    """Surat pernyataan mahasiswa — tata letak mengikuti contoh resmi Tridinanti:
    rata kiri, kolom titik dua lurus, butir bernomor menggantung, blok tanda tangan
    rata kanan bawah."""
    h1(doc, "SURAT PERNYATAAN")
    p(doc, "Saya yang bertanda tangan di bawah ini:", align=WD_ALIGN_PARAGRAPH.LEFT,
      indent=0, space_before=4, space_after=8, font_size=12, line_spacing=1.3)
    _identitas_hanging(doc, "Nama", penulis)
    _identitas_hanging(doc, "NPM", nim)
    _identitas_hanging(doc, "Program studi", "Magister Manajemen Universitas Tridinanti Palembang")
    _identitas_hanging(doc, "Judul tesis", _judul_rapi(judul))
    p(doc, "Dengan ini menyatakan dengan sesungguhnya bahwa:", align=WD_ALIGN_PARAGRAPH.LEFT,
      indent=0, space_before=10, space_after=8, font_size=12, line_spacing=1.3)
    _butir = [
        "Tesis yang saya susun sebagai salah satu syarat untuk memperoleh gelar Magister "
        "Manajemen (MM) dari Fakultas Ekonomi & Bisnis Program Studi Magister Manajemen "
        "Universitas Tridinanti Palembang seluruhnya adalah hasil karya saya sendiri.",
        "Apabila ada bagian tertentu dalam penulisan tesis yang saya kutip dari hasil karya "
        "orang lain, akan saya tulis sumbernya dengan jelas, sesuai norma-norma ilmiah yang berlaku.",
        "Jika dikemudian hari ditemukan salah satu atau sebagian tesis yang saya susun ini bukan "
        "hasil karya saya sendiri atau ada plagiat dalam bagian-bagian tertentu, saya bersedia "
        "menerima sanksi pencabutan gelar Akademik Magister Manajemen (MM) yang saya terima dan "
        "sanksi-sanksi lainnya sesuai dengan peraturan perundang-undangan yang berlaku.",
    ]
    for _i, _t in enumerate(_butir, 1):
        _baris_kolom(doc, f"{_i}.\t{_t}", bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                     tabs_cm=(1.0,), left_cm=1.0, first_cm=-1.0,
                     space_after=6, line_spacing=1.3)
    p(doc, "Demikianlah pernyataan ini saya buat dengan sesungguhnya dalam keadaan sadar dengan "
           "rasa tanggung jawab tanpa paksaan dari pihak manapun.",
      align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_before=8, space_after=0,
      font_size=12, line_spacing=1.3)
    _ttd = p(doc, "Palembang, ............................., 2026\nYang menyatakan,\n\n"
                  "(Materai Rp 10.000)\n\n\n" + str(penulis),
             align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_before=18, space_after=0,
             font_size=12, line_spacing=1.3)
    _ttd.paragraph_format.left_indent = Cm(7.6)


def _riwayat_hidup(doc, penulis, nim, konsentrasi, tahun):
    """Daftar riwayat hidup — rata kiri, kolom titik dua lurus, garis titik seragam,
    muat dalam satu halaman."""
    h1(doc, "DAFTAR RIWAYAT HIDUP")
    p(doc, "DATA PRIBADI", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0,
      space_before=6, space_after=4, font_size=12, line_spacing=1.0)
    _identitas_hanging(doc, "Nama", penulis.title(), tab_cm=4.6)
    _identitas_hanging(doc, "NPM", nim, tab_cm=4.6)
    _identitas_hanging(doc, "Konsentrasi", konsentrasi, tab_cm=4.6)
    _identitas_hanging(doc, "Program Studi", "Magister Manajemen", tab_cm=4.6)
    _identitas_hanging(doc, "Fakultas", "Ekonomi dan Bisnis", tab_cm=4.6)
    _identitas_hanging(doc, "Universitas", "Universitas Tridinanti", tab_cm=4.6)
    _identitas_hanging(doc, "Tempat/Tanggal Lahir",
                       "..................................................................", tab_cm=4.6)
    _identitas_hanging(doc, "Alamat",
                       "..................................................................", tab_cm=4.6)
    p(doc, "RIWAYAT PENDIDIKAN", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0,
      space_before=10, space_after=4, font_size=12, line_spacing=1.0)
    for _tgk in ("SD", "SMP", "SMA/SMK", "S-1", "S-2"):
        _nilai = ("Universitas Tridinanti Palembang\t" if _tgk == "S-2"
                  else "............................................\t")
        _baris_kolom(doc, f"{_tgk}\t: {_nilai}Tahun : ..............",
                     tabs_cm=(2.75, 10.2), space_after=2, line_spacing=1.0)
    p(doc, "RIWAYAT PEKERJAAN", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0,
      space_before=10, space_after=4, font_size=12, line_spacing=1.0)
    for _ in range(2):
        _baris_kolom(doc, "........................................................................................\t"
                          "Tahun : ..............",
                     tabs_cm=(10.2,), space_after=2, line_spacing=1.0)
    p(doc, "KARYA ILMIAH / PENELITIAN", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0,
      space_before=10, space_after=4, font_size=12, line_spacing=1.0)
    _baris_kolom(doc, "......................................................................................",
                 space_after=2, line_spacing=1.0)
    _ttd = p(doc, f"Palembang, ............................., {tahun}\n\n\n\n{penulis.title()}",
             align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_before=22, space_after=0,
             font_size=12, line_spacing=1.3)
    _ttd.paragraph_format.left_indent = Cm(7.6)


def build_bagian_awal(doc, config=None):
    if config is None:
        config = {}
    judul = config.get("judul", "PENGARUH GAYA KEPEMIMPINAN DAN BUDAYA ORGANISASI TERHADAP DISIPLIN KERJA SERTA DAMPAKNYA PADA KINERJA PEGAWAI BANK PEMBANGUNAN DAERAH SUMATERA SELATAN DAN BANGKA BELITUNG CABANG PEMBANTU PARAMESWARA").strip().upper()
    penulis = config.get("penulis", "RUSMIANITA").strip().upper()
    nim = config.get("nim", "255441029").strip()
    konsentrasi = config.get("konsentrasi", "Manajemen Sumber Daya Manusia").strip()
    instansi = config.get("instansi", "Bank Pembangunan Daerah Sumatera Selatan dan Bangka Belitung Cabang Pembantu Parameswara").strip()
    tahun = config.get("tahun", "2026").strip()
    x1 = config.get("x1", "Gaya Kepemimpinan").strip()
    x2 = config.get("x2", "Budaya Organisasi").strip()
    m = config.get("m", "Disiplin Kerja").strip()
    y = config.get("y", "Kinerja Pegawai").strip()

    # ---------------- HALAMAN JUDUL (SAMPUL) ----------------
    # Tata letak mengikuti sampul resmi: Judul – TESIS – subjudul – LOGO universitas
    # – penyusun – identitas program studi. Semua baris rata tengah tanpa indentasi
    # dan dirapatkan agar seluruh sampul muat dalam SATU halaman.
    p(doc, judul, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
      space_before=0, space_after=10, font_size=14, line_spacing=1.5)
    # Label "HALAMAN JUDUL" tetap ada (1 pt, putih) agar tercatat di DAFTAR ISI
    # otomatis, tetapi tidak terlihat pada halaman sampul.
    _penanda = p(doc, "HALAMAN JUDUL", align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
                 space_before=0, space_after=0, font_size=1, line_spacing=1.0)
    for _r in _penanda.runs:
        _r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    _set_outline_level(_penanda, 0)
    p(doc, "TESIS", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
      space_before=0, space_after=8, font_size=12, line_spacing=1.0)
    p(doc, "Ditulis untuk memenuhi Sebagian Persyaratan dalam\nmenempuh Gelar Magister Manajemen",
      align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=0, font_size=12, line_spacing=1.15)
    _sisip_logo_cover(doc, lebar_in=2.4)
    p(doc, "Disusun  Oleh :", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
      space_before=16, space_after=6, font_size=12, line_spacing=1.0)
    _baris_identitas_cover(doc, "Nama", penulis)
    _baris_identitas_cover(doc, "NPM", nim)
    _baris_identitas_cover(doc, "Konsentrasi", konsentrasi)
    p(doc, f"PROGRAM STUDI MAGISTER MANAJEMEN\nFAKULTAS EKONOMI DAN BISNIS\nUNIVERSITAS TRIDINANTI\n{tahun}",
      bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
      space_before=40, space_after=0, font_size=14, line_spacing=1.3)
    doc.add_page_break()

    # ---------------- LEMBAR PERSETUJUAN KOMISI PEMBIMBING ----------------
    _lembar_pers_pembimbing(doc, judul, penulis, nim, konsentrasi)
    doc.add_page_break()

    # ---------------- LEMBAR PERSETUJUAN KOMISI PENGUJI TESIS ----------------
    _lembar_pers_penguji(doc, judul, penulis, nim, konsentrasi, tahun)
    doc.add_page_break()

    # Kata Pengantar
    h1(doc, "KATA PENGANTAR")
    p(doc, f"Puji dan Syukur kepada Tuhan Yang Maha Esa, atas limpahan dan Hidayah-Nya sehingga Tesis yang Berjudul “{_judul_rapi(judul)}”.", indent=0.5, space_after=6)
    p(doc, "Tesis ini ditulis untuk memenuhi sebagian persyaratan mendapatkan gelar Magister Manajemen pada Universitas Tridinanti Palembang Program Studi Magister Manajemen. Penyelesaian Tesis ini melibatkan banyak pihak yang telah memberikan bantuan, baik langsung maupun tidak langsung, moril maupun materil. Oleh karena itu pada kesempatan ini penulis menyampaikan penghargaan dan rasa terima kasih kepada:", indent=0.5, space_after=6)
    points = [
        "1. Ketua Yayasan Pendidikan Nasional Tridinanti Palembang Bapak Prof.H. Mahmud Hasjim, MME.",
        "2. Rektor Universitas Tridinanti Palembang Bapak Prof. Dr. Ir. H. Edizal AE, MS.",
        "3. Dekan Fakultas Ekonomi & Bisnis Universitas Tridinanti Palembang Ibu Dr. Msy. Mikial, SE., M.Si., Ak.,CA.,CSRS.",
        "4. Wakil Dekan I Fakultas Ekonomi & Bisnis Universitas Tridinanti Palembang Ibu Dr. M. Ima Andriyani, SE., M.Si.",
        "5. Direktur Program Pascasarjana Magister Manajemen Universitas Tridinanti Palembang Bapak Dr. Djatmiko Noviantoro, S.E., M.Si.",
        "6. Kaprodi Program Pascasarjana Magister Manajemen Universitas Tridinanti Palembang Ibu Dr. Sari Sakarina, SE., MM., CHRM, sekaligus selaku Pembimbing I yang senantiasa meluangkan waktu, pikiran dan tenaganya untuk memberikan bimbingan dan arahan dalam menyelesaikan Tesis ini.",
        "7. Ibu Dr. Yuni Adinda Putri, S.E., M.Si, selaku Pembimbing II yang senantiasa meluangkan waktu, pikiran dan tenaganya untuk memberikan bimbingan dan arahan dalam menyelesaikan Tesis ini.",
        f"8. Pimpinan dan segenap pegawai pada {instansi} yang telah bersedia meluangkan waktu dan membantu proses pengumpulan data penelitian ini.",
        "9. Kedua Orang tua tercinta, bapak dan ibu atas doa dan kasih sayangnya yang tulus.",
        "10. Saudara, keluarga dan rekan-rekan yang senantiasa memberikan dukungan dan motivasi moril.",
        "11. Rekan-rekan mahasiswa Angkatan MM Universitas Tridinanti yang senantiasa saling menguatkan.",
        "12. Semua Pihak yang telah membimbing, membantu dan mendorong penyelesaian Tesis ini."
    ]
    for pt in points:
        p(doc, pt, indent=0.25, space_after=2)
    p(doc, f"Penulis menyadari mungkin dalam Tesis ini masih terdapat kekurangan, untuk itu diharapkan tanggapan dan masukan dari berbagai pihak sebagai bahan perbaikan dengan harapan Tesis ini dapat bermanfaat bagi pengembangan ilmu di bidang {konsentrasi}.", indent=0.5, space_before=6, space_after=12)
    p(doc, f"Palembang, {tahun}\n\n\n{penulis}", align=WD_ALIGN_PARAGRAPH.RIGHT)
    doc.add_page_break()

    # Abstrak Indonesia & Abstract Inggris — menggunakan angka riil PLS-SEM
    # bila analisis Bab IV/V telah tersedia.
    data_abstrak = _abstrak_otomatis(config)
    _abstrak_id(doc, data_abstrak)
    doc.add_page_break()

    _abstrak_en(doc, data_abstrak)
    doc.add_page_break()

    # Surat Pernyataan
    _surat_pernyataan(doc, penulis, nim, judul)
    doc.add_page_break()

    # Riwayat Hidup
    _riwayat_hidup(doc, penulis, nim, konsentrasi, tahun)
    doc.add_page_break()

    # Daftar Isi — FIELD OTOMATIS (Word yang menghitung judul & nomor halaman)
    # Daftar isi statis dihapus: dulu isinya ditulis manual dengan NOMOR HALAMAN
    # ESTIMASI, sehingga meleset. Sekarang memakai field TOC yang mengumpulkan
    # seluruh judul bertingkat pada dokumen (outline level h1/h2/h3).
    h1(doc, "DAFTAR ISI")
    _toc_field(doc)
    doc.add_page_break()

    # Daftar Tabel — FIELD OTOMATIS (mengumpulkan caption tabel: tingkat outline 4)
    h1(doc, "DAFTAR TABEL")
    _toc_field(doc, switch=r'TOC \o "4-4" \u \h \z \* CHARFORMAT', charformat=True)
    doc.add_page_break()

    # Daftar Gambar — FIELD OTOMATIS (caption gambar: tingkat outline 5)
    h1(doc, "DAFTAR GAMBAR")
    _toc_field(doc, switch=r'TOC \o "5-5" \u \h \z \* CHARFORMAT', charformat=True)
    doc.add_page_break()

    # Daftar Lampiran — FIELD OTOMATIS (judul lampiran: tingkat outline 6)
    h1(doc, "DAFTAR LAMPIRAN")
    _toc_field(doc, switch=r'TOC \o "6-6" \u \h \z \* CHARFORMAT', charformat=True)

    # Nyalakan "update fields saat dibuka" -> semua daftar otomatis langsung terhitung.
    _set_update_fields(doc)
    doc.add_page_break()

print("Bagian awal siap.")
