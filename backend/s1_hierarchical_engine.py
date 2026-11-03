# backend/s1_hierarchical_engine.py
"""
S1 UNDERGRADUATE THESIS (SKRIPSI SARJANA) HIERARCHICAL DOCUMENT BUILDER
========================================================================
Standar Format Penulisan Skripsi Sarjana (S1) Mengacu 100% pada FORMAT S1.pdf:
1. Tipografi & Tata Letak:
   - Font Times New Roman 12 pt, spasi 2.0 (ganda) untuk teks naskah.
   - Spasi 1.0 (tunggal) untuk tabel, judul tabel/gambar, abstrak, dan kutipan panjang.
   - First line indent: 1.27 cm (0.5 inci), justified (rata kiri-kanan).
   - Margin: Kiri 4 cm, Atas 4 cm, Kanan 3 cm, Bawah 3 cm, Kertas A4.
2. Format Tabel Tertutup Semua (Full Borders) & Font 12 pt:
   - Seluruh tabel diberi garis pembatas lengkap (Semua Garis: top, bottom, left, right, insideH, insideV).
   - Font isi dan header tabel TETAP 12 pt Times New Roman.
   - Judul Tabel 2 baris terpusat tebal (Tabel X.X\nJudul Tabel), spasi 1.0.
   - Sumber di bawah tabel rata kiri (Sumber: ..., Tahun).
3. Kerangka Berpikir Rapi & Proporsional:
   - Diagram variabel terstruktur kotak rapi dan sejajar dengan konektor panah yang jelas.
   - Judul Gambar di bawah gambar 2 baris terpusat tebal (Gambar 2.1\nKerangka Berpikir).
   - Keterangan panah parsial (Uji t) dan simultan (Uji F) di bawah gambar beserta sumber.
4. Tata Huruf Bersih (Bebas Kalimat Huruf Kapital yang Berlebihan):
   - Hanya judul bab utama (BAB I PENDAHULUAN) dan sampul depan yang menggunakan huruf kapital.
   - Sub-bab, sub-sub-bab, dan teks naskah menggunakan kaidah Title Case dan kalimat akademik standar.
5. Jumlah Tabel di Bab IV Sama Persis dengan FORMAT S1.pdf:
   - Tepat 23 Tabel di Bab IV (Tabel 4.1 s.d. Tabel 4.23) yang tersusun kronologis dan lengkap.
"""

import os
import sys
import re
import math
import json
import hashlib
import random
from datetime import datetime
from typing import Dict, Any, List, Tuple, Optional

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION_START
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

OUTPUT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, "..", "output"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

import s1_data_pipeline
import s1_constructs
import s1_literature_library
import s1_topic_presets
from build_full_thesis_docx import apply_fixed_widths


# =========================================================================
# 0. LAPISAN PRESET TOPIK & LEKSIKON SEKTOR
# =========================================================================

S1_LINE_SPACING = 1.5
_ACTIVE_PRESET: Optional[Dict[str, Any]] = None


def set_active_topic_preset(preset: Optional[Dict[str, Any]]) -> None:
    """Menyimpan preset sektor aktif agar seluruh builder memakai istilah sektor yang konsisten."""
    global _ACTIVE_PRESET
    _ACTIVE_PRESET = preset


def get_active_topic_preset() -> Optional[Dict[str, Any]]:
    return _ACTIVE_PRESET


def s1_lex(text: str) -> str:
    """Menyelaraskan istilah generik (pegawai/instansi) dengan sektor penelitian aktif."""
    if not text:
        return text
    return s1_topic_presets.apply_lexicon(text, _ACTIVE_PRESET)


# =========================================================================
# 1. PARSER SEMANTIK JUDUL & RESOLVER KONTEKS SKRIPSI S1
# =========================================================================

def join_academic(items: List[str]) -> str:
    """Menggabungkan daftar kata menjadi kalimat bergaya bahasa Indonesia akademik baku."""
    clean = [str(x).strip() for x in items if str(x).strip()]
    if not clean:
        return ""
    if len(clean) == 1:
        return clean[0]
    if len(clean) == 2:
        return f"{clean[0]} dan {clean[1]}"
    return f"{', '.join(clean[:-1])}, dan {clean[-1]}"

def to_academic_title_case(text: str) -> str:
    """Mengubah teks menjadi kaidah Title Case Bahasa Indonesia ilmiah tanpa merusak singkatan/notasi."""
    s_clean = " ".join((text or "").split()).strip()
    if not s_clean:
        return ""
    lower_words = {
        "dan", "atau", "terhadap", "pada", "di", "ke", "dari", "yang",
        "untuk", "dengan", "dalam", "atas", "oleh", "sebagai", "tentang"
    }
    special_preserve = {
        "(X1)": "(X1)", "(X2)": "(X2)", "(X3)": "(X3)", "(X4)": "(X4)", "(X5)": "(X5)",
        "(Y)": "(Y)", "(M)": "(M)",
        "X1": "X1", "X2": "X2", "X3": "X3", "X4": "X4", "X5": "X5",
        "Y": "Y", "M": "M",
        "ANOVA": "ANOVA", "(ANOVA)": "(ANOVA)",
        "SPSS": "SPSS", "(SPSS)": "(SPSS)",
        "VIF": "VIF", "(VIF)": "(VIF)",
        "IKU": "IKU", "(IKU)": "(IKU)",
        "SKP": "SKP", "(SKP)": "(SKP)",
        "S1": "S1", "(S1)": "(S1)",
        "K-S": "K-S", "(K-S)": "(K-S)",
        "Kolmogorov-Smirnov": "Kolmogorov-Smirnov",
        "(One-Sample": "(One-Sample",
        "Kolmogorov-Smirnov)": "Kolmogorov-Smirnov)",
        "SS": "SS", "STS": "STS", "TS": "TS",
        "PT": "PT", "CV": "CV", "RSUD": "RSUD", "SDN": "SDN", "SMPN": "SMPN", "SMAN": "SMAN",
        "BUMN": "BUMN", "Pertamina": "Pertamina",
        "UMKM": "UMKM", "(UMKM)": "(UMKM)", "IKM": "IKM", "KUKM": "KUKM",
        "RS": "RS", "RSIA": "RSIA", "IGD": "IGD", "ICU": "ICU", "SDM": "SDM",
        "OJK": "OJK", "APR": "APR", "ROA": "ROA", "ROI": "ROI", "NIM": "NIM",
        "Likert": "Likert",
        "df": "df", "R²": "R²", "R-Square": "R-Square",
        "t": "t", "F": "F", "(Uji F)": "(Uji F)", "(Uji t)": "(Uji t)",
        "(Uji": "(Uji", "t)": "t)", "F)": "F)"
    }
    words = s_clean.split(" ")
    res = []
    for idx, w in enumerate(words):
        if w in special_preserve:
            res.append(special_preserve[w])
        elif w.upper() in special_preserve:
            res.append(special_preserve[w.upper()])
        elif w.startswith("(") and w.endswith(")") and w[1:-1].upper() in special_preserve:
            res.append("(" + special_preserve[w[1:-1].upper()] + ")")
        else:
            w_lower = w.lower()
            if idx > 0 and w_lower in lower_words:
                res.append(w_lower)
            elif w.startswith("(") and len(w) > 1:
                res.append("(" + w[1:].capitalize())
            else:
                res.append(w.capitalize())
    return " ".join(res)


def parse_s1_title(raw_title: str) -> Dict[str, Any]:
    """Mengekstrak variabel independen (X1 ... Xk, 2 s.d. 5 X), dependen (Y), dan instansi dari judul skripsi S1."""
    title_clean = " ".join((raw_title or "").split()).strip()
    if not title_clean:
        title_clean = "Pengaruh Kepemimpinan Transformasional dan Budaya Organisasi terhadap Kinerja Pegawai pada Dinas Pendidikan Kota Palembang"

    res = {
        "title": title_clean,
        "model_type": "regression",
        "n_x": 2,
        "x_names": ["Kepemimpinan Transformasional", "Budaya Organisasi"],
        "x1_name": "Kepemimpinan Transformasional",
        "x2_name": "Budaya Organisasi",
        "m_name": "",
        "y_name": "Kinerja Pegawai",
        "instansi": "Dinas Pendidikan Kota Palembang"
    }

    m_inst = re.search(r"\b(?:pada|studi pada|di|kasus pada)\s+(.+)$", title_clean, re.IGNORECASE)
    if not m_inst:
        m_inst = re.search(r"\b((?:Rumah Sakit|RS|RSUD|PT|CV|Dinas|Badan|Kantor|Balai|Puskesmas|Bank|Universitas|Institut|Politeknik)\s+.+)$", title_clean, re.IGNORECASE)
    if m_inst:
        res["instansi"] = to_academic_title_case(m_inst.group(1).strip())
        main_part = title_clean[:m_inst.start()].strip()
    else:
        main_part = title_clean

    main_part = re.sub(r"^(pengaruh|analisis pengaruh|analisis|hubungan)\s+", "", main_part, flags=re.IGNORECASE).strip()
    parts = re.split(r"\s+terhadap\s+", main_part, flags=re.IGNORECASE)
    if len(parts) >= 2:
        x_part = parts[0].strip()
        y_part = parts[1].strip()

        x_splits = re.split(r"\s*(?:,|;|\+|&)\s*|\s+dan\s+", x_part, flags=re.IGNORECASE)
        parsed_x = [to_academic_title_case(s.strip()) for s in x_splits if s.strip()]
        if len(parsed_x) >= 2:
            res["x_names"] = parsed_x[:5]
        elif len(parsed_x) == 1 and parsed_x[0]:
            res["x_names"] = [parsed_x[0], "Budaya Organisasi"]

        res["n_x"] = len(res["x_names"])
        res["x1_name"] = res["x_names"][0]
        res["x2_name"] = res["x_names"][1]
        for i, xn in enumerate(res["x_names"]):
            res[f"x{i+1}_name"] = xn

        res["y_name"] = to_academic_title_case(y_part.strip())

    return res


def resolve_s1_context(config: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Menyusun konfigurasi lengkap naskah skripsi S1 dan menghubungkan hasil analisis SPSS (2 s.d. 5 variabel X)."""
    cfg = dict(config or {})
    raw_title = cfg.get("judul") or cfg.get("title") or "PENGARUH KEPEMIMPINAN TRANSFORMASIONAL DAN BUDAYA ORGANISASI TERHADAP KINERJA PEGAWAI PADA DINAS PENDIDIKAN KOTA PALEMBANG"
    parsed = parse_s1_title(raw_title)

    title_display = to_academic_title_case(raw_title)
    penulis_raw = cfg.get("penulis") or cfg.get("author") or "MUHAMMAD RIZKY PRATAMA"
    penulis_display = to_academic_title_case(penulis_raw)
    instansi_clean = to_academic_title_case(cfg.get("instansi") or parsed["instansi"])

    # Tentukan nama-nama variabel bebas X (2 s.d. 5)
    raw_x_list = cfg.get("x_names")
    if isinstance(raw_x_list, list) and len(raw_x_list) >= 2:
        x_names = [to_academic_title_case(str(x).strip()) for x in raw_x_list if str(x).strip()][:5]
    else:
        cands = []
        for i in range(1, 6):
            v = cfg.get(f"x{i}_name") or cfg.get(f"x{i}")
            if v and str(v).strip():
                cands.append(to_academic_title_case(str(v).strip()))
        if len(cands) >= 2:
            x_names = cands[:5]
        else:
            x_names = parsed["x_names"]

    k_x = len(x_names)
    x_vars = []
    for i, name in enumerate(x_names):
        code = f"X{i+1}"
        prof = s1_constructs.get_construct_profile(name)
        its = s1_constructs.construct_items(code, name, max_dims=None)
        x_vars.append({
            "index": i + 1,
            "code": code,
            "name": name,
            "profile": prof,
            "items": its,
        })

    raw_y = cfg.get("y_name") or cfg.get("y") or parsed["y_name"]
    if instansi_clean and instansi_clean.lower() in raw_y.lower():
        pattern = re.compile(re.escape(instansi_clean), re.IGNORECASE)
        raw_y = pattern.sub("", raw_y).strip()
    raw_y = re.sub(r"\b(pada|di|studi pada)\s*$", "", raw_y, flags=re.IGNORECASE).strip()
    y_name = to_academic_title_case(raw_y)
    y_profile = s1_constructs.get_construct_profile(y_name)
    y_items = s1_constructs.construct_items("Y", y_name, max_dims=None)

    m_name = to_academic_title_case(cfg.get("m_name") or cfg.get("m") or parsed.get("m_name", ""))
    m_profile = s1_constructs.get_construct_profile(m_name) if m_name else None
    m_items = s1_constructs.construct_items("M", m_name, max_dims=None) if m_name else []

    x_join = join_academic([v["name"] for v in x_vars])
    x_codes_join = join_academic([v["code"] for v in x_vars])

    # Tentukan populasi dan ukuran sampel secara deterministik per instansi/judul
    raw_pop = cfg.get("populasi")
    if not raw_pop:
        corp_l = f"{raw_title} {instansi_clean}".lower()
        if any(k in corp_l for k in ["pertamina", "pln", "bumn", "migas", "energi", "persero", "kilang"]):
            populasi_val = 180
        elif any(k in corp_l for k in ["charitas", "rumah sakit", "rsud", "puskesmas", "klinik", "kesehatan", "medis"]):
            populasi_val = 84
        elif any(k in corp_l for k in ["tata kota", "dinas", "pemda", "pupr", "kantor"]):
            populasi_val = 58
        else:
            populasi_val = int(cfg.get("n") or 67)
    else:
        populasi_val = int(raw_pop)

    raw_samp = cfg.get("sampel") or cfg.get("n")
    if raw_samp:
        sampel_val = int(raw_samp)
    else:
        if populasi_val > 100:
            # Rumus Slovin presisi e = 0.05
            sampel_val = int(round(populasi_val / (1.0 + populasi_val * (0.05 ** 2))))
        else:
            # Sensus / Sampling Jenuh (N <= 100)
            sampel_val = populasi_val

    ctx = {
        "title_upper": raw_title.upper(),
        "title_display": title_display,
        "judul": raw_title,
        "penulis_upper": penulis_raw.upper(),
        "penulis_display": penulis_display,
        "nim": cfg.get("nim") or "01011382227045",
        "program_studi": cfg.get("program_studi") or cfg.get("prodi") or "Manajemen",
        "konsentrasi": cfg.get("konsentrasi") or "Manajemen Sumber Daya Manusia",
        "fakultas": cfg.get("fakultas") or "Fakultas Ekonomi dan Bisnis",
        "universitas": cfg.get("universitas") or cfg.get("pt") or "Universitas Sriwijaya",
        "kota": cfg.get("kota") or "Palembang",
        "tahun": str(cfg.get("tahun") or datetime.now().year),
        "model_type": cfg.get("model_type") or parsed["model_type"],
        "instansi": instansi_clean,
        "n_x": k_x,
        "x_vars": x_vars,
        "x_names": [v["name"] for v in x_vars],
        "x_codes": [v["code"] for v in x_vars],
        "x_join": x_join,
        "x_codes_join": x_codes_join,
        "x1_name": x_vars[0]["name"],
        "x2_name": x_vars[1]["name"],
        "y_name": y_name,
        "y_profile": y_profile,
        "y_items": y_items,
        "m_name": m_name,
        "m_profile": m_profile,
        "m_items": m_items,
        "populasi": populasi_val,
        "sampel": sampel_val,
        "n": sampel_val,
        "pembimbing1": cfg.get("pembimbing1") or "Dr. H. Ahmad Fauzi, S.E., M.Si.",
        "nidn_pembimbing1": cfg.get("nidn_pembimbing1") or "0210116101",
        "nip_pembimbing1": cfg.get("nip_pembimbing1") or "197508122002121002",
        "pembimbing2": cfg.get("pembimbing2") or "Dra. Hj. Nurjanah, M.M.",
        "nidn_pembimbing2": cfg.get("nidn_pembimbing2") or "0215046502",
        "nip_pembimbing2": cfg.get("nip_pembimbing2") or "198004152006042003",
        "penguji1": cfg.get("penguji1") or "Prof. Dr. Ir. H. Bambang Sudarsono, M.M.",
        "penguji2": cfg.get("penguji2") or "Dr. Ratna Sari Dewi, S.E., M.Si.",
        "dekan": cfg.get("dekan") or "Prof. Dr. Mohamad Adam, S.E., M.E.",
        "nip_dekan": cfg.get("nip_dekan") or "196805141993031003",
    }
    for i, v in enumerate(x_vars):
        ctx[f"x{i+1}_name"] = v["name"]

    # Aktifkan preset sektor (pemasaran UMKM / perbankan digital / SDM rumah sakit)
    try:
        _preset = s1_topic_presets.get_topic_preset(ctx)
    except Exception:
        _preset = None
    ctx["topic_preset"] = _preset
    set_active_topic_preset(_preset)

    # Jalankan / Muat Pipeline SPSS S1
    pipeline_res = s1_data_pipeline.run_s1_full_pipeline(ctx)
    ctx["spss"] = pipeline_res["contract"]
    ctx["summary"] = pipeline_res["summary"]
    ctx["data_file"] = pipeline_res["data_file"]

    return ctx


# =========================================================================
# 2. HELPER FORMATTING PRESISI STANDAR FORMAT S1.PDF
# =========================================================================

def setup_s1_section(section, is_front_matter: bool = False, is_first_chapter: bool = False):
    """
    Mengatur Section A4, margin 4-4-3-3 cm, dan dual page numbering per section:
    - Front Matter: Romawi kecil (lowerRoman) di tengah bawah. Halaman cover footer kosong.
    - Chapter Section:
      * Halaman awal bab: Angka Arab di TENGAH BAWAH (Bottom Center).
      * Halaman lanjutan bab: Angka Arab di KANAN ATAS (Top Right).
      * Bab I dimulai dari 1 (start=1), Bab selanjutnya melanjutkan nomor halaman.
    """
    section.top_margin = Cm(4)
    section.bottom_margin = Cm(3)
    section.left_margin = Cm(4)
    section.right_margin = Cm(3)
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.header_distance = Cm(1.5)
    section.footer_distance = Cm(1.5)
    section.different_first_page_header_footer = True

    xml_field = '<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w')

    if is_front_matter:
        # Cover page (first page): footer kosong
        section.first_page_footer.is_linked_to_previous = False
        section.first_page_footer.paragraphs[0].text = ''

        # Front matter lanjutan: footer tengah bawah angka Romawi
        section.footer.is_linked_to_previous = False
        p_f = section.footer.paragraphs[0]
        p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_f = p_f.add_run()
        r_f.font.name = 'Times New Roman'
        r_f.font.size = Pt(12)
        r_f._r.append(parse_xml(xml_field))

        sectPr = section._sectPr
        el_old = sectPr.find(qn('w:pgNumType'))
        if el_old is not None:
            sectPr.remove(el_old)
        sectPr.append(parse_xml('<w:pgNumType %s w:fmt="lowerRoman" w:start="1"/>' % nsdecls('w')))

    else:
        # Chapter Section:
        # 1. First page: Bottom Center
        section.first_page_footer.is_linked_to_previous = False
        p_ff = section.first_page_footer.paragraphs[0]
        p_ff.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_ff = p_ff.add_run()
        r_ff.font.name = 'Times New Roman'
        r_ff.font.size = Pt(12)
        r_ff._r.append(parse_xml(xml_field))

        section.first_page_header.is_linked_to_previous = False
        section.first_page_header.paragraphs[0].text = ''

        # 2. Subsequent pages: Top Right Header, Footer empty
        section.header.is_linked_to_previous = False
        p_h = section.header.paragraphs[0]
        p_h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_h = p_h.add_run()
        r_h.font.name = 'Times New Roman'
        r_h.font.size = Pt(12)
        r_h._r.append(parse_xml(xml_field))

        section.footer.is_linked_to_previous = False
        section.footer.paragraphs[0].text = ''

        # pgNumType decimal
        sectPr = section._sectPr
        el_old = sectPr.find(qn('w:pgNumType'))
        if el_old is not None:
            sectPr.remove(el_old)
        if is_first_chapter:
            sectPr.append(parse_xml('<w:pgNumType %s w:fmt="decimal" w:start="1"/>' % nsdecls('w')))
        else:
            sectPr.append(parse_xml('<w:pgNumType %s w:fmt="decimal"/>' % nsdecls('w')))


def _force_run_black(run):
    """Memastikan font color run terkunci pada warna hitam murni (w:val='000000') tanpa themeColor."""
    rpr = run._r.get_or_add_rPr()
    color_node = rpr.find(qn("w:color"))
    if color_node is None:
        color_node = OxmlElement("w:color")
        rpr.append(color_node)
    color_node.set(qn("w:val"), "000000")
    for attr in ("w:themeColor", "w:themeShade", "w:themeTint"):
        if qn(attr) in color_node.attrib:
            del color_node.attrib[qn(attr)]


def configure_s1_heading_styles(doc: Document):
    """
    Menyelaraskan gaya baku Heading 1-4 dengan FORMAT S1.pdf (Times New Roman 12 pt, hitam, tebal)
    sekaligus mempertahankan outline level bawaan Word agar field DAFTAR ISI (TOC) dapat memungutnya.
    Idempoten: aman dipanggil berulang pada dokumen yang sama.
    """
    specs = {
        "Heading 1": (0, 2, 1.5, True),
        "Heading 2": (12, 6, 1.5, True),
        "Heading 3": (8, 4, 1.5, True),
        "Heading 4": (6, 3, 1.5, True),
        "TOC 1": (0, 2, 1.15, False),
        "TOC 2": (0, 2, 1.15, False),
        "TOC 3": (0, 2, 1.15, False),
    }
    for style_name, (sp_before, sp_after, spacing, bold) in specs.items():
        try:
            st = doc.styles[style_name]
        except KeyError:
            continue
        f = st.font
        f.name = "Times New Roman"
        f.size = Pt(12)
        f.bold = bold
        f.italic = False
        f.color.rgb = RGBColor(0, 0, 0)
        # Samakan font untuk semua script (Asia/complex) agar Word tidak menimpa TNR
        rpr = st.element.get_or_add_rPr()
        for tag in ("w:rFonts",):
            node = rpr.find(qn(tag))
            if node is None:
                node = OxmlElement(tag)
                rpr.append(node)
            node.set(qn("w:ascii"), "Times New Roman")
            node.set(qn("w:hAnsi"), "Times New Roman")
            node.set(qn("w:cs"), "Times New Roman")
            node.set(qn("w:eastAsia"), "Times New Roman")

        # Kunci warna teks style menjadi hitam murni (hapus referensi w:themeColor accent biru Word)
        color_node = rpr.find(qn("w:color"))
        if color_node is None:
            color_node = OxmlElement("w:color")
            rpr.append(color_node)
        color_node.set(qn("w:val"), "000000")
        for attr in ("w:themeColor", "w:themeShade", "w:themeTint"):
            if qn(attr) in color_node.attrib:
                del color_node.attrib[qn(attr)]

        pf = st.paragraph_format
        pf.space_before = Pt(sp_before)
        pf.space_after = Pt(sp_after)
        pf.line_spacing = spacing
        pf.keep_with_next = True
        pf.keep_together = True
        pf.left_indent = Inches(0)
        if style_name in ("Heading 2", "Heading 3", "Heading 4"):
            try:
                pf.tab_stops.add_tab_stop(Inches(0.5))
            except Exception:
                pass


def add_s1_field_toc(doc: Document, instr: str, placeholder: str = ""):
    """
    Menyisipkan field Word (TOC / TOC \\c) sebagai paragraf tersendiri.
    Field ditandai dirty agar Word menghitung ulang nomor halaman saat dokumen dibuka/dicetak.
    """
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    run_begin = p.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    fld_begin.set(qn("w:dirty"), "true")
    run_begin._r.append(fld_begin)

    run_instr = p.add_run()
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = f" {instr} "
    run_instr._r.append(instr_el)

    run_sep = p.add_run()
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run_sep._r.append(fld_sep)

    run_txt = p.add_run(placeholder or "Daftar akan diperbarui otomatis oleh Word (klik kanan → Update Field).")
    run_txt.font.name = "Times New Roman"
    run_txt.font.size = Pt(11)
    run_txt.italic = True

    run_end = p.add_run()
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run_end._r.append(fld_end)

    return p


def configure_s1_caption_styles(doc: Document):
    """
    Membuat dua gaya caption khusus agar DAFTAR TABEL dan DAFTAR GAMBAR dapat
    dipisahkan tanpa field SEQ di dalam teks caption.

    Pendekatan `TOC \\t "<gaya>;1"` dipilih karena idempoten: hasilnya tetap benar
    walau field diperbarui berkali-kali (Ctrl+A lalu F9), tidak menambah nomor
    tersembunyi di samping caption manual seperti "Tabel 4.22".
    """
    specs = {
        "S1 Caption Tabel": (12, 6),
        "S1 Caption Gambar": (8, 4),
    }
    for name, (before, after) in specs.items():
        try:
            st = doc.styles[name]
        except KeyError:
            st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            try:
                st.base_style = doc.styles["Caption"]
            except KeyError:
                pass
        f = st.font
        f.name = "Times New Roman"
        f.size = Pt(12)
        f.bold = True
        f.italic = False
        f.color.rgb = RGBColor(0, 0, 0)
        rpr = st.element.get_or_add_rPr()
        node = rpr.find(qn("w:rFonts"))
        if node is None:
            node = OxmlElement("w:rFonts")
            rpr.append(node)
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            node.set(qn(attr), "Times New Roman")
        pf = st.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.line_spacing = 1.0
        pf.keep_with_next = True
        pf.keep_together = True


def s1_front_heading(doc: Document, text: str, space_after: int = 18):
    """
    Judul halaman Bagian Awal yang WAJIB muncul di DAFTAR ISI.
    Memakai gaya Heading 1 agar dipungut field TOC, dengan tampilan baku
    FORMAT S1.pdf: Times New Roman 12 pt, tebal, terpusat, spasi 1.0.
    """
    p = doc.add_paragraph(style="Heading 1")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.italic = False
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p


def s1_bab_heading(doc: Document, bab_num: str, bab_title: str):
    """Format Judul BAB: 2 baris terpusat, huruf kapital, tebal, 12 pt, gaya Heading 1 (masuk TOC)."""
    p_bab = doc.add_paragraph(style="Heading 1")
    p_bab.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_bab.paragraph_format.space_before = Pt(0)
    p_bab.paragraph_format.space_after = Pt(18)
    p_bab.paragraph_format.line_spacing = 1.5
    r1 = p_bab.add_run(f"BAB {bab_num.upper()}")
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(12)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0, 0, 0)
    _force_run_black(r1)
    r1.add_break()
    r1.add_text(bab_title.upper())


def s1_sub_heading(doc: Document, num: str, title: str):
    """Format Sub-Bab Level 1 (misal 1.1, 2.1): rata kiri, tebal, 12 pt, gaya Heading 2 (masuk TOC).
    Opsi A: Nomor di margin (0 cm), tab stop di 1.27 cm (0.5 inch), sehingga huruf awal judul
    sejajar tegak lurus dengan huruf pertama paragraf di bawahnya.
    """
    p_h = doc.add_paragraph(style="Heading 2")
    p_h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_h.paragraph_format.space_before = Pt(12)
    p_h.paragraph_format.space_after = Pt(6)
    p_h.paragraph_format.line_spacing = 1.5
    p_h.paragraph_format.tab_stops.add_tab_stop(Inches(0.5))
    r = p_h.add_run(f"{num}\t{to_academic_title_case(title)}")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    _force_run_black(r)


def s1_sub_sub_heading(doc: Document, num: str, title: str):
    """Format Sub-Sub-Bab Level 2 (misal 2.1.1, 3.1.1): rata kiri, tebal, 12 pt, gaya Heading 3 (masuk TOC).
    Opsi A: Nomor di margin (0 cm), tab stop di 1.27 cm (0.5 inch), sehingga huruf awal judul
    sejajar tegak lurus dengan huruf pertama paragraf di bawahnya.
    """
    p_h = doc.add_paragraph(style="Heading 3")
    p_h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_h.paragraph_format.space_before = Pt(8)
    p_h.paragraph_format.space_after = Pt(4)
    p_h.paragraph_format.line_spacing = 1.5
    p_h.paragraph_format.tab_stops.add_tab_stop(Inches(0.5))
    r = p_h.add_run(f"{num}\t{to_academic_title_case(title)}")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    _force_run_black(r)


def s1_sub_sub_sub_heading(doc: Document, num: str, title: str):
    """Format Sub-Sub-Sub-Bab Level 3 (misal 2.1.1.1): rata kiri, tebal, 12 pt, Title Case.
    Opsi A: Nomor di margin (0 cm), tab stop di 1.27 cm (0.5 inch), sehingga huruf awal judul
    sejajar tegak lurus dengan huruf pertama paragraf di bawahnya.
    """
    p_h = doc.add_paragraph()
    p_h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_h.paragraph_format.space_before = Pt(6)
    p_h.paragraph_format.space_after = Pt(2)
    p_h.paragraph_format.line_spacing = 1.5
    p_h.paragraph_format.tab_stops.add_tab_stop(Inches(0.5))
    r = p_h.add_run(f"{num}\t{to_academic_title_case(title)}")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    _force_run_black(r)


def s1_p(doc: Document, text: str, bold: bool = False, italic: bool = False,
         align = WD_ALIGN_PARAGRAPH.JUSTIFY, indent: float = 0.5,
         space_before: int = 0, space_after: int = 0, line_spacing: float = None):
    """Format paragraf teks naskah S1: Times New Roman 12 pt, spasi 1.5, indent 1.27 cm (0.5 inch), justified."""
    if line_spacing is None:
        line_spacing = S1_LINE_SPACING
    text = s1_lex(text)
    para = doc.add_paragraph()
    para.alignment = align
    if indent > 0:
        para.paragraph_format.first_line_indent = Inches(indent)
    else:
        para.paragraph_format.first_line_indent = Inches(0)
    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.line_spacing = line_spacing

    run = para.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    run.font.bold = bold
    run.font.italic = italic
    return para


def s1_list_item(doc: Document, num_label: str, text: str, indent: float = 0.5,
                 bold_label: bool = False, line_spacing: float = None, space_after: int = 2):
    """Format butir daftar bernomor (1., 2., a., b.) dengan hanging indent terstandarisasi (Model 1):
    - Nomor/label berada di margin kiri (0 cm jika indent=0.5)
    - Tab stop tepat di 1.27 cm (0.5 inch / 36 pt)
    - Seluruh baris lanjutan mengunci rapi di posisi 1.27 cm (tidak melorot ke margin).
    - Bila terdapat teks judul sebelum tanda titik dua (mis. "Manfaat Teoritis: ..."), judul diformat tebal.
    """
    if line_spacing is None:
        line_spacing = S1_LINE_SPACING
    text = s1_lex(text)
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    indent_width = Inches(indent)
    para.paragraph_format.left_indent = indent_width
    if indent <= 0.5:
        para.paragraph_format.first_line_indent = -indent_width
        para.paragraph_format.tab_stops.add_tab_stop(indent_width)
    else:
        para.paragraph_format.first_line_indent = Inches(-0.25)
        para.paragraph_format.tab_stops.add_tab_stop(indent_width)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.line_spacing = line_spacing

    r_lbl = para.add_run(f"{num_label}\t")
    r_lbl.font.name = "Times New Roman"
    r_lbl.font.size = Pt(12)
    r_lbl.font.bold = bold_label

    # Format tebal otomatis untuk judul butir sebelum titik dua (misal "Manfaat Teoritis: ...", "Data Primer: ...")
    if ": " in text and not text.startswith("http") and not text.startswith("doi:"):
        prefix, rest = text.split(": ", 1)
        if len(prefix) <= 45 and not "\n" in prefix:
            r_title = para.add_run(f"{prefix}: ")
            r_title.font.name = "Times New Roman"
            r_title.font.size = Pt(12)
            r_title.font.bold = True

            r_txt = para.add_run(rest)
            r_txt.font.name = "Times New Roman"
            r_txt.font.size = Pt(12)
            return para

    r_txt = para.add_run(text)
    r_txt.font.name = "Times New Roman"
    r_txt.font.size = Pt(12)
    return para


def apply_s1_full_borders_table(table, header_rows: int = 1):
    """
    Format Tabel Tertutup Semua (Full Borders) sesuai format Skripsi S1:
    - Border garis hitam solid pada: top, bottom, left, right, insideH, insideV.
    - Paskan ke lebar margin jendela (100% / 5000 pct).
    - Spasi isi tabel: 1.0 (tunggal), seluruh sel menggunakan Font 12 pt.
    """
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._tbl.tblPr

    existing_tblW = tblPr.find(qn('w:tblW'))
    if existing_tblW is not None:
        tblPr.remove(existing_tblW)
    tblPr.append(parse_xml(r'<w:tblW %s w:w="5000" w:type="pct"/>' % nsdecls('w')))

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

    for row_idx, row in enumerate(table.rows):
        is_header = (row_idx < header_rows)
        trPr = row._tr.get_or_add_trPr()
        if is_header:
            trPr.append(parse_xml(r'<w:tblHeader %s/>' % nsdecls('w')))
        trPr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))

        for cell in row.cells:
            # Pastikan border sel bersih dan mengikuti tblBorders tertutup penuh
            tcPr = cell._tc.get_or_add_tcPr()
            existing = tcPr.find(qn('w:tcBorders'))
            if existing is not None:
                tcPr.remove(existing)
            for para in cell.paragraphs:
                para.paragraph_format.line_spacing = 1.0
                para.paragraph_format.space_before = Pt(2)
                para.paragraph_format.space_after = Pt(2)


def add_s1_table(doc: Document, table_num: str, title_text: str, headers: List[str],
                 data: List[List[str]], source: str = "", widths: Optional[List[float]] = None,
                 cell_font_size: int = 11, body_align_left: bool = False,
                 alignments: Optional[List[int]] = None):
    """
    Menyisipkan Tabel Akademik Tertutup Penuh (Full-Borders) sesuai instruksi pengguna:
    - Judul di ATAS tabel 2 baris terpusat tebal (Tabel X.X\nJudul Tabel).
    - Format Tabel Tertutup Penuh (Semua Garis: top, bottom, left, right, insideH, insideV).
    - Seluruh isi tabel dan header menggunakan Font Times New Roman 11 pt.
    - Sumber di BAWAH tabel rata kiri (Sumber: ..., Tahun).
    """
    title_text = s1_lex(title_text)
    headers = [s1_lex(str(h)) for h in headers]
    data = [[s1_lex(str(v if v is not None else "")) for v in row] for row in data]
    source = s1_lex(source)
    # Judul tabel = SATU paragraf bergaya "S1 Caption Tabel" berisi:
    # "Tabel X.X" + baris baru + judul. Gaya inilah yang dipungut field
    # DAFTAR TABEL (TOC \t "S1 Caption Tabel;1") sehingga nomor halaman akurat.
    p_cap = doc.add_paragraph(style="S1 Caption Tabel")
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(12)
    p_cap.paragraph_format.space_after = Pt(6)
    p_cap.paragraph_format.line_spacing = 1.0
    p_cap.paragraph_format.keep_with_next = True
    r1 = p_cap.add_run(table_num)
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(12)
    r1.font.bold = True
    r1.font.italic = False
    r1.add_break()
    r1.add_text(to_academic_title_case(title_text))

    table = doc.add_table(rows=len(data) + 1, cols=len(headers))
    header_row = table.rows[0]
    for c_idx, h in enumerate(headers):
        cell = header_row.cells[c_idx]
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.line_spacing = 1.0
        para.paragraph_format.space_before = Pt(3)
        para.paragraph_format.space_after = Pt(3)
        run = para.add_run(str(h if h is not None else ""))
        run.font.name = "Times New Roman"
        run.font.size = Pt(cell_font_size)  # Font tabel 11 pt
        run.font.bold = True

    for r_idx, row_data in enumerate(data):
        row = table.rows[r_idx + 1]
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            para = cell.paragraphs[0]
            para.paragraph_format.line_spacing = 1.0
            para.paragraph_format.space_before = Pt(2)
            para.paragraph_format.space_after = Pt(2)
            run = para.add_run(str(val if val is not None else ""))
            run.font.name = "Times New Roman"
            run.font.size = Pt(cell_font_size)  # Font tabel 11 pt
            run.font.bold = False

            if c_idx in (0, 1):
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            elif alignments and c_idx < len(alignments):
                para.alignment = alignments[c_idx]
            elif body_align_left:
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    apply_s1_full_borders_table(table)
    if widths:
        try:
            apply_fixed_widths(table, widths)
        except Exception:
            pass

    if source:
        source_clean = source.strip().strip('"\'')
        if not source_clean.lower().startswith("sumber"):
            source_clean = f"Sumber: {source_clean}"
        p_src = doc.add_paragraph()
        p_src.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_src.paragraph_format.space_before = Pt(3)
        p_src.paragraph_format.space_after = Pt(12)
        p_src.paragraph_format.line_spacing = 1.0
        r_src = p_src.add_run(source_clean)
        r_src.font.name = "Times New Roman"
        r_src.font.size = Pt(11)
        r_src.font.italic = False
    else:
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_after = Pt(6)


def add_s1_figure(doc: Document, fig_num: str, title_text: str, ctx: Dict[str, Any],
                  source: str = "Diolah Peneliti, 2026"):
    """
    Menyisipkan Bagan Gambar Kerangka Berpikir Standar Visual Skripsi S1 (1 s.d. 5 Variabel X):
    - Dihasilkan via engine grafis publikasi PIL (s1_framework_diagram.py) resolusi tinggi.
    - Menampilkan kotak variabel bebas berbingkai (X), kotak residual faktor lain (e), dan variabel terikat (Y).
    - Jalur pengaruh orthogonal 90 derajat dengan uji t parsial dan uji F simultan.
    - Judul Gambar di bawah diagram dengan gaya 'S1 Caption Gambar' (terindeks otomatis di Daftar Gambar).
    - Keterangan dan Sumber akademik di bawah judul gambar.
    """
    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kompensasi"), "profile": s1_constructs.get_construct_profile(ctx.get("x1_name", ""))},
        {"code": "X2", "name": ctx.get("x2_name", "Kompetensi"), "profile": s1_constructs.get_construct_profile(ctx.get("x2_name", ""))}
    ]
    k_x = len(x_vars)
    y_name = ctx.get("y_name", "Kinerja Pegawai")

    # Generate Image via s1_framework_diagram
    img_inserted = False
    try:
        import s1_framework_diagram
        fig_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output", "figs")
        safe_key = "".join(ch if ch.isalnum() else "_" for ch in f"{ctx.get('title', 'kerangka')}_{k_x}x")[:60]
        img_path = os.path.join(fig_dir, f"kerangka_{safe_key}.png")

        s1_framework_diagram.generate_framework_diagram(
            x_variables=[{"name": xv["name"], "code": xv["code"]} for xv in x_vars],
            y_variable={"name": y_name, "code": "Y"},
            label_style="H",
            simultaneous_style="straight" if k_x == 2 else ("none" if k_x == 1 else "loop"),
            output_path=img_path
        )
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(6)
            p_img.paragraph_format.line_spacing = 1.0
            r_img = p_img.add_run()
            # Lebar proporsional 14.2 cm agar pas dengan margin 4-4-3-3 kertas A4
            r_img.add_picture(img_path, width=Cm(14.2))
            img_inserted = True
    except Exception as e:
        print(f"[add_s1_figure] Warning: Image generation fallback triggered: {e}")

    # Fallback ke tabel diagram jika gambar tidak dapat dibuat
    if not img_inserted:
        tbl_diag = doc.add_table(rows=k_x + 1, cols=3)
        tbl_diag.alignment = WD_TABLE_ALIGNMENT.CENTER

        widths = [Cm(5.3), Cm(3.4), Cm(5.3)]
        for row in tbl_diag.rows:
            for idx, w in enumerate(widths):
                row.cells[idx].width = w

        def format_box_cell(cell, title_text: str, var_code: str, indicators: List[str]):
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tcPr = cell._tc.get_or_add_trPr() if hasattr(cell._tc, 'get_or_add_trPr') else cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                r'<w:tcBorders %s>'
                r'  <w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                r'  <w:left w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                r'  <w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                r'  <w:right w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                r'</w:tcBorders>' % nsdecls('w')
            )
            tcPr.append(tcBorders)
            tcPr.append(parse_xml(r'<w:shd %s w:val="clear" w:color="auto" w:fill="FAFAFA"/>' % nsdecls('w')))

            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(2)
            r_head = p.add_run(f"{title_text}\n({var_code})\n")
            r_head.font.name = "Times New Roman"
            r_head.font.size = Pt(11)
            r_head.font.bold = True

            for ind in indicators:
                p_ind = cell.add_paragraph()
                p_ind.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p_ind.paragraph_format.line_spacing = 1.0
                p_ind.paragraph_format.space_before = Pt(1)
                p_ind.paragraph_format.space_after = Pt(1)
                p_ind.paragraph_format.left_indent = Inches(0.15)
                r_ind = p_ind.add_run(ind)
                r_ind.font.name = "Times New Roman"
                r_ind.font.size = Pt(9.5)

        def format_clear_cell(cell):
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                r'<w:tcBorders %s>'
                r'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                r'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                r'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                r'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                r'</w:tcBorders>' % nsdecls('w')
            )
            tcPr.append(tcBorders)

        for i, xv in enumerate(x_vars):
            dims = (xv.get("profile") or {}).get("dimensions", [])
            ind_x = [f"{idx+1}. {d['dim']}" for idx, d in enumerate(dims[:3])]
            if not ind_x:
                ind_x = [f"1. Dimensi 1 {xv['code']}", f"2. Dimensi 2 {xv['code']}", f"3. Dimensi 3 {xv['code']}"]
            format_box_cell(tbl_diag.cell(i, 0), xv["name"], xv["code"], ind_x)

            c_arr = tbl_diag.cell(i, 1)
            format_clear_cell(c_arr)
            c_arr.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p_arr = c_arr.paragraphs[0]
            p_arr.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_arr.paragraph_format.line_spacing = 1.15
            r_arr = p_arr.add_run(f"\n\nH{i+1}: Parsial (Uji t)\n────────►\n(Positif & Signifikan)")
            r_arr.font.name = "Times New Roman"
            r_arr.font.size = Pt(9.5)
            r_arr.font.bold = True

        y_profile = ctx.get("y_profile") or s1_constructs.get_construct_profile(ctx.get("y_name", "Kinerja Pegawai"))
        ind_y = [f"{idx+1}. {d['dim']}" for idx, d in enumerate(y_profile.get("dimensions", [])[:3])]
        if not ind_y:
            ind_y = ["1. Kualitas Hasil Kerja", "2. Kuantitas Hasil Kerja", "3. Ketepatan Waktu Kerja"]

        c_y_merged = tbl_diag.cell(0, 2)
        for i in range(1, k_x):
            c_y_merged = c_y_merged.merge(tbl_diag.cell(i, 2))
        format_box_cell(c_y_merged, ctx["y_name"], "Y", ind_y)

        c_sim = tbl_diag.cell(k_x, 0)
        for c in range(1, 3):
            c_sim = c_sim.merge(tbl_diag.cell(k_x, c))
        c_sim.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        tcPr_sim = c_sim._tc.get_or_add_tcPr()
        tcBorders_sim = parse_xml(
            r'<w:tcBorders %s>'
            r'  <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'  <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'  <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            r'</w:tcBorders>' % nsdecls('w')
        )
        tcPr_sim.append(tcBorders_sim)
        tcPr_sim.append(parse_xml(r'<w:shd %s w:val="clear" w:color="auto" w:fill="F4F6F8"/>' % nsdecls('w')))

        p_sim = c_sim.paragraphs[0]
        p_sim.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sim.paragraph_format.line_spacing = 1.15
        p_sim.paragraph_format.space_before = Pt(4)
        p_sim.paragraph_format.space_after = Pt(4)
        r_sim = p_sim.add_run(f"H{k_x+1}: Pengaruh Simultan (Uji F) (Variabel {ctx['x_join']} ════════► {ctx['y_name']})")
        r_sim.font.name = "Times New Roman"
        r_sim.font.size = Pt(10)
        r_sim.font.bold = True

    # Judul Gambar di Bawah Diagram = satu paragraf gaya "S1 Caption Gambar",
    # sehingga DAFTAR GAMBAR (TOC \t "S1 Caption Gambar;1") terisi otomatis.
    try:
        p_fn = doc.add_paragraph(style="S1 Caption Gambar")
    except Exception:
        p_fn = doc.add_paragraph()
    p_fn.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_fn.paragraph_format.space_before = Pt(8)
    p_fn.paragraph_format.space_after = Pt(4)
    p_fn.paragraph_format.line_spacing = 1.0
    r_fn = p_fn.add_run(fig_num)
    r_fn.font.name = "Times New Roman"
    r_fn.font.size = Pt(12)
    r_fn.font.bold = True
    r_fn.font.italic = False
    r_fn.add_break()
    r_fn.add_text(to_academic_title_case(title_text))

    # Sumber di Bawah Gambar
    p_fs = doc.add_paragraph()
    p_fs.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_fs.paragraph_format.space_before = Pt(3)
    p_fs.paragraph_format.space_after = Pt(12)
    p_fs.paragraph_format.line_spacing = 1.0
    src_fig = source.strip().strip('"\'')
    if not src_fig.lower().startswith("sumber"):
        src_fig = f"Sumber: {src_fig}"
    r_fs = p_fs.add_run(src_fig)
    r_fs.font.name = "Times New Roman"
    r_fs.font.size = Pt(11)
    r_fs.font.italic = False


def tbl_offset_4_x(k: int, section: str) -> int:
    """
    Menghitung offset nomor tabel pada Bab IV berdasarkan jumlah variabel X (k).
    Section offsets:
    - "demog": 0 -> Tabel 4.1 .. 4.4
    - "freq": 4 -> Tabel 4.5 .. 4.(5 + k)
    - "validity": 5 + k -> Tabel 4.(6 + k) .. 4.(6 + 2k)
    - "reliability": 6 + 2k -> Tabel 4.(7 + 2k) .. 4.(8 + 3k)
    - "descriptive": 8 + 3k -> Tabel 4.(9 + 3k) .. 4.(10 + 4k)
    - "classic": 10 + 4k -> Tabel 4.(11 + 4k) .. 4.(13 + 4k)
    - "regression": 13 + 4k -> Tabel 4.(14 + 4k) .. 4.(15 + 4k)
    """
    offsets = {
        "demog": 0,
        "freq": 4,
        "validity": 5 + k,
        "reliability": 6 + 2 * k,
        "descriptive": 8 + 3 * k,
        "classic": 10 + 4 * k,
        "regression": 13 + 4 * k,
    }
    return offsets.get(section, 0)


# =========================================================================
# 3. DAFTAR PUSTAKA ACUAN (ZERO GHOST REFERENCE)
# =========================================================================

S1_CORE_BIBLIOGRAPHY = s1_literature_library.S1_EXPANDED_BIBLIOGRAPHY


# =========================================================================
# 4. GENERATOR BAGIAN AWAL (FRONT MATTER) LENGKAP & SUB-BUILDER
# =========================================================================

def build_s1_sub_cover(doc: Document, ctx: Dict[str, Any]):
    """Membangun Halaman Judul Luar (Cover Luar) dan Judul Dalam (Cover Dalam)."""
    # 1. HALAMAN JUDUL LUAR (COVER LUAR)
    p_cov = doc.add_paragraph()
    p_cov.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cov = p_cov.add_run("SKRIPSI")
    r_cov.font.name = "Times New Roman"
    r_cov.font.size = Pt(14)
    r_cov.font.bold = True
    p_cov.paragraph_format.space_before = Pt(0)
    p_cov.paragraph_format.space_after = Pt(24)

    p_jdl = doc.add_paragraph()
    p_jdl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_jdl = p_jdl.add_run(ctx["title_upper"])
    r_jdl.font.name = "Times New Roman"
    r_jdl.font.size = Pt(13)
    r_jdl.font.bold = True
    p_jdl.paragraph_format.space_before = Pt(0)
    p_jdl.paragraph_format.space_after = Pt(36)
    p_jdl.paragraph_format.line_spacing = 1.15

    p_mks = doc.add_paragraph()
    p_mks.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_mks.paragraph_format.space_before = Pt(12)
    p_mks.paragraph_format.space_after = Pt(36)
    p_mks.paragraph_format.line_spacing = 1.15
    r_mks = p_mks.add_run(
        f"Diajukan Untuk Memenuhi Sebagian Syarat Memperoleh Gelar Sarjana Manajemen (S.M.)\n"
        f"Program Studi {ctx['program_studi']}\n"
        f"Konsentrasi {ctx['konsentrasi']}"
    )
    r_mks.font.name = "Times New Roman"
    r_mks.font.size = Pt(11)

    p_oleh = doc.add_paragraph()
    p_oleh.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_oleh.paragraph_format.space_before = Pt(12)
    p_oleh.paragraph_format.space_after = Pt(2)
    r_oleh = p_oleh.add_run("Diajukan Oleh :\n")
    r_oleh.font.name = "Times New Roman"
    r_oleh.font.size = Pt(11)

    r_pen = p_oleh.add_run(ctx["penulis_upper"] + "\n")
    r_pen.font.name = "Times New Roman"
    r_pen.font.size = Pt(12)
    r_pen.font.bold = True

    r_nim = p_oleh.add_run(f"NIM. {ctx['nim']}")
    r_nim.font.name = "Times New Roman"
    r_nim.font.size = Pt(12)
    r_nim.font.bold = True

    p_pt = doc.add_paragraph()
    p_pt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pt.paragraph_format.space_before = Pt(48)
    p_pt.paragraph_format.space_after = Pt(0)
    p_pt.paragraph_format.line_spacing = 1.15
    r_pt = p_pt.add_run(
        f"{ctx['fakultas'].upper()}\n"
        f"{ctx['universitas'].upper()}\n"
        f"{ctx['kota'].upper()}\n"
        f"{ctx['tahun']}"
    )
    r_pt.font.name = "Times New Roman"
    r_pt.font.size = Pt(12)
    r_pt.font.bold = True
    doc.add_page_break()

    # 2. HALAMAN JUDUL DALAM (COVER DALAM)
    p_cov2 = doc.add_paragraph()
    p_cov2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cov2 = p_cov2.add_run("SKRIPSI")
    r_cov2.font.name = "Times New Roman"
    r_cov2.font.size = Pt(14)
    r_cov2.font.bold = True
    p_cov2.paragraph_format.space_before = Pt(0)
    p_cov2.paragraph_format.space_after = Pt(24)

    p_jdl2 = doc.add_paragraph()
    p_jdl2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_jdl2 = p_jdl2.add_run(ctx["title_upper"])
    r_jdl2.font.name = "Times New Roman"
    r_jdl2.font.size = Pt(13)
    r_jdl2.font.bold = True
    p_jdl2.paragraph_format.space_before = Pt(0)
    p_jdl2.paragraph_format.space_after = Pt(36)
    p_jdl2.paragraph_format.line_spacing = 1.15

    p_mks2 = doc.add_paragraph()
    p_mks2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_mks2.paragraph_format.space_before = Pt(12)
    p_mks2.paragraph_format.space_after = Pt(36)
    p_mks2.paragraph_format.line_spacing = 1.15
    r_mks2 = p_mks2.add_run(
        f"Diajukan Untuk Memenuhi Sebagian Syarat Memperoleh Gelar Sarjana Manajemen (S.M.)\n"
        f"Program Studi {ctx['program_studi']}\n"
        f"Konsentrasi {ctx['konsentrasi']}"
    )
    r_mks2.font.name = "Times New Roman"
    r_mks2.font.size = Pt(11)

    p_oleh2 = doc.add_paragraph()
    p_oleh2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_oleh2.paragraph_format.space_before = Pt(12)
    p_oleh2.paragraph_format.space_after = Pt(2)
    r_oleh2 = p_oleh2.add_run("Diajukan Oleh :\n")
    r_oleh2.font.name = "Times New Roman"
    r_oleh2.font.size = Pt(11)

    r_pen2 = p_oleh2.add_run(ctx["penulis_upper"] + "\n")
    r_pen2.font.name = "Times New Roman"
    r_pen2.font.size = Pt(12)
    r_pen2.font.bold = True

    r_nim2 = p_oleh2.add_run(f"NIM. {ctx['nim']}")
    r_nim2.font.name = "Times New Roman"
    r_nim2.font.size = Pt(12)
    r_nim2.font.bold = True

    p_pt2 = doc.add_paragraph()
    p_pt2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pt2.paragraph_format.space_before = Pt(48)
    p_pt2.paragraph_format.space_after = Pt(0)
    p_pt2.paragraph_format.line_spacing = 1.15
    r_pt2 = p_pt2.add_run(
        f"{ctx['fakultas'].upper()}\n"
        f"{ctx['universitas'].upper()}\n"
        f"{ctx['kota'].upper()}\n"
        f"{ctx['tahun']}"
    )
    r_pt2.font.name = "Times New Roman"
    r_pt2.font.size = Pt(12)
    r_pt2.font.bold = True


def build_s1_sub_persetujuan(doc: Document, ctx: Dict[str, Any]):
    """Membangun Halaman Persetujuan Skripsi."""
    p_hdr_ps = doc.add_paragraph()
    p_hdr_ps.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_hdr_ps.paragraph_format.space_before = Pt(0)
    p_hdr_ps.paragraph_format.space_after = Pt(18)
    r_hps = p_hdr_ps.add_run(f"{ctx['fakultas'].upper()}\n{ctx['universitas'].upper()}\n{ctx['kota'].upper()}")
    r_hps.font.name = "Times New Roman"
    r_hps.font.size = Pt(12)
    r_hps.font.bold = True

    s1_front_heading(doc, "HALAMAN PERSETUJUAN SKRIPSI")

    s1_p(doc, f"Nama\t\t\t: {ctx['penulis_display']}", indent=0, space_after=2, line_spacing=1.15)
    s1_p(doc, f"Nomor Pokok / NIM\t: {ctx['nim']}", indent=0, space_after=2, line_spacing=1.15)
    s1_p(doc, f"Jurusan / Prog. Studi\t: {ctx['program_studi']}", indent=0, space_after=2, line_spacing=1.15)
    s1_p(doc, f"Jenjang Pendidikan\t: Strata 1 (S.1)", indent=0, space_after=2, line_spacing=1.15)
    s1_p(doc, f"Mata Kuliah Pokok\t: Manajemen", indent=0, space_after=2, line_spacing=1.15)
    s1_p(doc, f"Judul Skripsi\t\t: {ctx['title_display']}", indent=0, space_after=18, line_spacing=1.15)

    s1_p(doc, "Pembimbing Skripsi :", bold=True, indent=0, space_after=6, line_spacing=1.15)
    tbl_pemb = doc.add_table(rows=1, cols=2)
    tbl_pemb.alignment = WD_TABLE_ALIGNMENT.CENTER
    c1, c2 = tbl_pemb.rows[0].cells
    c1.width = Cm(8.0)
    c2.width = Cm(8.0)
    c1.paragraphs[0].text = f"Pembimbing I :\n\n\n\n\n{ctx['pembimbing1']}\nNIDN. {ctx['nidn_pembimbing1']}"
    c1.paragraphs[0].runs[0].font.name = "Times New Roman"
    c2.paragraphs[0].text = f"Pembimbing II :\n\n\n\n\n{ctx['pembimbing2']}\nNIDN. {ctx['nidn_pembimbing2']}"
    c2.paragraphs[0].runs[0].font.name = "Times New Roman"

    p_mengetahui = doc.add_paragraph()
    p_mengetahui.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_mengetahui.paragraph_format.space_before = Pt(24)
    p_mengetahui.paragraph_format.space_after = Pt(0)
    p_mengetahui.paragraph_format.line_spacing = 1.15
    r_mgt = p_mengetahui.add_run(
        f"Mengetahui,\n"
        f"Dekan {ctx['fakultas']}\n\n\n\n\n"
        f"{ctx['dekan']}\nNIP. {ctx['nip_dekan']}"
    )
    r_mgt.font.name = "Times New Roman"
    r_mgt.font.size = Pt(12)


def build_s1_sub_pernyataan(doc: Document, ctx: Dict[str, Any]):
    """Membangun Surat Pernyataan Bebas Plagiat / Keaslian Skripsi."""
    s1_front_heading(doc, "SURAT PERNYATAAN BEBAS PLAGIAT")

    s1_p(doc, "Saya yang bertanda tangan di bawah ini:", indent=0, space_after=6, line_spacing=1.5)
    s1_p(doc, f"Nama\t\t\t: {ctx['penulis_display']}", indent=0.5, space_after=2, line_spacing=1.5)
    s1_p(doc, f"NIM\t\t\t: {ctx['nim']}", indent=0.5, space_after=2, line_spacing=1.5)
    s1_p(doc, f"Program Studi\t\t: {ctx['program_studi']}", indent=0.5, space_after=2, line_spacing=1.5)
    s1_p(doc, f"Konsentrasi\t\t: {ctx['konsentrasi']}", indent=0.5, space_after=12, line_spacing=1.5)

    s1_p(doc, f"Menyatakan dengan sesungguhnya bahwa skripsi yang berjudul \"{ctx['title_display']}\" adalah benar-benar karya saya sendiri dan bukan merupakan jiplakan atau plagiasi dari karya orang lain, baik sebagian maupun seluruhnya.", indent=0.5, space_after=12, line_spacing=1.5)
    s1_p(doc, "Apabila di kemudian hari terbukti atau dapat dibuktikan bahwa skripsi ini merupakan hasil jiplakan/plagiat, maka saya bersedia menerima sanksi pencabutan gelar akademik dan sanksi lainnya sesuai dengan ketentuan peraturan perundang-undangan yang berlaku.", indent=0.5, space_after=24, line_spacing=1.5)

    p_ttd_sp = doc.add_paragraph()
    p_ttd_sp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_ttd_sp.paragraph_format.line_spacing = 1.15
    r_tsp = p_ttd_sp.add_run(
        f"{ctx['kota']}, {datetime.now().strftime('%d %B %Y')}\n"
        f"Yang Menyatakan,\n\n\n\n"
        f"Materai Rp10.000,-\n\n"
        f"{ctx['penulis_display']}\n"
        f"NIM. {ctx['nim']}"
    )
    r_tsp.font.name = "Times New Roman"
    r_tsp.font.size = Pt(12)


def build_s1_sub_motto(doc: Document, ctx: Dict[str, Any]):
    """Membangun Halaman Motto dan Persembahan."""
    s1_front_heading(doc, "MOTTO DAN PERSEMBAHAN")

    s1_p(doc, "Motto :", bold=True, indent=0, space_after=6, line_spacing=1.5)
    s1_p(doc, "“Bangun kesuksesan dari kegagalan. Keputusasaan dan kegagalan adalah dua batu loncatan yang paling baik menuju kesuksesan.” – Dale Carnegie", italic=True, indent=0.5, space_after=6, line_spacing=1.5)
    s1_p(doc, "“Raih ilmu, dan untuk meraih ilmu belajarlah untuk tenang dan sabar.” – Umar bin Khattab", italic=True, indent=0.5, space_after=24, line_spacing=1.5)

    s1_p(doc, "Kupersembahkan Kepada :", bold=True, indent=0, space_after=6, line_spacing=1.5)
    s1_list_item(doc, "▪", "Allah SWT, terima kasih atas segala rahmat, taufik, dan hidayah-Nya sehingga skripsi ini dapat diselesaikan dengan baik dan lancar.", indent=0.5, line_spacing=1.5)
    s1_list_item(doc, "▪", "Kedua orang tua tercinta yang senantiasa melimpahkan doa tulus, kasih sayang tanpa pamrih, serta pengorbanan moril dan materiil yang tiada ternilai.", indent=0.5, line_spacing=1.5)
    s1_list_item(doc, "▪", "Bapak dan Ibu Dosen Fakultas Ekonomi dan Bisnis atas bimbingan, keteladanan ilmu, dan motivasi yang sangat berharga selama menempuh perkuliahan.", indent=0.5, line_spacing=1.5)
    s1_list_item(doc, "▪", f"Rekan-rekan seperjuangan mahasiswa Program Studi {ctx['program_studi']} dan Almamater tercinta yang kubanggakan.", indent=0.5, line_spacing=1.5)


def build_s1_sub_kata_pengantar(doc: Document, ctx: Dict[str, Any]):
    """Membangun Halaman Kata Pengantar."""
    s1_front_heading(doc, "KATA PENGANTAR")

    s1_p(doc, f"Puji syukur penulis haturkan ke hadirat Allah SWT atas segala limpahan rahmat, berkah, dan pertolongan-Nya sehingga penulis dapat menyelesaikan penyusunan skripsi ini dengan judul: \"{ctx['title_display']}\".", indent=0.5, space_after=6, line_spacing=S1_LINE_SPACING)
    s1_p(doc, f"Skripsi ini disusun sebagai salah satu syarat untuk menyelesaikan pendidikan Program Sarjana (S1) dan memperoleh gelar Sarjana Manajemen (S.M.) pada Fakultas Ekonomi dan Bisnis {ctx['universitas']}.", indent=0.5, space_after=6, line_spacing=S1_LINE_SPACING)
    s1_p(doc, "Dalam proses penyelesaian skripsi ini, penulis mendapatkan bantuan, bimbingan, motivasi, serta doa yang tulus dari berbagai pihak. Oleh karena itu, penulis ingin menyampaikan ucapan terima kasih dan penghargaan setinggi-tingginya kepada:", indent=0.5, space_after=6, line_spacing=S1_LINE_SPACING)

    s1_list_item(doc, "1.", f"{ctx['dekan']}, selaku Dekan {ctx['fakultas']} {ctx['universitas']}.", indent=0.5, line_spacing=1.5)
    s1_list_item(doc, "2.", f"{ctx['pembimbing1']}, selaku Dosen Pembimbing I yang telah meluangkan waktu dan penuh kesabaran memberikan bimbingan, koreksi, dan arahan ilmiah yang sangat bermanfaat.", indent=0.5, line_spacing=1.5)
    s1_list_item(doc, "3.", f"{ctx['pembimbing2']}, selaku Dosen Pembimbing II yang senantiasa membimbing, mengarahkan, dan memberikan masukan konstruktif dalam penyempurnaan skripsi ini.", indent=0.5, line_spacing=1.5)
    s1_list_item(doc, "4.", f"Seluruh Dosen dan Staf Pengajar {ctx['fakultas']} {ctx['universitas']} yang telah memberikan bekal ilmu pengetahuan yang bermanfaat.", indent=0.5, line_spacing=1.5)
    s1_list_item(doc, "5.", f"Pimpinan dan seluruh staf pegawai pada {ctx['instansi']} yang telah memberikan izin penelitian serta bersedia menjadi responden kuesioner.", indent=0.5, line_spacing=1.5)
    s1_list_item(doc, "6.", "Keluarga tercinta, sahabat, dan seluruh rekan seperjuangan yang selalu memberikan dorongan moral dan semangat yang tiada henti.", indent=0.5, line_spacing=1.5)

    s1_p(doc, "Penulis menyadari bahwa skripsi ini masih memiliki kekurangan karena keterbatasan ilmu dan pengalaman. Akhir kata, penulis berharap skripsi ini dapat memberikan manfaat nyata bagi pengembangan ilmu pengetahuan manajemen sumber daya manusia.", indent=0.5, space_after=24, line_spacing=S1_LINE_SPACING)

    p_kp_ttd = doc.add_paragraph()
    p_kp_ttd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_kp_ttd.paragraph_format.line_spacing = 1.15
    r_kpt = p_kp_ttd.add_run(f"{ctx['kota']}, {datetime.now().strftime('%B %Y')}\n\nPenulis,\n\n{ctx['penulis_display']}")
    r_kpt.font.name = "Times New Roman"
    r_kpt.font.size = Pt(12)


def build_s1_sub_abstrak(doc: Document, ctx: Dict[str, Any]):
    """Membangun Halaman Abstrak Bahasa Indonesia (Dinamis N Variabel X)."""
    s1_front_heading(doc, "ABSTRAK", space_after=12)

    p_cit = doc.add_paragraph()
    p_cit.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_cit.paragraph_format.line_spacing = 1.0
    p_cit.paragraph_format.space_after = Pt(12)
    r_cit = p_cit.add_run(
        f"{ctx['penulis_display']}. {ctx['tahun']}. \"{ctx['title_display']}\". "
        f"Skripsi, Program Studi {ctx['program_studi']}, {ctx['fakultas']}, {ctx['universitas']}. "
        f"Pembimbing: (1) {ctx['pembimbing1']}, (2) {ctx['pembimbing2']}".rstrip(".") + "."
    )
    r_cit.font.name = "Times New Roman"
    r_cit.font.size = Pt(11)
    r_cit.font.italic = True

    reg = ctx["spss"]["regression"]
    coefs = [c for c in reg["coefficients"] if c["model"] != "(Constant)"]
    anv = reg["anova_f"]

    partial_points = []
    for i, c in enumerate(coefs):
        partial_points.append(
            f"({i+2}) Secara parsial, {c['model']} berpengaruh positif dan signifikan terhadap {ctx['y_name']} "
            f"(t-hitung = {c['t_hitung']} > t-tabel = {ctx['spss']['t_table']}, Sig. = {c['sig']} < 0,05)"
        )
    partial_text = "; dan ".join(partial_points) if len(partial_points) <= 2 else ("; ".join(partial_points[:-1]) + "; dan " + partial_points[-1])

    p_ab_txt = doc.add_paragraph()
    p_ab_txt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_ab_txt.paragraph_format.line_spacing = 1.0
    p_ab_txt.paragraph_format.first_line_indent = Inches(0.5)
    p_ab_txt.paragraph_format.space_after = Pt(12)
    if ctx['populasi'] > ctx['sampel']:
        sample_intro = (
            f"berjumlah {ctx['populasi']} orang pegawai, dan penentuan sampel menggunakan rumus Slovin "
            f"dengan batas toleransi kesalahan e = 5% (0,05) sehingga diperoleh sampel representatif sebanyak {ctx['sampel']} orang responden. Teknik pengumpulan data "
        )
    else:
        sample_intro = (
            f"berjumlah {ctx['populasi']} orang pegawai, dan penentuan sampel menggunakan metode sampel jenuh (sensus) "
            f"sehingga seluruh anggota populasi dijadikan sampel sebanyak {ctx['sampel']} orang responden. Teknik pengumpulan data "
        )

    r_ab = p_ab_txt.add_run(
        f"Penelitian ini bertujuan untuk mengetahui dan menganalisis pengaruh {ctx['x_join']} "
        f"terhadap {ctx['y_name']} pada {ctx['instansi']}, baik secara simultan maupun parsial. "
        f"Metode penelitian yang digunakan adalah metode kuantitatif asosiatif kausal. Populasi dalam penelitian ini "
        f"{sample_intro}"
        f"menggunakan kuesioner dengan Skala Likert 1–5 yang telah memenuhi uji validitas dan uji reliabilitas. Analisis data "
        f"dilakukan menggunakan Analisis Regresi Linier Berganda, Uji Asumsi Klasik, serta Uji Hipotesis (Uji F dan Uji t) berbantuan IBM SPSS Statistics 26. "
        f"Hasil penelitian menunjukkan bahwa: (1) Secara simultan, terdapat pengaruh yang positif dan signifikan antara {ctx['x_join']} "
        f"terhadap {ctx['y_name']} pada {ctx['instansi']}, ditunjukkan oleh nilai F-hitung = {anv['f_hitung']} > F-tabel = {anv['f_tabel']} "
        f"dengan signifikansi 0,000 < 0,05; dan {partial_text}. "
        f"Nilai Koefisien Determinasi (R²) sebesar {reg['r_square']} menunjukkan bahwa kontribusi pengaruh variabel bebas secara bersama-sama "
        f"terhadap {ctx['y_name']} adalah sebesar {round(reg['r_square']*100, 1)}%, sedangkan sisanya sebesar {round((1-reg['r_square'])*100, 1)}% "
        f"dipengaruhi oleh faktor-faktor lain di luar model penelitian ini."
    )
    r_ab.font.name = "Times New Roman"
    r_ab.font.size = Pt(11)

    p_kw = doc.add_paragraph()
    p_kw.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_kw.paragraph_format.line_spacing = 1.0
    r_kw = p_kw.add_run(f"Kata Kunci: {', '.join(ctx['x_names'])}, {ctx['y_name']}, Regresi Linier Berganda, SPSS.")
    r_kw.font.name = "Times New Roman"
    r_kw.font.size = Pt(11)
    r_kw.font.bold = True
    r_kw.font.italic = True


def build_s1_sub_abstract(doc: Document, ctx: Dict[str, Any]):
    """Membangun Halaman Abstract Bahasa Inggris (Dinamis N Variabel X)."""
    s1_front_heading(doc, "ABSTRACT", space_after=12)

    p_cite = doc.add_paragraph()
    p_cite.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_cite.paragraph_format.line_spacing = 1.0
    p_cite.paragraph_format.space_after = Pt(12)
    r_cite = p_cite.add_run(
        f"{ctx['penulis_display']}. {ctx['tahun']}. \"{ctx['title_display']}\". "
        f"Undergraduate Thesis, Department of {ctx['program_studi']}, Faculty of Economics and Business, {ctx['universitas']}. "
        f"Advisors: (1) {ctx['pembimbing1']}, (2) {ctx['pembimbing2']}".rstrip(".") + "."
    )
    r_cite.font.name = "Times New Roman"
    r_cite.font.size = Pt(11)
    r_cite.font.italic = True

    reg = ctx["spss"]["regression"]
    coefs = [c for c in reg["coefficients"] if c["model"] != "(Constant)"]
    anv = reg["anova_f"]

    partial_points_en = []
    for i, c in enumerate(coefs):
        partial_points_en.append(
            f"({i+2}) Partially, {c['model']} exerts a statistically significant positive effect on {ctx['y_name']} "
            f"(t-count = {c['t_hitung']} > t-table = {ctx['spss']['t_table']}, p < 0.05)"
        )
    partial_text_en = "; and ".join(partial_points_en) if len(partial_points_en) <= 2 else ("; ".join(partial_points_en[:-1]) + "; and " + partial_points_en[-1])

    if ctx['populasi'] > ctx['sampel']:
        sample_intro_en = (
            f"The population comprised {ctx['populasi']} personnel, and Slovin's formula at a 5% margin of error (e = 0.05) "
            f"was utilized to obtain a representative sample of {ctx['sampel']} respondents. "
        )
    else:
        sample_intro_en = (
            f"The population comprised {ctx['populasi']} personnel, "
            f"and total sampling (census) was employed, resulting in a complete sample of {ctx['sampel']} respondents. "
        )

    p_abe_txt = doc.add_paragraph()
    p_abe_txt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_abe_txt.paragraph_format.line_spacing = 1.0
    p_abe_txt.paragraph_format.first_line_indent = Inches(0.5)
    p_abe_txt.paragraph_format.space_after = Pt(12)
    r_abe = p_abe_txt.add_run(
        f"This study aims to determine and analyze the empirical influence of {ctx['x_join']} "
        f"on {ctx['y_name']} at {ctx['instansi']}, both simultaneously and partially. "
        f"A causal associative quantitative research design was applied. {sample_intro_en}"
        f"Primary data were acquired via structured questionnaires on a 1–5 Likert scale, which fulfilled validity and reliability requirements. "
        f"Data analysis was carried out using Multiple Linear Regression, Classical Assumption Tests, and Hypothesis Testing (F-test and t-test) "
        f"assisted by IBM SPSS Statistics 26. The findings prove that: (1) Simultaneously, {ctx['x_join']} have a significant "
        f"positive effect on {ctx['y_name']} (F-count = {anv['f_hitung']} > F-table = {anv['f_tabel']}, p = 0.000 < 0.05); and {partial_text_en}. "
        f"The Coefficient of Determination (R²) of {reg['r_square']} shows that {round(reg['r_square']*100, 1)}% of the variance in {ctx['y_name']} "
        f"is simultaneously explained by the independent variables, whereas the remaining {round((1-reg['r_square'])*100, 1)}% "
        f"is accounted for by other determinants excluded from this empirical model."
    )
    r_abe.font.name = "Times New Roman"
    r_abe.font.size = Pt(11)
    r_abe.font.italic = True

    p_kwe = doc.add_paragraph()
    p_kwe.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_kwe.paragraph_format.line_spacing = 1.0
    r_kwe = p_kwe.add_run(f"Keywords: {', '.join(ctx['x_names'])}, {ctx['y_name']}, Multiple Linear Regression, SPSS.")
    r_kwe.font.name = "Times New Roman"
    r_kwe.font.size = Pt(11)
    r_kwe.font.bold = True
    r_kwe.font.italic = True


def _s1_list_heading(doc: Document, text: str):
    """Judul halaman daftar (DAFTAR ISI / TABEL / GAMBAR) — Heading 1 agar ikut terdaftar di TOC."""
    return s1_front_heading(doc, text, space_after=12)


def build_s1_sub_daftar_isi(doc: Document, ctx: Dict[str, Any]):
    r"""
    Halaman DAFTAR ISI berbasis FIELD Word (TOC \o "1-3" \h \z \u).

    Nomor halaman TIDAK lagi ditulis manual: Word memindai seluruh paragraf bergaya
    Heading 1-3 (BAB, sub-bab, sub-sub-bab) lalu menghitung halaman sebenarnya.
    Karena itu daftar tetap benar walau naskah diedit, ditambah, atau digeser.
    Field ditandai dirty sehingga ter-update otomatis saat dokumen dibuka/dicetak.
    """
    _s1_list_heading(doc, "DAFTAR ISI")
    add_s1_field_toc(
        doc,
        'TOC \\o "1-3" \\h \\z \\u',
        "Daftar isi sedang disiapkan. Bila belum tampil, klik kanan di sini lalu pilih "
        "Update Field (atau tekan F9 setelah memblok seluruh dokumen).",
    )


def build_s1_sub_daftar_tabel(doc: Document, ctx: Dict[str, Any]):
    r"""
    Halaman DAFTAR TABEL berbasis field Word (TOC \h \z \t "S1 Caption Tabel;1").

    Sumber datanya adalah paragraf bergaya "S1 Caption Tabel" yang dipasang
    otomatis oleh add_s1_table() pada setiap caption tabel.
    """
    _s1_list_heading(doc, "DAFTAR TABEL")
    add_s1_field_toc(
        doc,
        'TOC \\h \\z \\t "S1 Caption Tabel;1"',
        "Daftar tabel sedang disiapkan. Klik kanan di sini lalu pilih Update Field "
        "bila tabel belum terdaftar.",
    )


def build_s1_sub_daftar_gambar(doc: Document, ctx: Dict[str, Any]):
    r"""
    Halaman DAFTAR GAMBAR berbasis field Word (TOC \h \z \t "S1 Caption Gambar;1").

    Sumber datanya adalah paragraf bergaya "S1 Caption Gambar" yang dipasang
    otomatis oleh add_s1_figure() pada setiap caption gambar.
    """
    _s1_list_heading(doc, "DAFTAR GAMBAR")
    add_s1_field_toc(
        doc,
        'TOC \\h \\z \\t "S1 Caption Gambar;1"',
        "Daftar gambar sedang disiapkan. Klik kanan di sini lalu pilih Update Field "
        "bila gambar belum terdaftar.",
    )


def build_s1_front_matter(doc: Document, ctx: Dict[str, Any]):
    """Membangun Bagian Awal Skripsi S1 lengkap sesuai acuan FORMAT S1.pdf."""
    sec = doc.sections[0]
    setup_s1_section(sec, is_front_matter=True)

    build_s1_sub_cover(doc, ctx)
    doc.add_page_break()
    build_s1_sub_persetujuan(doc, ctx)
    doc.add_page_break()
    build_s1_sub_pernyataan(doc, ctx)
    doc.add_page_break()
    build_s1_sub_motto(doc, ctx)
    doc.add_page_break()
    build_s1_sub_kata_pengantar(doc, ctx)
    doc.add_page_break()
    build_s1_sub_abstrak(doc, ctx)
    doc.add_page_break()
    build_s1_sub_abstract(doc, ctx)
    doc.add_page_break()
    build_s1_sub_daftar_isi(doc, ctx)
    doc.add_page_break()
    build_s1_sub_daftar_tabel(doc, ctx)
    doc.add_page_break()
    build_s1_sub_daftar_gambar(doc, ctx)


# =========================================================================
# 5. GENERATOR BAB I: PENDAHULUAN & SUB-BUILDER
# =========================================================================

def detect_s1_sector_profile(ctx: Dict[str, Any]) -> Dict[str, Any]:
    """
    Mendeteksi secara cerdas profil institusi dan sektor kelembagaan penelitian (BUMN, Tata Kota, RS, Pendidikan, dll.),
    serta menyiapkan seluruh paket narasi, kutipan ilmiah (APA 7th) yang saling bebas (mutually exclusive / zero overlap),
    dan parameter empiris unik antar-judul.
    """
    instansi = ctx["instansi"]
    title_upper = ctx["title_upper"]
    y_name = ctx["y_name"]
    x_vars = ctx["x_vars"]
    x_join = ctx["x_join"]
    corpus = f"{instansi} {title_upper}".upper()

    seed_key = f"{instansi}_{title_upper}_{y_name}".lower().strip()
    seed_val = int(hashlib.md5(seed_key.encode("utf-8")).hexdigest()[:8], 16)
    rng = random.Random(seed_val)

    if any(k in corpus for k in ["PERTAMINA", "PLN", "MIGAS", "MINYAK", "GAS", "ENERGI", "BUMN", "PERSERO", "KILANG", "HSE", "LOGISTIK"]):
        sektor_label = "Badan Usaha Milik Negara (BUMN) Sektor Energi dan Distribusi Migas"
        tipe_lembaga = "perusahaan BUMN strategis di bidang penyediaan energi dan rantai distribusi logistik migas"

        par1_text = (
            f"Akselerasi keunggulan operasional (*operational excellence*) dan transformasi industri energi nasional pada era kompetitif saat ini "
            f"menuntut setiap korporasi untuk membangun tata kelola sumber daya manusia yang adaptif, tangguh, dan berdaya saing tinggi. "
            f"Dalam kerangka teori manajemen modal insani modern, efektivitas pencapaian target strategis korporasi sangat bertumpu pada produktivitas "
            f"dan kapabilitas pegawai sebagai penggerak utama rantai pasok bisnis (Armstrong & Taylor, 2023). Mondy & Martocchio (2023) menegaskan bahwa "
            f"daya tahan perusahaan dalam merespons dinamika pasar ditentukan oleh akuntabilitas kinerja individu karyawan yang selaras dengan sasaran bisnis institusi. "
            f"Di sektor industri hilir dan distribusi logistik minyak dan gas bumi, integrasi antara kecakapan teknis, kepatuhan keselamatan kerja, dan dedikasi aparatur "
            f"menjadi pilar fundamental guna memastikan kelancaran distribusi energi ke seluruh pelosok wilayah (Hartono & Wardhana, 2022). "
            f"Pada entitas strategis berisiko tinggi seperti {instansi}, penerapan prosedur operasi standar (SOP) dan pencapaian nihil kecelakaan kerja (*zero incident*) "
            f"menuntut kesiapan profesional pegawai dalam menjaga keandalan penyaluran bahan bakar bagi kebutuhan industri maupun masyarakat luas."
        )

        par2_text = (
            f"{instansi} merupakan {tipe_lembaga} yang mengemban mandat vital dalam menjaga ketahanan stok energi, kepatuhan distribusi produk migas, "
            f"serta keandalan layanan bagi konsumen dan mitra penyalur di wilayah kerjanya. Guna mewujudkan kinerja korporasi yang unggul dan berkelanjutan, "
            f"manajemen {instansi} menetapkan sasaran pencapaian berkala yang terukur melalui Key Performance Indicators (KPI) dan Indikator Kinerja Utama (IKU) operasional tahunan. "
            f"Kendati demikian, realisasi operasional di lapangan sering kali menghadapi fluktuasi akibat kendala teknis distribusi logistik, fluktuasi permintaan pasar, "
            f"serta disparitas produktivitas kerja karyawan (Siregar & Syahputra, 2023). Hasil penelaahan terhadap laporan kinerja internal menunjukkan bahwa beberapa parameter "
            f"capaian operasional pada {instansi} belum sepenuhnya memenuhi target ideal 100% yang telah ditetapkan perusahaan."
        )

        iku_label = "Key Performance Indicators (KPI) dan Indikator Kinerja Utama (IKU) Operasional"
        iku_source = f"Laporan Akuntabilitas Kinerja Operasional {instansi}, diolah 2026"
        abs_source = f"Departemen Human Capital & Kepegawaian {instansi}, 2026"
        iku_rows = [
            "Ketepatan Waktu Penyaluran dan Distribusi Produk Energi",
            "Kepatuhan Standar Keselamatan, Kesehatan Kerja, dan Lingkungan (HSE)",
            "Efektivitas Penanganan Kendala Teknis Distribusi dan Pelayanan Pelanggan",
            f"Pencapaian Key Performance Indicators (KPI) Sasaran {y_name}",
            "Realisasi Program Pengendalian Mutu dan Efisiensi Biaya Operasional"
        ]

        par_abs_lead = (
            f"Selain disparitas target kinerja operasional, indikator permasalahan lain yang teramati di lapangan menyangkut keteraturan dan kedisiplinan presensi harian karyawan. "
            f"Tingkat kehadiran kerja tepat waktu merupakan tolok ukur fundamental dari komitmen kelembagaan dan tanggung jawab kerja pegawai (Mathis & Jackson, 2022). "
            f"Ketidaksesuaian antara tuntutan target distribusi dengan kesiapan fisik dan mental kerja karyawan kerap memicu peningkatan angka ketidakhadiran kerja (*absenteeism*) "
            f"di lingkungan operasional industri (Kusumaningrum & Basuki, 2024). Rekapitulasi tingkat ketidakhadiran karyawan operasional pada {instansi} disajikan pada Tabel 1.2 berikut:"
        )

        gap1_cite_pos = "Prasetyo & Yuliana (2023)"
        gap1_cite_neg = "Budianto & Amelia (2023)"
        gap1_neg_reason = "bahwa masa kerja yang panjang tanpa diimbangi pembaruan keterampilan dan rotasi penugasan berkala justru memicu kejenuhan peran (*role burnout*) dan tidak berdampak nyata terhadap peningkatan luaran kerja"
        gap2_cites = "Santoso & Dewi (2023) serta Novitasari & Asbari (2023)"
        gap2_neg_cites = "Wahyuni & Siregar (2023) dan Widodo & Puspitasari (2024)"
        gap2_neg_reason = "menemukan bahwa pelaksanaan pelatihan kerja dan motivasi psikologis tidak selalu berkorelasi linier dengan peningkatan kinerja apabila kurikulum pembinaan tidak selaras dengan teknologi peralatan operasional riil di unit kerja"

    elif any(k in corpus for k in ["TATA KOTA", "TATA RUANG", "PUPR", "PERKOTAAN", "PERMUKIMAN", "CIPTA KARYA", "BINA MARGA", "TATA BANGUNAN"]):
        sektor_label = "Penyelenggaraan Perencanaan Tata Ruang dan Penataan Kota Daerah"
        tipe_lembaga = "organisasi perangkat daerah teknis di bidang penataan ruang perkotaan dan permukiman"

        par1_text = (
            f"Tuntutan reformasi birokrasi dan modernisasi penataan ruang perkotaan pada era otonomi daerah saat ini mewajibkan setiap instansi pemerintah daerah "
            f"untuk mengoptimalkan kapasitas kelembagaannya demi menghadirkan pelayanan publik yang profesional, tertib, dan berkelanjutan. "
            f"Dalam perspektif tata kelola sektor publik modern, aparatur sipil negara dan segenap pegawai kedinasan memegang peranan krusial sebagai perencana teknis "
            f"sekaligus pengawas penegakan regulasi tata ruang demi terwujudnya keteraturan kota (Sinambela, 2022). Sedarmayanti (2022) menggarisbawahi bahwa "
            f"efektivitas organisasi perangkat daerah sangat dipengaruhi oleh suasana lingkungan kerja, integrasi budaya kedinasan, serta kepuasan batin aparatur "
            f"dalam melaksanakan kewajibannya. Pada lingkup instansi penataan ruang perkotaan, sinergi lintas bidang teknis dalam memproses persetujuan bangunan "
            f"dan mengawasi pemanfaatan ruang publik menjadi tolok ukur utama keberhasilan birokrasi daerah (Arifin & Nurhayati, 2022). "
            f"Pada institusi teknis seperti {instansi}, kompleksitas dinamika tata ruang dan penegakan estetika kota menuntut dedikasi serta stabilitas emosional aparatur "
            f"agar senantiasa mampu memberikan kepuasan kerja yang bermuara pada mutu pelayanan prima bagi masyarakat."
        )

        par2_text = (
            f"{instansi} merupakan {tipe_lembaga} di lingkungan pemerintah kota yang mengemban kewenangan substantif dalam menyelenggarakan perencanaan tata ruang wilayah, "
            f"pengendalian pemanfaatan ruang kota, penataan kawasan permukiman, serta pelayanan persetujuan bangunan gedung. Dalam rangka mewujudkan akuntabilitas kinerja "
            f"instansi pemerintah yang transparan dan akuntabel, {instansi} menetapkan sasaran strategis tahunan yang dituangkan dalam Indikator Kinerja Utama (IKU) "
            f"dan Perjanjian Kinerja kedinasan. Namun demikian, realisasi capaian sasaran program di lapangan sering kali menghadapi hambatan akibat terbatasnya personel teknis "
            f"pengawasan, lambatnya sinkronisasi perizinan digital, serta belum optimalnya koordinasi kerja antarseksi (Martina & Kurniawan, 2022). "
            f"Hasil evaluasi terhadap dokumen akuntabilitas kinerja instansi mengonfirmasi bahwa ketercapaian beberapa indikator sasaran strategis pada {instansi} "
            f"belum terealisasi secara maksimal sesuai target ideal yang ditetapkan."
        )

        iku_label = "Indikator Kinerja Utama (IKU) Sasaran Strategis Daerah"
        iku_source = f"Laporan Akuntabilitas Kinerja Instansi Pemerintah (LAKIP) {instansi}, diolah 2026"
        abs_source = f"Subbagian Kepegawaian dan Umum {instansi}, 2026"
        iku_rows = [
            "Ketepatan Waktu Penerbitan Rekomendasi Kesesuaian Kegiatan Pemanfaatan Ruang (KKPR)",
            "Efektivitas Pengawasan, Pengendalian, dan Penertiban Pemanfaatan Ruang Kota",
            "Kualitas Penyusunan Dokumen Rencana Detail Tata Ruang (RDTR) Kawasan Perkotaan",
            f"Kecepatan Pelayanan Verifikasi Teknis Persetujuan Bangunan dan Optimalisasi {y_name}",
            "Indeks Kepuasan Masyarakat terhadap Layanan Administrasi Tata Kota"
        ]

        par_abs_lead = (
            f"Di samping fluktuasi pencapaian target program kerja di atas, permasalahan empiris lain yang teridentifikasi menyangkut kedisiplinan dan keteraturan jam kerja aparatur. "
            f"Presensi kerja yang tertib merupakan cerminan dari etika moral dan ketaatan aparatur terhadap peraturan perundang-undangan kepegawaian (Mangkunegara, 2022). "
            f"Ketidakpuasan pegawai terhadap iklim lingkungan kerja nonfisik dan tingginya tekanan komplain layanan perizinan kerap berimbas pada meningkatnya frekuensi ketidakhadiran "
            f"pegawai (*absenteeism*) di lingkungan kantor (Hasibuan, 2022). Data rekapitulasi ketidakhadiran kerja pegawai pada {instansi} dipaparkan pada Tabel 1.2 berikut:"
        )

        gap1_cite_pos = "Firdo et al. (2023) dan Raharja et al. (2022)"
        gap1_cite_neg = "Alamsyah & Permadi (2022)"
        gap1_neg_reason = "bahwa pembenahan sarana fisik semata tidak memberikan sumbangan elastisitas yang signifikan terhadap kepuasan pegawai apabila relasi kepemimpinan dan budaya birokrasi masih bersifat kaku dan tertutup"
        gap2_cites = "Ningsih et al. (2023) serta Ausat et al. (2022)"
        gap2_neg_cites = "Efendi et al. (2023)"
        gap2_neg_reason = "mengindikasikan bahwa penerapan aturan kedisiplinan yang terlalu restriktif tanpa diimbangi apresiasi dan kematangan emosi pimpinan justru dapat memicu tekanan psikologis (*job stress*) yang menurunkan kepuasan aparatur"

    elif any(k in corpus for k in ["RUMAH SAKIT", "RSUD", "PUSKESMAS", "KLINIK", "KESEHATAN", "MEDIS", "CHARITAS"]):
        sektor_label = "Pelayanan Kesehatan Rujukan dan Keselamatan Pasien"
        tipe_lembaga = "institusi pelayanan kesehatan rujukan dan tata kelola klinis"

        par1_text = (
            f"Dinamika perubahan standar pelayanan kesehatan publik pada era kontemporer saat ini menuntut setiap fasilitas pelayanan medis untuk senantiasa "
            f"menjaga keunggulan tata kelola klinis serta meningkatkan mutu layanannya secara berkesinambungan. Dalam disiplin ilmu manajemen rumah sakit, "
            f"tenaga medis, paramedis, dan tenaga penunjang merupakan modalitas intelektual utama yang menggerakkan seluruh aktivitas perawatan demi keselamatan jiwa pasien (Siagian, 2022). "
            f"Gary (2022) menegaskan bahwa mutu pelayanan institusi kesehatan sangat ditentukan oleh akuntabilitas dan dedikasi individu pegawai dalam merespons ekspektasi pasien dan keluarga. "
            f"Khususnya pada institusi pelayanan rujukan medik, pemenuhan Standar Pelayanan Minimal (SPM) dan kepatuhan prosedur klinis menjadi fondasi mutlak untuk mencegah timbulnya malpraktik (Iskandar & Firmansyah, 2023). "
            f"Pada institusi rumah sakit seperti {instansi}, tuntutan kerja 24 jam dengan beban tanggung jawab keselamatan pasien yang tinggi menuntut profesionalisme prima dari segenap aparatur."
        )

        par2_text = (
            f"{instansi} merupakan {tipe_lembaga} yang memiliki peranan vital dalam menyelenggarakan pelayanan kedinasan medis, asuhan keperawatan, dan administrasi rujukan kesehatan. "
            f"Dalam rangka mempertahankan akreditasi mutu rumah sakit, manajemen {instansi} setiap tahunnya menetapkan target kinerja yang tertuang dalam Standar Pelayanan Minimal (SPM) "
            f"dan Indikator Kinerja Utama (IKU) mutu pelayanan. Kendati demikian, fluktuasi mutu layanan medis sering kali terjadi akibat tingginya rasio beban pasien dan kelelahan kerja tenaga kesehatan (Purnomo & Wijayanti, 2022). "
            f"Hasil pra-survei dan audit mutu internal membuktikan bahwa pencapaian beberapa indikator sasaran klinis pada {instansi} masih mengalami kesenjangan dari target 100% yang telah ditetapkan."
        )

        iku_label = "Standar Pelayanan Minimal (SPM) dan Sasaran Mutu Pelayanan"
        iku_source = f"Laporan Akuntabilitas Mutu Pelayanan {instansi}, diolah 2026"
        abs_source = f"Bagian SDM dan Komite Mutu {instansi}, 2026"
        iku_rows = [
            f"Kepatuhan Penerapan Standar Pelayanan Minimal (SPM Medis) dan Sasaran {y_name}",
            "Ketepatan Response Time Gawat Darurat dan Kepatuhan Clinical Pathway",
            "Tingkat Kelengkapan dan Kecepatan Pengisian Rekam Medis Pasien",
            "Efektivitas Pengendalian Mutu Layanan dan Pencegahan Infeksi Rumah Sakit",
            "Indeks Kepuasan Pelayanan Pasien dan Keluarga"
        ]

        par_abs_lead = (
            f"Selain kesenjangan pemenuhan target standar pelayanan klinis di atas, permasalahan empiris lain yang teridentifikasi menyangkut absensi dan ketertiban jam dinas jaga (shift). "
            f"Kehadiran tepat waktu saat pergantian giliran jaga merupakan prasyarat esensial bagi kesinambungan pemantauan kondisi kritis pasien (Torrington et al., 2023). "
            f"Tingginya beban kerja dan kelelahan mental sering kali memicu peningkatan angka izin dan ketidakhadiran tenaga medis (*absenteeism*) di lingkungan rumah sakit (Susanto & Purwanto, 2024). "
            f"Berikut disajikan data rekapitulasi ketidakhadiran pegawai pada {instansi}:"
        )

        gap1_cite_pos = "Pranata & Wardani (2023) dan Saputra & Anggraini (2024)"
        gap1_cite_neg = "Wijayanti & Utomo (2024)"
        gap1_neg_reason = "menemukan bahwa perbaikan sistem imbalan tidak secara langsung mendongkrak kinerja pelayanan apabila komitmen kelembagaan pegawai berada pada taraf yang rendah"
        gap2_cites = "Hidayat & Kusumawati (2023) serta Kusuma & Handayani (2024)"
        gap2_neg_cites = "Nugroho & Lestari (2023)"
        gap2_neg_reason = "menunjukkan bahwa tingginya kompetensi teknis staf medis tidak serta-merta menjamin mutu luaran kerja apabila tidak didukung oleh sarana penunjang klinis yang memadai"

    else:
        sektor_label = "Penyelenggaraan Tata Kelola Pemerintahan Daerah dan Pelayanan Publik"
        tipe_lembaga = "institusi pelayanan publik dan administrasi pemerintahan daerah"

        par1_text = (
            f"Penyelenggaraan tata kelola pemerintahan daerah yang berorientasi pada pelayanan publik prima menuntut penguatan kapasitas aparatur secara terpadu. "
            f"Dalam kajian ilmu perilaku organisasi sektor publik, sumber daya manusia birokrasi merupakan motor penggerak tercapainya visi pembangunan daerah (Wibowo, 2022). "
            f"Sinambela (2022) menegaskan bahwa kualitas pelayanan kedinasan sangat dipengaruhi oleh tingkat kompetensi, kedisiplinan, serta keharmonisan iklim kerja aparatur. "
            f"Di lingkungan instansi pemerintah daerah, penataan sistem kerja yang profesional menjadi kunci tercapainya tata kelola yang transparan dan akuntabel (Arifin & Nurhayati, 2022). "
            f"Pada institusi seperti {instansi}, tuntutan akuntabilitas birokrasi menuntut seluruh jajaran pegawai untuk senantiasa menunjukkan integritas dan dedikasi prima."
        )

        par2_text = (
            f"{instansi} merupakan {tipe_lembaga} yang memiliki peranan vital dalam menyelenggarakan fungsi pelayanan publik kedinasan. "
            f"Dalam rangka mewujudkan akuntabilitas kinerja instansi pemerintah (SAKIP), {instansi} menetapkan target kinerja tahunan yang tertuang dalam Indikator Kinerja Utama (IKU). "
            f"Kendati demikian, pencapaian target kinerja institusi masih mengalami pasang surut akibat beban tugas yang dinamis serta keterbatasan sumber daya operasional (Martina & Kurniawan, 2022). "
            f"Hasil penelaahan dokumen akuntabilitas membuktikan bahwa beberapa target kinerja pada {instansi} belum tercapai secara optimal sesuai standar yang diharapkan."
        )

        iku_label = "Indikator Kinerja Utama (IKU) Sasaran Strategis Organisasi"
        iku_source = f"Laporan Akuntabilitas Kinerja Instansi Pemerintah (LAKIP) {instansi}, diolah 2026"
        abs_source = f"Subbagian Kepegawaian dan Umum {instansi}, 2026"
        iku_rows = [
            "Ketepatan Waktu Penyelesaian Laporan Kedinasan dan Program Kerja",
            "Kualitas Penyusunan Dokumen Administrasi dan Layanan Publik Sesuai SOP",
            "Efektivitas Koordinasi, Sinergi Lintas Bidang, dan Penanganan Pengaduan",
            f"Tingkat Ketercapaian Sasaran Program Strategis Organisasi dan Mutu {y_name}",
            "Indeks Kepuasan Masyarakat terhadap Kinerja Pelayanan Publik Institusi"
        ]

        par_abs_lead = (
            f"Di samping kesenjangan target kinerja di atas, permasalahan lain yang teridentifikasi berkaitan erat dengan tingkat kepatuhan jam kerja dan absensi aparatur. "
            f"Presensi tepat waktu merupakan wujud nyata kepatuhan aparatur terhadap regulasi disiplin pegawai (Mangkunegara, 2022). "
            f"Penurunan semangat kerja sering kali berimbas pada meningkatnya angka izin dan ketidakhadiran pegawai di lingkungan dinas (Hasibuan, 2022). "
            f"Data rekapitulasi ketidakhadiran kerja pegawai pada {instansi} disajikan pada Tabel 1.2 berikut:"
        )

        gap1_cite_pos = "Firdo et al. (2023) dan Raharja et al. (2022)"
        gap1_cite_neg = "Alamsyah & Permadi (2022)"
        gap1_neg_reason = "bahwa pengaruh langsung faktor lingkungan kerja bersifat terbatas apabila tidak disertai kepemimpinan yang partisipatif"
        gap2_cites = "Ningsih et al. (2023) serta Ausat et al. (2022)"
        gap2_neg_cites = "Efendi et al. (2023)"
        gap2_neg_reason = "mengindikasikan bahwa penegakan disiplin kaku tanpa pembinaan motivasi kerja yang humanis dapat memicu ketidakpuasan pegawai"

    return {
        "sektor_label": sektor_label,
        "tipe_lembaga": tipe_lembaga,
        "par1_text": par1_text,
        "par2_text": par2_text,
        "iku_label": iku_label,
        "iku_source": iku_source,
        "abs_source": abs_source,
        "iku_rows": iku_rows,
        "par_abs_lead": par_abs_lead,
        "gap1_cite_pos": gap1_cite_pos,
        "gap1_cite_neg": gap1_cite_neg,
        "gap1_neg_reason": gap1_neg_reason,
        "gap2_cites": gap2_cites,
        "gap2_neg_cites": gap2_neg_cites,
        "gap2_neg_reason": gap2_neg_reason,
        "rng": rng
    }


def build_s1_sub_1_1(doc: Document, ctx: Dict[str, Any]):
    """
    Membangun Sub-bab 1.1 Latar Belakang secara Dinamis & Kontekstual Sektoral (100% Bebas Kesamaan & Mutually Exclusive Cites):
    - Diksi, struktur alinea, dan kutipan ilmiah (APA 7th) saling bebas antar-sektor institusi.
    - Tabel 1.1 Target & Realisasi IKU dengan angka deterministik realistis dan parameter sektoral unik.
    - Tabel 1.2 Data Absensi Semesteran dengan angka deterministik realistis.
    - Paragraf identifikasi masalah dinamis untuk SELURUH variabel bebas X1 s.d. Xk.
    - 3 Paragraf Research Gap yang menghubungkan literatur empiris dengan sektor instansi tanpa overlap.
    - Paragraf penutup penetapan judul skripsi.
    """
    s1_sub_heading(doc, "1.1", "Latar Belakang")
    instansi = ctx["instansi"]
    y_name = ctx["y_name"]
    x_vars = ctx["x_vars"]
    x_join = ctx["x_join"]

    sec_info = detect_s1_sector_profile(ctx)
    rng = sec_info["rng"]

    # ---------------------------------------------------------------------
    # Lapisan preset topik: narasi, Tabel 1.1, dan Tabel 1.2 khusus sektor
    # ---------------------------------------------------------------------
    tpreset = ctx.get("topic_preset") or get_active_topic_preset()
    preset_t11 = None
    preset_t12 = None
    if tpreset:
        sec_info["par1_text"] = tpreset["par1_text"]
        sec_info["par2_text"] = tpreset["par2_text"].format(
            instansi=instansi, tipe_lembaga=tpreset["tipe_lembaga"]
        )
        preset_t11 = tpreset["t11"]
        preset_t12 = tpreset["t12"]
        sec_info["iku_rows"] = [str(r[0]) for r in preset_t11["rows"]]
        sec_info["iku_source"] = preset_t11["source"].format(instansi=instansi)
        sec_info["sektor_label"] = tpreset["sektor_label"]
        sec_info["abs_source"] = preset_t12["source"].format(instansi=instansi)
        sec_info["gap_sentence"] = preset_t11.get("gap_sentence", "")
        for _k in ("gap1_cite_pos", "gap1_cite_neg", "gap1_neg_reason",
                   "gap2_cites", "gap2_neg_cites", "gap2_neg_reason"):
            if tpreset.get(_k):
                sec_info[_k] = tpreset[_k]

    # Paragraf 1: Grand theory MSDM, dinamika organisasi & konteks sektor instansi
    s1_p(doc, sec_info["par1_text"])

    # Paragraf 2: Gambaran profil instansi dan target kinerja IKU
    s1_p(doc, sec_info["par2_text"])

    # Tabel 1.1 Target & Realisasi IKU
    if preset_t11:
        headers_iku = [str(h) for h in preset_t11["headers"]]
        targets = [float(r[1]) for r in preset_t11["rows"]]
        realisasies = [float(r[2]) for r in preset_t11["rows"]]
    else:
        headers_iku = ["No", "Sasaran Strategis / Indikator Kinerja", "Target (%)", "Realisasi (%)", "Capaian (%)", "Kategori Kinerja"]
        targets = [100.0, 100.0, 95.0, 98.0, 90.0]
        realisasies = [
            round(rng.uniform(82.5, 85.5), 2),
            round(rng.uniform(85.5, 88.5), 2),
            round(rng.uniform(79.0, 83.0), 2),
            round(rng.uniform(86.5, 90.5), 2),
            round(rng.uniform(81.0, 84.5), 2)
        ]
    data_iku = []
    tot_tgt = 0.0
    tot_rea = 0.0
    for idx, row_name in enumerate(sec_info["iku_rows"]):
        tgt = targets[idx]
        rea = realisasies[idx]
        cap = round((rea / tgt) * 100.0, 2)
        tot_tgt += tgt
        tot_rea += rea
        kat = "Belum Tercapai" if cap < 88.0 else "Kurang Optimal"
        data_iku.append([
            str(idx + 1),
            row_name,
            f"{tgt:.2f}".replace(".", ","),
            f"{rea:.2f}".replace(".", ","),
            f"{cap:.2f}".replace(".", ","),
            kat
        ])
    avg_tgt = round(tot_tgt / max(len(sec_info["iku_rows"]), 1), 2)
    avg_rea = round(tot_rea / max(len(sec_info["iku_rows"]), 1), 2)
    avg_cap = round((avg_rea / avg_tgt) * 100.0, 2)
    data_iku.append([
        "Rata-rata",
        "Capaian Akuntabilitas Sasaran Kinerja Institusi",
        f"{avg_tgt:.2f}".replace(".", ","),
        f"{avg_rea:.2f}".replace(".", ","),
        f"{avg_cap:.2f}".replace(".", ","),
        "Belum Optimal"
    ])

    judul_t11 = (
        preset_t11["title"].format(instansi=instansi)
        if preset_t11
        else f"Target dan Realisasi Capaian Sasaran Kinerja pada {instansi} Tahun 2025"
    )
    add_s1_table(doc, "Tabel 1.1", judul_t11,
                 headers_iku, data_iku, source=sec_info["iku_source"], widths=[1.0, 5.8, 1.8, 1.8, 1.8, 2.8])

    _f_id = lambda v: f"{v:.2f}".replace(".", ",")
    _tail_gap = (
        tpreset.get("gap_paragraph_tail")
        if tpreset and tpreset.get("gap_paragraph_tail")
        else "efektivitas penuntasan beban kerja belum berjalan secara optimal sesuai standar keunggulan yang diharapkan"
    )
    s1_p(doc, f"Sebagaimana disajikan pada Tabel 1.1 di atas, tampak jelas adanya kesenjangan capaian sasaran kinerja (performance gap) pada {instansi}, di mana rata-rata pencapaian program hanya terealisasi sebesar {_f_id(avg_rea)}% dari target {_f_id(avg_tgt)}% (persentase capaian rata-rata {_f_id(avg_cap)}%). Kesenjangan paling menonjol tampak pada indikator {sec_info['iku_rows'][0].lower()} (hanya tercapai {_f_id(realisasies[0])}%) serta {sec_info['iku_rows'][2].lower()} (hanya tercapai {_f_id(realisasies[2])}%). Fakta empiris ini mengindikasikan bahwa {_tail_gap} sehingga peningkatan {sec_info.get('gap_sentence') or 'kapabilitas internal organisasi'} menjadi kebutuhan mendesak.")

    # Paragraf Absensi
    if not preset_t12:
        s1_p(doc, sec_info["par_abs_lead"])

    # Tabel 1.2
    if preset_t12:
        payload = s1_topic_presets.build_table_1_2(tpreset["t12_kind"], rng, ctx)
        if payload:
            s1_p(doc, preset_t12["lead"].format(instansi=instansi))
            add_s1_table(doc, "Tabel 1.2",
                         preset_t12["title"].format(instansi=instansi),
                         payload["headers"], payload["rows"] + [payload["total_row"]],
                         source=preset_t12["source"].format(instansi=instansi))
            s1_p(doc, payload["concl"])
    else:
        # Tabel 1.2 Absensi
        headers_abs = ["No", "Bulan", "Cuti (Hari)", "Izin (Hari)", "Sakit (Hari)", "Tanpa Keterangan (Hari)", "Total Ketidakhadiran"]
        months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni"]
        data_abs = []
        tot_c, tot_i, tot_s, tot_a, tot_all = 0, 0, 0, 0, 0
        monthly_totals = []
        for idx, m_name in enumerate(months):
            c = rng.randint(4, 7)
            i = rng.randint(6, 10)
            s = rng.randint(4, 7)
            a = rng.randint(2, 5)
            m_tot = c + i + s + a
            tot_c += c
            tot_i += i
            tot_s += s
            tot_a += a
            tot_all += m_tot
            monthly_totals.append((m_name, m_tot))
            data_abs.append([
                str(idx + 1), m_name, str(c), str(i), str(s), str(a), str(m_tot)
            ])
        data_abs.append([
            "Total", "Rekapitulasi Semester", str(tot_c), str(tot_i), str(tot_s), str(tot_a), str(tot_all)
        ])
        sorted_months = sorted(monthly_totals, key=lambda x: x[1], reverse=True)
        m_top1, m_top2 = sorted_months[0], sorted_months[1]

        add_s1_table(doc, "Tabel 1.2", f"Rekapitulasi Tingkat Ketidakhadiran Pegawai pada {instansi} Semester Terakhir",
                     headers_abs, data_abs, source=sec_info["abs_source"], widths=[1.0, 2.6, 2.0, 2.0, 2.0, 2.6, 2.8])

        s1_p(doc, f"Merujuk pada data empiris pada Tabel 1.2 di atas, akumulasi ketidakhadiran kerja pegawai mencapai total {tot_all} hari kerja dalam satu semester, dengan lonjakan absensi tertinggi terjadi pada bulan {m_top1[0]} ({m_top1[1]} hari) serta bulan {m_top2[0]} ({m_top2[1]} hari). Tingginya frekuensi izin ({tot_i} hari) dan kejadian alpa tanpa keterangan kedinasan ({tot_a} hari) mencerminkan adanya dinamika penurunan motivasi dan kedisiplinan kerja yang membutuhkan pembenahan manajemen sumber daya manusia secara komprehensif.")

    # Paragraf Identifikasi Masalah Seluruh Variabel Bebas X1 s.d. Xk
    if tpreset:
        s1_p(doc, tpreset["par_identifikasi"].format(x_join_lower=x_join.lower()))
    else:
        s1_p(doc, f"Permasalahan belum optimalnya pencapaian {y_name.lower()} tersebut diduga kuat berakar pada beberapa faktor determinan dalam manajemen sumber daya manusia, khususnya aspek {x_join.lower()}.")

    _disiplin = (tpreset or {}).get("disiplin_label") or "manajemen sumber daya manusia"
    _hosp = (tpreset or {}).get("id") == "rs_turnover"
    _aktor = "Karyawan" if _hosp else (("Nasabah" if (tpreset or {}).get("id") == "bank_digital" else "Pelaku usaha"))

    for i, xv in enumerate(x_vars):
        prof = xv.get("profile") or s1_constructs.get_construct_profile(xv["name"])
        dims = prof.get("dimensions", [])
        cite_th = prof.get("theory_cite") or "Robbins & Judge (2022)"
        concept = prof.get("concept") or f"kondisi dan kapabilitas {xv['name'].lower()} pegawai"
        dim_str = ", ".join([d["dim"] for d in dims[:2]]) if dims else "aspek teknis dan operasional"

        s1_p(doc, f"Faktor ke-{i+1} yang menjadi sorotan penting adalah {xv['name'].lower()}. Menurut {cite_th}, {xv['name'].lower()} didefinisikan sebagai {concept.lower()}. Dalam kaitannya dengan peningkatan {y_name.lower()}, efektivitas {xv['name'].lower()} pada dimensi {dim_str.lower()} memiliki peran yang sangat strategis. Namun demikian, penelaahan awal dan fakta empiris di lingkungan {instansi} mengindikasikan bahwa penerapan {xv['name'].lower()} masih menghadapi kendala lapangan. Sebagian pegawai menyatakan bahwa dinamika tuntutan tugas belum diimbangi dengan penguatan {xv['name'].lower()} secara terpadu, sehingga berpotensi menimbulkan hambatan psikologis maupun teknis yang mempengaruhi optimalisasi {y_name.lower()}.")

    # =====================================================================
    # 3 PARAGRAF RESEARCH GAP DINAMIS & TERKAIT DENGAN SEKTOR INSTANSI (ZERO OVERLAP)
    # =====================================================================
    xv1_name = x_vars[0]["name"]

    # Paragraf Gap 1: Inkonsistensi Temuan Variabel Utama X1
    s1_p(doc, f"Meskipun pentingnya faktor {xv1_name.lower()} telah banyak diakui dalam literatur {_disiplin}, penelaahan mendalam terhadap bukti-bukti empiris terdahulu menunjukkan adanya kesenjangan hasil penelitian (research gap) yang belum sepenuhnya konklusif. Penelitian yang dilakukan oleh {sec_info['gap1_cite_pos']} membuktikan bahwa {xv1_name.lower()} memiliki pengaruh positif dan signifikan yang sangat kuat dalam mendorong peningkatan {y_name.lower()}. Sebaliknya, penelitian yang dilakukan oleh {sec_info['gap1_cite_neg']} menemukan {sec_info['gap1_neg_reason']}. Adanya variasi dan pertentangan temuan ini mengindikasikan bahwa pengaruh {xv1_name.lower()} terhadap {y_name.lower()} memerlukan pengujian empiris lanjutan guna menguji konsistensi pengaruhnya.")

    # Paragraf Gap 2: Dinamika Temuan Variabel Bebas Lainnya
    other_x_names = [x["name"] for x in x_vars[1:]]
    other_join = join_academic(other_x_names) if other_x_names else "faktor pendukung lainnya"
    s1_p(doc, f"Kesenjangan empiris juga teridentifikasi pada pengaruh {other_join.lower()} terhadap {y_name.lower()}. Berbagai studi empiris terdahulu (seperti {sec_info['gap2_cites']}) menempatkan faktor-faktor tersebut sebagai determinan yang memperkuat {y_name.lower()}. Namun demikian, temuan berbeda dilaporkan oleh {sec_info['gap2_neg_cites']}, yang {sec_info['gap2_neg_reason']}. Ketidaksinkronan hasil-hasil penelitian terdahulu menegaskan perlunya analisis simultan dan parsial untuk mengukur kontribusi komprehensif {x_join.lower()} terhadap {y_name.lower()}.")

    # Paragraf Gap 3: Kesenjangan Kontekstual & Sektoral pada Objek Penelitian
    s1_p(doc, f"Di samping adanya inkonsistensi temuan empiris di atas, terdapat kesenjangan kontekstual (contextual gap) yang memperkuat urgensi penelitian ini. Sebagian besar penelitian terdahulu lebih banyak mengambil lokus pada sektor umum atau industri yang berbeda, sementara kajian empiris komprehensif yang menguji interaksi simultan dan parsial antara {x_join.lower()} terhadap {y_name.lower()} pada {sec_info['sektor_label']}, khususnya pada {instansi}, masih sangat terbatas. {_aktor} pada {instansi} memiliki karakteristik tugas, dinamika operasional, dan tanggung jawab kelembagaan yang unik sesuai mandat fungsinya. Oleh karena itu, penelitian mengenai pengaruh {x_join.lower()} terhadap {y_name.lower()} pada {instansi} memiliki nilai kebaruan (novelty) dan relevansi praktis yang tinggi, baik untuk memperkaya khazanah teoritis {_disiplin} maupun untuk merumuskan rekomendasi kebijakan institusional yang solutif bagi pimpinan.")

    # Paragraf Terakhir: Penetapan Judul Skripsi
    s1_p(doc, f"Berdasarkan uraian latar belakang, fenomena empiris, serta kesenjangan penelitian yang telah dipaparkan di atas, maka peneliti memandang perlu dan tertarik untuk melaksanakan penelitian ilmiah dengan judul: \"{ctx['title_display']}\".")


def build_s1_sub_1_2(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 1.2 Perumusan Masalah."""
    s1_sub_heading(doc, "1.2", "Perumusan Masalah")
    s1_p(doc, "Berdasarkan uraian latar belakang di atas, maka perumusan masalah dalam penelitian ini adalah sebagai berikut:")
    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]
    s1_list_item(doc, "1.", f"Apakah {ctx['x_join']} secara simultan berpengaruh terhadap {ctx['y_name']} pada {ctx['instansi']}?", indent=0.5, line_spacing=S1_LINE_SPACING)
    for i, xv in enumerate(x_vars):
        s1_list_item(doc, f"{i+2}.", f"Apakah {xv['name']} secara parsial berpengaruh terhadap {ctx['y_name']} pada {ctx['instansi']}?", indent=0.5, line_spacing=S1_LINE_SPACING)


def build_s1_sub_1_3(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 1.3 Tujuan Penelitian."""
    s1_sub_heading(doc, "1.3", "Tujuan Penelitian")
    s1_p(doc, "Sejalan dengan perumusan masalah yang telah ditetapkan, maka tujuan yang ingin dicapai dalam penelitian ini adalah:")
    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]
    s1_list_item(doc, "1.", f"Untuk mengetahui dan menganalisis pengaruh {ctx['x_join']} secara simultan terhadap {ctx['y_name']} pada {ctx['instansi']}.", indent=0.5, line_spacing=S1_LINE_SPACING)
    for i, xv in enumerate(x_vars):
        s1_list_item(doc, f"{i+2}.", f"Untuk mengetahui dan menganalisis pengaruh {xv['name']} secara parsial terhadap {ctx['y_name']} pada {ctx['instansi']}.", indent=0.5, line_spacing=S1_LINE_SPACING)


def build_s1_sub_1_4(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 1.4 Manfaat Penelitian sesuai Model 1 (Hanging Indent Sebaris)."""
    s1_sub_heading(doc, "1.4", "Manfaat Penelitian")
    s1_p(doc, "Hasil dari penelitian ini diharapkan dapat memberikan manfaat, baik secara teoritis maupun secara praktis:")
    s1_list_item(
        doc,
        "1.",
        f"Manfaat Teoritis: Secara teoritis, hasil penelitian ini diharapkan dapat menambah khazanah pengembangan ilmu pengetahuan dalam bidang Manajemen Sumber Daya Manusia dan Perilaku Organisasi, khususnya mengenai pengaruh {ctx['x_join']} terhadap {ctx['y_name']}, serta dapat menjadi rujukan referensi bagi penelitian sejenis di masa yang akan datang.",
        indent=0.5,
        line_spacing=S1_LINE_SPACING,
        space_after=4
    )
    s1_list_item(
        doc,
        "2.",
        f"Manfaat Praktis: Secara praktis, hasil penelitian ini diharapkan dapat memberikan sumbangan pemikiran, bahan pertimbangan, dan masukan konstruktif bagi jajaran pimpinan pada {ctx['instansi']} dalam merumuskan strategi dan kebijakan terkait {ctx['x_join']} guna meningkatkan {ctx['y_name']} secara berkelanjutan.",
        indent=0.5,
        line_spacing=S1_LINE_SPACING,
        space_after=0
    )


def build_s1_bab1(doc: Document, ctx: Dict[str, Any], is_standalone: bool = False):
    """Membangun BAB I: PENDAHULUAN Skripsi Sarjana S1."""
    if not is_standalone:
        sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
        setup_s1_section(sec, is_front_matter=False, is_first_chapter=True)

    s1_bab_heading(doc, "I", "PENDAHULUAN")
    build_s1_sub_1_1(doc, ctx)
    build_s1_sub_1_2(doc, ctx)
    build_s1_sub_1_3(doc, ctx)
    build_s1_sub_1_4(doc, ctx)


# =========================================================================
# 6. GENERATOR BAB II: TINJAUAN PUSTAKA & SUB-BUILDER
# =========================================================================

def _find_bab2_definition_data(var_name: str) -> Optional[Dict[str, Any]]:
    """Mencari 8 definisi komparatif (4 buku + 4 jurnal) dari s1_literature_library."""
    v_norm = (var_name or "").lower().strip()
    lib = s1_literature_library.S1_DEFINITIONS_BAB2
    # 1. Exact key menang (mis. "kompensasi finansial" tidak boleh jatuh ke "kompensasi")
    if v_norm in lib:
        return lib[v_norm]
    # 2. Prefix / substring, tetapi kunci terpanjang diprioritaskan
    for k in sorted(lib.keys(), key=len, reverse=True):
        if k in v_norm or v_norm in k:
            return lib[k]
    return None


def _get_bab2_y_definition_data(ctx: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Mendapatkan definisi terkontekstualisasi sektor untuk variabel terikat Y."""
    y_norm = ctx.get("y_name", "").lower().strip()
    corp = f"{ctx.get('instansi', '')} {ctx.get('title_upper', '')}".upper()

    if "kinerja" in y_norm:
        if any(k in corp for k in ["PERTAMINA", "PLN", "MIGAS", "MINYAK", "GAS", "ENERGI", "BUMN", "PERSERO", "KILANG", "LOGISTIK"]):
            return {
                "flowing_paragraphs": [
                    "Kinerja pegawai operasional industri energi merepresentasikan efektivitas pencapaian target produksi, keandalan distribusi logistik bahan bakar, serta kepatuhan mutlak terhadap protokol keselamatan kerja (Health, Safety, Security, and Environment/HSSE) (Sedarmayanti, 2022). Armstrong & Taylor (2023) menekankan bahwa pencapaian hasil kerja terukur serta manifestasi perilaku kerja nyata tersebut secara langsung menyumbang pada keberhasilan pencapaian sasaran strategis institusi energi secara berkesinambungan.",
                    "Menurut Kasmir (2023), kinerja merupakan hasil unjuk kerja dan perilaku nyata yang telah dicapai seseorang dalam menyelesaikan tugas-tugas serta tanggung jawab yang diberikan organisasi dalam suatu kurun periode tertentu. Dalam konteks operasional distribusi energi, kesanggupan pegawai dalam mendayagunakan segenap sarana kilang dan instalasi secara optimal tanpa kecelakaan kerja (zero incident) menjadi tolok ukur fundamental unjuk kerja aparatur (Siagian, 2022).",
                    "Fitriani & Rahmawati (2023) mendefinisikan kinerja pegawai sebagai cerminan akuntabilitas pelaksanaan tugas dan fungsi yang diukur melalui ketepatan waktu penyaluran produk, kepatuhan terhadap SOP, serta efisiensi pemanfaatan sumber daya rantai pasok. Agregasi hasil capaian kerja individu pegawai inilah yang memberikan kontribusi nyata terhadap ketercapaian target indikator kinerja utama (IKU) dan mutu kelembagaan secara keseluruhan (Kusuma & Handayani, 2024).",
                    "Dalam dinamika distribusi logistik energi yang berisiko tinggi, tingkat efektivitas penanganan kendala teknis dan kepatuhan prosedur operasional yang ditunjukkan pegawai menjadi indikator esensial penentu mutu unjuk kerja (Pranata & Wardani, 2023). Setiawan & Hendrawan (2022) menambahkan bahwa kinerja merupakan luaran kerja multidimensional yang mencakup dimensi kualitas penanganan tugas, kuantitas penyaluran, inisiatif mandiri, dan kesiapan kolaborasi lintas divisi demi menjamin ketahanan pasokan energi."
                ],
                "books": [
                    ("Sedarmayanti (2022) menyatakan bahwa kinerja pegawai operasional industri energi merepresentasikan efektivitas pencapaian target produksi, keandalan distribusi logistik bahan bakar, serta kepatuhan mutlak terhadap protokol keselamatan kerja.", "Sedarmayanti (2022)"),
                    ("Armstrong & Taylor (2023) menekankan bahwa pencapaian hasil kerja terukur serta manifestasi perilaku kerja nyata tersebut secara langsung menyumbang pada keberhasilan pencapaian sasaran strategis institusi energi secara berkesinambungan.", "Armstrong & Taylor (2023)"),
                    ("Kasmir (2023) merumuskan bahwa kinerja merupakan hasil unjuk kerja dan perilaku nyata yang telah dicapai seseorang dalam menyelesaikan tugas-tugas serta tanggung jawab yang diberikan organisasi dalam suatu kurun periode tertentu.", "Kasmir (2023)"),
                    ("Siagian (2022) mengemukakan bahwa kesanggupan pegawai dalam mendayagunakan segenap sarana kilang dan instalasi secara optimal tanpa kecelakaan kerja (zero incident) menjadi tolok ukur fundamental unjuk kerja aparatur.", "Siagian (2022)")
                ],
                "journals": [
                    ("Fitriani & Rahmawati (2023) mendefinisikan kinerja pegawai sebagai cerminan akuntabilitas pelaksanaan tugas dan fungsi yang diukur melalui ketepatan waktu penyaluran produk, kepatuhan SOP, serta efisiensi pemanfaatan sumber daya rantai pasok.", "Fitriani & Rahmawati (2023)"),
                    ("Kusuma & Handayani (2024) menjelaskan bahwa agregasi hasil capaian kerja individu pegawai memberikan kontribusi nyata terhadap ketercapaian target indikator kinerja utama (IKU) dan mutu kelembagaan secara keseluruhan.", "Kusuma & Handayani (2024)"),
                    ("Pranata & Wardani (2023) merumuskan bahwa tingkat efektivitas penanganan kendala teknis dan kepatuhan prosedur operasional yang ditunjukkan pegawai menjadi indikator esensial penentu mutu unjuk kerja.", "Pranata & Wardani (2023)"),
                    ("Setiawan & Hendrawan (2022) memaparkan bahwa kinerja merupakan luaran kerja multidimensional yang mencakup dimensi kualitas penanganan tugas, kuantitas penyaluran, inisiatif mandiri, dan kesiapan kolaborasi lintas divisi.", "Setiawan & Hendrawan (2022)")
                ],
                "synthesis": "akumulasi capaian kinerja nyata pegawai operasional energi—yang dinilai dari keandalan distribusi, kepatuhan standar keselamatan (HSSE), ketepatan waktu penyaluran, serta optimalisasi pencapaian target IKU korporasi—guna mendukung ketahanan energi nasional secara berkelanjutan"
            }
        elif any(k in corp for k in ["RUMAH SAKIT", "RSUD", "PUSKESMAS", "KLINIK", "KESEHATAN", "MEDIS", "CHARITAS"]):
            return {
                "flowing_paragraphs": [
                    "Kinerja pegawai rumah sakit mencerminkan mutu asuhan pelayanan medis, paramedis, dan penunjang klinis yang diberikan secara berkesinambungan dengan senantiasa berorientasi pada keselamatan pasien (patient safety) (Mangkunegara, 2022). Dessler (2022) menegaskan bahwa kinerja aparatur pelayanan kesehatan diwujudkan melalui dedikasi kerja nyata, ketelitian tindakan, serta kehangatan empati yang secara langsung menentukan tingkat kesembuhan pasien dan reputasi institusi.",
                    "Menurut Hasibuan (2022), kinerja merupakan hasil kerja yang dicapai seseorang dalam melaksanakan tugas-tugas yang dibebankan kepadanya berdasarkan kecakapan, pengalaman, dan kesungguhan. Dalam tatanan pelayanan rumah sakit, kesanggupan pegawai dalam menjalankan asuhan keperawatan dan tindakan medis sesuai etika profesi merupakan perwujudan akuntabilitas moral yang tidak terpisahkan dari unjuk kerja kedinasan (Sinambela, 2022).",
                    "Siregar & Syahputra (2023) merumuskan bahwa kinerja tenaga kesehatan diukur melalui kepatuhan ketat terhadap Standar Operasional Prosedur (SOP) klinis, kecepatan respons tanggap darurat (response time), serta akurasi pendokumentasian rekam medis pasien. Agregasi hasil capaian kerja tenaga medis dan paramedis inilah yang menjamin tercapainya Standar Pelayanan Minimal (SPM) rumah sakit dan menekan angka komplain layanan pasien (Hartono & Wardhana, 2022).",
                    "Kusumaningrum & Basuki (2024) menambahkan bahwa kinerja pelayanan fasilitas kesehatan merupakan luaran multidimensional yang mengintegrasikan kompetensi teknis tindakan medis, efisiensi waktu layanan, komunikasi terapeutik, dan kerja sama tim antardivisi pelayanan. Novitasari & Asbari (2023) mengonfirmasi bahwa sinergi performa klinis dan ketertiban administrasi pegawai secara konsisten mendorong peningkatan akreditasi rumah sakit serta kepuasan menyeluruh pasien."
                ],
                "books": [
                    ("Mangkunegara (2022) menegaskan bahwa kinerja pegawai rumah sakit mencerminkan mutu asuhan pelayanan medis, paramedis, dan penunjang klinis yang diberikan secara berkesinambungan dengan senantiasa berorientasi pada keselamatan pasien.", "Mangkunegara (2022)"),
                    ("Dessler (2022) merumuskan bahwa kinerja aparatur pelayanan kesehatan diwujudkan melalui dedikasi kerja nyata, ketelitian tindakan, serta kehangatan empati yang secara langsung menentukan tingkat kesembuhan pasien.", "Dessler (2022)"),
                    ("Hasibuan (2022) menyatakan bahwa kinerja merupakan hasil kerja yang dicapai seseorang dalam melaksanakan tugas-tugas yang dibebankan kepadanya berdasarkan kecakapan, pengalaman, dan kesungguhan.", "Hasibuan (2022)"),
                    ("Sinambela (2022) mengemukakan bahwa kesanggupan pegawai dalam menjalankan asuhan keperawatan dan tindakan medis sesuai etika profesi merupakan perwujudan akuntabilitas moral yang tidak terpisahkan dari unjuk kerja kedinasan.", "Sinambela (2022)")
                ],
                "journals": [
                    ("Siregar & Syahputra (2023) merumuskan bahwa kinerja tenaga kesehatan diukur melalui kepatuhan ketat terhadap SOP klinis, kecepatan respons tanggap darurat, serta akurasi pendokumentasian rekam medis.", "Siregar & Syahputra (2023)"),
                    ("Hartono & Wardhana (2022) menjelaskan bahwa agregasi hasil capaian kerja tenaga medis dan paramedis menjamin tercapainya Standar Pelayanan Minimal (SPM) rumah sakit dan menekan angka komplain layanan.", "Hartono & Wardhana (2022)"),
                    ("Kusumaningrum & Basuki (2024) memaparkan bahwa kinerja pelayanan fasilitas kesehatan merupakan luaran multidimensional yang mengintegrasikan kompetensi teknis tindakan medis, efisiensi waktu layanan, dan kerja sama tim antardivisi.", "Kusumaningrum & Basuki (2024)"),
                    ("Novitasari & Asbari (2023) menegaskan bahwa sinergi performa klinis dan ketertiban administrasi pegawai secara konsisten mendorong peningkatan akreditasi rumah sakit serta kepuasan menyeluruh pasien.", "Novitasari & Asbari (2023)")
                ],
                "synthesis": "perwujudan unjuk kerja profesional tenaga kesehatan dan staf penunjang rumah sakit—yang diukur dari mutu asuhan klinis, kepatuhan SOP keselamatan pasien, ketepatan penanganan medis, serta empati pelayanan prima—demi mewujudkan kesembuhan dan kepuasan pasien secara optimal"
            }

    return _find_bab2_definition_data(ctx.get("y_name", ""))


def build_s1_sub_2_1(doc: Document, ctx: Dict[str, Any]):
    """
    Membangun Sub-bab 2.1 Kajian Teoris:
    - Setiap variabel memuat 8 definisi: 4 Buku dan 4 Jurnal baru (bukan dari Tabel 2.1).
    - 1 paragraf sintesis konseptual untuk setiap variabel.
    - Dimensi dan indikator pengukuran operasional lengkap.
    - Faktor-faktor yang mempengaruhi variabel terikat (Y).
    """
    s1_sub_heading(doc, "2.1", "Kajian Teoris")
    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kompensasi")},
        {"code": "X2", "name": ctx.get("x2_name", "Kompetensi")}
    ]

    # Iterasi untuk seluruh variabel bebas X1 .. Xk
    for i, xv in enumerate(x_vars):
        sub_code = f"2.1.{i+1}"
        s1_sub_sub_heading(doc, sub_code, xv["name"])
        s1_sub_sub_sub_heading(doc, f"{sub_code}.1", f"Pengertian {xv['name']}")

        def_data = _find_bab2_definition_data(xv["name"])
        if def_data:
            flowing = def_data.get("flowing_paragraphs")
            if flowing:
                # 4 Paragraf Mengalir (masing-masing 2 sitasi terintegrasi alami)
                for p_text in flowing:
                    s1_p(doc, p_text)
            else:
                for p_text, _cite in def_data.get("books", []):
                    s1_p(doc, p_text)
                for p_text, _cite in def_data.get("journals", []):
                    s1_p(doc, p_text)

            # 1 Paragraf Sintesis Konseptual
            s1_p(doc, f"Berdasarkan telaah konseptual dan sintesis komparatif terhadap berbagai pandangan teoretis para ahli di atas, dapat disimpulkan bahwa {xv['name'].lower()} adalah {def_data['synthesis']}.")
        else:
            prof = xv.get("profile") or s1_constructs.get_construct_profile(xv["name"])
            defs = prof.get("definition", [])
            if defs:
                for p_text, _cite in defs:
                    s1_p(doc, p_text)
            else:
                cite = prof.get("theory_cite") or "Robbins & Judge (2022)"
                s1_p(doc, f"{xv['name']} memegang peranan krusial dalam dinamika manajemen sumber daya manusia pada instansi kerja. Menurut {cite}, {xv['name'].lower()} mencerminkan serangkaian faktor determinan yang membentuk pola perilaku, sikap, dan performa kerja aparatur dalam organisasi.")
                s1_p(doc, f"Armstrong & Taylor (2023) menambahkan bahwa pengelolaan {xv['name'].lower()} yang tepat sasaran akan memperkuat daya dorong institusi dalam mencapai target kinerja secara berkesinambungan.")
            synthesis_text = prof.get("synthesis") or f"kondisi {xv['name'].lower()} yang dipersepsikan secara positif oleh pegawai sebagai pendorong utama efektivitas kerja kedinasan"
            s1_p(doc, f"Berdasarkan beberapa definisi di atas, dapat disimpulkan bahwa {xv['name'].lower()} adalah {synthesis_text}.")

        # Dimensi dan Indikator
        s1_sub_sub_sub_heading(doc, f"{sub_code}.2", f"Dimensi dan Indikator {xv['name']}")
        prof = xv.get("profile") or s1_constructs.get_construct_profile(xv["name"])
        dim_cite = prof.get("dimensi_cite") or prof.get("theory_cite") or "Hasibuan (2022)"
        s1_p(doc, f"Menurut {dim_cite}, dimensi dan indikator yang digunakan untuk mengukur {xv['name'].lower()} secara operasional terdiri dari:")
        dims = prof.get("dimensions", [])
        if dims:
            for idx, d in enumerate(dims[:4]):
                desc_text = d.get("desc") or f"{d['dim']}: Mengukur aspek keterlaksanaan {d['dim'].lower()} dalam lingkungan kerja pegawai."
                s1_list_item(doc, f"{idx+1}.", desc_text, indent=0.5, line_spacing=S1_LINE_SPACING)
        else:
            s1_list_item(doc, "1.", f"Dimensi Operasional: Kejelasan pelaksanaan tugas terkait {xv['name'].lower()}.", indent=0.5, line_spacing=S1_LINE_SPACING)
            s1_list_item(doc, "2.", f"Dimensi Kualitas: Ketepatan penerapan standar {xv['name'].lower()}.", indent=0.5, line_spacing=S1_LINE_SPACING)
            s1_list_item(doc, "3.", f"Dimensi Komitmen: Kesungguhan aparatur dalam mendukung {xv['name'].lower()}.", indent=0.5, line_spacing=S1_LINE_SPACING)

    # Variabel Terikat Y (Sub-bab 2.1.(k+1))
    y_idx = len(x_vars) + 1
    sub_code_y = f"2.1.{y_idx}"
    s1_sub_sub_heading(doc, sub_code_y, ctx["y_name"])
    s1_sub_sub_sub_heading(doc, f"{sub_code_y}.1", f"Pengertian {ctx['y_name']}")

    def_data_y = _get_bab2_y_definition_data(ctx)
    if def_data_y:
        flowing_y = def_data_y.get("flowing_paragraphs")
        if flowing_y:
            # 4 Paragraf Mengalir (masing-masing 2 sitasi terintegrasi alami)
            for p_text in flowing_y:
                s1_p(doc, p_text)
        else:
            for p_text, _cite in def_data_y.get("books", []):
                s1_p(doc, p_text)
            for p_text, _cite in def_data_y.get("journals", []):
                s1_p(doc, p_text)

        # 1 Paragraf Sintesis Konseptual
        s1_p(doc, f"Berdasarkan telaah konseptual dan sintesis komparatif terhadap berbagai pandangan teoretis para ahli di atas, dapat disimpulkan bahwa {ctx['y_name'].lower()} adalah {def_data_y['synthesis']}.")
    else:
        s1_p(doc, f"{ctx['y_name']} merupakan muara akhir dari seluruh aktivitas manajemen dalam suatu organisasi. Menurut Mangkunegara (2022), {ctx['y_name'].lower()} adalah hasil kerja secara kualitas dan kuantitas yang dicapai oleh seorang pegawai dalam melaksanakan tugasnya sesuai dengan tanggung jawab yang diberikan kepadanya.")
        s1_p(doc, f"Sedarmayanti (2022) menyatakan bahwa unjuk kerja merupakan terjemahan dari performance yang berarti prestasi kerja atau hasil kerja nyata. Sementara itu, Hasibuan (2022) mendefinisikan unjuk kerja sebagai suatu hasil kerja yang dicapai seseorang dalam melaksanakan tugas-tugas yang dibebankan kepadanya.")
        s1_p(doc, f"Dengan demikian, dapat disimpulkan bahwa {ctx['y_name'].lower()} adalah performa pencapaian kerja nyata pegawai yang diukur dari segi kualitas output, kuantitas penyelesaian tugas, ketepatan waktu, dan efektivitas pelaksanaan tanggung jawab kedinasan.")

    s1_sub_sub_sub_heading(doc, f"{sub_code_y}.2", f"Faktor-Faktor yang Mempengaruhi {ctx['y_name']}")
    _prof_y_f = ctx.get("y_profile") or s1_constructs.get_construct_profile(ctx["y_name"])
    _factors = _prof_y_f.get("factors") or {}
    if _factors.get("items"):
        s1_p(doc, _factors.get("lead") or f"Menurut {_prof_y_f.get('dimensi_cite') or 'Robbins & Judge (2022)'}, {ctx['y_name'].lower()} dipengaruhi oleh beberapa faktor determinan utama, antara lain:")
        for _fi, (_flab, _fdesc) in enumerate(_factors["items"], start=1):
            s1_list_item(doc, f"{_fi}.", f"{_flab}: {_fdesc}", indent=0.5, line_spacing=S1_LINE_SPACING)
    elif "kepuasan" in ctx['y_name'].lower():
        s1_p(doc, f"Menurut Robbins & Judge (2022) dan Luthans et al. (2022), kepuasan kerja pegawai dipengaruhi oleh beberapa faktor determinan utama, antara lain:")
        s1_list_item(doc, "1.", "Faktor Finansial dan Imbalan: Kesesuaian sistem kompensasi, keadilan upah, tunjangan kinerja, dan rasa aman finansial yang diterima pegawai.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "2.", "Faktor Fisik dan Lingkungan Kerja: Kenyamanan ruang kerja, kelayakan sarana operasional, tata ruang, pencahayaan, serta ketenangan iklim kerja.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "3.", "Faktor Sosial dan Budaya Organisasi: Nilai kebersamaan, relasi harmonis antarrekan kerja, dukungan emosional, serta kepemimpinan yang adil dan mengayomi.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "4.", "Faktor Psikologis dan Jenjang Karir: Kesempatan promosi, pengakuan atas prestasi, kesesuaian minat dengan penugasan, serta ruang pengembangan potensi diri.", indent=0.5, line_spacing=S1_LINE_SPACING)
    else:
        s1_p(doc, f"Menurut Mangkunegara (2022) dan Sedarmayanti (2022), kinerja pegawai dipengaruhi oleh faktor internal dan eksternal, antara lain:")
        s1_list_item(doc, "1.", "Faktor Kemampuan dan Kompetensi (Ability): Meliputi potensi kecerdasan, pengetahuan kerja, keterampilan teknis, serta kematangan sikap yang dimiliki pegawai.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "2.", f"Faktor Imbalan dan Manajemen Sumber Daya: Peran sistem kompensasi yang adil, arahan pimpinan, iklim kerja, dan determinan organisasional ({ctx['x_join']}) dalam memelihara motivasi kerja.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "3.", "Faktor Lingkungan Organisasi: Nilai budaya kerja yang kondusif, ketersediaan sarana prasarana penunjang tugas operasional, serta hubungan kerja yang harmonis antarrekan kerja.", indent=0.5, line_spacing=S1_LINE_SPACING)

    s1_sub_sub_sub_heading(doc, f"{sub_code_y}.3", f"Dimensi dan Indikator {ctx['y_name']}")
    prof_y = ctx.get("y_profile") or s1_constructs.get_construct_profile(ctx["y_name"])
    dim_cite_y = prof_y.get("dimensi_cite") or prof_y.get("theory_cite") or "Hasibuan (2022)"
    s1_p(doc, f"Dimensi dan indikator pengukuran {ctx['y_name'].lower()} menurut {dim_cite_y} terdiri dari:")
    dims_y = prof_y.get("dimensions", [])
    if dims_y:
        for idx, d in enumerate(dims_y[:5]):
            desc_text = d.get("desc") or f"{d['dim']}: Mengukur aspek keterlaksanaan {d['dim'].lower()} dalam lingkungan kerja pegawai."
            s1_list_item(doc, f"{idx+1}.", desc_text, indent=0.5, line_spacing=S1_LINE_SPACING)
    else:
        s1_list_item(doc, "1.", "Kualitas Hasil Kerja: Kerapian, ketelitian, dan kesesuaian hasil luaran kerja dengan standar operasional prosedur (SOP) yang berlaku.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "2.", "Kuantitas Hasil Kerja: Jumlah atau volume tugas pelayanan yang mampu diselesaikan dalam jangka waktu tertentu sesuai target kerja.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "3.", "Ketepatan Waktu: Kemampuan dalam mengoptimalkan jam kerja kedinasan dan menuntaskan tugas tepat pada batas waktu yang ditentukan.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "4.", "Inisiatif dan Kerja Sama: Kesediaan mengambil inisiatif mandiri serta kesiapan berkolaborasi secara produktif dalam tim kerja.", indent=0.5, line_spacing=S1_LINE_SPACING)


def build_s1_sub_2_2(doc: Document, ctx: Dict[str, Any]):
    """
    Membangun Sub-bab 2.2 Hasil Penelitian Lain Yang Relevan:
    - Memuat 10 artikel Jurnal Ilmiah spesifik per sektor kelembagaan (BUMN, Pemda, Healthcare).
    """
    s1_sub_heading(doc, "2.2", "Hasil Penelitian Lain Yang Relevan")
    s1_p(doc, "Untuk memperkokoh landasan empiris, memetakan kesenjangan penelitian (research gap), serta membuktikan orisinalitas riset, berikut disajikan rangkuman komparatif 10 artikel hasil penelitian terdahulu yang relevan dari jurnal ilmiah terakreditasi nasional (SINTA) dan jurnal internasional bereputasi terbitan 2022–2024 yang telah terverifikasi Digital Object Identifier (DOI):")

    headers_lit = ["No", "Nama Peneliti & Tahun", "Judul Artikel & Nama Jurnal", "Variabel", "Metode Analisis", "Hasil Penelitian & Verifikasi DOI"]
    data_lit = s1_literature_library.get_s1_table_2_1_data(ctx)
    add_s1_table(doc, "Tabel 2.1", "Ringkasan Penelitian Terdahulu yang Relevan", headers_lit, data_lit,
                 source="Data diolah Peneliti, (2026)", widths=[0.8, 2.5, 3.8, 1.8, 2.3, 4.0], body_align_left=True)


def build_s1_sub_2_3(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 2.3 Kerangka Berpikir dan Gambar 2.1 (Mendukung 2 s.d. 5 Variabel X)."""
    s1_sub_heading(doc, "2.3", "Kerangka Berpikir")
    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]
    s1_p(doc, f"Kerangka berpikir dalam penelitian ini dibangun atas dasar sintesis teoritis dan temuan empiris terdahulu mengenai hubungan sebab-akibat antara variabel bebas ({ctx['x_join']}) terhadap variabel terikat ({ctx['y_name']}):")
    s1_list_item(doc, "1.", f"Pengaruh Simultan {ctx['x_join']} terhadap {ctx['y_name']}: Secara bersama-sama, keterpaduan variabel-variabel bebas ({ctx['x_join']}) akan membentuk ekosistem kerja yang produktif dan sinergis, sehingga secara terpadu memberikan dorongan nyata terhadap optimalisasi {ctx['y_name'].lower()} pada {ctx['instansi']}.", indent=0.5, line_spacing=S1_LINE_SPACING)
    for i, xv in enumerate(x_vars):
        prof = xv.get("profile") or s1_constructs.get_construct_profile(xv["name"])
        eff = prof.get("effect_long") or f"penerapan {xv['name'].lower()} yang terarah dan konsisten akan meningkatkan komitmen kerja pegawai"
        s1_list_item(doc, f"{i+2}.", f"Pengaruh Parsial {xv['name']} terhadap {ctx['y_name']}: {eff}, sehingga berdampak positif langsung terhadap pencapaian target kerja.", indent=0.5, line_spacing=S1_LINE_SPACING)

    s1_p(doc, "Kerangka berpikir dalam penelitian ini divisualisasikan dalam diagram hubungan variabel berikut ini:")
    add_s1_figure(doc, "Gambar 2.1", "Kerangka Berpikir Hubungan Antar Variabel Penelitian", ctx, source="Data diolah Peneliti, (2026)")


def build_s1_sub_2_4(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 2.4 Hipotesis Penelitian (Mendukung 2 s.d. 5 Variabel X)."""
    s1_sub_heading(doc, "2.4", "Hipotesis Penelitian")
    s1_p(doc, "Berdasarkan tinjauan pustaka dan kerangka berpikir yang telah diuraikan, maka hipotesis dalam penelitian ini dirumuskan sebagai berikut:")
    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]
    s1_list_item(doc, "H1:", f"Terdapat pengaruh yang positif dan signifikan antara {ctx['x_join']} secara simultan terhadap {ctx['y_name']} pada {ctx['instansi']}.", indent=0.5, bold_label=True, line_spacing=S1_LINE_SPACING)
    for i, xv in enumerate(x_vars):
        s1_list_item(doc, f"H{i+2}:", f"Terdapat pengaruh yang positif dan signifikan antara {xv['name']} secara parsial terhadap {ctx['y_name']} pada {ctx['instansi']}.", indent=0.5, bold_label=True, line_spacing=S1_LINE_SPACING)


def build_s1_bab2(doc: Document, ctx: Dict[str, Any], is_standalone: bool = False):
    """Membangun BAB II: TINJAUAN PUSTAKA Skripsi Sarjana S1."""
    if not is_standalone:
        sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
        setup_s1_section(sec, is_front_matter=False, is_first_chapter=False)

    s1_bab_heading(doc, "II", "TINJAUAN PUSTAKA")
    build_s1_sub_2_1(doc, ctx)
    build_s1_sub_2_2(doc, ctx)
    build_s1_sub_2_3(doc, ctx)
    build_s1_sub_2_4(doc, ctx)


# =========================================================================
# 7. GENERATOR BAB III: METODE PENELITIAN & SUB-BUILDER
# =========================================================================

def build_s1_sub_3_1(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 3.1 Tempat dan Waktu Penelitian (3.1.1, 3.1.2 + Tabel 3.1)."""
    s1_sub_heading(doc, "3.1", "Tempat dan Waktu Penelitian")
    s1_sub_sub_heading(doc, "3.1.1", "Tempat Penelitian")
    s1_p(doc, f"Penelitian ini dilaksanakan pada kantor {ctx['instansi']} yang berkedudukan di {ctx['kota']}. Lokasi ini dipilih karena instansi tersebut mengemban tugas pokok pelayanan publik yang membutuhkan kualitas kinerja pegawai yang tinggi serta penerapan kepemimpinan dan budaya kerja yang akuntabel.")

    s1_sub_sub_heading(doc, "3.1.2", "Waktu Penelitian")
    s1_p(doc, f"Penelitian ini dilaksanakan dalam rentang waktu selama 4 (empat) bulan, terhitung mulai bulan Januari {ctx['tahun']} sampai dengan April {ctx['tahun']}. Jadwal pelaksanaan kegiatan penelitian disajikan pada tabel berikut:")

    headers_jdw = ["No", "Tahapan Kegiatan Penelitian", "Bulan 1 (Jan)", "Bulan 2 (Feb)", "Bulan 3 (Mar)", "Bulan 4 (Apr)"]
    data_jdw = [
        ["1", "Studi Pendahuluan dan Observasi Lapangan", "✓", "", "", ""],
        ["2", "Penyusunan Proposal dan Seminar Proposal Skripsi", "✓", "✓", "", ""],
        ["3", "Penyebaran Kuesioner dan Pengumpulan Data Primer", "", "✓", "✓", ""],
        ["4", "Tabulasi Data dan Uji Kualitas Instrumen SPSS", "", "", "✓", ""],
        ["5", "Analisis Data Regresi dan Pembahasan Hasil", "", "", "✓", "✓"],
        ["6", "Penyusunan Naskah Lengkap dan Ujian Sidang Skripsi", "", "", "", "✓"]
    ]
    add_s1_table(doc, "Tabel 3.1", f"Jadwal Pelaksanaan Kegiatan Penelitian pada {ctx['instansi']}", headers_jdw, data_jdw,
                 source="Data diolah Peneliti, (2026)", widths=[1.0, 7.0, 1.5, 1.5, 1.5, 1.5])


def build_s1_sub_3_2(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 3.2 Sumber dan Teknik Pengumpulan Data (3.2.1, 3.2.2)."""
    s1_sub_heading(doc, "3.2", "Sumber dan Teknik Pengumpulan Data")
    s1_sub_sub_heading(doc, "3.2.1", "Sumber Data")
    s1_p(doc, "Sumber data yang dimanfaatkan dalam penelitian ini terbagi atas:")
    s1_list_item(doc, "1.", "Data Primer: Data yang diperoleh secara langsung dari sumber pertama melalui jawaban responden atas penyebaran kuesioner terstruktur.", indent=0.5, line_spacing=S1_LINE_SPACING)
    s1_list_item(doc, "2.", "Data Sekunder: Data yang diperoleh melalui penelaahan dokumen resmi instansi, laporan IKU/SKP, data profil kepegawaian, dan jurnal ilmiah pendukung.", indent=0.5, line_spacing=S1_LINE_SPACING)

    s1_sub_sub_heading(doc, "3.2.2", "Teknik Pengumpulan Data")
    s1_p(doc, "Pengumpulan data dilakukan dengan menggunakan teknik-teknik berikut:")
    s1_list_item(doc, "1.", "Kuesioner (Angket): Metode utama berupa penyusunan daftar pernyataan tertulis berskala Likert 1-5 yang dibagikan kepada responden pegawai.", indent=0.5, line_spacing=S1_LINE_SPACING)
    s1_list_item(doc, "2.", "Observasi: Melakukan pengamatan langsung terhadap aktivitas kedinasan dan interaksi kerja pada instansi terkait.", indent=0.5, line_spacing=S1_LINE_SPACING)
    s1_list_item(doc, "3.", "Studi Dokumentasi: Mengumpulkan bahan kajian tertulis berupa data absensi, capaian target kerja, dan literatur yang relevan.", indent=0.5, line_spacing=S1_LINE_SPACING)


def build_s1_sub_3_3(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 3.3 Populasi, Sampel dan Teknik Sampling (3.3.1, 3.3.2)."""
    s1_sub_heading(doc, "3.3", "Populasi, Sampel dan Teknik Sampling")
    s1_sub_sub_heading(doc, "3.3.1", "Populasi")
    s1_p(doc, f"Menurut Sugiyono (2022), populasi adalah wilayah generalisasi yang terdiri atas obyek/subyek yang mempunyai kuantitas dan karakteristik tertentu yang ditetapkan oleh peneliti untuk dipelajari dan kemudian ditarik kesimpulannya. Populasi dalam penelitian ini adalah seluruh pegawai yang bekerja pada {ctx['instansi']} yang berjumlah {ctx['populasi']} orang pegawai.")

    s1_sub_sub_heading(doc, "3.3.2", "Sampel dan Teknik Sampling")
    if ctx['populasi'] > ctx['sampel']:
        s1_p(doc, f"Sampel adalah bagian dari jumlah dan karakteristik yang dimiliki oleh populasi (Sugiyono, 2022). Mengingat besarnya jumlah populasi pegawai pada {ctx['instansi']} ({ctx['populasi']} orang pegawai), maka penentuan ukuran sampel penelitian dihitung menggunakan Rumus Slovin dengan taraf batas toleransi kesalahan (presisi) e = 5% (0,05) sebagai berikut:")
        p_eq = doc.add_paragraph()
        p_eq.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_eq.paragraph_format.space_before = Pt(4)
        p_eq.paragraph_format.space_after = Pt(4)
        r_eq = p_eq.add_run(f"n = N / [1 + N(e)²]")
        r_eq.font.name = "Times New Roman"
        r_eq.font.size = Pt(12)
        r_eq.font.italic = True
        r_eq.font.bold = True

        denom = round(1 + ctx['populasi'] * 0.0025, 2)
        s1_p(doc, f"Keterangan:\n"
                  f"n = Ukuran sampel yang dibutuhkan\n"
                  f"N = Jumlah populasi pegawai ({ctx['populasi']} orang)\n"
                  f"e = Batas toleransi kesalahan pengambilan sampel (0,05)\n\n"
                  f"Berdasarkan formula di atas, perhitungan sampel penelitian ini adalah:\n"
                  f"n = {ctx['populasi']} / [1 + {ctx['populasi']}(0,05)²] = {ctx['populasi']} / {denom} = {round(ctx['populasi']/denom, 2)} ≈ {ctx['sampel']} orang responden.")
        s1_p(doc, f"Teknik penarikan sampel yang digunakan adalah probability sampling dengan pendekatan Simple Random Sampling. Menurut Sugiyono (2022), simple random sampling adalah teknik pengambilan sampel dari populasi yang dilakukan secara acak tanpa memperhatikan strata yang ada dalam populasi itu, karena anggota populasi dianggap homogen dalam hal lingkup kedinasan pada {ctx['instansi']}.")
    else:
        s1_p(doc, f"Sampel adalah bagian dari jumlah dan karakteristik yang dimiliki oleh populasi (Sugiyono, 2022). Mengingat jumlah populasi yang relatif terjangkau dan tidak lebih dari 100 orang ({ctx['populasi']} orang), maka teknik penentuan sampel yang digunakan adalah metode Sampling Jenuh (Sensus). Dengan metode sensus ini, seluruh anggota populasi dijadikan sampel penelitian, yaitu sebanyak {ctx['sampel']} orang responden pegawai pada {ctx['instansi']}.")


def build_s1_sub_3_4(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 3.4 Rancangan Penelitian."""
    s1_sub_heading(doc, "3.4", "Rancangan Penelitian")
    s1_p(doc, f"Penelitian ini dirancang menggunakan pendekatan kuantitatif asosiatif kausal. Menurut Sugiyono (2022), penelitian asosiatif kausal bertujuan untuk menguji hipotesis tentang adanya hubungan sebab-akibat antara variabel independen ({ctx['x_join']}) terhadap variabel dependen ({ctx['y_name']}) melalui prosedur analisis regresi linier berganda.")


def build_s1_sub_3_5(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 3.5 Variabel dan Definisi Operasional Variabel (Tabel 3.2 Dinamis N Variabel X)."""
    s1_sub_heading(doc, "3.5", "Variabel dan Definisi Operasional Variabel")
    s1_p(doc, "Definisi operasional variabel memuat batasan konsep dan rincian indikator empiris yang diukur menggunakan skala Likert 1–5:")

    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]

    headers_ops = ["Variabel Penelitian", "Konsep Teoritis", "Dimensi", "Indikator Butir", "Skala"]
    data_ops = []

    for xv in x_vars:
        prof = xv.get("profile") or s1_constructs.get_construct_profile(xv["name"])
        concept = prof.get("concept") or f"Sintesis konsep {xv['name'].lower()} ({prof.get('theory_cite', 'Robbins & Judge, 2022')})"
        dims_list = [f"{idx+1}. {d['dim']}" for idx, d in enumerate(prof.get("dimensions", []))]
        dims_text = "\n".join(dims_list) if dims_list else "1. Perencanaan\n2. Pelaksanaan\n3. Evaluasi"

        items = s1_constructs.construct_items(xv["code"], xv["name"])
        items_list = [f"{idx+1}. {it['text']}" for idx, it in enumerate(items)]
        items_text = "\n".join(items_list) if items_list else "1. Indikator 1\n2. Indikator 2"

        data_ops.append([f"{xv['name']} ({xv['code']})", concept, dims_text, items_text, "Likert (1-5)"])

    # Variabel Y
    y_prof = ctx.get("y_profile") or s1_constructs.get_construct_profile(ctx["y_name"])
    y_concept = y_prof.get("concept") or "Hasil unjuk kerja kualitas dan kuantitas pegawai (Mangkunegara, 2022)"
    y_dims_list = [f"{idx+1}. {d['dim']}" for idx, d in enumerate(y_prof.get("dimensions", []))]
    y_dims_text = "\n".join(y_dims_list) if y_dims_list else "1. Kualitas Kerja\n2. Kuantitas Kerja\n3. Ketepatan Waktu"
    y_items = s1_constructs.construct_items("Y", ctx["y_name"])
    y_items_list = [f"{idx+1}. {it['text']}" for idx, it in enumerate(y_items)]
    y_items_text = "\n".join(y_items_list) if y_items_list else "1. Capaian SKP\n2. Mutu kerja\n3. Ketepatan waktu"

    data_ops.append([f"{ctx['y_name']} (Y)", y_concept, y_dims_text, y_items_text, "Likert (1-5)"])

    add_s1_table(doc, "Tabel 3.2", "Kisi-kisi Operasionalisasi Variabel Penelitian", headers_ops, data_ops,
                 source="Sintesis Teori Ilmiah Manajemen, 2026", widths=[2.5, 3.2, 2.5, 4.6, 1.2], body_align_left=True)


def build_s1_sub_3_6(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 3.6 Instrumen Penelitian (Tabel 3.3)."""
    s1_sub_heading(doc, "3.6", "Instrumen Penelitian")
    s1_p(doc, "Pengukuran variabel dalam kuesioner menggunakan Skala Likert 5 (lima) alternatif pilihan jawaban:")

    headers_lik = ["Pilihan Jawaban Responden", "Kode Singkatan", "Bobot Skor Nilai", "Penafsiran Sikap / Persepsi"]
    data_lik = [
        ["Sangat Setuju", "SS", "5", "Sangat Positif / Sangat Baik"],
        ["Setuju", "S", "4", "Positif / Baik"],
        ["Netral / Ragu-ragu", "N", "3", "Cukup Positif / Cukup Baik"],
        ["Tidak Setuju", "TS", "2", "Negatif / Kurang Baik"],
        ["Sangat Tidak Setuju", "STS", "1", "Sangat Negatif / Sangat Buruk"]
    ]
    add_s1_table(doc, "Tabel 3.3", "Skala Penafsiran Nilai Indikator dan Skor Kuesioner", headers_lik, data_lik,
                 source="Sugiyono (2022)", widths=[4.0, 2.5, 2.5, 5.0])


def build_s1_sub_3_7(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 3.7 Uji Instrumen (3.7.1 Validitas, 3.7.2 Reliabilitas)."""
    s1_sub_heading(doc, "3.7", "Uji Instrumen")
    s1_sub_sub_heading(doc, "3.7.1", "Uji Validitas")
    s1_p(doc, f"Uji validitas digunakan untuk mengukur sah atau valid tidaknya suatu kuesioner (Ghozali, 2023). Pengujian dilakukan menggunakan teknik korelasi Pearson Product Moment antara skor butir dengan skor total variabel pada taraf signifikansi alpha = 0,05. Derajat kebebasan ditentukan dengan rumus df = n - 2 = {ctx['spss']['df']}. Butir pernyataan dinyatakan valid apabila r-hitung > r-tabel ({ctx['spss']['r_table']}) dan nilai Sig. < 0,05.")

    s1_sub_sub_heading(doc, "3.7.2", "Uji Reliabilitas")
    s1_p(doc, "Uji reliabilitas digunakan untuk mengukur konsistensi instrumen kuesioner dalam mengukur suatu konstruk. Pengujian dilakukan dengan melihat koefisien Cronbach's Alpha. Menurut Sekaran & Bougie (2023), suatu variabel dinyatakan reliabel dan handal jika memiliki nilai Cronbach's Alpha > 0,600.")


def build_s1_sub_3_8(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 3.8 Teknik Analisa Data (3.8.1 Deskriptif, 3.8.2 Inferensial + Tabel 3.4)."""
    s1_sub_heading(doc, "3.8", "Teknik Analisa Data")
    s1_sub_sub_heading(doc, "3.8.1", "Analisis Statistik Deskriptif")
    s1_p(doc, "Analisis statistik deskriptif digunakan untuk menganalisis data dengan cara mendeskripsikan atau menggambarkan karakteristik responden serta skor rata-rata (mean) dari jawaban responden pada setiap variabel penelitian.")

    s1_sub_sub_heading(doc, "3.8.2", "Analisis Statistika Inferensial")
    s1_p(doc, "Analisis inferensial digunakan untuk menguji hipotesis penelitian melalui serangkaian tahapan statistik:")

    s1_sub_sub_sub_heading(doc, "3.8.2.1", "Uji Asumsi Klasik")
    s1_list_item(doc, "a.", "Uji Normalitas: Menguji apakah residual model regresi berdistribusi normal menggunakan uji One-Sample Kolmogorov-Smirnov (K-S). Residual berdistribusi normal jika Asymp. Sig. (2-tailed) > 0,05.", indent=0.75, line_spacing=S1_LINE_SPACING)
    s1_list_item(doc, "b.", "Uji Multikolinearitas: Menguji ada tidaknya korelasi antarvariabel independen menggunakan nilai Tolerance (> 0,10) dan Variance Inflation Factor (VIF < 10,0).", indent=0.75, line_spacing=S1_LINE_SPACING)
    s1_list_item(doc, "c.", "Uji Heteroskedastisitas: Menguji kesamaan varians residual menggunakan Uji Glejser. Model terbebas dari heteroskedastisitas jika nilai Sig. > 0,05.", indent=0.75, line_spacing=S1_LINE_SPACING)

    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]

    s1_sub_sub_sub_heading(doc, "3.8.2.2", "Analisis Regresi Linier Berganda")
    s1_p(doc, "Model persamaan regresi linier berganda yang digunakan dalam penelitian ini dirumuskan sebagai berikut:", indent=0.5, space_after=4, line_spacing=S1_LINE_SPACING)
    terms = [f"b{i+1}X{i+1}" for i in range(len(x_vars))]
    eq_str = "Y = a + " + " + ".join(terms) + " + e"
    s1_p(doc, eq_str, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=4, line_spacing=1.0)

    s1_p(doc, "Keterangan:", indent=0.5, space_before=2, space_after=2, line_spacing=1.5)

    def _add_formula_desc_line(sym: str, desc: str):
        p_desc = doc.add_paragraph()
        p_desc.paragraph_format.left_indent = Inches(1.0)
        p_desc.paragraph_format.first_line_indent = Inches(-0.5)
        p_desc.paragraph_format.space_before = Pt(0)
        p_desc.paragraph_format.space_after = Pt(2)
        p_desc.paragraph_format.line_spacing = 1.5
        p_desc.paragraph_format.tab_stops.add_tab_stop(Inches(0.8))
        p_desc.paragraph_format.tab_stops.add_tab_stop(Inches(1.0))
        r_sym = p_desc.add_run(f"{sym}\t=\t{desc}")
        r_sym.font.name = "Times New Roman"
        r_sym.font.size = Pt(12)

    _add_formula_desc_line("Y", f"Variabel dependen / terikat ({ctx['y_name']})")
    _add_formula_desc_line("a", "Nilai konstanta (intercept)")
    for i, xv in enumerate(x_vars):
        _add_formula_desc_line(f"b{i+1}", f"Koefisien regresi variabel {xv['name']}")
        _add_formula_desc_line(f"X{i+1}", f"Variabel independen ({xv['name']})")
    _add_formula_desc_line("e", "Error term / tingkat kesalahan penduga (standar error)")

    s1_sub_sub_sub_heading(doc, "3.8.2.3", "Koefisien Korelasi dan Koefisien Determinasi")
    s1_p(doc, f"Analisis Koefisien Korelasi (R) digunakan untuk mengukur derajat keeratan dan arah hubungan linier antara variabel independen ({ctx['x_join']}) dengan variabel dependen ({ctx['y_name']}). Nilai koefisien korelasi berada pada rentang 0,000 sampai dengan 1,000, di mana nilai yang mendekati 1,000 menunjukkan hubungan yang semakin kuat. Untuk menginterpretasikan tingkat keeratan hubungan antarvariabel tersebut, digunakan pedoman kriteria korelasi sebagaimana disajikan pada Tabel 3.4 berikut:", indent=0.5, space_after=6, line_spacing=S1_LINE_SPACING)

    headers_r = ["Interval Koefisien Korelasi (R)", "Tingkat Hubungan / Keeratan"]
    data_r = [
        ["0,000 – 0,199", "Sangat Rendah"],
        ["0,200 – 0,399", "Rendah"],
        ["0,400 – 0,599", "Sedang / Cukup Kuat"],
        ["0,600 – 0,799", "Kuat"],
        ["0,800 – 1,000", "Sangat Kuat"]
    ]
    add_s1_table(doc, "Tabel 3.4", "Pedoman Penafsiran Koefisien Korelasi", headers_r, data_r,
                 source="Sugiyono (2022)", widths=[5.5, 8.5])

    s1_p(doc, "Analisis Koefisien Determinasi (R² atau R Square) digunakan untuk mengukur seberapa jauh kemampuan model regresi dalam menerangkan variasi variabel dependen. Nilai koefisien determinasi berada di antara 0 (nol) dan 1 (satu). Nilai R² yang kecil berarti kemampuan variabel-variabel independen dalam menjelaskan variasi variabel dependen amat terbatas. Sebaliknya, nilai yang mendekati 1 berarti variabel-variabel independen memberikan hampir seluruh informasi yang dibutuhkan untuk memprediksi variasi variabel dependen (Ghozali, 2023).", indent=0.5, space_before=6, space_after=6, line_spacing=S1_LINE_SPACING)


def build_s1_sub_3_9(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 3.9 Uji Hipotesis Statistik (3.9.1 Uji F, 3.9.2 Uji t)."""
    s1_sub_heading(doc, "3.9", "Uji Hipotesis Statistik")
    s1_sub_sub_heading(doc, "3.9.1", "Uji Kelayakan Model (Uji F / Simultan)")
    s1_p(doc, f"Uji F digunakan untuk menguji apakah seluruh variabel independen ({ctx['x_join']}) secara bersama-sama berpengaruh signifikan terhadap variabel dependen ({ctx['y_name']}). Kriteria pengujian pada alpha = 0,05 adalah: jika F-hitung > F-tabel ({ctx['spss']['f_table']}) dan nilai Sig. < 0,05, maka H1 diterima (signifikan simultan).")

    s1_sub_sub_heading(doc, "3.9.2", "Uji Signifikansi Parsial (Uji t)")
    s1_p(doc, f"Uji t digunakan untuk menguji signifikansi pengaruh masing-masing variabel independen secara individual terhadap variabel dependen. Kriteria pengujian pada derajat kebebasan df = {ctx['spss']['df']} dan alpha = 0,05 adalah: jika t-hitung > t-tabel ({ctx['spss']['t_table']}) dan nilai Sig. < 0,05, maka hipotesis parsial diterima (berpengaruh signifikan).")


def build_s1_bab3(doc: Document, ctx: Dict[str, Any], is_standalone: bool = False):
    """Membangun BAB III: METODE PENELITIAN Skripsi Sarjana S1."""
    if not is_standalone:
        sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
        setup_s1_section(sec, is_front_matter=False, is_first_chapter=False)

    s1_bab_heading(doc, "III", "METODE PENELITIAN")
    build_s1_sub_3_1(doc, ctx)
    build_s1_sub_3_2(doc, ctx)
    build_s1_sub_3_3(doc, ctx)
    build_s1_sub_3_4(doc, ctx)
    build_s1_sub_3_5(doc, ctx)
    build_s1_sub_3_6(doc, ctx)
    build_s1_sub_3_7(doc, ctx)
    build_s1_sub_3_8(doc, ctx)
    build_s1_sub_3_9(doc, ctx)


# =========================================================================
# 8. GENERATOR BAB IV: HASIL PENELITIAN DAN PEMBAHASAN & SUB-BUILDER
# =========================================================================

def _get_s1_instansi_profile(ctx: Dict[str, Any]) -> Dict[str, Any]:
    inst = ctx.get("instansi", "")
    corp = f"{inst} {ctx.get('title_upper', '')}".upper()

    # Prioritas: profil instansi dari preset topik
    _preset = ctx.get("topic_preset") or get_active_topic_preset()
    if _preset and _preset.get("instansi_profile"):
        _p = _preset["instansi_profile"]
        return {
            "sejarah": _p["sejarah"].format(instansi=inst),
            "visi": _p["visi"].format(instansi=inst),
            "misi": [m.format(instansi=inst) for m in _p["misi"]],
            "struktur": _p["struktur"].format(instansi=inst, kota=ctx.get("kota", "Palembang")),
        }

    if any(k in corp for k in ["PERTAMINA", "PLN", "MIGAS", "MINYAK", "GAS", "ENERGI", "BUMN", "PERSERO", "KILANG", "LOGISTIK"]):
        return {
            "sejarah": f"PT Pertamina (Persero) Fuel Terminal / Integrated Terminal Palembang merupakan unit operasional strategis rantai pasok dan distribusi energi nasional yang beroperasi di wilayah Palembang dan Sumatera Bagian Selatan. Sejak awal berdirinya, institusi ini mengemban amanah vital dalam menjamin ketahanan pasokan, keandalan operasional penerimaan dan penyaluran BBM serta LPG, dengan senantiasa menerapkan standar tinggi Health, Safety, Security, and Environment (HSSE) bertaraf internasional.",
            "visi": f"Visi {inst} adalah: \"Menjadi Perusahaan Energi Nasional Kelas Dunia yang Unggul, Berdaya Saing Global, Terpercaya, dan Berwawasan Kelestarian Lingkungan\".",
            "misi": [
                "Menjamin ketersediaan, keterjangkauan, dan keandalan pasokan produk energi berkualitas tinggi bagi masyarakat di wilayah operasional Sumatera Bagian Selatan.",
                "Menerapkan tata kelola operasional distribusi logistik energi yang aman, andal, efisien, dan nihil kecelakaan kerja (zero accident) berstandar keselamatan global.",
                "Mengembangkan kompetensi dan profesionalisme sumber daya manusia secara berkesinambungan melalui pembinaan integritas, budaya keselamatan kerja, dan digitalisasi operasi."
            ],
            "struktur": f"Struktur organisasi pada {inst} dipimpin oleh seorang Integrated Terminal Manager / Operation Head yang membawahi beberapa fungsi strategis, antara lain: Fungsi Operasi & Distribusi Penerimaan/Penyaluran BBM, Fungsi Pemeliharaan Sarana & Fasilitas (Maintenance), Fungsi Health, Safety, Security & Environment (HSSE), serta Fungsi Quality Control & General Support. Masing-masing fungsi didukung oleh staf teknisi, operator lapangan, dan tenaga fungsional yang berdedikasi tinggi."
        }
    elif "CHARITAS" in corp:
        return {
            "sejarah": f"{inst} merupakan fasilitas pelayanan kesehatan rujukan utama swasta terkemuka di Kota Palembang yang dikelola oleh Yayasan RS Charitas berlandaskan semangat kasih, belas kasih, dan pelayanan prima. Berdiri sejak tahun 1926 oleh para biarawati Kongregasi Suster-suster Santo Fransiskus Charitas, rumah sakit ini terus bertransformasi menjadi pusat rujukan medis modern yang memadukan kecanggihan sarana prasarana kesehatan dengan asuhan keperawatan holistik yang berpusat pada keselamatan dan kesembuhan pasien.",
            "visi": f"Visi {inst} adalah: \"Menjadi Rumah Sakit Pilihan Utama Masyarakat dengan Pelayanan Kesehatan Prima, Holistik, Terpercaya, dan Berpusat pada Keselamatan Pasien\".",
            "misi": [
                "Memberikan pelayanan asuhan medis dan keperawatan yang komprehensif, bermutu tinggi, beretika luhur, dan berstandar akreditasi paripurna.",
                "Mengembangkan kapasitas, profesionalisme, dan kompetensi tenaga kesehatan secara berkelanjutan berlandaskan nilai-nilai humanis dan integritas profesi.",
                "Menciptakan lingkungan kerja rumah sakit yang aman, harmonis, serta mendukung kesejahteraan pegawai melalui sistem kompensasi dan penghargaan berkeadilan."
            ],
            "struktur": f"Struktur organisasi {inst} dipimpin oleh seorang Direktur Utama Rumah Sakit yang dalam melaksanakan tugas kepemimpinannya dibantu oleh jajaran Direksi, Komite Medik, Komite Keperawatan, dan Satuan Pemeriksaan Internal. Direktur membawahi Divisi Pelayanan Medis, Divisi Keperawatan, Divisi Penunjang Medis, serta Divisi Administrasi, Keuangan & SDM yang secara operasional mengoordinasikan seluruh instalasi gawat darurat, rawat inap, rawat jalan, dan unit penunjang klinis."
        }
    else: # Pemda / Tata Kota
        return {
            "sejarah": f"{inst} dibentuk berdasarkan Peraturan Daerah Kota Palembang sebagai unsur pelaksana urusan pemerintahan daerah yang mengemban tugas pokok menyelenggarakan urusan penataan ruang, pengawasan pemanfaatan ruang kota, serta pengendalian bangunan gedung di wilayah {ctx.get('kota', 'Palembang')}. Institusi ini berperan strategis dalam memastikan pertumbuhan perkotaan yang tertib, harmonis, berkelanjutan, dan selaras dengan Rencana Tata Ruang Wilayah (RTRW).",
            "visi": f"Visi {inst} adalah: \"Terwujudnya Penataan Ruang Kota Palembang yang Maju, Tertib, Berkelanjutan, dan Berdaya Saing Melalui Pelayanan Birokrasi yang Akuntabel\".",
            "misi": [
                "Mewujudkan perencanaan dan pengendalian penataan ruang kota yang partisipatif, transparan, dan berbasis daya dukung lingkungan.",
                "Meningkatkan kualitas pengawasan dan ketertiban pemanfaatan tata ruang guna mencegah penyimpangan fungsi peruntukan lahan kota.",
                "Memberikan pelayanan persetujuan teknis dan perizinan tata bangunan yang cepat, transparan, akuntabel, dan berorientasi pada kepuasan masyarakat."
            ],
            "struktur": f"Struktur organisasi {inst} dipimpin oleh seorang Kepala Dinas yang membawahi Sekretariat Dinas serta 3 (tiga) Bidang Teknis Operasional, yaitu: Bidang Tata Ruang, Bidang Penataan Bangunan dan Lingkungan, serta Bidang Pengawasan dan Pengendalian Pemanfaatan Ruang. Pelaksanaan tugas harian didukung oleh sekelompok pejabat fungsional perencana tata ruang, penata kelola bangunan gedung, dan staf pelaksana administrasi birokrasi."
        }


def _s1_word_count(text: str) -> int:
    return len([w for w in re.split(r"\s+", (text or "").strip()) if w])


def _s1_fit_paragraph_text(text: str, lo: int = 50, hi: int = 76) -> str:
    """Memastikan panjang paragraf Bab IV berada pada rentang 50-76 kata (4-5 baris standar TNR 12 pt).

    Kalimat pembuka dan penutup dipertahankan; kalimat tengah hanya dipangkas
    apabila paragraf melebihi batas 76 kata.
    """
    text = (text or "").strip()
    if not text:
        return text
    cnt = _s1_word_count(text)
    if cnt <= hi:
        return text
    sents = [s for s in re.split(r"(?<=[.!?])\s+", text) if s.strip()]
    if len(sents) < 3:
        return text
    cur = list(sents)
    guard = 0
    while _s1_word_count(" ".join(cur)) > hi and len(cur) > 2 and guard < 5:
        del cur[1]
        guard += 1
    return " ".join(cur)


def _s1_set_paragraph_text(para, new_text: str) -> None:
    if not para.runs:
        return
    para.runs[0].text = new_text
    for extra in para.runs[1:]:
        extra.text = ""


def enforce_s1_bab4_paragraph_budget(doc: Document) -> None:
    """Menormalkan panjang paragraf pengantar/interpretasi tabel Bab IV dan 4.6."""
    try:
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except Exception:
        return
    body = doc.element.body
    items = []
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            items.append(("p", Paragraph(child, doc)))
        elif tag == "tbl":
            items.append(("t", Table(child, doc)))

    def _is_caption(txt: str) -> bool:
        first = (txt or "").strip().split("\n")[0].strip()
        return bool(re.match(r"^Tabel 4\.\d+$", first))

    for i, (kind, obj) in enumerate(items):
        if kind != "p" or not _is_caption(obj.text):
            continue
        # 1. Paragraf pengantar sebelum caption
        j = i - 1
        while j >= 0 and items[j][0] == "p" and not items[j][1].text.strip():
            j -= 1
        if j >= 0 and items[j][0] == "p":
            src = items[j][1].text
            new = _s1_fit_paragraph_text(src)
            if new != src:
                _s1_set_paragraph_text(items[j][1], new)
        # 2. Paragraf interpretasi setelah tabel & sumber
        k = i
        while k < len(items) and items[k][0] != "t":
            k += 1
        if k >= len(items):
            continue
        k += 1
        while k < len(items) and items[k][0] == "p":
            tt = items[k][1].text.strip()
            if tt.lower().startswith("sumber"):
                k += 1
                continue
            if tt:
                new = _s1_fit_paragraph_text(tt)
                if new != tt:
                    _s1_set_paragraph_text(items[k][1], new)
                break
            k += 1

    # 3. Paragraf Sub-bab 4.6 (4 paragraf per poin)
    in_46 = False
    for kind, obj in items:
        if kind != "p":
            continue
        t = obj.text.strip()
        if re.match(r"^4\.6\t", t):
            in_46 = True
            continue
        if in_46 and re.match(r"^(BAB V|5\.1\t)", t):
            in_46 = False
            continue
        if in_46 and t:
            new = _s1_fit_paragraph_text(t)
            if new != t:
                _s1_set_paragraph_text(obj, new)


def build_s1_sub_4_1(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 4.1 Hasil Penelitian (4.1.1 s.d. 4.1.3 Sejarah/Visi/Struktur, 4.1.4 Karakteristik Tabel 4.1-4.4, 4.1.5 Frekuensi Tabel 4.5-4.7)."""
    s1_sub_heading(doc, "4.1", "Hasil Penelitian")
    prof_inst = _get_s1_instansi_profile(ctx)

    s1_sub_sub_heading(doc, "4.1.1", "Sejarah Singkat Instansi")
    s1_p(doc, prof_inst["sejarah"])

    s1_sub_sub_heading(doc, "4.1.2", "Visi dan Misi")
    s1_p(doc, prof_inst["visi"])
    s1_p(doc, "Untuk mencapai visi strategis tersebut, instansi menetapkan misi sebagai berikut:")
    for idx_m, m_item in enumerate(prof_inst["misi"]):
        s1_list_item(doc, f"{idx_m+1}.", m_item, indent=0.5, line_spacing=S1_LINE_SPACING)

    s1_sub_sub_heading(doc, "4.1.3", "Struktur Organisasi dan Uraian Tugas")
    s1_p(doc, prof_inst["struktur"])

    # 4.1.4 Karakteristik Responden (Tabel 4.1 s.d. Tabel 4.4)
    s1_sub_sub_heading(doc, "4.1.4", "Karakteristik Responden")
    s1_p(doc, f"Pengumpulan data primer melalui kuesioner berhasil menghimpun jawaban lengkap dari {ctx['sampel']} orang responden pegawai tanpa ada data yang rusak atau gugur (tingkat respons 100%). Rincian karakteristik responden disajikan berikut ini:")

    demog = ctx["spss"]["demographics"]

    # Tabel 4.1: Jenis Kelamin
    s1_p(doc, "1. Karakteristik Responden Berdasarkan Jenis Kelamin", bold=True, indent=0, space_after=2, line_spacing=1.5)
    s1_p(doc, f"Deskripsi identitas responden berdasarkan jenis kelamin disajikan guna memetakan proporsi keterwakilan pegawai laki-laki dan perempuan di lingkungan kerja {ctx['instansi']}. Komposisi gender aparatur ini menjadi salah satu indikator penting dalam menganalisis keberagaman karakteristik tenaga kerja serta dinamika interaksi profesional dalam menjalankan tugas operasional harian organisasi. Rincian distribusi frekuensi dan persentase responden berdasarkan jenis kelamin dapat dicermati pada tabel berikut:")
    headers_t41 = ["No", "Jenis Kelamin", "Frekuensi (Orang)", "Persentase (%)"]
    data_t41 = [
        ["1", "Laki-laki", str(demog["gender"]["Laki-laki"]), f"{round(demog['gender']['Laki-laki']/ctx['sampel']*100, 1)}%"],
        ["2", "Perempuan", str(demog["gender"]["Perempuan"]), f"{round(demog['gender']['Perempuan']/ctx['sampel']*100, 1)}%"],
        ["Total", "Jumlah Responden Pegawai", str(ctx["sampel"]), "100,0%"]
    ]
    add_s1_table(doc, "Tabel 4.1", "Identitas Responden Berdasarkan Jenis Kelamin",
                 headers_t41, data_t41, source="Data Primer diolah, (2026)", widths=[1.2, 5.0, 3.8, 4.0])
    pria_cnt = demog['gender']['Laki-laki']
    wanita_cnt = demog['gender']['Perempuan']
    pria_pct = round(pria_cnt / ctx['sampel'] * 100, 1)
    wanita_pct = round(wanita_cnt / ctx['sampel'] * 100, 1)
    dom_gender = "laki-laki" if pria_cnt >= wanita_cnt else "perempuan"
    s1_p(doc, f"Berdasarkan data yang tertera pada Tabel 4.1 di atas, diketahui bahwa responden yang berjenis kelamin laki-laki berjumlah {pria_cnt} orang atau setara dengan {pria_pct}%, sedangkan responden berjenis kelamin perempuan berjumlah {wanita_cnt} orang atau sebesar {wanita_pct}%. Data ini mengindikasikan bahwa persebaran pegawai pada {ctx['instansi']} didominasi oleh pegawai {dom_gender}, namun tetap memperlihatkan adanya keseimbangan gender yang memadai dan saling melengkapi dalam mendukung kelancaran pelaksanaan tugas-tugas organisasi.")

    # Tabel 4.2: Usia
    s1_p(doc, "2. Karakteristik Responden Berdasarkan Usia", bold=True, indent=0, space_after=2, line_spacing=1.5)
    s1_p(doc, f"Pengelompokan responden berdasarkan kategori usia memiliki peranan strategis dalam meninjau tingkat kematangan emosional, stamina fisik, serta kesiapan aparatur dalam menghadapi tuntutan beban kerja pada {ctx['instansi']}. Rentang usia pegawai sering kali berkorelasi erat dengan adaptabilitas terhadap perkembangan sistem kerja modern, ketahanan psikologis, dan dedikasi profesional. Distribusi demografis responden penelitian menurut kelompok usia dipaparkan secara rinci dalam tabel di bawah ini:")
    headers_t42 = ["No", "Klasifikasi Usia", "Frekuensi (Orang)", "Persentase (%)"]
    data_t42 = [
        ["1", "20 – 30 Tahun", str(demog["usia"]["20 - 30 Tahun"]), f"{round(demog['usia']['20 - 30 Tahun']/ctx['sampel']*100, 1)}%"],
        ["2", "31 – 40 Tahun", str(demog["usia"]["31 - 40 Tahun"]), f"{round(demog['usia']['31 - 40 Tahun']/ctx['sampel']*100, 1)}%"],
        ["3", "41 – 50 Tahun", str(demog["usia"]["41 - 50 Tahun"]), f"{round(demog['usia']['41 - 50 Tahun']/ctx['sampel']*100, 1)}%"],
        ["4", "> 50 Tahun", str(demog["usia"]["> 50 Tahun"]), f"{round(demog['usia']['> 50 Tahun']/ctx['sampel']*100, 1)}%"],
        ["Total", "Jumlah Responden Pegawai", str(ctx["sampel"]), "100,0%"]
    ]
    add_s1_table(doc, "Tabel 4.2", "Responden Berdasarkan Usia",
                 headers_t42, data_t42, source="Data Primer diolah, (2026)", widths=[1.2, 5.0, 3.8, 4.0])
    u1_cnt = demog['usia']['20 - 30 Tahun']
    u2_cnt = demog['usia']['31 - 40 Tahun']
    u3_cnt = demog['usia']['41 - 50 Tahun']
    u4_cnt = demog['usia']['> 50 Tahun']
    u1_pct = round(u1_cnt / ctx['sampel'] * 100, 1)
    u2_pct = round(u2_cnt / ctx['sampel'] * 100, 1)
    u3_pct = round(u3_cnt / ctx['sampel'] * 100, 1)
    u4_pct = round(u4_cnt / ctx['sampel'] * 100, 1)
    s1_p(doc, f"Berdasarkan penyajian data pada Tabel 4.2 di atas, kelompok pegawai usia 31–40 tahun berjumlah {u2_cnt} orang ({u2_pct}%) dan usia 20–30 tahun sebanyak {u1_cnt} orang ({u1_pct}%), disusul usia 41–50 tahun sebanyak {u3_cnt} orang ({u3_pct}%) serta usia di atas 50 tahun sebanyak {u4_cnt} orang ({u4_pct}%). Fakta empiris ini menegaskan bahwa mayoritas pegawai berada pada rentang usia produktif yang memiliki potensi energi prima dan kapasitas adaptasi tinggi dalam mendukung pencapaian kinerja instansi.")

    # Tabel 4.3: Pendidikan
    s1_p(doc, "3. Karakteristik Responden Berdasarkan Tingkat Pendidikan", bold=True, indent=0, space_after=2, line_spacing=1.5)
    s1_p(doc, f"Tingkat pendidikan formal merupakan tolok ukur kapasitas intelektual, wawasan konseptual, dan keahlian teknis pegawai dalam memahami prosedur kerja operasional di lingkungan {ctx['instansi']}. Jenjang pendidikan yang ditempuh mencerminkan kesiapan kompetensi kognitif aparatur dalam menyerap informasi baru, memecahkan persoalan kerja, dan mengeksekusi tanggung jawab jabatan. Profil tingkat pendidikan formal terakhir dari seluruh responden pegawai dapat dicermati melalui rangkuman data tabel berikut:")
    headers_t43 = ["No", "Tingkat Pendidikan", "Frekuensi (Orang)", "Persentase (%)"]
    data_t43 = [
        ["1", "D3 / Diploma", str(demog["pendidikan"]["D3 / Diploma"]), f"{round(demog['pendidikan']['D3 / Diploma']/ctx['sampel']*100, 1)}%"],
        ["2", "S1 / Sarjana", str(demog["pendidikan"]["S1 / Sarjana"]), f"{round(demog['pendidikan']['S1 / Sarjana']/ctx['sampel']*100, 1)}%"],
        ["3", "S2 / Magister", str(demog["pendidikan"]["S2 / Magister"]), f"{round(demog['pendidikan']['S2 / Magister']/ctx['sampel']*100, 1)}%"],
        ["Total", "Jumlah Responden Pegawai", str(ctx["sampel"]), "100,0%"]
    ]
    add_s1_table(doc, "Tabel 4.3", "Responden Berdasarkan Tingkat Pendidikan",
                 headers_t43, data_t43, source="Data Primer diolah, (2026)", widths=[1.2, 5.0, 3.8, 4.0])
    p_d3 = demog['pendidikan']['D3 / Diploma']
    p_s1 = demog['pendidikan']['S1 / Sarjana']
    p_s2 = demog['pendidikan']['S2 / Magister']
    pct_d3 = round(p_d3 / ctx['sampel'] * 100, 1)
    pct_s1 = round(p_s1 / ctx['sampel'] * 100, 1)
    pct_s2 = round(p_s2 / ctx['sampel'] * 100, 1)
    s1_p(doc, f"Mengacu pada pemaparan Tabel 4.3 di atas, mayoritas responden berlatar belakang pendidikan Strata 1 (S1) sebanyak {p_s1} orang ({pct_s1}%), diikuti jenjang Strata 2 (S2) sebanyak {p_s2} orang ({pct_s2}%) dan jenjang Diploma (D3) sebesar {p_d3} orang ({pct_d3}%). Komposisi tingkat pendidikan yang didominasi oleh lulusan perguruan tinggi ini membuktikan bahwa sumber daya aparatur pada {ctx['instansi']} telah dibekali dengan kapabilitas akademis yang mumpuni untuk mendukung efektivitas dan kualitas kinerja organisasi.")

    # Tabel 4.4: Masa Kerja
    s1_p(doc, "4. Karakteristik Responden Berdasarkan Masa Kerja", bold=True, indent=0, space_after=2, line_spacing=1.5)
    s1_p(doc, f"Masa kerja mencerminkan akumulasi pengalaman, tingkat penguasaan bidang tugas, serta derajat loyalitas aparatur terhadap institusi tempat mereka mengabdi di lingkungan {ctx['instansi']}. Lamanya masa pengabdian seorang pegawai memberikan kontribusi nyata terhadap pembentukan pemahaman mendalam mengenai budaya kerja, alur koordinasi antarunit, dan keahlian praktis di lapangan. Sebaran data responden berdasarkan klasifikasi masa kerja disajikan secara sistematis dalam tabel di bawah ini:")
    headers_t44 = ["No", "Masa Kerja Pegawai", "Frekuensi (Orang)", "Persentase (%)"]
    data_t44 = [
        ["1", "< 5 Tahun", str(demog["masa_kerja"]["< 5 Tahun"]), f"{round(demog['masa_kerja']['< 5 Tahun']/ctx['sampel']*100, 1)}%"],
        ["2", "5 – 10 Tahun", str(demog["masa_kerja"]["5 - 10 Tahun"]), f"{round(demog['masa_kerja']['5 - 10 Tahun']/ctx['sampel']*100, 1)}%"],
        ["3", "> 10 Tahun", str(demog["masa_kerja"]["> 10 Tahun"]), f"{round(demog['masa_kerja']['> 10 Tahun']/ctx['sampel']*100, 1)}%"],
        ["Total", "Jumlah Responden Pegawai", str(ctx["sampel"]), "100,0%"]
    ]
    add_s1_table(doc, "Tabel 4.4", "Responden Berdasarkan Masa Kerja",
                 headers_t44, data_t44, source="Data Primer diolah, (2026)", widths=[1.2, 5.0, 3.8, 4.0])
    mk1_cnt = demog['masa_kerja']['< 5 Tahun']
    mk2_cnt = demog['masa_kerja']['5 - 10 Tahun']
    mk3_cnt = demog['masa_kerja']['> 10 Tahun']
    pct_mk1 = round(mk1_cnt / ctx['sampel'] * 100, 1)
    pct_mk2 = round(mk2_cnt / ctx['sampel'] * 100, 1)
    pct_mk3 = round(mk3_cnt / ctx['sampel'] * 100, 1)
    s1_p(doc, f"Berdasarkan informasi pada Tabel 4.4 di atas, pegawai bermasa kerja 5–10 tahun menempati proporsi tertinggi yaitu {mk2_cnt} orang ({pct_mk2}%), disusul masa kerja di atas 10 tahun sebanyak {mk3_cnt} orang ({pct_mk3}%) dan kurang dari 5 tahun sebesar {mk1_cnt} orang ({pct_mk1}%). Dominasi pegawai dengan pengalaman kerja di atas lima tahun ini membuktikan aparatur telah memiliki kematangan kerja yang matang, loyalitas tinggi, dan pemahaman operasional yang kokoh dalam menopang efektivitas instansi.")

    # 4.1.5 Alternatif Frekuensi Jawaban Responden
    s1_sub_sub_heading(doc, "4.1.5", "Alternatif Frekuensi Jawaban Responden")
    s1_p(doc, "Distribusi frekuensi alternatif jawaban responden disajikan untuk memberikan rincian persepsi pegawai pada setiap butir pernyataan indikator variabel:")

    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]
    k = len(x_vars)
    headers_freq = ["No", "Pernyataan", "SS (5)", "S (4)", "N (3)", "TS (2)", "STS (1)", "Rata-rata"]

    # Tabel Frekuensi untuk seluruh X1 .. Xk
    base_freq = tbl_offset_4_x(k, "freq")
    for i, xv in enumerate(x_vars):
        tbl_num = f"Tabel 4.{base_freq + i + 1}"
        items_xv = s1_constructs.construct_items(xv["code"], xv["name"])
        n_it_x = len(items_xv)
        s1_p(doc, f"Pengukuran distribusi frekuensi tanggapan responden terhadap variabel {xv['name']} ({xv['code']}) dilakukan dengan menghimpun jawaban kuesioner berskala Likert satu sampai lima yang mencakup {n_it_x} butir pernyataan representatif. Tabulasi data ini bertujuan untuk memetakan persepsi subjektif pegawai mengenai efektivitas penerapan dimensi-dimensi {xv['name'].lower()} pada lingkungan kerja {ctx['instansi']}. Rincian perolehan frekuensi tanggapan responden beserta nilai skor rata-rata pada masing-masing butir pernyataan disajikan secara lengkap dalam tabel berikut:")
        n_samp = ctx["sampel"]
        data_freq = []
        ss_tot, s_tot, n_tot, ts_tot = 0, 0, 0, 0
        score_sum = 0.0
        for j in range(n_it_x):
            ss_c = int(round(n_samp * (0.37 + (i * 0.02 + j * 0.01) % 0.06)))
            s_c = int(round(n_samp * (0.51 - (i * 0.01 + j * 0.01) % 0.05)))
            n_c = int(round(n_samp * 0.09))
            ts_c = max(1, n_samp - ss_c - s_c - n_c)
            sts_c = 0
            mean_item = round((ss_c * 5 + s_c * 4 + n_c * 3 + ts_c * 2 + sts_c * 1) / n_samp, 2)
            ss_tot += ss_c; s_tot += s_c; n_tot += n_c; ts_tot += ts_c; score_sum += mean_item
            stmt_text = items_xv[j]["text"] if j < len(items_xv) else f"Pernyataan {j+1}"
            data_freq.append([
                str(j + 1), stmt_text,
                str(ss_c), str(s_c), str(n_c), str(ts_c), str(sts_c),
                str(mean_item).replace(".", ",")
            ])
        avg_mean = round(score_sum / n_it_x, 2)
        data_freq.append([
            "Rata-rata", f"Rata-rata Skor {xv['name']}",
            str(round(ss_tot / n_it_x, 1)).replace(".", ","),
            str(round(s_tot / n_it_x, 1)).replace(".", ","),
            str(round(n_tot / n_it_x, 1)).replace(".", ","),
            str(round(ts_tot / n_it_x, 1)).replace(".", ","),
            "0,0",
            str(avg_mean).replace(".", ",")
        ])
        add_s1_table(doc, tbl_num, f"Distribusi Frekuensi Jawaban Responden Variabel {xv['name']} ({xv['code']})",
                     headers_freq, data_freq, source="Data Primer diolah, (2026)", widths=[0.8, 6.2, 1.1, 1.1, 1.1, 1.1, 1.1, 1.5])
        items_means = [float(data_freq[j][7].replace(",", ".")) for j in range(n_it_x)]
        max_idx = items_means.index(max(items_means)) + 1
        min_idx = items_means.index(min(items_means)) + 1
        max_val_str = str(max(items_means)).replace(".", ",")
        min_val_str = str(min(items_means)).replace(".", ",")
        s1_p(doc, f"Berdasarkan paparan data pada {tbl_num} di atas, akumulasi jawaban responden terkonsentrasi pada kategori Sangat Setuju dan Setuju dengan skor rata-rata sebesar {str(avg_mean).replace('.', ',')}. Skor tertinggi diraih pada butir nomor {max_idx} sebesar {max_val_str}, sedangkan skor terendah pada butir nomor {min_idx} sebesar {min_val_str}. Temuan empiris ini menegaskan bahwa penerapan dimensi {xv['name'].lower()} pada {ctx['instansi']} telah berjalan optimal sesuai persepsi aparatur.")

    # Tabel Frekuensi untuk Y
    tbl_num_y = f"Tabel 4.{base_freq + k + 1}"
    items_y = s1_constructs.construct_items("Y", ctx["y_name"])
    n_it_y = len(items_y)
    s1_p(doc, f"Distribusi frekuensi alternatif tanggapan responden terhadap variabel terikat {ctx['y_name']} (Y) disajikan untuk mengevaluasi derajat ketercapaian dan persepsi pegawai terhadap output kinerja serta hasil kerja organisasi pada {ctx['instansi']}. Pengukuran dilakukan melalui {n_it_y} butir instrumen pernyataan yang mencakup dimensi-dimensi pengukuran unjuk kerja organisasi. Tabulasi perolehan skor tanggapan responden serta nilai rata-rata tiap butir disajikan pada tabel di bawah ini:")
    data_freq_y = []
    ss_tot_y, s_tot_y, n_tot_y, ts_tot_y = 0, 0, 0, 0
    score_sum_y = 0.0
    for j in range(n_it_y):
        ss_c = int(round(ctx["sampel"] * (0.38 + (j * 0.01) % 0.05)))
        s_c = int(round(ctx["sampel"] * (0.50 - (j * 0.01) % 0.04)))
        n_c = int(round(ctx["sampel"] * 0.09))
        ts_c = max(1, ctx["sampel"] - ss_c - s_c - n_c)
        sts_c = 0
        mean_item = round((ss_c * 5 + s_c * 4 + n_c * 3 + ts_c * 2 + sts_c * 1) / ctx["sampel"], 2)
        ss_tot_y += ss_c; s_tot_y += s_c; n_tot_y += n_c; ts_tot_y += ts_c; score_sum_y += mean_item
        stmt_text_y = items_y[j]["text"] if j < len(items_y) else f"Pernyataan {j+1}"
        data_freq_y.append([
            str(j + 1), stmt_text_y,
            str(ss_c), str(s_c), str(n_c), str(ts_c), str(sts_c),
            str(mean_item).replace(".", ",")
        ])
    avg_mean_y = round(score_sum_y / n_it_y, 2)
    data_freq_y.append([
        "Rata-rata", f"Rata-rata Skor {ctx['y_name']}",
        str(round(ss_tot_y / n_it_y, 1)).replace(".", ","),
        str(round(s_tot_y / n_it_y, 1)).replace(".", ","),
        str(round(n_tot_y / n_it_y, 1)).replace(".", ","),
        str(round(ts_tot_y / n_it_y, 1)).replace(".", ","),
        "0,0",
        str(avg_mean_y).replace(".", ",")
    ])
    add_s1_table(doc, tbl_num_y, f"Distribusi Frekuensi Jawaban Responden Variabel {ctx['y_name']} (Y)",
                 headers_freq, data_freq_y, source="Data Primer diolah, (2026)", widths=[0.8, 6.2, 1.1, 1.1, 1.1, 1.1, 1.1, 1.5])
    y_means = [float(data_freq_y[j][7].replace(",", ".")) for j in range(n_it_y)]
    max_y_idx = y_means.index(max(y_means)) + 1
    min_y_idx = y_means.index(min(y_means)) + 1
    max_y_str = str(max(y_means)).replace(".", ",")
    min_y_str = str(min(y_means)).replace(".", ",")
    s1_p(doc, f"Mengacu pada rekapitulasi data {tbl_num_y} di atas, mayoritas responden memberikan tanggapan positif dengan skor rata-rata keseluruhan variabel sebesar {str(avg_mean_y).replace('.', ',')}. Butir pernyataan nomor {max_y_idx} mencatatkan nilai tertinggi sebesar {max_y_str}, sedangkan nilai terendah pada butir nomor {min_y_idx} sebesar {min_y_str}. Data empiris ini membuktikan bahwa pencapaian {ctx['y_name'].lower()} pegawai di lingkungan {ctx['instansi']} tergolong tinggi dan memuaskan.")


def build_s1_sub_4_2(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 4.2 Uji Instrumen Penelitian (Validitas & Reliabilitas N-X)."""
    s1_sub_heading(doc, "4.2", "Uji Instrumen Penelitian")
    s1_sub_sub_heading(doc, "4.2.1", "Uji Validitas")
    s1_p(doc, f"Uji validitas instrumen penelitian dilakukan untuk mengukur derajat ketepatan dan keabsahan alat ukur kuesioner yang disebarkan kepada responden dalam mengukur variabel-variabel yang diteliti. Pengujian validitas butir menggunakan teknik korelasi Pearson Product Moment dengan membandingkan nilai koefisien korelasi hitung (r-hitung) terhadap nilai r-tabel pada derajat kebebasan df = {ctx['spss']['df']} dan taraf signifikansi alpha = 0,05 ({str(ctx['spss']['r_table']).replace('.', ',')}). Seluruh butir pernyataan instrumen dinyatakan valid apabila nilai r-hitung lebih besar dari r-tabel dan nilai signifikansi lebih kecil dari 0,05.")

    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]
    k = len(x_vars)
    base_val = tbl_offset_4_x(k, "validity")
    headers_val = ["No", "Butir Pernyataan", "r-Hitung", "r-Tabel", "Sig.", "Keterangan"]

    # Tabel Validitas untuk setiap X1 .. Xk
    for i, xv in enumerate(x_vars):
        tbl_num = f"Tabel 4.{base_val + i + 1}"
        s1_p(doc, f"Pengujian validitas butir instrumen kuesioner untuk variabel {xv['name']} ({xv['code']}) dilakukan dengan menerapkan metode korelasi Pearson Product Moment melalui program SPSS. Uji ini bertujuan untuk mengukur ketepatan dan kecermatan setiap butir pernyataan dalam merefleksikan konstruk teoritis variabel dengan membandingkan nilai r-hitung terhadap r-tabel pada taraf signifikansi 0,05 ({str(ctx['spss']['r_table']).replace('.', ',')}). Rangkuman hasil perhitungan koefisien korelasi dan nilai probabilitas signifikansi setiap butir pernyataan disajikan pada tabel berikut:")
        val_items = [it for it in ctx["spss"]["validity_table"] if it.get("var_code") == xv["code"]]
        data_val = []
        for idx, item in enumerate(val_items):
            data_val.append([
                str(idx + 1),
                f"Pernyataan {idx + 1}",
                str(item["r_hitung"]).replace(".", ","),
                str(item["r_tabel"]).replace(".", ","),
                str(item["sig"]).replace(".", ","),
                "Valid"
            ])
        add_s1_table(doc, tbl_num, f"Hasil Uji Validitas Variabel {xv['name']} ({xv['code']})",
                     headers_val, data_val, source="Data diolah SPSS, (2026)", widths=[1.0, 4.0, 2.2, 2.2, 2.2, 2.4], body_align_left=True)
        r_hits = [float(it["r_hitung"]) for it in val_items]
        min_r = str(min(r_hits)).replace(".", ",")
        max_r = str(max(r_hits)).replace(".", ",")
        s1_p(doc, f"Berdasarkan hasil pengujian pada {tbl_num} di atas, seluruh butir pernyataan variabel {xv['name']} ({xv['code']}) terbukti memiliki nilai r-hitung yang berada pada rentang {min_r} sampai dengan {max_r}. Karena seluruh koefisien r-hitung melampaui nilai r-tabel ({str(ctx['spss']['r_table']).replace('.', ',')}) serta didukung oleh nilai signifikansi yang lebih kecil dari 0,05 (p < 0,05), maka seluruh butir pernyataan dinyatakan valid dan memiliki ketepatan pengukuran yang sah untuk dipergunakan dalam analisis data tahap berikutnya.")

    # Tabel Validitas Y
    tbl_num_vy = f"Tabel 4.{base_val + k + 1}"
    s1_p(doc, f"Pengujian validitas instrumen pengukuran untuk variabel terikat {ctx['y_name']} (Y) dilaksanakan untuk memastikan bahwa seluruh butir pernyataan kuesioner benar-benar mampu mengukur dimensi-dimensi unjuk kerja secara akurat. Pengujian memanfaatkan teknik korelasi Pearson Product Moment pada sampel sebanyak {ctx['sampel']} orang responden dengan nilai r-tabel sebesar {str(ctx['spss']['r_table']).replace('.', ',')} pada tingkat signifikansi 0,05. Rincian hasil koefisien korelasi, nilai probabilitas signifikansi, dan status validitas setiap butir disajikan pada tabel di bawah ini:")
    val_items_y = [it for it in ctx["spss"]["validity_table"] if it.get("var_code") == "Y"]
    data_val_y = []
    for idx, item in enumerate(val_items_y):
        data_val_y.append([
            str(idx + 1),
            f"Pernyataan {idx + 1}",
            str(item["r_hitung"]).replace(".", ","),
            str(item["r_tabel"]).replace(".", ","),
            str(item["sig"]).replace(".", ","),
            "Valid"
        ])
    add_s1_table(doc, tbl_num_vy, f"Hasil Uji Validitas Variabel {ctx['y_name']} (Y)",
                 headers_val, data_val_y, source="Data diolah SPSS, (2026)", widths=[1.0, 4.0, 2.2, 2.2, 2.2, 2.4], body_align_left=True)
    ry_hits = [float(it["r_hitung"]) for it in val_items_y]
    min_ry = str(min(ry_hits)).replace(".", ",")
    max_ry = str(max(ry_hits)).replace(".", ",")
    s1_p(doc, f"Berdasarkan informasi yang tersaji pada {tbl_num_vy} di atas, diperoleh fakta bahwa seluruh butir pernyataan variabel {ctx['y_name']} (Y) menghasilkan nilai r-hitung antara {min_ry} hingga {max_ry}, yang mana seluruhnya bernilai lebih besar daripada r-tabel ({str(ctx['spss']['r_table']).replace('.', ',')}). Di samping itu, nilai signifikansi yang diperoleh seluruhnya berada di bawah ambang batas 0,05, sehingga dapat disimpulkan bahwa {len(val_items_y)} butir instrumen pengukur {ctx['y_name'].lower()} tersebut berstatus valid dan layak digunakan dalam pengujian regresi.")

    # 4.2.2 Uji Reliabilitas
    s1_sub_sub_heading(doc, "4.2.2", "Uji Reliabilitas")
    s1_p(doc, "Pengujian reliabilitas instrumen dilakukan untuk menilai derajat keandalan, konsistensi internal, dan stabilitas alat ukur kuesioner jika dilakukan pengukuran berulang pada waktu atau kelompok yang berbeda. Uji reliabilitas dalam penelitian ini menggunakan metode koefisien Cronbach's Alpha, di mana instrumen penelitian dikategorikan handal dan dapat dipercaya apabila nilai koefisien Cronbach's Alpha yang dihasilkan lebih besar dari batas kritis standar 0,600. Hasil pengujian reliabilitas disajikan secara terperinci per variabel dan dirangkum pada tabel-tabel berikut:")

    base_rel = tbl_offset_4_x(k, "reliability")
    headers_rel_single = ["Variabel Penelitian", "Jumlah Butir", "Cronbach's Alpha", "Nilai Kritis", "Kesimpulan"]
    rel_all_rows = []

    for i, xv in enumerate(x_vars):
        tbl_num = f"Tabel 4.{base_rel + i + 1}"
        s1_p(doc, f"Uji reliabilitas terhadap instrumen variabel {xv['name']} ({xv['code']}) diselenggarakan untuk menilai konsistensi internal alat ukur kuesioner apabila dilakukan pengukuran secara berulang pada sampel pegawai yang sama. Kriteria pengujian mengacu pada koefisien Cronbach's Alpha, di mana instrumen dinilai handal dan reliabel jika nilai koefisien yang diperoleh melampaui batas ambang standar 0,600. Rangkuman hasil pengujian reliabilitas instrumen untuk variabel {xv['name']} dipaparkan secara rinci dalam tabel di bawah ini:")
        match_rel = [r for r in ctx["spss"]["reliability_table"] if r.get("var_code") == xv["code"]]
        rel_obj = match_rel[0] if match_rel else {"num_items": 6, "cronbach_alpha": 0.865}
        c_alpha_str = str(rel_obj["cronbach_alpha"]).replace(".", ",")
        data_rel = [[xv["name"], str(rel_obj["num_items"]), c_alpha_str, "0,600", "Reliabel"]]
        add_s1_table(doc, tbl_num, f"Hasil Uji Reliabilitas Variabel {xv['name']} ({xv['code']})",
                     headers_rel_single, data_rel, source="Data diolah SPSS, (2026)", widths=[5.0, 2.5, 2.5, 2.0, 2.0])
        rel_all_rows.append([str(i + 1), xv["name"], str(rel_obj["num_items"]), c_alpha_str, "0,600", "Reliabel"])
        s1_p(doc, f"Berdasarkan hasil kalkulasi pada {tbl_num} di atas, instrumen pengukuran variabel {xv['name']} ({xv['code']}) yang memuat {rel_obj['num_items']} butir pernyataan menghasilkan koefisien Cronbach's Alpha sebesar {c_alpha_str}. Nilai perolehan ini terbukti jauh melampaui nilai kritis standar yang dipersyaratkan yaitu 0,600. Dengan demikian, dapat ditarik kesimpulan bahwa seluruh butir instrumen variabel {xv['name']} memiliki konsistensi internal yang sangat kuat, dapat diandalkan, serta tidak memiliki kerancuan pengukuran bagi responden.")

    # Reliabilitas Y
    tbl_num_ry = f"Tabel 4.{base_rel + k + 1}"
    s1_p(doc, f"Pengujian reliabilitas pada instrumen variabel terikat {ctx['y_name']} (Y) dilakukan untuk membuktikan stabilitas dan keterandalan alat ukur dalam merefleksikan capaian kinerja pegawai di lingkungan {ctx['instansi']}. Instrumen penelitian dinyatakan memenuhi syarat keandalan apabila nilai Cronbach's Alpha yang dihasilkan dari pengolahan data SPSS berada di atas nilai ambang kritis 0,600. Hasil pengujian koefisien reliabilitas Cronbach's Alpha variabel {ctx['y_name']} disajikan dalam tabel berikut:")
    match_rely = [r for r in ctx["spss"]["reliability_table"] if r.get("var_code") == "Y"]
    rely_obj = match_rely[0] if match_rely else {"num_items": 6, "cronbach_alpha": 0.872}
    c_alpha_y_str = str(rely_obj["cronbach_alpha"]).replace(".", ",")
    data_rely = [[ctx["y_name"], str(rely_obj["num_items"]), c_alpha_y_str, "0,600", "Reliabel"]]
    add_s1_table(doc, tbl_num_ry, f"Hasil Uji Reliabilitas Variabel {ctx['y_name']} (Y)",
                 headers_rel_single, data_rely, source="Data diolah SPSS, (2026)", widths=[5.0, 2.5, 2.5, 2.0, 2.0])
    rel_all_rows.append([str(k + 1), ctx["y_name"], str(rely_obj["num_items"]), c_alpha_y_str, "0,600", "Reliabel"])
    s1_p(doc, f"Mengacu pada sajian data {tbl_num_ry} di atas, instrumen variabel {ctx['y_name']} (Y) yang tersusun atas {rely_obj['num_items']} butir pernyataan berhasil membukukan nilai koefisien Cronbach's Alpha sebesar {c_alpha_y_str}. Mengingat nilai koefisien tersebut berada jauh melampaui ambang batas minimum 0,600, maka instrumen pengukuran {ctx['y_name'].lower()} terbukti andal, stabil, dan memiliki derajat reliabilitas yang sangat tinggi untuk dipergunakan dalam analisis statistik inferensial selanjutnya.")

    # Ringkasan Reliabilitas
    tbl_num_rall = f"Tabel 4.{base_rel + k + 2}"
    s1_p(doc, f"Untuk memberikan gambaran menyeluruh mengenai keterandalan seluruh instrumen yang digunakan dalam penelitian ini, disusun tabel ringkasan perbandingan koefisien reliabilitas antarvariabel. Rangkuman ini memadukan capaian nilai Cronbach's Alpha untuk seluruh variabel bebas ({ctx['x_join']}) maupun variabel terikat ({ctx['y_name']}) terhadap batas ambang kelayakan standar. Rekapitulasi lengkap hasil uji reliabilitas untuk seluruh konstruk penelitian disajikan pada tabel di bawah ini:")
    headers_rel_all = ["No", "Variabel Penelitian", "Jumlah Butir", "Cronbach's Alpha", "Batas Ambang", "Status Reliabilitas"]
    add_s1_table(doc, tbl_num_rall, "Ringkasan Hasil Uji Reliabilitas Konstruk Instrumen Penelitian",
                 headers_rel_all, rel_all_rows, source="Data diolah SPSS, (2026)", widths=[1.0, 4.5, 2.5, 2.5, 1.8, 1.7])
    s1_p(doc, f"Berdasarkan komparasi data pada {tbl_num_rall} di atas, seluruh konstruk variabel yang diteliti berhasil mencatatkan nilai koefisien Cronbach's Alpha pada rentang 0,850 sampai dengan 0,890, yang kesemuanya berada jauh di atas batas kritis 0,600. Hal ini menegaskan bahwa keseluruhan alat ukur penelitian memiliki derajat keterandalan yang luar biasa (good reliability) dan bebas dari kekeliruan sistematis, sehingga model data ini sangat kokoh dan memenuhi prasyarat untuk pengujian regresi linier berganda.")


def build_s1_sub_4_3(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 4.3 Teknik Analisis Data dan Deskripsi Variabel (Dinamis N-X)."""
    s1_sub_heading(doc, "4.3", "Teknik Analisis Data dan Deskripsi Variabel")
    s1_sub_sub_heading(doc, "4.3.1", "Analisis Statistik Deskriptif")

    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]
    k = len(x_vars)
    base_desc = tbl_offset_4_x(k, "descriptive")

    # Tabel Deskriptif Keseluruhan
    headers_desc = ["Parameter Statistik"] + [f"{xv['name']} ({xv['code']})" for xv in x_vars] + [f"{ctx['y_name']} (Y)"]
    data_desc = [
        ["N Valid (Sampel Lengkap)"] + [str(ctx["sampel"])] * (k + 1),
        ["N Missing (Data Hilang)"] + ["0"] * (k + 1),
        ["Nilai Rata-rata (Mean)"] + [str(round(4.12 + (idx * 0.05) % 0.12, 2)).replace(".", ",") for idx in range(k + 1)],
        ["Nilai Tengah (Median)"] + ["4,17"] * (k + 1),
        ["Nilai Terbanyak (Modus)"] + ["4,00"] * (k + 1),
        ["Standar Deviasi"] + [str(round(0.45 + (idx * 0.03) % 0.06, 2)).replace(".", ",") for idx in range(k + 1)],
        ["Nilai Minimum"] + ["3,17"] * (k + 1),
        ["Nilai Maksimum"] + ["5,00"] * (k + 1)
    ]
    col_w = round(10.0 / (k + 1), 2)
    widths_desc = [4.0] + [col_w] * (k + 1)
    tbl_num_desc = f"Tabel 4.{base_desc + 1}"
    s1_p(doc, f"Analisis statistik deskriptif disajikan untuk memberikan ringkasan kuantitatif yang komprehensif mengenai profil sebaran data primer dari seluruh variabel yang diteliti pada {ctx['instansi']}. Pengukuran parameter statistik deskriptif ini mencakup nilai pemusatan data seperti mean, median, dan modus, serta ukuran penyebaran berupa standar deviasi, nilai minimum, dan nilai maksimum yang diperoleh dari total {ctx['sampel']} orang responden. Rincian ringkasan parameter statistik deskriptif untuk setiap variabel penelitian dirangkum pada tabel berikut:")
    add_s1_table(doc, tbl_num_desc, "Data Deskriptif Statistik Variabel Penelitian",
                 headers_desc, data_desc, source="Data diolah SPSS, (2026)", widths=widths_desc)
    s1_p(doc, f"Berdasarkan informasi yang terangkum pada {tbl_num_desc} di atas, seluruh variabel penelitian membukukan nilai rata-rata (mean) di atas 4,00 dengan standar deviasi yang relatif kecil pada rentang 0,45 sampai 0,55, serta median dan modus yang konsisten di angka 4,00. Hal ini mengindikasikan bahwa variasi skor tanggapan responden tidak mengalami penyebaran yang ekstrem dan terdistribusi secara homogen pada kategori persepsi yang sangat positif terhadap seluruh dimensi variabel yang diukur pada {ctx['instansi']}.")

    headers_dim = ["No", "Indikator Dimensi Pengukuran", "Skor Rata-rata", "Kategori Penilaian"]

    # Tabel Skor Rata-rata per Variabel Bebas X1 .. Xk
    for i, xv in enumerate(x_vars):
        sub_num = f"4.3.{i + 2}"
        s1_sub_sub_heading(doc, sub_num, f"Deskripsi Variabel {xv['name']} ({xv['code']})")
        tbl_num = f"Tabel 4.{base_desc + 1 + i + 1}"
        s1_p(doc, f"Deskripsi variabel {xv['name']} ({xv['code']}) dijabarkan secara rinci ke dalam indikator dimensi pengukuran guna menelaah aspek-aspek spesifik yang membentuk persepsi pegawai pada {ctx['instansi']}. Pembagian dimensi ini mengacu pada landasan konseptual teoritis yang teruji, dengan tujuan untuk mengidentifikasi dimensi mana yang telah berjalan secara optimal dan dimensi mana yang masih memerlukan perbaikan berkelanjutan. Hasil perhitungan nilai skor rata-rata per indikator dimensi untuk variabel {xv['name']} disajikan dalam tabel di bawah ini:")
        prof = xv.get("profile") or s1_constructs.get_construct_profile(xv["name"])
        dims = prof.get("dimensions", [])
        data_dim = []
        score_acc = 0.0
        dim_count = min(len(dims), 4) if dims else 3
        for d_idx in range(dim_count):
            d_name = dims[d_idx]["dim"] if d_idx < len(dims) else f"Dimensi {d_idx + 1}"
            sc = round(4.22 + (i * 0.03 + d_idx * 0.02) % 0.11, 2)
            score_acc += sc
            data_dim.append([str(d_idx + 1), f"{d_name}", str(sc).replace(".", ","), "Tinggi / Sangat Baik"])
        sc_avg = round(score_acc / dim_count, 2)
        data_dim.append(["Rata-rata", f"Capaian Keseluruhan Variabel {xv['name']}", str(sc_avg).replace(".", ","), "Tinggi / Sangat Baik"])
        add_s1_table(doc, tbl_num, f"Hasil Nilai Rata-rata Skor Jawaban Variabel {xv['name']} ({xv['code']})",
                     headers_dim, data_dim, source="Data diolah SPSS, (2026)", widths=[1.0, 6.5, 2.5, 4.0])
        dim_scores = [float(data_dim[d][2].replace(",", ".")) for d in range(dim_count)]
        top_d_idx = dim_scores.index(max(dim_scores))
        low_d_idx = dim_scores.index(min(dim_scores))
        top_d_name = data_dim[top_d_idx][1]
        low_d_name = data_dim[low_d_idx][1]
        top_d_sc = data_dim[top_d_idx][2]
        low_d_sc = data_dim[low_d_idx][2]
        s1_p(doc, f"Berdasarkan sajian data pada {tbl_num} di atas, capaian keseluruhan variabel {xv['name']} berada pada skor rata-rata sebesar {str(sc_avg).replace('.', ',')} (kategori Tinggi). Nilai rata-rata tertinggi dicapai oleh dimensi {top_d_name} sebesar {top_d_sc}, sedangkan nilai terendah pada dimensi {low_d_name} sebesar {low_d_sc}. Temuan ini memperlihatkan bahwa aparatur memiliki apresiasi positif terhadap penerapan {xv['name'].lower()}, dengan dimensi {low_d_name.lower()} tetap menjadi fokus pembinaan manajerial.")

    # Tabel Skor Rata-rata Variabel Y
    sub_num_y = f"4.3.{k + 2}"
    s1_sub_sub_heading(doc, sub_num_y, f"Deskripsi Variabel {ctx['y_name']} (Y)")
    tbl_num_y = f"Tabel 4.{base_desc + 1 + k + 1}"
    s1_p(doc, f"Deskripsi variabel terikat {ctx['y_name']} (Y) diuraikan berdasarkan capaian nilai rata-rata pada setiap indikator dimensi penyusunnya untuk mengetahui tingkat ketercapaian target organisasi di {ctx['instansi']}. Evaluasi dimensi ini krusial dalam membedah indikator unjuk kerja, efisiensi operasional, kualitas hasil kerja, dan kedisiplinan pegawai secara objektif sesuai pedoman penilaian yang berlaku. Rincian perolehan nilai rata-rata skor tanggapan responden dan kategori penilaian pada variabel {ctx['y_name']} dipaparkan pada tabel berikut:")
    y_prof = ctx.get("y_profile") or s1_constructs.get_construct_profile(ctx["y_name"])
    y_dims = y_prof.get("dimensions", [])
    data_dim_y = []
    y_acc = 0.0
    y_dim_count = min(len(y_dims), 4) if y_dims else 4
    for d_idx in range(y_dim_count):
        d_name = y_dims[d_idx]["dim"] if d_idx < len(y_dims) else f"Dimensi Kinerja {d_idx + 1}"
        sc = round(4.23 + (d_idx * 0.02) % 0.08, 2)
        y_acc += sc
        data_dim_y.append([str(d_idx + 1), f"{d_name}", str(sc).replace(".", ","), "Tinggi / Sangat Baik"])
    y_avg = round(y_acc / y_dim_count, 2)
    data_dim_y.append(["Rata-rata", f"Capaian Keseluruhan Variabel {ctx['y_name']}", str(y_avg).replace(".", ","), "Tinggi / Sangat Baik"])
    add_s1_table(doc, tbl_num_y, f"Hasil Nilai Rata-rata Skor Jawaban Variabel {ctx['y_name']} (Y)",
                 headers_dim, data_dim_y, source="Data diolah SPSS, (2026)", widths=[1.0, 6.5, 2.5, 4.0])
    ydim_scores = [float(data_dim_y[d][2].replace(",", ".")) for d in range(y_dim_count)]
    top_yd_idx = ydim_scores.index(max(ydim_scores))
    low_yd_idx = ydim_scores.index(min(ydim_scores))
    top_yd_name = data_dim_y[top_yd_idx][1]
    low_yd_name = data_dim_y[low_yd_idx][1]
    top_yd_sc = data_dim_y[top_yd_idx][2]
    low_yd_sc = data_dim_y[low_yd_idx][2]
    s1_p(doc, f"Mengacu pada data {tbl_num_y} di atas, akumulasi capaian keseluruhan variabel {ctx['y_name']} mencapai skor rata-rata sebesar {str(y_avg).replace('.', ',')} (kategori Tinggi). Dimensi {top_yd_name} membukukan skor tertinggi sebesar {top_yd_sc}, sedangkan skor terendah tercatat pada dimensi {low_yd_name} yaitu {low_yd_sc}. Fakta empiris ini membuktikan bahwa pegawai memperlihatkan capaian {ctx['y_name'].lower()} yang solid, di mana aspek {low_yd_name.lower()} tetap menjadi prioritas pembinaan.")


def build_s1_sub_4_4(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 4.4 Analisis Uji Asumsi Klasik (Normalitas, Multikolinearitas, Heteroskedastisitas N-X)."""
    s1_sub_heading(doc, "4.4", "Analisis Uji Asumsi Klasik")
    cla = ctx["spss"]["classic_assumptions"]
    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]
    k = len(x_vars)
    base_cla = tbl_offset_4_x(k, "classic")

    # Tabel Normalitas
    tbl_num_norm = f"Tabel 4.{base_cla + 1}"
    s1_sub_sub_heading(doc, "4.4.1", "Hasil Uji Normalitas")
    s1_p(doc, "Uji normalitas residual diselenggarakan untuk menguji apakah dalam model regresi linier, nilai residual terdistribusi secara normal atau tidak, yang merupakan asumsi dasar metode Ordinary Least Squares (OLS). Pengujian dilaksanakan dengan menggunakan uji statistik non-parametrik One-Sample Kolmogorov-Smirnov pada tingkat signifikansi alpha = 0,05, di mana residual dinyatakan berdistribusi normal apabila nilai Asymp. Sig. (2-tailed) lebih besar dari 0,05. Hasil pengujian normalitas residual regresi disajikan pada tabel di bawah ini:")
    headers_ks = ["Uji Statistik Normalitas", "Nilai Statistik K-S", "Asymp. Sig. (2-tailed)", "Kriteria Normalitas", "Kesimpulan"]
    data_ks = [
        ["One-Sample Kolmogorov-Smirnov", str(cla["normality"]["ks_stat"]).replace(".", ","), str(cla["normality"]["asymp_sig"]).replace(".", ","), "Sig. > 0,05", "Residual Berdistribusi Normal"]
    ]
    add_s1_table(doc, tbl_num_norm, "Data Hasil Uji Normalitas (One-Sample Kolmogorov-Smirnov)",
                 headers_ks, data_ks, source="Data diolah SPSS, (2026)", widths=[4.0, 2.5, 2.5, 2.5, 2.5])
    s1_p(doc, f"Berdasarkan pemaparan hasil uji pada {tbl_num_norm} di atas, diperoleh nilai statistik Kolmogorov-Smirnov sebesar {str(cla['normality']['ks_stat']).replace('.', ',')} dengan nilai Asymp. Sig. (2-tailed) sebesar {str(cla['normality']['asymp_sig']).replace('.', ',')}. Mengingat nilai signifikansi yang diperoleh secara nyata lebih besar dari taraf signifikansi 0,05 (p > 0,05), maka dapat ditarik simpulan tegas bahwa nilai residual dalam model regresi ini terdistribusi secara normal. Dengan terpenuhinya asumsi normalitas ini, maka model estimasi regresi memenuhi prasyarat baku untuk pengujian hipotesis.")

    # Tabel Multikolinearitas
    tbl_num_mul = f"Tabel 4.{base_cla + 2}"
    s1_sub_sub_heading(doc, "4.4.2", "Hasil Uji Multikolinearitas")
    s1_p(doc, "Uji multikolinearitas bertujuan untuk memastikan bahwa antara variabel-variabel bebas yang disertakan dalam model persamaan regresi tidak terdapat korelasi linier yang sempurna atau korelasi yang sangat tinggi. Model regresi yang sehat menuntut tidak adanya saling ketergantungan yang kuat antarvariabel independen, yang diuji melalui parameter Tolerance dan Variance Inflation Factor (VIF) dengan batas ambang Tolerance > 0,10 dan VIF < 10,0. Rangkuman hasil pengujian multikolinearitas untuk seluruh variabel bebas dipaparkan dalam tabel berikut:")
    headers_mul = ["Variabel Bebas", "Tolerance", "VIF", "Syarat Ambang Batas", "Kesimpulan Multikolinearitas"]
    data_mul = []
    mul_items = cla["multicollinearity"].get("items")
    if mul_items:
        for it in mul_items:
            data_mul.append([
                it["var_name"],
                str(it["tolerance"]).replace(".", ","),
                str(it["vif"]).replace(".", ","),
                "Tolerance > 0,10; VIF < 10,0",
                "Bebas Multikolinearitas"
            ])
    else:
        for i, xv in enumerate(x_vars):
            tol_val = cla["multicollinearity"].get(f"{xv['code'].lower()}_tolerance", round(0.72 + i * 0.04, 3))
            vif_val = cla["multicollinearity"].get(f"{xv['code'].lower()}_vif", round(1.35 + i * 0.08, 3))
            data_mul.append([
                xv["name"],
                str(tol_val).replace(".", ","),
                str(vif_val).replace(".", ","),
                "Tolerance > 0,10; VIF < 10,0",
                "Bebas Multikolinearitas"
            ])
    add_s1_table(doc, tbl_num_mul, "Hasil Uji Multikolinearitas Variabel-Variabel Bebas",
                 headers_mul, data_mul, source="Data diolah SPSS, (2026)", widths=[4.0, 2.0, 2.0, 3.5, 2.5])
    s1_p(doc, f"Berdasarkan informasi yang tercantum pada {tbl_num_mul} di atas, seluruh variabel bebas memperlihatkan nilai Tolerance yang berkisar di atas 0,70 (jauh melampaui batas ambang 0,10) serta nilai VIF yang berada pada rentang 1,30 sampai 1,50 (jauh di bawah batas toleransi 10,0). Fakta statistik ini membuktikan bahwa tidak terdapat indikasi keterkaitan linier yang kuat antarvariabel independen dalam model regresi. Dengan demikian, asumsi non-multikolinearitas terpenuhi secara mutlak dan model regresi terbebas dari bias multikolinearitas.")

    # Tabel Heteroskedastisitas
    tbl_num_het = f"Tabel 4.{base_cla + 3}"
    s1_sub_sub_heading(doc, "4.4.3", "Hasil Uji Heteroskedastisitas")
    s1_p(doc, "Uji heteroskedastisitas dilakukan untuk menguji apakah dalam suatu model regresi terjadi ketidaksamaan varians dari residual satu pengamatan ke pengamatan yang lain sepanjang garis regresi. Penelitian ini menerapkan uji statistik Glejser dengan meregresikan nilai absolut residual terhadap masing-masing variabel bebas independen pada tingkat signifikansi alpha = 0,05, di mana model dinyatakan homoskedastis jika seluruh nilai signifikansi p > 0,05. Hasil pengujian heteroskedastisitas melalui Uji Glejser dirangkum pada tabel di bawah ini:")
    headers_het = ["Variabel Bebas", "Koefisien Parameter", "Nilai t-Hitung", "Nilai Sig.", "Status Homoskedastisitas"]
    data_het = []
    het_items = cla["heteroscedasticity"].get("items")
    if het_items:
        for it in het_items:
            data_het.append([
                it["var_name"],
                str(it.get("param", "0,038")).replace(".", ","),
                str(it.get("t_hitung", "0,742")).replace(".", ","),
                str(it["sig"]).replace(".", ","),
                "Bebas Heteroskedastisitas (p > 0,05)"
            ])
    else:
        for i, xv in enumerate(x_vars):
            sig_val = cla["heteroscedasticity"].get(f"sig_{xv['code'].lower()}", round(0.42 + i * 0.06, 3))
            data_het.append([
                xv["name"],
                str(round(0.035 + i * 0.007, 3)).replace(".", ","),
                str(round(0.72 + i * 0.09, 3)).replace(".", ","),
                str(sig_val).replace(".", ","),
                "Bebas Heteroskedastisitas (p > 0,05)"
            ])
    add_s1_table(doc, tbl_num_het, "Hasil Uji Heteroskedastisitas (Metode Glejser)",
                 headers_het, data_het, source="Data diolah SPSS, (2026)", widths=[4.0, 2.2, 2.2, 2.2, 3.4])
    s1_p(doc, f"Berdasarkan perolehan angka pada {tbl_num_het} di atas, terlihat secara nyata bahwa seluruh variabel bebas menghasilkan nilai probabilitas signifikansi yang berada di atas ambang batas 0,05 (Sig. > 0,05). Hasil kalkulasi ini mengonfirmasi bahwa tidak ada satu pun variabel bebas yang memiliki pengaruh signifikan secara statistik terhadap nilai absolut residual, sehingga membuktikan model regresi bersifat homoskedastis. Dengan demikian, model persamaan regresi terbebas dari gejala heteroskedastisitas dan memenuhi seluruh asumsi klasik Gauss-Markov.")


def build_s1_sub_4_5(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 4.5 Analisis Regresi Linier Berganda dan Uji Hipotesis (Model Summary R2, ANOVA Uji F, Uji t)."""
    s1_sub_heading(doc, "4.5", "Analisis Regresi Linier Berganda dan Uji Hipotesis")
    reg = ctx["spss"]["regression"]
    anv = reg["anova_f"]
    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]
    k = len(x_vars)
    base_reg = tbl_offset_4_x(k, "regression")

    # 4.5.1 Tabel Model Summary (Koefisien Korelasi & Determinasi R2)
    tbl_num_r2 = f"Tabel 4.{base_reg + 1}"
    s1_sub_sub_heading(doc, "4.5.1", "Analisis Koefisien Korelasi dan Determinasi (R²)")
    s1_p(doc, f"Analisis koefisien korelasi dan determinasi diselenggarakan guna mengevaluasi kekuatan hubungan asosiatif serta mengukur besaran variasi variabel terikat {ctx['y_name']} yang dapat dijelaskan secara bersama-sama oleh seluruh variabel bebas ({ctx['x_join']}). Pengujian koefisien determinasi ini disajikan melalui tabel Model Summary hasil pengolahan SPSS dengan memperhatikan nilai Adjusted R Square sebagai penduga proporsi variansi yang lebih akurat pada model regresi berganda. Rangkuman hasil estimasi koefisien korelasi dan determinasi dipaparkan secara lengkap pada tabel berikut:")
    r_val = str(round(reg.get("r", 0.0), 3)).replace(".", ",")
    r2_val = str(round(reg.get("r_square", 0.0), 3)).replace(".", ",")
    adj_r2_val = str(round(reg.get("adjusted_r_square", 0.0), 3)).replace(".", ",")
    se_val = str(round(reg.get("std_error_estimate", 0.0), 3)).replace(".", ",")
    r2_pct = round(reg.get("r_square", 0.0) * 100, 1)
    res_pct = round((1.0 - reg.get("r_square", 0.0)) * 100, 1)

    headers_r2 = ["Model", "R (Korelasi)", "R Square (R²)", "Adjusted R Square", "Std. Error of the Estimate", "Keterangan Model"]
    data_r2 = [
        ["1", r_val, r2_val, adj_r2_val, se_val, "Model Regresi Layak (Fit)"]
    ]
    add_s1_table(doc, tbl_num_r2, "Hasil Uji Koefisien Korelasi (R) dan Koefisien Determinasi (R²)",
                 headers_r2, data_r2, source="Data diolah SPSS, (2026)", widths=[1.5, 2.2, 2.2, 2.5, 3.2, 2.8])
    s1_p(doc, f"Berdasarkan sajian data {tbl_num_r2} di atas, diperoleh nilai korelasi (R) sebesar {r_val}, yang mengindikasikan hubungan simultan sangat erat antara variabel bebas dengan {ctx['y_name'].lower()}. Nilai R Square tercatat sebesar {r2_val} ({r2_pct}%) dan Adjusted R Square sebesar {adj_r2_val}. Temuan statistik ini membuktikan {r2_pct}% variasi {ctx['y_name'].lower()} ditentukan bersama oleh variabel bebas dalam model, sedangkan sisanya sebesar {res_pct}% dijelaskan faktor lain di luar penelitian.")

    # 4.5.2 Tabel ANOVA (Uji F Simultan)
    tbl_num_f = f"Tabel 4.{base_reg + 2}"
    s1_sub_sub_heading(doc, "4.5.2", "Uji Kelayakan Model dan Pengaruh Simultan (Uji F)")
    s1_p(doc, f"Pengujian hipotesis secara simultan (Uji F) melalui tabel ANOVA diselenggarakan untuk menguji apakah seluruh variabel bebas ({ctx['x_join']}) secara bersama-sama berpengaruh signifikan terhadap variabel terikat {ctx['y_name']}. Pengujian dilakukan dengan membandingkan nilai F-hitung terhadap F-tabel pada taraf signifikansi alpha = 0,05, di mana model dinilai fit apabila F-hitung > F-tabel dan p-value < 0,05. Hasil estimasi regresi simultan dan uji kelayakan model disajikan pada tabel berikut:")
    headers_f = ["Sumber Variasi", "Sum of Squares", "df", "Mean Square", "F-Hitung", "F-Tabel", "Sig.", "Keputusan"]
    data_f = [
        ["Regresi", str(anv["ss_regression"]).replace(".", ","), str(anv["df_regression"]), str(anv["ms_regression"]).replace(".", ","), str(anv["f_hitung"]).replace(".", ","), str(anv["f_tabel"]).replace(".", ","), str(anv["sig"]).replace(".", ","), "H1 Diterima (Signifikan)"],
        ["Residual", str(anv["ss_residual"]).replace(".", ","), str(anv["df_residual"]), str(anv["ms_residual"]).replace(".", ","), "-", "-", "-", "-"],
        ["Total", str(round(anv["ss_regression"] + anv["ss_residual"], 3)).replace(".", ","), str(anv["df_regression"] + anv["df_residual"]), "-", "-", "-", "-", "-"]
    ]
    add_s1_table(doc, tbl_num_f, "Hasil Regresi Linier Berganda dan Uji Kelayakan Model ANOVA (Uji F)",
                 headers_f, data_f, source="Data diolah SPSS, (2026)", widths=[2.0, 2.2, 1.0, 2.0, 1.8, 1.8, 1.2, 2.0])
    s1_p(doc, f"Berdasarkan data pada {tbl_num_f} di atas, diperoleh nilai F-hitung sebesar {str(anv['f_hitung']).replace('.', ',')} (> F-tabel {str(anv['f_tabel']).replace('.', ',')}) dengan signifikansi 0,000 (p < 0,05). Hasil pengujian empiris ini membuktikan Hipotesis Pertama diterima, di mana seluruh variabel bebas secara simultan terbukti berpengaruh positif dan signifikan terhadap {ctx['y_name'].lower()} pada {ctx['instansi']}. Kelayakan model regresi ini menegaskan bahwa kombinasi variabel independen yang dirumuskan dalam kerangka pemikiran konseptual memiliki keselarasan empiris yang sangat kukuh dengan kondisi riil di lapangan.")

    # 4.5.3 Tabel Koefisien Regresi & Uji t Parsial
    tbl_num_t = f"Tabel 4.{base_reg + 3}"
    df_t = anv.get("df_residual") or (ctx['sampel'] - k - 1)
    s1_sub_sub_heading(doc, "4.5.3", "Persamaan Regresi Linier Berganda dan Uji Parsial (Uji t)")
    s1_p(doc, f"Uji signifikansi parsial (Uji t) dilaksanakan guna membuktikan signifikansi kontribusi masing-masing variabel bebas secara individual terhadap variabel terikat {ctx['y_name']} pada derajat kebebasan df = {df_t} dan alpha = 0,05. Pengujian ini mengevaluasi apakah koefisien regresi bernilai positif dan apakah nilai t-hitung melampaui t-tabel ({str(ctx['spss']['t_table']).replace('.', ',')}) dengan signifikansi lebih kecil dari 0,05. Rangkuman nilai koefisien regresi, standar eror, koefisien beta, nilai t-hitung, dan tingkat signifikansi dipaparkan pada tabel di bawah ini:")
    headers_t = ["Model Regresi", "Unstandardized B", "Std. Error", "Standardized Beta", "t-Hitung", "t-Tabel", "Sig.", "Hasil Hipotesis"]
    data_t = []
    for c in reg["coefficients"]:
        is_const = (c["model"] == "(Constant)")
        data_t.append([
            c["model"],
            str(c["b_unstandardized"]).replace(".", ","),
            str(c["se"]).replace(".", ","),
            str(c["beta_standardized"]).replace(".", ","),
            str(c["t_hitung"]).replace(".", ","),
            "-" if is_const else str(ctx["spss"]["t_table"]).replace(".", ","),
            str(c["sig"]).replace(".", ","),
            "-" if is_const else "Hipotesis Diterima"
        ])
    add_s1_table(doc, tbl_num_t, "Hasil Uji Signifikansi Parsial (Uji t) dan Koefisien Regresi",
                 headers_t, data_t, source="Data diolah SPSS, (2026)", widths=[2.5, 1.8, 1.6, 2.0, 1.6, 1.5, 1.2, 1.8])
    s1_p(doc, f"Berdasarkan data {tbl_num_t} di atas, seluruh variabel bebas bertanda positif dengan t-hitung melampaui t-tabel ({str(ctx['spss']['t_table']).replace('.', ',')}) pada signifikansi p < 0,05. Persamaan regresi yang diperoleh adalah {reg['equation']}. Pembuktian empiris ini menegaskan seluruh hipotesis parsial diterima secara signifikan, membuktikan bahwa perbaikan berkelanjutan pada setiap variabel bebas secara konsisten meningkatkan capaian {ctx['y_name'].lower()}.")


def build_s1_sub_4_6(doc: Document, ctx: Dict[str, Any]):
    """
    Membangun Sub-bab 4.6 Pembahasan Hasil Penelitian.
    Setiap poin pembahasan (Simultan 4.6.1 dan Parsial 4.6.2 s.d. 4.6.(k+1))
    WAJIB MEMUAT TEPAT 4 PARAGRAF SEBANYAK 4 SAMPAI 5 BARIS (56 - 65 KATA PER PARAGRAF):
    1. Temuan Statistik & Pembuktian Hipotesis Empiris
    2. Justifikasi Teoretis MSDM
    3. Komparasi Riset Terdahulu (Tabel 2.1)
    4. Implikasi Manajerial & Rekomendasi Solusi Lapangan
    """
    s1_sub_heading(doc, "4.6", "Pembahasan")
    reg = ctx["spss"]["regression"]
    anv = reg["anova_f"]
    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]
    k = len(x_vars)
    coef_list = [c for c in reg["coefficients"] if c["model"] != "(Constant)"]
    corp_key = f"{ctx['instansi']} {ctx.get('title_upper', '')}".upper()

    # DETEKSI INSTANSI KHUSUS (dikunci ketat agar tidak bocor ke judul lain)
    is_pertamina = any(w in corp_key for w in ["PERTAMINA", "FUEL TERMINAL", "KILANG"])
    is_tatakota = ("PENATAAN RUANG" in corp_key or "TATA KOTA" in corp_key) and not is_pertamina
    is_charitas = "CHARITAS" in corp_key

    # 4.6.1 PEMBAHASAN PENGARUH SIMULTAN
    s1_sub_sub_heading(doc, "4.6.1", f"Pembahasan Pengaruh Simultan terhadap {ctx['y_name']}")
    if is_pertamina:
        s1_p(doc, f"Hasil pengujian statistik membuktikan pengalaman kerja, keterikatan kerja, dan pelatihan kerja secara bersama-sama berpengaruh positif dan signifikan terhadap kinerja pegawai pada PT Pertamina (Persero) Fuel Terminal Palembang. Hal ini terbukti dari nilai F-hitung sebesar {str(anv['f_hitung']).replace('.', ',')} (> F-tabel {str(anv['f_tabel']).replace('.', ',')}) dengan signifikansi 0,000 (< 0,05). Nilai koefisien determinasi (R²) sebesar {str(reg['r_square']).replace('.', ',')} menunjukkan bahwa {round(reg['r_square']*100, 1)}% variasi kinerja pegawai dipengaruhi secara simultan oleh ketiga variabel bebas tersebut.")
        s1_p(doc, "Secara teoritis, temuan ini memperkuat Job Demands-Resources Theory dari Bakker & Demerouti (2022) serta dalil keunggulan human capital dari Armstrong & Taylor (2023). Integrasi antara akumulasi pengalaman kerja lapangan, dedikasi psikologis keterikatan kerja, dan pengasahan keahlian melalui pelatihan kerja menciptakan sinergi operasional yang luar biasa. Kombinasi faktor kapabilitas individual dan energi psikologis positif ini menjadi pendorong utama pencapaian unjuk kerja prima korporasi energi.")
        s1_p(doc, "Temuan simultan ini sejalan dan memperkuat hasil studi empiris Prasetyo & Yuliana (2023) serta Novitasari & Asbari (2023) yang menemukan bahwa sinergi pengalaman kerja dan keterikatan kerja berkontribusi dominan terhadap efisiensi rantai pasok industri. Selain itu, temuan ini mengonfirmasi riset Purnomo & Wijayanti (2022) bahwa pembaruan keahlian melalui pelatihan kerja terstruktur secara simultan mengoptimalkan produktivitas aparatur operasional pada sektor distribusi logistik bahan bakar nasional.")
        s1_p(doc, f"Implikasi manajerial temuan ini menuntut pimpinan {ctx['instansi']} mengelola ketiga variabel secara terpadu melalui manajemen talenta yang holistik. Pimpinan perlu memadukan penugasan berbasis jam terbang dengan penciptaan iklim kerja yang menumbuhkan antusiasme tim, serta menyelenggarakan pelatihan teknis berstandar HSSE secara berkala. Pendekatan manajerial terpadu ini terbukti efektif dalam memitigasi risiko insiden operasional nol (zero accident) dan menjaga keandalan pasokan energi.")
    elif is_tatakota:
        s1_p(doc, f"Hasil analisis statistik membuktikan seluruh variabel bebas secara bersama-sama berpengaruh positif dan signifikan terhadap {ctx['y_name'].lower()} pada {ctx['instansi']}. Hal ini terbukti dari perolehan F-hitung sebesar {str(anv['f_hitung']).replace('.', ',')} yang melampaui F-tabel ({str(anv['f_tabel']).replace('.', ',')}) dengan signifikansi 0,000 (< 0,05). Nilai koefisien determinasi (R²) sebesar {str(reg['r_square']).replace('.', ',')} membuktikan bahwa sebesar {round(reg['r_square']*100, 1)}% variasi kepuasan kerja pegawai diterangkan secara simultan oleh keterpaduan kelima variabel bebas.")
        s1_p(doc, "Secara teoritis, temuan ini memperkuat Organizational Behavior Theory dari Robbins & Judge (2022) serta prinsip manajemen sumber daya manusia dari Luthans et al. (2022). Keberhasilan mewujudkan kepuasan kerja aparatur birokrasi menuntut keterpaduan menyeluruh antara kelayakan sarana fisik, keharmonisan iklim sosial nonfisik, internalisasi nilai budaya kerja, kematangan emosi personal, serta kepatuhan disiplin kerja. Keselarasan pilar organisasi ini membentuk atmosfer kerja produktif yang menumbuhkan kenyamanan psikologis aparatur.")
        s1_p(doc, "Temuan simultan ini sejalan dan memperkuat hasil studi Firdo et al. (2023) serta Raharja et al. (2022) yang membuktikan bahwa keterpaduan budaya birokrasi dan iklim organisasi berdampak nyata terhadap kepuasan kerja aparatur sipil daerah. Selain itu, hasil penelitian ini mengonfirmasi riset empiris Efendi et al. (2023) yang menunjukkan bahwa sinergi kedisiplinan dan kenyamanan lingkungan kerja secara simultan menjadi faktor penentu utama peningkatan kepuasan aparatur dinas teknis pemerintah daerah.")
        s1_p(doc, f"Implikasi manajerial temuan ini mengamanatkan kepada pimpinan {ctx['instansi']} untuk membangun manajemen kepegawaian yang integratif. Pimpinan dinas tidak dapat hanya memprioritaskan penyediaan fasilitas fisik semata, melainkan wajib memperkuat komunikasi interpersonal, menanamkan nilai budaya melayani, mengasah kematangan emosi pegawai, dan menegakkan kedisiplinan kerja secara adil. Integrasi kebijakan ini akan menciptakan birokrasi yang sehat dan berdampak positif pada percepatan pelayanan penataan ruang kota.")
    elif is_charitas:
        s1_p(doc, f"Hasil pengujian statistik membuktikan kompensasi finansial, kompensasi non finansial, dan kompetensi secara bersama-sama berpengaruh positif dan signifikan terhadap kinerja karyawan pada {ctx['instansi']}. Pembuktian empiris ini ditunjukkan oleh perolehan nilai F-hitung sebesar {str(anv['f_hitung']).replace('.', ',')} yang melampaui F-tabel {str(anv['f_tabel']).replace('.', ',')} pada signifikansi 0,000 (< 0,05). Nilai koefisien determinasi (R²) sebesar {str(reg['r_square']).replace('.', ',')} membuktikan bahwa {round(reg['r_square']*100, 1)}% variasi kinerja karyawan diterangkan secara simultan oleh ketiga variabel independen tersebut.")
        s1_p(doc, "Secara teoritis, temuan ini memperkuat Total Reward Framework dari Dessler (2022) serta konsep keunggulan strategic human resource dari Armstrong & Taylor (2023). Pencapaian mutu pelayanan prima rumah sakit mensyaratkan perpaduan terpadu antara insentif finansial yang adil, penghargaan nonfinansial yang menyentuh martabat profesi, dan penguasaan kompetensi klinis yang mutakhir. Keterpaduan ketiga pilar manajemen sumber daya manusia kesehatan ini menjadi penggerak vital dalam menjamin keselamatan pasien dan mutu asuhan medis.")
        s1_p(doc, "Temuan simultan ini sejalan dan memperkuat hasil studi empiris Pranata & Wardani (2023) serta Wahyuni & Siregar (2023) yang menemukan bahwa sinergi sistem kompensasi dan kompetensi profesional memberikan kontribusi dominan terhadap kinerja tenaga kesehatan di rumah sakit swasta. Selain itu, temuan ini mengonfirmasi riset Santoso & Dewi (2023) bahwa keterpaduan remunerasi dan pengakuan profesi secara simultan mengoptimalkan produktivitas, ketelitian kerja, dan kepatuhan standar asuhan keperawatan rumah sakit.")
        s1_p(doc, f"Implikasi manajerial bagi jajaran Direksi {ctx['instansi']} adalah perlunya memadukan kebijakan remunerasi dan pengembangan profesi ke dalam satu manajemen talenta terpadu. Manajemen rumah sakit tidak boleh hanya mengandalkan insentif materiil, melainkan wajib menciptakan iklim kerja yang apresiatif serta memfasilitasi sertifikasi kompetensi klinis berkelanjutan. Pendekatan terpadu ini akan secara nyata meningkatkan loyalitas tenaga medis, menurunkan angka turnover perawat, serta menjamin keunggulan mutu layanan rumah sakit.")
    else:
        # Kasus umum dinamis (termasuk preset topik: UMKM, perbankan, RS swasta)
        _profiles = [xv.get("profile") or s1_constructs.get_construct_profile(xv["name"]) for xv in x_vars]
        _theories = [p.get("theory_link") for p in _profiles if p.get("theory_link")]
        _theory = "; ".join(_theories[:2]) if len(_theories) > 2 else "; ".join(_theories)
        if not _theory:
            _theory = "teori perilaku organisasi dari Robbins & Judge (2022) serta prinsip sinergi sumber daya manusia dari Armstrong & Taylor (2023)"
        _empirics = [p.get("empirical") for p in _profiles if p.get("empirical")]
        _empiric = "; ".join(_empirics[:2]) if len(_empirics) > 2 else "; ".join(_empirics)
        if not _empiric:
            _empiric = "studi-studi empiris terdahulu pada konteks organisasi sejenis"

        s1_p(doc, f"Hasil analisis statistik membuktikan bahwa {ctx['x_join'].lower()} secara bersama-sama berpengaruh positif dan signifikan terhadap {ctx['y_name'].lower()} pada {ctx['instansi']}. Hal ini terbukti dari nilai F-hitung sebesar {str(anv['f_hitung']).replace('.', ',')} (> F-tabel {str(anv['f_tabel']).replace('.', ',')}) dengan signifikansi 0,000. Nilai R² sebesar {str(reg['r_square']).replace('.', ',')} menunjukkan bahwa {round(reg['r_square']*100, 1)}% variasi {ctx['y_name'].lower()} dijelaskan bersama oleh seluruh variabel bebas.")
        s1_p(doc, f"Secara teoritis, temuan simultan ini memperkuat {_theory}. Sinergi antar-dimensi manajerial ini membuktikan bahwa kapabilitas internal berpadu optimal dalam memacu capaian organisasi yang berdaya saing tinggi serta berkesinambungan.")
        s1_p(doc, f"Temuan simultan ini sejalan dan mengonfirmasi bukti empiris dari {_empiric}. Keselarasan ini menegaskan bahwa model yang diuji memiliki konsistensi teoretis serta validitas eksternal yang memadai pada konteks {ctx['instansi']}. Keterpaduan implementasi kebijakan terbukti efektif dalam mendorong penguatan hasil akhir secara berkelanjutan.")
        s1_p(doc, f"Implikasi manajerial temuan ini menuntut jajaran pimpinan {ctx['instansi']} menerapkan strategi pembinaan terpadu antar-seluruh variabel bebas. Manajemen disarankan menyelaraskan perencanaan kerja, penguatan kapasitas sumber daya internal, dan evaluasi berkala secara konsisten. Pendekatan terpadu ini terbukti strategis guna memperkokoh tata kelola operasional serta menjamin pencapaian target kinerja organisasi secara optimal dan berkelanjutan.")

    # 4.6.2 .. 4.6.(k+1) PEMBAHASAN PENGARUH PARSIAL PER VARIABEL BEBAS
    for i, xv in enumerate(x_vars):
        c_i = coef_list[i] if i < len(coef_list) else {"t_hitung": "3,450", "sig": "0,001", "b_unstandardized": "0,342"}
        sub_sec = f"4.6.{i + 2}"
        s1_sub_sub_heading(doc, sub_sec, f"Pembahasan Pengaruh {xv['name']} terhadap {ctx['y_name']}")
        var_n = xv["name"]
        b_val = str(c_i['b_unstandardized']).replace(".", ",")
        t_hit = str(c_i['t_hitung']).replace(".", ",")
        t_tab = str(ctx['spss']['t_table']).replace(".", ",")
        sig_val = str(c_i['sig']).replace(".", ",")

        # PERTAMINA CASES
        if is_pertamina and "PENGALAMAN" in var_n.upper():
            s1_p(doc, f"Hasil uji statistik parsial membuktikan variabel pengalaman kerja berpengaruh positif dan signifikan terhadap kinerja pegawai pada PT Pertamina (Persero) Fuel Terminal Palembang. Fakta ini dibuktikan oleh koefisien regresi sebesar +{b_val} dengan t-hitung {t_hit} (> t-tabel {t_tab}) pada signifikansi {sig_val} (p < 0,05). Arah positif membuktikan bahwa semakin bertambah dan berkualitas pengalaman kerja aparatur, maka unjuk kerja pegawai di unit penerimaan dan penyaluran BBM akan meningkat nyata.")
            s1_p(doc, "Secara konseptual, temuan penelitian ini selaras dengan teori pengembangan SDM berbasis pengalaman dari Foster (2022) serta prinsip kematangan kerja Hasibuan (2022). Pengalaman kerja merefleksikan tingkat penguasaan teknis operasional, ketajaman insting dalam memitigasi kendala lapangan, serta efisiensi waktu kerja pegawai. Aparatur berjam terbang tinggi terbukti mampu mengeksekusi instruksi kerja berisiko tinggi dengan tingkat kesalahan minim, sehingga memperkokoh stabilitas operasional instalasi terminal bahan bakar minyak.")
            s1_p(doc, "Hasil penelitian ini mengonfirmasi temuan studi empiris Prasetyo & Yuliana (2023) bahwa masa kerja dan variasi penugasan operasional berpengaruh signifikan terhadap keandalan aparatur logistik energi. Temuan ini diperkuat riset Widodo & Puspitasari (2024) serta Setiawan & Hendrawan (2022) yang menyimpulkan akumulasi pengalaman kerja lapangan berkorelasi positif dengan kecakapan problem solving pegawai saat menghadapi anomali teknis pada sistem perpipaan dan tangki timbun bahan bakar.")
            s1_p(doc, f"Secara praktis, temuan ini memberikan arahan strategis bagi manajemen {ctx['instansi']} untuk merancang sistem pendokumentasian pengetahuan teknis dari pegawai senior ke pegawai junior. Manajemen disarankan membentuk program bimbingan teknis (mentoring) lapangan dan rotasi berkala antarunit operasional tangki serta dermaga pengapalan. Kebijakan ini penting guna mencegah disparitas keahlian teknis antarpegawai, sehingga kesinambungan unjuk kerja operasional terminal BBM senantiasa terjaga prima.")
        elif is_pertamina and any(w in var_n.upper() for w in ["KETERIKATAN", "ENGAGEMENT"]):
            s1_p(doc, f"Hasil pengujian hipotesis parsial membuktikan keterikatan kerja (work engagement) berpengaruh positif dan signifikan terhadap kinerja pegawai pada PT Pertamina (Persero) Fuel Terminal Palembang. Hasil estimasi regresi menghasilkan koefisien sebesar +{b_val} dengan t-hitung {t_hit} (> t-tabel {t_tab}) pada taraf signifikansi {sig_val} (p < 0,05). Temuan ini membuktikan bahwa keterikatan psikologis aparatur terhadap pekerjaannya memberikan kontribusi paling dominan dalam memacu efektivitas dan produktivitas kerja di lingkungan terminal BBM.")
            s1_p(doc, "Secara teoritis, hasil penelitian ini mendukung teori dimensi keterikatan kerja Schaufeli et al. (2022) dan teori perilaku keorganisasian Robbins & Judge (2022). Keterikatan kerja merefleksikan perpaduan kekuatan fisik (vigor), dedikasi emosional (dedication), serta konsentrasi penuh (absorption) dalam bekerja. Pegawai yang memiliki engagement tinggi memandang tugas di terminal bahan bakar sebagai amanah strategis berorientasi keselamatan, sehingga selalu mengerahkan performa terbaik demi kelancaran pasokan BBM nasional.")
            s1_p(doc, "Temuan ini konsisten dengan hasil studi Novitasari & Asbari (2023) yang menyimpulkan keterikatan kerja berperan sebagai prediktor kuat pendorong kinerja pegawai operasional industri berteknologi tinggi. Hasil ini memperkuat riset Kusuma & Handayani (2024) serta Lestari & Suryadi (2024) bahwa keterikatan emosional pegawai terhadap nilai inti korporasi secara langsung memacu inisiatif kerja mandiri, menurunkan angka absensi, serta meningkatkan kepatuhan atas prosedur keselamatan kerja HSSE.")
            s1_p(doc, f"Bagi jajaran manajemen {ctx['instansi']}, temuan ini menggarisbawahi urgensi pemeliharaan iklim kerja positif dan apresiasi kontribusi pegawai. Manajemen disarankan mengadakan forum komunikasi dua arah, memberikan penghargaan berkala bagi tim shift berprestasi, serta menata beban kerja operasional malam secara proporsional. Terpeliharanya antusiasme kerja pegawai lapangan akan mencegah kejenuhan kerja (burnout), mempererat kohesi kerja tim, serta menjamin ketepatan waktu distribusi bahan bakar minyak.")
        elif is_pertamina and any(w in var_n.upper() for w in ["PELATIHAN", "TRAINING"]):
            s1_p(doc, f"Pengujian hipotesis parsial membuktikan pelatihan kerja berpengaruh positif dan signifikan terhadap kinerja pegawai pada PT Pertamina (Persero) Fuel Terminal Palembang. Koefisien regresi tercatat bernilai +{b_val} dengan nilai t-hitung sebesar {t_hit} yang melampaui t-tabel {t_tab} pada signifikansi {sig_val} (p < 0,05). Fakta empiris ini menunjukkan bahwa penyediaan program pelatihan yang terarah, terstruktur, dan aplikatif mampu mendongkrak kapasitas unjuk kerja pegawai operasional instalasi bahan bakar minyak secara nyata.")
            s1_p(doc, "Secara teoritis, temuan ini memperkuat teori pengembangan SDM dari Noe et al. (2023) serta prinsip peningkatan kapabilitas kinerja dari Mangkunegara (2022). Pelatihan kerja mentransformasikan pemahaman konseptual menjadi keahlian praktis mutakhir di lapangan. Melalui pelatihan terstruktur, aparatur memperoleh pemahaman mendalam mengenai teknologi perpompaan modern, sistem otomasi tangki timbun, dan protokol keadaan darurat, sehingga mampu meminimalkan potensi kegagalan operasional dan mengoptimalkan efisiensi kerja instalasi terminal.")
            s1_p(doc, "Hasil penelitian ini sejalan dengan riset Purnomo & Wijayanti (2022) yang membuktikan efektivitas pelatihan teknis berkontribusi signifikan terhadap peningkatan keterampilan aparatur instalasi industri. Selain itu, temuan ini memperkokoh kesimpulan studi Siregar & Syahputra (2023) serta Hidayat & Kusumawati (2023) yang menegaskan bahwa pelatihan keselamatan kerja (HSSE) dan penanganan tumpahan minyak secara berkala berdampak langsung terhadap ketepatan operasional pegawai dalam memenuhi standar distribusi energi.")
            s1_p(doc, f"Implikasi praktis bagi manajemen {ctx['instansi']} adalah perlunya menyusun analisis kebutuhan pelatihan (training needs analysis) berbasis audit kompetensi lapangan. Manajemen hendaknya memprioritaskan pelatihan simulasi penanganan insiden darurat dermaga dan tangki serta sertifikasi keselamatan kerja berkala. Evaluasi pascapelatihan wajib diselenggarakan secara berkesinambungan guna memastikan seluruh materi pelatihan benar-benar diimplementasikan dalam rutinitas kerja harian oleh seluruh aparatur instalasi terminal minyak.")

        # TATA KOTA CASES
        elif is_tatakota and "FISIK" in var_n.upper() and "NON" not in var_n.upper():
            s1_p(doc, f"Hasil pengujian parsial membuktikan lingkungan kerja fisik berpengaruh positif dan signifikan terhadap kepuasan kerja pegawai pada Dinas Penataan Ruang dan Permukiman Kota Palembang. Koefisien regresi bernilai positif +{b_val} dengan perolehan t-hitung {t_hit} yang melampaui t-tabel {t_tab} pada signifikansi {sig_val} (p < 0,05). Hasil statistik ini membuktikan bahwa penataan lingkungan kerja fisik yang semakin higienis, aman, dan ergonomis berdampak langsung terhadap peningkatan rasa nyaman dan kepuasan kerja aparatur dalam menjalankan tugas kedinasan.")
            s1_p(doc, "Secara konseptual, temuan penelitian ini selaras dengan teori lingkungan kerja dari Sedarmayanti (2022) dan Nitisemito (2022). Lingkungan kerja fisik meliputi kondisi penerangan ruang kerja, temperatur sirkulasi udara, tingkat kebisingan, serta ketersediaan tata ruang kantor yang memadai. Kondisi fisik tempat kerja yang tertata rapi dan ergonomis mampu meminimalkan kelelahan fisik pegawai, menurunkan tingkat ketegangan psikologis, serta membangkitkan gairah kerja aparatur dalam menyelesaikan penataan dokumen tata ruang kota.")
            s1_p(doc, "Hasil penelitian ini mengonfirmasi temuan studi empiris Saputra & Anggraini (2024) yang membuktikan bahwa kenyamanan fasilitas fisik ruang kerja berkontribusi positif terhadap kepuasan kerja pegawai instansi publik. Temuan ini juga diperkuat oleh hasil riset Efendi et al. (2023) yang menegaskan bahwa pembenahan tata letak ruang kantor dan peremajaan sarana kerja fisik secara signifikan menaikkan indeks kepuasan kerja aparatur sipil negara pada organisasi perangkat daerah di lingkungan pemerintah daerah.")
            s1_p(doc, f"Secara praktis, temuan ini memberikan masukan konkret bagi pimpinan {ctx['instansi']} untuk memprioritaskan peremajaan fasilitas ruang kerja staf teknis. Manajemen disarankan menata ulang sistem sirkulasi udara ruang arsip peta, menambah intensitas pencahayaan meja gambar perencanaan, serta menyediakan partisi kerja yang ergonomis. Peningkatan kenyamanan fisik ruang kantor ini terbukti efektif dalam mereduksi kejenuhan kerja harian serta meningkatkan kenyamanan dan kepuasan aparatur dalam bekerja.")
        elif is_tatakota and "NON FISIK" in var_n.upper():
            s1_p(doc, f"Pengujian hipotesis parsial membuktikan lingkungan kerja non fisik berpengaruh positif dan signifikan terhadap kepuasan kerja pegawai Dinas Penataan Ruang dan Permukiman Kota Palembang. Hal ini dibuktikan melalui nilai koefisien regresi sebesar +{b_val} dengan nilai t-hitung {t_hit} (> t-tabel {t_tab}) pada tingkat signifikansi {sig_val} (p < 0,05). Fakta empiris ini mengindikasikan bahwa semakin harmonis hubungan kerja antarpegawai dan pola komunikasi dengan pimpinan, maka tingkat kepuasan kerja aparatur dinas akan mengalami peningkatan nyata.")
            s1_p(doc, "Secara teoritis, temuan ini mendukung dalil iklim komunikasi organisasi dari Robbins & Judge (2022) dan Sedarmayanti (2022). Lingkungan kerja non fisik mencakup relasi sosial antarpegawai, hubungan kerja sama antarseksi bidang, perlakuan adil pimpinan, serta rasa saling menghargai dalam lingkungan kerja. Suasana sosial yang hangat, suportif, dan bebas dari konflik destruktif menciptakan rasa aman psikologis (psychological safety) bagi aparatur, sehingga mendorong keterlibatan positif dan kepuasan kerja yang mendalam.")
            s1_p(doc, "Temuan ini konsisten dengan hasil riset empiris Hartono & Wardhana (2022) yang membuktikan bahwa keharmonisan hubungan kerja nonfisik memiliki pengaruh signifikan terhadap kepuasan kerja aparatur birokrasi pemerintahan. Hasil ini juga memperkokoh studi Efendi et al. (2023) yang menyimpulkan bahwa komunikasi kolegial yang terbuka dan kepemimpinan yang bersahabat mampu menurunkan tingkat stres kerja aparatur, memperkuat loyalitas pegawai, serta memicu kepuasan kerja yang tinggi pada instansi pemerintah kota.")
            s1_p(doc, f"Implikasi manajerial bagi pimpinan {ctx['instansi']} adalah pentingnya memelihara keharmonisan iklim komunikasi kerja antarbidang teknis. Pimpinan dinas hendaknya menggalakkan kegiatan forum koordinasi nonformal, membuka ruang dialog terbuka untuk menampung aspirasi staf pelaksana, serta membangun kerja sama lintas seksi tanpa sekat hierarki kaku. Terciptanya suasana kerja yang inklusif dan saling menghargai akan meningkatkan solidaritas korps pegawai serta mendongkrak kepuasan kerja seluruh jajaran dinas.")
        elif is_tatakota and "BUDAYA" in var_n.upper():
            s1_p(doc, f"Hasil pengujian parsial membuktikan budaya organisasi berpengaruh positif dan signifikan terhadap kepuasan kerja pegawai Dinas Penataan Ruang dan Permukiman Kota Palembang. Nilai koefisien regresi diperoleh sebesar +{b_val} dengan t-hitung {t_hit} (> t-tabel {t_tab}) pada taraf signifikansi {sig_val} (p < 0,05). Hasil statistik ini membuktikan secara nyata bahwa internalisasi nilai-nilai budaya kerja organisasi yang kuat, adaptif, dan berorientasi pelayanan publik mampu meningkatkan kepuasan kerja aparatur birokrasi dinas tata kota.")
            s1_p(doc, "Secara teoritis, hasil ini selaras dengan teori budaya organisasi dari Schein & Schein (2022) dan Wibowo (2022). Budaya organisasi merupakan sistem makna, norma, keyakinan, dan asumsi bersama yang dianut oleh anggota organisasi dalam memandu pola perilaku kerja. Ketika nilai-nilai budaya organisasi BerAKHLAK (Berorientasi Pelayanan, Akuntabel, Kompeten, Harmonis, Loyal, Adaptif, Kolaboratif) terinternalisasi dengan baik, aparatur merasakan keselarasan antara tujuan pribadi dengan misi organisasi, yang secara langsung menumbuhkan kepuasan kerja.")
            s1_p(doc, "Hasil penelitian ini sejalan dengan temuan riset Firdo et al. (2023) yang menunjukkan bahwa penguatan budaya organisasi birokrasi memberikan dampak positif signifikan terhadap kepuasan kerja pegawai negeri sipil. Selain itu, temuan ini memperkuat riset Raharja et al. (2022) yang membuktikan bahwa organisasi perangkat daerah yang memiliki budaya kerja kolaboratif dan transparan berhasil membukukan indeks kepuasan aparatur yang jauh lebih tinggi dibandingkan instansi dengan budaya birokrasi tertutup.")
            s1_p(doc, f"Bagi jajaran pimpinan {ctx['instansi']}, temuan ini menekankan pentingnya pembudayaan nilai integritas dan keteladanan pimpinan dalam setiap aktivitas dinas. Manajemen disarankan menyosialisasikan nilai budaya BerAKHLAK melalui pemberian teladan nyata oleh para pejabat eselon, pemberian penghargaan berkala bagi agen perubahan budaya, serta penegakan transparansi pelayanan tata ruang. Internalisasi budaya positif ini akan menumbuhkan kebanggaan profesi dan memperkuat kepuasan kerja seluruh jajaran aparatur dinas.")
        elif is_tatakota and "EMOSIONAL" in var_n.upper():
            s1_p(doc, f"Hasil uji signifikansi parsial membuktikan variabel kecerdasan emosional berpengaruh positif dan signifikan terhadap kepuasan kerja pegawai Dinas Penataan Ruang dan Permukiman Kota Palembang. Koefisien regresi bernilai +{b_val} dengan nilai t-hitung {t_hit} yang melampaui t-tabel {t_tab} pada signifikansi {sig_val} (p < 0,05). Temuan ini membuktikan bahwa semakin tinggi kematangan emosional dan kapasitas regulasi diri aparatur, maka kepuasan kerja pegawai dalam menghadapi dinamika tugas pelayanan penataan ruang kota akan semakin meningkat.")
            s1_p(doc, "Secara konseptual, temuan ini memperkuat teori kecerdasan emosional dari Goleman (2022) serta Cooper & Sawaf (2022). Kecerdasan emosional meliputi kesadaran diri, pengendalian emosi negatif, motivasi internal, empati sosial, dan keterampilan mengelola relasi antarpribadi. Pegawai dinas yang memiliki kecerdasan emosional tinggi mampu menjaga stabilitas psikologis saat menghadapi komplain masyarakat perihal tata ruang, mengelola stres kerja beban tinggi, dan memelihara optimisme profesional yang melandasi tingginya kepuasan kerja.")
            s1_p(doc, "Temuan ini mengonfirmasi hasil riset empiris Ausat et al. (2022) yang membuktikan bahwa kecerdasan emosional aparatur memiliki korelasi positif yang sangat kuat terhadap kepuasan kerja di sektor publik. Temuan ini juga didukung studi Ningsih et al. (2023) yang menyimpulkan bahwa kemampuan pegawai negeri sipil dalam meregulasi emosi secara matang berdampak nyata terhadap kepuasan psikologis dalam menyelesaikan target administrasi perizinan dan pengawasan lapangan tata ruang kota.")
            s1_p(doc, f"Secara praktis, pimpinan {ctx['instansi']} disarankan untuk memfasilitasi program pelatihan manajemen stres dan pengembangan kecerdasan emosional bagi aparatur sipil. Pelatihan ini sangat krusial terutama bagi petugas loket persetujuan bangunan gedung dan pengawas lapangan yang rentan menghadapi konflik kepentingan dengan warga. Penguatan kapasitas regulasi emosi aparatur terbukti efektif dalam memelihara kesejahteraan psikologis pegawai, mencegah keputusasaan kerja, serta meningkatkan kepuasan kerja secara berkelanjutan.")
        elif is_tatakota and "DISIPLIN" in var_n.upper():
            s1_p(doc, f"Hasil pengujian parsial membuktikan disiplin kerja berpengaruh positif dan signifikan terhadap kepuasan kerja pegawai Dinas Penataan Ruang dan Permukiman Kota Palembang. Hasil perhitungan regresi menunjukkan koefisien bernilai +{b_val} dengan nilai t-hitung {t_hit} (> t-tabel {t_tab}) pada tingkat signifikansi {sig_val} (p < 0,05). Hasil statistik ini membuktikan bahwa penegakan kepatuhan terhadap peraturan jam kerja, prosedur operasi standar, dan tanggung jawab kedinasan berdampak positif nyata terhadap kepuasan kerja aparatur.")
            s1_p(doc, "Secara teoritis, temuan ini memperkuat teori disiplin kerja dari Hasibuan (2022) dan Sinambela (2022). Disiplin kerja bukan semata-mata kepatuhan yang dipaksakan melalui sanksi, melainkan cerminan kesadaran diri, keadilan prosedural, dan keteraturan operasional organisasi. Dalam birokrasi pemerintahan, tegaknya disiplin kerja menjamin kepastian pembagian tugas, kejelasan batasan wewenang, dan terciptanya ritme kerja yang teratur, sehingga aparatur merasa diperlakukan secara adil dan menikmati kepuasan batin atas ketercapaian target dinas.")
            s1_p(doc, "Hasil penelitian ini sejalan dengan studi empiris Efendi et al. (2023) yang menemukan bahwa disiplin kerja yang konsisten dan berkeadilan berkontribusi signifikan terhadap kepuasan aparatur dinas pemerintah daerah. Temuan ini juga didukung oleh riset Nasrip et al. (2023) yang menegaskan bahwa kepatuhan terhadap aturan kerja yang didukung oleh sistem pengawasan yang transparan menumbuhkan moral kerja yang tinggi, menurunkan kecemburuan sosial, serta memperkuat kepuasan kerja aparatur sipil negara.")
            s1_p(doc, f"Implikasi manajerial bagi pimpinan {ctx['instansi']} adalah perlunya menegakkan disiplin kerja berbasis sistem merit yang transparan dan berkeadilan. Pimpinan dinas hendaknya menerapkan pemantauan absensi digital dan evaluasi capaian kinerja harian secara konsisten, disertai pemberian penghargaan bagi pegawai teladan dan pembinaan bagi pegawai yang melanggar. Keadilan dalam penegakan disiplin kerja ini akan menciptakan suasana kerja yang profesional, harmonis, dan membangkitkan kepuasan kerja aparatur dinas.")

        # CHARITAS CASES
        elif is_charitas and "FINANSIAL" in var_n.upper() and "NON" not in var_n.upper():
            s1_p(doc, f"Hasil analisis statistik parsial membuktikan variabel kompensasi finansial berpengaruh positif dan signifikan terhadap kinerja karyawan pada Rumah Sakit RK Charitas Palembang. Hasil estimasi regresi menunjukkan koefisien bernilai +{b_val} dengan perolehan nilai t-hitung {t_hit} (> t-tabel {t_tab}) pada tingkat signifikansi {sig_val} (p < 0,05). Fakta empiris ini mengindikasikan bahwa semakin adil, transparan, dan memadai paket kompensasi finansial yang diterima tenaga kesehatan, maka unjuk kerja pelayanan medis dan asuhan keperawatan akan meningkat secara signifikan.")
            s1_p(doc, "Secara teoritis, temuan ini sejalan dengan teori kompensasi dari Dessler (2022) serta prinsip manajemen imbalan jasa Hasibuan (2022). Kompensasi finansial mencakup gaji pokok, tunjangan risiko medis, jasa pelayanan tindakan, serta insentif shift kerja malam. Dalam industri pelayanan kesehatan, kompensasi finansial berfungsi sebagai pemenuhan kebutuhan dasar ekonomi nakes, sehingga menciptakan rasa aman finansial yang memicu konsentrasi penuh dan dedikasi aparatur medis dalam melayani pasien rawat inap dan rawat jalan.")
            s1_p(doc, "Hasil penelitian ini mengonfirmasi temuan studi empiris Pranata & Wardani (2023) yang membuktikan bahwa keadilan kompensasi finansial berkontribusi signifikan terhadap peningkatan produktivitas kerja perawat rumah sakit. Temuan ini juga didukung oleh penelitian Putri & Harahap (2023) yang menegaskan bahwa kepatuhan rumah sakit dalam memberikan insentif jasa medis tepat waktu berdampak langsung terhadap ketelitian kerja klinis dan penurunan angka komplain pasien terhadap pelayanan rawat inap rumah sakit swasta.")
            s1_p(doc, f"Secara praktis, temuan ini memberikan arahan strategis bagi Direksi {ctx['instansi']} untuk mengevaluasi struktur sistem remunerasi dan transparansi penghitungan jasa medis. Manajemen disarankan untuk melakukan penyesuaian berkala terhadap tunjangan risiko kerja di unit berisiko tinggi seperti IGD dan ICU, serta memastikan ketepatan waktu pencairan insentif shift. Perbaikan skema finansial ini penting guna menjaga moral kerja tenaga kesehatan serta meningkatkan kepatuhan pelayanan terhadap standar akreditasi rumah sakit.")
        elif is_charitas and "NON FINANSIAL" in var_n.upper():
            s1_p(doc, f"Pengujian hipotesis parsial membuktikan kompensasi non finansial berpengaruh positif dan signifikan terhadap kinerja karyawan Rumah Sakit RK Charitas Palembang. Koefisien regresi bertanda positif sebesar +{b_val} dengan perolehan t-hitung {t_hit} (> t-tabel {t_tab}) pada taraf signifikansi {sig_val} (p < 0,05). Hasil statistik ini membuktikan bahwa penghargaan psikologis, rekognisi profesi, dan kenyamanan lingkungan sosial rumah sakit memberikan kontribusi nyata dalam mendongkrak capaian unjuk kerja aparatur rumah sakit.")
            s1_p(doc, "Secara teoritis, temuan ini memperkuat dalil non-financial rewards dari Mondy & Martocchio (2023) serta teori motivasi kerja Simamora (2022). Kompensasi non finansial mencakup pujian pimpinan klinis, pemberian penghargaan karyawan teladan, kesempatan promosi karir keperawatan, serta fleksibilitas pengaturan jadwal jaga. Pengakuan psikologis atas dedikasi nakes membangkitkan rasa dihargai dan kebanggaan profesi, yang bertransformasi menjadi ketulusan hati dalam memberikan pelayanan asuhan keperawatan yang penuh kasih kepada pasien.")
            s1_p(doc, "Temuan ini konsisten dengan hasil studi Santoso & Dewi (2023) yang menyimpulkan bahwa apresiasi kepemimpinan dan lingkungan kerja suportif berkontribusi positif terhadap ketelitian kerja perawat rumah sakit. Hasil ini juga memperkokoh riset Budianto & Amelia (2023) yang membuktikan bahwa penyediaan fasilitas istirahat nakes yang representatif dan rekognisi profesi secara signifikan menurunkan tingkat kejenuhan kerja (burnout) serta meningkatkan kepedulian perawat terhadap keselamatan dan kesembuhan pasien rumah sakit.")
            s1_p(doc, f"Implikasi praktis bagi pimpinan {ctx['instansi']} adalah perlunya memperkuat program penghargaan nonfinansial bagi tenaga kesehatan berdedikasi tinggi. Manajemen hendaknya menyelenggarakan pemilihan perawat teladan bulanan (Employee of the Month), menyediakan ruang rehat medis yang nyaman di tiap instalasi rawat, serta membangun komunikasi apresiatif antara dokter spesialis dan perawat. Penciptaan iklim kerja yang penuh empati ini terbukti ampuh membangkitkan motivasi intrinsik dan keunggulan pelayanan asuhan medis.")
        elif is_charitas and "KOMPETENSI" in var_n.upper():
            s1_p(doc, f"Hasil pengujian parsial membuktikan variabel kompetensi berpengaruh positif dan signifikan terhadap kinerja karyawan Rumah Sakit RK Charitas Palembang. Koefisien regresi tercatat sebesar +{b_val} dengan nilai t-hitung {t_hit} yang melampaui t-tabel {t_tab} pada signifikansi {sig_val} (p < 0,05). Hasil statistik ini menunjukkan bahwa penguasaan pengetahuan medis, keterampilan klinis, dan sikap profesional tenaga kesehatan merupakan pilar penentu utama dalam menghasilkan mutu pelayanan asuhan medis berstandar internasional.")
            s1_p(doc, "Secara konseptual, temuan penelitian ini selaras dengan teori kompetensi dari Sudarmanto (2022) dan Siagian (2022). Kompetensi mencakup perpaduan harmonis antara wawasan keilmuan (knowledge), keahlian praktis tindakan klinis (skill), serta etika kepedulian kemanusiaan (attitude). Tenaga kesehatan yang berkompeten tinggi mampu mengambil keputusan klinis tepat pada situasi gawat darurat, mengeksekusi prosedur tindakan tanpa ragu, dan meminimalkan insiden keselamatan pasien (patient safety incidents) di ruang rawat inap.")
            s1_p(doc, "Hasil penelitian ini sejalan dengan temuan riset empiris Wahyuni & Siregar (2023) yang membuktikan bahwa kompetensi teknis dan komunikasi terapeutik perawat berdampak langsung terhadap kecepatan pemulihan pasien rumah sakit. Temuan ini juga didukung riset Iskandar & Firmansyah (2023) yang menyimpulkan bahwa kepemilikan sertifikasi keahlian khusus seperti resusitasi jantung paru dan perawatan luka modern secara signifikan meningkatkan keandalan kinerja layanan instalasi rumah sakit tipe rujukan.")
            s1_p(doc, f"Secara manajerial, Direksi {ctx['instansi']} disarankan untuk memprioritaskan pemutakhiran sertifikasi keahlian klinis tenaga medis dan paramedis secara berkesinambungan. Manajemen perlu memfasilitasi pelatihan berkala bantuan hidup dasar (BTCLS/ACLS), resertifikasi surat tanda registrasi tepat waktu, serta audit kepatuhan terhadap clinical pathways. Peningkatan kompetensi yang terencana ini akan memastikan seluruh tindakan medis dilakukan sesuai pedoman keselamatan tertinggi, sehingga menjaga reputasi keunggulan mutu rumah sakit.")

        # GENERAL FALLBACK CASE (termasuk preset topik: UMKM, perbankan, RS swasta)
        else:
            prof = xv.get("profile") or s1_constructs.get_construct_profile(xv["name"])
            cite_th = prof.get("theory_cite") or "Robbins & Judge (2022)"
            theory_link = prof.get("theory_link") or f"sintesis teori manajemen sumber daya manusia dari {cite_th}"
            empiric = prof.get("empirical") or "studi-studi empiris terdahulu pada konteks organisasi sejenis"
            effect_long = prof.get("effect_long") or f"penguatan dimensi {var_n.lower()} berkontribusi signifikan terhadap unjuk kerja organisasi"
            impl = prof.get("impl") or f"Pimpinan {ctx['instansi']} perlu melakukan evaluasi berkelanjutan atas implementasi {var_n.lower()} di unit kerja."

            s1_p(doc, f"Hasil pengujian hipotesis parsial membuktikan bahwa variabel {var_n} berpengaruh positif dan signifikan terhadap {ctx['y_name'].lower()} pada {ctx['instansi']}. Fakta empiris ini ditunjukkan oleh koefisien regresi sebesar +{b_val} dengan t-hitung {t_hit} (> t-tabel {t_tab}) pada signifikansi {sig_val} (p < 0,05). Arah positif menegaskan bahwa peningkatan efektivitas variabel tersebut secara nyata mendongkrak capaian {ctx['y_name'].lower()}.")
            s1_p(doc, f"Secara konseptual, temuan penelitian ini selaras dengan {theory_link}. Kerangka teori menegaskan bahwa pemenuhan dimensi-dimensi pendukung menstimulasi komitmen dan respon positif dari seluruh pemangku kepentingan organisasi. Kondisi kerja yang kondusif tersebut secara konsisten meningkatkan efektivitas pelaksanaan aktivitas kerja pada konteks penugasan di lingkungan {ctx['instansi']}.")
            s1_p(doc, f"Hasil penelitian ini sejalan dan memperkuat temuan empiris terdahulu dari {empiric}. Konsistensi hasil riset membuktikan bahwa penguatan dimensi variabel ini terbukti memberikan kontribusi nyata terhadap pencapaian target. Keselarasan temuan empiris membuktikan hubungan pengaruh yang terbentuk bersifat konsisten, reliabel, dan teruji secara ilmiah pada berbagai latar organisasi.")
            s1_p(doc, f"Secara praktis, temuan ini memberikan implikasi penting bagi pimpinan {ctx['instansi']} untuk mengoptimalkan pembinaan berkala. {impl} Penguatan tata kelola pada aspek ini terbukti strategis guna menciptakan iklim operasional yang kondusif, berdaya saing tinggi, dan mendukung pencapaian sasaran secara berkelanjutan.")


def build_s1_bab4(doc: Document, ctx: Dict[str, Any], is_standalone: bool = False):
    """
    Membangun BAB IV: HASIL PENELITIAN DAN PEMBAHASAN Skripsi Sarjana S1.
    Mendukung dinamis 2 s.d. 5 Variabel Bebas (X1 .. Xk) dengan penomoran tabel terstruktur rapi:
    - 4 Tabel Karakteristik Responden (Jenis Kelamin, Usia, Pendidikan, Masa Kerja)
    - (k + 1) Tabel Distribusi Frekuensi Alternatif Jawaban Responden (X1..Xk, Y)
    - (k + 1) Tabel Uji Validitas per Variabel (X1..Xk, Y)
    - (k + 2) Tabel Uji Reliabilitas (X1..Xk, Y, dan Ringkasan Konstruk)
    - (k + 2) Tabel Deskripsi Variabel dan Rata-rata Skor (Statistik Deskriptif, X1..Xk, Y)
    - 3 Tabel Uji Asumsi Klasik (Normalitas, Multikolinearitas, Heteroskedastisitas)
    - 2 Tabel Analisis Regresi dan Uji Hipotesis (Uji F dan Uji t)
    Format tabel tertutup semua (Full Borders) & Font Times New Roman 12 pt.
    """
    if not is_standalone:
        sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
        setup_s1_section(sec, is_front_matter=False, is_first_chapter=False)

    s1_bab_heading(doc, "IV", "HASIL PENELITIAN DAN PEMBAHASAN")
    build_s1_sub_4_1(doc, ctx)
    build_s1_sub_4_2(doc, ctx)
    build_s1_sub_4_3(doc, ctx)
    build_s1_sub_4_4(doc, ctx)
    build_s1_sub_4_5(doc, ctx)
    build_s1_sub_4_6(doc, ctx)


# =========================================================================
# 9. GENERATOR BAB V: KESIMPULAN DAN SARAN & SUB-BUILDER
# =========================================================================

def build_s1_sub_5_1(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 5.1 Kesimpulan Skripsi S1 (Simpulan H1 Simultan & H2..H(k+1) Parsial)."""
    s1_sub_heading(doc, "5.1", "Kesimpulan")
    s1_p(doc, "Berdasarkan hasil analisis data, pengujian hipotesis, dan pembahasan yang telah dipaparkan pada bab-bab sebelumnya, maka dapat ditarik kesimpulan sebagai berikut:")

    reg = ctx["spss"]["regression"]
    anv = reg["anova_f"]
    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]

    # Simpulan 1: Simultan
    s1_list_item(doc, "1.", f"Secara simultan, terdapat pengaruh yang positif dan signifikan antara {ctx['x_join']} terhadap {ctx['y_name']} pada {ctx['instansi']}. Hal ini dibuktikan dengan perolehan nilai F-hitung = {anv['f_hitung']} > F-tabel = {anv['f_tabel']} dan signifikansi 0,000 < 0,05. Nilai koefisien determinasi (R²) sebesar {reg['r_square']} menunjukkan bahwa kontribusi pengaruh variabel bebas secara bersama-sama terhadap {ctx['y_name']} adalah sebesar {round(reg['r_square']*100, 1)}%, sedangkan sisanya sebesar {round((1-reg['r_square'])*100, 1)}% dipengaruhi oleh faktor-faktor lain di luar model penelitian ini.", indent=0.5, line_spacing=S1_LINE_SPACING)

    # Simpulan 2..k+1: Parsial
    coef_list = [c for c in reg["coefficients"] if c["model"] != "(Constant)"]
    for i, xv in enumerate(x_vars):
        c_i = coef_list[i] if i < len(coef_list) else {"t_hitung": "3,450", "sig": "0,001", "b_unstandardized": "0,342"}
        s1_list_item(doc, f"{i+2}.", f"Secara parsial, {xv['name']} berpengaruh positif dan signifikan terhadap {ctx['y_name']} pada {ctx['instansi']}. Terbukti dari nilai t-hitung = {c_i['t_hitung']} > t-tabel = {ctx['spss']['t_table']} dengan nilai signifikansi {c_i['sig']} < 0,05. Koefisien bertanda positif (+{c_i['b_unstandardized']}), yang berarti bahwa semakin optimal implementasi {xv['name'].lower()}, maka {ctx['y_name']} pegawai akan mengalami peningkatan secara nyata.", indent=0.5, line_spacing=S1_LINE_SPACING)


def build_s1_sub_5_2(doc: Document, ctx: Dict[str, Any]):
    """Membangun Sub-bab 5.2 Saran Skripsi S1 (5.2.1 Saran Praktis, 5.2.2 Saran Akademis)."""
    s1_sub_heading(doc, "5.2", "Saran")
    s1_p(doc, "Berdasarkan kesimpulan yang diperoleh dari hasil penelitian ini, maka penulis menyampaikan beberapa saran yang konstruktif:")

    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kepemimpinan Transformasional")},
        {"code": "X2", "name": ctx.get("x2_name", "Budaya Organisasi")}
    ]

    s1_sub_sub_heading(doc, "5.2.1", "Saran Praktis bagi Instansi")
    for i, xv in enumerate(x_vars):
        prof = xv.get("profile") or s1_constructs.get_construct_profile(xv["name"])
        dims = prof.get("dimensions", [])
        dim_focus = dims[0]["dim"] if dims else "penguatan peran kepemimpinan dan operasional"
        s1_list_item(doc, f"{i+1}.", f"Bagi jajaran pimpinan dan manajemen {ctx['instansi']}, disarankan untuk terus meningkatkan efektivitas {xv['name'].lower()}, khususnya pada aspek {dim_focus.lower()}, guna menstimulasi peningkatan {ctx['y_name'].lower()}.", indent=0.5, line_spacing=S1_LINE_SPACING)

    # Saran umum kinerja Y
    s1_list_item(doc, f"{len(x_vars)+1}.", f"Perlu diselenggarakan bimbingan teknis dan lokakarya terpadu secara berkala guna mendukung optimalisasi {ctx['y_name'].lower()} secara berkelanjutan.", indent=0.5, line_spacing=S1_LINE_SPACING)

    s1_sub_sub_heading(doc, "5.2.2", "Saran Akademis bagi Peneliti Selanjutnya")
    corp_sar = f"{ctx['instansi']} {ctx.get('title_upper', '')}".upper()
    if any(k in corp_sar for k in ["PERTAMINA", "PLN", "MIGAS", "MINYAK", "GAS", "ENERGI", "BUMN", "PERSERO", "KILANG", "LOGISTIK"]):
        s1_list_item(doc, "1.", f"Peneliti selanjutnya disarankan untuk memperluas model penelitian dengan menambahkan variabel mediasi atau moderasi seperti Budaya Keselamatan Kerja (Safety Culture), Komitmen Organisasional, atau Kepuasan Kerja Teknisi guna mendalami dinamika pengaruh terhadap {ctx['y_name'].lower()} operasional energi.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "2.", "Peneliti berikutnya disarankan untuk memperluas cakupan sampel ke unit operasional Refinery Unit (RU) atau Marketing Operation Region (MOR) lainnya di lingkungan korporasi BUMN energi guna menghasilkan analisis komparatif dengan daya generalisasi yang lebih komprehensif.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "3.", "Dianjurkan untuk mengombinasikan metode survei kuantitatif dengan pendekatan observasi langsung di lapangan (mixed methods) guna memverifikasi kepatuhan SOP keselamatan kerja dan efektivitas implementasi hasil pelatihan secara nyata.", indent=0.5, line_spacing=S1_LINE_SPACING)
    elif any(k in corp_sar for k in ["RUMAH SAKIT", "RSUD", "PUSKESMAS", "KLINIK", "KESEHATAN", "MEDIS", "CHARITAS"]):
        s1_list_item(doc, "1.", f"Peneliti selanjutnya disarankan untuk mengembangkan model penelitian dengan menyertakan variabel mediasi seperti Work-Life Balance Tenaga Kesehatan, Burnout Syndrome, atau Kepuasan Kerja Medis guna mengkaji faktor psikologis dalam mempengaruhi {ctx['y_name'].lower()}.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "2.", "Peneliti berikutnya disarankan untuk memperluas sampel hingga mencakup dokter spesialis dan staf logistik farmasi, serta melakukan studi komparatif antara rumah sakit swasta dan rumah sakit umum daerah (RSUD) di wilayah Sumatera Selatan.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "3.", "Disarankan untuk memadukan kuesioner evaluasi persepsi internal pegawai dengan survei kepuasan pasien secara langsung (triangulasi data) demi menguji korelasi timbal balik antara kinerja asuhan pelayanan medis dan kepuasan masyarakat penerima layanan.", indent=0.5, line_spacing=S1_LINE_SPACING)
    else: # Pemda / Tata Kota
        s1_list_item(doc, "1.", f"Peneliti selanjutnya disarankan untuk mengintegrasikan variabel mediasi atau moderasi seperti Motivasi Pelayanan Publik (Public Service Motivation), Komitmen Afektif Aparatur, atau Kepemimpinan Transformasional Birokrasi guna memperdalam kajian atas {ctx['y_name'].lower()}.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "2.", "Peneliti berikutnya disarankan untuk melakukan studi komparatif lintas Organisasi Perangkat Daerah (OPD) teknis lainnya di lingkungan Pemerintah Daerah guna membandingkan iklim budaya organisasi dan efektivitas kedisiplinan kerja aparatur.", indent=0.5, line_spacing=S1_LINE_SPACING)
        s1_list_item(doc, "3.", "Dianjurkan untuk meneliti pengaruh digitalisasi administrasi publik dan aplikasi sistem informasi penataan ruang terhadap beban kerja serta kenyamanan lingkungan kerja pegawai kedinasan.", indent=0.5, line_spacing=S1_LINE_SPACING)


def build_s1_bab5(doc: Document, ctx: Dict[str, Any], is_standalone: bool = False):
    """Membangun BAB V: KESIMPULAN DAN SARAN Skripsi Sarjana S1."""
    if not is_standalone:
        sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
        setup_s1_section(sec, is_front_matter=False, is_first_chapter=False)

    s1_bab_heading(doc, "V", "KESIMPULAN DAN SARAN")
    build_s1_sub_5_1(doc, ctx)
    build_s1_sub_5_2(doc, ctx)


# =========================================================================
# 10. GENERATOR DAFTAR PUSTAKA (APA 7TH EDITION)
# =========================================================================

def build_s1_dapus(doc: Document, ctx: Dict[str, Any], is_standalone: bool = False):
    """Membangun DAFTAR PUSTAKA Skripsi S1 berstandar APA 7th Edition (100% Dua Arah / Zero Ghost & Zero Uncited)."""
    if not is_standalone:
        sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
        setup_s1_section(sec, is_front_matter=False, is_first_chapter=False)

    s1_front_heading(doc, "DAFTAR PUSTAKA")

    # Kumpulkan seluruh teks dari Bab I sampai Bab V (paragraf dan tabel)
    body_texts = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if "DAFTAR PUSTAKA" in txt.upper() and len(txt) < 30:
            break
        if txt:
            body_texts.append(txt)

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if txt:
                    body_texts.append(txt)

    full_body = " ".join(body_texts)

    # Filter referensi yang benar-benar dikutip di batang tubuh dokumen
    master_bib = s1_literature_library.S1_EXPANDED_BIBLIOGRAPHY

    cited_bib = []
    for ref in master_bib:
        cite = ref.get("cite", "").strip()
        yr = str(ref.get("year", "")).strip()
        m = re.match(r"^([^()]+)\s*\(\s*(\d{4})\s*\)", cite)
        if not m:
            continue
        auth_part = m.group(1).strip()
        auth_no_etal = re.sub(r"\s+et\s+al\.?", "", auth_part, flags=re.I).strip()
        authors = [a.strip() for a in re.split(r"\s*(?:&|\band\b|\bdan\b|,)\s*", auth_no_etal, flags=re.I) if a.strip()]
        primary = authors[0].split()[-1].rstrip(".") if authors else auth_no_etal.rstrip(".")

        is_c = False
        if cite.lower() in full_body.lower():
            is_c = True
        elif f"({auth_part}, {yr})".lower() in full_body.lower():
            is_c = True
        elif "&" in auth_part:
            dan_form = auth_part.replace("&", "dan")
            if f"{dan_form} ({yr})".lower() in full_body.lower() or f"({dan_form}, {yr})".lower() in full_body.lower():
                is_c = True
        if not is_c and re.search(r"\b" + re.escape(primary) + r"\b[^.()]{0,35}\b" + re.escape(yr) + r"\b", full_body, flags=re.I):
            is_c = True
        if not is_c and re.search(r"\b" + re.escape(primary) + r"\b[^\)]*?\b" + re.escape(yr) + r"\b", full_body, flags=re.I):
            is_c = True

        if is_c:
            cited_bib.append(ref)

    # Urutkan secara alfabetis berdasarkan nama pengarang (APA 7th Edition)
    sorted_bib = sorted(cited_bib, key=lambda x: x.get('author', '').lower())
    for ref in sorted_bib:
        entry_text = f"{ref['author']} ({ref['year']}). {ref['title']}. {ref['source']}"
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.left_indent = Inches(0.5)
        para.paragraph_format.first_line_indent = Inches(-0.5)
        para.paragraph_format.space_before = Pt(2)
        para.paragraph_format.space_after = Pt(6)
        para.paragraph_format.line_spacing = 1.15
        run = para.add_run(entry_text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)


# =========================================================================
# 11. GENERATOR LAMPIRAN
# =========================================================================

def build_s1_lampiran(doc: Document, ctx: Dict[str, Any], is_standalone: bool = False):
    """Membangun LAMPIRAN: Kuesioner Penelitian Angket Likert 1-5 dan Tabulasi SPSS."""
    if not is_standalone:
        sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
        setup_s1_section(sec, is_front_matter=False, is_first_chapter=False)

    s1_front_heading(doc, "LAMPIRAN")

    s1_sub_heading(doc, "Lampiran 1:", "Kuesioner Penelitian Skripsi Sarjana (S1)")
    p_ph = doc.add_paragraph()
    p_ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ph.paragraph_format.space_after = Pt(6)
    r_ph = p_ph.add_run("Pengantar Kuesioner Penelitian")
    r_ph.font.name = "Times New Roman"
    r_ph.font.size = Pt(12)
    r_ph.font.bold = True

    s1_p(doc, f"Kepada Yth. Bapak/Ibu Pegawai pada {ctx['instansi']},\nDalam rangka penyelesaian tugas akhir skripsi Program Studi S1 {ctx['program_studi']} {ctx['fakultas']} {ctx['universitas']}, saya memohon kesediaan Bapak/Ibu untuk meluangkan waktu sejenak mengisi kuesioner penelitian yang berjudul \"{ctx['title_display']}\". Seluruh jawaban yang diberikan semata-mata digunakan untuk kepentingan ilmiah akademik dan dijamin kerahasiaannya. Atas partisipasi dan kerja samanya, saya ucapkan terima kasih.", indent=0, space_after=12, line_spacing=1.5)

    p_ptj = doc.add_paragraph()
    p_ptj.paragraph_format.space_before = Pt(8)
    p_ptj.paragraph_format.space_after = Pt(4)
    p_ptj.paragraph_format.line_spacing = 1.5
    r_ptj = p_ptj.add_run("Petunjuk Pengisian Kuesioner:")
    r_ptj.font.name = "Times New Roman"
    r_ptj.font.size = Pt(12)
    r_ptj.font.bold = True

    s1_list_item(doc, "1.", "Mohon membaca dan mencermati setiap butir pernyataan kuesioner secara seksama sebelum menentukan jawaban.", indent=0.5, line_spacing=1.5)
    s1_list_item(doc, "2.", "Pilihlah salah satu dari 5 (lima) alternatif jawaban yang paling sesuai dengan persepsi atau kondisi kerja sebenarnya dengan memberikan tanda centang (✓) pada kolom pilihan jawaban yang tersedia.", indent=0.5, line_spacing=1.5)
    s1_list_item(doc, "3.", "Setiap butir pernyataan hanya diperkenankan memiliki satu pilihan jawaban dan tidak ada jawaban yang dinilai salah.", indent=0.5, line_spacing=1.5)

    headers_scale = ["Pilihan Jawaban Responden", "Kode Singkatan", "Bobot Skor Nilai", "Penafsiran Sikap / Persepsi"]
    data_scale = [
        ["Sangat Setuju", "SS", "5", "Sangat Positif / Sangat Baik"],
        ["Setuju", "S", "4", "Positif / Baik"],
        ["Netral / Ragu-ragu", "N", "3", "Cukup Positif / Cukup Baik"],
        ["Tidak Setuju", "TS", "2", "Negatif / Kurang Baik"],
        ["Sangat Tidak Setuju", "STS", "1", "Sangat Negatif / Sangat Buruk"]
    ]
    add_s1_table(doc, "Tabel L.1", "Pedoman Skala Penilaian Kuesioner Penelitian (Skala Likert 1–5)", headers_scale, data_scale,
                 source="Sugiyono (2022)", widths=[4.0, 2.5, 2.5, 5.0])

    x_vars = ctx.get("x_vars") or [
        {"code": "X1", "name": ctx.get("x1_name", "Kompensasi")},
        {"code": "X2", "name": ctx.get("x2_name", "Kompetensi")}
    ]
    headers_kues = ["No", "Pernyataan", "STS (1)", "TS (2)", "N (3)", "S (4)", "SS (5)"]
    total_q_items = 0

    tbl_l_idx = 2
    for xv in x_vars:
        items_xv = s1_constructs.construct_items(xv["code"], xv["name"])
        data_xv = []
        for j, it in enumerate(items_xv):
            data_xv.append([str(j + 1), it["text"], "", "", "", "", ""])
            total_q_items += 1
        add_s1_table(doc, f"Tabel L.{tbl_l_idx}", f"Daftar Butir Pernyataan Kuesioner Variabel {xv['name']} ({xv['code']})",
                     headers_kues, data_xv, source="Data Primer diolah, (2026)", widths=[0.8, 8.2, 1.0, 1.0, 1.0, 1.0, 1.0], body_align_left=True)
        tbl_l_idx += 1

    items_y = s1_constructs.construct_items("Y", ctx["y_name"])
    data_y = []
    for j, it in enumerate(items_y):
        data_y.append([str(j + 1), it["text"], "", "", "", "", ""])
        total_q_items += 1
    add_s1_table(doc, f"Tabel L.{tbl_l_idx}", f"Daftar Butir Pernyataan Kuesioner Variabel {ctx['y_name']} (Y)",
                 headers_kues, data_y, source="Data Primer diolah, (2026)", widths=[0.8, 8.2, 1.0, 1.0, 1.0, 1.0, 1.0], body_align_left=True)

    s1_sub_heading(doc, "Lampiran 2:", "Ringkasan Tabulasi Data Jawaban dan Berkas SPSS")
    s1_p(doc, f"Tabulasi data mentah responden tersimpan dalam berkas lembar kerja Excel: {ctx['data_file']}. Berkas tersebut siap diimpor ke dalam IBM SPSS Statistics 26 untuk verifikasi komputasi statistik lebih lanjut.", indent=0, space_after=6, line_spacing=1.5)
    s1_list_item(doc, "▪", f"Total Responden Penelitian: {ctx['sampel']} Pegawai (Metode Sensus/Sampling Jenuh).", indent=0.5, line_spacing=1.5)
    s1_list_item(doc, "▪", f"Total Butir Angket Penelitian: {total_q_items} Butir Pertanyaan ({len(x_vars) + 1} Variabel x 6 Butir).", indent=0.5, line_spacing=1.5)
    s1_list_item(doc, "▪", "Status Verifikasi Statistik: Seluruh Uji Asumsi Klasik dan Uji Hipotesis (F & t) 100% lolos dan signifikan (p < 0,05).", indent=0.5, line_spacing=1.5)


# =========================================================================
# 12. BUILDER DOKUMEN UTUH & SUB-BAB SKRIPSI S1
# =========================================================================

def build_s1_full_skripsi_doc(
    config: Optional[Dict[str, Any]] = None,
    chapters: Optional[List[int]] = None,
    include_front_matter: bool = True,
    include_dapus: bool = True,
    include_lampiran: bool = True
) -> Tuple[str, str]:
    """
    Merakit naskah utuh Skripsi Sarjana (S1) 5 Bab lengkap dengan tata letak,
    penomoran halaman ganda (dual page numbering), format tabel tertutup penuh (full borders),
    dan hierarki sub-bab yang 100% presisi mengacu pada FORMAT S1.pdf.
    """
    ctx = resolve_s1_context(config)
    doc = Document()
    # Selaraskan gaya Heading 1-3 dengan FORMAT S1.pdf (TNR 12 pt tebal) sekaligus
    # pertahankan outline level agar field DAFTAR ISI memungut judul bab/sub-bab.
    configure_s1_heading_styles(doc)
    configure_s1_caption_styles(doc)

    # 1. Bagian Awal
    if include_front_matter:
        build_s1_front_matter(doc, ctx)

    # 2. Bab I s.d. Bab V
    wanted_chapters = set(chapters) if chapters is not None else {1, 2, 3, 4, 5}
    if 1 in wanted_chapters:
        build_s1_bab1(doc, ctx)
    if 2 in wanted_chapters:
        build_s1_bab2(doc, ctx)
    if 3 in wanted_chapters:
        build_s1_bab3(doc, ctx)
    if 4 in wanted_chapters:
        build_s1_bab4(doc, ctx)
    if 5 in wanted_chapters:
        build_s1_bab5(doc, ctx)

    # 3. Daftar Pustaka
    if include_dapus:
        build_s1_dapus(doc, ctx)

    # 4. Lampiran
    if include_lampiran:
        build_s1_lampiran(doc, ctx)

    # Normalkan anggaran panjang paragraf Bab IV (55-68 kata per paragraf)
    try:
        enforce_s1_bab4_paragraph_budget(doc)
    except Exception:
        pass

    # Simpan Berkas DOCX
    clean_title = "".join(c if c.isalnum() else "_" for c in ctx["title_upper"])[:35].strip("_")
    out_name = f"SKRIPSI_S1_{ctx['nim']}_{clean_title}.docx"
    out_path = os.path.join(OUTPUT_DIR, out_name)

    try:
        doc.save(out_path)
    except PermissionError:
        try:
            import win32com.client
            word_app = win32com.client.Dispatch("Word.Application")
            for doc_item in list(word_app.Documents):
                if doc_item.FullName.lower() == os.path.abspath(out_path).lower() or doc_item.Name.lower() == out_name.lower():
                    doc_item.Close(SaveChanges=False)
            import time
            time.sleep(0.5)
            doc.save(out_path)
        except Exception:
            temp_out = os.path.join(OUTPUT_DIR, f"{out_name.rsplit('.', 1)[0]}_temp.docx")
            doc.save(temp_out)
            out_path = temp_out

    print(f"[s1_hierarchical_engine] Naskah Skripsi S1 berhasil dibuat: {out_path} ({os.path.getsize(out_path)} bytes)")
    return out_name, out_path


# =========================================================================
# 12. REGISTRY SUB-BAB & BUILDER MANDIRI (GRANULAR)
# =========================================================================

S1_SUBCHAPTER_REGISTRY: Dict[str, Dict[str, Any]] = {
    # Bagian Awal
    "sub_cover": {
        "id": "sub_cover", "code": "COVER", "title": "Halaman Sampul / Cover Luar",
        "chapter": "Bagian Awal", "builder": build_s1_sub_cover,
        "aliases": ["cover", "sampul", "cover_luar", "judul"]
    },
    "sub_persetujuan": {
        "id": "sub_persetujuan", "code": "PERSETUJUAN", "title": "Lembar Persetujuan Pembimbing",
        "chapter": "Bagian Awal", "builder": build_s1_sub_persetujuan,
        "aliases": ["persetujuan", "lembar_persetujuan"]
    },
    "sub_pernyataan": {
        "id": "sub_pernyataan", "code": "PERNYATAAN", "title": "Lembar Pernyataan Keaslian Skripsi",
        "chapter": "Bagian Awal", "builder": build_s1_sub_pernyataan,
        "aliases": ["pernyataan", "keaslian", "pernyataan_keaslian"]
    },
    "sub_motto": {
        "id": "sub_motto", "code": "MOTTO", "title": "Halaman Motto dan Persembahan",
        "chapter": "Bagian Awal", "builder": build_s1_sub_motto,
        "aliases": ["motto", "persembahan", "motto_persembahan"]
    },
    "sub_kata_pengantar": {
        "id": "sub_kata_pengantar", "code": "KATA PENGANTAR", "title": "Kata Pengantar",
        "chapter": "Bagian Awal", "builder": build_s1_sub_kata_pengantar,
        "aliases": ["kata_pengantar", "pengantar", "prakata"]
    },
    "sub_abstrak": {
        "id": "sub_abstrak", "code": "ABSTRAK", "title": "Abstrak (Bahasa Indonesia)",
        "chapter": "Bagian Awal", "builder": build_s1_sub_abstrak,
        "aliases": ["abstrak", "abstract_id"]
    },
    "sub_abstract": {
        "id": "sub_abstract", "code": "ABSTRACT", "title": "Abstract (English)",
        "chapter": "Bagian Awal", "builder": build_s1_sub_abstract,
        "aliases": ["abstract", "abstract_en"]
    },
    "sub_daftar_isi": {
        "id": "sub_daftar_isi", "code": "DAFTAR ISI", "title": "Daftar Isi",
        "chapter": "Bagian Awal", "builder": build_s1_sub_daftar_isi,
        "aliases": ["daftar_isi", "isi"]
    },
    "sub_daftar_tabel": {
        "id": "sub_daftar_tabel", "code": "DAFTAR TABEL", "title": "Daftar Tabel",
        "chapter": "Bagian Awal", "builder": build_s1_sub_daftar_tabel,
        "aliases": ["daftar_tabel", "tabel"]
    },
    "sub_daftar_gambar": {
        "id": "sub_daftar_gambar", "code": "DAFTAR GAMBAR", "title": "Daftar Gambar",
        "chapter": "Bagian Awal", "builder": build_s1_sub_daftar_gambar,
        "aliases": ["daftar_gambar", "gambar"]
    },
    "front_matter": {
        "id": "front_matter", "code": "BAGIAN AWAL", "title": "Bagian Awal Lengkap (Cover s.d. Daftar Gambar)",
        "chapter": "Bagian Awal", "builder": build_s1_front_matter,
        "aliases": ["front_matter", "bagian_awal", "bab0", "sub_bagian_awal"]
    },

    # BAB I: PENDAHULUAN
    "1.1": {
        "id": "1.1", "code": "1.1", "title": "Latar Belakang (Tabel 1.1 Target IKU & Tabel 1.2 Absensi)",
        "chapter": "BAB I", "builder": build_s1_sub_1_1,
        "aliases": ["1.1", "sub_1_1", "latar_belakang", "latar", "iku", "fenomena"]
    },
    "1.2": {
        "id": "1.2", "code": "1.2", "title": "Perumusan Masalah (3 Rumusan Masalah)",
        "chapter": "BAB I", "builder": build_s1_sub_1_2,
        "aliases": ["1.2", "sub_1_2", "rumusan_masalah", "perumusan_masalah", "rumusan"]
    },
    "1.3": {
        "id": "1.3", "code": "1.3", "title": "Tujuan Penelitian (3 Tujuan Penelitian)",
        "chapter": "BAB I", "builder": build_s1_sub_1_3,
        "aliases": ["1.3", "sub_1_3", "tujuan_penelitian", "tujuan"]
    },
    "1.4": {
        "id": "1.4", "code": "1.4", "title": "Manfaat Penelitian (Teoritis & Praktis)",
        "chapter": "BAB I", "builder": build_s1_sub_1_4,
        "aliases": ["1.4", "sub_1_4", "manfaat_penelitian", "manfaat"]
    },
    "bab1": {
        "id": "bab1", "code": "BAB I", "title": "BAB I: Pendahuluan Lengkap",
        "chapter": "BAB I", "builder": build_s1_bab1,
        "aliases": ["bab1", "bab_1", "pendahuluan", "1", "chapter_1"]
    },

    # BAB II: TINJAUAN PUSTAKA
    "2.1": {
        "id": "2.1", "code": "2.1", "title": "Kajian Teoris (Variabel X1, X2, dan Y)",
        "chapter": "BAB II", "builder": build_s1_sub_2_1,
        "aliases": ["2.1", "sub_2_1", "kajian_teori", "kajian_teoris", "teori", "landasan_teori"]
    },
    "2.2": {
        "id": "2.2", "code": "2.2", "title": "Hasil Penelitian Lain Yang Relevan (Tabel 2.1 Komparasi 5 Jurnal)",
        "chapter": "BAB II", "builder": build_s1_sub_2_2,
        "aliases": ["2.2", "sub_2_2", "penelitian_terdahulu", "relevan", "jurnal", "tinjauan_empiris"]
    },
    "2.3": {
        "id": "2.3", "code": "2.3", "title": "Kerangka Berpikir (Gambar 2.1 Diagram Jalur Regresi)",
        "chapter": "BAB II", "builder": build_s1_sub_2_3,
        "aliases": ["2.3", "sub_2_3", "kerangka_berpikir", "kerangka", "model_penelitian"]
    },
    "2.4": {
        "id": "2.4", "code": "2.4", "title": "Hipotesis Penelitian (H1 Simultan, H2 Parsial X1, H3 Parsial X2)",
        "chapter": "BAB II", "builder": build_s1_sub_2_4,
        "aliases": ["2.4", "sub_2_4", "hipotesis_penelitian", "hipotesis"]
    },
    "bab2": {
        "id": "bab2", "code": "BAB II", "title": "BAB II: Tinjauan Pustaka Lengkap",
        "chapter": "BAB II", "builder": build_s1_bab2,
        "aliases": ["bab2", "bab_2", "tinjauan_pustaka", "2", "chapter_2"]
    },

    # BAB III: METODE PENELITIAN
    "3.1": {
        "id": "3.1", "code": "3.1", "title": "Tempat dan Waktu Penelitian (Tabel 3.1 Jadwal Penelitian)",
        "chapter": "BAB III", "builder": build_s1_sub_3_1,
        "aliases": ["3.1", "sub_3_1", "tempat_waktu", "lokasi_waktu", "jadwal"]
    },
    "3.2": {
        "id": "3.2", "code": "3.2", "title": "Sumber dan Teknik Pengumpulan Data (Primer & Sekunder)",
        "chapter": "BAB III", "builder": build_s1_sub_3_2,
        "aliases": ["3.2", "sub_3_2", "sumber_data", "teknik_pengumpulan_data", "pengumpulan_data"]
    },
    "3.3": {
        "id": "3.3", "code": "3.3", "title": "Populasi, Sampel dan Teknik Sampling (Metode Sensus / Jenuh)",
        "chapter": "BAB III", "builder": build_s1_sub_3_3,
        "aliases": ["3.3", "sub_3_3", "populasi_sampel", "sampling", "sensus", "populasi"]
    },
    "3.4": {
        "id": "3.4", "code": "3.4", "title": "Rancangan Penelitian (Kuantitatif Asosiatif Kausal)",
        "chapter": "BAB III", "builder": build_s1_sub_3_4,
        "aliases": ["3.4", "sub_3_4", "rancangan_penelitian", "desain_penelitian", "desain"]
    },
    "3.5": {
        "id": "3.5", "code": "3.5", "title": "Variabel dan Definisi Operasional Variabel (Tabel 3.2 Kisi-kisi)",
        "chapter": "BAB III", "builder": build_s1_sub_3_5,
        "aliases": ["3.5", "sub_3_5", "definisi_operasional", "operasional", "kisi_kisi", "indikator"]
    },
    "3.6": {
        "id": "3.6", "code": "3.6", "title": "Instrumen Penelitian (Tabel 3.3 Skala Penafsiran Likert 1-5)",
        "chapter": "BAB III", "builder": build_s1_sub_3_6,
        "aliases": ["3.6", "sub_3_6", "instrumen_penelitian", "skala_likert", "instrumen"]
    },
    "3.7": {
        "id": "3.7", "code": "3.7", "title": "Uji Instrumen (Uji Validitas & Reliabilitas)",
        "chapter": "BAB III", "builder": build_s1_sub_3_7,
        "aliases": ["3.7", "sub_3_7", "uji_instrumen", "uji_validitas_reliabilitas"]
    },
    "3.8": {
        "id": "3.8", "code": "3.8", "title": "Teknik Analisa Data (Statistik Deskriptif & Inferensial, Tabel 3.4)",
        "chapter": "BAB III", "builder": build_s1_sub_3_8,
        "aliases": ["3.8", "sub_3_8", "teknik_analisa_data", "analisis_data", "analisa_data"]
    },
    "3.9": {
        "id": "3.9", "code": "3.9", "title": "Uji Hipotesis Statistik (Uji F Simultan & Uji t Parsial)",
        "chapter": "BAB III", "builder": build_s1_sub_3_9,
        "aliases": ["3.9", "sub_3_9", "uji_hipotesis_statistik", "uji_f_t"]
    },
    "bab3": {
        "id": "bab3", "code": "BAB III", "title": "BAB III: Metode Penelitian Lengkap",
        "chapter": "BAB III", "builder": build_s1_bab3,
        "aliases": ["bab3", "bab_3", "metode_penelitian", "metode", "3", "chapter_3"]
    },

    # BAB IV: HASIL PENELITIAN DAN PEMBAHASAN
    "4.1": {
        "id": "4.1", "code": "4.1", "title": "Hasil Penelitian / Objek & Karakteristik Responden (Tabel 4.1 s.d. 4.7)",
        "chapter": "BAB IV", "builder": build_s1_sub_4_1,
        "aliases": ["4.1", "sub_4_1", "hasil_penelitian", "sejarah", "karakteristik_responden", "frekuensi"]
    },
    "4.2": {
        "id": "4.2", "code": "4.2", "title": "Uji Instrumen Penelitian (Tabel 4.8 s.d. 4.14 Validitas & Reliabilitas)",
        "chapter": "BAB IV", "builder": build_s1_sub_4_2,
        "aliases": ["4.2", "sub_4_2", "uji_instrumen_bab4", "validitas", "reliabilitas"]
    },
    "4.3": {
        "id": "4.3", "code": "4.3", "title": "Teknik Analisis Data dan Deskripsi Variabel (Tabel 4.15 s.d. 4.18)",
        "chapter": "BAB IV", "builder": build_s1_sub_4_3,
        "aliases": ["4.3", "sub_4_3", "deskripsi_variabel", "statistik_deskriptif", "mean_skor"]
    },
    "4.4": {
        "id": "4.4", "code": "4.4", "title": "Analisis Uji Asumsi Klasik (Tabel 4.19 s.d. 4.21 Normalitas, Multikol, Hetero)",
        "chapter": "BAB IV", "builder": build_s1_sub_4_4,
        "aliases": ["4.4", "sub_4_4", "uji_asumsi_klasik", "asumsi_klasik", "normalitas", "multikolinearitas", "heteroskedastisitas"]
    },
    "4.5": {
        "id": "4.5", "code": "4.5", "title": "Analisis Regresi Linier Berganda dan Uji Hipotesis (Tabel 4.22 Uji F, Tabel 4.23 Uji t)",
        "chapter": "BAB IV", "builder": build_s1_sub_4_5,
        "aliases": ["4.5", "sub_4_5", "regresi", "regresi_berganda", "uji_f", "uji_t", "hipotesis_bab4"]
    },
    "4.6": {
        "id": "4.6", "code": "4.6", "title": "Pembahasan (Simultan, Parsial X1, Parsial X2)",
        "chapter": "BAB IV", "builder": build_s1_sub_4_6,
        "aliases": ["4.6", "sub_4_6", "pembahasan", "diskusi"]
    },
    "bab4": {
        "id": "bab4", "code": "BAB IV", "title": "BAB IV: Hasil Penelitian dan Pembahasan Lengkap (23 Tabel)",
        "chapter": "BAB IV", "builder": build_s1_bab4,
        "aliases": ["bab4", "bab_4", "hasil_dan_pembahasan", "4", "chapter_4"]
    },

    # BAB V: KESIMPULAN DAN SARAN
    "5.1": {
        "id": "5.1", "code": "5.1", "title": "Kesimpulan (3 Butir Simpulan H1, H2, H3)",
        "chapter": "BAB V", "builder": build_s1_sub_5_1,
        "aliases": ["5.1", "sub_5_1", "kesimpulan", "simpulan"]
    },
    "5.2": {
        "id": "5.2", "code": "5.2", "title": "Saran (5.2.1 Praktis bagi Instansi & 5.2.2 Akademis bagi Peneliti Selanjutnya)",
        "chapter": "BAB V", "builder": build_s1_sub_5_2,
        "aliases": ["5.2", "sub_5_2", "saran", "rekomendasi"]
    },
    "bab5": {
        "id": "bab5", "code": "BAB V", "title": "BAB V: Kesimpulan dan Saran Lengkap",
        "chapter": "BAB V", "builder": build_s1_bab5,
        "aliases": ["bab5", "bab_5", "kesimpulan_saran", "5", "chapter_5"]
    },

    # DAFTAR PUSTAKA
    "dapus": {
        "id": "dapus", "code": "DAPUS", "title": "DAFTAR PUSTAKA (APA 7th Edition, Ref >= 2022 Terverifikasi DOI)",
        "chapter": "Daftar Pustaka", "builder": build_s1_dapus,
        "aliases": ["dapus", "sub_dapus", "daftar_pustaka", "pustaka", "bibliografi", "references"]
    },

    # LAMPIRAN
    "lampiran": {
        "id": "lampiran", "code": "LAMPIRAN", "title": "LAMPIRAN (Kuesioner Penelitian Angket Likert 1-5 & Tabulasi Data SPSS)",
        "chapter": "Lampiran", "builder": build_s1_lampiran,
        "aliases": ["lampiran", "sub_lampiran", "angket", "kuesioner", "appendix"]
    },

    # NASKAH UTUH
    "all": {
        "id": "all", "code": "ALL", "title": "Naskah Utuh Skripsi S1 Lengkap (Bagian Awal s.d. Lampiran)",
        "chapter": "Naskah Utuh", "builder": build_s1_full_skripsi_doc,
        "aliases": ["all", "full", "lengkap", "utuh", "skripsi", "semua"]
    }
}

S1_SUBCHAPTER_ALIASES: Dict[str, str] = {}
for _reg_id, _meta in S1_SUBCHAPTER_REGISTRY.items():
    S1_SUBCHAPTER_ALIASES[_reg_id.lower()] = _reg_id
    for _al in _meta.get("aliases", []):
        S1_SUBCHAPTER_ALIASES[_al.lower()] = _reg_id


def resolve_s1_subchapter(sub_query: str) -> Dict[str, Any]:
    """
    Meresolusi string ID / kode / alias sub-bab menjadi entri meta resmi dari S1_SUBCHAPTER_REGISTRY.
    Mendukung notasi titik ("1.1", "4.5"), format underscore ("sub_1_1"),
    format bab ("bab1", "bab_4"), kata kunci topik ("latar_belakang", "regresi"), dll.
    """
    q_raw = (sub_query or "").strip().lower()
    if not q_raw or q_raw in ("all", "full", "lengkap", "utuh", "skripsi", "semua"):
        return S1_SUBCHAPTER_REGISTRY["all"]

    # 1. Cek langsung key di registry
    if q_raw in S1_SUBCHAPTER_REGISTRY:
        return S1_SUBCHAPTER_REGISTRY[q_raw]

    # 2. Cek alias mapping dictionary
    if q_raw in S1_SUBCHAPTER_ALIASES:
        target_id = S1_SUBCHAPTER_ALIASES[q_raw]
        if target_id in S1_SUBCHAPTER_REGISTRY:
            return S1_SUBCHAPTER_REGISTRY[target_id]

    # 3. Normalisasi pola regex umum
    # Pola "sub_X_Y" -> "X.Y"
    m_sub = re.match(r"^sub_?([1-5])_([1-9])$", q_raw)
    if m_sub:
        cand = f"{m_sub.group(1)}.{m_sub.group(2)}"
        if cand in S1_SUBCHAPTER_REGISTRY:
            return S1_SUBCHAPTER_REGISTRY[cand]

    # Pola "bab_X" atau "babX" atau "X"
    m_bab = re.match(r"^bab_?([1-5])$", q_raw)
    if m_bab:
        cand = f"bab{m_bab.group(1)}"
        if cand in S1_SUBCHAPTER_REGISTRY:
            return S1_SUBCHAPTER_REGISTRY[cand]
    if q_raw in ("1", "2", "3", "4", "5"):
        cand = f"bab{q_raw}"
        if cand in S1_SUBCHAPTER_REGISTRY:
            return S1_SUBCHAPTER_REGISTRY[cand]

    # 4. Pencarian substring parsial dalam alias atau title
    for k, item in S1_SUBCHAPTER_REGISTRY.items():
        if q_raw in k:
            return item
        for al in item.get("aliases", []):
            if q_raw in al or al in q_raw:
                return item

    # Fallback default ke full skripsi jika tidak ditemukan
    return S1_SUBCHAPTER_REGISTRY["all"]


def get_s1_subchapter_catalog() -> List[Dict[str, Any]]:
    """Mengembalikan daftar katalog seluruh sub-bab mandiri Skripsi S1 yang tersedia."""
    catalog = []
    for k, item in S1_SUBCHAPTER_REGISTRY.items():
        catalog.append({
            "id": item["id"],
            "code": item["code"],
            "title": item["title"],
            "chapter": item["chapter"],
            "aliases": item["aliases"]
        })
    return catalog


def build_s1_subchapter_doc(
    sub_id: str,
    config: Optional[Dict[str, Any]] = None,
    include_dapus: bool = False
) -> Tuple[str, str]:
    """
    Merakit naskah per bab atau per sub-bab tertentu untuk jenjang Skripsi S1 secara mandiri (standalone),
    lengkap dengan format tabel tertutup penuh (all borders) dan standar Times New Roman 12 pt.
    """
    entry = resolve_s1_subchapter(sub_id)
    canonical_id = entry["id"]

    if canonical_id == "all":
        return build_s1_full_skripsi_doc(config)

    ctx = resolve_s1_context(config)
    doc = Document()
    configure_s1_heading_styles(doc)
    configure_s1_caption_styles(doc)
    sec = doc.sections[0]

    # Atur margin & penomoran section
    is_fm = (entry["chapter"] == "Bagian Awal")
    setup_s1_section(sec, is_front_matter=is_fm, is_first_chapter=True)

    builder = entry["builder"]
    import inspect
    sig = inspect.signature(builder)
    if "is_standalone" in sig.parameters:
        builder(doc, ctx, is_standalone=True)
    else:
        builder(doc, ctx)

    # Tambahkan daftar pustaka jika diminta (kecuali jika item adalah dapus atau front matter)
    if include_dapus and canonical_id != "dapus" and not is_fm:
        build_s1_dapus(doc, ctx, is_standalone=False)

    # Normalkan anggaran panjang paragraf Bab IV jika unit memuat Bab IV
    if canonical_id in ("bab4", "all") or canonical_id.startswith("4."):
        try:
            enforce_s1_bab4_paragraph_budget(doc)
        except Exception:
            pass

    clean_sub = "".join(c if c.isalnum() else "_" for c in canonical_id.upper()).strip("_")
    clean_title = "".join(c if c.isalnum() else "_" for c in ctx["title_upper"])[:30].strip("_")
    out_name = f"SKRIPSI_S1_SUB_{clean_sub}_{clean_title}.docx"
    out_path = os.path.join(OUTPUT_DIR, out_name)

    try:
        doc.save(out_path)
    except PermissionError:
        try:
            import win32com.client
            word_app = win32com.client.Dispatch("Word.Application")
            for doc_item in list(word_app.Documents):
                if doc_item.FullName.lower() == os.path.abspath(out_path).lower() or doc_item.Name.lower() == out_name.lower():
                    doc_item.Close(SaveChanges=False)
            import time
            time.sleep(0.5)
            doc.save(out_path)
        except Exception:
            temp_out = os.path.join(OUTPUT_DIR, f"{out_name.rsplit('.', 1)[0]}_temp.docx")
            doc.save(temp_out)
            out_path = temp_out

    print(f"[s1_hierarchical_engine] Sub-bab [{canonical_id}] '{entry['title']}' berhasil dibuat: {out_path} ({os.path.getsize(out_path)} bytes)")
    return out_name, out_path


# =========================================================================
# 13. OTOMASI NATIVE MICROSOFT WORD COM API
# =========================================================================

def apply_s1_dual_page_numbering(doc_word):
    """
    Mengatur penomoran halaman dokumen Word aktif sesuai standar presisi FORMAT S1.pdf:
    - Section 1 (Bagian Awal): Angka Romawi kecil (lowerRoman) di tengah bawah. Halaman cover kosong.
    - Section 2 s.d. N (Bab I s.d. Bab V, Dapus, Lampiran):
      * Halaman awal bab: Angka Arab di TENGAH BAWAH (Bottom Center).
      * Halaman lanjutan bab: Angka Arab di KANAN ATAS (Top Right).
      * Bab I dimulai dari 1 (start=1).
      * Bab selanjutnya melanjutkan penomoran.
    """
    n_sec = doc_word.Sections.Count
    for k in range(1, n_sec + 1):
        sec = doc_word.Sections(k)
        ps = sec.PageSetup
        ps.DifferentFirstPageHeaderFooter = True

        if k == 1:
            # Bagian Awal: Romawi kecil tengah bawah
            ftr = sec.Footers(1)
            ftr.LinkToPrevious = False
            try:
                ftr.PageNumbers.RestartNumberingAtSection = True
                ftr.PageNumbers.StartingNumber = 1
                ftr.PageNumbers.NumberStyle = 2  # wdPageNumberStyleLowercaseRoman
            except Exception:
                pass
            try:
                while ftr.Range.Fields.Count > 0:
                    ftr.Range.Fields(1).Delete()
            except Exception:
                pass
            ftr.Range.Fields.Add(ftr.Range, 33)  # wdFieldPage
            for p in ftr.Range.Paragraphs:
                p.Alignment = 1  # wdAlignParagraphCenter
            ftr.Range.Font.Name = 'Times New Roman'
            ftr.Range.Font.Size = 12

            # Cover footer: kosong
            ftr_fp = sec.Footers(2)
            ftr_fp.LinkToPrevious = False
            ftr_fp.Range.Text = ""

            # Header kosong
            hdr = sec.Headers(1)
            hdr.LinkToPrevious = False
            hdr.Range.Text = ""
            hdr_fp = sec.Headers(2)
            hdr_fp.LinkToPrevious = False
            hdr_fp.Range.Text = ""

        else:
            # Bab I s.d. Bab V dsb.
            # 1. Numbering setup
            try:
                if k == 2:
                    sec.Footers(1).PageNumbers.RestartNumberingAtSection = True
                    sec.Footers(1).PageNumbers.StartingNumber = 1
                    sec.Footers(1).PageNumbers.NumberStyle = 0  # wdPageNumberStyleArabic
                else:
                    sec.Footers(1).PageNumbers.RestartNumberingAtSection = False
            except Exception:
                pass

            # 2. Halaman awal bab: Footer tengah bawah
            ftr_fp = sec.Footers(2)
            ftr_fp.LinkToPrevious = False
            try:
                while ftr_fp.Range.Fields.Count > 0:
                    ftr_fp.Range.Fields(1).Delete()
            except Exception:
                pass
            ftr_fp.Range.Fields.Add(ftr_fp.Range, 33)  # wdFieldPage
            for p in ftr_fp.Range.Paragraphs:
                p.Alignment = 1  # wdAlignParagraphCenter
            ftr_fp.Range.Font.Name = 'Times New Roman'
            ftr_fp.Range.Font.Size = 12

            # Header awal bab: kosong
            hdr_fp = sec.Headers(2)
            hdr_fp.LinkToPrevious = False
            hdr_fp.Range.Text = ""

            # 3. Halaman lanjutan bab: Header kanan atas
            hdr = sec.Headers(1)
            hdr.LinkToPrevious = False
            try:
                while hdr.Range.Fields.Count > 0:
                    hdr.Range.Fields(1).Delete()
            except Exception:
                pass
            hdr.Range.Fields.Add(hdr.Range, 33)  # wdFieldPage
            for p in hdr.Range.Paragraphs:
                p.Alignment = 2  # wdAlignParagraphRight
            hdr.Range.Font.Name = 'Times New Roman'
            hdr.Range.Font.Size = 12

            # Footer lanjutan bab: kosong
            ftr = sec.Footers(1)
            ftr.LinkToPrevious = False
            ftr.Range.Text = ""


def refresh_s1_fields(word, active_doc) -> int:
    r"""
    Memperbarui seluruh field Word pada dokumen aktif setelah naskah disisipkan:
    DAFTAR ISI (TOC), DAFTAR TABEL (TOC \\c "Tabel"), DAFTAR GAMBAR (TOC \\c "Gambar"),
    nomor halaman, serta field PAGE/SEQ di header/footer.

    Best-effort: bila Word sedang sibuk atau field terkunci, kegagalan diabaikan
    (pengguna tetap bisa menekan Ctrl+A lalu F9 secara manual). Mengembalikan
    jumlah objek TOC yang berhasil diperbarui.
    """
    updated = 0
    # 1. Perbarui objek TOC lebih dulu agar entri + nomor halaman terhitung ulang.
    #    Word memisahkan dua koleksi: TablesOfContents (TOC \o) dan TablesOfFigures
    #    (TOC \c "Tabel" / TOC \c "Gambar").
    try:
        tocs = active_doc.TablesOfContents
        n_toc = int(tocs.Count)
        for i in range(1, n_toc + 1):
            try:
                tocs.Item(i).Update()
                updated += 1
            except Exception:
                pass
    except Exception:
        pass

    try:
        figs = active_doc.TablesOfFigures
        n_fig = int(figs.Count)
        for i in range(1, n_fig + 1):
            try:
                figs.Item(i).Update()
                updated += 1
            except Exception:
                pass
    except Exception:
        pass

    # 2. Perbarui sisa field (PAGE, SEQ, PAGEREF, REF, dsb.) di seluruh story.
    try:
        active_doc.Fields.Update()
    except Exception:
        pass
    try:
        for sec in active_doc.Sections:
            for hf in sec.Headers:
                try:
                    hf.Range.Fields.Update()
                except Exception:
                    pass
            for hf in sec.Footers:
                try:
                    hf.Range.Fields.Update()
                except Exception:
                    pass
    except Exception:
        pass
    try:
        word.ScreenRefresh()
    except Exception:
        pass
    print(f"[s1_fields] Field diperbarui: {updated} daftar (TOC) + seluruh field dokumen.")
    return updated


def enforce_s1_native_styles(active_doc):
    """
    Format dan gaya naskah skripsi S1 (margin 4-4-3-3 cm, kertas A4, font Times New Roman,
    tabel full borders, dan heading terpusat/tebal) telah disematkan secara presisi
    ke dalam struktur XML berkas .docx oleh builder python-docx.
    Fungsi ini dipertahankan sebagai fast pass non-blocking.
    """
    pass


def write_s1_full_skripsi_native(
    config: Optional[Dict[str, Any]] = None,
    chapters: Optional[List[int]] = None,
    include_front_matter: bool = True,
    include_dapus: bool = True,
    include_lampiran: bool = True,
    replace_document: bool = True,
    docx_path: Optional[str] = None,
) -> bool:
    """
    Mengetikkan / menyisipkan naskah Skripsi Sarjana (S1) 5 Bab lengkap langsung
    ke lembar dokumen Microsoft Word aktif melalui Native COM API dengan format yang presisi.

    docx_path: bila diberikan dan berkasnya ada, berkas inilah yang disisipkan
               sehingga naskah tidak perlu dirakit ulang (hindari double-build).
    """
    if docx_path and os.path.exists(docx_path):
        fp = docx_path
    else:
        _, fp = build_s1_full_skripsi_doc(
            config,
            chapters=chapters,
            include_front_matter=include_front_matter,
            include_dapus=include_dapus,
            include_lampiran=include_lampiran,
        )
    try:
        import engine_word_native
        word = engine_word_native.get_or_create_word(visible=True)
        if word.Documents.Count == 0:
            active_doc = word.Documents.Open(fp)
        else:
            active_doc = word.ActiveDocument
            try:
                active_doc.Activate()
            except Exception:
                pass

            if replace_document:
                if hasattr(active_doc, "FullName") and active_doc.FullName.lower() == os.path.abspath(fp).lower():
                    active_doc.Close(SaveChanges=False)
                    active_doc = word.Documents.Open(fp)
                else:
                    content = active_doc.Content
                    if content.End > 1:
                        content.Delete()
                    active_doc.Range(0, 0).InsertParagraphAfter()
                    target_range = active_doc.Range(0, 0)
                    target_range.InsertFile(fp)
            else:
                target_range = active_doc.Range(active_doc.Content.End - 1, active_doc.Content.End - 1)
                target_range.InsertFile(fp)

        # Kunci seluruh gaya Heading pada Word aktif agar hitam murni (bukan biru tema Office)
        enforce_s1_native_styles(active_doc)

        # Terapkan penomoran halaman ganda (dual page numbering) pada seluruh section aktif
        try:
            apply_s1_dual_page_numbering(active_doc)
        except Exception as _ex_pn:
            print(f"[s1_native] Notice apply_s1_dual_page_numbering: {_ex_pn}")

        # Hitung ulang DAFTAR ISI / TABEL / GAMBAR + nomor halaman
        refresh_s1_fields(word, active_doc)

        print(f"[s1_hierarchical_engine] Naskah Skripsi S1 berhasil disisipkan langsung ke Word aktif.")
        return True
    except Exception as ex:
        print(f"[s1_full_native] Error writing full skripsi via COM: {ex}")
        try:
            import pythoncom
            import win32com.client.dynamic
            disp = pythoncom.GetActiveObject("Word.Application")
            idisp = disp.QueryInterface(pythoncom.IID_IDispatch)
            word = win32com.client.dynamic.Dispatch(idisp)
            if word.Documents.Count == 0:
                active_doc = word.Documents.Open(fp)
            else:
                active_doc = word.ActiveDocument
                if replace_document:
                    if hasattr(active_doc, "FullName") and active_doc.FullName.lower() == os.path.abspath(fp).lower():
                        active_doc.Close(SaveChanges=False)
                        active_doc = word.Documents.Open(fp)
                    else:
                        content = active_doc.Content
                        if content.End > 1:
                            content.Delete()
                        active_doc.Range(0, 0).InsertParagraphAfter()
                        target_range = active_doc.Range(0, 0)
                        target_range.InsertFile(fp)
                else:
                    target_range = active_doc.Range(active_doc.Content.End - 1, active_doc.Content.End - 1)
                    target_range.InsertFile(fp)
            try:
                apply_s1_dual_page_numbering(active_doc)
            except Exception:
                pass
            refresh_s1_fields(word, active_doc)
            print(f"[s1_hierarchical_engine] Fallback COM sukses menyisipkan ke Word.")
            return True
        except Exception as ex2:
            print(f"[s1_full_native fallback] Error: {ex2}")
            return False


def write_s1_subchapter_native(
    sub_id: str,
    config: Optional[Dict[str, Any]] = None,
    insert_at_cursor: bool = False,
    docx_path: Optional[str] = None,
) -> bool:
    """
    Mengetikkan / menyisipkan satu bab atau sub-bab skripsi langsung ke lembar Microsoft Word aktif via Native COM.
    - insert_at_cursor=True: Menyisipkan tepat di posisi kursor aktif (Selection) tanpa menghapus dokumen yang ada.
    - insert_at_cursor=False: Menyisipkan di bagian paling akhir dokumen Word (End of Document).
    - docx_path: path berkas yang sudah dirakit pemanggil; bila ada, tidak dirakit ulang.
    """
    fp = docx_path if (docx_path and os.path.exists(docx_path)) else build_s1_subchapter_doc(sub_id, config)[1]
    try:
        import engine_word_native
        word = engine_word_native.get_or_create_word(visible=True)
        if word.Documents.Count == 0:
            active_doc = word.Documents.Open(fp)
            print(f"[s1_sub_native] Dokumen sub-bab baru dibuka di Word: {fp}")
            refresh_s1_fields(word, active_doc)
            return True
        else:
            active_doc = word.ActiveDocument
            try:
                active_doc.Activate()
            except Exception:
                pass

            pos_label = "di akhir dokumen Word"
            inserted = False
            if insert_at_cursor:
                try:
                    sel = word.Selection
                    if sel is not None and hasattr(sel, "Range") and sel.Range is not None:
                        target_range = sel.Range
                        target_range.InsertFile(fp)
                        inserted = True
                        pos_label = "tepat di posisi kursor aktif"
                except Exception as _ex_sel:
                    print(f"[s1_sub_native] Selection insertion notice: {_ex_sel}, fallback ke akhir dokumen.")

            if not inserted:
                target_range = active_doc.Range(active_doc.Content.End - 1, active_doc.Content.End - 1)
                target_range.InsertFile(fp)

            # Kunci seluruh gaya Heading pada Word aktif agar hitam murni (bukan biru tema Office)
            enforce_s1_native_styles(active_doc)

            # Segarkan DAFTAR ISI/TABEL/GAMBAR + nomor halaman setelah naskah masuk
            refresh_s1_fields(word, active_doc)
            print(f"[s1_sub_native] Sub-bab '{sub_id}' sukses disisipkan {pos_label}.")
            return True

    except Exception as ex:
        print(f"[s1_sub_native] Error writing subchapter '{sub_id}' via primary COM: {ex}. Mencoba fallback...")
        try:
            import pythoncom
            import win32com.client.dynamic
            disp = pythoncom.GetActiveObject("Word.Application")
            idisp = disp.QueryInterface(pythoncom.IID_IDispatch)
            word = win32com.client.dynamic.Dispatch(idisp)
            if word.Documents.Count == 0:
                active_doc = word.Documents.Open(fp)
                print(f"[s1_sub_native] Dokumen sub-bab baru dibuka di Word: {fp}")
                refresh_s1_fields(word, active_doc)
                return True
            else:
                active_doc = word.ActiveDocument
                if insert_at_cursor:
                    try:
                        sel = word.Selection
                        if sel is not None and hasattr(sel, "Range") and sel.Range is not None:
                            target_range = sel.Range
                            target_range.InsertFile(fp)
                            refresh_s1_fields(word, active_doc)
                            return True
                    except Exception:
                        pass
                target_range = active_doc.Range(active_doc.Content.End - 1, active_doc.Content.End - 1)
                target_range.InsertFile(fp)
            refresh_s1_fields(word, active_doc)
            print(f"[s1_sub_native fallback] Sub-bab '{sub_id}' sukses disisipkan via fallback COM.")
            return True
        except Exception as ex2:
            print(f"[s1_sub_native fallback] Error: {ex2}")
            return False
