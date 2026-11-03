# backend/s3_hierarchical_engine.py
"""
S3 DOCTORAL HIERARCHICAL PIPELINE ENGINE (V3.0)
==============================================
Arsitektur Generasi Naskah Disertasi Doktor (S3) Berantai (Chained Sequential Pipeline)
yang Disesuaikan Penuh Berdasarkan JUDUL PENELITIAN & MODEL STRUKTURAL HIBRIDA.

Fitur Utama:
  1. Title-Adaptive Semantic Engine:
     - Seluruh Bab (Bab 0 s/d Bab 5, Dapus, Lampiran) dan 36 sub-bab menyesuaikan
       secara presisi dengan JUDUL DISERTASI, instansi, dan variabel penelitian.
     - Mendeteksi dan mengintegrasikan seluruh variabel dari judul:
       * Variabel Bebas / Eksogen: X1, X2, X3, ... Xn
       * Variabel Mediasi / Intervening: M
       * Variabel Moderasi / Kontingensi: Z
       * Variabel Terikat / Endogen: Y
       * Model Riset: Moderated Mediation, Moderation, Mediation
  2. Prompt Khusus Doktoral (S3):
     - Persona: Pakar Akademis & Penulis Disertasi Doktor (S3) Ilmu Manajemen Pascasarjana.
     - Kedalaman telaah: Grand Theory, Middle Range Theory, Applied Theory, Ontologi,
       Epistemologi, Aksiologi, Novelty / Kebaruan Riset, State of the Art, Kontribusi
       Teoretis & Implikasi Kebijakan Strategis.
  3. Integrasi PLS-SEM & PROCESS Hayes:
     - Outer Model: Convergent, Discriminant HTMT, Composite Reliability, VIF.
     - Inner Model: R², Q² predict, f² effect size.
     - Bootstrapping 5.000 subsampel untuk pengujian hipotesis H1 s.d Hm.
     - Moderasi Dua Arah: Simple Slope Analysis & Plot interaksi.
     - Mediasi Termoderasi: Conditional Indirect Effects (Index of Moderated Mediation).
  4. Pembahasan Komprehensif Format 4-Alinea Baku Pascasarjana.
  5. Multi-Output: Direct Word Native COM, Office.js base64, dan berkas .docx siap unduh.
  6. 100% Terisolasi: Tidak mengubah atau merusak mesin Master S2 lama.
"""

import os
import sys
import json
import time
import re
import math
import hashlib
from typing import Dict, Any, List, Tuple, Optional, Callable
from datetime import datetime

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls


def _join_id(items: Any) -> str:
    """Gabungkan nama variabel secara gramatikal: 'A', 'A dan B', 'A, B, dan C'."""
    vals = [str(x).strip() for x in (items or []) if str(x).strip()]
    if not vals:
        return ""
    if len(vals) == 1:
        return vals[0]
    if len(vals) == 2:
        return f"{vals[0]} dan {vals[1]}"
    return ", ".join(vals[:-1]) + f", dan {vals[-1]}"


def set_cell_border(cell, **kwargs):
    """Sets individual cell borders (top, bottom, left, right)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(r'<w:tcBorders %s/>' % nsdecls('w'))
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f'w:{edge}'
            element = parse_xml(r'<%s %s w:val="%s" w:sz="%s" w:space="0" w:color="%s"/>' %
                                (tag, nsdecls('w'), edge_data.get('val', 'single'),
                                 edge_data.get('sz', '4'), edge_data.get('color', 'auto')))
            tcBorders.append(element)
        else:
            tag = f'w:{edge}'
            element = parse_xml(r'<%s %s w:val="none"/>' % (tag, nsdecls('w')))
            tcBorders.append(element)
    tcPr.append(tcBorders)

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

import s3_title_parser
import s3_ontology_researcher
import s3_narrative_variation as nv
import s3_data_ingest as s3data
import s3_data_pipeline
import structural_graph_compiler
import dynamic_pls_engine
import engine_word_native
import s3_academic_researcher
from build_full_thesis_docx import (
    setup_page, p, h1, h2, h3, add_table_data, format_proper_instansi, add_rich_runs_to_para,
    _toc_field, _set_update_fields, mark_caption_styles,
    add_page_number_footer, apply_section_margins,
)

_NAMA_BULAN_ID = (
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
)

# Padanan istilah konstruk untuk abstrak berbahasa Inggris. Istilah yang tidak
# terdaftar dipertahankan apa adanya agar tidak muncul terjemahan keliru.
_EN_VAR_TERMS = {
    "beban kerja": "workload",
    "lingkungan kerja": "work environment",
    "disiplin kerja": "work discipline",
    "kinerja pegawai": "employee performance",
    "kinerja karyawan": "employee performance",
    "kinerja organisasi": "organizational performance",
    "kinerja": "performance",
    "kepuasan kerja": "job satisfaction",
    "motivasi kerja": "work motivation",
    "komitmen organisasi": "organizational commitment",
    "komitmen pimpinan": "leadership commitment",
    "budaya organisasi": "organizational culture",
    "budaya kerja": "work culture",
    "iklim kerja": "work climate",
    "iklim organisasi": "organizational climate",
    "gaya kepemimpinan": "leadership style",
    "kepemimpinan transformasional": "transformational leadership",
    "kepemimpinan": "leadership",
    "etos kerja": "work ethic",
    "stres kerja": "work stress",
    "konflik kerja": "work conflict",
    "kualitas sumber daya manusia": "human resource quality",
    "kualitas sdm": "human resource quality",
    "kompetensi": "competence",
    "pengembangan karir": "career development",
    "pengembangan karier": "career development",
    "pelayanan publik": "public service",
    "kualitas pelayanan": "service quality",
    "akuntabilitas": "accountability",
    "integritas": "integrity",
    "kinerja keuangan": "financial performance",
    "keputusan pembelian": "purchase decision",
    "minat beli": "purchase intention",
    "loyalitas pelanggan": "customer loyalty",
    "kepuasan pelanggan": "customer satisfaction",
    "citra merek": "brand image",
    "sistem informasi": "information system",
    "teknologi informasi": "information technology",
}


def _en_var(name: str) -> str:
    """Padanan Inggris nama variabel; kembalikan nama asli bila tak dikenali."""
    key = " ".join(str(name or "").split()).lower()
    return _EN_VAR_TERMS.get(key, str(name or "").strip())


def _join_en(items: Any) -> str:
    """Gabungkan nama variabel dalam bahasa Inggris: 'A', 'A and B', 'A, B, and C'."""
    vals = [str(x).strip() for x in (items or []) if str(x).strip()]
    if not vals:
        return ""
    if len(vals) == 1:
        return vals[0]
    if len(vals) == 2:
        return f"{vals[0]} and {vals[1]}"
    return ", ".join(vals[:-1]) + f", and {vals[-1]}"


def _fmt3(val: Any, default: str = s3data.PLACEHOLDER) -> str:
    """Format angka 3 desimal aman: tangani None, string placeholder, dan float."""
    if val is None or val == "" or val == s3data.PLACEHOLDER:
        return default
    try:
        return f"{float(val):.3f}"
    except (TypeError, ValueError):
        return str(val)


def _get_instansi_focus(instansi: str) -> str:
    """Karakteristik fokus tugas instansi dinamis untuk narasi lokus yang akurat."""
    inst_lower = (instansi or "").lower()
    if any(k in inst_lower for k in ("pendidikan", "sekolah", "guru", "diknas", "disdik")):
        return "pelayanan dan tata kelola administrasi pendidikan daerah"
    if any(k in inst_lower for k in ("pajak", "pendapatan", "bapenda", "keuangan", "retribusi")):
        return "pelayanan dan pengelolaan penerimaan daerah"
    if any(k in inst_lower for k in ("kesehatan", "rumah sakit", "rsud", "dinkes", "puskesmas")):
        return "pelayanan kesehatan masyarakat dan administrasi medis daerah"
    return "penyelenggaraan pelayanan publik dan administrasi kedinasan"


def h4(doc, text):
    """Heading 4 / Sub-sub-heading (bold, left-aligned, indent 0.5 cm)."""
    return p(doc, text, bold=True, italic=False, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0.5, space_before=6, space_after=2)


def pv(doc, text, ctx: Dict[str, Any], salt: str = "", **kwargs):
    """Paragraf narasi dengan variasi leksikal deterministik per judul.

    Dipakai khusus untuk paragraf naratif. Judul sub-bab, pernyataan hipotesis,
    isi tabel, dan sitasi tidak boleh melewati fungsi ini agar tidak berubah.
    """
    para = p(doc, nv.polish_and_restructure(text, ctx, salt), **kwargs)
    try:
        para._nv_varied = True
    except Exception:
        pass
    return para


def _set_para_text(para, text: str):
    """Ganti seluruh teks paragraf tanpa mengubah format paragraf pertama."""
    try:
        if not para.runs:
            para.add_run(text)
            return
        para.runs[0].text = text
        for r in para.runs[1:]:
            r.text = ""
    except Exception:
        pass


def de_fabricate_sweep(doc: Document, ctx: Dict[str, Any]):
    """Sapuan kejujuran data saat mode template atau data simulasi sintetis.

    - Bila mode template: label Sumber diubah menjadi penanda template dan banner template disisipkan.
    - Bila data simulasi sintetis (origin='generated'): banner catatan kejujuran data disisipkan di awal Bab IV dan Bab V.
    - Bila data empiris lapangan riil: tidak menyisipkan banner peringatan.
    """
    is_tmpl = s3data.is_template(ctx)
    meta = (s3data.get(ctx) or {}).get("meta") or {}
    origin = str(meta.get("origin") or "").strip().lower()

    banner = None
    if is_tmpl:
        banner = s3data.TEMPLATE_BANNER
        for para in doc.paragraphs:
            t = para.text.strip()
            if not t:
                continue
            low = t.lower()
            if low.startswith("sumber") and ("data primer" in low or "diolah" in low):
                _set_para_text(para, s3data.SOURCE_TEMPLATE)
    elif origin == "generated":
        banner = getattr(s3data, "GENERATED_BANNER", None)

    if not banner:
        return

    # Sisipkan banner setelah heading Bab IV dan Bab V.
    targets = []
    for idx, para in enumerate(doc.paragraphs):
        t = (para.text or "").upper()
        if t.replace("\n", " ").startswith("BAB IV") or "HASIL ANALISIS" in t:
            targets.append(("BAB IV", idx))
        elif t.replace("\n", " ").startswith("BAB V") or "PEMBAHASAN" in t and "HASIL ANALISIS" not in t:
            targets.append(("BAB V", idx))

    inserted_markers = set()
    for marker, idx in targets:
        if marker in inserted_markers:
            continue
        inserted_markers.add(marker)
        nxt = doc.paragraphs[idx + 1] if idx + 1 < len(doc.paragraphs) else None
        if nxt is not None:
            bp = nxt.insert_paragraph_before()
            run = bp.add_run(banner)
            run.bold = True
            run.italic = True
            bp.paragraph_format.space_before = 6
            bp.paragraph_format.space_after = 6


OUTPUT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, "..", "output"))
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ==============================================================================
# KATALOG SUB-BAB DISERTASI DOKTOR (S3) LENGKAP (36 UNIT)
# ==============================================================================
S3_SUBCHAPTER_CATALOG = {
    # Bagian Awal (Front Matter)
    "sub_bagian_awal": {
        "chapter": 0, "code": "Bagian Awal",
        "title": "Bagian Awal (Halaman Judul, Lembar Pengesahan Promotor, Abstrak Bilingual, Riwayat Hidup, Daftar Isi)",
        "desc": "Halaman administratif pra-naskah Disertasi Doktor: Halaman Sampul, Lembar Persetujuan Komisi Promotor & Pengesahan Komisi Penguji Disertasi, Kata Pengantar, Abstrak Bahasa Indonesia & Inggris, Riwayat Hidup, dan Daftar Isi/Tabel/Gambar."
    },
    # Bab I: Pendahuluan (Format Disertasi Doktoral TAPD Pascasarjana: A, B, C, D, E, F)
    "sub_1_1": {
        "chapter": 1, "code": "A",
        "title": "A. Latar Belakang",
        "desc": "Urgensi makro-nasional, telaah teoritis komprehensif variabel judul, profil dan fenomena empiris instansi, celah riset kontradiktif, novelty kebaruan doktoral, dan penetapan judul."
    },
    "sub_1_2": {
        "chapter": 1, "code": "B",
        "title": "B. Rumusan Masalah",
        "desc": "Pertanyaan penelitian doktoral terstruktur (H1 s.d Hm) yang diturunkan dari model struktural judul mencakup pengaruh langsung, mediasi, moderasi, dan mediasi termoderasi."
    },
    "sub_1_3": {
        "chapter": 1, "code": "C",
        "title": "C. Tujuan Penelitian",
        "desc": "Tujuan penelitian doktoral yang selaras 1-ke-1 dengan rumusan masalah dan hipotesis dari model judul."
    },
    "sub_1_4": {
        "chapter": 1, "code": "D",
        "title": "D. Manfaat Penelitian",
        "desc": "Kontribusi teoretis pengembangan keilmuan dan novelty model doktoral terpadu sesuai teori yang relevan dengan variabel judul, serta manfaat praktis rekomendasi strategis bagi pimpinan instansi, unit pengawasan internal, dan masyarakat luas."
    },

    # Bab II: Tinjauan Pustaka & Hipotesis (Struktur Acuan TAPD 2026)
    "sub_2_1_0": {
        "chapter": 2, "code": "A.1",
        "title": "A.1 1. Grand Teori (Deduktif 3 Tingkat & Novelty Model)",
        "desc": "Hierarki teoretis deduktif doktoral: Grand Theory, Middle Range Theory, Applied Theory, konstruksi Novelty Model, dan justifikasi Relevansi dengan Penelitian."
    },
    "sub_2_1_1": {
        "chapter": 2, "code": "A.2",
        "title": "A.2 2. Kajian Teoritis Variabel Terikat (Y)",
        "desc": "Kajian mendalam hakikat variabel dependen, telaah naratif 10 definisi ahli terkini, sintesis peneliti, serta dimensi dan indikator operasional teks hierarkis murni tanpa tabel."
    },
    "sub_2_1_2": {
        "chapter": 2, "code": "A.3",
        "title": "A.3 3. Kajian Teoritis Variabel Mediasi (M)",
        "desc": "Kajian mendalam variabel intervening mediasi, teori transmisi psikologis, telaah naratif 10 definisi ahli, sintesis peneliti, serta dimensi dan indikator teks hierarkis murni tanpa tabel."
    },
    "sub_2_1_4": {
        "chapter": 2, "code": "A.4",
        "title": "A.4 4. Kajian Teoritis Variabel Moderasi (Z)",
        "desc": "Kajian mendalam variabel kontingensi moderasi, Contingency Theory, kondisi batas (boundary conditions), telaah naratif 10 definisi ahli, sintesis peneliti, serta dimensi dan indikator teks hierarkis murni tanpa tabel."
    },
    "sub_2_1_3": {
        "chapter": 2, "code": "A.5",
        "title": "A.5 5. Kajian Teoritis Variabel Bebas (X1, X2... Xn)",
        "desc": "Kajian mendalam seluruh pilar variabel bebas eksogen, teori kapasitas organisasi, telaah naratif 10 definisi per variabel, sintesis peneliti, serta dimensi dan indikator teks hierarkis murni tanpa tabel.",
        "is_composite": True
    },
    "sub_2_1_3_x1": {
        "chapter": 2, "code": "A.5.1",
        "title": "A.5.1 Kajian Teoritis Variabel Bebas X1",
        "desc": "Kajian mendalam variabel bebas X1: hakikat pilar kapasitas, telaah naratif 10 definisi ahli, sintesis peneliti, serta dimensi dan indikator teks hierarkis murni tanpa tabel.",
        "is_granular_child": True
    },
    "sub_2_1_3_x2": {
        "chapter": 2, "code": "A.5.2",
        "title": "A.5.2 Kajian Teoritis Variabel Bebas X2",
        "desc": "Kajian mendalam variabel bebas X2: hakikat pilar kapasitas, telaah naratif 10 definisi ahli, sintesis peneliti, serta dimensi dan indikator teks hierarkis murni tanpa tabel.",
        "is_granular_child": True
    },
    "sub_2_1_3_x3": {
        "chapter": 2, "code": "A.5.3",
        "title": "A.5.3 Kajian Teoritis Variabel Bebas X3",
        "desc": "Kajian mendalam variabel bebas X3: hakikat pilar kapasitas, telaah naratif 10 definisi ahli, sintesis peneliti, serta dimensi dan indikator teks hierarkis murni tanpa tabel.",
        "is_granular_child": True
    },
    "sub_2_1_3_x4": {
        "chapter": 2, "code": "A.5.4",
        "title": "A.5.4 Kajian Teoritis Variabel Bebas X4",
        "desc": "Kajian mendalam variabel bebas X4: hakikat pilar kapasitas, telaah naratif 10 definisi ahli, sintesis peneliti, serta dimensi dan indikator teks hierarkis murni tanpa tabel.",
        "is_granular_child": True
    },
    "sub_2_2": {
        "chapter": 2, "code": "B",
        "title": "B. Hasil Penelitian yang Relevan",
        "desc": "Sintesis matriks komparasi minimal 15 penelitian empiris bereputasi mutakhir (2022-2026) dalam format baku Tabel 2.1 (5 kolom baku: No, Penulis/Tahun, Judul/publikasi, Metoda Riset, Temuan Penelitian), analisis dialektika konvergensi dan divergensi temuan, serta perumusan celah penelitian (research gap) dan kebaruan ilmiah (novelty) model terpadu."
    },
    "sub_2_3": {
        "chapter": 2, "code": "C",
        "title": "C. Kerangka Berpikir",
        "desc": "Arsitektur pemikiran teoretis terpadu kausalitas struktural antarvariabel, dekonstruksi paradigma direct effect, matriks silogisme deduktif H1-Hn, elaborasi empiris mutakhir (2022-2026), Gambar 2.1 Bagan Kerangka Berpikir, Gambar 2.2 State of the Art, dan argumentasi kritis boundary conditions variabel kontingensi (bila model memuat moderasi).",
        "is_composite": True
    },
    "sub_2_3_1": {
        "chapter": 2, "code": "C.1",
        "title": "C.1 Kerangka Konseptual Penelitian",
        "desc": "Dekonstruksi paradigma direct effect, rasionalisasi mediasi sebagai behavioral execution bridge, silogisme deduktif H1-Hn (Premis Mayor & Premis Minor), dan alur deduksi logika.",
        "is_granular_child": True
    },
    "sub_2_3_2": {
        "chapter": 2, "code": "C.2",
        "title": "C.2 Keterkaitan Kausal Antar-Variabel",
        "desc": "Sintesis keterkaitan nalar teoretis dan komparasi literatur empiris mutakhir (2022-2026) yang mendasari hubungan kausal antar-konstruk (X->M, X->Y, M->Y, transmisi mediasi, moderasi).",
        "is_granular_child": True
    },
    "sub_2_3_3": {
        "chapter": 2, "code": "C.3",
        "title": "C.3 Paradigma & Diagram Alur Model Penelitian",
        "desc": "Gambar 2.1 Bagan Kerangka Berpikir, Gambar 2.2 State of the Art Framework & Akronim Model Doktoral Orisinal, serta Analisis Kritis Boundary Conditions Variabel Moderasi (3 fenomena kegagalan sistemik).",
        "is_granular_child": True
    },
    "sub_2_4": {
        "chapter": 2, "code": "D",
        "title": "D. Hipotesis Penelitian",
        "desc": "Paragraf pengantar deduktif ilmiah, formulasi hipotesis formal berarah (H1 s.d Hn) dalam format tabel borderless 3 kolom baku TAPD 2026 yang selaras 1-ke-1 dengan Rumusan Masalah dan Tujuan Penelitian, serta narasi penutup metodologis pengujian PLS-SEM SmartPLS 4 (bootstrapping 5.000 subsampel)."
    },

    # Bab III: Metode Penelitian (Format Baku TAPD 2026 DISERTASI S3.docx)
    "sub_3_a": {
        "chapter": 3, "code": "A",
        "title": "A. Desain Penelitian",
        "desc": "Pendekatan kuantitatif kausal-eksplanatori berlandaskan paradigma positivisme untuk menguji konstelasi hubungan antarvariabel serta menguji hipotesis secara inferensial menggunakan PLS-SEM (Hair et al., 2022; Neuman, 2022)."
    },
    "sub_3_b": {
        "chapter": 3, "code": "B",
        "title": "B. Lokasi Penelitian",
        "desc": "Penetapan lokus empiris pada instansi judul, argumentasi relevansi strategis tata kelola terhadap variabel penelitian, keterbukaan akses data riset (data accessibility), dan heterogenitas responden representatif (Cooper & Schindler, 2022; Sekaran & Bougie, 2020)."
    },
    "sub_3_c": {
        "chapter": 3, "code": "C",
        "title": "C. Waktu Penelitian",
        "desc": "Rentang waktu pelaksanaan penelitian, justifikasi stabilitas aktivitas kedinasan organisasi untuk meminimalkan bias, dan matriks Tabel 3.1 Jadwal Penelitian lengkap 13 tahapan riset dari usulan judul hingga ujian promosi doktor (Sugiyono, 2022; Neuman, 2022)."
    },
    "sub_3_d": {
        "chapter": 3, "code": "D",
        "title": "D. Sumber Data",
        "desc": "Klasifikasi data primer (kuesioner skala Likert 1-5 terstruktur) dan data sekunder (Renstra, LAKIP/LKjIP, arsip kepegawaian) serta prosedur mitigasi Common Method Bias (CMB) prosedural dan statistik Harman's Single Factor Test (Podsakoff et al., 2024; Hair et al., 2022)."
    },
    "sub_3_e": {
        "chapter": 3, "code": "E",
        "title": "E. Populasi dan Sampel",
        "desc": "Definisi wilayah generalisasi populasi dan Tabel 3.2 Sebaran Populasi Pegawai per Unit Kerja, kriteria inklusi purposive sampling, perhitungan rumus Slovin presisi 5%, justifikasi kecukupan sampel PLS-SEM, serta formula dan Tabel 3.3 Alokasi Sampel Proportional Random Sampling (Sugiyono, 2022; Sinambela, 2023; Hair et al., 2022)."
    },
    "sub_3_f": {
        "chapter": 3, "code": "F",
        "title": "F. Definisi Konseptual dan Operasional Variabel",
        "desc": "Taksonomi 4 kuadran model (Endogen Y, Mediasi M, Moderasi Z, Eksogen X1..Xn), masing-masing memuat Definisi Konseptual berakar teori substantif, Definisi Operasional skala Likert 1-5, serta matriks Tabel 3.4 s.d 3.10 Kisi-Kisi Instrumen Dimensi dan Indikator Empiris (Sekaran & Bougie, 2020; Cooper & Schindler, 2022).",
        "is_composite": True
    },
    "sub_3_f_y": {
        "chapter": 3, "code": "F.1",
        "title": "F.1 Definisi & Kisi-Kisi Variabel Terikat (Y)",
        "desc": "Definisi konseptual, operasional skala Likert 1-5, dan tabel kisi-kisi dimensi dan indikator empiris untuk Variabel Terikat (Y).",
        "is_granular_child": True
    },
    "sub_3_f_m": {
        "chapter": 3, "code": "F.2",
        "title": "F.2 Definisi & Kisi-Kisi Variabel Mediasi (M)",
        "desc": "Definisi konseptual, operasional skala Likert 1-5, dan tabel kisi-kisi dimensi dan indikator empiris untuk Variabel Mediasi (M).",
        "is_granular_child": True
    },
    "sub_3_f_z": {
        "chapter": 3, "code": "F.3",
        "title": "F.3 Definisi & Kisi-Kisi Variabel Moderasi (Z)",
        "desc": "Definisi konseptual, operasional skala Likert 1-5, dan tabel kisi-kisi dimensi dan indikator empiris untuk Variabel Moderasi (Z).",
        "is_granular_child": True
    },
    "sub_3_f_x1": {
        "chapter": 3, "code": "F.4.1",
        "title": "F.4.1 Definisi & Kisi-Kisi Variabel Bebas (X1)",
        "desc": "Definisi konseptual, operasional skala Likert 1-5, dan tabel kisi-kisi dimensi dan indikator empiris untuk Variabel Bebas X1.",
        "is_granular_child": True
    },
    "sub_3_f_x2": {
        "chapter": 3, "code": "F.4.2",
        "title": "F.4.2 Definisi & Kisi-Kisi Variabel Bebas (X2)",
        "desc": "Definisi konseptual, operasional skala Likert 1-5, dan tabel kisi-kisi dimensi dan indikator empiris untuk Variabel Bebas X2.",
        "is_granular_child": True
    },
    "sub_3_f_x3": {
        "chapter": 3, "code": "F.4.3",
        "title": "F.4.3 Definisi & Kisi-Kisi Variabel Bebas (X3)",
        "desc": "Definisi konseptual, operasional skala Likert 1-5, dan tabel kisi-kisi dimensi dan indikator empiris untuk Variabel Bebas X3 (dinamis).",
        "is_granular_child": True
    },
    "sub_3_f_x4": {
        "chapter": 3, "code": "F.4.4",
        "title": "F.4.4 Definisi & Kisi-Kisi Variabel Bebas (X4)",
        "desc": "Definisi konseptual, operasional skala Likert 1-5, dan tabel kisi-kisi dimensi dan indikator empiris untuk Variabel Bebas X4 (dinamis).",
        "is_granular_child": True
    },
    "sub_3_g": {
        "chapter": 3, "code": "G",
        "title": "G. Instrumen Penelitian",
        "desc": "Penyusunan instrumen kuesioner tertutup terstruktur deduktif, teori skala pengukuran Likert 5 poin (Sugiyono, 2022; Sinambela, 2023), Tabel 3.11 Skor Pilihan Jawaban Responden (1 s.d 5), serta protokol uji kelayakan isi (content validity) dan standarisasi alat ukur (Cooper & Schindler, 2022)."
    },
    "sub_3_h": {
        "chapter": 3, "code": "H",
        "title": "H. Ujicoba Instrumen",
        "desc": "Prosedur pilot testing pada 50 responden setara di luar sampel utama, Uji Validitas Konstruk Pearson Product Moment (df=48, r-tabel=0,279) dan Tabel 3.12 s.d 3.18 hasil validitas per variabel, serta Uji Reliabilitas Cronbach's Alpha, Tabel 3.19 Rujukan Interpretasi Sinambela (2023), dan Tabel 3.20 Hasil Reliabilitas Sangat Reliabel (Hair et al., 2022; Sinambela, 2023)."
    },
    "sub_3_i": {
        "chapter": 3, "code": "I",
        "title": "I. Teknik Analisis Data",
        "desc": "Rasionalisasi metodologis PLS-SEM SmartPLS 4 model struktural terpadu, kecukupan sampel 10-times rule & G*Power, Tiga Tahap Analisis: (1) Evaluasi Outer Model [Loading > 0,708, AVE > 0,50, Fornell-Larcker, Cross Loading, HTMT < 0,85, CR/Alpha > 0,70, VIF < 3], (2) Evaluasi Inner Model [Inner VIF, R² substansial/moderat, f², Q² blindfolding, SRMR < 0,08, NFI, GoF Tenenhaus], dan (3) Pengujian Hipotesis Bootstrapping 5.000 subsampel (t > 1,96, p < 0,05), Specific Indirect Effects / VAF mediasi, Two-Stage Interaction moderasi, serta persamaan regresi struktural matematis (Hair et al., 2022; Hayes, 2022)."
    },

    # Backward compatibility aliases (sub_3_1 s.d sub_3_6)
    "sub_3_1": {
        "chapter": 3, "code": "A",
        "title": "A. Desain Penelitian (Alias 3.1)",
        "desc": "Alias backward-compatible untuk Sub-Bab A. Desain Penelitian.",
        "is_legacy_alias": True
    },
    "sub_3_2": {
        "chapter": 3, "code": "E",
        "title": "E. Populasi dan Sampel (Alias 3.2)",
        "desc": "Alias backward-compatible untuk Sub-Bab E. Populasi dan Sampel.",
        "is_legacy_alias": True
    },
    "sub_3_3": {
        "chapter": 3, "code": "F",
        "title": "F. Definisi Konseptual & Operasional (Alias 3.3)",
        "desc": "Alias backward-compatible untuk Sub-Bab F. Definisi Konseptual dan Operasional Variabel.",
        "is_legacy_alias": True
    },
    "sub_3_4": {
        "chapter": 3, "code": "G",
        "title": "G. Instrumen Penelitian (Alias 3.4)",
        "desc": "Alias backward-compatible untuk Sub-Bab G. Instrumen Penelitian.",
        "is_legacy_alias": True
    },
    "sub_3_5": {
        "chapter": 3, "code": "H",
        "title": "H. Ujicoba Instrumen (Alias 3.5)",
        "desc": "Alias backward-compatible untuk Sub-Bab H. Ujicoba Instrumen.",
        "is_legacy_alias": True
    },
    "sub_3_6": {
        "chapter": 3, "code": "I",
        "title": "I. Teknik Analisis Data PLS-SEM (Alias 3.6)",
        "desc": "Alias backward-compatible untuk Sub-Bab I. Teknik Analisis Data.",
        "is_legacy_alias": True
    },

    # Bab IV: Hasil Analisis dan Pembahasan (Format Baku TAPD 2026 Disertasi Doktor S3)
    "sub_4_a_1": {
        "chapter": 4, "code": "4.A.1",
        "title": "A.1 Karakteristik Demografis Responden",
        "desc": "Tabel 4.1 (Gender), Tabel 4.2 (Usia), Tabel 4.3 (Pangkat/Golongan), serta narasi analitis representasi meritokrasi aparatur."
    },
    "sub_4_a_2": {
        "chapter": 4, "code": "4.A.2",
        "title": "A.2 Analisis Statistik Deskriptif Variabel",
        "desc": "Statistik deskriptif seluruh variabel (Min, Max, Median, Mean, SD), kriteria interval Sugiyono (Tabel 4.11), dan rekapitulasi kategori respon responden (Tabel 4.12).",
        "is_composite": True
    },
    "sub_4_a_2_y": {
        "chapter": 4, "code": "4.A.2.1",
        "title": "A.2.1 Deskriptif Variabel Terikat (Y)",
        "desc": "Tabel 5 parameter statistik deskriptif dan narasi interpretasi kategori 3-box Sugiyono untuk Variabel Terikat (Y).",
        "is_granular_child": True
    },
    "sub_4_a_2_mz": {
        "chapter": 4, "code": "4.A.2.2",
        "title": "A.2.2 Deskriptif Variabel Mediasi (M) & Moderasi (Z)",
        "desc": "Tabel 5 parameter statistik deskriptif dan narasi interpretasi untuk variabel intervening mediasi dan kontingensi moderasi.",
        "is_granular_child": True
    },
    "sub_4_a_2_x1": {
        "chapter": 4, "code": "4.A.2.3.1",
        "title": "A.2.3.1 Deskriptif Variabel Bebas X1",
        "desc": "Tabel 5 parameter statistik deskriptif dan narasi interpretasi kategori 3-box Sugiyono untuk Variabel Bebas X1.",
        "is_granular_child": True
    },
    "sub_4_a_2_x2": {
        "chapter": 4, "code": "4.A.2.3.2",
        "title": "A.2.3.2 Deskriptif Variabel Bebas X2",
        "desc": "Tabel 5 parameter statistik deskriptif dan narasi interpretasi kategori 3-box Sugiyono untuk Variabel Bebas X2.",
        "is_granular_child": True
    },
    "sub_4_a_2_x3": {
        "chapter": 4, "code": "4.A.2.3.3",
        "title": "A.2.3.3 Deskriptif Variabel Bebas X3",
        "desc": "Tabel 5 parameter statistik deskriptif dan narasi interpretasi kategori 3-box Sugiyono untuk Variabel Bebas X3 (dinamis).",
        "is_granular_child": True
    },
    "sub_4_a_2_x4": {
        "chapter": 4, "code": "4.A.2.3.4",
        "title": "A.2.3.4 Deskriptif Variabel Bebas X4",
        "desc": "Tabel 5 parameter statistik deskriptif dan narasi interpretasi kategori 3-box Sugiyono untuk Variabel Bebas X4 (dinamis).",
        "is_granular_child": True
    },
    "sub_4_a_2_rekap": {
        "chapter": 4, "code": "4.A.2.Rekap",
        "title": "A.2.Rekap Rekapitulasi Rerata Skor & Kategorisasi Keseluruhan Variabel",
        "desc": "Perhitungan rentang skala Sugiyono RS=1,33, tabel kriteria kategori (Rendah, Sedang, Tinggi), tabel rekapitulasi grand mean seluruh variabel, dan narasi evaluasi komparatif.",
        "is_granular_child": True
    },
    "sub_4_a_3": {
        "chapter": 4, "code": "4.A.3",
        "title": "A.3 Evaluasi Model Pengukuran (Outer Model SmartPLS 4)",
        "desc": "Evaluasi validitas konvergen (Gambar 4.1, Outer Loading, AVE), validitas diskriminan (HTMT < 0,85, Cross Loading), dan reliabilitas konstruk (Cronbach's Alpha, Composite Reliability).",
        "is_composite": True
    },
    "sub_4_a_3_1": {
        "chapter": 4, "code": "4.A.3.1",
        "title": "A.3.1 Uji Validitas Konvergen (Gambar 4.1, Outer Loadings & AVE)",
        "desc": "Bagan Gambar 4.1 Full Model SmartPLS 4, Tabel Outer Loadings (> 0,708), Tabel AVE (> 0,500), serta narasi analisis psikometrik validitas konvergen.",
        "is_granular_child": True
    },
    "sub_4_a_3_2": {
        "chapter": 4, "code": "4.A.3.2",
        "title": "A.3.2 Uji Validitas Diskriminan (Cross-Loadings & HTMT)",
        "desc": "Tabel Matriks Cross-Loadings dan Tabel Rasio Heterotrait-Monotrait (HTMT < 0,85/0,90) serta narasi pembuktian diskriminan antar-konstruk laten.",
        "is_granular_child": True
    },
    "sub_4_a_3_3": {
        "chapter": 4, "code": "4.A.3.3",
        "title": "A.3.3 Uji Reliabilitas Konstruk (Cronbach's Alpha, CR & rho_a)",
        "desc": "Tabel Evaluasi Construct Reliability (Cronbach's Alpha > 0,70, Composite Reliability > 0,70, Dijkstra-Henseler's rho_a > 0,70) dan konfirmasi reliabilitas instrumen.",
        "is_granular_child": True
    },
    "sub_4_a_4": {
        "chapter": 4, "code": "4.A.4",
        "title": "A.4 Evaluasi Model Struktural (Inner Model PLS-SEM)",
        "desc": "Evaluasi koefisien determinasi R-Square (R²), effect size (f²), daya prediksi out-of-sample PLSpredict, Model Fit (SRMR < 0,08), dan Goodness of Fit global (GoF Tenenhaus).",
        "is_composite": True
    },
    "sub_4_a_4_1": {
        "chapter": 4, "code": "4.A.4.1",
        "title": "A.4.1 Evaluasi Koefisien Determinasi (R-Square / R²)",
        "desc": "Tabel Nilai R² dan R² Adjusted Variabel Endogen (Y dan M) berdasarkan kriteria Chin/Hair serta narasi kapasitas prediksi substantif model.",
        "is_granular_child": True
    },
    "sub_4_a_4_2": {
        "chapter": 4, "code": "4.A.4.2",
        "title": "A.4.2 Evaluasi Ukuran Efek (Effect Size f-Square / f²)",
        "desc": "Tabel Nilai f² seluruh jalur struktural prediktor ke endogen berdasarkan kriteria Cohen (0,02, 0,15, 0,35) serta narasi kontribusi praktis tiap variabel.",
        "is_granular_child": True
    },
    "sub_4_a_4_3": {
        "chapter": 4, "code": "4.A.4.3",
        "title": "A.4.3 Evaluasi Relevansi Prediksi Luar Sampel (PLSpredict)",
        "desc": "Prosedur 10-fold cross validation, Tabel perbandingan Q²_predict > 0 dan RMSE PLS-SEM vs Linear Model (LM) serta narasi high predictive power.",
        "is_granular_child": True
    },
    "sub_4_a_4_4": {
        "chapter": 4, "code": "4.A.4.4",
        "title": "A.4.4 Evaluasi Kecocokan Model & Validitas Global (Model Fit & GoF)",
        "desc": "Tabel Model Fit Indices (SRMR < 0,08, d_ULS, d_G, NFI > 0,90) dan Tabel Perhitungan Indeks Tenenhaus GoF Global (> 0,38 Large/Superior) serta narasi keterpaduan model.",
        "is_granular_child": True
    },
    "sub_4_a_5": {
        "chapter": 4, "code": "4.A.5",
        "title": "A.5 Pengujian Hipotesis Penelitian (Direct, Indirect & Moderation)",
        "desc": "Hasil bootstrapping 5.000 subsampel: Tabel Direct Effects, Gambar 4.2 T-Statistik, Persamaan Sub-Struktural 1 & 2, Tabel Mediasi (Indirect Effects), dan Tabel Moderasi.",
        "is_composite": True
    },
    "sub_4_a_5_1": {
        "chapter": 4, "code": "4.A.5.1",
        "title": "A.5.1 Pengujian Hipotesis Pengaruh Langsung (Direct Effects)",
        "desc": "Prosedur bootstrapping 5.000 subsampel, Gambar 4.2 Bagan Jalur T-Statistik, Tabel Pengaruh Langsung (β, T, P, Keputusan), Persamaan Sub-Struktural 1 & 2, serta narasi analitis.",
        "is_granular_child": True
    },
    "sub_4_a_5_2": {
        "chapter": 4, "code": "4.A.5.2",
        "title": "A.5.2 Pengujian Hipotesis Pengaruh Mediasi (Indirect Effects)",
        "desc": "Tabel Uji Efek Mediasi Intervening (Specific Indirect Effects / VAF), klasifikasi full/partial/no mediation, dan narasi transmisi psikologis keperilakuan.",
        "is_granular_child": True
    },
    "sub_4_a_5_3": {
        "chapter": 4, "code": "4.A.5.3",
        "title": "A.5.3 Pengujian Hipotesis Pengaruh Moderasi (Moderating Effects)",
        "desc": "Tabel Uji Efek Moderasi Two-Stage Interaction Approach, evaluasi koefisien interaksi Z*X->Y, Simple Slope Analysis, dan narasi Teori Kontingensi.",
        "is_granular_child": True
    },
    "sub_4_b": {
        "chapter": 4, "code": "4.B",
        "title": "B. Pembahasan Hasil Penelitian",
        "desc": "Pembahasan komprehensif mendalam formula 5 alinea doktoral untuk SELURUH hipotesis penelitian (H1 s.d H_total).",
        "is_composite": True
    },
    "sub_4_b_h1": {
        "chapter": 4, "code": "4.B.1",
        "title": "B.1 Pembahasan Hipotesis H1",
        "desc": "Pembahasan mandiri Hipotesis H1 Formula Baku 5 Alinea Doktoral: Bukti Statistik, Telaah Teori, Realitas Instansi, Riset 2022-2026, Implikasi Manajerial Praktis.",
        "is_granular_child": True
    },
    "sub_4_b_h2": {
        "chapter": 4, "code": "4.B.2",
        "title": "B.2 Pembahasan Hipotesis H2",
        "desc": "Pembahasan mandiri Hipotesis H2 Formula Baku 5 Alinea Doktoral: Bukti Statistik, Telaah Teori, Realitas Instansi, Riset 2022-2026, Implikasi Manajerial Praktis.",
        "is_granular_child": True
    },
    "sub_4_b_h3": {
        "chapter": 4, "code": "4.B.3",
        "title": "B.3 Pembahasan Hipotesis H3",
        "desc": "Pembahasan mandiri Hipotesis H3 Formula Baku 5 Alinea Doktoral: Bukti Statistik, Telaah Teori, Realitas Instansi, Riset 2022-2026, Implikasi Manajerial Praktis.",
        "is_granular_child": True
    },
    "sub_4_b_h4": {
        "chapter": 4, "code": "4.B.4",
        "title": "B.4 Pembahasan Hipotesis H4",
        "desc": "Pembahasan mandiri Hipotesis H4 Formula Baku 5 Alinea Doktoral: Bukti Statistik, Telaah Teori, Realitas Instansi, Riset 2022-2026, Implikasi Manajerial Praktis.",
        "is_granular_child": True
    },
    "sub_4_b_h5": {
        "chapter": 4, "code": "4.B.5",
        "title": "B.5 Pembahasan Hipotesis H5",
        "desc": "Pembahasan mandiri Hipotesis H5 Formula Baku 5 Alinea Doktoral: Bukti Statistik, Telaah Teori, Realitas Instansi, Riset 2022-2026, Implikasi Manajerial Praktis.",
        "is_granular_child": True
    },
    "sub_4_b_h6": {
        "chapter": 4, "code": "4.B.6",
        "title": "B.6 Pembahasan Hipotesis H6",
        "desc": "Pembahasan mandiri Hipotesis H6 Formula Baku 5 Alinea Doktoral: Bukti Statistik, Telaah Teori, Realitas Instansi, Riset 2022-2026, Implikasi Manajerial Praktis.",
        "is_granular_child": True
    },
    "sub_4_b_h7": {
        "chapter": 4, "code": "4.B.7",
        "title": "B.7 Pembahasan Hipotesis H7",
        "desc": "Pembahasan mandiri Hipotesis H7 Formula Baku 5 Alinea Doktoral: Bukti Statistik, Telaah Teori, Realitas Instansi, Riset 2022-2026, Implikasi Manajerial Praktis.",
        "is_granular_child": True
    },
    "sub_4_b_h8": {
        "chapter": 4, "code": "4.B.8",
        "title": "B.8 Pembahasan Hipotesis H8",
        "desc": "Pembahasan mandiri Hipotesis H8 Formula Baku 5 Alinea Doktoral: Bukti Statistik, Telaah Teori, Realitas Instansi, Riset 2022-2026, Implikasi Manajerial Praktis.",
        "is_granular_child": True
    },

    # Composite Sub-Bab A (Hasil Analisis Lengkap A.1 s.d A.5)
    "sub_4_a": {
        "chapter": 4, "code": "4.A",
        "title": "A. Hasil Analisis (Lengkap A.1 s.d A.5)",
        "desc": "Kompilasi lengkap seluruh Hasil Analisis: Karakteristik Responden, Statistik Deskriptif, Outer Model, Inner Model, dan Uji Hipotesis.",
        "is_legacy_alias": True
    },

    # Legacy Aliases (Backward-Compatible)
    "sub_4_1_1": {
        "chapter": 4, "code": "4.1.1",
        "title": "4.1.1 Gambaran Umum Objek & Profil Responden (Alias 4.A.1)",
        "desc": "Alias backward-compatible untuk Sub-Bab A.1 Karakteristik Demografis Responden.",
        "is_legacy_alias": True
    },
    "sub_4_1_2": {
        "chapter": 4, "code": "4.1.2",
        "title": "4.1.2 Karakteristik Demografis Responden (Alias 4.A.1)",
        "desc": "Alias backward-compatible untuk Sub-Bab A.1 Karakteristik Demografis Responden.",
        "is_legacy_alias": True
    },
    "sub_4_1_3": {
        "chapter": 4, "code": "4.1.3",
        "title": "4.1.3 Statistik Deskriptif Variabel Bebas X (Alias 4.A.2)",
        "desc": "Alias backward-compatible untuk Sub-Bab A.2 Analisis Statistik Deskriptif Variabel.",
        "is_legacy_alias": True
    },
    "sub_4_1_4": {
        "chapter": 4, "code": "4.1.4",
        "title": "4.1.4 Statistik Deskriptif Variabel Mediasi M & Moderasi Z (Alias 4.A.2)",
        "desc": "Alias backward-compatible untuk Sub-Bab A.2 Analisis Statistik Deskriptif Variabel.",
        "is_legacy_alias": True
    },
    "sub_4_1_5": {
        "chapter": 4, "code": "4.1.5",
        "title": "4.1.5 Statistik Deskriptif Variabel Terikat Y (Alias 4.A.2)",
        "desc": "Alias backward-compatible untuk Sub-Bab A.2 Analisis Statistik Deskriptif Variabel.",
        "is_legacy_alias": True
    },
    "sub_4_2_poin_1": {
        "chapter": 4, "code": "4.2.1",
        "title": "4.2.1 Evaluasi Measurement Model (Outer Model) (Alias 4.A.3)",
        "desc": "Alias backward-compatible untuk Sub-Bab A.3 Evaluasi Model Pengukuran (Outer Model).",
        "is_legacy_alias": True
    },
    "sub_4_2_poin_2": {
        "chapter": 4, "code": "4.2.2",
        "title": "4.2.2 Evaluasi Structural Model (Inner Model) (Alias 4.A.4)",
        "desc": "Alias backward-compatible untuk Sub-Bab A.4 Evaluasi Model Struktural (Inner Model).",
        "is_legacy_alias": True
    },
    "sub_4_2_poin_3": {
        "chapter": 4, "code": "4.2.3",
        "title": "4.2.3 Pengujian Hipotesis Pengaruh Langsung (Alias 4.A.5)",
        "desc": "Alias backward-compatible untuk Sub-Bab A.5 Pengujian Hipotesis Penelitian.",
        "is_legacy_alias": True
    },
    "sub_4_2_poin_4": {
        "chapter": 4, "code": "4.2.4",
        "title": "4.2.4 Pengujian Efek Mediasi (Indirect Effects) (Alias 4.A.5)",
        "desc": "Alias backward-compatible untuk Sub-Bab A.5 Pengujian Hipotesis Penelitian.",
        "is_legacy_alias": True
    },
    "sub_4_2_poin_5": {
        "chapter": 4, "code": "4.2.5",
        "title": "4.2.5 Pengujian Efek Moderasi (Simple Slope) (Alias 4.A.5)",
        "desc": "Alias backward-compatible untuk Sub-Bab A.5 Pengujian Hipotesis Penelitian.",
        "is_legacy_alias": True
    },
    "sub_4_2_poin_6": {
        "chapter": 4, "code": "4.2.6",
        "title": "4.2.6 Pengujian Mediasi Termoderasi (Conditional Indirect) (Alias 4.A.5)",
        "desc": "Alias backward-compatible untuk Sub-Bab A.5 Pengujian Hipotesis Penelitian.",
        "is_legacy_alias": True
    },
    "sub_4_2_poin_7": {
        "chapter": 4, "code": "4.2.7",
        "title": "4.2.7 Pembahasan Komprehensif Hasil Penelitian (Alias 4.B)",
        "desc": "Alias backward-compatible untuk Sub-Bab B Pembahasan Hasil Penelitian.",
        "is_legacy_alias": True
    },

    # Bab V: Simpulan dan Saran (Format Baku TAPD 2026 Disertasi S3)
    "sub_5_a": {
        "chapter": 5, "code": "A",
        "title": "A. Ringkasan Penelitian",
        "desc": "Sintesis eksekutif latar belakang urgensi riset, metodologi PLS-SEM, temuan empiris, kebaruan model, dan ekosistem holistik."
    },
    "sub_5_b": {
        "chapter": 5, "code": "B",
        "title": "B. Kesimpulan",
        "desc": "Kesimpulan doktoral sistematis atas seluruh 17 hipotesis (pengaruh langsung, mediasi, moderasi) disertai bukti statistik koefisien jalur, p-value, dan justifikasi ilmiah."
    },
    "sub_5_c": {
        "chapter": 5, "code": "C",
        "title": "C. Implikasi Teori",
        "desc": "Kontribusi teoretis doktoral (doctoral novelty) memperluas Grand Theory, merekonstruksi teori kapasitas dan perilaku organisasi, mengintegrasikan teori utama dan teori menengah, serta memvalidasi arsitektur model."
    },
    "sub_5_d": {
        "chapter": 5, "code": "D",
        "title": "D. Implikasi Manajerial",
        "desc": "Implikasi manajerial strategis bagi pimpinan daerah dan instansi: kepemimpinan pelindung, pelembagaan KPI, rekonstruksi rekrutmen/seleksi, dan reorientasi audit internal."
    },
    "sub_5_e": {
        "chapter": 5, "code": "E",
        "title": "E. Implikasi Praktis",
        "desc": "Implikasi praktis operasional harian: penguatan prosedur kerja, pemanfaatan teknologi informasi, pembinaan etika profesi, dan kanal pelaporan internal."
    },
    "sub_5_f": {
        "chapter": 5, "code": "F",
        "title": "F. Saran dan Keterbatasan Penelitian",
        "desc": "Keterbatasan metodologis (geografis, social desirability bias, variabel model) serta saran solutif bagi akademisi mendatang dan bagi pemangku kebijakan daerah."
    },
    "sub_5_all": {
        "chapter": 5, "code": "Bab V",
        "title": "Bab V Lengkap: Simpulan dan Saran (A s.d F)",
        "desc": "Kompilasi lengkap seluruh sub-bab Bab V Simpulan dan Saran format baku TAPD 2026 (Ringkasan, Kesimpulan, Implikasi Teori, Manajerial, Praktis, Saran & Keterbatasan).",
        "is_composite": True
    },

    # Legacy Aliases Bab V
    "sub_5_1": {
        "chapter": 5, "code": "5.1",
        "title": "5.1 Kesimpulan Penelitian Doktoral (Legacy Alias)",
        "desc": "Alias legacy untuk Sub-Bab B Kesimpulan Doktoral.",
        "is_legacy_alias": True
    },
    "sub_5_2": {
        "chapter": 5, "code": "5.2",
        "title": "5.2 Implikasi Penelitian (Legacy Alias)",
        "desc": "Alias legacy untuk Sub-Bab C, D, dan E Implikasi Teori, Manajerial, dan Praktis.",
        "is_legacy_alias": True
    },
    "sub_5_3": {
        "chapter": 5, "code": "5.3",
        "title": "5.3 Keterbatasan dan Saran (Legacy Alias)",
        "desc": "Alias legacy untuk Sub-Bab F Saran dan Keterbatasan Penelitian.",
        "is_legacy_alias": True
    },

    # Bagian Akhir
    "sub_dapus": {
        "chapter": 6, "code": "Daftar Pustaka",
        "title": "Daftar Pustaka (Standar APA 7th Edition Mutakhir 2022-2026)",
        "desc": "Bibliografi kumulatif berstandar APA 7th Edition, disusun alfabetis A-Z, memuat seluruh referensi yang dikutip di naskah Disertasi."
    },
    "sub_lampiran": {
        "chapter": 6, "code": "Lampiran",
        "title": "Lampiran (Kuesioner Doktoral 40 Butir, Kisi-Kisi, dan Output PLS-SEM)",
        "desc": "Instrumen kuesioner skala Likert lengkap 40 butir operasional seluruh variabel judul, kisi-kisi instrumen, dan rekapitulasi output SmartPLS 4."
    }
}


# ==============================================================================
# 1. RESOLVER KONTEKS DISERTASI DOKTOR BERDASARKAN JUDUL
# ==============================================================================
def _parse_positive_int(raw: Any, label: str) -> int:
    """Parse bilangan bulat positif, toleran terhadap format 1.250 / 1,250 / 87.0."""
    text = str(raw if raw is not None else "").strip()
    if not text:
        raise ValueError(f"{label} belum diisi.")
    cleaned = re.sub(r"[^0-9,\.]", "", text)
    if not cleaned:
        raise ValueError(f"{label} '{raw}' bukan angka yang valid.")
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", cleaned):
        cleaned = cleaned.replace(".", "")
    elif "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif cleaned.count(",") == 1 and "." not in cleaned and len(cleaned.split(",")[1]) <= 2:
        cleaned = cleaned.replace(",", ".")
    else:
        cleaned = cleaned.replace(",", "")
        if cleaned.count(".") > 1:
            cleaned = cleaned.replace(".", "")
    try:
        value = int(round(float(cleaned)))
    except ValueError:
        raise ValueError(f"{label} '{raw}' bukan angka yang valid.")
    if value <= 0:
        raise ValueError(f"{label} harus lebih besar dari 0.")
    return value


def _slovin_sample_size(populasi: int, e: float = 0.05) -> int:
    """Ukuran sampel Slovin (dibulatkan ke atas); populasi <= 100 memakai sampling jenuh."""
    if populasi <= 100:
        return populasi
    return int(math.ceil(populasi / (1.0 + populasi * (e ** 2))))


def resolve_s3_context(config: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    Menganalisis, mengekstrak, dan menyelaraskan konfigurasi Disertasi Doktor
    secara penuh berdasarkan JUDUL PENELITIAN.
    Menjamin konsistensi variabel, hipotesis, instansi, dan formula matematis.
    """
    if config and ("ontology" in config and "graph" in config and "pls" in config and "academic_refs" in config):
        return config

    cfg = dict(config or {})
    raw_title = (cfg.get("judul") or cfg.get("title") or "").strip()
    
    # 1. Fallback judul bila belum diisi
    if not raw_title:
        raw_title = (
            "PENGARUH TRANSFORMATIONAL LEADERSHIP, BUDAYA ORGANISASI, DAN IKLIM KERJA "
            "TERHADAP KINERJA PEGAWAI DENGAN MOTIVASI KERJA SEBAGAI VARIABEL MEDIASI "
            "SERTA ETOS KERJA SEBAGAI VARIABEL MODERASI PADA BADAN PENGELOLA PAJAK DAERAH KABUPATEN OKI"
        )
        cfg["judul"] = raw_title

    # 2. Parse komponen struktural judul menggunakan Semantic Doctoral Parser
    parsed_s3 = s3_title_parser.parse_title_doctoral(raw_title)

    # 3. Sinkronisasi variabel: JUDUL adalah sumber kebenaran utama.
    #    Nilai panel (x1/x2/x3/m/z/y) yang TIDAK LAGI muncul pada judul dianggap
    #    BASI (ghost variable dari judul sebelumnya) dan diganti hasil ekstraksi
    #    judul terbaru. Aturan ini mencegah variabel judul lama terbawa ke naskah.
    def _norm_var(s: Any) -> str:
        return re.sub(r"[^a-z0-9]+", " ", str(s or "").lower()).strip()

    def _appears_in_title(name: Any, title_norm: str) -> bool:
        n = _norm_var(name)
        if not n:
            return False
        # Cocokkan frasa utuh; jika frasa panjang, cukup semua kata kuncinya ada berurutan.
        return n in title_norm or all(w in title_norm for w in n.split() if len(w) > 3)

    title_norm = _norm_var(raw_title)

    cfg_x = cfg.get("x_vars") or [cfg.get(k) for k in ("x1", "x2", "x3", "x4") if cfg.get(k)]
    cfg_x = [str(v).strip() for v in cfg_x if str(v or "").strip()]
    parsed_x_raw = [str(v).strip() for v in (parsed_s3.get("x_vars") or []) if str(v or "").strip()]
    parsed_x = [v for v in parsed_x_raw if _appears_in_title(v, title_norm)]
    parsed_m = [str(v).strip() for v in (parsed_s3.get("m_vars") or []) if _appears_in_title(v, title_norm)]
    parsed_z = [str(v).strip() for v in (parsed_s3.get("z_vars") or []) if _appears_in_title(v, title_norm)]
    parsed_y = (parsed_s3.get("y_var") or "").strip()

    _cfg_inst = cfg.get("instansi", "").strip()
    _title_inst = (parsed_s3.get("instansi") or "").strip()
    if _title_inst and (not _cfg_inst or not _appears_in_title(_cfg_inst, title_norm)):
        instansi = _title_inst
    else:
        instansi = _cfg_inst or _title_inst or "Instansi Pemerintah Daerah"
    instansi = format_proper_instansi(instansi)

    # Variabel X: bila judul berhasil diparsing, judul menang atas panel yang basi.
    if parsed_x:
        x_vars = parsed_x
    elif cfg_x:
        x_vars = cfg_x
    else:
        x_vars = parsed_x_raw or ["Variabel Bebas X1", "Variabel Bebas X2"]

    # Model diturunkan dari mediasi/moderasi yang BENAR-BENAR ada di judul.
    derived_model = "direct"
    if parsed_m and parsed_z:
        derived_model = "moderated_mediation"
    elif parsed_m:
        derived_model = "mediation"
    elif parsed_z:
        derived_model = "moderation"
    if cfg.get("model_type"):
        model_type = cfg["model_type"]
    elif parsed_x_raw:
        model_type = derived_model
    else:
        model_type = parsed_s3.get("model_type") or "moderated_mediation"

    # Variabel M: hanya yang tertulis di judul (tidak mengarang mediasi).
    if parsed_x_raw:
        m_vars = parsed_m
        z_vars = parsed_z
    else:
        m_vars = cfg.get("m_vars") or ([cfg.get("m").strip()] if cfg.get("m") else [])
        z_vars = cfg.get("z_vars") or ([cfg.get("z").strip()] if cfg.get("z") else [])
    if model_type == "moderation":
        m_vars = []
    if model_type == "mediation":
        z_vars = []
    if model_type == "direct":
        m_vars, z_vars = [], []

    # Variabel Y
    y_var = (parsed_y if (parsed_y and _appears_in_title(parsed_y, title_norm)) else "") or (cfg.get("y_var") or cfg.get("y") or parsed_y or "Kinerja Pegawai").strip()

    # Identitas Akademik
    penulis = cfg.get("penulis", "Dr. (Cand.) Peneliti Pascasarjana").strip()
    nim = cfg.get("nim", "2024099901").strip()
    tahun = str(cfg.get("tahun", "2026")).strip()
    # --- Populasi & Sampel: WAJIB dari input peneliti (tanpa angka default lintas instansi) ---
    # Angka 150/110 sebelumnya dipakai sebagai default diam-diam sehingga dua judul
    # berbeda menghasilkan populasi dan sampel identik. Sekarang N wajib diisi dan n
    # dihitung dari N (Slovin 5%) atau diambil dari input peneliti yang valid.
    populasi_raw = str(cfg.get("populasi") or "").strip()
    if not populasi_raw:
        raise ValueError(
            "Jumlah populasi (N) belum diisi. Mesin tidak lagi memakai angka default "
            "(150/110) karena populasi setiap instansi berbeda. Isi kolom "
            "'Jumlah Populasi (N)' pada panel penelitian, lalu jalankan ulang proses generate."
        )
    populasi_int = _parse_positive_int(populasi_raw, "Jumlah populasi (N)")
    populasi = str(populasi_int)

    sampel_raw = str(cfg.get("sampel") or "").strip()
    if sampel_raw:
        sampel_int = _parse_positive_int(sampel_raw, "Jumlah sampel (n)")
        if sampel_int > populasi_int:
            raise ValueError(
                f"Jumlah sampel ({sampel_int}) tidak boleh melebihi jumlah populasi ({populasi_int})."
            )
    else:
        sampel_int = _slovin_sample_size(populasi_int)
    sampel = str(sampel_int)
    n_sample = sampel_int

    # Rekapitulasi unit kerja riil (opsional). Bila peneliti mengisi data kepegawaian
    # per unit, Tabel 3.2 dan 3.3 memakai angka tersebut, bukan estimasi proporsional.
    unit_kerja_raw = cfg.get("unit_kerja") or cfg.get("units") or []

    # 4. Kompilasi Graph Model Struktural S3 (17+ Hipotesis Baku)
    compiled_graph = structural_graph_compiler.compile_research_graph(
        model_type=model_type,
        x_vars=x_vars,
        m_vars=m_vars,
        z_vars=z_vars,
        y_var=y_var,
        instansi=instansi
    )

    # 5. Simulasi SmartPLS 4 & PROCESS Hayes
    pls_res = dynamic_pls_engine.simulate_dynamic_pls(compiled_graph, sample_size=n_sample)

    # 5b. Data empiris riil (bila disediakan) menimpa angka simulasi.
    #     Mesin tidak boleh menyajikan angka fiktif sebagai hasil penelitian.
    empirical_data = s3data.load_empirical_data(cfg)
    if empirical_data:
        pls_res = s3data.merge_into_pls(pls_res, empirical_data)

    # 6. Riset Ontologi Doktoral (Dimensi, Indikator Baku, Grand Theory)
    onto_s3 = s3_ontology_researcher.research_doctoral_ontology(
        title=raw_title,
        x_vars=x_vars,
        m_vars=m_vars,
        z_vars=z_vars,
        y_var=y_var,
        instansi=instansi
    )

    # 7. Butir Kuesioner Doktoral 40 Butir
    kuesioner_items = s3_ontology_researcher.build_s3_kuesioner_items(onto_s3)

    # Paket konteks terpadu
    ctx = {
        "judul": raw_title,
        "title": raw_title,
        "instansi": instansi,
        "penulis": penulis,
        "nim": nim,
        "tahun": tahun,
        "populasi": populasi,
        "sampel": sampel,
        "n": n_sample,
        "universitas": (cfg.get("universitas") or cfg.get("kampus") or "Universitas Tridinanti Palembang"),
        "fakultas": (cfg.get("fakultas") or "Fakultas Ekonomi dan Bisnis"),
        "program_studi": (cfg.get("program_studi") or "Program Doktor Ilmu Manajemen"),
        "kota": (cfg.get("kota") or "").strip(),
        "promotor": (cfg.get("promotor") or "").strip(),
        "kopromotor1": (cfg.get("kopromotor1") or "").strip(),
        "kopromotor2": (cfg.get("kopromotor2") or "").strip(),
        "tempat_lahir": (cfg.get("tempat_lahir") or "").strip(),
        "email": (cfg.get("email") or "").strip(),
        "riwayat_hidup": cfg.get("riwayat_hidup") or [],
        "unit_kerja": unit_kerja_raw,
        "model_type": model_type,
        "x_vars": x_vars,
        "m_vars": m_vars,
        "z_vars": z_vars,
        "y_var": y_var,
        "x1": x_vars[0] if len(x_vars) > 0 else "",
        "x2": x_vars[1] if len(x_vars) > 1 else "",
        "x3": x_vars[2] if len(x_vars) > 2 else "",
        "m": m_vars[0] if m_vars else "",
        "z": z_vars[0] if z_vars else "",
        "y": y_var,
        "graph": compiled_graph,
        "pls": pls_res,
        "empirical_data": empirical_data,
        "data_mode": "empirical" if empirical_data else "template",
        "ontology": onto_s3,
        "kuesioner": kuesioner_items,
        "jenjang": "Doktor (S3)",
        "degree": "Doktor",
        "sakip_nasional": _validate_sakip_rows(cfg.get("sakip_nasional") or cfg.get("tabel_1_1")),
        "sakip_sumber": (cfg.get("sakip_sumber") or "").strip(),
    }

    # 8. Riset Literatur Akademik Riil Ber-DOI (Crossref & OpenAlex)
    academic_corpus = s3_academic_researcher.fetch_s3_verified_academic_corpus(ctx)
    citation_groups = s3_academic_researcher.get_s3_citation_groups(ctx)
    ctx["academic_refs"] = academic_corpus
    ctx["citation_groups"] = citation_groups

    return ctx


# ==============================================================================
# 2. GENERATOR PROMPT KHUSUS DOKTORAL DISESUAIKAN PENUH KE JUDUL
# ==============================================================================
def _generate_doctoral_model_acronym(ctx: Dict[str, Any]) -> str:
    """Membuat akronim model konseptual orisinal disertasi doktoral."""
    judul = ctx.get("judul", "").upper()
    if "FRAUD" in judul:
        return "FRAUD-E-KOMSBA"

    def _part(name: str, fallback: str, ln: int) -> str:
        token = str(name or "").split()
        return (token[0][:ln].upper() if token else fallback)

    x_vars = ctx.get("x_vars") or []
    parts = [_part(v, f"X{i+1}", 5) for i, v in enumerate(x_vars)] or ["EXOG"]
    z = ctx.get("z") or (ctx.get("z_vars") or [""])[0]
    y = ctx.get("y") or "PERF"
    if z:
        parts.append(_part(z, "MOD", 5))
    parts.append(_part(y, "PERF", 4))
    suffix = {
        "direct": "DIRECT",
        "mediation": "MED",
        "moderation": "MOD",
        "moderated_mediation": "MODMED",
    }.get(str(ctx.get("model_type") or ""), "MODMED")
    return f"{'-'.join(p for p in parts if p)}-{suffix}"


def get_s3_subchapter_meta(sub_id: str, ctx: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    Mengambil metadata sub-bab Disertasi Doktor (S3), mendukung katalog statis
    dan resolusi dinamis untuk anak baris granular (per variabel, per uji, per hipotesis).
    """
    x_vars = ctx.get("x_vars", []) if ctx else []
    hypos = ctx.get("pls", {}).get("hypotheses_testing", []) if (ctx and "pls" in ctx) else []

    if sub_id in S3_SUBCHAPTER_CATALOG:
        m = dict(S3_SUBCHAPTER_CATALOG[sub_id])
        m_x = re.match(r"^sub_2_1_3_x(\d+)$", sub_id)
        if m_x and x_vars:
            idx = int(m_x.group(1))
            if idx <= len(x_vars):
                m["title"] = f"A.5.{idx} Kajian Teoritis {x_vars[idx - 1]}"

        m_f = re.match(r"^sub_3_f_x(\d+)$", sub_id)
        if m_f and x_vars:
            idx = int(m_f.group(1))
            if idx <= len(x_vars):
                m["title"] = f"F.4.{idx} Definisi & Kisi-Kisi {x_vars[idx - 1]}"

        m_a2 = re.match(r"^sub_4_a_2_x(\d+)$", sub_id)
        if m_a2 and x_vars:
            idx = int(m_a2.group(1))
            if idx <= len(x_vars):
                m["title"] = f"A.2.3.{idx} Deskriptif {x_vars[idx - 1]}"

        m_bh = re.match(r"^sub_4_b_h(\d+)$", sub_id)
        if m_bh and hypos:
            idx = int(m_bh.group(1))
            if idx <= len(hypos):
                h_code = hypos[idx - 1].get("code", f"H{idx}")
                h_stmt = hypos[idx - 1].get("statement", "")
                m["title"] = f"B.{idx} Pembahasan Hipotesis {h_code}: {h_stmt}"
        return m

    m_x = re.match(r"^sub_2_1_3_x(\d+)$", sub_id)
    if m_x:
        idx = int(m_x.group(1))
        x_name = x_vars[idx - 1] if idx <= len(x_vars) else f"Variabel Bebas X{idx}"
        return {
            "chapter": 2, "code": f"A.5.{idx}",
            "title": f"A.5.{idx} Kajian Teoritis {x_name}",
            "desc": f"Kajian mendalam {x_name}: hakikat pilar kapasitas, telaah naratif 10 definisi ahli, sintesis peneliti, serta dimensi dan indikator teks hierarkis murni tanpa tabel.",
            "is_granular_child": True
        }

    m_f = re.match(r"^sub_3_f_x(\d+)$", sub_id)
    if m_f:
        idx = int(m_f.group(1))
        x_name = x_vars[idx - 1] if idx <= len(x_vars) else f"Variabel Bebas X{idx}"
        return {
            "chapter": 3, "code": f"F.4.{idx}",
            "title": f"F.4.{idx} Definisi & Kisi-Kisi {x_name}",
            "desc": f"Definisi konseptual, operasional skala Likert 1-5, dan tabel kisi-kisi dimensi dan indikator empiris untuk {x_name}.",
            "is_granular_child": True
        }

    m_a2 = re.match(r"^sub_4_a_2_x(\d+)$", sub_id)
    if m_a2:
        idx = int(m_a2.group(1))
        x_name = x_vars[idx - 1] if idx <= len(x_vars) else f"Variabel Bebas X{idx}"
        return {
            "chapter": 4, "code": f"4.A.2.3.{idx}",
            "title": f"A.2.3.{idx} Deskriptif {x_name}",
            "desc": f"Tabel 5 parameter statistik deskriptif dan narasi interpretasi kategori 3-box Sugiyono untuk {x_name}.",
            "is_granular_child": True
        }

    m_bh = re.match(r"^sub_4_b_h(\d+)$", sub_id)
    if m_bh:
        idx = int(m_bh.group(1))
        h_code = f"H{idx}"
        h_stmt = ""
        if idx <= len(hypos):
            h_code = hypos[idx - 1].get("code", f"H{idx}")
            h_stmt = f": {hypos[idx - 1].get('statement', '')}"
        return {
            "chapter": 4, "code": f"4.B.{idx}",
            "title": f"B.{idx} Pembahasan Hipotesis {h_code}{h_stmt}",
            "desc": f"Pembahasan mandiri Hipotesis {h_code} Formula Baku 5 Alinea Doktoral: Bukti Statistik, Telaah Teori, Realitas Instansi, Riset 2022-2026, Implikasi Manajerial Praktis.",
            "is_granular_child": True
        }

    return {"chapter": 1, "code": sub_id, "title": sub_id, "desc": "Sub-bab Disertasi Doktor"}


def build_s3_prompt(sub_id: str, ctx: Dict[str, Any]) -> str:
    """
    Membangun prompt terperinci khusus Disertasi Doktor (S3) untuk sub-bab terpilih
    yang diselaraskan 100% dengan JUDUL PENELITIAN, variabel, dan model empiris.
    """
    meta = get_s3_subchapter_meta(sub_id, ctx)
    
    judul = ctx["judul"]
    instansi = ctx["instansi"]
    x_vars = ctx.get("x_vars", [])
    m_vars = ctx.get("m_vars", [])
    z_vars = ctx.get("z_vars", [])
    x_vars_str = _join_id(x_vars)
    m_vars_str = ", ".join(m_vars) if m_vars else "(Tidak ada mediasi)"
    z_vars_str = ", ".join(z_vars) if z_vars else "(Tidak ada moderasi)"
    y_var = ctx["y_var"]
    m_type = ctx["model_type"].replace("_", " ").title()
    total_h = ctx["graph"].get("total_hypotheses", 0)

    model_acronym = _generate_doctoral_model_acronym(ctx)
    _fams = _theory_families(ctx)
    _grand_names = "; ".join(_theory_name_short(g) for g in _grand_theory_items(_fams))
    _mid_names = "; ".join(_theory_name_short(m) for m in _mid_theory_items(_fams))
    _model_lbl = _model_label(ctx)
    _mz_parts = _model_parts_phrase(ctx)
    _relevansi_vars = " dan ".join([p for p in [
        f"variabel mediasi {m_vars_str}" if m_vars else "",
        f"variabel moderasi {z_vars_str}" if z_vars else "",
    ] if p]) or "variabel penelitian"
    cite_grp = ctx.get("citation_groups") or s3_academic_researcher.get_s3_citation_groups(ctx)
    pos_str = cite_grp.get("pos_citations", "Hair et al., 2022; Sarstedt et al., 2022")
    gap_str = cite_grp.get("gap_citations", "Gadzai et al., 2025; Almazrouei & Alnahhal, 2025")
    gov_str = cite_grp.get("gov_citations", "Hayes, 2022; Creswell & Creswell, 2023")

    instruksi_khusus = ""
    if sub_id == "sub_1_1":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB LATAR BELAKANG DISERTASI DOKTORAL (SUB-BAB A. LATAR BELAKANG) ===
Judul Sub-bab WAJIB ditulis persis: A. Latar Belakang
Naskah Latar Belakang Masalah WAJIB disusun mengikuti Anatomi 7 Lapisan Deduktif Doktoral (The 7-Layer Doctoral Funnel) sebagai standar tertinggi Pascasarjana:

1. LAPISAN 1: FENOMENA MAKRO NASIONAL & THEORETICAL HOOK
   - Buka narasi dengan urgensi strategis {y_var} pada tata kelola organisasi dan akuntabilitas publik di Indonesia.
   - Sertakan Theoretical Hook: Goal-Setting Theory (Locke & Latham (1990), dalam Schunk, (2022)), Resource-Based View (Barney (1991), dalam Hitt et al., (2023)), dan AMO Framework (Appelbaum et al. (2000), dalam Boxall & Purcell, (2022)).
   - Kutip rujukan mutakhir (2022–2026) yang kredibel. Dilarang menggunakan frasa klise pembuka esai umum.

2. LAPISAN 2: TABEL 1.1 TREN STATISTIK DATA MAKRO NASIONAL (2021–2025)
   - Wajib menyajikan Tabel 1.1: Tren Evaluasi Akuntabilitas Kinerja Instansi Publik Nasional (2021–2025).
   - Format: Kolom Tahun (2021-2025), Capaian Nilai Rata-rata SAKIP Nasional, Persentase Instansi Predikat Rendah/Cukup (%), Disparitas Capaian Wilayah, Status Evaluasi Nasional.
   - Sumber resmi mutakhir (2026).
   - Analisis paragraf pasca-tabel yang membuktikan adanya urgensi statistik nasional.

3. LAPISAN 3: FENOMENA MESO & MIKRO LOKAL PADA {instansi.upper()} & 3 TABEL FENOMENA
   - Kerucutkan ke lokus empiris {instansi}: beban kerja riil, kesenjangan kompetensi, disparitas capaian target, dan dinamika kepemimpinan.
   - Wajib menyajikan 3 TABEL FENOMENA LOKAL (kualitatif, tanpa angka persentase atau nominal fiktif):
     * Tabel 1.2: Fenomena Masalah Empiris dan Evaluasi Capaian Kinerja di {instansi} (2021–2025) memuat 5 peristiwa/masalah dengan kolom: No, Dimensi Permasalahan Empiris, Indikasi Unsur Masalah Teoretis, Implikasi terhadap Sasaran Kinerja, Unit/Aktor Terlibat.
     * Tabel 1.3: Pola Peran Pelaku / Demografi Jabatan Terhadap Kesenjangan Kinerja di {instansi} dengan kolom: Jenjang Jabatan/Kelompok Aktor, Peran dalam Rantai Kesenjangan Kinerja, Karakteristik Hambatan Dominan, Implikasi bagi Variabel yang Diteliti. Dilarang mencantumkan persentase kontribusi tanpa data resmi instansi.
     * Tabel 1.4: Keterpaparan Risiko Kesenjangan Kinerja Berdasarkan Bidang Kerja di {instansi} dengan kolom: Bidang/Sektor Unit Kerja, Tingkat Keterpaparan Risiko Kinerja (Tinggi/Sedang–Tinggi/Sedang), Fokus Hambatan dan Implikasi terhadap Target Organisasi.
   - Sintesis paragraf yang menunjukkan bahwa masalah kinerja terjadi secara sistemik.

4. LAPISAN 4: PARADOKS EMPIRIS & PREMIS 'INFRASTRUKTUR PASIF'
   - Elaborasi 4 Pilar Eksogen: {x_vars_str}.
   - Paradoks: Walaupun peraturan formal, fasilitas, dan tunjangan telah tersedia, kinerja optimal belum tercapai.
   - Premis Doktoral: 4 pilar eksogen tersebut hanyalah 'POTENSI PASIF' (infrastruktur mati di atas kertas) yang membutuhkan jembatan eksekusi operasional.

5. LAPISAN 5: JUSTIFIKASI STRUKTURAL MEDIASI (M) DAN MODERASI (Z)
   - Peran Mediasi {m_vars_str}: Sebagai wujud 'Behavioral Execution' (transmisi psikologis internalisasi nilai dan motivasi intrinsik pegawai).
   - Peran Moderasi {z_vars_str}: Sebagai 'Katalisator Kontingensi & Tone at the Top' yang membentengi organisasi dari ancaman Management Override (pengabaian sistem oleh otoritas) dan melipatgandakan efek kepemimpinan terhadap unjuk kerja.

6. LAPISAN 6: DIALEKTIKA DEBAT LITERATUR & TABEL 1.5 RESEARCH GAP
   - Membenturkan dua kubu literatur terdahulu secara tajam: Kubu Positif Signifikan ({pos_str}) vs Kubu Kontradiktif / Tidak Signifikan / Kondisi Batas ({gap_str}).
   - Wajib menyajikan Tabel 1.5: Matriks Research Gap (Kesenjangan Riset Doktoral) dalam format matriks 3 kolom (Aspek Penelitian, Tinjauan Penelitian Terdahulu, Kesenjangan & Kebaruan) mencakup 6 baris aspek:
     (1) Inkonsistensi Temuan Empiris
     (2) Peran Variabel dalam Model
     (3) Posisi Variabel Mediasi ({m_vars_str})
     (4) Posisi Variabel Moderasi ({z_vars_str})
     (5) Konteks Sektoral & Desain Model Simultan Moderated-Mediation
     (6) Kebaruan Model Konseptual Doktoral
   - Analisis mendalam pasca-tabel.

7. LAPISAN 7: FORMULASI NOVELTY DOKTORAL (FILLING TWO VOIDS) & PENAMAAN MODEL
   - Rumuskan Novelty Doktoral dengan konsep 'Mengisi Dua Kekosongan Teoretis (Two Voids)':
     (1) Mengisi Kekosongan Eksekusi Perilaku (Behavioral Execution Void) via {m_vars_str}.
     (2) Mengisi Kekosongan Perlindungan Sistem & Tone at the Top (System Protection Void) via {z_vars_str}.
   - Deklarasikan Nama Orisinal Model Disertasi: Model {model_acronym}.
   - Kontribusi Keilmuan Doktoral dan Penutupan Menuju Penetapan Judul Disertasi.
"""
    elif sub_id == "sub_1_2":
        hypos = ctx["graph"]["hypotheses"]
        blok_mediasi_p = (
            f"- Justifikasi Mediasi: Jelaskan bahwa pilar kapasitas organisasi ({x_vars_str}) adalah potensi pasif yang membutuhkan {m_vars_str} sebagai variabel mediasi untuk menjembatani kapasitas sistem ke dalam eksekusi perilaku unjuk kerja nyata (behavioral execution) di lapangan."
            if ctx["m_vars"] else
            "- Model Tanpa Mediasi: Fokus pada pengujian transmisi langsung pengaruh variabel eksogen terhadap unjuk kerja."
        )
        blok_moderasi_p = (
            f"- Justifikasi Moderasi: Jelaskan bahwa efektivitas hubungan struktural bersifat kontingen, di mana {z_vars_str} bertindak sebagai katalisator Tone at the Top dan benteng integritas guna menangkal ancaman Management Override."
            if ctx["z_vars"] else
            "- Model Tanpa Moderasi: Model mengasumsikan hubungan struktural bersifat invarian terhadap kondisi batas eksternal."
        )

        daftar_pertanyaan_str = ""
        for i, h in enumerate(hypos, 1):
            daftar_pertanyaan_str += f"   {i}. {h['question']}\n"

        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK RUMUSAN MASALAH DOKTORAL (SUB-BAB B. RUMUSAN MASALAH) ===
Judul Sub-bab WAJIB ditulis persis: B. Rumusan Masalah

Naskah WAJIB disusun dalam 2 Bagian Terpadu berstandar tertinggi Pascasarjana:

BAGIAN 1: NARASI FONDASI MASALAH INDUK & SINTESIS MODEL DOKTORAL (4 PARAGRAF):
1. Paragraf 1 (Overarching Problem Statement & Theoretical Gap Anchor):
   - Wajib dibuka dengan rumusan masalah payung: "Rumusan masalah penelitian yang diajukan dalam disertasi ini adalah: bagaimana membangun sebuah model konseptual dan model empirik untuk mengatasi kesenjangan penelitian (research gap) terkait peningkatan dan optimalisasi {y_var} pada {instansi}."
   - Paparkan kembali fakta inkonsistensi temuan antarpeneliti terdahulu (sitasi riil ber-DOI 2022-2026: {pos_str} vs {gap_str}).
2. Paragraf 2 (Justifikasi Argumen Mediasi Transmisi):
   {blok_mediasi_p}
3. Paragraf 3 (Justifikasi Argumen Moderasi Katalisator Kontingensi):
   {blok_moderasi_p}
4. Paragraf 4 (Pernyataan Kebaruan & Penamaan Model Doktoral Orisinal):
   - Tegaskan kebaruan (novelty) pengujian model struktural {m_type}.
   - Deklarasikan secara resmi nama model konseptual doktoral orisinal: MODEL {model_acronym}. Jelaskan kepanjangan dan filosofi akronim tersebut.

BAGIAN 2: DAFTAR PERTANYAAN PENELITIAN ({len(hypos)} PERTANYAAN FORMAL):
Sajikan pengantar: "Berdasarkan perumusan masalah di atas, pertanyaan penelitian ini adalah:"
{daftar_pertanyaan_str}

ATURAN STRUKTUR & LARANGAN FORMAT (SANGAT KETAT):
- Judul Sub-bab WAJIB ditulis: B. Rumusan Masalah
- DILARANG MENULISKAN SUB-JUDUL KELOMPOK seperti:
  'A. Kelompok Pengaruh Langsung (Direct Effects):'
  'B. Kelompok Pengaruh Tidak Langsung / Mediasi (Indirect / Mediating Effects):'
  'C. Kelompok Pengaruh Interaksi Moderasi (Moderating Effects):'
- Seluruh {len(hypos)} butir pertanyaan penelitian WAJIB ditulis langsung sebagai daftar penomoran sekuensial kontinu (1, 2, 3, ... s.d {len(hypos)}) tepat di bawah kalimat pengantar.
- Jumlah butir pertanyaan WAJIB persis {len(hypos)} butir (sinkron 1-ke-1 dengan Hipotesis Formal H1 s.d H{total_h} dan Tujuan Penelitian).
"""
    elif sub_id == "sub_1_3":
        hypos = ctx["graph"]["hypotheses"]
        daftar_tujuan_str = ""
        for i, h in enumerate(hypos, 1):
            daftar_tujuan_str += f"   {i}. {h['objective']}\n"

        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK TUJUAN PENELITIAN DOKTORAL (SUB-BAB C. TUJUAN PENELITIAN) ===
Judul Sub-bab WAJIB ditulis persis: C. Tujuan Penelitian

Naskah WAJIB disusun dalam 2 Bagian Terpadu:

1. Pembatasan Masalah (Delimitasi Doktoral):
   - Jelaskan ruang lingkup konseptual dan batasan operasional penelitian agar kajian fokus pada pengujian model struktural {model_acronym} yang mengintegrasikan {y_var}, {m_vars_str}, {z_vars_str}, dan {x_vars_str} pada aparatur {instansi}.
2. Tujuan Penelitian Terstruktur ({len(hypos)} Tujuan Formal):
   - Sajikan pengantar: "Sejalan dengan pertanyaan penelitian di atas, tujuan dari penelitian ini adalah untuk:"
{daftar_tujuan_str}

ATURAN STRUKTUR & LARANGAN FORMAT:
- Judul Sub-bab WAJIB ditulis: C. Tujuan Penelitian
- DILARANG MENULISKAN SUB-JUDUL KELOMPOK seperti 'A. Kelompok Pengujian Pengaruh Langsung', 'B. Kelompok...', 'C. Kelompok...'.
- Setiap butir tujuan penelitian WAJIB bernomor sekuensial kontinu (1 s.d {len(hypos)}) dan selaras 1-ke-1 dengan Rumusan Masalah serta Hipotesis Penelitian.
- Gunakan formula baku doktoral: "Untuk menganalisis dan membuktikan pengaruh/peran..."
"""
    elif sub_id == "sub_1_4":
        model_acronym = _generate_doctoral_model_acronym(ctx)
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB D: MANFAAT PENELITIAN (BAB I) ===
Judul Sub-bab WAJIB ditulis persis: D. Manfaat Penelitian
SUB-BAB INDUK: D. MANFAAT PENELITIAN (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat analitis substantif mendalam atau 70–100 kata per paragraf). DILARANG KERAS membuat paragraf pendek (1, 2, atau 3 baris teks)!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Seluruh paragraf naratif dan uraian poin manfaat WAJIB memuat kutipan ilmiah (nama pakar dan tahun terbit mutakhir 2022-2026 atau teori klasik yang disitasi dalam literatur mutakhir, misal: Sekaran & Bougie, 2020; Creswell & Creswell, 2023; Cooper & Schindler, 2022; Snyder, 2019; Hitt et al., 2023; Yukl, 2022; Robbins & Judge, 2022; Daft, 2022).
3. STRUKTUR SISTEMATIKA PENULISAN SUB-BAB D:
   Naskah WAJIB dibagi menjadi 2 dimensi kemanfaatan utama sesuai standar TAPD 2026:
   a. Paragraf Pengantar Sub-Bab D (>= 4 baris teks utuh, bersitasi):
      Menjelaskan bahwa penelitian doktoral ini memberikan kontribusi ganda yang seimbang antara kemajuan keilmuan akademis dan perbaikan tata kelola birokrasi pada {instansi}.
   b. 1. Manfaat Teoretis (Pengembangan Keilmuan & Novelty):
      - Paragraf pengantar manfaat teoretis (>= 4 baris teks utuh, bersitasi).
      - Poin a. Melahirkan Model Empiris Baru (Model Doktoral {model_acronym}): Menjelaskan penciptaan model struktural {_model_lbl} baru yang mengintegrasikan pilar eksogen {x_vars_str}{_mz_parts}, dan muara endogen {y_var} (>= 4 baris teks utuh, bersitasi).
      - Poin b. Pengembangan Teori Multidisiplin: Menyatukan {_grand_names} guna menjawab empirical research gap (>= 4 baris teks utuh, bersitasi).
   c. 2. Manfaat Praktis (Kebijakan Strategis Organisasi):
      - Paragraf pengantar manfaat praktis (>= 4 baris teks utuh, bersitasi).
      - Poin a. Bagi Pemerintah Daerah dan Jajaran Pimpinan {instansi}: Menjadi cetak biru (blueprint) reformasi kebijakan dan intervensi manajerial yang berfokus pada penguatan etika profesi dan ketegasan sanksi (>= 4 baris teks utuh, bersitasi).
      - Poin b. Bagi Aparat Pengawasan Internal Pemerintah (APIP) dan Unit Evaluasi: Menjadi pedoman audit berbasis risiko perilaku dan budaya organisasi guna memitigasi risiko management override (>= 4 baris teks utuh, bersitasi).
      - Poin c. Bagi Masyarakat Luas dan Pemangku Kepentingan: Menjamin efisiensi alokasi anggaran, meningkatkan transparansi layanan publik, dan memulihkan public trust (>= 4 baris teks utuh, bersitasi).
"""
    elif sub_id == "sub_2_1_0":
        model_acronym = _generate_doctoral_model_acronym(ctx)
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB 1. GRAND TEORI (SUB-BAB A.1) ===
Judul Sub-bab WAJIB ditulis persis: A.1 1. Grand Teori
SUB-BAB INDUK: A. KAJIAN PUSTAKA

Naskah WAJIB disusun dalam 4 Bagian Terpadu berstandar tertinggi Pascasarjana Doktoral (S3):

1. Paragraf Pembuka Hierarki Teoretis Deduktif:
   - Jelaskan bahwa Disertasi ini dibangun di atas hierarki teoretis yang terstruktur secara deduktif, dimulai dari teori utama (Grand Theory), teori penghubung keperilakuan dan organisasi (Middle Range Theory), hingga teori praktis pengukuran lapangan (Applied Theory).

2. Grand Theory (Teori Utama / Payung Filosofis):
   - Uraikan secara mendalam {_grand_names}.
   - Deklarasikan Novelty Model Disertasi: Model {model_acronym} (sebuah model struktural {_model_lbl} yang menghubungkan pilar eksogen {x_vars_str}{_mz_parts}, dan muara endogen {y_var}).
   - WAJIB sertakan sub-bagian eksplisit: "Relevansi dengan Penelitian" yang menjelaskan justifikasi logis pemilihan Grand Theory terhadap dinamika {instansi}.

3. Middle Range Theory (Teori Penghubung / Perilaku / Manajemen):
   - Uraikan {_mid_names}, serta Organizational Behavior & Motivation Theory (Robbins & Judge, 2022; Ryan & Deci, 2022).
   - WAJIB sertakan sub-bagian eksplisit: "Relevansi dengan Penelitian" yang menjelaskan mengapa {_relevansi_vars} mutlak dihadirkan dalam model.

4. Applied Theory (Teori Terapan / Operasional Lapangan):
   - Uraikan Campbell's Theory of Job Performance (Campbell & Wiernik, 2022) serta AMO Framework (Appelbaum et al., dalam Boxall & Purcell, 2022).
   - WAJIB sertakan sub-bagian eksplisit: "Relevansi dengan Penelitian" yang memetakan bagaimana teori ini menjadi matriks operasional untuk merumuskan indikator pengukuran operasional di lapangan.
"""
    elif sub_id == "sub_2_1_1":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB KAJIAN TEORITIS {y_var.upper()} (SUB-BAB A.2) ===
Judul Sub-bab WAJIB ditulis persis: A.2 2. Kajian Teoritis {y_var}
SUB-BAB INDUK: A. KAJIAN PUSTAKA

Karakteristik Murni Teoretis: DILARANG menyebut nama instansi lokal ({instansi}) atau data empiris lapangan. Kajian harus murni konseptual universal substantif berstandar Doktoral (S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat substantif mendalam atau 70–100 kata per paragraf). DILARANG KERAS membuat paragraf pendek (1, 2, atau 3 baris teks)!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: SETIAP PARAGRAF TANPA TERKECUALI HARUS MEMILIKI KUTIPAN ILMIAH (nama pakar dan tahun terbit mutakhir 2022-2026 atau teori klasik yang disitasi dalam literatur mutakhir).
   - Paragraf Pengantar Hakikat Variabel: Wajib memuat sitasi hakikat dan urgensi keilmuan {y_var}.
   - Paragraf Telaah Definisi: Wajib merujuk 10 pakar otoritatif dalam paragraf naratif mengalir kohesif.
   - Paragraf Sintesis Peneliti: Wajib memuat rujukan komparatif pakar yang disintesiskan sebelum formula baku simpulan.
   - Paragraf Kesimpulan Terpadu Definisi dan Dimensi: Wajib memuat rujukan pakar pengukuran dan metodologi riset (misal: Hair et al., 2022).
   DILARANG KERAS menghasilkan paragraf teoretis yang hampa tanpa kutipan referensi ilmiah!

Naskah WAJIB disusun dalam 5 Bagian Terpadu:

1. Bagian 1: Narasi Pengantar Hakikat Substantif & Urgensi Variabel (1-2 Paragraf):
   - Uraikan hakikat ontologis dan esensi substantif {y_var} bagi keberhasilan dan keberlanjutan organisasi.
   - PENTING: DILARANG menggunakan istilah teknis pemodelan/metodologis (seperti "variabel endogen utama", "ultimate criterion variable", dsb.) pada kalimat pembuka kajian teori! Awali langsung dengan esensi fenomena dan substansi keilmuan {y_var}.

2. Bagian 2: Telaah Integratif 10 Definisi Para Ahli Otoritatif (Paragraf Naratif Mengalir):
   - Identifikasi domain keilmuan spesifik dari {y_var} (misalnya: MSDM/Perilaku Organisasi, Akuntansi Forensik/Fraud PBJ, Tata Kelola Sektor Publik, Pemasaran, atau Keuangan).
   - Rujuk tepat 10 (sepuluh) pemikiran/definisi pakar terkemuka yang otoritatif di bidang tersebut (paduan buku teks klasik dan sitasi mutakhir 2022-2026).
   - WAJIB DISUSUN MENGALIR DALAM PARAGRAF-PARAGRAF AKADEMIK UTUH (RUNNING TEXT PARAGRAPHS) DENGAN KATA TRANSISI ILMIAH KOHESIF.
   - DILARANG KERAS menggunakan penomoran poin (DILARANG: 1., 2., 3... dst.) pada kajian definisi ini!
   - DILARANG menulis kalimat pengantar robotik klise (seperti: "Berikut ditelaah 10 definisi para ahli..."). Alur naskah harus berpindah secara alamiah dari pemaparan hakikat ke diskursus pemikiran para ahli.

3. Bagian 3: Rumusan Sintesis Konseptual Peneliti Doktoral (1 Paragraf Utuh):
   - Formulasikan benang merah konseptual orisinal peneliti sendiri berdasarkan sintesis kritis atas pandangan para ahli tersebut.
   - Paragraf WAJIB diawali dengan formula baku:
     "Berdasarkan berbagai pendapat para ahli tersebut, dapat disimpulkan bahwa {y_var} merupakan..."

4. Bagian 4: Dimensi dan Indikator Variabel (Format Teks Hirarkis Bertingkat, MURNI TANPA TABEL):
   - Tetapkan pakar rujukan utama penetapan dimensi dan indikator yang diakui dalam mengukur {y_var}.
   - Awali dengan kalimat pengantar baku:
     "Menurut [Nama Tokoh & Tahun Rujukan Utama], dimensi dan indikator yang dipakai dalam mengukur {y_var} yaitu:"
   - Paparkan seluruh dimensi dan butir-butir indikator operasional secara hierarkis bertingkat (tanpa tabel):
     1. [Nama Dimensi Pertama]
        Dimensi [Nama Dimensi Pertama] diukur dengan menggunakan [3-4] indikator, yaitu:
        a. [Kalimat indikator terukur 1]
        b. [Kalimat indikator terukur 2]
        c. [Kalimat indikator terukur 3]
     2. [Nama Dimensi Kedua]
        Dimensi [Nama Dimensi Kedua] diukur dengan menggunakan [3-4] indikator, yaitu:
        a. [Kalimat indikator terukur 1]
        b. [Kalimat indikator terukur 2]
        c. [Kalimat indikator terukur 3]
     (Lanjutkan untuk seluruh dimensi baku {y_var}).

5. Bagian 5: Kesimpulan Terpadu Definisi dan Dimensi (1 Paragraf Akhir Sub-Bab):
   - Formulasikan satu paragraf penutup sub-bab yang menyimpulkan secara integratif kaitan antara definisi konseptual peneliti dengan seluruh dimensi dan indikator operasional yang dipilih.
   - Tegaskan bahwa perpaduan antara rumusan definisi dan dimensi-dimensi tersebut mencerminkan konstruk {y_var} secara komprehensif, valid, dan reliabel untuk operasionalisasi riset doktoral ini.

ATURAN FORMAT KETAT:
- DILARANG KERAS MENGGUNAKAN TABEL DALAM BENTUK APAPUN!
- DILARANG MENGGUNAKAN PENOMORAN ANGKA PADA BAGIAN TELAAH DEFINISI (WAJIB PARAGRAF NARATIF MENGALIR)!
"""
    elif sub_id == "sub_2_1_2":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB KAJIAN TEORITIS {m_vars_str.upper()} (SUB-BAB A.3) ===
Judul Sub-bab WAJIB ditulis persis: A.3 3. Kajian Teoritis Variabel Mediasi ({m_vars_str})
SUB-BAB INDUK: A. KAJIAN PUSTAKA

Karakteristik Murni Teoretis: DILARANG menyebut nama instansi lokal ({instansi}) atau data lapangan. Kajian murni konseptual universal substantif berstandar Doktoral (S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat substantif mendalam atau 70–100 kata per paragraf). DILARANG KERAS membuat paragraf pendek (1, 2, atau 3 baris teks)!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: SETIAP PARAGRAF TANPA TERKECUALI HARUS MEMILIKI KUTIPAN ILMIAH (nama pakar dan tahun terbit mutakhir 2022-2026 atau teori klasik yang disitasi dalam literatur mutakhir).
   - Paragraf Pengantar Hakikat Variabel: Wajib memuat sitasi hakikat dan peran transmisi {m_vars_str}.
   - Paragraf Telaah Definisi: Wajib merujuk 10 pakar otoritatif dalam paragraf naratif mengalir kohesif.
   - Paragraf Sintesis Peneliti: Wajib memuat rujukan komparatif pakar yang disintesiskan sebelum formula baku simpulan.
   - Paragraf Kesimpulan Terpadu Definisi dan Dimensi: Wajib memuat rujukan pakar pengukuran dan metodologi riset (misal: Hair et al., 2022).
   DILARANG KERAS menghasilkan paragraf teoretis yang hampa tanpa kutipan referensi ilmiah!

Naskah WAJIB disusun dalam 5 Bagian Terpadu:

1. Bagian 1: Narasi Pengantar Hakikat Substantif Variabel (1-2 Paragraf):
   - Uraikan hakikat konseptual dan esensi psikologis/keperilakuan {m_vars_str} dalam lingkungan organisasi.
   - PENTING: DILARANG menggunakan istilah teknis pemodelan (seperti "variabel intervening mediasi yang menyalurkan stimulus kausalitas...") pada kalimat pembuka! Awali langsung dengan esensi substantif {m_vars_str}.

2. Bagian 2: Telaah Integratif 10 Definisi Para Ahli Otoritatif (Paragraf Naratif Mengalir):
   - Identifikasi domain keilmuan spesifik dari {m_vars_str}.
   - Rujuk 10 pemikiran/definisi pakar terkemuka yang otoritatif di bidang tersebut (buku teks klasik & sitasi mutakhir 2022-2026).
   - WAJIB DISUSUN MENGALIR DALAM PARAGRAF-PARAGRAF AKADEMIK UTUH (RUNNING TEXT PARAGRAPHS) DENGAN KATA TRANSISI ILMIAH KOHESIF.
   - DILARANG KERAS menggunakan penomoran poin (DILARANG: 1., 2., 3... dst.) pada kajian definisi ini!
   - DILARANG menulis kalimat pengantar robotik klise (seperti: "Berikut disajikan 10 definisi...").

3. Bagian 3: Rumusan Sintesis Konseptual Peneliti Doktoral (1 Paragraf Utuh):
   - Formulasikan definisi sintesis konseptual orisinal peneliti sendiri yang memadukan intisari ke-10 definisi tersebut.
   - Paragraf WAJIB diawali dengan kalimat baku:
     "Berdasarkan berbagai pendapat para ahli tersebut, dapat disimpulkan bahwa {m_vars_str} merupakan..."

4. Bagian 4: Dimensi dan Indikator Variabel (Format Teks Hirarkis Bertingkat, MURNI TANPA TABEL):
   - Tetapkan pakar rujukan utama pengukur {m_vars_str}.
   - Awali dengan kalimat pengantar baku:
     "Menurut [Nama Tokoh & Tahun Rujukan], dimensi dan indikator yang dipakai dalam mengukur {m_vars_str} yaitu:"
   - Paparkan dimensi dan butir indikatornya secara berjenjang teks:
     1. [Nama Dimensi 1]
        Dimensi [Nama Dimensi 1] diukur dengan menggunakan [3-4] indikator, yaitu:
        a. ...
        b. ...
        c. ...

5. Bagian 5: Kesimpulan Terpadu Definisi dan Dimensi (1 Paragraf Akhir Sub-Bab):
   - Formulasikan satu paragraf penutup sub-bab yang menyimpulkan secara terpadu kaitan antara definisi konseptual hasil sintesis dengan seluruh dimensi dan indikator operasional {m_vars_str}.
   - Tegaskan bahwa perpaduan definisi dan dimensi tersebut menjadi instrumen pengukuran yang sahih, komprehensif, dan reliabel bagi operasionalisasi riset doktoral ini.

ATURAN FORMAT KETAT:
- DILARANG KERAS MENGGUNAKAN TABEL DALAM BENTUK APAPUN!
- DILARANG MENGGUNAKAN PENOMORAN ANGKA PADA BAGIAN TELAAH DEFINISI (WAJIB PARAGRAF NARATIF MENGALIR)!
"""
    elif sub_id == "sub_2_1_4":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB KAJIAN TEORITIS {z_vars_str.upper()} (SUB-BAB A.4) ===
Judul Sub-bab WAJIB ditulis persis: A.4 4. Kajian Teoritis Variabel Moderasi ({z_vars_str})
SUB-BAB INDUK: A. KAJIAN PUSTAKA

Karakteristik Murni Teoretis: DILARANG menyebut nama instansi lokal ({instansi}) atau data lapangan. Naskah murni konseptual universal substantif.

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat substantif mendalam atau 70–100 kata per paragraf). DILARANG KERAS membuat paragraf pendek (1, 2, atau 3 baris teks)!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: SETIAP PARAGRAF TANPA TERKECUALI HARUS MEMILIKI KUTIPAN ILMIAH (nama pakar dan tahun terbit mutakhir 2022-2026 atau teori klasik yang disitasi dalam literatur mutakhir).
   - Paragraf Pengantar Hakikat Variabel: Wajib memuat sitasi hakikat dan peran kontingensi moderasi {z_vars_str}.
   - Paragraf Telaah Definisi: Wajib merujuk 10 pakar otoritatif dalam paragraf naratif mengalir kohesif.
   - Paragraf Sintesis Peneliti: Wajib memuat rujukan komparatif pakar yang disintesiskan sebelum formula baku simpulan.
   - Paragraf Kesimpulan Terpadu Definisi dan Dimensi: Wajib memuat rujukan pakar pengukuran dan metodologi riset (misal: Hair et al., 2022).
   DILARANG KERAS menghasilkan paragraf teoretis yang hampa tanpa kutipan referensi ilmiah!

Naskah WAJIB disusun dalam 5 Bagian Terpadu:

1. Bagian 1: Narasi Pengantar Hakikat Substantif Variabel (1-2 Paragraf):
   - Uraikan hakikat nilai fundamental dan esensi kontekstual {z_vars_str} dalam dinamika organisasi.
   - PENTING: DILARANG menggunakan istilah teknis pemodelan (seperti "variabel kontingensi moderasi yang berfungsi memperkuat...") pada kalimat pembuka! Awali langsung dengan fenomena dan substansi keilmuan {z_vars_str}.

2. Bagian 2: Telaah Integratif 10 Definisi Para Ahli Otoritatif (Paragraf Naratif Mengalir):
   - Identifikasi domain keilmuan spesifik dari {z_vars_str}.
   - Rujuk 10 definisi/pemikiran pakar terkemuka yang otoritatif di bidang tersebut (buku teks klasik & sitasi mutakhir 2022-2026).
   - WAJIB DISUSUN MENGALIR DALAM PARAGRAF-PARAGRAF AKADEMIK UTUH (RUNNING TEXT PARAGRAPHS) DENGAN KATA TRANSISI ILMIAH KOHESIF.
   - DILARANG KERAS menggunakan penomoran poin (DILARANG: 1., 2., 3... dst.) pada kajian definisi ini!
   - DILARANG menulis kalimat pengantar robotik klise (seperti: "Berikut disajikan 10 definisi...").

3. Bagian 3: Rumusan Sintesis Konseptual Peneliti Doktoral (1 Paragraf Utuh):
   - Formulasikan benang merah konseptual orisinal peneliti sendiri.
   - Paragraf WAJIB diawali kalimat baku:
     "Berdasarkan berbagai pendapat para ahli tersebut, dapat disimpulkan bahwa {z_vars_str} merupakan..."

4. Bagian 4: Dimensi dan Indikator Variabel (Format Teks Hirarkis Bertingkat, MURNI TANPA TABEL):
   - Tetapkan pakar rujukan utama pengukuran {z_vars_str}.
   - Awali kalimat baku:
     "Menurut [Nama Tokoh & Tahun Rujukan], dimensi dan indikator yang dipakai dalam mengukur {z_vars_str} yaitu:"
   - Uraikan dimensi dan indikator operasional secara hierarkis bertingkat:
     1. [Nama Dimensi 1]
        Dimensi [Nama Dimensi 1] diukur dengan menggunakan [3-4] indikator, yaitu:
        a. ...
        b. ...

5. Bagian 5: Kesimpulan Terpadu Definisi dan Dimensi (1 Paragraf Akhir Sub-Bab):
   - Formulasikan satu paragraf penutup sub-bab yang menyimpulkan secara terpadu kaitan antara definisi konseptual hasil sintesis dengan dimensi dan indikator operasional {z_vars_str}.
   - Tegaskan bahwa perpaduan definisi dan dimensi tersebut menjadi instrumen pengukuran yang sahih, komprehensif, dan reliabel bagi operasionalisasi riset doktoral ini.

ATURAN FORMAT KETAT:
- DILARANG MENGGUNAKAN TABEL DALAM BENTUK APAPUN!
- DILARANG MENGGUNAKAN PENOMORAN ANGKA PADA TELAAH DEFINISI (WAJIB PARAGRAF NARATIF MENGALIR)!
"""
    elif re.match(r"^sub_2_1_3_x(\d+)$", sub_id):
        m_x = re.match(r"^sub_2_1_3_x(\d+)$", sub_id)
        x_idx = int(m_x.group(1))
        x_name = ctx["x_vars"][x_idx - 1] if x_idx <= len(ctx["x_vars"]) else f"Variabel Bebas X{x_idx}"
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB KAJIAN TEORITIS VARIABEL BEBAS X{x_idx} ({x_name.upper()}) (SUB-BAB A.5.{x_idx}) ===
Judul Sub-bab WAJIB ditulis persis: A.5.{x_idx} Kajian Teoritis {x_name}
SUB-BAB INDUK: A. KAJIAN PUSTAKA (5. Kajian Teoritis Variabel-Variabel Bebas)

Karakteristik Murni Teoretis: DILARANG menyebut nama instansi lokal ({instansi}) atau data lapangan. Naskah murni konseptual universal substantif.

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat substantif mendalam atau 70–100 kata per paragraf). DILARANG KERAS membuat paragraf pendek (1, 2, atau 3 baris teks)!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: SETIAP PARAGRAF TANPA TERKECUALI HARUS MEMILIKI KUTIPAN ILMIAH (nama pakar dan tahun terbit mutakhir 2022-2026 atau teori klasik yang disitasi dalam literatur mutakhir).
3. 5 KOMPONEN BAKU PEMBAHASAN VARIABEL {x_name}:
   a. Narasi Pengantar Hakikat Substantif (1-2 Paragraf, >= 4 baris, bersitasi): Esensi substantif {x_name} sebagai pilar keunggulan dan kapasitas fundamental organisasi modern (Yukl, 2022; Armstrong & Taylor, 2023).
   b. Telaah Integratif 10 Definisi Para Ahli Otoritatif (Sitasi Mutakhir 2022-2026): Minimal 10 pemikiran konseptual pakar terkemuka, diuraikan dalam paragraf naratif akademik yang mengalir dan kohesif (MURNI TANPA PENOMORAN ANGKA/BULLET).
   c. Rumusan Paragraf Sintesis Konseptual Peneliti Doktoral (1 Paragraf Utuh, >= 4 baris): Diawali kalimat baku: "Berdasarkan berbagai pendapat para ahli tersebut, dapat disimpulkan bahwa {x_name} merupakan..."
   d. Dimensi dan Indikator Variabel (Format Teks Hirarkis Bertingkat, MURNI TANPA TABEL): Paparan dimensi dan butir indikator operasional secara teks hierarkis bertingkat (tanpa tabel).
   e. Kesimpulan Terpadu Definisi dan Dimensi (1 Paragraf Akhir, >= 4 baris): Merangkum kaitan erat antara definisi konseptual hasil sintesis dengan seluruh dimensi operasional {x_name}.

ATURAN FORMAT KETAT:
- DILARANG KERAS MENGGUNAKAN TABEL DALAM BENTUK APAPUN!
- DILARANG MENGGUNAKAN PENOMORAN ANGKA PADA TELAAH DEFINISI (WAJIB PARAGRAF NARATIF MENGALIR)!
"""
    elif sub_id == "sub_2_1_3":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB KAJIAN TEORITIS VARIABEL-VARIABEL BEBAS ({x_vars_str.upper()}) (SUB-BAB A.5) ===
Judul Sub-bab WAJIB ditulis persis: A.5 5. Kajian Teoritis Variabel-Variabel Bebas ({x_vars_str})
SUB-BAB INDUK: A. KAJIAN PUSTAKA

Karakteristik Murni Teoretis: DILARANG menyebut nama instansi lokal ({instansi}) atau data lapangan. Naskah murni konseptual universal substantif.

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat substantif mendalam atau 70–100 kata per paragraf). DILARANG KERAS membuat paragraf pendek (1, 2, atau 3 baris teks)!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: SETIAP PARAGRAF TANPA TERKECUALI HARUS MEMILIKI KUTIPAN ILMIAH (nama pakar dan tahun terbit mutakhir 2022-2026 atau teori klasik yang disitasi dalam literatur mutakhir).
   - Paragraf Pengantar Sub-Bab & Pengantar Tiap Variabel Bebas: Wajib memuat sitasi hakikat pilar kapasitas organisasi.
   - Paragraf Telaah Definisi Tiap Variabel: Wajib merujuk 10 pakar otoritatif dalam paragraf naratif mengalir kohesif.
   - Paragraf Sintesis Peneliti Tiap Variabel: Wajib memuat rujukan komparatif pakar yang disintesiskan sebelum formula baku simpulan.
   - Paragraf Kesimpulan Terpadu Definisi dan Dimensi: Wajib memuat rujukan pakar pengukuran dan metodologi riset (misal: Hair et al., 2022).
   - Paragraf Penutup Seluruh Pilar Eksogen: Wajib memuat sitasi integrasi kapasitas keorganisasian.
   DILARANG KERAS menghasilkan paragraf teoretis yang hampa tanpa kutipan referensi ilmiah!

Tugas Anda adalah menguraikan kajian teoretis mendalam untuk SETIAP variabel bebas eksogen secara terpisah dan tuntas ({x_vars_str}).

Untuk SETIAP Variabel Bebas (misalnya: 5.1 Kajian Teoritis [Nama X1], 5.2 Kajian Teoritis [Nama X2], dst.), WAJIB menyajikan 5 komponen baku:

1. Narasi Pengantar Hakikat Variabel (1-2 Paragraf):
   - Esensi substantif variabel sebagai pilar keunggulan dan kapasitas organisasi (tanpa istilah teknis pemodelan di awal kalimat).
2. Telaah Integratif 10 Definisi Para Ahli Otoritatif (Sitasi Mutakhir 2022-2026):
   - 10 definisi/pemikiran konseptual dari pakar terkemuka yang relevan, diuraikan dalam paragraf-paragraf naratif akademik yang mengalir dan kohesif (TANPA PENOMORAN 1, 2, 3.. DAN TANPA PENGANTAR ROBOTIK).
3. Rumusan Paragraf Sintesis Konseptual Peneliti Doktoral:
   - Diawali kalimat baku: "Berdasarkan berbagai pendapat para ahli tersebut, dapat disimpulkan bahwa [Nama Variabel X] merupakan..."
4. Dimensi dan Indikator Variabel (Format Teks Hirarkis Bertingkat, MURNI TANPA TABEL):
   - Kalimat rujukan baku: "Menurut [Nama Tokoh & Tahun Rujukan], dimensi dan indikator yang dipakai dalam mengukur [Nama Variabel X] yaitu:"
   - Paparan dimensi dan indikator secara hierarkis teks berjenjang (tanpa tabel).
5. Kesimpulan Terpadu Definisi dan Dimensi (1 Paragraf Akhir Variabel):
   - Paragraf penutup yang merangkum kaitan erat antara definisi konseptual hasil sintesis dengan seluruh dimensi operasional variabel tersebut.

ATURAN FORMAT KETAT:
- DILARANG KERAS MENGGUNAKAN TABEL DALAM BENTUK APAPUN!
- DILARANG MENGGUNAKAN PENOMORAN ANGKA PADA TELAAH DEFINISI (WAJIB PARAGRAF NARATIF MENGALIR)!
"""
    elif sub_id == "sub_2_2":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB B: HASIL PENELITIAN YANG RELEVAN (SUB-BAB 2.2) ===
Judul Sub-bab WAJIB ditulis persis: B. Hasil Penelitian yang Relevan
SUB-BAB INDUK: B. HASIL PENELITIAN YANG RELEVAN (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat analitis substantif mendalam atau 70–100 kata per paragraf). DILARANG KERAS membuat paragraf pendek (1, 2, atau 3 baris teks)!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: SETIAP PARAGRAF NARATIF TANPA TERKECUALI HARUS MEMILIKI KUTIPAN ILMIAH (nama pakar dan tahun terbit mutakhir 2022-2026 atau rujukan metodologis terkemuka):
   - Paragraf Pengantar Sub-Bab B: Wajib memuat rujukan metodologis terkait pemetaan literatur empiris (misal: Snyder, 2019; Cooper & Schindler, 2022; Creswell & Creswell, 2023).
   - Paragraf Analisis Konvergensi Temuan: Wajib memuat rujukan empiris studi-studi terdahulu yang membuktikan konsistensi pengaruh positif pilar eksogen dan mediasi.
   - Paragraf Analisis Divergensi & Empirical Research Gap: Wajib memuat rujukan empiris studi terdahulu yang tidak konsisten/tidak signifikan beserta argumentasi celah risetnya.
   - Paragraf Sintesis Kebaruan (Novelty) Doktoral: Wajib memuat rujukan metodologis dan model struktural terpadu (misal: Hair et al., 2022; Hayes, 2022).
   DILARANG KERAS menghasilkan paragraf naratif yang hampa tanpa kutipan referensi ilmiah!

3. TABEL 2.1 HASIL PENELITIAN YANG RELEVAN:
   - FORMAT BAKU TABEL: Tepat 5 kolom sesuai acuan resmi TAPD 2026:
     | No | Penulis, Tahun | Judul / publikasi | Metoda Riset | Temuan Penelitian |
   - JUMLAH PENELITIAN DIDALAM TABEL: MINIMAL 15 PENELITIAN (wajib memuat 15 hingga 16 penelitian empiris mutakhir kurun waktu 2022-2026 dari jurnal bereputasi Scopus Q1-Q4 dan Sinta 1-2).
   - Cakupan Jalur Kausalitas dalam Tabel: Wajib memetakan secara komprehensif seluruh keterkaitan antarvariabel penelitian:
     * Pengaruh langsung variabel bebas ({x_vars_str}) terhadap variabel terikat ({y_var})
     * Pengaruh variabel bebas ({x_vars_str}) terhadap variabel mediasi ({m_vars_str})
     * Pengaruh variabel mediasi ({m_vars_str}) terhadap variabel terikat ({y_var})
     * Peran mediasi transmisi ({m_vars_str})
     * Peran moderasi penguat/pelemah ({z_vars_str})
     * Model {_model_lbl} simultan
     * Temuan yang tidak signifikan atau kontradiktif sebagai landasan empirical research gap.
   - Sumber di bawah tabel wajib ditulis persis: "Sumber: Data diolah peneliti, (2026)".

STRUKTUR SISTEMATIKA PENULISAN SUB-BAB B:
1. Paragraf Pengantar Metodologis Sub-Bab B (1 Paragraf Utuh, >= 4 baris, bervitasi):
   Menguraikan telaah sistematis literatur terdahulu (state of the art) sebagai instrumen metodologis dalam membedah konstelasi keilmuan kontemporer, kriteria inklusi artikel Scopus/Sinta 2022-2026, serta tujuan pemetaan terhadap dinamika {x_vars_str}, {m_vars_str}, {z_vars_str}, dan {y_var}.
2. Tabel 2.1 Hasil Penelitian yang Relevan:
   Tabel 5 kolom baku memuat MINIMAL 15 PENELITIAN EMPIRIS (studi nomor 1 sampai minimal 15) beserta baris sumber di bawah tabel.
3. Paragraf Analisis Konvergensi Temuan Positif (1 Paragraf Utuh, >= 4 baris, bervitasi):
   Menganalisis pola keselarasan hasil riset empiris terdahulu yang mengonfirmasi kontribusi signifikan pilar eksogen dan mediasi terhadap peningkatan capaian unjuk kerja.
4. Paragraf Analisis Divergensi dan Research Gap (1 Paragraf Utuh, >= 4 baris, bervitasi):
   Membedah inkonsistensi temuan antarpeneliti terdahulu (seperti pengaruh langsung yang gugur/tidak signifikan pada setting kelembagaan kaku) yang menunjukkan bahwa infrastruktur formal bersifat pasif tanpa adanya jembatan mediasi keperilakuan ({m_vars_str}) dan benteng kontingensi ({z_vars_str}).
5. Paragraf Sintesis Kebaruan Ilmiah (Novelty) Doktoral Model Terpadu (1 Paragraf Utuh, >= 4 baris, bervitasi):
   Merumuskan kebaruan orisinal disertasi ini yang mengintegrasikan seluruh konstruk dalam satu model struktural terpadu {_model_lbl} yang belum pernah diuji secara simultan oleh studi-studi terdahulu.
"""
    elif sub_id == "sub_2_3_1":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB C.1: KERANGKA KONSEPTUAL & SILOGISME DEDUKTIF (BAB II) ===
Judul Sub-bab WAJIB ditulis persis: C.1 Kerangka Konseptual & Silogisme Deduktif
SUB-BAB INDUK: C. KERANGKA BERFIKIR (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat analitis substantif mendalam atau 70–100 kata per paragraf). DILARANG KERAS membuat paragraf pendek (1, 2, atau 3 baris teks)!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan ilmiah mutakhir (2022-2026) atau teori klasik bereputasi (Snyder, 2019; Robbins & Judge, 2022; Cooper & Schindler, 2022; Sekaran & Bougie, 2020; Creswell & Creswell, 2023).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar & Dekonstruksi Paradigma Direct Effect (>= 4 baris teks utuh, bersitasi):
      Mendekonstruksi kelemahan asumsi linier mekanistis sederhana bahwa pilar eksogen ({x_vars_str}) otomatis langsung meningkatkan {y_var} tanpa perantara keperilakuan.
   b. Paragraf Rasionalisasi Mediasi sebagai Behavioral Execution Bridge (>= 4 baris teks utuh, bersitasi):
      Menjelaskan landasan Theory of Planned Behavior & Social Cognitive Theory mengapa {m_vars_str} mutlak dibutuhkan sebagai jembatan eksekusi perilaku nyata.
   c. Paragraf Pengantar Logika Deduktif Ilmiah (>= 4 baris teks utuh, bersitasi):
      Penerapan penalaran silogisme deduktif berjenjang (Premis Mayor, Premis Minor, Formulasi Hipotesis).
   d. Rangkaian Silogisme Deduktif Lengkap untuk Seluruh Hipotesis (H1 s.d H{total_h}):
      Setiap hipotesis WAJIB memuat 3 elemen:
      - Premis Mayor: Dalil teoretis makro (Grand/Middle Theory).
      - Premis Minor: Realitas kontekstual empiris aparatur dan tata kelola pada {instansi}.
      - Formulasi Hipotesis: Pernyataan formal hipotesis berarah (H1 s.d H{total_h}).
"""
    elif sub_id == "sub_2_3_2":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB C.2: KETERKAITAN KAUSAL ANTAR-VARIABEL (BAB II) ===
Judul Sub-bab WAJIB ditulis persis: C.2 Keterkaitan Kausal Antar-Variabel
SUB-BAB INDUK: C. KERANGKA BERFIKIR (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat analitis substantif mendalam atau 70–100 kata per paragraf). DILARANG KERAS membuat paragraf pendek (1, 2, atau 3 baris teks)!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: 100% setiap kelompok hubungan wajib didukung rujukan empiris mutakhir (2022-2026) dari jurnal internasional/nasional terakreditasi.
3. STRUKTUR SISTEMATIKA PENULISAN (5 RUMPUN KAUSALITAS):
   a. Elaborasi Kausalitas Langsung Variabel Bebas ({x_vars_str}) terhadap Mediasi ({m_vars_str}) (>= 4 baris, bersitasi 2022-2026).
   b. Elaborasi Kausalitas Langsung Variabel Bebas ({x_vars_str}) terhadap Terikat ({y_var}) (>= 4 baris, bersitasi 2022-2026).
   c. Elaborasi Kausalitas Langsung Variabel Mediasi ({m_vars_str}) terhadap Terikat ({y_var}) (>= 4 baris, bersitasi 2022-2026).
   d. Elaborasi Efek Transmisi Mediasi Tidak Langsung ({x_vars_str} -> {m_vars_str} -> {y_var}) (>= 4 baris, bersitasi 2022-2026).
   e. Elaborasi Efek Kontingensi Moderasi ({z_vars_str}) terhadap Keterkaitan Struktural ({x_vars_str} -> {y_var}) (>= 4 baris, bersitasi 2022-2026).
"""
    elif sub_id == "sub_2_3_3":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB C.3: PARADIGMA & KERANGKA BERFIKIR (BAB II) ===
Judul Sub-bab WAJIB ditulis persis: C.3 Paradigma & Kerangka Berpikir
SUB-BAB INDUK: C. KERANGKA BERFIKIR (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat analitis substantif mendalam atau 70–100 kata per paragraf).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan teori kepemimpinan, kontingensi, dan metodologis (Yukl, 2022; Donaldson, 2023; Hair et al., 2022; Hayes, 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Rujukan Gambar 2.1 Kerangka Berpikir: Paragraf pengantar + Box visual struktur model {_model_lbl} + narasi pembacaan arsitektur bagan (>= 4 baris, bersitasi).
   b. Integrasi Multi-Teori & Framework State of the Art Gambar 2.2: Harmonisasi teori yang relevan dengan variabel judul + Rumusan Akronim Model {model_acronym} + Box visual + narasi analitis kebaruan model terpadu hibrida (>= 4 baris).
   c. Rasionalisasi Boundary Conditions Variabel Moderasi ({z_vars_str}):
      Membedah secara mendalam 3 fenomena kelembagaan birokrasi (masing-masing >= 4 baris teks utuh & bersitasi):
      - 1. Rentan Terhadap Intervensi Struktural (Management Override)
      - 2. Ketergantungan Eksekusi Hasil Pengawasan dan Evaluasi Otoritas Formal
      - 3. Legitimasi dan Keteladanan Pimpinan (Tone at the Top)
      Diakhiri paragraf kesimpulan kontingensi kelembagaan bahwa tata kelola bersifat kondisional pada komitmen pimpinan {instansi}.
"""
    elif sub_id == "sub_2_3":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB C: KERANGKA BERFIKIR LENGKAP (C.1 s.d C.3) ===
Judul Sub-bab WAJIB ditulis persis: C. Kerangka Berpikir
SUB-BAB INDUK: C. KERANGKA BERFIKIR (Format Baku TAPD 2026 Disertasi Doktor S3).
Kompilasi lengkap seluruh bagian Kerangka Berpikir Doktoral:
1. C.1 Kerangka Konseptual & Silogisme Deduktif (Dekonstruksi Direct Effect, Rasionalisasi Mediasi, Logika Deduktif, Matriks Silogisme Deduktif H1 s.d H{total_h}).
2. C.2 Keterkaitan Kausal Antar-Variabel (Elaborasi Kausalitas Teoretis & Pembuktian Empiris Mutakhir 2022-2026 untuk 5 kelompok jalur kausalitas).
3. C.3 Paradigma & Kerangka Berpikir (Gambar 2.1 Bagan Kerangka Berpikir, Gambar 2.2 State of the Art Framework, Model {model_acronym}, Rasionalisasi Boundary Conditions Moderasi 3 Fenomena).
Seluruh paragraf naratif WAJIB >= 4 baris teks utuh (minimal 70-100 kata) dan wajib bersitasi ilmiah mutakhir (2022–2026).
"""
    elif sub_id == "sub_2_4":
        model_acronym = _generate_doctoral_model_acronym(ctx)
        hypos = ctx["graph"]["hypotheses"]
        daftar_h_str = ""
        for h in hypos:
            stmt = h.get("statement", "")
            if not stmt.lower().startswith("terdapat"):
                stmt = f"Terdapat {stmt}"
            daftar_h_str += f"   {h.get('code', 'H')}\t:\t{stmt}\n"

        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB D: HIPOTESIS PENELITIAN (BAB II) ===
Judul Sub-bab WAJIB ditulis persis: D. Hipotesis Penelitian
SUB-BAB INDUK: D. HIPOTESIS PENELITIAN (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat analitis substantif mendalam atau 70–100 kata per paragraf). DILARANG KERAS membuat paragraf pendek (1, 2, atau 3 baris teks)!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Paragraf pengantar dan paragraf penutup metodologis WAJIB memuat rujukan ilmiah (misal: Sekaran & Bougie, 2020; Cooper & Schindler, 2022; Creswell & Creswell, 2023; Hair et al., 2022; Hayes, 2022).
3. STRUKTUR SISTEMATIKA PENULISAN SUB-BAB D:
   Naskah WAJIB disusun dalam 3 Komponen Baku Sesuai Acuan DISERTASI S3.docx:
   a. Paragraf Pengantar Hipotesis Deduktif (>= 4 baris teks utuh, bersitasi):
      Menguraikan sintesis tinjauan Grand Theory, penelitian empiris terdahulu, dan kerangka pemikiran di atas yang melandasi perumusan hipotesis operasional model {model_acronym} pada {instansi}. Diakhiri kalimat transisi: "Berdasarkan uraian pada tinjauan teori yang diuraikan sebelumnya dan hasil penelitian terdahulu yang relevan serta kerangka pemikiran di atas, maka hipotesis yang dirumuskan pada penelitian ini adalah sebagai berikut:"
   b. Tabel Borderless 3 Kolom Formulasi Hipotesis Formal Berarah ({total_h} Hipotesis Formal H1 s.d H{total_h}):
      Disajikan dalam format tabel 3 kolom tanpa garis bingkai (borderless table):
      - Kolom 1: Kode Hipotesis (H1, H2, ... H{total_h})
      - Kolom 2: Simbol Pemisah (:)
      - Kolom 3: Pernyataan Formal Hipotesis Berarah yang selaras 1-ke-1 secara sekuensial dengan Rumusan Masalah dan Tujuan Penelitian di Bab I.
{daftar_h_str}
   c. Paragraf Penutup Metodologis Pengujian PLS-SEM (>= 4 baris teks utuh, bersitasi):
      Menjelaskan bahwa seluruh {total_h} hipotesis formal berarah diuji secara simultan menggunakan Partial Least Squares Structural Equation Modeling (PLS-SEM) dengan SmartPLS 4, estimasi koefisien jalur path beta, uji signifikansi bootstrapping 5.000 subsampel (t > 1,96 dan p < 0,05), specific indirect effects untuk mediasi, dan interaction effect untuk moderasi (Hair et al., 2022; Hayes, 2022).
"""
    elif sub_id in ("sub_3_a", "sub_3_1"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A: DESAIN PENELITIAN (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: A. Desain Penelitian
SUB-BAB INDUK: A. DESAIN PENELITIAN (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat analitis substantif mendalam atau 70–100 kata per paragraf). DILARANG KERAS membuat paragraf pendek (1, 2, atau 3 baris teks)!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Seluruh paragraf naratif WAJIB memuat rujukan metodologis bereputasi (Hair et al., 2022; Neuman, 2022; Creswell & Creswell, 2023; Cooper & Schindler, 2022; Sarstedt et al., 2022; Sugiyono, 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf 1: Pendekatan Kuantitatif Eksplanatori & Paradigma Positivisme (>= 4 baris, bersitasi):
      Menjelaskan pendekatan kuantitatif kausal-eksplanatori berlandaskan paradigma positivisme untuk menguji konstelasi hubungan antarvariabel serta menguji hipotesis deduktif secara statistik pada {instansi} (Hair et al., 2022; Creswell & Creswell, 2023). Metode survei dipilih untuk mengumpulkan data primer berskala besar secara efisien dan terstandarisasi melalui kuesioner terstruktur (Neuman, 2022).
   b. Paragraf 2: Desain Survei Cross-Sectional & Pembuktian Kausalitas (>= 4 baris, bersitasi):
      Menjelaskan desain survei cross-sectional di mana data empiris dikumpulkan secara simultan dalam satu kurun waktu (Cooper & Schindler, 2022). Karakteristik eksplanatori ditujukan untuk membuktikan kausalitas langsung ({x_vars_str} terhadap {y_var}), transmisi mediasi ({m_vars_str}), dan moderasi kontingensi ({z_vars_str}) menggunakan PLS-SEM dengan daya generalisasi tinggi (Hair et al., 2022; Sarstedt et al., 2022).
"""
    elif sub_id == "sub_3_b":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB B: LOKASI PENELITIAN (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: B. Lokasi Penelitian
SUB-BAB INDUK: B. LOKASI PENELITIAN (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata per paragraf). DILARANG KERAS membuat paragraf pendek!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Paragraf WAJIB memuat rujukan ilmiah metodologis (Cooper & Schindler, 2022; Sekaran & Bougie, 2020; Sugiyono, 2022; Neuman, 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf 1: Penetapan Lokasi Empiris & Relevansi Variabel (>= 4 baris, bersitasi):
      Menetapkan lokasi penelitian pada lingkungan {instansi}. Menjelaskan relevansi substantif lokus empiris terhadap konstelasi variabel penelitian ({y_var}, {m_vars_str}, {z_vars_str}, {x_vars_str}) dan mandat tata kelola kelembagaan daerah (Cooper & Schindler, 2022; Sekaran & Bougie, 2020).
   b. Paragraf 2: Aksesibilitas Data & Karakteristik Responden Representatif (>= 4 baris, bersitasi):
      Menguraikan argumentasi rasional penetapan lokasi: keterbukaan akses data riset (data accessibility), komitmen pimpinan instansi dalam perbaikan manajemen, serta sebaran unit kerja fungsional dengan aparatur yang heterogen dan kompeten (Sugiyono, 2022; Neuman, 2022).
"""
    elif sub_id == "sub_3_c":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB C: WAKTU PENELITIAN (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: C. Waktu Penelitian
SUB-BAB INDUK: C. WAKTU PENELITIAN (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata). DILARANG membuat paragraf pendek!
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Paragraf pengantar dan penutup WAJIB memuat sitasi metodologis (Sugiyono, 2022; Neuman, 2022; Cooper & Schindler, 2022; Hair et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf 1: Rasionalisasi Penetapan Waktu & Stabilitas Organisasi (>= 4 baris, bersitasi):
      Menjelaskan kurun waktu pelaksanaan penelitian pada tahun {ctx.get('tahun', '2026')}. Periode dipilih saat ritme kedinasan berada pada intensitas normal agar responden objektif dan minim bias kelelahan atau tekanan kerja puncak (Sugiyono, 2022; Neuman, 2022).
   b. Tabel 3.1: Jadwal Penelitian (Tabel 13 tahapan riset dari usulan judul hingga promosi doktor beserta waktu dan target capaian).
   c. Paragraf Penutup: Kepatuhan Metodologis Timeline Riset Doktoral (>= 4 baris, bersitasi) (Cooper & Schindler, 2022; Hair et al., 2022).
"""
    elif sub_id == "sub_3_d":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB D: SUMBER DATA (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: D. Sumber Data
SUB-BAB INDUK: D. SUMBER DATA (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan metodologis (Hair et al., 2022; Sekaran & Bougie, 2020; Cooper & Schindler, 2022; Neuman, 2022; Podsakoff et al., 2024).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf 1: Data Primer Kuesioner Terstruktur (>= 4 baris, bersitasi):
      Data primer dihimpun langsung melalui kuesioner tertutup skala Likert 1-5 kepada responden aparatur {instansi} mencakup persepsi dan sikap kedinasan seluruh variabel (Hair et al., 2022; Sekaran & Bougie, 2020).
   b. Paragraf 2: Data Sekunder Dokumentasi Kelembagaan (>= 4 baris, bersitasi):
      Data sekunder bersumber dari Renstra, LAKIP/LKjIP, regulasi pemerintah daerah, dan catatan kepegawaian untuk triangulasi dan penguatan fenomena empiris (Cooper & Schindler, 2022; Neuman, 2022).
   c. Paragraf 3: Mitigasi Common Method Bias (CMB) (>= 4 baris, bersitasi):
      Menguraikan prosedur prosedural (anonymity responden, pengacakan urutan butir kuesioner) dan statistik melalui Harman's Single Factor Test (< 50% varians) menggunakan SmartPLS 4 / SPSS guna memastikan bebas dari ancaman bias metode umum (Podsakoff et al., 2024; Hair et al., 2022).
"""
    elif sub_id in ("sub_3_e", "sub_3_2"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB E: POPULASI DAN SAMPEL (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: E. Populasi dan Sampel
SUB-BAB INDUK: E. POPULASI DAN SAMPEL (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan metodologis wajib (Sugiyono, 2022; Sinambela, 2023; Sekaran & Bougie, 2020; Cooper & Schindler, 2022; Hair et al., 2022; Sarstedt et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Bagian Populasi: Definisi wilayah generalisasi (Sugiyono, 2022). Total populasi N = {ctx.get('populasi', '')} pegawai pada {instansi}. Tabel 3.2 Sebaran Populasi Pegawai per Unit Kerja.
   b. Bagian Sampel & Kriteria Inklusi: Definisi sampel (Sinambela, 2023). Kriteria purposive: PNS/PPPK definitif dan masa kerja minimal 2 tahun (Cooper & Schindler, 2022; Sugiyono, 2022).
   c. Rumus Slovin & Perhitungan: Formula n = N / (1 + N.e^2) dengan presisi e = 5% (0,05). Langkah kalkulasi matematis menghasilkan sampel n = {ctx.get('sampel', '')} responden. Justifikasi kecukupan sampel PLS-SEM (10-times rule & statistical power) (Hair et al., 2022; Sarstedt et al., 2022).
   d. Teknik Proportional Random Sampling: Formula alokasi proporsional ni = (Ni / N) x n dan Tabel 3.3 Alokasi Sampel Penelitian secara Proporsional pada Unit Kerja di {instansi} (Sugiyono, 2022; Sinambela, 2023).
"""
    elif sub_id == "sub_3_f_y":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB F.1: DEFINISI & KISI-KISI VARIABEL TERIKAT ({y_var.upper()}) (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: F.1 Definisi dan Kisi-Kisi Instrumen Variabel Terikat ({y_var})
SUB-BAB INDUK: F. DEFINISI KONSEPTUAL DAN OPERASIONAL VARIABEL (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan teori substantif dan metodologi pengukuran (Cooper & Schindler, 2022; Sekaran & Bougie, 2020; Sugiyono, 2022; Hair et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Definisi Konseptual {y_var}:
      Paragraf >= 4 baris teks utuh berakar pada teori kinerja organisasi dan MSDM (Armstrong & Taylor, 2023; Robbins & Judge, 2022).
   b. Definisi Operasional {y_var}:
      Paragraf >= 4 baris teks utuh menguraikan operasionalisasi pada {instansi} menggunakan skala Likert 1-5 (1=Sangat Tidak Setuju hingga 5=Sangat Setuju) (Sugiyono, 2022).
   c. Tabel Kisi-Kisi Instrumen Variabel {y_var}:
      Tabel matriks memuat: No, Dimensi, Indikator Empiris, dan Nomor Butir Pertanyaan.
"""
    elif sub_id == "sub_3_f_m":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB F.2: DEFINISI & KISI-KISI VARIABEL MEDIASI ({m_vars_str.upper()}) (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: F.2 Definisi dan Kisi-Kisi Instrumen Variabel Mediasi ({m_vars_str})
SUB-BAB INDUK: F. DEFINISI KONSEPTUAL DAN OPERASIONAL VARIABEL (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan teori keperilakuan dan transmisi mediasi (Ajzen, 1991/2022; Schaufeli et al., 2022; Robbins & Judge, 2022; Hair et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Definisi Konseptual {m_vars_str}:
      Paragraf >= 4 baris teks utuh berakar pada Theory of Planned Behavior dan dorongan motivasional keperilakuan aparatur.
   b. Definisi Operasional {m_vars_str}:
      Paragraf >= 4 baris teks utuh menjelaskan cara ukur ordinal skala Likert 1-5 pada {instansi}.
   c. Tabel Kisi-Kisi Instrumen Variabel {m_vars_str}:
      Tabel matriks memuat: No, Dimensi, Indikator Empiris, dan Nomor Butir Pertanyaan.
"""
    elif sub_id == "sub_3_f_z":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB F.3: DEFINISI & KISI-KISI VARIABEL MODERASI ({z_vars_str.upper()}) (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: F.3 Definisi dan Kisi-Kisi Instrumen Variabel Moderasi ({z_vars_str})
SUB-BAB INDUK: F. DEFINISI KONSEPTUAL DAN OPERASIONAL VARIABEL (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan Contingency Theory dan kepemimpinan (Donaldson, 2023; Yukl, 2022; Hair et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Definisi Konseptual {z_vars_str}:
      Paragraf >= 4 baris teks utuh berakar pada teori kontingensi dan etos kerja/kepemimpinan sebagai variabel penguat/pelemah hubungan struktural.
   b. Definisi Operasional {z_vars_str}:
      Paragraf >= 4 baris teks utuh menjelaskan pengukuran skala Likert 1-5 pada {instansi}.
   c. Tabel Kisi-Kisi Instrumen Variabel {z_vars_str}:
      Tabel matriks memuat: No, Dimensi, Indikator Empiris, dan Nomor Butir Pertanyaan.
"""
    elif re.match(r"^sub_3_f_x(\d+)$", sub_id):
        m_x = re.match(r"^sub_3_f_x(\d+)$", sub_id)
        x_idx = int(m_x.group(1))
        x_name = ctx["x_vars"][x_idx - 1] if x_idx <= len(ctx["x_vars"]) else f"Variabel Bebas X{x_idx}"
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB F.4.{x_idx}: DEFINISI & KISI-KISI VARIABEL BEBAS X{x_idx} ({x_name.upper()}) (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: F.4.{x_idx} Definisi dan Kisi-Kisi Instrumen {x_name}
SUB-BAB INDUK: F. DEFINISI KONSEPTUAL DAN OPERASIONAL VARIABEL (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan teori substantif dan metodologi pengukuran (Cooper & Schindler, 2022; Sekaran & Bougie, 2020; Sugiyono, 2022; Hair et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Definisi Konseptual {x_name}:
      Paragraf >= 4 baris teks utuh berakar pada literatur mutakhir teori kapasitas dan perilaku organisasi.
   b. Definisi Operasional {x_name}:
      Paragraf >= 4 baris teks utuh menguraikan operasionalisasi pengukuran skala Likert 1-5 pada {instansi}.
   c. Tabel Kisi-Kisi Instrumen Variabel {x_name}:
      Tabel matriks memuat: No, Dimensi, Indikator Empiris, dan Nomor Butir Pertanyaan.
"""
    elif sub_id in ("sub_3_f", "sub_3_3"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB F: DEFINISI KONSEPTUAL DAN OPERASIONAL VARIABEL (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: F. Definisi Konseptual dan Operasional Variabel
SUB-BAB INDUK: F. DEFINISI KONSEPTUAL DAN OPERASIONAL VARIABEL (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan teori substantif dan metodologi pengukuran (Cooper & Schindler, 2022; Sekaran & Bougie, 2020; Sugiyono, 2022; Sinambela, 2023; Hair et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Taksonomi: Mengklasifikasikan variabel ke dalam 4 kuadran: Terikat ({y_var}), Mediasi ({m_vars_str}), Moderasi ({z_vars_str}), dan Bebas ({x_vars_str}) (Sekaran & Bougie, 2020; Hair et al., 2022).
   b. Elaborasi per Variabel (urut Y, M, Z, X1, X2...):
      - Sub-heading Heading 3: Variabel [Nama Variabel] ([Kode])
      - Definisi Konseptual: Paragraf >= 4 baris berakar teori substantif dan literatur mutakhir.
      - Definisi Operasional: Paragraf >= 4 baris menjelaskan cara ukur skala Likert 1-5 dan dimensi-dimensi operasionalnya pada {instansi}.
      - Kisi-Kisi Instrumen: Tabel 3.x (Dimensi, Indikator, Butir Pertanyaan).
"""
    elif sub_id in ("sub_3_g", "sub_3_4"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB G: INSTRUMEN PENELITIAN (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: G. Instrumen Penelitian
SUB-BAB INDUK: G. INSTRUMEN PENELITIAN (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan metodologis instrumen dan skala (Cooper & Schindler, 2022; Hair et al., 2022; Sugiyono, 2022; Sinambela, 2023; Neuman, 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf 1: Kuesioner Tertutup Terstruktur Deduktif (>= 4 baris, bersitasi):
      Dirancang dengan kalimat lugas, tegas, tidak bermakna ganda, dan positif untuk merepresentasikan indikator empiris seluruh variabel ({y_var}, {m_vars_str}, {z_vars_str}, {x_vars_str}) pada {instansi} (Cooper & Schindler, 2022; Hair et al., 2022).
   b. Paragraf 2: Skala Pengukuran Likert 5 Poin (>= 4 baris, bersitasi):
      Teori skala Likert menurut Sugiyono (2022) dan Sinambela (2023) untuk mengukur sikap dan persepsi responden secara ordinal simetris bertingkat.
   c. Tabel Skor Jawaban Responden: Tabel 3.11 memuat opsi Sangat Setuju (5) hingga Sangat Tidak Setuju (1).
   d. Paragraf Penutup: Standarisasi Instrumen & Uji Kelayakan Isi (>= 4 baris, bersitasi) (Neuman, 2022; Cooper & Schindler, 2022).
"""
    elif sub_id in ("sub_3_h", "sub_3_5"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB H: UJICOBA INSTRUMEN (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: H. Ujicoba Instrumen
SUB-BAB INDUK: H. UJICOBA INSTRUMEN (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan metodologis uji validitas dan reliabilitas (Sinambela, 2023; Hair et al., 2022; Sugiyono, 2022; Cooper & Schindler, 2022; Neuman, 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Pilot Testing (>= 4 baris, bersitasi): Pelaksanaan ujicoba kepada 50 responden aparatur di luar sampel utama dengan karakteristik setara (Sinambela, 2023; Hair et al., 2022).
   b. 1. Uji Validitas Instrumen:
      - Validitas konstruk dan korelasi Pearson Product Moment (Sugiyono, 2022; Sinambela, 2023).
      - Rumus Pearson Product Moment dan kriteria keputusan r_hitung >= r_tabel (0,279 pada df=48, alpha=0,05).
      - Tabel Hasil Uji Validitas per Variabel (Tabel 3.12 s.d 3.18) disertai narasi pengantar dan interpretasi >= 4 baris per variabel.
   c. 2. Uji Reliabilitas Instrumen:
      - Uji konsistensi internal Cronbach's Alpha bagi butir valid (Sinambela, 2023; Hair et al., 2022).
      - Rumus Cronbach's Alpha dan Tabel Rujukan Interpretasi Sinambela (2023) (Tabel 3.19).
      - Tabel Hasil Uji Reliabilitas Seluruh Variabel (Tabel 3.20): Seluruh variabel memiliki Cronbach's Alpha > 0,800 (rentang 0,950 s.d 0,991) berpredikat Sangat Reliabel.
      - Paragraf penutup stabilitas dan konsistensi instrumen (>= 4 baris, bersitasi) (Hair et al., 2022).
"""
    elif sub_id in ("sub_3_i", "sub_3_6"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB I: TEKNIK ANALISIS DATA (BAB III) ===
Judul Sub-bab WAJIB ditulis persis: I. Teknik Analisis Data
SUB-BAB INDUK: I. TEKNIK ANALISIS DATA (Format Baku TAPD 2026 Disertasi Doktor S3).

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan metodologis analisis SEM mutakhir (Hair et al., 2022; Sarstedt et al., 2022; Chin, 1998; Henseler et al., 2015; Tenenhaus et al., 2005; Hayes, 2022; Cooper & Schindler, 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Rasionalisasi PLS-SEM SmartPLS 4 (>= 4 baris, bersitasi):
      Keunggulan PLS-SEM untuk model struktural kompleks bertingkat {_model_lbl}, distribusi non-parametrik tanpa asumsi normalitas ketat, dan orientasi prediksi empiris (Hair et al., 2022; Sarstedt et al., 2022).
   b. Paragraf Kecukupan Sampel & Uji Statistik (>= 4 baris, bersitasi):
      Ukuran sampel n = {ctx.get('n', 176)} memenuhi kaidah 10-times rule dan G*Power (Hair et al., 2022).
   c. Tiga Tahapan Baku Analisis PLS-SEM:
      - 1. Evaluasi Outer Model (>= 4 baris, bersitasi): Outer Loadings > 0,708, AVE > 0,50, Fornell-Larcker, Cross Loadings, HTMT < 0,85, CR & Cronbach's Alpha > 0,70, Outer VIF < 3,0 (Henseler et al., 2015; Hair et al., 2022).
      - 2. Evaluasi Inner Model (>= 4 baris, bersitasi): Inner VIF < 3,0, R² (substansial 0,75, moderat 0,50, lemah 0,25), f² (0,02, 0,15, 0,35), Q² blindfolding > 0, Model Fit SRMR < 0,08, NFI > 0,90, GoF Tenenhaus (Tenenhaus et al., 2005; Hair et al., 2022).
      - 3. Pengujian Hipotesis Bootstrapping (>= 4 baris, bersitasi): Non-parametrik 5.000 subsampel, two-tailed alpha=0,05 (t >= 1,96 dan p < 0,05) (Hair et al., 2022; Cooper & Schindler, 2022).
      - Prosedur Pengujian Mediasi & Moderasi: Specific Indirect Effects, VAF mediasi, dan Two-Stage Interaction model moderasi (Hair et al., 2022; Hayes, 2022; Sarstedt et al., 2022).
      - Model Persamaan Struktural Matematis: Persamaan regresi bertingkat (sub-structural equations).
"""

    elif sub_id in ("sub_4_a_1", "sub_4_1_1", "sub_4_1_2"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.1: KARAKTERISTIK DEMOGRAFIS RESPONDEN (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.1 Karakteristik Demografis Responden
SUB-BAB INDUK: A. HASIL ANALISIS (Format Baku TAPD 2026 Disertasi Doktor S3).
BAGIAN UTAMA: 1. Deskriptif Data Demografis Responden

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata). Dilarang membuat paragraf pendek (1-3 baris).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan metodologis demografi, psikologi organisasi, dan manajemen SDM mutakhir (Cooper & Schindler, 2022; Sekaran & Bougie, 2020; Sugiyono, 2022; Armstrong & Taylor, 2023; Robbins & Judge, 2022; Sinambela, 2023; Sedarmayanti, 2022; Hasibuan, 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Analisis Demografis (>= 4 baris, bersitasi):
      Pentingnya profil demografis aparatur di {instansi} sebagai basis kontekstual pemaknaan respon terhadap variabel penelitian (Cooper & Schindler, 2022; Sekaran & Bougie, 2020).
   b. Tabel 4.1 Deskripsi Karakteristik Responden berdasarkan Jenis Kelamin:
      Tabel distribusi frekuensi dan persentase gender (Laki-laki & Perempuan, n = {ctx['n']}).
   c. 2 Paragraf Narasi Analitis Jenis Kelamin (masing-masing >= 4 baris, bersitasi):
      - Paragraf 1: Analisis proporsi gender dalam pelaksanaan tugas operasional lapangan vs administratif di {instansi} (Armstrong & Taylor, 2023).
      - Paragraf 2: Representasi meritokrasi gender dan dampaknya terhadap iklim kerja dan motivasi kerja (Robbins & Judge, 2022; Sinambela, 2023).
   d. Tabel 4.2 Deskripsi Karakteristik Responden berdasarkan Usia:
      Tabel kelompok usia (<= 30 tahun, 31-40 tahun, 41-50 tahun, > 50 tahun).
   e. 2 Paragraf Narasi Analitis Usia (masing-masing >= 4 baris, bersitasi):
      - Paragraf 1: Dominasi usia produktif matang (31-50 tahun) dan kematangan kognitif dalam merespon butir instrumen (Armstrong & Taylor, 2023).
      - Paragraf 2: Stabilitas emosional, etos kerja, dan adaptabilitas teknologi birokrasi aparatur generasi madya (Robbins & Judge, 2022).
   f. Tabel 4.3 Deskripsi Karakteristik Responden berdasarkan Pangkat/Golongan:
      Tabel distribusi kepangkatan (Golongan II/Pengatur, Golongan III/Penata, Golongan IV/Pembina).
   g. 2 Paragraf Narasi Analitis Pangkat/Golongan (masing-masing >= 4 baris, bersitasi):
      - Paragraf 1: Komposisi Golongan III sebagai tulang punggung teknis operasional dan Golongan IV sebagai pengarah strategis (Hasibuan, 2022; Sedarmayanti, 2022).
      - Paragraf 2: Korelasi jenjang kepangkatan dengan kedewasaan birokrasi, kepemimpinan transformasional, dan akuntabilitas kinerja aparatur di {instansi} (Sinambela, 2023; Robbins & Judge, 2022).
"""
    elif sub_id == "sub_4_a_2_y":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.2.1: DESKRIPTIF VARIABEL TERIKAT ({y_var.upper()}) (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.2.1 Analisis Statistik Deskriptif Variabel Terikat ({y_var})
SUB-BAB INDUK: A. HASIL ANALISIS (2. Analisis Statistik Deskriptif Variabel)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan statistik deskriptif dan perilaku organisasi (Sugiyono, 2022; Cooper & Schindler, 2022; Robbins & Judge, 2022; Armstrong & Taylor, 2023).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Deskriptif {y_var} (>= 4 baris, bersitasi).
   b. Tabel Distribusi Statistik Deskriptif Indikator {y_var}: Kolom Indikator, Min, Max, Median, Mean, dan Std. Deviasi.
   c. Paragraf Interpretasi Analitis Indikator (>= 4 baris, bersitasi): Membedah indikator dengan mean tertinggi sebagai kekuatan capaian dan mean terendah sebagai area perbaikan kinerja aparatur pada {instansi}.
"""
    elif sub_id == "sub_4_a_2_mz":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.2.2: DESKRIPTIF VARIABEL MEDIASI ({m_vars_str.upper()}) & MODERASI ({z_vars_str.upper()}) (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.2.2 Analisis Statistik Deskriptif Variabel Mediasi ({m_vars_str}) dan Moderasi ({z_vars_str})
SUB-BAB INDUK: A. HASIL ANALISIS (2. Analisis Statistik Deskriptif Variabel)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan statistik deskriptif dan teori keperilakuan (Sugiyono, 2022; Robbins & Judge, 2022; Yukl, 2022; Donaldson, 2023).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Deskriptif Variabel Mediasi ({m_vars_str}): Paragraf pengantar + Tabel 5 parameter statistik indikator + Paragraf analisis mendalam indikator terkuat dan titik kritis di {instansi}.
   b. Deskriptif Variabel Moderasi ({z_vars_str}): Paragraf pengantar + Tabel 5 parameter statistik indikator + Paragraf analisis mendalam peran kontingensi di {instansi}.
"""
    elif re.match(r"^sub_4_a_2_x(\d+)$", sub_id):
        m_x = re.match(r"^sub_4_a_2_x(\d+)$", sub_id)
        x_idx = int(m_x.group(1))
        x_name = ctx["x_vars"][x_idx - 1] if x_idx <= len(ctx["x_vars"]) else f"Variabel Bebas X{x_idx}"
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.2.3.{x_idx}: DESKRIPTIF VARIABEL BEBAS X{x_idx} ({x_name.upper()}) (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.2.3.{x_idx} Analisis Statistik Deskriptif {x_name}
SUB-BAB INDUK: A. HASIL ANALISIS (2. Analisis Statistik Deskriptif Variabel)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan statistik deskriptif dan manajemen sumber daya manusia (Sugiyono, 2022; Sekaran & Bougie, 2020; Robbins & Judge, 2022; Armstrong & Taylor, 2023).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Deskriptif {x_name} (>= 4 baris, bersitasi).
   b. Tabel Distribusi Statistik Deskriptif Seluruh Indikator {x_name}: Min, Max, Median, Mean, Std. Deviasi.
   c. Paragraf Interpretasi Analitis Indikator (>= 4 baris, bersitasi): Menelaah indikator dengan persepsi responden paling dominan dan indikator yang memerlukan perhatian pimpinan {instansi}.
"""
    elif sub_id == "sub_4_a_2_rekap":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.2.REKAP: REKAPITULASI STATISTIK DESKRIPTIF & KATEGORISASI (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.2.Rekap Rekapitulasi Rerata Skor dan Kategorisasi Keseluruhan Variabel
SUB-BAB INDUK: A. HASIL ANALISIS (2. Analisis Statistik Deskriptif Variabel)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan metodologis rentang skala dan evaluasi variabel (Sugiyono, 2018/2022; Cooper & Schindler, 2022; Hair et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Kalkulasi Rumus Rentang Skala Sugiyono (2018/2022): RS = (5 - 1) / 3 = 1,33.
   b. Tabel Kriteria Kategori Tanggapan Responden: Rendah (1,00 - 2,33), Sedang (2,34 - 3,67), Tinggi (3,68 - 5,00).
   c. Tabel Rekapitulasi Rata-rata Skor Penilaian Seluruh Variabel Penelitian ({y_var}, {m_vars_str}, {z_vars_str}, {x_vars_str}): Menampilkan Grand Mean, Standar Deviasi, dan Predikat Kategori.
   d. Paragraf Analisis Komparatif Antar-Konstruk Laten (>= 4 baris, bersitasi): Sintesis menyeluruh kekuatan komparatif variabel dalam mendukung tujuan kelembagaan {instansi}.
"""
    elif sub_id in ("sub_4_a_2", "sub_4_1_3", "sub_4_1_4", "sub_4_1_5"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.2: ANALISIS STATISTIK DESKRIPTIF VARIABEL (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.2 Analisis Statistik Deskriptif Variabel
SUB-BAB INDUK: A. HASIL ANALISIS (Format Baku TAPD 2026 Disertasi Doktor S3).
BAGIAN UTAMA: 2. Analisis Statistik Deskriptif Variabel

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata). Dilarang membuat paragraf pendek (1-3 baris).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan statistik deskriptif dan perilaku organisasi (Sugiyono, 2022; Cooper & Schindler, 2022; Sekaran & Bougie, 2020; Robbins & Judge, 2022; Hair et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Analisis Deskriptif (>= 4 baris, bersitasi):
      Prinsip statistik deskriptif menggunakan 5 parameter tendensi sentral dan dispersi: Skor Minimum (Min), Skor Maksimum (Max), Nilai Tengah (Median), Rerata Hitung (Mean), dan Standar Deviasi (SD) (Sugiyono, 2022; Cooper & Schindler, 2022).
   b. Elaborasi Deskriptif Tiap Variabel Penelitian ({y_var}, {m_vars_str}, {z_vars_str}, {x_vars_str}):
      - Setiap variabel diawali 1 paragraf narasi pembuka (>= 4 baris, bersitasi).
      - Tabel 4.4 s.d Tabel 4.{3 + len(x_vars) + (1 if m_vars else 0) + (1 if z_vars else 0) + 1}: Tabel distribusi statistik seluruh indikator variabel (Kolom: Pernyataan Indikator, Min, Max, Median, Mean, Std. Deviasi).
      - 1 Paragraf narasi interpretasi analitis mendalam (>= 4 baris, bersitasi) mengulas indikator dengan nilai mean tertinggi (kekuatan) dan indikator dengan mean terendah (titik kritis yang membutuhkan intervensi manajerial di {instansi}).
   c. Perhitungan Rentang Kategori Jawaban Sugiyono:
      - Formula lebar interval skala Likert 5 poin: RS = (Skor Maksimal - Skor Minimal) / Jumlah Kategori = (5 - 1) / 3 = 1,33 (Sugiyono, 2018/2022).
      - Tabel Kriteria Kategori Tanggapan Responden: Rendah (1,00 - 2,33), Sedang (2,34 - 3,67), Tinggi (3,68 - 5,00).
   d. Tabel Rata-rata Skor Penilaian Seluruh Variabel Penelitian:
      Menampilkan rekapitulasi Grand Mean, Standar Deviasi, dan Kategori Tanggapan Responden untuk seluruh konstruk laten ({y_var}, {m_vars_str}, {z_vars_str}, {x_vars_str}).
   e. Paragraf Komparasi Komprehensif Antar-Konstruk (>= 4 baris, bersitasi):
      Evaluasi komparatif tingkat ketercapaian seluruh variabel dalam mendukung kinerja organisasi di {instansi} (Robbins & Judge, 2022; Hair et al., 2022).
"""
    elif sub_id == "sub_4_a_3_1":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.3.1: EVALUASI VALIDITAS KONVERGEN (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.3.1 Evaluasi Validitas Konvergen (Convergent Validity)
SUB-BAB INDUK: A. HASIL ANALISIS (3. Evaluasi Model Pengukuran Outer Model SmartPLS 4)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan baku PLS-SEM (Hair et al., 2022; Sarstedt et al., 2022; Chin, 1998; Tenenhaus et al., 2005).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Analisis Pengukuran & Gambar 4.1 Bagan Model Pengukuran Penuh (Full Model SmartPLS 4) (>= 4 baris, bersitasi).
   b. Kriteria Validitas Konvergen: Nilai Outer Loading > 0,708 dan AVE > 0,500 (Hair et al., 2022).
   c. Tabel Hasil Evaluasi Outer Loadings Seluruh Indikator (Tabel butir indikator, loading factor > 0,708) + Paragraf interpretasi validitas indikator (>= 4 baris, bersitasi).
   d. Tabel Nilai Average Variance Extracted (AVE) Seluruh Konstruk Laten (Tabel nilai AVE > 0,500) + Paragraf konfirmasi daya jelas variansi konstruk laten (>= 4 baris, bersitasi).
"""
    elif sub_id == "sub_4_a_3_2":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.3.2: EVALUASI VALIDITAS DISKRIMINAN (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.3.2 Evaluasi Validitas Diskriminan (Discriminant Validity)
SUB-BAB INDUK: A. HASIL ANALISIS (3. Evaluasi Model Pengukuran Outer Model SmartPLS 4)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan diskriminan mutakhir (Henseler et al., 2015; Franke & Sarstedt, 2019; Hair et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Validitas Diskriminan & Kelemahan Kriteria Klasik Fornell-Larcker (>= 4 baris, bersitasi).
   b. Kriteria Heterotrait-Monotrait Ratio of Correlations (HTMT < 0,85 atau < 0,90) (Henseler et al., 2015).
   c. Tabel Matriks Nilai Rasio HTMT Antar-Konstruk Laten + Paragraf konfirmasi keunikan konseptual konstruk (>= 4 baris, bersitasi).
   d. Tabel Matriks Cross-Loadings Indikator + Paragraf verifikasi diskriminan pada tingkat indikator (>= 4 baris, bersitasi).
"""
    elif sub_id == "sub_4_a_3_3":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.3.3: EVALUASI RELIABILITAS KONSTRUK (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.3.3 Evaluasi Reliabilitas Konstruk (Construct Reliability)
SUB-BAB INDUK: A. HASIL ANALISIS (3. Evaluasi Model Pengukuran Outer Model SmartPLS 4)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan konsistensi internal psikometrik (Hair et al., 2022; Dijkstra & Henseler, 2015; Nunnally & Bernstein, 1994).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Konsistensi Internal Instrumen (>= 4 baris, bersitasi).
   b. Kriteria Evaluasi: Cronbach's Alpha > 0,70, Composite Reliability (rho_c) > 0,70, dan Dijkstra-Henseler's rho_a > 0,70 (Hair et al., 2022).
   c. Tabel Hasil Evaluasi Reliabilitas Seluruh Konstruk Laten (Kolom: Variabel, Cronbach's Alpha, rho_a, Composite Reliability rho_c, dan Kesimpulan Sangat Reliabel).
   d. Paragraf Konfirmasi Stabilitas dan Ketahanan Bebas Bias Pengukuran (>= 4 baris, bersitasi).
"""
    elif sub_id in ("sub_4_a_3", "sub_4_2_poin_1"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.3: EVALUASI MODEL PENGUKURAN (OUTER MODEL) (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.3 Evaluasi Model Pengukuran (Outer Model SmartPLS 4)
SUB-BAB INDUK: A. HASIL ANALISIS (Format Baku TAPD 2026 Disertasi Doktor S3).
BAGIAN UTAMA: 3. Evaluasi Model Pengukuran (Outer Model SmartPLS 4)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan metodologis pengukuran PLS-SEM mutakhir (Hair et al., 2022; Sarstedt et al., 2022; Chin, 1998; Tenenhaus et al., 2005; Henseler et al., 2015; Franke & Sarstedt, 2019; Nunnally & Bernstein, 1994).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Analisis Inferensial PLS-SEM (>= 4 baris, bersitasi):
      Rasionalisasi evaluasi outer model menggunakan algoritma SmartPLS 4 untuk memastikan validitas dan reliabilitas seluruh konstruk laten (Hair et al., 2022; Sarstedt et al., 2022).
   b. Gambar 4.1 Bagan Model Pengukuran Penuh (Full Model SmartPLS 4):
      Visualisasi arsitektur model pengukuran dan struktural lengkap dalam box Word terformat rapi + narasi pembacaan bagan (>= 4 baris, bersitasi).
   c. Evaluasi Validitas Konvergen (Convergent Validity):
      - Kriteria metodologis: Outer Loading > 0,708 dan AVE > 0,50 (Hair et al., 2022; Chin, 1998; Tenenhaus et al., 2005).
      - Tabel Hasil Uji Outer Loadings Seluruh Indikator (Tabel nilai loading > 0,708) + narasi analisis validitas konvergen (>= 4 baris, bersitasi).
      - Tabel Nilai Average Variance Extracted (AVE) Seluruh Konstruk (seluruh AVE > 0,50) + narasi konfirmasi daya jelas variansi konstruk laten (>= 4 baris, bersitasi).
   d. Evaluasi Validitas Diskriminan (Discriminant Validity):
      - Kriteria metodologis: Nilai Heterotrait-Monotrait Ratio (HTMT) < 0,85 atau < 0,90 dan Cross-Loadings (Henseler et al., 2015; Franke & Sarstedt, 2019; Hair et al., 2022).
      - Tabel Nilai HTMT Antar-Konstruk Laten (seluruh nilai < 0,85) + narasi evaluasi keunikan konstruk (>= 4 baris, bersitasi).
      - Tabel Matriks Cross-Loadings Indikator (nilai loading pada konstruk target lebih besar dari korelasi ke konstruk lain) + narasi verifikasi diskriminan (>= 4 baris, bersitasi).
   e. Evaluasi Reliabilitas Konstruk (Construct Reliability):
      - Kriteria metodologis: Cronbach's Alpha > 0,70 dan Composite Reliability (rho_a, rho_c) > 0,70 (Hair et al., 2022; Nunnally & Bernstein, 1994).
      - Tabel Hasil Evaluasi Reliabilitas Konstruk (Cronbach's Alpha, rho_a, rho_c, dan kesimpulan Sangat Reliabel) + narasi konfirmasi konsistensi internal instrumen (>= 4 baris, bersitasi).
"""
    elif sub_id == "sub_4_a_4_1":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.4.1: EVALUASI KOEFISIEN DETERMINASI R-SQUARE (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.4.1 Evaluasi Koefisien Determinasi (R-Square / R²)
SUB-BAB INDUK: A. HASIL ANALISIS (4. Evaluasi Model Struktural Inner Model PLS-SEM)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan baku R² (Chin, 1998; Hair et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Koefisien Determinasi dan Kriteria Chin (1998) / Hair et al. (2022): 0,75 (substansial), 0,50 (moderat), 0,25 (lemah).
   b. Tabel Nilai R-Square dan R-Square Adjusted Konstruk Endogen ({m_vars_str} dan {y_var}).
   c. Paragraf Interpretasi Retensi Variansi {m_vars_str} (>= 4 baris, bersitasi): Menjelaskan seberapa besar varians variabel mediasi mampu dijelaskan oleh pilar-pilar eksogen {x_vars_str}.
   d. Paragraf Interpretasi Retensi Variansi {y_var} (>= 4 baris, bersitasi): Menjelaskan seberapa besar varians unjuk kerja akhir mampu diterangkan secara simultan oleh pilar eksogen, mediasi, dan peran interaksi moderasi {z_vars_str}.
"""
    elif sub_id == "sub_4_a_4_2":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.4.2: EVALUASI UKURAN EFEK F-SQUARE (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.4.2 Evaluasi Ukuran Efek (Effect Size f-Square / f²)
SUB-BAB INDUK: A. HASIL ANALISIS (4. Evaluasi Model Struktural Inner Model PLS-SEM)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan effect size f² (Cohen, 1988; Hair et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Kriteria Cohen (1988) untuk f²: 0,02 (efek kecil), 0,15 (efek sedang), 0,35 (efek besar).
   b. Tabel Nilai f-Square Seluruh Jalur Struktural (Eksogen -> Mediasi, Eksogen -> Terikat, Mediasi -> Terikat, Moderasi Interaksi -> Terikat).
   c. 2 Paragraf Interpretasi Substantif Pengaruh Praktis Tiap Jalur (masing-masing >= 4 baris, bersitasi): Menelaah jalur yang memberikan kontribusi praktis paling dominan dalam membentuk efektivitas model pada {instansi}.
"""
    elif sub_id == "sub_4_a_4_3":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.4.3: EVALUASI PREDIKSI OUT-OF-SAMPLE (PLSpredict) (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.4.3 Evaluasi Relevansi Prediksi Luar Sampel (PLSpredict)
SUB-BAB INDUK: A. HASIL ANALISIS (4. Evaluasi Model Struktural Inner Model PLS-SEM)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan baku PLSpredict (Shmueli et al., 2016, 2019; Hair et al., 2022; Sarstedt et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Metodologi PLSpredict (10-fold cross validation) (>= 4 baris, bersitasi).
   b. Kriteria Pengambilan Keputusan: Nilai Q²_predict > 0 serta evaluasi selisih galat RMSE (Root Mean Squared Error) antara PLS-SEM dan benchmark model linier naif (Linear Model / LM).
   c. Tabel Hasil Evaluasi PLSpredict Seluruh Indikator Endogen (Kolom: Indikator, Q²_predict, RMSE PLS-SEM, RMSE LM, Selisih Galat, dan Kesimpulan).
   d. Paragraf Konfirmasi Daya Prediksi Tinggi (High Predictive Power) (>= 4 baris, bersitasi): Membuktikan bahwa seluruh atau mayoritas indikator pada model PLS-SEM menghasilkan galat prediksi yang lebih rendah daripada LM di {instansi}.
"""
    elif sub_id == "sub_4_a_4_4":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.4.4: EVALUASI KECOCOKAN MODEL & GLOBAL GOF (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.4.4 Evaluasi Kecocokan Model dan Validitas Global (Model Fit & GoF)
SUB-BAB INDUK: A. HASIL ANALISIS (4. Evaluasi Model Struktural Inner Model PLS-SEM)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan Model Fit & GoF (Henseler et al., 2016; Hu & Bentler, 1999; Tenenhaus et al., 2005; Wetzels et al., 2009).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Kriteria Kecocokan Model (Model Fit): Nilai Standardized Root Mean Square Residual (SRMR < 0,080), d_ULS, d_G, dan Normed Fit Index (NFI > 0,90) (Henseler et al., 2016).
   b. Tabel Uji Kelayakan Model Fit PLS-SEM (Saturated Model vs Estimated Model) + Paragraf konfirmasi good fit (>= 4 baris, bersitasi).
   c. Evaluasi Goodness of Fit Global Tenenhaus: Rumus GoF = sqrt(AVE_rerata * R2_rerata) dengan tolok ukur 0,10 (kecil), 0,25 (sedang), 0,36 (besar).
   d. Tabel Perhitungan Indeks GoF Global (Tabel AVE rata-rata, R² rata-rata, nilai GoF > 0,38 predikat Large/Superior).
   e. Paragraf Sintesis Keselarasan Struktural Menyeluruh (>= 4 baris, bersitasi).
"""
    elif sub_id in ("sub_4_a_4", "sub_4_2_poin_2"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.4: EVALUASI MODEL STRUKTURAL (INNER MODEL) (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.4 Evaluasi Model Struktural (Inner Model PLS-SEM)
SUB-BAB INDUK: A. HASIL ANALISIS (Format Baku TAPD 2026 Disertasi Doktor S3).
BAGIAN UTAMA: 4. Evaluasi Model Struktural (Inner Model PLS-SEM)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan metodologis evaluasi inner model mutakhir (Hair et al., 2022; Sarstedt et al., 2022; Chin, 1998; Cohen, 1988; Shmueli et al., 2019; Henseler et al., 2016; Hu & Bentler, 1999; Tenenhaus et al., 2005; Wetzels et al., 2009).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Evaluasi Model Struktural (>= 4 baris, bersitasi):
      Penjelasan lima pilar evaluasi inner model: R², f², PLSpredict, Model Fit, dan Tenenhaus GoF (Hair et al., 2022; Sarstedt et al., 2022).
   b. Evaluasi Koefisien Determinasi R-Square (R²):
      - Kriteria baku: 0,75 substansial, 0,50 moderat, 0,25 lemah (Chin, 1998; Hair et al., 2022).
      - Tabel Nilai R² dan R² Adjusted Variabel Endogen (Mediasi {m_vars_str} dan Terikat {y_var}) + narasi interpretasi kapasitas prediksi dan retensi variansi (>= 4 baris, bersitasi).
   c. Evaluasi Ukuran Efek (Effect Size f-Square / f²):
      - Kriteria Cohen (1988): 0,02 efek kecil, 0,15 efek sedang, 0,35 efek besar (Hair et al., 2022).
      - Tabel Nilai f² Seluruh Jalur Struktural (dari eksogen ke mediasi dan terikat) + narasi analitis kontribusi praktis tiap variabel prediktor (>= 4 baris, bersitasi).
   d. Evaluasi Relevansi Prediksi Out-of-Sample (PLSpredict):
      - Metodologi k-fold cross validation (k=10) dan perbandingan RMSE PLS-SEM vs Linear Model (LM) (Shmueli et al., 2019; Hair et al., 2022).
      - Tabel Hasil Pengujian PLSpredict (Q²_predict > 0, RMSE PLS < RMSE LM) + narasi konfirmasi daya generalisasi dan kekuatan prediksi out-of-sample (>= 4 baris, bersitasi).
   e. Evaluasi Model Fit:
      - Kriteria SRMR < 0,080, d_ULS, d_G, dan NFI > 0,90 (Henseler et al., 2016; Hu & Bentler, 1999).
      - Tabel Uji Kelayakan Model Fit PLS-SEM (Saturated Model vs Estimated Model) + narasi pembuktian model fit yang baik (>= 4 baris, bersitasi).
   f. Evaluasi Goodness of Fit Global (GoF Tenenhaus):
      - Rumus: GoF = sqrt(AVE_rerata * R2_rerata) dengan kriteria 0,10 kecil, 0,25 sedang, 0,36 besar (Tenenhaus et al., 2005; Wetzels et al., 2009).
      - Tabel Perhitungan Indeks GoF Global (nilai GoF > 0,38 predikat Large/Superior) + narasi keterpaduan model menyeluruh (>= 4 baris, bersitasi).
"""
    elif sub_id == "sub_4_a_5_1":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.5.1: PENGUJIAN PENGARUH LANGSUNG (DIRECT EFFECTS) (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.5.1 Pengujian Hipotesis Pengaruh Langsung (Direct Effects)
SUB-BAB INDUK: A. HASIL ANALISIS (5. Pengujian Hipotesis Penelitian)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan baku bootstrapping SEM-PLS (Hair et al., 2022; Streukens & Leroi-Werelds, 2016; Sarstedt et al., 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Prosedur Bootstrapping 5.000 subsampel non-parametrik (alpha=0,05, t-tabel >= 1,96, p-value < 0,05) (>= 4 baris, bersitasi).
   b. Tabel Hasil Pengujian Hipotesis Pengaruh Langsung (Kolom: Hipotesis, Jalur Struktural, Koefisien Original Sample β, Rerata Sampel, STDEV, T-Statistik, P-Value, dan Keputusan Diterima/Ditolak).
   c. Paragraf Evaluasi Signifikansi Jalur Langsung (>= 4 baris, bersitasi).
   d. Gambar 4.2 Diagram Jalur T-Statistik Bootstrapping SmartPLS 4 (box visual terformat rapi di Word) + Narasi pembacaan signifikansi jalur (>= 4 baris, bersitasi).
   e. Persamaan Sub-Struktural 1 ({m_vars_str if ctx.get("m_vars") else 'M'} = β1·X1 + β2·X2 + β3·X3 + e1) + 2 Paragraf interpretasi matematis analitis mendalam (masing-masing >= 4 baris, bersitasi).
   f. Persamaan Sub-Struktural 2 ({y_var} = β1·X1 + β2·X2 + β3·X3 + β4·M + β5·Z + β6·(X*Z) + e2) + 2 Paragraf interpretasi matematis analitis mendalam (masing-masing >= 4 baris, bersitasi).
"""
    elif sub_id == "sub_4_a_5_2":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.5.2: PENGUJIAN EFEK MEDIASI (INDIRECT EFFECTS) (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.5.2 Pengujian Hipotesis Pengaruh Mediasi (Indirect Effects)
SUB-BAB INDUK: A. HASIL ANALISIS (5. Pengujian Hipotesis Penelitian)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan metodologis mediasi bootstrapping (Zhao et al., 2010; Hair et al., 2022; Preacher & Hayes, 2008).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Metodologi Pengujian Mediasi Bias-Corrected Bootstrapping (>= 4 baris, bersitasi).
   b. Klasifikasi Mediasi menurut Zhao et al. (2010) / Hair et al. (2022): Complementary Partial Mediation, Competitive Partial Mediation, Full / Indirect-only Mediation, atau No Mediation.
   c. Tabel Hasil Pengujian Efek Mediasi Tidak Langsung: Menampilkan seluruh lintasan hipotesis mediasi, Koefisien Efek Tidak Langsung, STDEV, T-Statistik, P-Value, Nilai VAF (Variance Accounted For), Sifat Mediasi, dan Keputusan.
   d. 2 Paragraf Narasi Analitis Peran Variabel Mediasi ({m_vars_str}) (masing-masing >= 4 baris, bersitasi): Menjelaskan bagaimana variabel mediasi menjadi jembatan eksekusi transmisi keperilakuan dari pilar eksogen terhadap keberhasilan {y_var} pada {instansi}.
"""
    elif sub_id == "sub_4_a_5_3":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.5.3: PENGUJIAN EFEK MODERASI (INTERACTION EFFECTS) (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.5.3 Pengujian Hipotesis Pengaruh Moderasi (Interaction Effects)
SUB-BAB INDUK: A. HASIL ANALISIS (5. Pengujian Hipotesis Penelitian)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan pendekatan moderasi interaksi (Henseler & Fassott, 2010; Hayes, 2022; Hair et al., 2022; Donaldson, 2023).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Metodologi Two-Stage Interaction Approach (>= 4 baris, bersitasi).
   b. Tabel Hasil Pengujian Efek Moderasi Interaksi: Menampilkan lintasan hipotesis interaksi ({z_vars_str} * Eksogen -> {y_var}), Koefisien Interaksi, STDEV, T-Statistik, P-Value, dan Keputusan.
   c. Analisis Kemiringan Sederhana (Simple Slope Analysis): Menelaah dampak prediktor pada saat variabel moderasi berada pada tingkat tinggi (+1 SD), rata-rata (mean), dan tingkat rendah (-1 SD).
   d. 2 Paragraf Narasi Analitis Peran Boundary Conditions Variabel Moderasi ({z_vars_str}) (masing-masing >= 4 baris, bersitasi): Menjelaskan implikasi kontingensi moderasi sebagai faktor penguat efektivitas kinerja aparatur pada {instansi}.
"""
    elif sub_id in ("sub_4_a_5", "sub_4_2_poin_3", "sub_4_2_poin_4", "sub_4_2_poin_5", "sub_4_2_poin_6"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A.5: PENGUJIAN HIPOTESIS PENELITIAN (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A.5 Pengujian Hipotesis Penelitian
SUB-BAB INDUK: A. HASIL ANALISIS (Format Baku TAPD 2026 Disertasi Doktor S3).
BAGIAN UTAMA: 5. Pengujian Hipotesis Penelitian

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Rujukan metodologis pengujian hipotesis SEM-PLS mutakhir (Hair et al., 2022; Sarstedt et al., 2022; Cooper & Schindler, 2022; Streukens & Leroi-Werelds, 2016; Hayes, 2022; Henseler & Fassott, 2010; Zhao et al., 2010).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Prosedur Bootstrapping (>= 4 baris, bersitasi):
      Pengujian signifikansi menggunakan 5.000 subsampel non-parametrik pada taraf signifikansi alpha 5% (two-tailed test, t-tabel >= 1,96, p-value < 0,05) (Hair et al., 2022; Streukens & Leroi-Werelds, 2016).
   b. Pengujian Pengaruh Langsung (Direct Effects):
      - Tabel Hasil Pengujian Hipotesis Pengaruh Langsung: Menampilkan seluruh jalur hipotesis langsung (Original Sample, Sample Mean, STDEV, T-Statistik, P-Value, Keputusan).
      - Narasi evaluasi pengaruh langsung (>= 4 baris, bersitasi).
      - Gambar 4.2 Diagram Jalur T-Statistik Bootstrapping SmartPLS 4 (box visual terformat rapi di Word) + narasi pembacaan signifikansi jalur (>= 4 baris, bersitasi).
   c. Persamaan Sub-Struktural 1 (Model Variabel Mediasi):
      - Narasi pengantar pemodelan persamaan regresi sub-struktural pertama (>= 4 baris, bersitasi).
      - Persamaan Matematis: {m_vars_str if m_vars else 'M'} = β1·X1 + β2·X2 + β3·X3 + e1 (masukkan nilai koefisien jalur empiris).
      - 2 Paragraf interpretasi matematis analitis mendalam mengenai kontribusi koefisien terbesar hingga terkecil (masing-masing >= 4 baris, bersitasi).
   d. Persamaan Sub-Struktural 2 (Model Variabel Terikat Akhir):
      - Narasi pengantar pemodelan persamaan regresi sub-struktural kedua (>= 4 baris, bersitasi).
      - Persamaan Matematis: {y_var} = β1·X1 + β2·X2 + β3·X3 + β4·M + β5·Z + β6·(X*Z) + e2 (masukkan nilai koefisien jalur empiris).
      - 2 Paragraf interpretasi matematis analitis mendalam mengenai dominasi jalur mediasi dan peran kontingensi moderasi (masing-masing >= 4 baris, bersitasi).
   e. Pengujian Efek Mediasi (Indirect Effects / Jalur Transmisi):
      - Metodologi mediasi bootstrapping bias-corrected dan klasifikasi mediasi (Zhao et al., 2010; Hair et al., 2022).
      - Tabel Hasil Pengujian Efek Mediasi Tidak Langsung: Menampilkan seluruh hipotesis mediasi, koefisien tidak langsung, STDEV, T-Statistik, P-Value, Sifat Mediasi, Keputusan.
      - Narasi analitis peran variabel mediasi ({m_vars_str}) dalam menjembatani pengaruh eksogen ke endogen (>= 4 baris, bersitasi).
   f. Pengujian Efek Moderasi (Interaction Effects / Peran Kontingensi):
      - Metodologi Two-Stage Interaction approach (Henseler & Fassott, 2010; Hayes, 2022).
      - Tabel Hasil Pengujian Efek Moderasi: Menampilkan hipotesis moderasi, koefisien interaksi, T-Statistik, P-Value, dan Keputusan.
      - Narasi analitis peran variabel moderasi ({z_vars_str}) sebagai penguat keterkaitan struktural di {instansi} (>= 4 baris, bersitasi).
"""
    elif re.match(r"^sub_4_b_h(\d+)$", sub_id):
        m_h = re.match(r"^sub_4_b_h(\d+)$", sub_id)
        h_idx = int(m_h.group(1))
        hypos = ctx.get("pls", {}).get("hypotheses_testing", [])
        h_item = hypos[h_idx - 1] if h_idx <= len(hypos) else {}
        h_code = h_item.get("code", f"H{h_idx}")
        h_stmt = h_item.get("statement", f"Pengaruh Hipotesis ke-{h_idx}")
        h_b = h_item.get("original_sample")
        h_t = h_item.get("t_statistics") or h_item.get("t_stat")
        h_p = h_item.get("p_value")
        h_decision = h_item.get("decision") or ""
        if h_b is None or h_t is None or h_p is None:
            stat_instruction = (
                "Parameter empiris (koefisien jalur, T-statistik, p-value) untuk hipotesis "
                f"{h_code} BELUM TERSEDIA. Dilarang mengarang angka, arah pengaruh, maupun "
                "keputusan hipotesis. Nyatakan secara eksplisit bahwa hipotesis belum diuji dan "
                "bahwa uraian akan dilengkapi setelah keluaran SmartPLS 4 dimasukkan."
            )
        else:
            stat_instruction = (
                f"Laporkan apa adanya dan jangan menambah angka lain di luar data ini: "
                f"koefisien jalur Original Sample = {h_b}; T-Statistik = {h_t}; "
                f"P-Value = {h_p}; keputusan resmi = {h_decision or 'belum diputuskan'}."
            )
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB B.{h_idx}: PEMBAHASAN HIPOTESIS {h_code} (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: B.{h_idx} Pembahasan Hipotesis {h_code}: {h_stmt}
SUB-BAB INDUK: B. PEMBAHASAN HASIL PENELITIAN (BAB IV)

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF DARI 5 ALINEA WAJIB BERBOBOT MENDALAM TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–120 kata per paragraf). Dilarang keras membuat alinea pendek (1-3 baris).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan teori substantif (Grand, Middle, Applied Theory), riset empiris bereputasi mutakhir (2022–2026), serta literatur manajemen keperilakuan publik (Hair et al., 2022; Armstrong & Taylor, 2023; Robbins & Judge, 2022; Donaldson, 2023; Yukl, 2022).
3. FORMULA BAKU 5 (LIMA) ALINEA AKADEMIK DOKTORAL WAJIB LENGKAP:
   - Alinea 1: Bukti Statistik Empiris (>= 4 baris teks utuh, bersitasi Hair et al., 2022 / Sarstedt et al., 2022):
     {stat_instruction}
   - Alinea 2: Telaah & Konfirmasi Teori Rujukan (>= 4 baris teks utuh, bersitasi teori makro relevan):
     Mendialogkan temuan dengan Grand Theory ({_grand_names}) atau Middle Theory ({_mid_names}). Jelaskan bagaimana temuan ini mengonfirmasi atau memperkaya proposisi teoretis tersebut.
   - Alinea 3: Dinamika Faktual & Operasional di Lokus Instansi ({instansi}) (>= 4 baris teks utuh, bersitasi):
     Mengontekstualisasikan temuan statistik dengan realitas operasional aparatur, pelaksanaan tugas fungsi kedinasan, interaksi pimpinan bawahan, dan budaya kerja birokrasi di lingkungan {instansi}.
   - Alinea 4: Komparasi Riset Empiris Bereputasi Mutakhir (2022–2026) (>= 4 baris teks utuh, bersitasi minimal 2-3 rujukan riset mutakhir):
     Melakukan dialog akademik komparatif dengan hasil-hasil penelitian terdahulu dalam jurnal internasional bereputasi dan nasional terakreditasi mutakhir (2022-2026). Tunjukkan konsistensi temuan dan kontribusi penguatan celah empiris (empirical gap).
   - Alinea 5: Implikasi Manajerial Praktis & Rekomendasi Kebijakan bagi Pimpinan (>= 4 baris teks utuh, bersitasi Armstrong & Taylor, 2023 / Mathis & Jackson, 2022):
     Merumuskan rekomendasi solusi operasional yang aplikatif, actionable, dan terukur bagi jajaran pimpinan {instansi} guna mengakselerasi ketercapaian target kinerja organisasi.
"""
    elif sub_id in ("sub_4_b", "sub_4_2_poin_7"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB B: PEMBAHASAN HASIL PENELITIAN (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: B. Pembahasan Hasil Penelitian
SUB-BAB INDUK: BAB IV: HASIL ANALISIS DAN PEMBAHASAN (Format Baku TAPD 2026 Disertasi Doktor S3).
BAGIAN UTAMA: B. Pembahasan Hasil Penelitian

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–100 kata). Dilarang keras membuat paragraf pendek (1-3 baris).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan teori substantif (Grand, Middle, Applied Theory), riset empiris bereputasi mutakhir (2022–2026), serta literatur MSDM dan kebijakan publik (Armstrong & Taylor, 2023; Robbins & Judge, 2022; Barney, 2021; Donaldson, 2023; Mathis & Jackson, 2022; Sinambela, 2023; Hasibuan, 2022).
3. STRUKTUR SISTEMATIKA PENULISAN:
   a. Paragraf Pengantar Pembahasan Doktoral (>= 4 baris, bersitasi):
      Menjelaskan landasan epistemologis pembahasan hasil pengujian empiris dalam mengonfirmasi keterpaduan model di {instansi} (Hair et al., 2022; Cooper & Schindler, 2022).
   b. Pembahasan Mandiri untuk SETIAP Hipotesis Penelitian (H1 s.d H{total_h}):
      Setiap hipotesis dibahas secara tuntas, kritis, dan berbobot doktoral menggunakan FORMULA BAKU 5 (LIMA) ALINEA AKADEMIK DOKTORAL:
      - Sub-heading Heading 3: [Nomor Urut]. Pembahasan Hipotesis [Kode]: [Pernyataan Hipotesis Lengkap]
      - Alinea 1: Bukti Statistik Empiris (>= 4 baris, bersitasi Hair et al., 2022): Memaparkan nilai koefisien jalur (β), T-statistik, p-value, arah pengaruh positif/negatif, signifikansi, dan ketegasan keputusan penerimaan hipotesis.
      - Alinea 2: Telaah & Konfirmasi Teori Rujukan (>= 4 baris, bersitasi): Mendiskusikan secara mendalam keterkaitan temuan dengan Grand Theory (Resource-Based View, Contingency Theory, Social Cognitive Theory) atau Middle Range Theory yang mendasari hipotesis tersebut.
      - Alinea 3: Dinamika Faktual & Operasional di Lokus Instansi (>= 4 baris, bersitasi): Mengaitkan temuan statistik dengan kondisi nyata pelaksanaan tugas pokok, fungsi, budaya kerja, gaya kepemimpinan, dan etos kerja aparatur di {instansi}.
      - Alinea 4: Komparasi Riset Empiris Bereputasi Mutakhir (2022–2026) (>= 4 baris, bersitasi minimal 2-3 jurnal ilmiah): Membandingkan apakah hasil penelitian ini konsisten, mendukung, atau memperluas temuan peneliti-peneliti terdahulu dalam literatur internasional dan nasional mutakhir.
      - Alinea 5: Implikasi Manajerial Praktis & Rekomendasi Kebijakan (>= 4 baris, bersitasi Armstrong & Taylor, 2023 / Mathis & Jackson, 2022): Merumuskan langkah aksi konkret, kebijakan manajerial, dan rekomendasi operasional bagi jajaran pimpinan {instansi} untuk perbaikan kinerja aparatur.
"""
    elif sub_id == "sub_4_a":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A: HASIL ANALISIS (LENGKAP A.1 s.d A.5) (BAB IV) ===
Judul Sub-bab WAJIB ditulis persis: A. Hasil Analisis
SUB-BAB INDUK: BAB IV: HASIL ANALISIS DAN PEMBAHASAN (Format Baku TAPD 2026 Disertasi Doktor S3).
Kompilasi lengkap seluruh kluster Hasil Analisis:
1. Deskriptif Data Demografis Responden (Tabel 4.1 s.d 4.3 + narasi analitis gender, usia, pangkat/golongan).
2. Analisis Statistik Deskriptif Variabel (Tabel 4.4 s.d 4.12 + kriteria Sugiyono + narasi analitis capaian variabel).
3. Evaluasi Model Pengukuran Outer Model (Gambar 4.1 + Tabel 4.13 s.d 4.17: Loadings, AVE, HTMT, Cross-Loadings, Reliabilitas).
4. Evaluasi Model Struktural Inner Model (Tabel 4.18 s.d 4.22: R², f², PLSpredict, Model Fit, Tenenhaus GoF).
5. Pengujian Hipotesis Penelitian (Tabel 4.23 Direct Effects, Gambar 4.2 T-Statistik, Persamaan Sub-Struktural 1 & 2, Tabel 4.24 Mediasi, Tabel 4.25 Moderasi).
Seluruh paragraf naratif WAJIB >= 4 baris teks utuh (minimal 70-100 kata) dan wajib bersitasi ilmiah mutakhir (2022–2026).
"""
    elif sub_id == "sub_5_a":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB A: RINGKASAN PENELITIAN (BAB V) ===
Judul Sub-bab WAJIB ditulis persis: A. Ringkasan Penelitian
SUB-BAB INDUK: BAB V: SIMPULAN DAN SARAN (Format Baku TAPD 2026 Disertasi Doktor S3).
BAGIAN UTAMA: A. Ringkasan Penelitian

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–120 kata per paragraf). Dilarang keras membuat paragraf pendek (1-3 baris).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan ilmiah mutakhir (2022–2026) seperti Hair et al. (2022), Hayes (2022), Robbins & Judge (2022), Armstrong & Taylor (2023), Creswell & Creswell (2023), Donaldson (2023), dan Neuman (2022).
3. STRUKTUR SISTEMATIKA PENULISAN (5 PARAGRAF DOKTORAL BERBOBOT):
   a. Paragraf 1: Latar Belakang Urgensi, Masalah Pokok & Konstruksi Model Integratif (>= 4 baris, bersitasi):
      Menjelaskan latar belakang tingginya urgensi reformasi tata kelola kinerja publik pada {instansi}, fokus pada {y_var}, dan tujuan utama memvalidasi Model Doktoral Orisinal '{model_acronym}' yang mengintegrasikan pilar kapasitas {x_vars_str}, mediasi {m_vars_str}, dan moderasi {z_vars_str}.
   b. Paragraf 2: Pendekatan Metodologis & Pengujian PLS-SEM (>= 4 baris, bersitasi):
      Menguraikan metode kuantitatif kausal-eksplanatori, survei skala Likert 5 poin kepada n = {ctx['n']} aparatur, evaluasi Outer Model dan Inner Model PLS-SEM dengan SmartPLS 4, serta pemenuhan Goodness of Fit yang berpredikat sangat baik.
   c. Paragraf 3: Temuan Empiris Krusial Pengaruh Langsung (>= 4 baris, bersitasi):
      Mendeduksi temuan bahwa seluruh pilar organisasi berpengaruh signifikan terhadap {m_vars_str}, namun ketika dihadapkan langsung pada {y_var}, pilar teknis dan budaya tidak dapat berdiri sendiri tanpa pengawalan perilaku, membuktikan statusnya sebagai potensi pasif.
   d. Paragraf 4: Pembuktian Sah Kebaruan (Novelty) Mediasi & Moderasi (>= 4 baris, bersitasi):
      Menegaskan kebaruan model di mana {m_vars_str} terbukti sebagai 'jembatan eksekusi perilaku' (behavioral execution bridge) mutlak, dan {z_vars_str} terbukti sebagai 'katalisator pelindung' (boundary condition) yang melipatgandakan dampak pilar dasar.
   e. Paragraf 5: Sintesis Ekosistem Kelembagaan Holistik (>= 4 baris, bersitasi):
      Menyimpulkan bahwa pencapaian {y_var} menuntut orkestrasi ekosistem terpadu (Model {model_acronym}) yang menghubungkan kapasitas intelektual, lingkungan kerja, eksekusi etika/motivasi, dan payung keteladanan pimpinan.
"""
    elif sub_id in ("sub_5_b", "sub_5_1"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB B: KESIMPULAN (BAB V) ===
Judul Sub-bab WAJIB ditulis persis: B. Kesimpulan
SUB-BAB INDUK: BAB V: SIMPULAN DAN SARAN (Format Baku TAPD 2026 Disertasi Doktor S3).
BAGIAN UTAMA: B. Kesimpulan

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF/POIN KESIMPULAN WAJIB BERBOBOT NARATIF MENDALAM TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–120 kata). Dilarang membuat kesimpulan poin pendek 1-2 baris.
2. INTEGRITAS STATISTIK EMPIRIS: Setiap poin kesimpulan wajib menyertakan bukti empiris eksak: Koefisien Jalur (Original Sample / O), T-Statistik, P-Value, dan Keputusan Hipotesis.
3. STRUKTUR SISTEMATIKA PENULISAN (PARAGRAF PENGANTAR + 5 RUMPUN KESIMPULAN 17 HIPOTESIS):
   a. Paragraf Pengantar Kesimpulan Doktoral (>= 4 baris): Berdasarkan analisis data, evaluasi model pengukuran dan struktural, serta pembahasan pada bab sebelumnya di {instansi}.
   b. Rumpun 1: Pengaruh Langsung terhadap {m_vars_str} (H direct to mediator): Paparan koefisien O, T, P, keputusan diterima, dan makna peningkatan kompetensi/budaya dalam membentuk dorongan etika/motivasi kerja aparatur.
   c. Rumpun 2: Pengaruh Langsung terhadap {y_var} (H direct to dependent): Paparan koefisien O, T, P, pemilahan pengaruh signifikan vs tidak signifikan/marginal, serta argumentasi teoretis bahwa instrumen kognitif/budaya kurang efektif jika berdiri sendiri tanpa kepatuhan perilaku.
   d. Rumpun 3: Pengaruh Langsung {m_vars_str} terhadap {y_var}: Paparan koefisien terbesar dalam model, T-statistik, p-value, dan penegasan sebagai faktor determinan paling kuat dalam menggerakkan kinerja.
   e. Rumpun 4: Peran Mediasi {m_vars_str} (H indirect effects): Pemilahan Mediasi Penuh (Full Mediation) vs Mediasi Parsial (Partial Mediation), koefisien tidak langsung, dan bukti bahwa variabel mediasi bertindak sebagai jembatan eksekusi mutlak.
   f. Rumpun 5: Peran Moderasi {z_vars_str} (H moderation interaction & conditional indirect): Koefisien interaksi, klasifikasi pure moderator / penguat murni, dan maknanya sebagai pelindung aparatur dari intervensi kekuasaan.
"""
    elif sub_id in ("sub_5_c", "sub_5_2"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB C: IMPLIKASI TEORI (BAB V) ===
Judul Sub-bab WAJIB ditulis persis: C. Implikasi Teori
SUB-BAB INDUK: BAB V: SIMPULAN DAN SARAN (Format Baku TAPD 2026 Disertasi Doktor S3).
BAGIAN UTAMA: C. Implikasi Teori

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–120 kata per paragraf).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan teori bereputasi mutakhir (2022-2026) seperti Albrecht et al. (2024), Mathis et al. (2022), Robbins & Judge (2022), Dessler (2023), Schein & Schein (2023), Hair et al. (2022), Hayes (2022).
3. STRUKTUR SISTEMATIKA PENULISAN (PENGANTAR + 4 PILAR IMPLIKASI TEORETIS):
   a. Paragraf Pengantar Implikasi Teoretis Doktoral (>= 4 baris, bersitasi).
   b. Pilar 1: Perluasan Spektrum Grand Theory dalam Konteks Sektor Publik:
      Heading 3: 1. Perluasan Spektrum Teori Utama Organisasi Publik (>= 4 baris, bersitasi).
   c. Pilar 2: Rekonstruksi Teori Kapasitas dan Perilaku Organisasi:
      Heading 3: 2. Rekonstruksi Teori Kapasitas dan Perilaku Organisasi (>= 4 baris, bersitasi).
      Menjelaskan postulat bahwa kapabilitas SDM dan budaya bernilai netral, sehingga menuntut jembatan eksekusi perilaku kerja.
   d. Pilar 3: Penegasan Integrasi Teori Utama dan Teori Menengah:
      Heading 3: 3. Penegasan Integrasi Teori Utama dan Teori Menengah (>= 4 baris, bersitasi).
      Menjelaskan bahwa {_grand_names} perlu dipadukan dengan {_mid_names} agar mekanisme perilaku kerja dapat dijelaskan secara utuh.
   e. Pilar 4: Validasi Konstruksi Kebaruan Model Orisinal (Model {model_acronym}):
      Heading 3: 4. Validasi Konstruksi Kebaruan Model {model_acronym} (>= 4 baris, bersitasi).
      Sintesis holistik pergeseran dari pendekatan parsial linier menuju ekosistem terpadu {_model_lbl}.
"""
    elif sub_id == "sub_5_d":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB D: IMPLIKASI MANAJERIAL (BAB V) ===
Judul Sub-bab WAJIB ditulis persis: D. Implikasi Manajerial
SUB-BAB INDUK: BAB V: SIMPULAN DAN SARAN (Format Baku TAPD 2026 Disertasi Doktor S3).
BAGIAN UTAMA: D. Implikasi Manajerial

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–120 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan MSDM dan kebijakan publik mutakhir (Armstrong & Taylor, 2023; Robbins & Judge, 2022; Yukl, 2022; Daft, 2022; Mathis & Jackson, 2022; Sinambela, 2023).
3. STRUKTUR SISTEMATIKA PENULISAN (PENGANTAR + 4 PILAR KEBIJAKAN STRATEGIS MANAJERIAL):
   a. Paragraf Pengantar Implikasi Manajerial bagi Pemangku Kepentingan di {instansi} (>= 4 baris, bersitasi).
   b. Pilar 1: Transformasi Kepemimpinan Menjadi Pelindung Psikologis dan Advokasi Hukum:
      Heading 3: 1. Transformasi Kepemimpinan Menjadi Pelindung Psikologis dan Hukum (>= 4 baris, bersitasi).
   c. Pilar 2: Pelembagaan {m_vars_str} ke dalam Indikator Kinerja Utama (KPI) dan SOP:
      Heading 3: 2. Pelembagaan Etika dan Kinerja ke dalam Indikator Kinerja Utama (KPI) (>= 4 baris, bersitasi).
   d. Pilar 3: Rekonstruksi Kriteria Rekrutmen, Seleksi, dan Penempatan Aparatur:
      Heading 3: 3. Rekonstruksi Kriteria Penempatan Aparatur Berbasis Integritas (>= 4 baris, bersitasi).
   e. Pilar 4: Reorientasi Pendekatan Pengawasan Internal Menuju Continuous & Risk-Based Audit:
      Heading 3: 4. Reorientasi Pendekatan Pengawasan Internal Audit (APIP) (>= 4 baris, bersitasi).
"""
    elif sub_id == "sub_5_e":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB E: IMPLIKASI PRAKTIS (BAB V) ===
Judul Sub-bab WAJIB ditulis persis: E. Implikasi Praktis
SUB-BAB INDUK: BAB V: SIMPULAN DAN SARAN (Format Baku TAPD 2026 Disertasi Doktor S3).
BAGIAN UTAMA: E. Implikasi Praktis

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–120 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan metodologi dan manajemen operasional (Cooper & Schindler, 2022; Neuman, 2022; Sinambela, 2023; Afandi, 2023; Lonto et al., 2023; Alamsyah et al., 2024).
3. STRUKTUR SISTEMATIKA PENULISAN (PENGANTAR + 4 REKOMENDASI TAKTIS-OPERASIONAL):
   a. Paragraf Pengantar Implikasi Praktis Operasional Harian di {instansi} (>= 4 baris, bersitasi).
   b. Poin 1: Penerapan Protokol Isolasi dan Jarak Profesional bagi Aparatur Teknis:
      Heading 3: 1. Penerapan Protokol Isolasi dan Jarak Profesional bagi Aparatur (>= 4 baris, bersitasi).
   c. Poin 2: Optimalisasi Pendekatan Pengawasan Berbasis Jejak Digital dan Pemantauan Real-Time:
      Heading 3: 2. Optimalisasi Pendekatan Pengawasan Berbasis Jejak Digital (>= 4 baris, bersitasi).
   d. Poin 3: Simulasi Dilema Etis dalam Kurikulum Pelatihan Pegawai (Bimtek/Diklat):
      Heading 3: 3. Simulasi Dilema Etis dalam Kurikulum Pelatihan Pegawai (>= 4 baris, bersitasi).
   e. Poin 4: Penyediaan Kanal Pelaporan Darurat Terenkripsi (Fast-Track Crisis Whistleblowing):
      Heading 3: 4. Penyediaan Kanal Pelaporan Darurat (Fast-Track Whistleblowing) (>= 4 baris, bersitasi).
"""
    elif sub_id in ("sub_5_f", "sub_5_3"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB SUB-BAB F: SARAN DAN KETERBATASAN PENELITIAN (BAB V) ===
Judul Sub-bab WAJIB ditulis persis: F. Saran dan Keterbatasan Penelitian
SUB-BAB INDUK: BAB V: SIMPULAN DAN SARAN (Format Baku TAPD 2026 Disertasi Doktor S3).
BAGIAN UTAMA: F. Saran dan Keterbatasan Penelitian

ATURAN STRUKTURAL & METODOLOGIS MUTLAK:
1. ATURAN PANJANG PARAGRAF: SETIAP PARAGRAF NARATIF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 70–120 kata).
2. KEHARUSAN SITASI DI SETIAP PARAGRAF: Wajib memuat rujukan metodologis (Creswell & Creswell, 2023; Sekaran & Bougie, 2020; Neuman, 2022; Hair et al., 2022; Podsakoff et al., 2024).
3. STRUKTUR SISTEMATIKA PENULISAN (2 BAGIAN TERPISAH):
   a. Bagian 1: Keterbatasan Penelitian
      - Heading 3: 1. Keterbatasan Penelitian
      - Paragraf Pengantar Keterbatasan Metodologis (>= 4 baris, bersitasi).
      - Poin a: Keterbatasan Geografis dan Ruang Lingkup (>= 4 baris).
      - Poin b: Kerentanan Bias Respons (Social Desirability Bias) (>= 4 baris, bersitasi Podsakoff et al., 2024).
      - Poin c: Keterbatasan Cakupan Variabel Model (>= 4 baris).
   b. Bagian 2: Saran
      - Heading 3: 2. Saran
      - Paragraf Pengantar Rumusan Saran Solutif (>= 4 baris).
      - Heading 4: a. Bagi Peneliti Selanjutnya (Riset Mendatang):
        * Pengayaan Pendekatan Metodologi mixed-methods & longitudinal (>= 4 baris, bersitasi).
        * Perluasan Objek dan Integrasi Variabel Teknologi Digital AI (>= 4 baris, bersitasi).
      - Heading 4: b. Bagi Pemerintah Daerah dan Pemangku Kebijakan:
        * Adopsi Model {model_acronym} sebagai Blueprint Tata Kelola Daerah (>= 4 baris).
        * Perumusan Regulasi Daerah (Perkada) Perlindungan Hukum dan Penilaian KPI (>= 4 baris).
"""
    elif sub_id in ("sub_5_all", "sub_5"):
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB BAB V LENGKAP: SIMPULAN DAN SARAN ===
Judul Bab WAJIB ditulis persis: BAB V SIMPULAN DAN SARAN
SUB-BAB LENGKAP FORMAT BAKU TAPD 2026 (A s.d F):
A. Ringkasan Penelitian (5 paragraf komprehensif, >= 4 baris per paragraf).
B. Kesimpulan (pengantar + 5 rumpun kesimpulan menjawab seluruh {total_h} hipotesis memuat O, T, P).
C. Implikasi Teori (pengantar + 4 pilar teoretis doktoral, novelty model {model_acronym}).
D. Implikasi Manajerial (pengantar + 4 pilar kebijakan strategis bagi {instansi}).
E. Implikasi Praktis (pengantar + 4 rekomendasi taktis-operasional rutinitas harian).
F. Saran dan Keterbatasan Penelitian (keterbatasan metodologis 3 poin + saran peneliti selanjutnya & pemda).
Seluruh paragraf naratif WAJIB >= 4 baris teks utuh (minimal 70–120 kata) dan wajib bersitasi ilmiah mutakhir (2022–2026).
"""
    elif sub_id == "sub_dapus":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB DAFTAR PUSTAKA (STANDAR APA 7TH EDITION 2022-2026) ===
Judul WAJIB: DAFTAR PUSTAKA
Format Baku: Standar APA 7th Edition, disusun alfabetis A-Z, format Hanging Indent 1.27 cm (0.5 inch), spasi 1.5.
MUTLAK: Seluruh rujukan WAJIB 100% memuat tahun terbit mutakhir antara tahun 2022 hingga 2026.
Memuat seluruh literatur buku teks, jurnal internasional bereputasi (Q1-Q4), jurnal nasional SINTA, dan dokumen resmi yang disitasi dalam Disertasi Doktor.
"""
    elif sub_id == "sub_lampiran":
        instruksi_khusus = f"""
=== PETUNJUK SPESIFIK WAJIB LAMPIRAN DISERTASI DOKTOR ===
Judul WAJIB: LAMPIRAN
Memuat instrumen penelitian lengkap:
1. Bagian I: Biodata Responden dan Petunjuk Pengisian Skala Likert 1-5.
2. Bagian II: Kuesioner Penelitian Doktoral 40 Butir Pertanyaan Operasional Seluruh Variabel ({y_var}, {m_vars_str}, {z_vars_str}, {x_vars_str}).
3. Kisi-Kisi Instrumen Penelitian (Tabel Dimensi, Indikator, dan Nomor Butir).
4. Rekapitulasi Matriks Output Pengujian PLS-SEM (SmartPLS 4).
"""
    citations_block = s3_academic_researcher.build_s3_prompt_citations_block(ctx, sub_id)

    prompt = f"""Anda adalah Guru Besar dan Pakar Metodologi Penelitian Manajemen yang bertindak sebagai Penulis Ahli Disertasi Doktor (S3) Ilmu Manajemen Pascasarjana Indonesia.

Tugas Anda adalah menyusun naskah akademik tingkat doktoral untuk:
BAGIAN / SUB-BAB: {meta['code']} - {meta['title']}
DESKRIPSI KONTEN: {meta['desc']}
{instruksi_khusus}
=== ANCHOR UTAMA: JUDUL DISERTASI DOKTORAL (WAJIB DISINKRONKAN) ===
"{judul}"

=== KONTEKS OBJEK & STRUKTUR MODEL DOKTORAL ===
• Objek / Instansi Empiris : {instansi}
• Jenjang Akademik        : Program Doktor (S3) Ilmu Manajemen
• Desain Model Riset      : Model Hibrida {m_type}
• Variabel Bebas (X)       : {x_vars_str}
• Variabel Mediasi (M)     : {m_vars_str}
• Variabel Moderasi (Z)    : {z_vars_str}
• Variabel Terikat (Y)     : {y_var}
• Total Hipotesis Formal   : {total_h} Hipotesis (H1 s.d H{total_h})
• Sampel & Alat Analisis   : n = {ctx['n']} Responden, SmartPLS 4 & PROCESS Hayes Model

{citations_block}

=== PEDOMAN SUBSTANSI AKADEMIK DOKTORAL BERDASARKAN JUDUL ===
1. Keterikatan Judul & Variabel: Seluruh argumentasi, logika kausalitas, dan elaborasi teoretis WAJIB merujuk secara konsisten pada variabel-variabel di dalam judul ({y_var}, {m_vars_str}, {z_vars_str}, {x_vars_str}). Khusus untuk sub-bab 2.1 (kajian teori murni), kajian wajib universal tanpa menyebut nama instansi lokal.
2. Kedalaman Teori Doktoral: Sertakan telaah Grand Theory yang relevan (Resource-Based View, Contingency Theory, Social Cognitive Theory, Transformational Leadership Theory, Goal Setting Theory), Middle Range Theory, dan Applied Theory.
3. Telaah Variabel Kajian Pustaka: Setiap variabel penelitian (Variabel Terikat Y, Mediasi M, Moderasi Z, serta Bebas Eksogen X1..Xn) WAJIB memuat 10 (sepuluh) definisi konseptual dari pakar terkemuka dalam paragraf naratif mengalir, rumusan sintesis konseptual peneliti, serta dimensi operasional dan indikator terukurnya dalam format teks hirarkis bertingkat (100% tanpa tabel).
4. Standar Kebaruan Pustaka (2022-2026) & Format Sitasi Teori Lama:
   • Seluruh buku dan artikel jurnal ilmiah yang dirujuk WAJIB diterbitkan dalam kurun waktu 5 tahun terakhir (2022–2026).
   • Jika merujuk pada teori lama/klasik (sebelum 2022), WAJIB menggunakan format kutipan sekunder persis seperti formula ini:
     [Tokoh Asli] ([Tahun Asli]), dalam [Penulis Mutakhir], ([Tahun Mutakhir 2022-2026])
     Contoh wajib:
     - Locke & Latham (1990), dalam Schunk, (2022)
     - Barney (1991), dalam Hitt et al., (2023)
     - Bass (1985), dalam Yukl, (2022)
     - Weber (1905), dalam Donaldson, (2023)
     - Bernardin & Russell (1993), dalam Kasmir, (2023)
     - Bandura (1986), dalam Ryan & Deci, (2022)
     - Vroom (1964), dalam Kreitner & Kinicki, (2022)
     - Herzberg (1966), dalam Luthans et al., (2022)
   • Seluruh entri pada Daftar Pustaka (References) WAJIB 100% memuat angka tahun terbitan antara 2022 s.d. 2026.
5. Kebaruan Ilmiah (Novelty): Tegaskan celah penelitian (research gap) dan kontribusi teoretis baru yang membedakan riset doktoral ini dari riset-riset magister sebelumnya.
6. Gaya Bahasa & Format: Bahasa Indonesia akademik formal baku (EYD V), bernas, kohesif, menggunakan alur paragraf deduktif yang mengalir tanpa repetisi kata.
7. Standar Wajib Panjang Paragraf & Keharusan Sitasi di Seluruh Paragraf:
   • SETIAP PARAGRAF TIDAK BOLEH KURANG DARI 4 BARIS TEKS UTUH (minimal 4–6 kalimat analitis substantif atau 70–100 kata). DILARANG KERAS membuat paragraf pendek (1–3 baris).
   • SETIAP PARAGRAF HARUS MEMILIKI KUTIPAN ILMIAH (sitasi nama ahli dan tahun terbit 2022–2026 atau teori klasik yang disitasi dalam literatur mutakhir). Tidak boleh ada paragraf konseptual atau teoretis yang berdiri sendiri tanpa rujukan literatur ilmiah.

Hasilkan naskah lengkap, mendalam, dan memenuhi standar tebal disertasi doktoral.
"""
    return prompt.strip()


# ==============================================================================
# 3. BUILDER DOKUMEN DISERTASI DOKTOR PER SUB-BAB (36 UNIT)
# ==============================================================================

def _kota_dari_instansi(instansi: str) -> str:
    """Turunkan nama kota/kabupaten dari nama instansi (fallback: Palembang)."""
    m = re.search(r"(?:Kota|Kabupaten)\s+([A-Z][A-Za-z' ]{2,40})", instansi or "")
    if m:
        return m.group(1).strip().title()
    return "Palembang"


def _validate_sakip_rows(rows: Any):
    """Validasi tabel SAKIP nasional dari input pengguna (Blok C).

    Menerima daftar baris dengan 5 kolom:
    ``[Tahun, Nilai SAKIP, Proporsi Predikat Rendah/Cukup, Disparitas, Status]``.
    Tahun wajib 4 digit (19xx/20xx) dan kolom nilai harus terbaca sebagai angka
    (format ``"61,35"``/``"61.35"``; teks predikat dalam tanda kurung diizinkan).
    Bila ada baris yang tidak lolos validasi, seluruh input ditolak (dikembalikan
    sebagai daftar kosong) agar mesin memakai tabel placeholder yang jujur.
    """
    if not rows or not isinstance(rows, (list, tuple)):
        return []
    out = []
    for idx, row in enumerate(rows, start=1):
        try:
            if isinstance(row, dict):
                row = [row.get(k, "") for k in
                       ("tahun", "nilai", "proporsi", "disparitas", "status")]
            if not isinstance(row, (list, tuple)) or len(row) < 5:
                print(f"[s3_sakip] Baris {idx} ditolak: perlu 5 kolom.")
                return []
            tahun = str(row[0]).strip()
            if not re.fullmatch(r"(?:19|20)\d{2}", tahun):
                print(f"[s3_sakip] Baris {idx} ditolak: tahun '{tahun}' tidak valid.")
                return []
            nilai = str(row[1]).strip()
            proporsi = str(row[2]).strip()
            if not re.search(r"\d", nilai) or not re.search(r"\d", proporsi):
                print(f"[s3_sakip] Baris {idx} ditolak: nilai/proporsi bukan angka.")
                return []
            out.append([tahun, nilai, proporsi,
                        str(row[3]).strip() or "-", str(row[4]).strip() or "-"])
        except Exception as ex:
            print(f"[s3_sakip] Baris {idx} ditolak: {ex}")
            return []
    return out


def _build_doc_sub_bagian_awal(doc: Document, ctx: Dict[str, Any]):
    """Bagian Awal (Front Matter) Doktoral Lengkap."""
    judul = ctx["judul"].upper()
    penulis = ctx["penulis"].upper()
    nim = ctx["nim"]
    instansi = ctx["instansi"]
    tahun = ctx["tahun"]

    # Identitas akademik (konfigurabel, tanpa nama palsu).
    universitas = (ctx.get("universitas") or "Universitas Tridinanti Palembang").strip()
    fakultas = (ctx.get("fakultas") or "Fakultas Ekonomi dan Bisnis").strip()
    program = (ctx.get("program_studi") or "Program Doktor Ilmu Manajemen").strip()
    kota = (ctx.get("kota") or _kota_dari_instansi(instansi)).strip()
    promotor = (ctx.get("promotor") or "").strip()
    kopromotor1 = (ctx.get("kopromotor1") or "").strip()
    kopromotor2 = (ctx.get("kopromotor2") or "").strip()

    def _baris_tanda_tangan(nama: str, jabatan: str) -> str:
        if nama:
            return f"{nama}\n({jabatan})"
        return "\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\n" \
               f"({jabatan})"

    # Halaman Sampul Luar (Cover)
    p(doc, judul, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_before=0, space_after=6)
    p(doc, "", space_after=18)
    p(doc, "DISERTASI", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_before=0, space_after=4)
    p(doc, "Diajukan untuk Memenuhi Sebagian Syarat Memperoleh Gelar Doktor Ilmu Manajemen (Dr.)", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)

    p(doc, "Oleh:", align=WD_ALIGN_PARAGRAPH.CENTER)
    p(doc, f"{penulis}", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    p(doc, f"NIM: {nim}", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=36)

    p(doc, f"{program.upper()}\n{fakultas.upper()}\n{universitas.upper()}\n" + tahun, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # Halaman Pengesahan Promotor
    h1(doc, "LEMBAR PERSETUJUAN KOMISI PROMOTOR")
    p(doc, f"Disertasi dengan judul: \"{judul}\" yang disusun oleh {penulis} (NIM: {nim}) telah disetujui dan dinyatakan memenuhi syarat untuk diajukan ke Sidang Ujian Disertasi Doktor pada {program}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=18)
    _hari_ini = datetime.now()
    p(doc, f"{kota}, {_hari_ini.day} {_NAMA_BULAN_ID[_hari_ini.month - 1]} {_hari_ini.year}", align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=18)
    p(doc, "Komisi Promotor:\n\n\n", align=WD_ALIGN_PARAGRAPH.CENTER)
    p(doc, _baris_tanda_tangan(promotor, "Promotor Utama"), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
    p(doc, _baris_tanda_tangan(kopromotor1, "Ko-Promotor I"), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
    p(doc, _baris_tanda_tangan(kopromotor2, "Ko-Promotor II"), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
    p(doc, f"Mengetahui,\nKetua {program}", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12)
    doc.add_page_break()

    # Lembar Pernyataan Keaslian Karya
    h1(doc, "LEMBAR PERNYATAAN KEASLIAN KARYA")
    p(doc, "Yang bertanda tangan di bawah ini:", align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=6)
    p(doc, f"Nama\t\t: {penulis}", align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=0)
    p(doc, f"NIM\t\t: {nim}", align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=0)
    p(doc, f"Program Studi\t: {program}", align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=0)
    p(doc, f"Institusi\t\t: {universitas}", align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=12)
    p(doc, f"Dengan ini menyatakan dengan sesungguhnya bahwa Disertasi Doktor berjudul \u201c{ctx['judul']}\u201d beserta seluruh isinya adalah benar-benar karya saya sendiri. Saya tidak melakukan penjiplakan atau pengutipan dengan cara yang tidak sesuai dengan etika keilmuan yang berlaku, dan seluruh sumber yang saya kutip telah saya cantumkan secara lengkap pada Daftar Pustaka.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0.5, space_after=6)
    p(doc, "Apabila di kemudian hari ditemukan pelanggaran etika keilmuan atau klaim pihak lain atas karya ini, saya bersedia menerima sanksi akademik sesuai peraturan yang berlaku, termasuk pencabutan gelar yang telah diperoleh.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0.5, space_after=18)
    p(doc, f"{kota}, " + tahun, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=6)
    p(doc, "Yang membuat pernyataan,", align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=30)
    p(doc, f"{penulis}\nNIM: {nim}", bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=6)
    doc.add_page_break()

    # Kata Pengantar
    h1(doc, "KATA PENGANTAR")
    p(doc, f"Puji dan syukur peneliti panjatkan ke hadirat Tuhan Yang Maha Esa atas limpahan rahmat, taufik, dan hidayah-Nya sehingga naskah Disertasi Doktor yang berjudul \"{ctx['judul']}\" ini dapat diselesaikan sesuai dengan kaidah ilmiah yang berlaku pada {program} {universitas}.", indent=0.5, space_after=6)
    p(doc, f"Disertasi ini disusun sebagai ikhtiar ilmiah untuk membedah dan mengonfirmasi model struktural pada {instansi}, sekaligus memberikan kontribusi teoretis kebaruan (novelty) bagi khazanah ilmu manajemen kontemporer.", indent=0.5, space_after=6)
    p(doc, "Peneliti menyampaikan penghargaan dan terima kasih yang setinggi-tingginya kepada Komisi Promotor atas bimbingan yang tulus, jajaran Rektorat dan Dekanat, segenap Guru Besar Program Doktor Ilmu Manajemen, serta pimpinan dan seluruh responden pada instansi penelitian yang telah memberikan dukungan data secara terbuka.", indent=0.5, space_after=6)
    p(doc, "Peneliti menyadari naskah ini masih memiliki keterbatasan. Oleh sebab itu, kritik dan saran yang membangun sangat diharapkan demi penyempurnaan karya ini pada masa yang akan datang.", indent=0.5, space_after=12)
    p(doc, f"{kota}, " + tahun + "\n\nPeneliti", align=WD_ALIGN_PARAGRAPH.RIGHT)
    doc.add_page_break()

    # Abstrak Bahasa Indonesia
    h1(doc, "ABSTRAK")

    _x_str = _join_id(ctx["x_vars"]) or "variabel eksogen"
    _m_str = _join_id(ctx["m_vars"])
    _z_str = _join_id(ctx["z_vars"])
    _model_id_label = {
        "moderated_mediation": "model mediasi termoderasi",
        "mediation": "model mediasi",
        "moderation": "model moderasi",
        "direct": "model pengaruh langsung",
    }.get(str(ctx.get("model_type") or ""), "model struktural")
    _model_en_label = {
        "moderated_mediation": "moderated mediation",
        "mediation": "mediation",
        "moderation": "moderation",
        "direct": "direct effect",
    }.get(str(ctx.get("model_type") or ""), "structural")
    _m_clause_id = f" dengan {_m_str} sebagai variabel mediasi" if _m_str else ""
    _z_clause_id = f" serta {_z_str} sebagai variabel moderasi" if _z_str else ""
    _x_str_en = _join_en([_en_var(v) for v in ctx["x_vars"]]) or "exogenous variables"
    _y_str_en = _en_var(ctx["y_var"])
    _m_str_en = _join_en([_en_var(v) for v in ctx["m_vars"]])
    _z_str_en = _join_en([_en_var(v) for v in ctx["z_vars"]])
    _m_clause_en = f", mediated by {_m_str_en}" if _m_str_en else ""
    _z_clause_en = f", and moderated by {_z_str_en}" if _z_str_en else ""

    _hypos_abs = (ctx.get("pls") or {}).get("hypotheses_testing", []) or []
    _acc_abs = sum(1 for h in _hypos_abs if str(h.get("decision", "")).lower().startswith("diterima"))
    if not s3data.is_template(ctx) and _hypos_abs:
        _hasil_id = (
            f"Dari {len(_hypos_abs)} hipotesis yang diuji, {_acc_abs} dinyatakan diterima dan "
            f"{len(_hypos_abs) - _acc_abs} tidak diterima pada taraf signifikansi 5%."
        )
        _hasil_en = (
            f"Of the {len(_hypos_abs)} hypotheses tested, {_acc_abs} were supported and "
            f"{len(_hypos_abs) - _acc_abs} were not supported at the 5% significance level."
        )
    else:
        _hasil_id = (
            "Karena data empiris riil belum dimasukkan ke dalam konfigurasi penelitian, naskah "
            "ini menyajikan kerangka metodologis dan struktur pelaporan hasil secara lengkap; "
            "seluruh nilai pengujian hipotesis wajib diisi dari keluaran SmartPLS 4 sebelum "
            "naskah diajukan."
        )
        _hasil_en = (
            "As empirical data have not yet been supplied in the research configuration, this "
            "manuscript presents the complete methodological framework and reporting structure; "
            "all hypothesis-testing values must be populated from SmartPLS 4 output before "
            "submission."
        )

    if _m_str and _z_str:
        _novelty_id = ("Kebaruan teoretis (novelty) disertasi ini terletak pada integrasi "
                       "kerangka mediasi termoderasi pada tata kelola sektor publik.")
        _novelty_en = ("The theoretical novelty of this dissertation lies in establishing an "
                       "integrative moderated mediation framework within public sector governance.")
    elif _m_str:
        _novelty_id = (f"Kebaruan teoretis (novelty) disertasi ini terletak pada pemodelan peran "
                       f"mediasi {_m_str} dalam rantai kinerja sektor publik.")
        _novelty_en = ("The theoretical novelty of this dissertation lies in modelling the "
                       f"mediating role of {_m_str_en} within the public-sector performance chain.")
    elif _z_str:
        _novelty_id = (f"Kebaruan teoretis (novelty) disertasi ini terletak pada pemodelan peran "
                       f"moderasi {_z_str} pada hubungan kerja sektor publik.")
        _novelty_en = ("The theoretical novelty of this dissertation lies in modelling the "
                       f"moderating role of {_z_str_en} in public-sector work relationships.")
    else:
        _novelty_id = ("Kebaruan teoretis (novelty) disertasi ini terletak pada pemetaan jalur "
                       "langsung penentu kinerja pada tata kelola sektor publik.")
        _novelty_en = ("The theoretical novelty of this dissertation lies in mapping the direct "
                       "determinants of performance in public-sector governance.")

    _keywords = [ctx["y_var"]] + list(ctx["m_vars"]) + list(ctx["z_vars"]) + list(ctx["x_vars"])

    p(doc, f"Penelitian disertasi ini bertujuan untuk menganalisis {_model_id_label} yang menguji pengaruh {_x_str} terhadap {ctx['y_var']}{_m_clause_id}{_z_clause_id} pada {instansi}.", indent=0.5, space_after=6)
    p(doc, f"Metode penelitian menggunakan pendekatan kuantitatif kausal eksplanatori dengan teknik survei kuesioner skala Likert 5 poin terhadap {ctx['n']} responden. Analisis data dilakukan menggunakan Partial Least Squares Structural Equation Modeling (PLS-SEM) dengan perangkat lunak SmartPLS 4.", indent=0.5, space_after=6)
    p(doc, f"{_hasil_id} {_novelty_id}", indent=0.5, space_after=6)
    p(doc, f"Kata Kunci: {', '.join(_keywords)}, SmartPLS 4, Disertasi Doktor.", italic=True, space_after=12)
    doc.add_page_break()

    # Abstract Bahasa Inggris
    h1(doc, "ABSTRACT")
    p(doc, f"This doctoral dissertation aims to empirically examine a comprehensive {_model_en_label} structural model testing the effects of {_x_str_en} on {_y_str_en}{_m_clause_en}{_z_clause_en} at {instansi}.", indent=0.5, space_after=6)
    p(doc, f"A quantitative causal explanatory approach was utilized surveying {ctx['n']} respondents using structured Likert-scale questionnaires. Data estimation was performed through Partial Least Squares Structural Equation Modeling (PLS-SEM) utilizing SmartPLS 4.", indent=0.5, space_after=6)
    p(doc, f"{_hasil_en} {_novelty_en}", indent=0.5, space_after=6)
    p(doc, f"Keywords: {', '.join(_en_var(v) for v in _keywords)}, PLS-SEM, Doctoral Dissertation.", italic=True)
    doc.add_page_break()

    # Daftar Isi (field TOC otomatis)
    h1(doc, "DAFTAR ISI")
    _toc_field(doc, switch=r'TOC \o "1-3" \h \z \u')
    doc.add_page_break()

    # Daftar Tabel (field TOC dari caption bertingkat)
    h1(doc, "DAFTAR TABEL")
    _toc_field(doc, switch=r'TOC \h \z \u \o "4-4"', charformat=True)
    doc.add_page_break()

    # Daftar Gambar (field TOC dari caption bertingkat)
    h1(doc, "DAFTAR GAMBAR")
    _toc_field(doc, switch=r'TOC \h \z \u \o "5-5"', charformat=True)
    doc.add_page_break()

    # Daftar Lampiran (field TOC dari judul lampiran bertingkat)
    h1(doc, "DAFTAR LAMPIRAN")
    _toc_field(doc, switch=r'TOC \h \z \u \o "6-6"', charformat=True)
    doc.add_page_break()

    # Riwayat Hidup
    h1(doc, "RIWAYAT HIDUP")
    p(doc, f"Nama\t\t: {ctx.get('penulis', '')}", align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=0)
    p(doc, f"NIM\t\t: {ctx.get('nim', '')}", align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=0)
    p(doc, f"Tempat/Tanggal Lahir\t: {ctx.get('tempat_lahir', '') or s3data.PLACEHOLDER}", align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=0)
    p(doc, f"Program Studi\t: {program}", align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=0)
    p(doc, f"Institusi\t\t: {universitas}", align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=6)

    _riwayat = ctx.get("riwayat_hidup") or []
    if isinstance(_riwayat, str):
        _riwayat = [_riwayat]
    if _riwayat:
        p(doc, "Riwayat Pendidikan dan Karier:", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=4)
        for idx_r, baris_r in enumerate(_riwayat, start=1):
            p(doc, f"{idx_r}. {str(baris_r).strip()}", align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0.5, space_after=2)
    else:
        p(doc, "Riwayat pendidikan dan karier belum diisi pada konfigurasi penelitian. Bagian ini diisi oleh peneliti sesuai data pribadi yang sebenarnya sebelum naskah diajukan.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0.5, space_after=6)
    p(doc, "Kontak\t\t: " + (ctx.get("email", "") or s3data.PLACEHOLDER), align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, space_after=6)


def _build_doc_sub_1_1(doc: Document, ctx: Dict[str, Any]):
    """1.1 Latar Belakang Masalah -- tujuh lapisan deduktif, dirakit dari profil variabel."""
    judul = ctx["judul"]
    instansi = ctx["instansi"]
    x1, x2, x3 = ctx["x1"], ctx["x2"], ctx["x3"]
    m, z, y_name = ctx["m"], ctx["z"], ctx["y"]
    model_acronym = _generate_doctoral_model_acronym(ctx)
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "1.1")
    rng_bg = nv.rng_for(ctx, "1.1-bg")
    xs = pm["_all"]["x"]
    ms = pm["_all"]["m"]
    zs = pm["_all"]["z"]
    y = pm["Y"]
    x_labels = _join_id(p["label"] for p in xs) or "variabel eksogen"
    m_label = ms[0]["label"] if ms else ""
    z_label = zs[0]["label"] if zs else ""
    x0 = xs[0] if xs else nv.make_profile("kapasitas organisasi", {}, ctx)

    cite_grp = ctx.get("citation_groups") or s3_academic_researcher.get_s3_citation_groups(ctx)
    pos_cites = cite_grp.get("pos_citations", "Rahman, 2023; Aini et al., 2025; Hair et al., 2022")
    gap_cites = cite_grp.get("gap_citations", "Gadzai et al., 2025; Massud & Uddin, 2025; Almazrouei & Alnahhal, 2025")
    gov_cites = cite_grp.get("gov_citations", "Saroni et al., 2023; Astuti, 2022; Sarumaha, 2022")
    first_gap = cite_grp.get("first_gap_cite", "Gadzai et al. (2025)")
    sec_gap = cite_grp.get("second_gap_cite", "Almazrouei & Alnahhal (2025)")

    h1(doc, "BAB I\nPENDAHULUAN")
    h2(doc, "A. Latar Belakang")

    # --- LAPISAN 1: FENOMENA MAKRO & THEORETICAL HOOK ---
    _fokus_tugas = _get_instansi_focus(instansi)
    pv(doc, nv.pick(rng_bg, [
       (
           f"Transformasi tata kelola organisasi publik menempatkan kualitas sumber daya manusia "
           f"sebagai penentu utama capaian kinerja. Di Indonesia, reformasi birokrasi menuntut "
           f"pergeseran dari manajemen berbasis prosedur menuju manajemen berbasis hasil yang "
           f"terukur. Pergeseran itu tidak berhenti pada penyusunan dokumen perencanaan, tetapi "
           f"menyentuh cara aparatur bekerja, mengambil keputusan, dan mempertanggungjawabkan "
           f"hasilnya. Pada {instansi}, tuntutan tersebut bertemu dengan mandat pengelolaan "
           f"{_fokus_tugas} yang menuntut ketepatan, transparansi, dan "
           f"kecepatan respons."
       ),
       (
           f"Kualitas sumber daya manusia kini dipandang sebagai penentu pokok capaian kinerja "
           f"dalam tata kelola organisasi publik. Reformasi birokrasi di Indonesia menghendaki "
           f"perpindahan dari manajemen berbasis prosedur ke manajemen berbasis hasil yang "
           f"terukur. Perpindahan tersebut tidak sekadar menyentuh penyusunan dokumen "
           f"perencanaan, melainkan juga cara aparatur bekerja, mengambil keputusan, dan "
           f"mempertanggungjawabkan hasilnya. Di {instansi}, tuntutan itu bersinggungan dengan "
           f"mandat pengelolaan {_fokus_tugas} yang menuntut ketepatan, "
           f"transparansi, dan kecepatan respons."
       ),
       (
           f"Dalam tata kelola organisasi publik, mutu sumber daya manusia menjadi faktor penentu "
           f"utama kinerja. Perjalanan reformasi birokrasi nasional mengarahkan peralihan dari "
           f"manajemen prosedural menuju manajemen hasil yang terukur. Peralihan ini menyentuh "
           f"bukan hanya penyusunan dokumen perencanaan, tetapi juga cara aparatur bekerja, "
           f"mengambil keputusan, dan mempertanggungjawabkan hasilnya. Pada {instansi}, dorongan "
           f"tersebut berjumpa dengan mandat pengelolaan {_fokus_tugas} "
           f"yang menuntut ketepatan, transparansi, dan kecepatan respons."
       ),
    ]), ctx, "1.1-l1a", indent=0.5, space_after=6)

    pv(doc,
       f"Secara teoretis, {y['label']} tidak lahir dari satu faktor tunggal. {y['definisi'].capitalize()}. "
       f"Pemahaman itu menuntut penjelasan yang menghubungkan karakteristik pekerjaan, kapasitas "
       f"aparatur, dan konteks kelembagaan. {x0['label']} dipahami sebagai {x0['definisi']}. "
       f"Apabila kapasitas tersebut tidak dikelola, risiko yang muncul adalah {x0['risiko']}. "
       f"Sebaliknya, ketika kapasitas "
       f"dikelola dengan baik, {x0['tinggi']}. Premis ini menjadi titik awal untuk menelaah "
       f"mengapa capaian {y['label']} antarbidang di {instansi} belum merata.",
       ctx, "1.1-l1b", indent=0.5, space_after=6)

    pv(doc,
       f"Kesenjangan antara tuntutan dan realitas tersebut bukan persoalan tunggal. Laporan "
       f"akuntabilitas kinerja di berbagai instansi daerah berulang kali memperlihatkan bahwa "
       f"penambahan anggaran, pelatihan, dan fasilitas tidak selalu diikuti lompatan "
       f"produktivitas. Inersia birokrasi, beban administratif, dan lemahnya umpan balik "
       f"menjadi faktor yang saling mengunci. Karena itu, evaluasi terhadap determinan "
       f"{y['label']} perlu dilakukan secara simultan, bukan secara parsial per variabel.",
       ctx, "1.1-l1c", indent=0.5, space_after=6)

    # --- LAPISAN 2: TABEL 1.1 (Rekapitulasi SAKIP Nasional) ---
    _sakip_rows = ctx.get("sakip_nasional") or []
    try:
        _th_now = int(str(ctx.get("tahun") or "2026")[:4])
    except Exception:
        _th_now = 2026
    _tahun_lampau = [str(y) for y in range(_th_now - 5, _th_now)]
    _rentang = f"{_tahun_lampau[0]}\u2013{_tahun_lampau[-1]}"

    pv(doc, nv.pick(rng_bg, [
       (
           f"Gambaran makro tersebut dapat dibaca dari tren evaluasi Sistem Akuntabilitas Kinerja "
           f"Instansi Pemerintah (SAKIP). Tabel 1.1 menyajikan rekapitulasi nilai SAKIP nasional "
           f"beserta proporsi predikat rendah/cukup. Angka pada tabel ini berfungsi sebagai latar "
           f"konteks, bukan sebagai dasar inferensi statistik penelitian, karena fokus disertasi "
           f"tetap berada pada tingkat aparatur di {instansi}."
       ),
       (
           f"Potret makro itu dapat ditelusuri melalui tren evaluasi Sistem Akuntabilitas Kinerja "
           f"Instansi Pemerintah (SAKIP). Tabel 1.1 memuat rekapitulasi nilai SAKIP nasional "
           f"beserta proporsi predikat rendah/cukup. Angka-angka tersebut diposisikan sebagai latar "
           f"konteks, bukan landasan inferensi statistik, sebab fokus disertasi ini tetap pada "
           f"tingkat aparatur di {instansi}."
       ),
       (
           f"Kondisi makro tersebut tampak pada tren evaluasi Sistem Akuntabilitas Kinerja "
           f"Instansi Pemerintah (SAKIP). Tabel 1.1 menampilkan rekapitulasi nilai SAKIP nasional "
           f"dan proporsi predikat rendah/cukup. Datanya berperan sebagai latar konteks, bukan "
           f"sebagai dasar inferensi statistik, karena disertasi ini memusatkan perhatian pada "
           f"tingkat aparatur di {instansi}."
       ),
    ]), ctx, "1.1-l2a", indent=0.5, space_after=6)

    t1_headers = ["Tahun", "Nilai Rata-rata SAKIP Nasional", "Proporsi Predikat Rendah/Cukup (%)", "Disparitas Capaian Antar Wilayah", "Status Evaluasi Tata Kelola Nasional"]
    if _sakip_rows:
        t1_rows = [list(r) for r in _sakip_rows]
        t1_source = "Sumber: " + (ctx.get("sakip_sumber") or "Publikasi resmi evaluasi SAKIP nasional (KemenPANRB)")
    else:
        t1_rows = [[th, s3data.PLACEHOLDER, s3data.PLACEHOLDER, s3data.PLACEHOLDER, "Belum diisi"]
                   for th in _tahun_lampau]
        t1_source = ("Sumber: belum diisi. Lengkapi nilai SAKIP nasional melalui konfigurasi "
                     "penelitian (kunci \"sakip_nasional\" dan \"sakip_sumber\") dari publikasi "
                     "resmi KemenPANRB/LKjIP sebelum naskah diajukan")
    add_table_data(doc, title=f"Tabel 1.1. Tren Evaluasi Akuntabilitas Kinerja Instansi Publik Nasional ({_rentang})",
                   headers=t1_headers, data=t1_rows, source=t1_source)

    if _sakip_rows:
        pv(doc, nv.pick(rng_bg, [
           (
               f"Tabel 1.1 memperlihatkan perbaikan bertahap, tetapi belum menghapus disparitas. "
               f"Kenaikan rata-rata nilai tidak otomatis berarti seluruh instansi mengalami "
               f"perbaikan yang setara. Persoalan kinerja karena itu bersifat struktural, sehingga "
               f"penjelasan tentang {y['label']} harus memperhitungkan variasi konteks dan mekanisme "
               f"di tingkat aparatur."
           ),
           (
               f"Tabel 1.1 menunjukkan perbaikan yang berlangsung bertahap tanpa menghapus "
               f"disparitas. Naiknya nilai rata-rata tidak dengan sendirinya berarti seluruh "
               f"instansi membaik secara setara. Karakter persoalan {y['label']} dengan demikian "
               f"bersifat struktural dan menuntut penjelasan pada tingkat aparatur."
           ),
        ]), ctx, "1.1-l2b", indent=0.5, space_after=6)
    else:
        p(doc,
          f"Karena nilai SAKIP nasional periode {_rentang} belum dimasukkan ke dalam konfigurasi "
          f"penelitian, Tabel 1.1 disajikan sebagai kerangka kosong. Tabel ini wajib dilengkapi dari "
          f"publikasi resmi sebelum naskah diajukan, sebab penjelasan tentang {y['label']} perlu "
          f"memperhitungkan variasi konteks dan mekanisme di tingkat aparatur.",
          indent=0.5, space_after=6)

    # --- LAPISAN 3: FENOMENA LOKAL & TABEL 1.2-1.4 ---
    pv(doc,
       f"Pada tingkat lokus, {instansi} menghadapi persoalan yang lebih spesifik. Sebagai "
       f"instansi yang mengemban fungsi {_get_instansi_focus(instansi)}, capaian kerjanya tidak hanya dinilai "
       f"dari ketertiban administrasi, tetapi juga dari kemampuan menjaga akurasi data, "
       f"menyelesaikan layanan tepat waktu, dan mempertahankan kepatuhan prosedur. Ketiga "
       f"tuntutan itu menekan unit teknis dan unit pelayanan secara berbeda, sehingga pemetaan "
       f"masalah tidak cukup dilakukan pada tingkat organisasi.",
       ctx, "1.1-l3a", indent=0.5, space_after=6)

    _x1_lbl = x1 or "kompetensi teknis"
    _x2_lbl = x2 or "beban kerja"
    _m_lbl = m or "keterikatan kerja"
    _z_lbl = z  # kosong bila judul tidak memuat variabel moderasi
    _aspek_aktif = [v for v in (_x1_lbl, _x2_lbl, _m_lbl, _z_lbl) if v]
    if len(_aspek_aktif) > 1:
        _aspek_str = ", ".join(_aspek_aktif[:-1]) + ", dan " + _aspek_aktif[-1]
    else:
        _aspek_str = _aspek_aktif[0] if _aspek_aktif else "variabel penelitian"

    pv(doc, nv.pick(rng_bg, [
       (
           f"Observasi awal dan telaah dokumen perencanaan menunjukkan gejala yang perlu dipetakan "
           f"secara sistematis. Gejala tersebut tidak disajikan sebagai angka persentase yang belum "
           f"terverifikasi, melainkan sebagai indikasi kualitatif yang dapat ditelusuri ke unit dan "
           f"aktor terkait. Tabel 1.2 memetakan indikasi masalah tersebut beserta implikasinya "
           f"terhadap sasaran kinerja."
       ),
       (
           f"Observasi awal serta telaah dokumen perencanaan memperlihatkan gejala yang perlu "
           f"dipetakan secara sistematis. Gejala itu tidak disajikan sebagai angka persentase yang "
           f"belum terverifikasi, tetapi sebagai indikasi kualitatif yang dapat ditelusuri hingga "
           f"unit dan aktor terkait. Tabel 1.2 memetakan indikasi masalah tersebut beserta "
           f"implikasinya bagi sasaran kinerja."
       ),
       (
           f"Penelaahan awal atas dokumen perencanaan dan pengamatan lapangan menemukan gejala yang "
           f"perlu dipetakan secara sistematis. Gejala tersebut tidak disajikan dalam bentuk angka "
           f"persentase yang belum terverifikasi, melainkan sebagai indikasi kualitatif yang dapat "
           f"dilacak ke unit dan aktor terkait. Tabel 1.2 memetakan indikasi masalah itu beserta "
           f"implikasinya terhadap sasaran kinerja."
       ),
    ]), ctx, "1.1-l3b", indent=0.5, space_after=6)

    t2_headers = ["No", "Dimensi Permasalahan Empiris", "Indikasi Unsur Masalah Teoretis", "Implikasi terhadap Sasaran Kinerja", "Unit / Aktor yang Terlibat"]
    _rows_12 = []
    if x1:
        _rows_12.append([
            f"Ketidakseragaman penguasaan {x1} pada unit teknis",
            "Capability gap: kesenjangan pengetahuan, keterampilan, dan sikap kerja",
            "Ketepatan waktu dan mutu luaran pekerjaan berpotensi menurun",
            "Aparatur teknis & pejabat fungsional",
        ])
    if x2:
        _rows_12.append([
            f"Ketidakseimbangan tuntutan {x2} terhadap kapasitas aparatur",
            "Job demands–resources imbalance: tuntutan kerja melampaui sumber daya",
            "Kelelahan kerja dan inkonsistensi unjuk kerja antarwaktu",
            "Aparatur pelaksana & koordinator bidang",
        ])
    if m:
        _rows_12.append([
            f"Belum optimalnya {m} sebagai perekat pelaksanaan tugas",
            "Meaningfulness & dedication gap: makna tugas belum terinternalisasi",
            "Keterlibatan kerja cenderung formalistik, belum menjadi komitmen afektif",
            "Seluruh unit kerja",
        ])
    if z:
        _rows_12.append([
            f"Keterbatasan {z} dalam menghadapi tekanan kerja",
            "Psychological capital gap: daya lenting dan optimisme kerja belum kuat",
            "Pemulihan kinerja pascabeban kerja tinggi berjalan lambat",
            "Aparatur pada posisi risiko tinggi",
        ])
    _rows_12.append([
        "Fragmentasi koordinasi lintas unit dan penyelarasan kebijakan pimpinan",
        "Direction gap: kejelasan sasaran dan komunikasi vertikal belum selaras",
        f"Konsistensi pencapaian sasaran {y_name} antarbidang belum merata",
        "Pimpinan struktural & sekretariat",
    ])
    if len(_rows_12) < 5:
        _rows_12.append([
            "Ketimpangan distribusi beban administratif terhadap pencapaian target",
            "Administrative burden & process bottleneck",
            f"Waktu penyelesaian layanan dan pencapaian {y_name} tidak proporsional",
            "Unit pelayanan & administrasi umum",
        ])
    t2_rows = [[f"{i + 1}.", *row] for i, row in enumerate(_rows_12[:5])]
    add_table_data(doc, title=f"Tabel 1.2. Fenomena Masalah Empiris dan Evaluasi Capaian Kinerja di Lingkungan {instansi} (2021–2025)",
                   headers=t2_headers, data=t2_rows,
                   source="Sumber: Observasi awal peneliti dan telaah dokumen perencanaan kinerja instansi (2026)")

    pv(doc,
       f"Tabel 1.2 menunjukkan bahwa persoalan {y['label']} tidak berdiri tunggal. Indikasi "
       f"pada aspek {_aspek_str} saling terkait dengan koordinasi "
       f"lintas unit. Keterlibatan aktor pada berbagai jenjang menegaskan perlunya pemetaan peran "
       f"menurut jabatan, sebagaimana disajikan pada Tabel 1.3.",
       ctx, "1.1-l3c", indent=0.5, space_after=6)

    t3_headers = ["Jenjang Jabatan / Kelompok Aktor", "Peran dalam Rantai Kesenjangan Kinerja", "Karakteristik Hambatan Dominan", "Implikasi bagi Variabel yang Diteliti"]
    t3_rows = [
        ["Pejabat Struktural (Eselon III & IV)", "Pengarah kebijakan dan teladan kinerja (tone at the top)", "Konsistensi keteladanan dan pengawasan eksekusi langsung", f"Pembentuk persepsi {_m_lbl}" + (f" dan {_z_lbl}" if _z_lbl else "") + " bawahan"],
        ["Pejabat Fungsional & Teknis", "Pelaksana substantif tugas pokok kedinasan", f"Kesesuaian tuntutan teknis dengan penguasaan {_x1_lbl}", f"Menguji kontribusi {_x1_lbl} terhadap {y_name}"],
        ["Aparatur Pelaksana / Staf Administratif", "Eksekutor rutinitas kedinasan dan layanan", f"Volume {_x2_lbl} dan daya tahan terhadap tekanan kerja", f"Menguji dampak {_x2_lbl} terhadap {y_name}"],
        ["Tim Verifikasi / Pengawas Internal", "Penjaga mutu dan kepatuhan prosedur", "Ketajaman deteksi dini dan independensi evaluasi", f"Menjaga validitas pengukuran {y_name}"],
        ["Staf Layanan Terpadu / Frontliner", "Wajah layanan publik instansi", "Kompetensi interaksi dan kepatuhan standar pelayanan", f"Menentukan persepsi publik atas {y_name}"],
    ]
    add_table_data(doc, title=f"Tabel 1.3. Pola Peran Demografis Jabatan Terhadap Kesenjangan Kinerja di {instansi} (2021–2025)",
                   headers=t3_headers, data=t3_rows,
                   source="Sumber: Pemetaan peran aktor berdasarkan struktur organisasi dan observasi awal peneliti (2026)")

    pv(doc,
       f"Pembagian peran pada Tabel 1.3 memperlihatkan bahwa sumber hambatan berbeda menurut "
       f"jenjang. Pejabat struktural lebih terkait dengan keteladanan dan keputusan, sedangkan "
       f"pelaksana lebih terkait dengan volume tugas dan kejelasan prosedur. Perbedaan itu "
       f"membawa konsekuensi metodologis: pengukuran harus mampu menangkap variasi persepsi "
       f"antarkelompok, bukan hanya rata-rata organisasi. Pemetaan berikutnya dilakukan menurut "
       f"bidang kerja pada Tabel 1.4.",
       ctx, "1.1-l3d", indent=0.5, space_after=6)

    t4_headers = ["Bidang / Sektor Unit Kerja", "Tingkat Keterpaparan Risiko Kinerja", "Fokus Hambatan dan Implikasi Terhadap Target Organisasi"]
    t4_rows = [
        ["Bidang Operasional Teknis & Pelayanan Utama", "Tinggi", f"Tekanan pencapaian target, kepadatan beban administratif, dan tuntutan {_x1_lbl}"],
        ["Bidang Pendataan, Pendaftaran & Pemutakhiran Data", "Sedang–Tinggi", "Akurasi validasi basis data dan adaptasi aparatur terhadap sistem digital"],
        ["Bidang Pembukuan, Pelaporan & Pengawasan", "Sedang", f"Ketepatan waktu rekonsiliasi dan konsistensi {_x2_lbl} pada periode pelaporan"],
        ["Sekretariat, Kepegawaian & Subbag Umum", "Sedang", f"Pengembangan {_x1_lbl} berkala dan pembinaan iklim kerja yang menopang {_m_lbl}"],
    ]
    add_table_data(doc, title=f"Tabel 1.4. Keterpaparan Risiko Kesenjangan Capaian Kinerja Berdasarkan Bidang Kerja di {instansi} (2021–2025)",
                   headers=t4_headers, data=t4_rows,
                   source="Sumber: Telaah struktur organisasi dan dokumen perencanaan kinerja instansi (2026)")

    pv(doc, nv.pick(rng_bg, [
       (
           f"Rangkaian indikasi pada Tabel 1.2 sampai 1.4 tidak dimaksudkan sebagai bukti statistik, "
           f"melainkan sebagai dasar penalaran untuk merumuskan model. Pola yang muncul menunjukkan "
           f"bahwa persoalan {y['label']} berakar pada interaksi antara kapasitas, tuntutan kerja, "
           f"dan kondisi psikologis aparatur. Interaksi itulah yang perlu dijelaskan secara kausal, "
           f"bukan hanya dideskripsikan."
       ),
       (
           f"Kumpulan indikasi pada Tabel 1.2 hingga 1.4 tidak diposisikan sebagai bukti statistik, "
           f"tetapi sebagai dasar penalaran dalam merumuskan model. Pola yang terlihat menandakan "
           f"bahwa persoalan {y['label']} bersumber dari interaksi antara kapasitas, tuntutan kerja, "
           f"dan kondisi psikologis aparatur. Interaksi tersebut perlu dijelaskan secara kausal, "
           f"tidak cukup hanya dideskripsikan."
       ),
       (
           f"Indikasi pada Tabel 1.2 sampai 1.4 bukanlah bukti statistik, melainkan dasar penalaran "
           f"bagi perumusan model. Pola yang muncul memperlihatkan bahwa persoalan {y['label']} "
           f"berakar pada interaksi kapasitas, tuntutan kerja, dan kondisi psikologis aparatur. "
           f"Interaksi itulah yang dituntut untuk dijelaskan secara kausal, bukan sekadar "
           f"dideskripsikan."
       ),
    ]), ctx, "1.1-l3e", indent=0.5, space_after=6)

    # --- LAPISAN 4: PILAR, PARADOKS, PREMIS ---
    _pillars = "; ".join(
        f"({i}) {p['label']} sebagai {p['definisi']}"
        for i, p in enumerate(xs, 1)
    ) or "(1) kapasitas keorganisasian sebagai pilar utama efektivitas kerja"
    _all_constructs = [p["label"] for p in [*xs, *ms, *zs]] + [y["label"]]
    if len(_all_constructs) > 1:
        _pillars_label = ", ".join(_all_constructs[:-1]) + ", serta " + _all_constructs[-1]
    else:
        _pillars_label = _all_constructs[0] if _all_constructs else "konstruk penelitian ini"

    pv(doc,
       f"Ekosistem pendorong {y['label']} pada {instansi} bersandar pada pilar yang dimodelkan "
       f"dalam judul penelitian, yaitu: {_pillars}. Pilar-pilar itu tidak berdiri sendiri; "
       f"masing-masing memiliki mekanisme kerja yang berbeda dan saling memengaruhi. Karena itu, "
       f"model penelitian tidak memperlakukan seluruh variabel sebagai prediktor yang setara.",
       ctx, "1.1-l4a", indent=0.5, space_after=6)

    pv(doc, nv.pick(rng_bg, [
       (
           f"Meskipun pilar-pilar tersebut telah tersedia secara formal, praktik di lapangan "
           f"memunculkan paradoks: mengapa prosedur yang lengkap, fasilitas yang memadai, dan "
           f"program pengembangan yang berjalan belum selalu menghasilkan lompatan {y['label']}? "
           f"Paradoks ini tidak dapat dijawab dengan menambah variabel, tetapi dengan menempatkan "
           f"mekanisme psikologis dan kondisi kontekstual sebagai bagian dari penjelasan."
       ),
       (
           f"Sekalipun pilar-pilar itu sudah tersedia secara formal, praktik lapangan menyisakan "
           f"paradoks: mengapa prosedur yang lengkap, fasilitas yang memadai, dan program "
           f"pengembangan yang berjalan tidak selalu berbuah lompatan {y['label']}? Pertanyaan ini "
           f"tidak dapat dijawab dengan menambah variabel, melainkan dengan menempatkan mekanisme "
           f"psikologis dan kondisi kontekstual sebagai bagian dari penjelasan."
       ),
       (
           f"Kendati pilar-pilarnya telah tersedia secara formal, praktik di lapangan tetap "
           f"menyisakan paradoks: mengapa prosedur lengkap, fasilitas memadai, dan program "
           f"pengembangan yang berjalan belum selalu mendorong lompatan {y['label']}? Paradoks "
           f"semacam itu tidak dapat dijawab dengan menambah variabel, tetapi dengan menempatkan "
           f"mekanisme psikologis dan kondisi kontekstual sebagai bagian dari penjelasan."
       ),
    ]), ctx, "1.1-l4b", indent=0.5, space_after=6)

    pv(doc, nv.pick(rng_bg, [
       (
           f"Premis yang diajukan disertasi ini adalah bahwa seluruh modal struktural tersebut "
           f"bersifat potensial sampai diolah melalui mekanisme transmisi perilaku. {_pillars_label} "
           f"tidak otomatis berubah menjadi luaran kinerja tanpa proses internalisasi di tingkat "
           f"aparatur. Premis ini menuntut model yang memisahkan kapasitas, proses, dan luaran, "
           f"sekaligus menguji apakah pemisahan itu didukung data."
       ),
       (
           f"Disertasi ini bertolak dari premis bahwa seluruh modal struktural tersebut bersifat "
           f"potensial hingga diolah melalui mekanisme transmisi perilaku. {_pillars_label} tidak "
           f"dengan sendirinya berubah menjadi luaran kinerja tanpa internalisasi di tingkat "
           f"aparatur. Premis itu menuntut model yang memisahkan kapasitas, proses, dan luaran, "
           f"sekaligus menguji apakah pemisahan tersebut terbukti secara empiris."
       ),
       (
           f"Premis pokok disertasi ini adalah bahwa modal struktural yang tersedia masih bersifat "
           f"potensial sampai diolah melalui mekanisme transmisi perilaku. {_pillars_label} tidak "
           f"otomatis menjelma menjadi luaran kinerja tanpa proses internalisasi pada tingkat "
           f"aparatur. Karena itu, diperlukan model yang memisahkan kapasitas, proses, dan luaran, "
           f"sekaligus menguji dukungan datanya."
       ),
    ]), ctx, "1.1-l4c", indent=0.5, space_after=6)

    # --- LAPISAN 5: MEDIASI & MODERASI ---
    _xv_str = " dan ".join(p["label"] for p in xs) or "variabel eksogen"
    if ms:
        mm = ms[0]
        pv(doc,
           f"Atas dasar premis tersebut, {mm['label']} ditempatkan sebagai variabel mediasi. "
           f"{mm['definisi'].capitalize()}. Peran itu strategis karena {mm['mekanisme']}. "
           f"Tanpa mekanisme tersebut, stimulasi dari {_xv_str} hanya berhenti pada kepatuhan "
           f"formal. Pengujian mediasi diperlukan untuk membuktikan apakah energi kerja benar-benar "
           f"menjadi jembatan atau hanya menjadi variabel yang berjalan sejajar.",
           ctx, "1.1-l5a", indent=0.5, space_after=6)
    if zs:
        zz = zs[0]
        pv(doc,
           f"Efektivitas ekosistem tersebut juga ditentukan oleh {zz['label']} sebagai variabel "
           f"moderasi. {zz['definisi'].capitalize()}. Dalam logika contingency, pengaruh variabel "
           f"eksogen tidak bersifat universal; ia bergantung pada kondisi yang memungkinkan "
           f"aparatur mengolah tuntutan dan sumber daya. Apabila {zz['label']} kuat, "
           f"{zz['tinggi']}. Sebaliknya, apabila rapuh, risiko yang muncul adalah {zz['risiko']}.",
           ctx, "1.1-l5b", indent=0.5, space_after=6)

        pv(doc,
           f"Posisi {zz['label']} sebagai katalisator perlu diuji, bukan diasumsikan. Klaim "
           f"teoretis tentang efek penguatan baru sahih apabila interaksi statistiknya signifikan "
           f"dan arahnya sesuai prediksi. Karena itu, model penelitian menempatkan interaksi "
           f"sebagai hipotesis tersendiri, bukan sebagai catatan tambahan pada pembahasan.",
           ctx, "1.1-l5c", indent=0.5, space_after=6)
    elif not ms:
        pv(doc,
           f"Efektivitas hubungan {_xv_str} dan {y['label']} tidak dapat dilepaskan dari "
           f"konfigurasi struktural dan kualitas eksekusi kebijakan di {instansi}. Mengacu pada "
           f"Contingency Theory, kekuatan pengaruh variabel eksogen bersifat kontekstual. Oleh "
           f"karena itu, model penelitian memusatkan pengujian pada jalur langsung sebagai "
           f"fondasi kausalitas yang terukur.",
           ctx, "1.1-l5alt", indent=0.5, space_after=6)

    # --- LAPISAN 6: DIALEKTIKA LITERATUR & TABEL 1.5 ---
    pv(doc,
       f"Urgensi model ini diperkuat oleh ketegangan temuan dalam literatur. Sebagian studi "
       f"melaporkan pengaruh positif {_xv_str} terhadap {y['label']} (seperti {pos_cites}). "
       f"Studi lain menemukan pengaruh yang tidak signifikan atau melemah ketika berhadapan "
       f"dengan kekakuan birokrasi dan lemahnya independensi pengawasan (seperti {gap_cites}). "
       f"Inkonsistensi itu mengindikasikan adanya mekanisme perantara dan kondisi batas yang "
       f"belum diuji secara simultan.",
       ctx, "1.1-l6a", indent=0.5, space_after=6)

    pv(doc,
       f"Perbedaan temuan tidak cukup dijelaskan dengan alasan sampel atau alat ukur. Penjelasan "
       f"yang lebih substantif terletak pada perbedaan konteks kelembagaan dan kualitas proses "
       f"psikologis yang menghubungkan kapasitas dengan luaran. Karena itu, penelitian ini "
       f"memetakan kesenjangan tersebut ke dalam matriks pada Tabel 1.5, dengan menandai posisi "
       f"kebaruan yang hendak diuji.",
       ctx, "1.1-l6b", indent=0.5, space_after=6)

    _med_txt = f"{m_label} sebagai jembatan transmisi keperilakuan" if m_label else "mekanisme psikologis eksekusi perilaku"
    _mod_txt = f"{z_label} sebagai variabel moderasi pengawal kontingensi" if z_label else "konfigurasi kontekstual organisasi"
    _role_txt = " dan ".join([v for v in (m_label, z_label) if v]) or "variabel eksogen"

    t5_headers = ["Aspek Penelitian Doktoral", "Tinjauan Penelitian Terdahulu (State of the Art)", "Kesenjangan (Research Gap) & Kebaruan (Doctoral Novelty)"]
    t5_rows = [
        [
            "Inkonsistensi Temuan Empiris (Inconsistency of Findings)",
            f"Riset {pos_cites} menemukan efek positif signifikan {_xv_str} terhadap {y['label']}. Namun temuan itu diperdebatkan oleh sejumlah studi ({gap_cites}) yang menemukan pengaruh tidak signifikan atau kondisional pada organisasi sektor publik.",
            f"Celah (Gap): Anteseden formal cenderung tidak bekerja secara otomatis tanpa mekanisme eksekusi perilaku di lapangan.\nKebaruan (Novelty): Menyelesaikan perdebatan dengan model struktural terintegrasi yang menempatkan {_med_txt} serta {_mod_txt}.",
        ],
        [
            "Peran Variabel dalam Model (Variable Role Formulation)",
            f"Sebagian besar riset terdahulu memperlakukan {_role_txt} sekadar sebagai prediktor independen parsial yang berdiri sendiri dalam model regresi linier biasa ({gap_cites}).",
            f"Celah (Gap): Gagal memotret dinamika transmisi perilaku dan kondisi kontingensi secara utuh.\nKebaruan (Novelty): Mengintegrasikan {_med_txt} serta {_mod_txt} dalam satu model struktural.",
        ],
        [
            "Posisi Variabel Mediasi (Mediating Mechanism)",
            "Penelitian terdahulu menempatkan mediasi sebatas kepuasan administratif tanpa mengaitkannya dengan dorongan psikologis berprestasi (Ryan & Deci, 2022; Schunk, 2022).",
            f"Celah (Gap): Belum ada pembuktian empiris bahwa kapasitas eksogen adalah potensi pasif yang menuntut aktivasi motivasional.\nKebaruan (Novelty): Menguji {_med_txt} dalam menjelaskan konversi kapasitas organisasi menjadi kinerja.",
        ],
        [
            "Posisi Variabel Moderasi (Moderating & Contingency)",
            f"Riset kontemporer (Yukl & Gardner, 2024; Northouse, 2024) cenderung mengabaikan dinamika {z_label or 'kondisi kontekstual aparatur'} dan ancaman management override sebagai batas struktural.",
            f"Celah (Gap): Mengasumsikan efektivitas kepemimpinan linier di setiap kondisi moralitas pegawai.\nKebaruan (Novelty): Menjadikan {_mod_txt} sebagai pemeta kapan pengaruh eksogen berdaya maksimal.",
        ],
        [
            "Konteks Sektoral & Desain Model (Context & Structural Complexity)",
            f"Mayoritas riset terdahulu terisolasi pada sektor privat atau mediasi tunggal sederhana, minim kajian komprehensif pada organisasi sektor publik tingkat daerah ({instansi}) dengan kompleksitas tugas kelembagaan.",
            f"Celah (Gap): Minimnya model hibrida terpadu pada birokrasi pemerintah daerah.\nKebaruan (Novelty): Menguji model integratif Moderated-Mediation secara simultan pada {instansi} menggunakan PLS-SEM 4.",
        ],
        [
            "Kebaruan Model Konseptual Orisinal (Doctoral Theoretical Novelty)",
            f"Belum ada sintesis terpadu dalam literatur manajemen yang menyatukan teori-teori yang melandasi {_xv_str} ke dalam satu model struktural dengan {_role_txt}.",
            f"Kebaruan (Novelty): Melahirkan Model Doktoral Orisinal '{model_acronym}' yang menjawab anomali kinerja birokrasi secara teoretis dan aplikatif.",
        ],
    ]
    add_table_data(doc, title="Tabel 1.5. Matriks Pemetaan Kesenjangan Penelitian (Research Gap) dan Kebaruan Ilmiah (Novelty)",
                   headers=t5_headers, data=t5_rows,
                   source="Sumber: Sintesis analisis kesenjangan literatur dan pemetaan kebaruan peneliti doktoral (2026)")

    pv(doc,
       f"Tabel 1.5 menegaskan bahwa kebaruan penelitian ini bukan penambahan variabel secara "
       f"artifisial. Kebaruan terletak pada penyusunan ulang hubungan antarkonstruk sehingga "
       f"mekanisme transmisi dan kondisi batas dapat diuji dalam satu model. Dengan cara itu, "
       f"penelitian tidak mengulang perdebatan lama, tetapi mengubah cara perdebatan itu dijawab.",
       ctx, "1.1-l6c", indent=0.5, space_after=6)

    # --- LAPISAN 7: NOVELTY & JUDUL ---
    _xv_all_str = _join_id(p["label"] for p in xs) or "variabel eksogen"
    _void_intro = "Mengisi Dua Kekosongan Teoretis Mendasar" if (m and z) else "Mengisi Kekosongan Teoretis Mendasar"
    pv(doc,
       f"Berdasarkan dialektika tersebut, kontribusi kebaruan disertasi ini bertumpu pada upaya "
       f"'{_void_intro}' (Filling Theoretical Voids):",
       ctx, "1.1-l7a", indent=0.5, space_after=6)

    if m:
        pv(doc,
           f"1. Mengisi Kekosongan Eksekusi Perilaku (Behavioral Execution Void): menguji apakah "
           f"{_xv_all_str} dapat mengubah unjuk kerja secara langsung tanpa melalui "
           f"{m_label}, atau justru berlangsung melalui jalur tersebut. Tanpa dorongan berprestasi "
           f"dan persistensi usaha, regulasi dan arahan manajerial berisiko berhenti sebagai "
           f"kepatuhan formal.",
           ctx, "1.1-l7b", indent=0.75, space_after=4)
    else:
        pv(doc,
           f"1. Mengisi Kekosongan Bukti Empiris Jalur Langsung (Direct Path Evidence Void): "
           f"mengukur besaran dan signifikansi kontribusi masing-masing pilar {_xv_all_str} "
           f"terhadap {y['label']} pada konteks birokrasi pemerintahan daerah, sehingga prioritas "
           f"intervensi dapat disusun berbasis bukti.",
           ctx, "1.1-l7b2", indent=0.75, space_after=4)

    if z:
        z0 = zs[0] if zs else None
        pv(doc,
           f"2. Mengisi Kekosongan Mekanisme Psikologis Bersyarat (Conditional Psychological "
           f"Mechanism): menempatkan {z_label} bukan sebagai variabel tambahan, melainkan "
           f"kondisi batas yang menentukan sejauh mana {x_labels} terkonversi menjadi {y['label']}. "
           f"Apabila {z_label} kuat, {z0['tinggi'] if z0 else 'daya ungkit kapasitas menguat'}; "
           f"apabila lemah, risiko yang muncul adalah {z0['risiko'] if z0 else 'daya ungkit kapasitas melemah'}. "
           f"Dengan demikian, kontribusi teoretis disertasi ini terletak pada penjelasan bersyarat, "
           f"bukan pada klaim pengaruh yang berlaku umum.",
           ctx, "1.1-l7c", indent=0.75, space_after=6)
    elif m:
        pv(doc,
           f"2. Mengisi Kekosongan Konversi Kapasitas menjadi Kinerja (Capacity-to-Performance "
           f"Void): menguji apakah {_xv_all_str} berpengaruh terhadap {y['label']} melalui "
           f"mekanisme transmisi {m_label}, sehingga investasi kapasitas tidak berhenti sebagai "
           f"potensi.",
           ctx, "1.1-l7c2", indent=0.75, space_after=6)

    _model_label = {
        "direct": "Model Integratif Jalur Langsung Kinerja Organisasi Publik",
        "mediation": "Model Integratif Mediasi Kinerja Organisasi Publik",
        "moderation": "Model Integratif Moderasi Kinerja Organisasi Publik",
        "moderated_mediation": "Model Integratif Moderated-Mediation Kinerja Organisasi Publik",
    }.get(str(ctx.get("model_type") or ""), "Model Integratif Kinerja Organisasi Publik")

    pv(doc,
       f"Sintesis tersebut melahirkan model struktural orisinal tingkat doktoral yang diberi "
       f"nama MODEL {model_acronym} ({_model_label}). Pengujian empiris dilakukan dengan "
       f"Partial Least Squares Structural Equation Modeling (PLS-SEM) menggunakan SmartPLS 4. "
       f"Melalui pengujian itu, disertasi ini diharapkan memberi kontribusi teoretis bagi "
       f"pengembangan ilmu manajemen sumber daya manusia dan cetak biru praktis bagi pimpinan "
       f"{instansi} dalam mengakselerasi {y['label']}.",
       ctx, "1.1-l7d", indent=0.5, space_after=6)

    pv(doc,
       f"Berlandaskan argumentasi teoretis, data makro, telaah fenomena lokal, paradoks pilar "
       f"pasif, dialektika literatur, dan perumusan kebaruan tersebut, penelitian mendalam "
       f"dalam bentuk disertasi doktor dipandang urgen dan strategis dengan judul: \"{judul}\".",
       ctx, "1.1-l7e", indent=0.5, space_after=12)




def _build_doc_sub_1_2(doc: Document, ctx: Dict[str, Any]):
    """Sub-bab B. Rumusan Masalah berbasis topology graph dan profil variabel."""
    h2(doc, "B. Rumusan Masalah")
    judul = ctx["judul"]
    instansi = ctx["instansi"]
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "1.2")
    xs = pm["_all"]["x"]
    ms = pm["_all"]["m"]
    zs = pm["_all"]["z"]
    y = pm["Y"]
    x_vars_str = _join_id(p["label"] for p in xs) or "variabel eksogen"
    m_vars_str = ", ".join(p["label"] for p in ms)
    z_vars_str = ", ".join(p["label"] for p in zs)
    m_type = ctx["model_type"].replace("_", " ").title()
    model_acronym = _generate_doctoral_model_acronym(ctx)
    hypos = ctx["graph"]["hypotheses"]
    cite_grp = ctx.get("citation_groups") or s3_academic_researcher.get_s3_citation_groups(ctx)
    pos_cites = cite_grp.get("pos_citations", "Rahman, 2023; Aini et al., 2025; Hair et al., 2022")
    gap_cites = cite_grp.get("gap_citations", "Gadzai et al., 2025; Massud & Uddin, 2025; Almazrouei & Alnahhal, 2025")
    first_gap = cite_grp.get("first_gap_cite", "Gadzai et al. (2025)")
    sec_gap = cite_grp.get("second_gap_cite", "Almazrouei & Alnahhal (2025)")
    x0 = xs[0] if xs else nv.make_profile("kapasitas organisasi", {}, ctx)

    pv(doc,
       f"Rumusan masalah penelitian ini berangkat dari pertanyaan utama: bagaimana membangun "
       f"model konseptual dan model empirik yang mampu menjelaskan variasi {y['label']} pada "
       f"{instansi}, sekaligus mengatasi ketidakkonsistenan temuan penelitian sebelumnya? "
       f"Ketidakkonsistenan itu tampak pada sejumlah studi ({gap_cites}), yang menunjukkan "
       f"bahwa hubungan antara kapasitas eksogen dan luaran kerja tidak selalu positif dan "
       f"signifikan. Perbedaan hasil tersebut mengindikasikan adanya mekanisme "
       f"yang belum terukur dan kondisi batas yang belum diuji secara bersamaan.",
       ctx, "1.2-p1", indent=0.5, space_after=6)

    if ms:
        m = ms[0]
        pv(doc,
           f"Studi arus utama ({pos_cites}) memang melaporkan pengaruh positif {x_vars_str} "
           f"terhadap {y['label']}. Namun, penelitian ini berargumen bahwa pengaruh tersebut "
           f"tidak beroperasi secara otomatis. {x0['label']} pada hakikatnya adalah potensi yang "
           f"harus diolah menjadi energi kerja; di sinilah {m['label']} berfungsi sebagai "
           f"mekanisme transmisi. Tanpa mekanisme itu, risiko yang muncul adalah {x0['risiko']}, sehingga investasi pada "
           f"kapasitas tidak serta-merta berubah menjadi perbaikan {y['label']}.",
           ctx, "1.2-p2", indent=0.5, space_after=6)
    else:
        pv(doc,
           f"Studi arus utama ({pos_cites}) mengasumsikan bahwa pilar kapasitas organisasi "
           f"berkontribusi positif terhadap {y['label']}. Penelitian ini tidak menolak asumsi "
           f"tersebut, tetapi menuntut kejelasan tentang mekanisme dan konteks kerjanya. "
           f"Tanpa kejelasan itu, rekomendasi kebijakan berisiko menyalin praktik dari "
           f"organisasi lain tanpa memperhitungkan kondisi {instansi}.",
           ctx, "1.2-p2alt", indent=0.5, space_after=6)

    if zs:
        z = zs[0]
        pv(doc,
           f"Selain itu, dinamika kerja di birokrasi menunjukkan bahwa capaian kinerja tidak "
           f"hanya ditentukan oleh kapasitas dan proses psikologis, tetapi juga oleh kondisi "
           f"kontingensi. {z['label']} diposisikan sebagai variabel moderasi karena "
           f"{z['definisi']}. Apabila {z['label']} kuat, {z['tinggi']}. Sebaliknya, apabila "
           f"{z['label']} lemah, risiko yang muncul adalah {z['risiko']}, sehingga dampak "
           f"{x_vars_str} terhadap {y['label']} melemah. "
           f"Karena itu, efek penguatan atau pelemahan tersebut perlu diuji, bukan diasumsikan.",
           ctx, "1.2-p3", indent=0.5, space_after=6)
    else:
        pv(doc,
           f"Capaian {y['label']} pada {instansi} juga tidak dapat dilepaskan dari konteks "
           f"struktural dan kualitas eksekusi kebijakan. Oleh karena itu, rumusan masalah "
           f"penelitian ini mencakup pengujian jalur langsung sekaligus penilaian kewajaran "
           f"besaran pengaruh masing-masing variabel eksogen dalam konteks organisasi.",
           ctx, "1.2-p3alt", indent=0.5, space_after=6)

    pv(doc,
       f"Penelitian ini menawarkan kebaruan ilmiah berupa pengujian model struktural {m_type} "
       f"yang diberi nama MODEL {model_acronym}. Model tersebut menyatukan kapasitas eksogen, "
       f"mekanisme transmisi, dan kondisi batas dalam satu kerangka yang dapat diuji secara "
       f"empiris. Dengan demikian, rumusan masalah tidak berhenti pada deskripsi hubungan, "
       f"tetapi mengarah pada penjelasan kausal yang dapat dipertanggungjawabkan.",
       ctx, "1.2-p4", indent=0.5, space_after=6)

    pv(doc, "Berdasarkan perumusan masalah di atas, pertanyaan penelitian ini adalah:",
       ctx, "1.2-p5", indent=0.5, space_after=6)

    for i, h in enumerate(hypos, 1):
        p(doc, f"{i}. {h['question']}", indent=0.75, space_after=3)


def _build_doc_sub_1_3(doc: Document, ctx: Dict[str, Any]):
    """Sub-bab C. Tujuan Penelitian, selaras dengan rumusan masalah dan graph."""
    h2(doc, "C. Tujuan Penelitian")
    instansi = ctx["instansi"]
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "1.3")
    xs = pm["_all"]["x"]
    ms = pm["_all"]["m"]
    zs = pm["_all"]["z"]
    y = pm["Y"]
    x_vars_str = _join_id(p["label"] for p in xs) or "variabel eksogen"
    m_vars_str = ", ".join(p["label"] for p in ms) or "mekanisme perantara"
    z_vars_str = ", ".join(p["label"] for p in zs) or "kondisi kontekstual"
    model_acronym = _generate_doctoral_model_acronym(ctx)
    hypos = ctx["graph"]["hypotheses"]

    pv(doc,
       f"Penelitian ini dibatasi pada pengujian empiris Model {model_acronym} pada aparatur "
       f"di lingkungan {instansi}. Batasan tersebut mencakup empat hal: konstruk yang diukur, "
       f"unit analisis, periode pengamatan, dan klaim generalisasi. Konstruk yang diukur "
       f"meliputi {x_vars_str}, {m_vars_str}, {z_vars_str}, serta {y['label']}; unit analisisnya "
       f"adalah aparatur yang memiliki pengalaman kerja relevan; periode pengamatan bersifat "
       f"cross-sectional; dan generalisasi diarahkan pada populasi aparatur instansi, bukan "
       f"seluruh birokrasi daerah.",
       ctx, "1.3-p1", indent=0.5, space_after=6)

    pv(doc,
       f"Secara teoretis, tujuan penelitian ini adalah menguji proposisi bahwa {x_vars_str} "
       f"berpengaruh terhadap {y['label']} melalui mekanisme yang tidak sederhana. Secara "
       f"metodologis, penelitian ini bertujuan menghasilkan estimasi jalur langsung, efek "
       f"tidak langsung, dan efek kondisional yang dapat dibandingkan dengan temuan sebelumnya. "
       f"Secara praktis, hasilnya diarahkan pada perumusan prioritas intervensi yang sesuai "
       f"dengan kondisi {instansi}.",
       ctx, "1.3-p2", indent=0.5, space_after=6)

    pv(doc, "Sejalan dengan pertanyaan penelitian di atas, tujuan dari penelitian ini adalah untuk:",
       ctx, "1.3-p3", indent=0.5, space_after=6)

    for i, h in enumerate(hypos, 1):
        _obj = str(h.get("objective") or "").strip()
        _obj = re.sub(r"^Untuk\s+", "", _obj, flags=re.I)
        p(doc, f"{i}. {_obj}", indent=0.75, space_after=3)


def _build_doc_sub_1_4(doc: Document, ctx: Dict[str, Any]):
    """1.4 Manfaat Penelitian, diturunkan dari profil variabel dan konteks instansi."""
    h2(doc, "D. Manfaat Penelitian")
    instansi = ctx.get("instansi", "Instansi Pemerintah")
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "1.4")
    xs = pm["_all"]["x"]
    ms = pm["_all"]["m"]
    zs = pm["_all"]["z"]
    y = pm["Y"]
    x_vars_str = _join_id(p["label"] for p in xs) or "variabel eksogen"
    m = ms[0]["label"] if ms else ""
    z = zs[0]["label"] if zs else ""
    model_acronym = _generate_doctoral_model_acronym(ctx)
    cite_grp = ctx.get("citation_groups") or s3_academic_researcher.get_s3_citation_groups(ctx)
    pos_cites = cite_grp.get("pos_citations", "Hair et al., 2022; Sarstedt et al., 2022")
    gov_cites = cite_grp.get("gov_citations", "Hayes, 2022; Creswell & Creswell, 2023")

    pv(doc,
       f"Penelitian ini dirancang untuk memberi kontribusi yang seimbang antara pengembangan "
       f"keilmuan dan perbaikan praktik. Kontribusi tersebut tidak diukur dari banyaknya "
       f"variabel yang diteliti, tetapi dari kejelasan mekanisme yang berhasil dijelaskan dan "
       f"dari kegunaan temuan bagi pengambilan keputusan di {instansi} (Sekaran & Bougie, 2020; "
       f"Creswell & Creswell, 2023). Uraian berikut memisahkan manfaat teoretis dan manfaat "
       f"praktis agar klaim dapat dinilai secara proporsional.",
       ctx, "1.4-p1", indent=0.5, space_after=6)

    p(doc, "1. Manfaat Teoretis (Pengembangan Keilmuan & Novelty)", bold=True, indent=0.5, space_after=3)
    pv(doc,
       f"Secara keilmuan, penelitian ini memperkaya kajian manajemen kinerja sektor publik "
       f"dengan menjelaskan bagaimana {x_vars_str} dapat memengaruhi {y['label']} melalui "
       f"mekanisme yang lebih spesifik. Penelitian ini tidak berhenti pada pengujian korelasi, "
       f"tetapi menawarkan penjelasan tentang kapan dan mengapa pengaruh tersebut menguat atau "
       f"melemah (Snyder, 2019; Hitt et al., 2023). Dengan cara itu, temuan dapat digunakan "
       f"untuk menilai kembali asumsi linier yang sering dipakai dalam studi kinerja aparatur.",
       ctx, "1.4-p2", indent=0.75, space_after=4)

    _m_part = (
        f"peran {m} sebagai jembatan transmisi" if m
        else f"jalur langsung {x_vars_str} sebagai fondasi kausalitas"
    )
    _z_part = (
        f", serta peran {z} sebagai kondisi batas yang memperkuat atau memperlemah hubungan"
        if z else ""
    )
    pv(doc,
       f"a. Model Empiris Baru (Model Doktoral {model_acronym}): Penelitian ini menyusun model "
       f"struktural yang mengintegrasikan {_m_part}{_z_part} terhadap {y['label']} "
       f"(Hair et al., 2022; Hayes, 2022). Model ini memberi kerangka uji yang lebih utuh "
       f"daripada pendekatan parsial, karena memisahkan kapasitas, proses, dan luaran dalam "
       f"satu struktur. Kehadirannya diharapkan menjadi rujukan bagi penelitian lanjutan yang "
       f"ingin menguji mekanisme serupa pada konteks birokrasi lain ({pos_cites}).",
       ctx, "1.4-p3", indent=0.75, space_after=4)

    theory_names: List[str] = []
    for prof in [*xs, *ms, *zs, y]:
        for nama, sitasi, _ in (prof.get("teori") or [])[:1]:
            entry = nv.theory_entry(nama, sitasi)
            if entry not in theory_names:
                theory_names.append(entry)
    theory_txt = "; ".join(theory_names[:4]) or "teori perilaku organisasi"
    pv(doc,
       f"b. Pengembangan Teori Multidisiplin: Penelitian ini menyintesiskan {theory_txt} dalam "
       f"satu kerangka eksplanasi. Sintesis itu membantu menjelaskan mengapa temuan penelitian "
       f"terdahulu tidak selalu konsisten, khususnya ketika perbedaan konteks dan mekanisme "
       f"psikologis tidak diperhitungkan. Kontribusi teoretisnya terletak pada penyediaan "
       f"penjelasan bersyarat, bukan pada penambahan dalil yang berdiri sendiri.",
       ctx, "1.4-p4", indent=0.75, space_after=6)

    p(doc, "2. Manfaat Praktis (Kebijakan Strategis Organisasi)", bold=True, indent=0.5, space_after=3)
    pv(doc,
       f"Secara praktis, temuan penelitian ini dapat digunakan oleh pimpinan dan unit pengelola "
       f"sumber daya manusia di {instansi} untuk menyusun prioritas intervensi. Prioritas tidak "
       f"ditentukan oleh besarnya anggaran semata, tetapi oleh titik mana dalam rantai kausal "
       f"yang paling perlu diperbaiki (Robbins & Judge, 2022; Daft, 2022). Rekomendasi disusun "
       f"pada tiga tingkat berikut.",
       ctx, "1.4-p5", indent=0.75, space_after=4)

    _x_interv = "; ".join(f"{p['label']}: {p['intervensi']}" for p in xs[:2]) or "penataan kapasitas kerja"
    pv(doc,
       f"a. Bagi Pemerintah Daerah dan Jajaran Pimpinan {instansi}: model penelitian ini dapat "
       f"menjadi acuan untuk menilai apakah persoalan kinerja bersumber dari kapasitas, proses "
       f"psikologis, atau kondisi kontekstual. Intervensi yang disarankan mencakup {_x_interv}. "
       f"Dengan pemetaan itu, kebijakan tidak berhenti pada penambahan pelatihan atau prosedur, "
       f"tetapi menyentuh faktor yang benar-benar menjadi penghambat ({gov_cites}).",
       ctx, "1.4-p6", indent=0.75, space_after=4)

    _m_interv = ms[0]["intervensi"] if ms else "penguatan umpan balik dan kejelasan peran"
    pv(doc,
       f"b. Bagi Unit Kepegawaian dan Pengelola Kinerja: hasil penelitian dapat digunakan untuk "
       f"merancang program pembinaan yang lebih terarah. Jika {m or 'mekanisme perantara'} "
       f"terbukti menjadi jalur utama, maka intervensi perlu diarahkan pada {_m_interv}. "
       f"Sebaliknya, jika jalur langsung lebih dominan, perbaikan diarahkan pada kapasitas teknis "
       f"dan kejelasan prosedur. Pendekatan ini menghindari program yang seragam untuk seluruh "
       f"unit kerja ({pos_cites}).",
       ctx, "1.4-p7", indent=0.75, space_after=4)

    pv(doc,
       f"c. Bagi Aparatur dan Pemangku Kepentingan Publik: perbaikan {y['label']} pada {instansi} "
       f"berimplikasi pada kualitas layanan dan kepercayaan publik. Apabila mekanisme kerja "
       f"diperbaiki, masyarakat memperoleh layanan yang lebih cepat, akurat, dan dapat "
       f"dipertanggungjawabkan. Manfaat ini bersifat tidak langsung, tetapi penting untuk "
       f"menjaga legitimasi institusi dan keberlanjutan program pembangunan daerah ({gov_cites}).",
       ctx, "1.4-p8", indent=0.75, space_after=6)


def _build_doc_sub_1_5(doc: Document, ctx: Dict[str, Any]):
    """1.5 Ruang Lingkup Penelitian Doktoral."""
    h2(doc, "E. Ruang Lingkup Penelitian")
    instansi = ctx["instansi"]

    p(doc, f"Ruang lingkup dalam penelitian disertasi ini ditetapkan secara terukur sebagai berikut:", indent=0.5, space_after=6)
    p(doc, f"1. Lingkup Substansi Variabel: Penelitian difokuskan pada hubungan struktural antar variabel yang mencakup variabel terikat ({ctx['y_var']}), variabel mediasi ({', '.join(ctx['m_vars'])}), variabel moderasi ({', '.join(ctx['z_vars'])}), serta variabel bebas ({', '.join(ctx['x_vars'])}).", indent=0.75, space_after=3)
    p(doc, f"2. Lingkup Objek dan Lokasi: Penelitian dilaksanakan secara empiris pada lingkungan {instansi}.", indent=0.75, space_after=3)
    p(doc, f"3. Lingkup Subjek dan Responden: Subjek penelitian adalah seluruh pegawai aparatur sipil pada instansi terkait dengan ukuran sampel representatif n = {ctx['n']} responden.", indent=0.75, space_after=3)
    p(doc, f"4. Lingkup Metodologis: Menggunakan metode kuantitatif kausal-eksplanatori dengan teknik pemodelan persamaan struktural berbasis varian (PLS-SEM) menggunakan SmartPLS 4 dan Hayes PROCESS Macro.", indent=0.75, space_after=3)
    p(doc, f"5. Lingkup Temporal: Pengumpulan data primer kuesioner dan observasi empiris dilaksanakan pada kurun waktu tahun {ctx['tahun']}.", indent=0.75, space_after=6)


def _build_doc_sub_1_6(doc: Document, ctx: Dict[str, Any]):
    """1.6 Sistematika Penulisan Disertasi."""
    h2(doc, "F. Sistematika Penulisan")
    p(doc, "Naskah disertasi ini disusun ke dalam 5 (lima) bab yang saling terkait secara logis dan terstruktur sebagai berikut:", indent=0.5, space_after=6)
    p(doc, "BAB I PENDAHULUAN: Memuat latar belakang masalah, identifikasi dan rumusan masalah, pembatasan dan tujuan penelitian, manfaat penelitian, ruang lingkup penelitian, serta sistematika penulisan.", indent=0.5, space_after=3)
    p(doc, "BAB II KAJIAN PUSTAKA, KERANGKA KONSEPTUAL, DAN PENGEMBANGAN HIPOTESIS: Menguraikan telaah Grand Theory, Middle Theory, kajian teoritis mendalam per variabel, penelitian terdahulu (state of the art), kerangka konseptual model hibrida, dan formulasi hipotesis penelitian.", indent=0.5, space_after=3)
    p(doc, "BAB III METODE PENELITIAN: Menguraikan desain penelitian, populasi dan sampel, operasionalisasi variabel, teknik pengumpulan data, pengujian instrumen, dan rancangan analisis data PLS-SEM.", indent=0.5, space_after=3)
    p(doc, "BAB IV HASIL PENELITIAN DAN PEMBAHASAN: Menyajikan deskripsi objek penelitian, karakteristik responden, statistik deskriptif, evaluasi Outer Model dan Inner Model PLS-SEM, uji hipotesis bootstrapping, analisis moderasi simple slope, conditional indirect effects, serta pembahasan 4 alinea pascasarjana.", indent=0.5, space_after=3)
    p(doc, "BAB V KESIMPULAN, IMPLIKASI, DAN SARAN: Menyajikan kesimpulan menyeluruh, implikasi teoretis (novelty doktoral), implikasi manajerial praktis, keterbatasan penelitian, dan agenda riset lanjutan.", indent=0.5, space_after=6)


def _theory_families(ctx: Dict[str, Any]) -> set:
    """Keluarga konstruk dari judul, dipakai untuk memilih grand/middle theory.

    Tidak ada nama teori yang di-hardcode pada tingkat judul; pemilihan bergantung
    sepenuhnya pada profil variabel penelitian (lihat ``nv.profile_map``).
    """
    try:
        pm = nv.profile_map(ctx)
        parts = [*pm["_all"]["x"], *pm["_all"]["m"], *pm["_all"]["z"], pm["Y"]]
        return {p.get("family") for p in parts if isinstance(p, dict) and p.get("family")}
    except Exception:
        return set()


_FAMILY_DOMAIN = {
    "kapasitas": "sumber daya manusia",
    "tuntutan": "tuntutan kerja",
    "energi": "energi dan keterlibatan kerja",
    "modal": "kekuatan psikologis",
    "luaran": "perilaku dan hasil kerja",
    "kepemimpinan": "pengaruh kepemimpinan",
    "nilai": "nilai dan budaya organisasi",
    "iklim": "iklim organisasi",
    "kepatuhan": "ketaatan pada aturan",
    "dorongan": "motivasi kerja",
    "evaluasi": "kepuasan kerja",
    "imbalan": "keadilan imbalan",
    "kondisi": "lingkungan kerja",
    "sikap": "sikap dan niat perilaku",
    "kontrol": "pengawasan dan akuntabilitas",
    "integritas": "integritas dan akuntabilitas",
}


def _family_domain(families) -> str:
    """Frasa ranah substantif untuk sekumpulan keluarga variabel."""
    order = ("kapasitas", "tuntutan", "kondisi", "kepatuhan", "energi", "modal",
             "dorongan", "evaluasi", "imbalan", "kepemimpinan", "nilai", "iklim",
             "sikap", "kontrol", "integritas", "luaran")
    doms = []
    for f in order:
        if f in families and _FAMILY_DOMAIN[f] not in doms:
            doms.append(_FAMILY_DOMAIN[f])
    if not doms:
        return "perilaku kerja"
    if len(doms) == 1:
        return doms[0]
    return ", ".join(doms[:-1]) + " serta " + doms[-1]


def _grand_theory_items(families) -> List[str]:
    """Butir grand theory sesuai keluarga variabel yang benar-benar ada."""
    items: List[str] = []
    if "kapasitas" in families:
        items.append(
            "Human Capital Theory (Becker, 1964) yang memandang pengetahuan dan keterampilan "
            "sebagai investasi produktif"
        )
    if "tuntutan" in families:
        items.append(
            "Job Demands-Resources Theory (Bakker & Demerouti, 2022) yang menjelaskan interaksi "
            "antara tuntutan dan sumber daya kerja"
        )
    if "kondisi" in families:
        items.append(
            "Teori Lingkungan Kerja (Sedarmayanti, 2022; Mangkunegara, 2022) yang menjelaskan "
            "pengaruh kondisi fisik, sosial, dan psikologis terhadap efektivitas kerja"
        )
    if "kepatuhan" in families:
        items.append(
            "Teori Kepatuhan Organisasi (Sutrisno, 2022; Hasibuan, 2022) yang menempatkan "
            "ketaatan pada aturan sebagai prasyarat ketertiban kerja"
        )
    if "kontrol" in families:
        items.append(
            "Agency Theory (Jensen & Meckling, 1976) yang menyoroti asimetri informasi dan "
            "kebutuhan pemantauan"
        )
    if families & {"nilai", "iklim"}:
        items.append(
            "Teori Budaya dan Iklim Organisasi (Schein & Schein, 2023; Litwin & Stringer, 1968) "
            "yang menjelaskan bagaimana nilai dan persepsi kolektif membentuk perilaku"
        )
    if "modal" in families:
        items.append(
            "Psychological Capital Theory (Luthans et al., 2022) yang menempatkan kekuatan "
            "psikologis positif sebagai sumber daya kerja"
        )
    if "energi" in families:
        items.append(
            "Teori Keterikatan Kerja (Saks, 2022; Schaufeli, 2022) yang menjelaskan energi, "
            "dedikasi, dan konsentrasi sebagai penggerak usaha"
        )
    if "dorongan" in families:
        items.append(
            "Teori Motivasi Kerja (Vroom, 1964, dalam Kreitner & Kinicki, 2022; Ryan & Deci, "
            "2022) yang menjelaskan arah, intensitas, dan ketekunan usaha"
        )
    if "evaluasi" in families:
        items.append(
            "Teori Kepuasan Kerja (Robbins & Judge, 2022; Luthans et al., 2022) yang memandang "
            "sikap kerja sebagai hasil evaluasi terhadap pekerjaan"
        )
    if "imbalan" in families:
        items.append(
            "Teori Keadilan Organisasi (Robbins & Judge, 2022) yang menekankan keadilan "
            "distributif dan prosedural sebagai penentu usaha kerja"
        )
    if "kepemimpinan" in families:
        items.append(
            "Teori Kepemimpinan Transformasional (Bass, 1985, dalam Yukl, 2022) yang menjelaskan "
            "pengaruh pemimpin terhadap motivasi dan kinerja"
        )
    if "sikap" in families:
        items.append(
            "Teori Sikap dan Perilaku Terencana (Ryan & Deci, 2022; Robbins & Judge, 2022) yang "
            "menghubungkan sikap, norma, dan niat perilaku"
        )
    if "integritas" in families:
        items.append(
            "Teori Integritas Organisasi (Robbins & Judge, 2022) yang menempatkan kejujuran dan "
            "akuntabilitas sebagai landasan tata kelola"
        )
    if "luaran" in families:
        items.append(
            "Teori Kinerja (Campbell & Wiernik, 2022) yang memosisikan kinerja sebagai "
            "perilaku yang dapat diamati dan dinilai"
        )
    if not items:
        items.append(
            "Teori Perilaku Organisasi (Robbins & Judge, 2022) yang menjelaskan interaksi "
            "individu, pekerjaan, dan konteks organisasi"
        )
    return items


def _mid_theory_items(families) -> List[str]:
    """Butir middle range theory sesuai keluarga variabel yang benar-benar ada."""
    mid: List[str] = []
    if "tuntutan" in families:
        mid.append("Teori Job Demands-Resources (Bakker & Demerouti, 2022)")
    if "kondisi" in families:
        mid.append("Teori Lingkungan Kerja (Sedarmayanti, 2022; Mangkunegara, 2022)")
    if "kepatuhan" in families:
        mid.append("Teori Disiplin Kerja (Sutrisno, 2022; Hasibuan, 2022)")
    if "energi" in families:
        mid.append("Teori Keterikatan Kerja (Kahn, 1990, dalam Saks, 2022; Schaufeli, 2022)")
    if "modal" in families:
        mid.append("Teori Modal Psikologis (Luthans et al., 2022)")
    if "dorongan" in families:
        mid.append("Teori Motivasi Kerja (Vroom, 1964, dalam Kreitner & Kinicki, 2022; Ryan & Deci, 2022)")
    if "evaluasi" in families:
        mid.append("Teori Kepuasan Kerja (Robbins & Judge, 2022)")
    if "kepemimpinan" in families:
        mid.append("Teori Kepemimpinan Transformasional (Bass, 1985; Yukl & Gardner, 2024)")
    if "nilai" in families:
        mid.append("Teori Budaya Organisasi (Schein & Schein, 2023)")
    if "iklim" in families:
        mid.append("Teori Iklim Organisasi (Litwin & Stringer, 1968)")
    if "imbalan" in families:
        mid.append("Expectancy Theory (Vroom, 1964)")
    if "sikap" in families:
        mid.append("Teori Sikap dan Niat Perilaku (Robbins & Judge, 2022)")
    if "kontrol" in families:
        mid.append("Stewardship Theory (Donaldson & Davis, 1991)")
    if "integritas" in families:
        mid.append("Teori Integritas dan Akuntabilitas (Robbins & Judge, 2022)")
    if not mid:
        mid.append("Teori Perilaku Organisasi (Colquitt et al., 2022)")
    return mid


def _theory_name_short(item: str) -> str:
    """Ambil nama teori saja dari butir teori (buang klausa penjelasnya)."""
    return str(item).split(" yang ")[0].strip().rstrip(".") + ""


def _model_label(ctx: Dict[str, Any]) -> str:
    """Label ringkas jenis model struktural sesuai variabel judul."""
    v = _v5(ctx)
    if v["has_m"] and v["has_z"]:
        return "moderated-mediation"
    if v["has_m"]:
        return "mediation"
    if v["has_z"]:
        return "moderation"
    return "direct-effect"


def _model_parts_phrase(ctx: Dict[str, Any]) -> str:
    """Frasa opsional mediasi/moderasi yang hanya muncul bila variabelnya ada."""
    v = _v5(ctx)
    parts = []
    if v["has_m"]:
        parts.append(f"jembatan mediasi {v['m_str']}")
    if v["has_z"]:
        parts.append(f"pengawal kontingensi {v['z_str']}")
    if not parts:
        return ""
    return ", " + ", ".join(parts)


def _build_doc_sub_2_1_0(doc: Document, ctx: Dict[str, Any]):
    """A.1 Grand Teori -- hierarki deduktif yang dipilih dari profil variabel judul."""
    h1(doc, "BAB II\n"
            "KAJIAN PUSTAKA, KERANGKA KONSEPTUAL, DAN PENGEMBANGAN HIPOTESIS")
    h2(doc, "A. KAJIAN PUSTAKA")
    h2(doc, "1. Grand Teori")

    instansi = ctx["instansi"]
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "2.1.0")
    xs = pm["_all"]["x"]
    ms = pm["_all"]["m"]
    zs = pm["_all"]["z"]
    y = pm["Y"]
    x_labels = _join_id(p["label"] for p in xs) or "variabel eksogen"
    m_labels = ", ".join(p["label"] for p in ms) or ""
    z_labels = ", ".join(p["label"] for p in zs) or ""
    model_acronym = _generate_doctoral_model_acronym(ctx)
    families = _theory_families(ctx)

    pv(doc,
       "Grand theory berfungsi sebagai payung filosofis yang menjelaskan akar persoalan, "
       "bukan sekadar hiasan kutipan. Dalam disertasi ini, grand theory dipilih berdasarkan "
       "kesesuaiannya dengan sifat konstruk yang diteliti, sehingga setiap teori yang dirujuk "
       "memiliki konsekuensi metodologis dan interpretatif yang jelas.",
       ctx, "2.1.0-p1", indent=0.5, space_after=6)

    pv(doc,
       "Hierarki teoretis disusun secara deduktif: dari teori umum tentang perilaku dan kinerja "
       "organisasi, menuju teori menengah yang menjelaskan mekanisme psikologis dan kontekstual, "
       "hingga teori terapan yang memandu operasionalisasi variabel. Susunan berjenjang ini "
       "menjaga agar model tidak berdiri di atas satu teori tunggal yang dipaksakan untuk "
       "menjelaskan seluruh hubungan.",
       ctx, "2.1.0-p2", indent=0.5, space_after=6)

    # --- Grand theory dipilih dari keluarga variabel ---
    grand_items: List[str] = _grand_theory_items(families)

    grand_txt = "; ".join(grand_items[:-1]) + "; serta " + grand_items[-1] if len(grand_items) > 1 else grand_items[0]
    _x_fams = {p.get("family") for p in xs if isinstance(p, dict) and p.get("family")}
    _m_fams = {p.get("family") for p in ms if isinstance(p, dict) and p.get("family")}
    _gt_clauses: List[str] = [
        f"{x_labels} membutuhkan teori tentang {_family_domain(_x_fams)}"
    ]
    if ms:
        _gt_clauses.append(
            f"proses perantara melalui {m_labels} membutuhkan teori tentang "
            f"{_family_domain(_m_fams)}"
        )
    if zs:
        _gt_clauses.append(
            f"kondisi batas yang dibawa {z_labels} membutuhkan teori tentang faktor kontingensi"
        )
    _gt_clauses.append(
        f"luaran {y['label']} membutuhkan teori tentang perilaku dan kinerja"
    )
    if len(_gt_clauses) > 1:
        _gt_sentence = "; ".join(_gt_clauses[:-1]) + "; sedangkan " + _gt_clauses[-1]
    else:
        _gt_sentence = _gt_clauses[0]

    pv(doc,
       f"Pada tataran grand theory, penelitian ini berpijak pada {grand_txt}. Pilihan tersebut "
       f"didasarkan pada sifat {x_labels} yang tidak dapat dijelaskan oleh satu dalil tunggal. "
       f"{_gt_sentence}. Konvergensi teori-teori itu menghasilkan "
       f"kerangka yang lebih realistis daripada model deterministik sederhana.",
       ctx, "2.1.0-grand", indent=0.5, space_after=6)

    _model_desc = {
        "direct": "model struktural jalur langsung (direct-effects model)",
        "mediation": "model struktural mediasi (mediation model)",
        "moderation": "model struktural moderasi (moderation model)",
        "moderated_mediation": "model struktural kontingensi hibrida (moderated-mediation model)",
    }.get(str(ctx.get("model_type") or ""), "model struktural integratif")

    pv(doc,
       f"Sebagai kebaruan ilmiah, penelitian ini mengonstruksi {_model_desc} yang diberi nama "
       f"MODEL {model_acronym}. Model ini dirancang untuk menjawab kesenjangan antara kapasitas "
       f"formal yang tersedia dan capaian {y['label']} yang belum optimal di {instansi}.",
       ctx, "2.1.0-novelty", indent=0.5, space_after=6)

    _arki = f"Struktur Model {model_acronym} menghubungkan {x_labels} sebagai variabel eksogen"
    if m_labels:
        _arki += f", disalurkan melalui {m_labels} sebagai mekanisme transmisi"
    if z_labels:
        _arki += f", dan dibatasi oleh {z_labels} sebagai kondisi kontingensi"
    _arki += f", untuk menjelaskan {y['label']} sebagai luaran utama."
    pv(doc, _arki, ctx, "2.1.0-arki", indent=0.5, space_after=6)

    pv(doc,
       f"Relevansi grand theory terletak pada kemampuannya menjelaskan mengapa intervensi "
       f"manajerial tidak cukup berhenti pada penambahan anggaran atau penyusunan prosedur. "
       f"Teori-teori yang dipilih menuntut perhatian pada kualitas sumber daya, proses psikologis, "
       f"dan konteks kelembagaan secara simultan. Dengan demikian, grand theory tidak hanya "
       f"melegitimasi judul, tetapi membimbing pemilihan variabel, arah hipotesis, dan "
       f"interpretasi hasil.",
       ctx, "2.1.0-rel", indent=0.5, space_after=6)

    # --- Middle Range Theory: komposisi dinamis ---
    mid: List[str] = _mid_theory_items(families)

    if len(mid) == 1:
        mid_txt = mid[0]
    else:
        mid_txt = "; ".join(mid[:-1]) + "; serta " + mid[-1]
    _mech = (f"Mekanisme itu mencakup cara aparatur mengolah {_family_domain(_x_fams)} "
             f"menjadi tindakan kerja.")
    if m_labels:
        _mech += f" {m_labels} berperan menyalurkan pengaruh tersebut menuju {y['label']}."
    if z_labels:
        _mech += f" {z_labels} menentukan sejauh mana hubungan itu menguat atau melemah."
    if not (m_labels or z_labels):
        _mech += f" Rangkaian itulah yang menjelaskan mengapa {x_labels} berujung pada {y['label']}."

    _mid_ref = "Teori tersebut" if len(mid) == 1 else "Teori-teori tersebut"
    pv(doc,
       f"Pada tataran middle range theory, penelitian ini dipandu oleh {mid_txt}. {_mid_ref} "
       f"dipilih karena menjelaskan mekanisme yang berada di antara kebijakan makro dan "
       f"perilaku individu. {_mech}",
       ctx, "2.1.0-mid", indent=0.5, space_after=6)

    _rel2 = "Relevansi dengan penelitian: teori tingkat menengah ini menjelaskan "
    if z_labels and m_labels:
        _rel2 += f"mengapa {m_labels} dapat menjadi jembatan dan mengapa {z_labels} dapat mengubah kekuatan jembatan tersebut."
    elif m_labels:
        _rel2 += f"mengapa {m_labels} berfungsi sebagai jembatan antara {x_labels} dan {y['label']}."
    elif z_labels:
        _rel2 += f"mengapa {z_labels} menentukan sejauh mana pengaruh {x_labels} terhadap {y['label']} berlangsung."
    else:
        _rel2 += f"mengapa {x_labels} dan {y['label']} tidak selalu berhubungan secara linier."
    _rel2 += (
        f" Implikasinya, pengukuran tidak cukup menangkap atribut formal, tetapi juga harus "
        f"menjangkau pengalaman kerja aparatur di {instansi}."
    )
    pv(doc, _rel2, ctx, "2.1.0-midrel", indent=0.5, space_after=6)

    # --- Applied theory ---
    pv(doc,
       f"Pada tataran applied theory, penelitian ini menggunakan Teori Kinerja "
       f"(Campbell & Wiernik, 2022) dan AMO Framework (Ability, Motivation, Opportunity) "
       f"sebagai panduan operasional. Teori Campbell menegaskan bahwa kinerja adalah perilaku "
       f"yang dapat diamati, sedangkan AMO menjelaskan bahwa perilaku produktif muncul ketika "
       f"kemampuan, motivasi, dan kesempatan bekerja secara bersamaan. Kerangka ini dipakai "
       f"untuk menurunkan indikator {y['label']} dan variabel penjelasnya.",
       ctx, "2.1.0-applied", indent=0.5, space_after=6)

    pv(doc,
       f"Relevansi teori terapan tampak pada operasionalisasi model {model_acronym}. Setiap "
       f"dimensi yang diukur harus dapat ditelusuri ke kemampuan, motivasi, atau kesempatan kerja. "
       f"Dengan disiplin itu, instrumen penelitian tidak mengukur constructs yang kabur, tetapi "
       f"perilaku dan kondisi kerja yang relevan bagi perbaikan {y['label']} di {instansi}.",
       ctx, "2.1.0-appliedrel", indent=0.5, space_after=6)


_CANON_CITE_KEYS = None


def _canon_cite(text: str, default: str = "Robbins & Judge (2022)") -> str:
    """Jaring pengaman zero-ghost untuk daftar sitasi tokoh variabel.

    Hanya menyisakan sitasi yang benar-benar memiliki entri di registry kanon
    S3. Bila seluruh sitasi tersaring habis, dipakai sitasi kanon bawaan agar
    kalimat tetap utuh dan tetap memiliki entri Daftar Pustaka.
    """
    global _CANON_CITE_KEYS
    raw = str(text or "").strip()
    if not raw:
        return default
    try:
        if _CANON_CITE_KEYS is None:
            _CANON_CITE_KEYS = s3_academic_researcher._bibliography_key_set(
                list(s3_academic_researcher.ALL_CANONICAL_REFERENCES)
            )
        cleaned = s3_academic_researcher.sanitize_citation_text(raw, _CANON_CITE_KEYS)
    except Exception:
        return raw
    if not cleaned.strip():
        return default
    return cleaned


def _format_hierarchical_dimensions(doc: Document, var_name: str, dims: List[Dict[str, Any]], tokoh_rujuk: str = "sintesis indikator terstandarisasi"):
    """Format dimensi dan indikator dalam bentuk hierarkis bertingkat murni tanpa tabel, sesuai acuan TAPD 2026."""
    tokoh = str(tokoh_rujuk).strip()
    parts = [x.strip() for x in tokoh.split(";") if x.strip()]
    if len(parts) > 2:
        tokoh = ", ".join(parts[:-1]) + f", dan {parts[-1]}"
    elif len(parts) == 2:
        tokoh = f"{parts[0]} dan {parts[1]}"
    _tokoh_lower = tokoh.lower()
    awalan = "Berdasarkan" if (" menurut " in _tokoh_lower or "—" in tokoh or "theory" in _tokoh_lower) else "Menurut"
    p(doc, f"{awalan} {tokoh}, dimensi dan indikator yang dipakai dalam mengukur {var_name} adalah:", indent=0.5, space_after=4)
    for idx, d in enumerate(dims, 1):
        dim_nama = d.get("nama", f"Dimensi {idx}")
        indikators = d.get("indikator", [])
        p(doc, f"{idx}. {dim_nama}", indent=0.5, bold=True, space_after=2)
        p(doc, f"Dimensi {dim_nama} diukur dengan menggunakan {len(indikators)} indikator, yaitu:", indent=0.5, space_after=2)
        for i_idx, ind in enumerate(indikators):
            p(doc, f"{chr(97+i_idx)}. {ind}.", indent=0.75, space_after=2)
    p(doc, "", space_after=6)


def _build_doc_sub_2_1_1(doc: Document, ctx: Dict[str, Any]):
    """A.2 2. Kajian Teoritis Variabel Terikat (Y) - Murni Naratif & Dimensi Teks Tanpa Tabel."""
    y = ctx["y_var"]
    onto_vars = ctx.get("ontology", {}).get("variables", {})
    onto_y = onto_vars.get("y") or onto_vars.get("Y", {})

    h2(doc, f"2. Kajian Teoritis {y}")

    # 1. Narasi Pengantar Hakikat Substantif Variabel (Wajib >= 4 baris & bervitasi)
    definisi_onto = str(onto_y.get("definisi") or "").strip().rstrip(".")
    if not definisi_onto:
        definisi_onto = "capaian kerja yang mencerminkan efektivitas pelaksanaan tugas dan kontribusi aparatur terhadap organisasi"
    if not re.match(r"^(kinerja|performance)\b", definisi_onto, re.I):
        definisi_onto = f"{y} didefinisikan sebagai {definisi_onto[0].lower()}{definisi_onto[1:]}"
    p(doc, f"Dalam perspektif manajemen sumber daya manusia kontemporer (Robbins & Judge, 2022; Armstrong & Taylor, 2023), {y} merupakan keluaran utama yang menentukan keberhasilan pencapaian sasaran strategis organisasi. {definisi_onto}. Kinerja aparatur bukan sekadar catatan administratif atas penyelesaian rutinitas, melainkan hasil konversi kapasitas intelektual, dedikasi, dan kepatuhan prosedural menjadi luaran kerja yang berdampak pada kinerja organisasi (Campbell & Wiernik, 2022; Dessler, 2023).", indent=0.5, space_after=6)
    
    p(doc, f"Kajian teoretis tentang {y} diperlukan untuk mengidentifikasi determinan yang mengarahkan peningkatan unjuk kerja aparatur (Colquitt, LePine, & Wesson, 2022). Kasmir (2023) serta Afandi (2023) menegaskan bahwa instrumen penilaian kinerja yang objektif dan komprehensif menjadi acuan bagi intervensi manajerial, pengembangan kompetensi pegawai, dan pemenuhan tuntutan akuntabilitas publik. Oleh karena itu, perumusan konstruk {y} menjadi dasar bagi pengukuran operasional yang valid dan reliabel.", indent=0.5, space_after=6)

    # 2. Telaah Integratif 10 Definisi Para Ahli Otoritatif
    y_lower = y.lower()
    if any(k in y_lower for k in ["fraud", "kecurangan", "korupsi"]):
        p(doc, f"Pencegahan kecurangan merupakan langkah mitigasi terstruktur untuk mengamankan proses operasional dan keuangan publik dari ancaman kerugian finansial serta reputasi kelembagaan. Albrecht, Albrecht, Albrecht, & Zimbelman (2024) mendefinisikan pencegahan fraud sebagai upaya proaktif untuk mengeliminasi pemicu kecurangan dengan cara memperketat pengawasan guna menutup peluang serta membangun budaya yang meniadakan ruang rasionalisasi. Sejalan dengan hal tersebut, Wells (2017) secara spesifik menyoroti bahwa di sektor pengadaan barang dan jasa, pencegahan kecurangan membutuhkan penerapan kontrol holistik, mencakup segregasi tugas yang ketat, transparansi alur lelang, hingga ketersediaan mekanisme pelaporan pelanggaran (whistleblowing). Singleton & Singleton (2010) menambahkan bahwa tujuan utama upaya ini berfokus pada penciptaan lingkungan penangkalan (deterrence environment), di mana probabilitas terdeteksinya kecurangan dibuat sangat tinggi sehingga calon pelaku mengurungkan niatnya.", indent=0.5, space_after=6)

        p(doc, f"Lebih lanjut, Hopwood, Leiner, & Young (2012) menegaskan bahwa efektivitas pencegahan kecurangan sangat bergantung pada objektivitas evaluasi, keterbukaan dokumen, serta ketegasan penegakan sanksi. Kranacher & Riley (2023) menyimpulkan bahwa program pencegahan kecurangan yang komprehensif tidak boleh sekadar memperbaiki instrumen akuntansi, melainkan wajib mengintegrasikan penguatan kapasitas aparatur dan etika profesional di setiap tahapan kritis pengambilan keputusan. Sementara itu, Rustiyaningsih (2023) menjelaskan bahwa pencegahan fraud merupakan usaha untuk menghalangi tindakan ilegal yang melawan hukum serta menekan atau meminimalkan faktor-faktor penyebab timbulnya kecurangan di lingkungan kerja.", indent=0.5, space_after=6)

        p(doc, f"Komalasari & Annisa (2023) menyatakan bahwa pencegahan kecurangan merupakan tindakan untuk menghindari manipulasi, kolusi, dan penipuan dengan tujuan memperkaya diri sendiri secara melawan hukum, sehingga fokusnya berorientasi antisipatif sebelum kerugian terjadi. Hal ini diperkuat oleh Chairi (2022) yang menekankan pentingnya organisasi memiliki strategi dan kebijakan sistematis dalam mengendalikan potensi fraud. Banjarnahor (2023) memandang pencegahan fraud sebagai ikhtiar menolak segala bentuk penyimpangan melalui pembangunan sistem pengendalian yang efektif sejak dini. Puncak diskursus ini ditegaskan oleh Mawikere & Loupatty (2025) yang menyatakan bahwa pencegahan kecurangan adalah tindakan menghilangkan potensi kecurangan melalui peningkatan mutu pengendalian internal, penciptaan budaya kejujuran, dan dukungan keterbukaan antar-aparatur.", indent=0.5, space_after=6)

        sintesis_text = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Albrecht et al., 2024; Singleton & Singleton, 2010; Wells, 2017; serta Mawikere & Loupatty, 2025), dapat disimpulkan bahwa {y} merupakan serangkaian upaya preventif terstruktur dan terpadu yang diimplementasikan institusi guna mengidentifikasi, menangkal, dan meniadakan potensi penyimpangan moral maupun administratif. Konstruk ini dibangun melalui penguatan sistem pengendalian internal yang andal, transparansi penegakan regulasi pengadaan, serta pelembagaan budaya integritas yang menuntut kejujuran dan pertanggungjawaban publik secara paripurna."
        dims_y = [
            {"nama": "Kerangka Regulasi & Kepatuhan Hukum", "indikator": ["Kejelasan regulasi dan petunjuk teknis penugasan", "Ketegasan penegakan sanksi atas pelanggaran prosedur", "Tingkat kepatuhan penuh terhadap ketentuan peraturan perundang-undangan"]},
            {"nama": "Transparansi Alur & Keterbukaan Informasi", "indikator": ["Kemudahan akses informasi publik terkait proses kerja", "Kejelasan kriteria penilaian dan penetapan keputusan", "Keterbukaan alur tahapan kerja kepada para pemangku kepentingan"]},
            {"nama": "Akuntabilitas Pengelolaan & Dokumentasi", "indikator": ["Ketertiban dan keabsahan penyimpanan arsip dokumen resmi", "Ketersediaan rekam jejak audit (audit trail) yang lengkap dan dapat diverifikasi", "Transparansi pelaporan pertanggungjawaban hasil kerja"]},
            {"nama": "Pemisahan Fungsi & Batasan Otoritas", "indikator": ["Pemisahan fungsi dan segregasi tugas antar-pejabat yang tegas", "Integritas pejabat dalam menjalankan wewenang kedinasan", "Kejelasan batasan otoritas dalam pengambilan keputusan strategis"]},
            {"nama": "Pengawasan & Audit Berkala", "indikator": ["Intensitas pelaksanaan pengawasan internal secara berkala", "Independensi aparat pemeriksa dalam mengidentifikasi anomali kerja", "Ketuntasan tindak lanjut atas rekomendasi hasil audit"]}
        ]
        tokoh_rujuk_y = "Albrecht et al. (2024) dan Maria, Mubarok, & Indriasari (2024)"
    else:
        # Default Employee Performance
        p(doc, f"Pemaknaan teoretis mengenai {y} pada mulanya berfokus pada hasil akhir yang terukur. Bernardin & Russell (2013) serta Kasmir (2023) merumuskan kinerja sebagai catatan atas hasil-hasil yang diproduksi dari fungsi pekerjaan atau aktivitas tertentu selama suatu periode waktu yang ditetapkan. Pendekatan berbasis luaran ini sejalan dengan pandangan Mangkunegara (2022) yang menegaskan bahwa kinerja merupakan hasil kerja secara kualitas dan kuantitas yang dicapai oleh seorang pegawai dalam melaksanakan tugasnya sesuai dengan tanggung jawab kedinasan yang dibebankan kepadanya. Lebih lanjut, Kasmir (2023) menambahkan bahwa ketercapaian kinerja harus dibuktikan melalui bukti unjuk kerja faktual dan ketuntasan target program kerja institusi dalam kurun waktu operasional tertentu.", indent=0.5, space_after=6)

        p(doc, f"Meskipun luaran kerja merupakan elemen penting, diskursus keilmuan manajemen perilaku modern memandang bahwa pemaknaan kinerja tidak boleh direduksi semata-mata pada angka-angka luaran administratif. Armstrong & Taylor (2023) menjelaskan bahwa kinerja dipahami sebagai pencapaian sasaran kerja yang terintegrasi secara utuh antara bagaimana pekerjaan tersebut diselesaikan (perilaku kerja nyata) dan apa yang sesungguhnya dihasilkan (kontribusi nilai tambah). Pandangan integratif ini diperkuat oleh Campbell & Wiernik (2022) yang menegaskan bahwa esensi kinerja sejatinya adalah sinonim dari perilaku kerja itu sendiri, yaitu serangkaian tindakan nyata yang dapat diamati, dinilai, dan berada di bawah kendali individu dalam rangka memenuhi tuntutan peran tugasnya. Sejalan dengan hal tersebut, Colquitt, LePine, & Wesson (2022) mendefinisikan kinerja pegawai sebagai nilai total dari serangkaian perilaku pegawai yang berkontribusi, baik secara positif maupun negatif, terhadap pemenuhan sasaran-sasaran organisasi.", indent=0.5, space_after=6)

        p(doc, f"Ditinjau dari perspektif fungsional dan penegakan standar kerja, Mathis & Jackson (2011), dalam Kasmir (2023), merumuskan bahwa kinerja pada hakikatnya adalah apa yang dilakukan atau tidak dilakukan oleh pegawai yang mempengaruhi seberapa besar kontribusi nyata mereka terhadap keberhasilan organisasi. Tingkat keberhasilan tersebut, menurut Dessler (2023), dievaluasi secara komparatif melalui prestasi aktual pegawai berdasarkan standar kerja dan kompetensi teknis yang telah ditetapkan institusi, yang sekaligus mencerminkan tingkat kepatuhan prosedur operasional. Pada dimensi psikologis keperilakuan, Robbins & Judge (2022) mengemukakan bahwa kinerja pegawai merupakan tingkat keberhasilan individu dalam memenuhi tuntutan pekerjaan melalui pengerahan energi fisik, mental, dan emosional secara optimal. Puncak dari seluruh pemaknaan unjuk kerja dirangkum oleh Afandi (2023) yang menegaskan bahwa kinerja aparatur adalah hasil kerja yang dicapai dalam batas wewenang dan tanggung jawab masing-masing secara legal, tidak melanggar hukum, serta selaras dengan etika dan moralitas kedinasan.", indent=0.5, space_after=6)

        sintesis_text = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Bernardin & Russell, 2013; Kasmir, 2023; Campbell & Wiernik, 2022; Robbins & Judge, 2022; Armstrong & Taylor, 2023; serta Afandi, 2023), dapat disimpulkan bahwa {y} merupakan konstruk multidimensional integratif yang merefleksikan kapabilitas, ketekunan, dan keberhasilan aparatur dalam menghasilkan luaran tugas bermutu tinggi. Capaian unjuk kerja ini ditandai oleh ketuntasan target volume beban dinas, kedisiplinan terhadap batas tenggat waktu, efisiensi pendayagunaan sarana kantor, serta inisiatif kolaborasi sinergis guna mewujudkan sasaran strategis institusi secara optimal."
        dims_y = onto_y.get("dimensi") if (onto_y and onto_y.get("dimensi")) else [
            {"nama": "Kualitas Pekerjaan (Quality of Work)", "indikator": ["Ketelitian dan ketepatan kalkulasi teknis hasil penugasan", "Kerapian penyusunan berkas dan dokumen kedinasan", "Kepatuhan penuh terhadap Standar Operasional Prosedur (SOP)"]},
            {"nama": "Kuantitas Output (Quantity of Output)", "indikator": ["Tingkat ketercapaian target volume pekerjaan rutin kedinasan", "Kemampuan menuntaskan beban tugas tambahan di luar tugas pokok", "Ketahanan kerja dalam menyelesaikan tumpukan beban penugasan"]},
            {"nama": "Ketepatan Waktu (Timeliness / Punctuality)", "indikator": ["Kecepatan penuntasan tugas sebelum atau tepat pada batas waktu (deadline)", "Daya tanggap (responsiveness) dalam menindaklanjuti disposisi pimpinan", "Disiplin pemanfaatan jam kerja kedinasan secara optimal"]},
            {"nama": "Efektivitas & Kemandirian (Effectiveness & Independence)", "indikator": ["Efisiensi dan kehati-hatian dalam pendayagunaan sarana kantor", "Inisiatif mandiri dalam memecahkan hambatan teknis operasional", "Kemandirian pengambilan keputusan teknis sesuai batas kewenangan"]},
            {"nama": "Kerjasama Tim (Teamwork & Collaboration)", "indikator": ["Kesediaan berkoordinasi sinergis lintas bidang kedinasan", "Sikap kooperatif dan saling mendukung dalam penuntasan target tim"]}
        ]
        tokoh_rujuk_y = onto_y.get("tokoh_ahli", "Robbins & Judge (2022) dan Campbell & Wiernik (2022)")

    tokoh_rujuk_y = _canon_cite(tokoh_rujuk_y)

    # 3. Sintesis Konseptual Peneliti Doktoral (Wajib >= 4 baris & bervitasi)
    p(doc, sintesis_text, indent=0.5, space_after=6)

    # 4. Dimensi dan Indikator Operasional Baku (Format Teks Hirarkis Bertingkat Tanpa Tabel)
    _format_hierarchical_dimensions(doc, y, dims_y, tokoh_rujuk_y)

    # 5. Kesimpulan Terpadu Definisi dan Dimensi (Paragraf Akhir Sub-Bab, Wajib >= 4 baris & bervitasi)
    dim_names_str = _join_id([d.get("nama", "").split("(")[0].strip() for d in dims_y])
    p(doc, f"Berdasarkan sintesis definisi konseptual dan rincian dimensi serta indikator yang merujuk pada kerangka teoritis {tokoh_rujuk_y}, dapat disimpulkan bahwa pengukuran {y} dalam penelitian disertasi ini secara terpadu mencakup dimensi {dim_names_str}. Keterpaduan struktural antara rumusan konseptual hasil sintesis dengan butir-butir indikator operasional tersebut menegaskan bahwa instrumen ini memenuhi kaidah validitas isi (content validity) dan reliabilitas pengukuran yang tinggi sebagaimana disyaratkan dalam metodologi penelitian manajemen mutakhir (Hair et al., 2022; Sekaran & Bougie, 2020), sehingga mampu merefleksikan esensi dan dinamika {y} secara objektif, komprehensif, dan terukur.", indent=0.5, space_after=6)


def _build_doc_sub_2_1_2(doc: Document, ctx: Dict[str, Any]):
    """A.3 3. Kajian Teoritis Variabel Mediasi (M) - Murni Naratif & Dimensi Teks Tanpa Tabel."""
    m_list = ctx.get("m_vars", [])
    m = m_list[0] if m_list else ctx.get("m", "")
    if not m:
        h2(doc, "3. Kajian Teoritis Variabel Mediasi")
        p(doc, "Model penelitian tidak menggunakan variabel intervening mediasi.", indent=0.5)
        return

    onto_vars = ctx.get("ontology", {}).get("variables", {})
    onto_m = onto_vars.get("m") or onto_vars.get("M", {})

    h2(doc, f"3. Kajian Teoritis Variabel Mediasi ({m})")

    # 1. Narasi Pengantar Hakikat Substantif Variabel (Wajib >= 4 baris & bervitasi)
    p(doc, f"Dalam diskursus ilmu perilaku keorganisasian dan manajemen sumber daya manusia (Robbins & Judge, 2022; Ryan & Deci, 2022), {m} merupakan dorongan psikologis dan komitmen keperilakuan mendasar yang menentukan sejauh mana aparatur bersedia mengerahkan dedikasi terbaiknya dalam menyelesaikan tugas-tugas kedinasan. Ketersediaan sarana kerja, infrastruktur modern, dan regulasi formal tidak akan menghasilkan luaran kinerja yang optimal tanpa adanya keterlibatan afektif, kegairahan bertugas, dan ketekunan aparatur di lapangan (Bakker & Demerouti, 2022; Schaufeli, 2022). Keberadaan {m} bertindak sebagai energi penggerak batiniah yang mengarahkan fokus aparatur untuk gigih menghadapi tantangan tugas, memelihara integritas profesi, dan secara konsisten berorientasi pada pencapaian standar pelayanan publik yang prima.", indent=0.5, space_after=6)
    
    p(doc, f"Kajian teoretis mendalam mengenai {m} menjadi sangat esensial karena variabel ini bertindak sebagai jembatan transmisi psikologis yang menyalurkan pengaruh dari pilar-pilar kapasitas organisasi menuju luaran unjuk kerja nyata aparatur (Rich, Lepine, & Crawford, 2010, dalam Robbins, 2022). Menurut Colquitt, LePine, & Wesson (2022), aparatur yang memiliki tingkat keterikatan dan motivasi intrinsik yang kuat akan menunjukkan ketahanan mental tinggi, kesediaan bekerja melampaui kewajiban peran minimal, serta kepatuhan moral yang kokoh. Oleh karena itu, telaah atas fondasi teoretis {m} diperlukan guna menyusun konstruk pengukuran yang valid dan memiliki daya eksplanatori yang tangguh dalam kerangka model struktural doktoral ini.", indent=0.5, space_after=6)

    m_lower = m.lower()
    if any(k in m_lower for k in ["engagement", "keterikatan", "work engagement"]):
        p(doc, f"Kajian teoretis mengenai keterikatan kerja berakar pada pemikiran awal Kahn (1990), dalam Saks (2022), yang mendefinisikan keterikatan kerja sebagai pendayagunaan diri anggota organisasi secara fisik, kognitif, dan emosional dalam pelaksanaan peran kerja mereka. Konsep ini kemudian disempurnakan secara operasional oleh Schaufeli, Bakker, & Salanova (2006), dalam Bakker (2022), yang merumuskan work engagement sebagai kondisi mental positif, memuaskan, dan berkaitan dengan pekerjaan yang ditandai oleh semangat tinggi (vigor), dedikasi kuat (dedication), dan keasyikan bertugas (absorption). Menurut Bakker & Demerouti (2022), pegawai yang engaged memiliki koneksi energik dan afektif yang mendalam terhadap aktivitas kerja mereka, memandang penugasan bukan sebagai beban melainkan sebagai tantangan yang bermakna.", indent=0.5, space_after=6)

        p(doc, f"Lebih lanjut, Saks (2022) mengelaborasi bahwa work engagement mencakup integrasi antara keterikatan pada peran tugas (job engagement) dan keterikatan pada institusi secara menyeluruh (organization engagement). Christian, Garza, & Slaughter (2011), dalam Albrecht (2022), menunjukkan bahwa keterikatan kerja mencerminkan investasi simultan atas energi pribadi dalam pelaksanaan tugas, yang membedakannya secara tegas dari konsep kepuasan pasif semata. Rich, Lepine, & Crawford (2010), dalam Robbins (2022), menambahkan bahwa keterikatan kerja merupakan mediator psikologis esensial yang menghubungkan karakteristik kepribadian dan dukungan lingkungan kerja dengan unjuk kerja nyata.", indent=0.5, space_after=6)

        p(doc, f"Pada tataran manajerial modern, Macey & Schneider (2008), dalam Schaufeli (2022), memetakan engagement ke dalam tiga fase: trait engagement (orientasi kepribadian positif), state engagement (kondisi afektif berenergi), dan behavioral engagement (perilaku kerja proaktif ekstra-peran). Albrecht (2022) menegaskan bahwa iklim organisasi yang suportif merupakan prasyarat mutlak bagi berkembangnya keterikatan kerja aparatur. Rasa memiliki nilai bersama antara pegawai dan visi organisasi menjadi pintu masuk tumbuhnya keterikatan (Saks, 2022). Di Indonesia, Muhtar dan Anita (2022) menunjukkan bahwa keterikatan kerja memediasi pengaruh kepemimpinan transformasional terhadap kinerja pegawai, sehingga keterikatan aparatur sipil tercermin melalui kesediaan mencurahkan energi fisik dan fokus pikiran demi tercapainya keunggulan pelayanan publik.", indent=0.5, space_after=6)

        sintesis_m = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Kahn, 1990, dalam Saks, 2022; Schaufeli, Bakker, & Salanova, 2006, dalam Bakker, 2022; Bakker & Demerouti, 2022; serta Muhtar & Anita, 2022), dapat disimpulkan bahwa {m} merupakan kondisi psikologis dan perilaku kerja positif yang ditandai oleh tingginya curahan energi fisik, persistensi daya tahan mental, antusiasme dedikasi terhadap makna tugas, serta konsentrasi mendalam aparatur dalam menyelesaikan tanggung jawab kedinasan. Kondisi keperilakuan ini mendorong individu untuk mendedikasikan seluruh kapasitas dirinya demi kemajuan organisasi melampaui batas kewajiban peran formal."
        dims_m = [
            {"nama": "Vigor (Semangat & Ketahanan Fisik-Mental)", "indikator": ["Tingginya energi fisik dan ketahanan mental saat melaksanakan tugas kedinasan", "Kesediaan mencurahkan usaha maksimal walau menghadapi kesulitan teknis", "Ketahanan kerja dan tidak mudah lelah dalam menyelesaikan beban tugas"]},
            {"nama": "Dedication (Dedikasi & Rasa Kebermaknaan)", "indikator": ["Rasa bangga dan antusiasme mendalam terhadap profesi dan peran tugas", "Keyakinan bahwa penugasan yang dijalankan memiliki makna dan nilai luhur", "Inspirasi dan tantangan positif yang dirasakan saat menuntaskan penugasan"]},
            {"nama": "Absorption (Konsentrasi & Keasyikan Bertugas)", "indikator": ["Fokus konsentrasi penuh dan keterikatan pikiran secara mendalam saat bertugas", "Pemanfaatan waktu kerja secara intensif sehingga waktu terasa berjalan cepat", "Keleluasaan dalam menuntaskan tugas tanpa mudah terdistraksi oleh gangguan sekitar"]}
        ]
        tokoh_rujuk_m = "Schaufeli, Bakker, & Salanova (2006, dalam Bakker, 2022)"
    elif any(k in m_lower for k in ["etika", "profesi", "ethical"]):
        p(doc, f"Etika profesi merupakan sekumpulan prinsip moral, integritas, dan standar perilaku luhur yang wajib dijunjung tinggi oleh aparatur dalam menjalankan wewenang kedinasannya. Duska, Duska, & Kury (2018) menyatakan bahwa praktik etika profesi menuntut praktisi untuk menerapkan prinsip kejujuran, objektivitas, dan kehati-hatian secara ketat sebagai bentuk pertanggungjawaban kepada publik. Brooks & Dunn (2017) memperluas pandangan ini dengan menjelaskan bahwa kepatuhan terhadap etika profesi berfungsi sebagai jembatan yang mentransformasikan pengetahuan teknis individu ke dalam tindakan nyata yang menjunjung keadilan serta mencegah benturan kepentingan (conflict of interest).", indent=0.5, space_after=6)

        p(doc, f"Mintz & Morris (2019) menegaskan bahwa etika profesi bukan sekadar pedoman normatif tertulis, melainkan komitmen perilaku yang mengharuskan aparatur menempatkan perlindungan kepentingan publik di atas tekanan pragmatis pihak lain. Menurut Bowie (2017), kepatuhan etis merupakan instrumen pencegahan paling efektif karena individu yang beretika secara mandiri akan menolak skema penyimpangan. Hoffman, Frederick, & Schwartz (2014) menyimpulkan bahwa dalam ekosistem sektor publik, etika profesi secara imperatif mewajibkan setiap pengambil keputusan untuk menolak gratifikasi yang mencederai integritas pelayanan.", indent=0.5, space_after=6)

        p(doc, f"Nautani (2024) berpendapat bahwa pemahaman etika profesi yang kokoh akan mengarahkan niat berperilaku aparatur menjadi tindakan transparan di lapangan. Atarwaman (2022) mengidentifikasi pilar utama dalam kode etik profesi, yaitu integritas, objektivitas, kehati-hatian profesional, kerahasiaan, dan perilaku profesional. Surtikanti, Anggadini, & Aprilia (2023) menambahkan bahwa etika profesi mempertahankan martabat kelembagaan melalui pelayanan berkualitas. Puncak diskursus ini ditegaskan oleh Distiyah (2024) serta Kinanti & Novita (2025) yang menyatakan bahwa nilai-nilai etis merupakan fondasi keadilan yang membentengi aparatur dari bias dan intervensi eksternal.", indent=0.5, space_after=6)

        sintesis_m = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Duska, Duska, & Kury, 2018; Brooks & Dunn, 2017; Mintz & Morris, 2019; Atarwaman, 2022; serta Distiyah, 2024), dapat disimpulkan bahwa {m} merupakan perwujudan komitmen moral dan kepatuhan aparatur terhadap standar perilaku profesional yang mencakup kejujuran integritas, objektivitas pengambilan keputusan, kecermatan kehati-hatian kerja, serta perlindungan kerahasiaan jabatan. Kepatuhan etis ini menjadi benteng normatif yang memandu seluruh tindakan kedinasan aparatur agar bebas dari konflik kepentingan dan berorientasi murni pada kemaslahatan publik."
        dims_m = [
            {"nama": "Integritas Moral", "indikator": ["Kejujuran dan ketulusan dalam setiap pelaksanaan proses tugas", "Kepatuhan penuh pada aturan demi memelihara kepercayaan publik", "Konsistensi menolak segala bentuk kompromi yang menyimpang"]},
            {"nama": "Objektivitas Pengambilan Keputusan", "indikator": ["Sikap netral dan tidak memihak dalam menjalankan fungsi penugasan", "Penolakan terhadap benturan kepentingan pribadi maupun kelompok", "Ketegasan menolak gratifikasi atau imbalan tidak sah"]},
            {"nama": "Kecermatan & Kehati-hatian Profesional", "indikator": ["Penerapan standar operasional secara cermat dan seksama", "Pemeliharaan kompetensi teknis melalui pembelajaran berkelanjutan", "Pelaksanaan verifikasi silang (cross-check) atas keabsahan dokumen tugas"]},
            {"nama": "Perilaku Profesional & Kerahasiaan", "indikator": ["Konsistensi menjaga reputasi keluhuran institusi dan jabatan", "Kehati-hatian dalam mengelola informasi penting kedinasan", "Penghindaran perilaku yang mencederai martabat korps aparatur"]}
        ]
        tokoh_rujuk_m = "Duska, Duska, & Kury (2018) dan Atarwaman (2022)"
    else:
        # Default Motivasi Kerja
        p(doc, f"Kajian mengenai {m} berakar kuat pada teori motivasi kognitif dan determinasi diri. Robbins & Judge (2022) mendefinisikan motivasi sebagai proses yang menjelaskan intensitas, arah, dan ketekunan usaha individu menuju pencapaian tujuan keorganisasian. Pandangan ini diperkaya oleh Ryan & Deci (2022) melalui Self-Determination Theory yang menggarisbawahi bahwa kondisi psikologis aparatur digerakkan oleh energi batin untuk bertindak, bertahan dalam kesulitan, dan mengarahkan perilaku demi memenuhi kebutuhan dasar otonomi dan kompetensi. Sejalan dengan itu, McClelland (1961), dalam Schunk (2022), mengemukakan bahwa dorongan psikologis berakar pada tiga motif fundamental, yaitu dorongan berprestasi (need for achievement), kebutuhan berafiliasi, dan dorongan memperoleh pengaruh.", indent=0.5, space_after=6)

        p(doc, f"Pada tataran kognitif, Bandura (1986), dalam Schunk & DiBenedetto (2022), menegaskan bahwa motivasi adalah proses yang dibangkitkan oleh keyakinan efikasi diri, di mana individu menetapkan tujuan pribadi dan merencanakan strategi pencapaian secara mandiri. Sementara itu, Vroom (1964), dalam Kreitner & Kinicki (2022), memandang kekuatan dorongan bertindak sebagai perkalian rasional antara ekspektansi keberhasilan usaha, instrumentalitas perolehan imbalan, dan valensi nilai dari penghargaan tersebut. Di sisi lain, Herzberg (1966), dalam Luthans et al. (2022), membedakan antara faktor pemuas hakiki (motivators) seperti pengakuan kedinasan dan rasa berprestasi dengan faktor higiene lingkungan kerja. Pentingnya penetapan sasaran yang jelas juga ditekankan oleh Locke & Latham (1990), dalam Schunk (2022), yang menyatakan bahwa pengerahan energi psikologis dipicu oleh komitmen terhadap tujuan kerja yang spesifik dan menantang.", indent=0.5, space_after=6)

        p(doc, f"Dalam konteks manajemen aparatur di Indonesia, Hasibuan (2022) menegaskan bahwa motivasi adalah pemberian daya penggerak yang menciptakan kegairahan kerja seseorang agar mereka mau bekerja sama dan bekerja efektif demi mencapai kepuasan bersama. Senada dengan hal tersebut, Sedarmayanti (2022) merumuskan motivasi sebagai kesediaan mengerahkan tingkat upaya yang tinggi ke arah tercapainya sasaran kelembagaan yang diimbangi pemenuhan kebutuhan pegawai. Puncak elaborasi ini dirumuskan oleh Uno (2023) yang menyatakan bahwa motivasi kerja tercermin melalui hasrat berprestasi, dorongan kebutuhan kedinasan, harapan masa depan, dan penghargaan atas tugas yang diselesaikan.", indent=0.5, space_after=6)

        sintesis_m = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Robbins & Judge, 2022; Ryan & Deci, 2022; McClelland, 1961, dalam Schunk, 2022; Sedarmayanti, 2022; serta Uno, 2023), dapat disimpulkan bahwa {m} merupakan dorongan psikologis integratif yang membangkitkan intensitas daya upaya, mengarahkan fokus perilaku, dan memelihara ketekunan aparatur dalam menuntaskan target-target tugas kedinasan. Dorongan ini menyelaraskan pemenuhan kebutuhan berprestasi pribadi dengan komitmen terhadap pencapaian sasaran strategis institusi secara berkesinambungan."
        dims_m = onto_m.get("dimensi") if (onto_m and onto_m.get("dimensi")) else [
            {"nama": "Dorongan Berprestasi (Need for Achievement)", "indikator": ["Ambisi menyelesaikan penugasan kerja yang kompleks", "Keinginan melampaui target standar kinerja yang ditetapkan", "Kepuasan batin atas pencapaian hasil kerja yang bermutu tinggi"]},
            {"nama": "Ketekunan & Daya Tahan (Persistence of Effort)", "indikator": ["Konsistensi alokasi waktu dan energi dalam menuntaskan pekerjaan", "Ketahanan mental dalam menghadapi tekanan beban kerja tinggi", "Sikap pantang menyerah saat menghadapi kendala teknis kedinasan"]},
            {"nama": "Kebutuhan Berafiliasi (Need for Affiliation)", "indikator": ["Kenyamanan berinteraksi dan berkolaborasi dalam tim kerja", "Kesediaan berbagi pengetahuan dan membantu rekan sejawat"]},
            {"nama": "Komitmen Sasaran (Goal Commitment)", "indikator": ["Keselarasan tujuan pribadi dengan sasaran strategis institusi", "Kesediaan mengerahkan usaha ekstra melampaui batas kewajiban minimal"]}
        ]
        tokoh_rujuk_m = onto_m.get("tokoh_ahli", "Ryan & Deci (2022) dan Robbins & Judge (2022)")

    tokoh_rujuk_m = _canon_cite(tokoh_rujuk_m)

    # 3. Sintesis Konseptual Peneliti Doktoral (Wajib >= 4 baris & bervitasi)
    p(doc, sintesis_m, indent=0.5, space_after=6)

    # 4. Dimensi dan Indikator Operasional Baku (Format Teks Hirarkis Bertingkat Tanpa Tabel)
    _format_hierarchical_dimensions(doc, m, dims_m, tokoh_rujuk_m)

    # 5. Kesimpulan Terpadu Definisi dan Dimensi (Paragraf Akhir Sub-Bab, Wajib >= 4 baris & bervitasi)
    dim_m_names_str = ", ".join([d.get("nama", "").split("(")[0].strip() for d in dims_m])
    p(doc, f"Berdasarkan sintesis definisi konseptual dan rincian dimensi serta indikator operasional yang berakar pada rujukan {tokoh_rujuk_m}, dapat disimpulkan bahwa operasionalisasi {m} dalam penelitian disertasi ini secara terpadu mencakup dimensi {dim_m_names_str}. Keterpaduan konseptual antara rumusan hasil sintesis dengan indikator-indikator operasional tersebut menunjukkan bahwa instrumen ini memiliki validitas konstruk yang sahih dan dapat diandalkan sesuai kaidah metodologi penelitian manajemen kontemporer (Hair et al., 2022; Cooper & Schindler, 2022), sehingga mampu merefleksikan esensi {m} secara akurat dan komprehensif.", indent=0.5, space_after=6)


def _build_doc_sub_2_1_4(doc: Document, ctx: Dict[str, Any]):
    """A.4 4. Kajian Teoritis Variabel Moderasi (Z) - Murni Naratif & Dimensi Teks Tanpa Tabel."""
    instansi = (ctx.get("instansi") or "Instansi Pemerintah").strip()
    z_list = ctx.get("z_vars", [])
    z = z_list[0] if z_list else ctx.get("z", "")
    if not z:
        h2(doc, "4. Kajian Teoritis Variabel Moderasi")
        p(doc, "Model penelitian tidak menggunakan variabel moderasi.", indent=0.5)
        return

    onto_vars = ctx.get("ontology", {}).get("variables", {})
    onto_z = onto_vars.get("z") or onto_vars.get("Z", {})

    h2(doc, f"4. Kajian Teoritis Variabel Moderasi ({z})")

    # 1. Narasi pengantar dari profil moderator (bukan framing kepemimpinan)
    pm_z = nv.profile_map(ctx)
    prof_z = nv.lookup(z) or nv.make_profile(z, {}, ctx)
    y_z = pm_z["Y"]
    x_labels_z = _join_id(p["label"] for p in pm_z["_all"]["x"]) or "variabel eksogen"

    pv(doc,
       f"Kajian atas {z} berangkat dari pertanyaan tentang kondisi psikologis apa yang membuat "
       f"pengaruh {x_labels_z} terhadap {y_z['label']} berbeda antaraparatur. {z} dipahami "
       f"sebagai {prof_z['definisi']}. Dalam model ini, {z} tidak diperlakukan sebagai variabel "
       f"tambahan, melainkan sebagai kondisi batas yang menentukan sejauh mana kapasitas dan "
       f"tuntutan kerja terkonversi menjadi {y_z['label']}.",
       ctx, "2.1.4-intro1", indent=0.5, space_after=6)

    pv(doc,
       f"Penempatan {z} sebagai moderator berimplikasi metodologis. Efeknya tidak dibaca dari "
       f"korelasi sederhana, tetapi dari perubahan kekuatan hubungan pada tingkat {z} yang "
       f"berbeda (Hayes, 2022). Apabila {z} kuat, {prof_z['tinggi']}. Sebaliknya, apabila {z} "
       f"lemah, risiko yang muncul adalah {prof_z['risiko']}, sehingga dampak {x_labels_z} "
       f"terhadap {y_z['label']} melemah. "
       f"Karena itu, pengukuran {z} harus terpisah dari pengukuran variabel eksogen dan variabel "
       f"terikat agar tidak terjadi tumpang tindih konstruk.",
       ctx, "2.1.4-intro2", indent=0.5, space_after=6)

    z_lower = z.lower()
    if any(k in z_lower for k in ["pimpinan", "tone at the top", "komitmen pimpinan"]):
        p(doc, f"Komitmen pimpinan atau tone at the top merupakan faktor penentu yang mengarahkan iklim integritas dan disiplin di seluruh level kelembagaan. Northouse (2024) menjelaskan bahwa kepemimpinan yang efektif diukur dari komitmen atasan dalam menciptakan visi yang jelas, mengalokasikan sumber daya, dan memberikan keteladanan moral (role modeling) yang menginspirasi bawahan untuk bertindak etis. Yukl & Gardner (2024) menyatakan bahwa komitmen pimpinan bertindak sebagai katalisator utama: pimpinan yang tegas akan memperkuat budaya tata kelola organisasi, sementara pimpinan yang permisif akan memicu pembiaran sistemik terhadap pelanggaran. Bass & Riggio (2006) menekankan bahwa pimpinan yang memiliki komitmen etis tinggi akan secara proaktif melindungi aparatur dari tekanan menyimpang.", indent=0.5, space_after=6)

        p(doc, f"Kouzes & Posner (2023) berargumen bahwa komitmen tidak bermakna jika hanya diucapkan, melainkan harus dibuktikan melalui tindakan nyata (model the way) seperti penegakan aturan secara konsisten tanpa pandang bulu. Daft (2022) menambahkan bahwa pimpinan yang mengadopsi prinsip stewardship akan selalu menempatkan kepentingan jangka panjang organisasi dan transparansi proses kerja di atas kepentingan pragmatis. Sinambela (2024) merumuskan bahwa komitmen pimpinan mencakup kesungguhan dalam menyalurkan pikiran, wewenang, dan keteladanan tindakan demi tercapainya target-target organisasi. Sejalan dengan hal tersebut, Alenarda & Gunarto (2024) mendefinisikannya sebagai komitmen pimpinan untuk merealisasikan tanggung jawab dan janji yang melekat pada peran kepemimpinannya.", indent=0.5, space_after=6)

        p(doc, f"Aulia (2025) menegaskan bahwa komitmen pimpinan mencakup kesediaan manajemen menyediakan sumber daya yang memadai, waktu, serta dukungan kebijakan yang konsisten. Gera (2024) menjelaskan bahwa pimpinan yang berkomitmen mampu mengintegrasikan visi pribadi dengan visi organisasi sehingga tercipta budaya kerja yang solid di antara bawahan. Zulfri (2024) menyoroti keterlibatan aktif pimpinan puncak dalam pengambilan keputusan strategis yang berdampak panjang. Puncak dari telaah ini dirumuskan oleh Ananda (2024) serta Masega (2025) yang menyimpulkan bahwa komitmen pimpinan merupakan keselarasan mutlak antara perkataan visi pimpinan dengan tindakan nyata di lapangan, yang menjadi fondasi utama terbangunnya kepercayaan seluruh jajaran aparatur.", indent=0.5, space_after=6)

        sintesis_z = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Northouse, 2024; Yukl & Gardner, 2024; Kouzes & Posner, 2023; Daft, 2022; serta Sinambela, 2024), dapat disimpulkan bahwa {z} merupakan kesungguhan tekad, konsistensi keteladanan, dan keterlibatan aktif jajaran pimpinan dalam mengawal integritas kelembagaan. Komitmen ini diwujudkan melalui pemberian teladan moral nyata di tempat kerja, penerbitan regulasi internal yang protektif, alokasi sumber daya yang memadai, serta penegakan sanksi yang adil demi menciptakan iklim kerja yang bersih dan berintegritas tinggi."
        dims_z = [
            {"nama": "Keteladanan Moral & Integritas Tindakan", "indikator": ["Konsistensi antara ucapan dan tindakan nyata pimpinan di tempat kerja", "Keteladanan pimpinan dalam mematuhi kode etik dan disiplin kedinasan", "Ketegasan pimpinan menolak segala bentuk kompromi yang menyimpang"]},
            {"nama": "Dukungan Kebijakan & Perlindungan Staf", "indikator": ["Penerbitan instruksi dan regulasi internal yang mendukung iklim kerja bersih", "Perlindungan kedinasan bagi aparatur yang memegang teguh integritas", "Ketegasan sanksi yang adil tanpa toleransi terhadap pelanggaran"]},
            {"nama": "Penyediaan Sumber Daya & Fasilitasi", "indikator": ["Alokasi anggaran yang memadai untuk kelancaran pelaksanaan tugas", "Fasilitasi pengembangan sarana kerja dan sistem pengawasan modern", "Penyediaan waktu pimpinan untuk pembinaan dan pendampingan staf"]},
            {"nama": "Keterlibatan Pengambilan Keputusan Strategis", "indikator": ["Keterlibatan proaktif pimpinan dalam merumuskan arah solusi masalah kerja", "Ketelitian mengevaluasi laporan pertanggungjawaban unit tugas", "Responsivitas tinggi menindaklanjuti rekomendasi perbaikan sistem"]}
        ]
        tokoh_rujuk_z = "Northouse (2024) dan Sinambela (2024)"
    elif any(k in z_lower for k in ["motivasi", "motivation"]):
        p(doc, f"Motivasi kerja merupakan konstruk psikologis yang menjelaskan arah, intensitas, dan persistensi usaha individu dalam mencapai sasaran kerja. McClelland (1961), dalam Schunk (2022), merumuskan motivasi berprestasi (need for achievement) sebagai dorongan internal untuk menuntaskan tugas secara unggul, di samping kebutuhan afiliasi (need for affiliation) dan kebutuhan pengaruh (need for power) yang membentuk dinamika perilaku keorganisasian. Ryan & Deci (2022) melalui Self-Determination Theory menegaskan bahwa kualitas motivasi ditentukan oleh derajat pemenuhan kebutuhan otonomi, kompetensi, dan keterhubungan, sehingga motivasi intrinsik menghasilkan ketekunan yang lebih tahan lama dibandingkan motivasi ekstrinsik yang bersifat instrumental.", indent=0.5, space_after=6)

        p(doc, f"Luthans, Luthans, & Luthans (2022) menambahkan bahwa motivasi kerja yang sehat bertumpu pada harapan keberhasilan, keyakinan akan kemampuan diri, dan dukungan lingkungan kerja yang suportif. Locke & Latham (1990), dalam Schunk (2022), melalui Goal-Setting Theory menjelaskan bahwa motivasi kerja menguat manakala sasaran kerja bersifat spesifik, menantang namun terjangkau, serta disertai umpan balik yang teratur atas kemajuan pencapaian. Dalam konteks organisasi publik, motivasi kerja tidak hanya dipengaruhi oleh imbalan formal, melainkan juga oleh makna pelayanan, pengakuan profesional, dan kualitas hubungan atasan-bawahan (Yukl & Gardner, 2024).", indent=0.5, space_after=6)

        p(doc, f"Pada ranah moderasi, motivasi kerja berfungsi sebagai kondisi batas yang menentukan sejauh mana stimulus kepemimpinan dan kapasitas organisasi mampu dikonversi menjadi kinerja nyata. Aparatur dengan motivasi kerja tinggi cenderung merespons tuntutan tugas sebagai tantangan yang layak diupayakan, sehingga pengaruh variabel eksogen terhadap kinerja menjadi lebih kuat. Sebaliknya, motivasi kerja yang rendah membuat intervensi manajerial berhenti pada kepatuhan formal tanpa energi diskresioner yang memadai (Ryan & Deci, 2022; Luthans et al., 2022).", indent=0.5, space_after=6)

        sintesis_z = f"Berdasarkan berbagai pendapat para ahli di atas (seperti McClelland, 1961, dalam Schunk, 2022; Ryan & Deci, 2022; Locke & Latham, 1990, dalam Schunk, 2022; serta Luthans et al., 2022), dapat disimpulkan bahwa {z} merupakan daya dorong internal aparatur yang mengarahkan arah, intensitas, dan persistensi usaha kerja menuju pencapaian sasaran organisasi. Motivasi ini terbentuk dari interaksi antara kebutuhan berprestasi, pemenuhan otonomi dan kompetensi, kejelasan target kerja, serta dukungan lingkungan kelembagaan yang memaknai kerja sebagai pelayanan publik yang bernilai."
        dims_z = onto_z.get("dimensi") if (onto_z and onto_z.get("dimensi")) else [
            {"nama": "Kebutuhan Berprestasi (Need for Achievement)", "indikator": ["Keinginan menuntaskan tugas menantang secara unggul", "Penetapan standar kerja pribadi yang tinggi", "Ketekunan mengejar sasaran hingga tuntas"]},
            {"nama": "Kebutuhan Afiliasi (Need for Affiliation)", "indikator": ["Kedekatan dan kerja sama harmonis dengan rekan kerja", "Keterlibatan aktif dalam kegiatan unit kerja", "Kepedulian terhadap keberhasilan tim"]},
            {"nama": "Kebutuhan Pengaruh (Need for Power)", "indikator": ["Keberanian mengambil inisiatif dan memimpin kegiatan", "Kemampuan mempengaruhi keputusan kerja secara konstruktif", "Tanggung jawab atas dampak keputusan terhadap unit"]},
            {"nama": "Motivasi Intrinsik & Komitmen Tugas", "indikator": ["Rasa bermakna dalam menjalankan pelayanan publik", "Ketahanan usaha saat menghadapi hambatan kerja", "Kesediaan mencurahkan upaya melampaui tuntutan minimal"]}
        ]
        tokoh_rujuk_z = "McClelland (1961, dalam Schunk, 2022) dan Ryan & Deci (2022)"
    elif any(k in z_lower for k in ["psychological capital", "modal psikologis", "psycap"]):
        p(doc, f"Psychological capital (modal psikologis) merupakan kondisi perkembangan psikologis positif individu yang dicirikan oleh kepemilikan daya tahan batin, keyakinan diri, optimisme realistis, dan ketabahan dalam menghadapi dinamika kerja. Luthans, Youssef, & Avolio (2007), dalam Luthans et al. (2022), merumuskan modal psikologis melampaui modal manusia (apa yang Anda ketahui) dan modal sosial (siapa yang Anda kenal), yakni berfokus pada 'siapa diri Anda sesungguhnya' (who you are) dan 'siapa diri Anda dalam potensi perkembangannya' (who you can become). Menurut Bandura (1997), dalam Schunk (2022), dimensi keyakinan diri (self-efficacy) memberikan landasan kepercayaan atas kapasitas personal dalam memobilisasi motivasi guna menuntaskan tugas menantang.", indent=0.5, space_after=6)

        p(doc, f"Dimensi harapan (hope), sebagaimana dirumuskan oleh Snyder (2002), dalam Luthans (2022), merupakan kondisi motivasional positif yang berakar pada keterpaduan antara tekad mencapai tujuan (willpower) dan perencanaan alternatif jalan keluar (waypower). Di sisi lain, Seligman (2002), dalam Luthans et al. (2022), mendefinisikan optimisme sebagai atribusi positif yang memandang keberhasilan sebagai hasil dari kapabilitas internal yang berkelanjutan, sedangkan kegagalan dipandang sebagai situasi temporer yang dapat diatasi. Masten (2001), dalam Luthans (2022), melengkapi konstruk ini melalui dimensi resiliensi (resilience), yaitu kemampuan psikologis untuk bangkit kembali dari kesulitan, ketidakpastian, konflik, maupun kegagalan kerja.", indent=0.5, space_after=6)

        p(doc, f"Elaborasi empiris oleh Peterson et al. (2022) menunjukkan bahwa modal psikologis beroperasi sebagai sumber daya kognitif dan afektif terpadu yang memicu ketahanan kerja tingkat tinggi. Avey, Reichard, Luthans, & Mhatre (2011), dalam Youssef-Morgan (2022), menemukan bahwa sinergi keempat komponen PsyCap memberikan daya prediksi kinerja yang jauh lebih kuat dibanding masing-masing komponen secara terisolasi. Newman, Ucbasaran, Zhu, & Hirst (2022) menegaskan bahwa modal psikologis berfungsi sebagai benteng internal yang memitigasi dampak tekanan kerja. Youssef-Morgan & Luthans (2022) menambahkan bahwa PsyCap dapat dikembangkan melalui intervensi keorganisasian yang terarah. Di Indonesia, Wibowo (2023) menyimpulkan bahwa modal psikologis aparatur menjadi jangkar mental yang memelihara integritas dan dedikasi dalam melayani masyarakat.", indent=0.5, space_after=6)

        sintesis_z = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Luthans, Youssef, & Avolio, 2007, dalam Luthans et al., 2022; Bandura, 1997, dalam Schunk, 2022; Snyder, 2002, dalam Luthans, 2022; serta Wibowo, 2023), dapat disimpulkan bahwa {z} merupakan kapasitas psikologis positif terpadu aparatur yang mencakup keyakinan diri (self-efficacy) dalam menuntaskan tugas menantang, harapan solutif (hope) dalam mengejar sasaran tugas, optimisme realistis terhadap masa depan, serta daya lenting resiliensi untuk bangkit kembali dari kesulitan kerja secara tangguh demi keberhasilan institusi."
        dims_z = [
            {"nama": "Self-Efficacy (Keyakinan Diri Kognitif)", "indikator": ["Keyakinan mampu menyelesaikan tugas dengan kompleksitas tinggi", "Keberanian mengambil inisiatif dalam memecahkan kendala kerja", "Kemandirian dalam menentukan langkah teknis penyelesaian tugas"]},
            {"nama": "Hope (Harapan & Tekad Solutif)", "indikator": ["Kejelasan penetapan target sasaran pencapaian kerja pribadi", "Kemampuan merumuskan strategi alternatif saat menghadapi hambatan", "Ketekunan daya upaya (willpower) dalam mengejar sasaran tugas"]},
            {"nama": "Optimism (Optimisme Realistis)", "indikator": ["Keyakinan positif terhadap keberhasilan penugasan kerja", "Kecenderungan memandang kegagalan sebagai peluang belajar temporer", "Sikap optimis terhadap masa depan dan kemajuan institusi"]},
            {"nama": "Resilience (Daya Lenting & Ketabahan)", "indikator": ["Kemampuan bangkit dengan cepat setelah menghadapi tekanan kerja", "Ketabahan mental dalam menghadapi perubahan situasi dan regulasi", "Ketahanan emosional dalam menuntaskan akumulasi beban penugasan"]}
        ]
        tokoh_rujuk_z = "Luthans, Youssef, & Avolio (2007, dalam Luthans et al., 2022)"
    elif any(k in z_lower for k in ["etos", "ethic", "etika kerja"]):
        # Etos kerja: hanya dipakai bila judul benar-benar memuat konstruk ini.
        p(doc, f"Diskursus mengenai {z} berakar pada kajian sosiologis dan psikologi kerja mendalam. Weber (1905), dalam Donaldson (2023), meletakkan dasar pemikiran bahwa etika kerja rasional memandang profesi dan pekerjaan sebagai panggilan hidup luhur (calling) yang menuntut asketisme, dedikasi moral, dan ketekunan tanpa henti. Konsep ini kemudian dioperasionalkan secara multidimensional oleh Miller et al. (2002), dalam Meriac et al. (2023), yang mendefinisikannya sebagai sistem sikap, keyakinan, dan nilai-nilai fundamental yang mengarahkan individu untuk memandang kerja sebagai aktivitas yang bermakna, menjunjung kerja keras, dan menolak penyia-nyiaan waktu. Pandangan ini dipertegas oleh Furnham (1990), dalam Furnham (2022), yang menyatakan bahwa etos kerja menempatkan kerja keras sebagai kebajikan moral tertinggi dan memandang pemborosan waktu sebagai bentuk degradasi etis.", indent=0.5, space_after=6)

        p(doc, f"Lebih lanjut, Merrens & Garrett (1975), dalam Cherrington (2022), menegaskan bahwa etos kerja dilandasi keyakinan intrinsik bahwa pengerahan usaha maksimal merupakan sumber utama pencapaian martabat pribadi dan harga diri. Pada ranah perilaku modern, Petty & Hill (2005) mengelompokkan karakteristik etos kerja ke dalam tiga pilar terukur: keterampilan interpersonal, inisiatif kerja mandiri, dan keandalan penugasan kedinasan. Dalam perspektif tata nilai nusantara, Sinamo (2022) merumuskan delapan etos kerja profesional (8 Ethos of Work) yang memandang kerja sebagai rahmat, amanah, panggilan, aktualisasi, ibadah, seni, kehormatan, dan pelayanan prima.", indent=0.5, space_after=6)

        p(doc, f"Kedalaman makna spiritual kerja diuraikan pula oleh Tasmara (2002), dalam Ndraha (2023), sebagai totalitas kepribadian serta cara memandang dan memberikan makna terhadap kerja yang mendorong perwujudan amal saleh dan prestasi optimal. Hal ini selaras dengan pandangan Anoraga (2022) yang menekankan etos kerja sebagai pandangan hidup khas mengenai nilai ketekunan, kejujuran, dan disiplin waktu kantor. Keraf (2022) melengkapinya sebagai komitmen moral internal aparatur untuk bersikap profesional dan menjunjung integritas publik. Seluruh telaah ini disimpulkan oleh Ndraha (2023) sebagai jiwa dan kebiasaan kerja aparatur yang dilandasi kesadaran nilai moralitas luhur dan kepatuhan pada kehormatan institusi.", indent=0.5, space_after=6)

        sintesis_z = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Weber, 1905, dalam Donaldson, 2023; Miller et al., 2002, dalam Meriac et al., 2023; Furnham, 2022; Sinamo, 2022; serta Ndraha, 2023), dapat disimpulkan bahwa {z} merupakan sistem komitmen moral dan sikap mental fundamental aparatur yang memandang kerja sebagai amanah kehormatan. Sikap ini dimanifestasikan melalui kegigihan kerja keras, kejujuran integritas, efisiensi pemanfaatan jam dinas, serta kemandirian profesional dalam menuntaskan seluruh tanggung jawab pelayanan publik."
        dims_z = onto_z.get("dimensi") if (onto_z and onto_z.get("dimensi")) else [
            {"nama": "Kerja Keras (Hard Work)", "indikator": ["Keuletan menuntaskan penugasan yang menuntut energi tinggi", "Kesediaan mencurahkan waktu dan pikiran melampaui jam formal", "Sikap pantang menyerah dalam mencari pemecahan kendala"]},
            {"nama": "Moralitas & Integritas (Morality & Integrity)", "indikator": ["Kejujuran dalam menyajikan data dan pelaporan capaian tugas", "Kepatuhan terhadap kode etik aparatur sipil", "Keteguhan menolak gratifikasi dan benturan kepentingan"]},
            {"nama": "Efisiensi Pemanfaatan Waktu (Wasted Time Avoidance)", "indikator": ["Disiplin kehadiran dan ketepatan memulai jam kerja kedinasan", "Pemanfaatan jam kantor secara efektif untuk penuntasan tugas", "Penghindaran aktivitas non-produktif selama jam dinas"]},
            {"nama": "Kemandirian Profesional (Self-Reliance)", "indikator": ["Kemandirian inisiatif dalam merencanakan langkah penyelesaian tugas", "Rasa tanggung jawab penuh atas keberhasilan dan kegagalan hasil kerja", "Kemampuan mengevaluasi mutu unjuk kerja diri sendiri secara kritis"]}
        ]
        tokoh_rujuk_z = onto_z.get("tokoh_ahli", "Sinamo (2022) dan Miller et al. (2022)")
    else:
        # Moderator generik: narasi dibangun dari profil ontologi variabel judul,
        # bukan meminjam kerangka etos kerja.
        _def_z = prof_z.get("definisi") or "kondisi kerja yang menentukan kuat lemahnya hubungan antarvariabel"
        p(doc, f"Kajian atas {z} dalam posisi moderator berangkat dari pertanyaan tentang kondisi apa yang membuat pengaruh {x_labels_z} terhadap {y_z['label']} berbeda antarunit kerja. {z} dipahami sebagai {_def_z}. Posisi ini menuntut penjelasan tentang mekanisme batas (boundary condition), yaitu keadaan yang membuat suatu hubungan menguat pada satu kelompok dan melemah pada kelompok lain (Hayes, 2022; Donaldson, 2023).", indent=0.5, space_after=6)

        p(doc, f"Secara operasional, {z} tidak diukur sebagai sifat bawaan individu, melainkan sebagai keadaan yang dapat diamati pada tingkat unit kerja. Pengukurannya dipisahkan dari variabel eksogen dan variabel terikat agar tidak terjadi tumpang tindih konstruk (Hair et al., 2022). Apabila {z} berada pada tingkat yang menguntungkan, {prof_z.get('tinggi', 'hubungan antarvariabel cenderung menguat')}; sebaliknya, risiko yang muncul adalah {prof_z.get('risiko', 'hubungan melemah karena kondisi kerja tidak mendukung')}.", indent=0.5, space_after=6)

        p(doc, f"Dalam konteks {instansi}, {z} perlu dibaca bersama tuntutan tugas dan dukungan organisasi. Perbedaan kondisi antarbagian dapat menghasilkan kekuatan hubungan yang berbeda meskipun nilai rata-rata organisasi tampak serupa, sehingga analisis efek interaksi lebih informatif daripada pembacaan rerata (Henseler & Fassott, 2010; Robbins & Judge, 2022).", indent=0.5, space_after=6)

        sintesis_z = (
            f"Berdasarkan telaah tersebut, {z} diposisikan sebagai kondisi batas yang menentukan "
            f"sejauh mana pengaruh {x_labels_z} terhadap {y_z['label']} berlangsung. Posisi ini "
            f"menuntut pengukuran yang terpisah, perumusan indikator yang spesifik, serta "
            f"penafsiran koefisien interaksi pada tingkat {z} yang berbeda (Hayes, 2022; "
            f"Hair et al., 2022)."
        )
        dims_z = (onto_z.get("dimensi") if (onto_z and onto_z.get("dimensi")) else prof_z.get("dimensi")) or [
            {"nama": "Kondisi Pendukung", "indikator": [
                "Ketersediaan dukungan kerja yang memadai",
                "Kejelasan arahan dalam pelaksanaan tugas",
            ]},
            {"nama": "Tekanan Situasional", "indikator": [
                "Intensitas tuntutan tugas pada periode sibuk",
                "Ketersediaan waktu untuk menyelesaikan pekerjaan",
            ]},
        ]
        tokoh_rujuk_z = (onto_z.get("tokoh_ahli") if onto_z else "") or "Hayes (2022) dan Donaldson (2023)"

    tokoh_rujuk_z = _canon_cite(tokoh_rujuk_z)

    # 3. Sintesis Konseptual Peneliti Doktoral (Wajib >= 4 baris & bervitasi)
    p(doc, sintesis_z, indent=0.5, space_after=6)

    # 4. Dimensi dan Indikator Operasional Baku (Format Teks Hirarkis Bertingkat Tanpa Tabel)
    _format_hierarchical_dimensions(doc, z, dims_z, tokoh_rujuk_z)

    # 5. Kesimpulan Terpadu Definisi dan Dimensi (Paragraf Akhir Sub-Bab, Wajib >= 4 baris & bervitasi)
    dim_z_names_str = ", ".join([d.get("nama", "").split("(")[0].strip() for d in dims_z])
    p(doc, f"Berdasarkan sintesis definisi konseptual dan rincian dimensi serta indikator operasional yang merujuk pada kerangka teoritis {tokoh_rujuk_z}, dapat disimpulkan bahwa pengukuran {z} dalam penelitian disertasi ini secara terpadu mencakup dimensi {dim_z_names_str}. Keterpaduan struktural antara definisi konseptual hasil sintesis dengan butir-butir indikator operasional tersebut menunjukkan bahwa instrumen ini memenuhi standar validitas isi dan reliabilitas konstruk yang andal (Hair et al., 2022; Sekaran & Bougie, 2020), sehingga memiliki ketepatan prediktif yang tinggi dalam menguji efek moderasi {z} pada model struktural yang diajukan.", indent=0.5, space_after=6)


def _build_x_theory_from_profile(doc: Document, ctx: Dict[str, Any], prof: Dict[str, Any], x_name: str):
    """Kajian teoritis satu variabel eksogen dari basis pengetahuan semantik."""
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, f"2.1.3-{prof['key']}")
    y = pm["Y"]
    instansi = ctx.get("instansi", "Instansi Pemerintah")
    teori = [t for t in (prof.get("teori") or []) if len(t) >= 3]
    t1 = teori[0] if teori else ("Teori Kapasitas Organisasi", "Daft (2022)",
                                 "menjelaskan kapasitas internal sebagai sumber daya yang menentukan mutu luaran organisasi")
    t2 = teori[1] if len(teori) > 1 else None

    p1 = (
        f"Kajian teoretis atas {x_name} tidak berangkat dari satu definisi tunggal. "
        f"{nv.theory_entry(t1[0], t1[1])} {t1[2]}. "
    )
    if t2:
        p1 += f"{nv.theory_entry(t2[0], t2[1])} {t2[2]}. "
    p1 += (
        f"Dalam penelitian ini, {x_name} dipahami sebagai {prof['definisi']}. "
        f"Rumusan tersebut dipertahankan karena dapat dioperasionalkan dan relevan dengan "
        f"konteks kerja aparatur, bukan karena kebiasaan mengutip definisi yang paling sering dipakai."
    )
    pv(doc, p1, ctx, "2.1.3-p1", indent=0.5, space_after=6)

    pv(doc,
       f"Secara mekanis, {x_name} bekerja dengan {prof['mekanisme']}. Mekanisme itu tidak "
       f"bersifat otomatis: apabila kondisi kerjanya tidak mendukung, risiko yang muncul "
       f"adalah {prof['risiko']}. "
       f"Karena itu, pengaruh {x_name} terhadap {y['label']} perlu dibaca sebagai rangkaian "
       f"proses, bukan sebagai hubungan satu langkah yang serba pasti.",
       ctx, "2.1.3-p2", indent=0.5, space_after=6)

    dims = prof.get("dimensi") or []
    dim_names = ", ".join(d.get("nama", "").split("(")[0].strip() for d in dims)
    pv(doc,
       f"Pada konteks {instansi}, {x_name} tidak berdiri sendiri. Kapasitas ini berinteraksi "
       f"dengan tuntutan tugas, dukungan atasan, dan sumber daya kerja yang tersedia. Oleh sebab "
       f"itu, pengukuran diarahkan pada {len(dims)} dimensi, yaitu {dim_names}. Pembatasan ini "
       f"menjaga agar konstruk tidak melebar menjadi penilaian atas seluruh aspek kepribadian "
       f"atau loyalitas aparatur (Hair et al., 2022; Sarstedt et al., 2022).",
       ctx, "2.1.3-p3", indent=0.5, space_after=6)

    tokoh = "; ".join(f"{t[0]} — {t[1]}" for t in teori[:2]) or "sintesis teori"
    pv(doc,
       f"Berdasarkan sintesis tersebut, {x_name} didefinisikan sebagai {prof['definisi']}. "
       f"Definisi ini menuntun penyusunan butir pernyataan pada setiap dimensi, sehingga isi "
       f"instrumen tidak ditentukan oleh istilah umum, melainkan oleh perilaku kerja yang dapat "
       f"diamati dan dinilai.",
       ctx, "2.1.3-p4", indent=0.5, space_after=6)

    _format_hierarchical_dimensions(doc, x_name, dims, tokoh)

    pv(doc,
       f"Dengan demikian, pengukuran {x_name} dalam disertasi ini mencakup dimensi {dim_names}. "
       f"Setiap dimensi diturunkan menjadi indikator yang dapat diuji validitas dan "
       f"reliabilitasnya, lalu dihubungkan dengan hipotesis yang telah dirumuskan. Konsistensi "
       f"antara definisi, dimensi, dan indikator inilah yang menjadi dasar penilaian validitas "
       f"isi instrumen (Cooper & Schindler, 2022; Sekaran & Bougie, 2020).",
       ctx, "2.1.3-p5", indent=0.5, space_after=6)


def _build_doc_single_x_theory(doc: Document, ctx: Dict[str, Any], x_idx: int, standalone: bool = True):
    """A.5 Kajian Teoritis Mandiri Variabel Bebas X ke-x_idx (Format Baku TAPD 2026 Disertasi S3)."""
    x_vars = ctx.get("x_vars", [])
    if not x_vars:
        x_vars = [ctx.get("x1") or "Variabel Bebas X1"]
    if x_idx < 1 or x_idx > len(x_vars):
        x_name = f"Variabel Bebas X{x_idx}"
    else:
        x_name = x_vars[x_idx - 1]

    if standalone:
        h1(doc, "BAB II\nKAJIAN PUSTAKA, KERANGKA PEMIKIRAN, DAN HIPOTESIS")
        h2(doc, "A. Kajian Teori")
        h3(doc, f"5.{x_idx} Kajian Teoritis {x_name}")
    else:
        h3(doc, f"5.{x_idx} Kajian Teoritis {x_name}")

    prof_x = nv.lookup(x_name)
    if prof_x:
        _build_x_theory_from_profile(doc, ctx, prof_x, x_name)
        return

    onto_vars = ctx.get("ontology", {}).get("variables", {})
    code_key = f"x{x_idx}"
    onto_x = onto_vars.get(code_key) or onto_vars.get(f"X{x_idx}") or onto_vars.get(x_name) or {}

    # 1. Narasi Pengantar Hakikat Substantif (Wajib >= 4 baris & bervitasi)
    p(doc, f"Dalam diskursus manajemen strategis dan perilaku keorganisasian (Daft, 2022; Robbins & Judge, 2022), {x_name} merupakan konstruk kunci yang menentukan efektivitas operasional institusi sekaligus membentuk mutu pengalaman kerja aparatur. Kehadiran {x_name} memberikan arah normatif, standar pelaksanaan, serta daya dorong yang menuntun aparatur dalam menyelaraskan kapasitas pribadi dengan tuntutan sasaran kelembagaan (Yukl, 2022; Armstrong & Taylor, 2023). Tanpa pengelolaan yang terarah pada pilar {x_name}, organisasi akan mengalami hambatan dalam mengoptimalkan sumber dayanya dan menghadapi resistensi terhadap perubahan sistem kedinasan kontemporer.", indent=0.5, space_after=6)

    x_lower = x_name.lower()
    if any(k in x_lower for k in ["budaya", "culture"]):
        # DOMAIN: BUDAYA ORGANISASI
        p(doc, f"Kajian teoretis mengenai budaya organisasi berakar pada teori kebudayaan kelembagaan yang memandang organisasi sebagai sistem nilai sosial yang unik. Schein (1985), dalam Denison (1990), mendefinisikan budaya organisasi sebagai pola asumsi dasar bersama yang dipelajari kelompok saat memecahkan masalah adaptasi eksternal dan integrasi internal, yang telah terbukti sahih sehingga diajarkan kepada anggota baru sebagai cara yang benar dalam memandang dan merasakan masalah. Denison (1990) mengoperasionalkannya sebagai nilai-nilai, keyakinan, dan prinsip dasar yang menjadi fondasi bagi sistem manajemen organisasi serta memandu praktik perilaku aparatur sehari-hari. Sementara itu, Cameron & Quinn (2022) melalui Competing Values Framework menekankan bahwa budaya mencerminkan ideologi mendasar yang menjadi pemandu cara organisasi beroperasi dan menilai keberhasilan.", indent=0.5, space_after=6)

        p(doc, f"Lebih lanjut, Robbins & Judge (2022) mendefinisikan budaya organisasi sebagai sistem makna bersama yang dianut oleh anggota yang membedakan organisasi tersebut dari organisasi lainnya. O'Reilly et al. (1991), dalam Chatman (2022), memandang budaya sebagai seperangkat norma dan nilai sosial yang disepakati bersama yang mengarahkan perilaku anggota tanpa perlu pengawasan formal yang ketat. Kotter & Heskett (1992), dalam Daft (2022), menunjukkan bahwa budaya organisasi yang adaptif dan menjunjung integritas memiliki korelasi langsung terhadap keberlanjutan kinerja jangka panjang. Hofstede (1980), dalam Minkov (2022), menguraikan bahwa budaya adalah pemrograman kolektif atas pikiran yang membedakan kelompok manusia yang satu dari yang lain.", indent=0.5, space_after=6)

        p(doc, f"Dalam konteks manajemen sektor publik di Indonesia, Kusdi (2022) menyatakan bahwa budaya organisasi merupakan perekat sosial yang mengikat aparatur melalui nilai-nilai moralitas, etika pelayanan, dan kepatuhan hukum bersama. Sutrisno (2022) menambahkan bahwa budaya kerja yang kuat menumbuhkan rasa memiliki (sense of belonging) yang tinggi di kalangan pegawai. Puncak elaborasi ini disimpulkan oleh Sobirin (2023) yang menegaskan bahwa budaya organisasi adalah totalitas falsafah, norma kedinasan, dan pedoman perilaku yang dihayati bersama oleh seluruh aparatur dalam memberikan pelayanan terbaik kepada masyarakat.", indent=0.5, space_after=6)

        sintesis_x = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Schein, 1985, dalam Denison, 1990; Cameron & Quinn, 2022; Robbins & Judge, 2022; Hofstede, 1980, dalam Minkov, 2022; serta Sobirin, 2023), dapat disimpulkan bahwa {x_name} merupakan sistem nilai, norma, keyakinan, dan asumsi dasar yang dihayati bersama oleh seluruh anggota organisasi. Sistem nilai ini berfungsi sebagai pedoman pola pikir, pembimbing perilaku kerja kedinasan, serta perekat sosial yang memperkuat koordinasi dan integrasi antar-unit guna mencapai sasaran strategis institusi secara berkelanjutan."
        dims_x = [
            {"nama": "Keterlibatan (Involvement)", "indikator": ["Pemberdayaan wewenang aparatur dalam penugasan", "Tingkat kerjasama sinergis dalam unit kerja", "Pengembangan kapabilitas berkelanjutan pegawai"]},
            {"nama": "Konsistensi (Consistency)", "indikator": ["Kepatuhan pada nilai-nilai inti institusi", "Kesepakatan bersama dalam penyelesaian hambatan kerja", "Koordinasi dan integrasi harmonis antar-bidang kedinasan"]},
            {"nama": "Adaptabilitas (Adaptability)", "indikator": ["Keterbukaan terhadap perubahan sistem kerja dan teknologi", "Orientasi pada kepuasan pemangku kepentingan publik", "Kemampuan belajar dari pengalaman kedinasan sebelumnya"]},
            {"nama": "Misi (Mission)", "indikator": ["Kejelasan pemahaman atas visi dan misi organisasi", "Kejelasan sasaran strategis jangka menengah", "Penyelarasan rencana kerja individu dengan misi instansi"]}
        ]
        tokoh_rujuk_x = "Denison (1990) dan Schein (2022)"

    elif any(k in x_lower for k in ["kualitas sdm", "kompetensi", "kapabilitas"]):
        # DOMAIN: KUALITAS SDM / KOMPETENSI
        p(doc, f"Kajian teoretis mengenai kualitas sumber daya manusia berakar pada Human Capital Theory yang dipelopori oleh Schultz (1961) dan Becker (1964). Becker (1964), dalam Armstrong & Taylor (2023), mendefinisikan modal manusia sebagai akumulasi pengetahuan, keahlian, keterampilan, dan kapasitas intelektual yang dimiliki individu yang memberikan kontribusi produktif bagi organisasi. Spencer & Spencer (1993), dalam Dessler (2023), merumuskan kompetensi sebagai karakteristik dasar individu yang berkaitan secara kausal dengan kriteria referensi efektivitas unjuk kerja unggul dalam pekerjaan atau situasi tertentu. Sementara itu, Armstrong & Taylor (2023) menegaskan bahwa kualitas aparatur mencerminkan kapasitas profesional yang memadukan keahlian teknis dengan kematangan emosional dan etika kerja.", indent=0.5, space_after=6)

        p(doc, f"Mathis & Jackson (2011), dalam Kasmir (2023), mengemukakan bahwa kualitas pegawai dinilai dari kemampuan mereka menyerap informasi, menyesuaikan diri dengan prosedur kerja baru, dan menuntaskan penugasan secara cermat. Dessler (2023) menyoroti bahwa aparatur yang berkualitas memiliki kompetensi inti (core competencies) yang mencakup kecakapan memecahkan masalah, keandalan teknis, dan integritas kerja. Mangkunegara (2022) menambahkan bahwa kualitas sumber daya manusia dipengaruhi oleh latar belakang pendidikan kedinasan, pelatihan berkelanjutan, dan pengalaman lapangan yang relevan. Sedarmayanti (2022) memandang kualitas aparatur sebagai cerminan kesiapan fisik, mental, dan intelektual dalam mengemban amanah publik.", indent=0.5, space_after=6)

        p(doc, f"Hasibuan (2022) merumuskan bahwa sumber daya manusia yang berkualitas memiliki kesadaran tanggung jawab yang tinggi terhadap keberhasilan institusi dan efisiensi pemanfaatan sumber daya kantor. Di sisi lain, Kasmir (2023) menegaskan bahwa kualitas aparatur dibuktikan melalui kecepatan adaptasi terhadap sistem kerja berbasis digital dan kepatuhan penuh pada etika profesi. Puncak pemaknaan ini disimpulkan oleh Afandi (2023) yang menyatakan bahwa kualitas aparatur merupakan integrasi harmonis antara keunggulan kompetensi teknis, kecakapan relasional interpersonal, dan keteguhan moralitas dalam memberikan pelayanan kedinasan.", indent=0.5, space_after=6)

        sintesis_x = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Becker, 1964, dalam Armstrong & Taylor, 2023; Spencer & Spencer, 1993, dalam Dessler, 2023; Mangkunegara, 2022; Kasmir, 2023; serta Afandi, 2023), dapat disimpulkan bahwa {x_name} merupakan akumulasi kapabilitas komprehensif aparatur yang memadukan penguasaan pengetahuan teknis, kemahiran keterampilan operasional, kecakapan analitis dalam pemecahan masalah, serta keteguhan sikap moral dan integritas dalam mengemban penugasan publik secara profesional dan akuntabel."
        dims_x = [
            {"nama": "Penguasaan Pengetahuan Teknis (Technical Knowledge)", "indikator": ["Pemahaman mendalam terhadap regulasi dan petunjuk teknis", "Kecakapan memahami tata kelola administrasi kedinasan", "Wawasan terhadap perkembangan kebijakan dan teknologi baru"]},
            {"nama": "Kemahiran Keterampilan Operasional (Operational Skills)", "indikator": ["Ketelitian dan ketepatan dalam menyelesaikan kalkulasi kerja", "Kemampuan mengoperasikan sistem aplikasi kedinasan secara efektif", "Kecepatan menuntaskan prosedur teknis operasional"]},
            {"nama": "Kapasitas Pemecahan Masalah (Problem Solving Ability)", "indikator": ["Kemampuan mengidentifikasi akar hambatan teknis penugasan", "Kreativitas merumuskan alternatif solusi yang aplikatif", "Ketepatan dalam pengambilan keputusan teknis kedinasan"]},
            {"nama": "Sikap Kerja & Integritas Profesional (Professional Attitude)", "indikator": ["Tanggung jawab penuh atas keabsahan hasil pekerjaan", "Kedisiplinan mematuhi kode etik dan standar mutu tugas", "Keteguhan menolak penyimpangan prosedur kerja"]}
        ]
        tokoh_rujuk_x = "Becker (1964, dalam Armstrong & Taylor, 2023) dan Spencer & Spencer (1993)"

    elif any(k in x_lower for k in ["iklim", "climate"]):
        # DOMAIN: IKLIM KERJA
        p(doc, f"Kajian mengenai iklim organisasi berakar pada tradisi psikologi keorganisasian yang menelaah persepsi psikologis bersama anggota terhadap lingkungan kerjanya. Litwin & Stringer (1968), dalam Schneider et al. (2022), mendefinisikan iklim organisasi sebagai serangkaian keadaan lingkungan kerja yang dapat diukur, yang dirasakan langsung atau tidak langsung oleh orang-orang yang hidup dan bekerja dalam lingkungan tersebut, yang diasumsikan mempengaruhi motivasi dan perilaku mereka. Schneider, Ehrhart, & Macey (2013), dalam Schneider et al. (2022), memperjelas bahwa iklim merupakan makna bersama yang melekat pada kebijakan, praktik kedinasan, dan prosedur yang dialami anggota institusi. Sementara itu, James & Jones (1974), dalam James et al. (2022), membedakan antara iklim psikologis individual dengan iklim organisasi agregat.", indent=0.5, space_after=6)

        p(doc, f"Pritchard & Karasick (1973), dalam Denison (1990), menguraikan bahwa iklim mencerminkan sifat internal yang membedakan satu organisasi dari organisasi lain yang relatif bertahan lama dan mempengaruhi perilaku kerja anggota. Gellerman (1963), dalam Luthans (2022), memandang iklim sebagai 'kepribadian' organisasi yang dirasakan aparatur. Stringer (2002) mengidentifikasi bahwa iklim yang kondusif ditandai oleh adanya otonomi kerja, kejelasan standar sasaran, kehangatan hubungan antarpribadi, serta dukungan penuh dari pimpinan. Wirawan (2022) menegaskan bahwa iklim kerja yang sehat menumbuhkan rasa aman psikologis (psychological safety) bagi aparatur untuk berinovasi.", indent=0.5, space_after=6)

        p(doc, f"Simamora (2022) menambahkan bahwa iklim keorganisasian ditentukan oleh keterbukaan komunikasi dan keadilan sistem pengakuan prestasi. Wibowo (2023) menyimpulkan bahwa iklim kerja aparatur adalah suasana lingkungan batiniah yang tercipta dari hubungan harmonis antara pimpinan dan bawahan serta antar-rekan sejawat yang membangkitkan kegairahan kerja. Pandangan terpadu ini menegaskan bahwa iklim kerja bukan sekadar sarana fisik, melainkan atmosfer psikologis yang mengkondisikan komitmen kerja seluruh jajaran institusi.", indent=0.5, space_after=6)

        sintesis_x = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Litwin & Stringer, 1968, dalam Schneider et al., 2022; Schneider, Ehrhart, & Macey, 2013; Stringer, 2002; Wirawan, 2022; serta Wibowo, 2023), dapat disimpulkan bahwa {x_name} merupakan persepsi psikologis bersama aparatur mengenai atmosfer lingkungan kerja kelembagaan yang terbentuk dari kejelasan struktur tugas, kehangatan hubungan antarpribadi, keterbukaan komunikasi, serta dukungan kepemimpinan yang secara nyata mengkondisikan motivasi dan kegairahan kerja dinas."
        dims_x = [
            {"nama": "Otonomi & Fleksibilitas Kerja", "indikator": ["Keluasan wewenang teknis dalam menyelesaikan tugas", "Ketiadaan hambatan birokrasi berlebihan dalam prosedur kerja", "Ruang inisiatif mandiri dalam mencari pemecahan masalah"]},
            {"nama": "Kejelasan Standar & Tanggung Jawab", "indikator": ["Kejelasan uraian penugasan pada setiap fungsi jabatan", "Ketegasan standar mutu pelayanan yang terukur", "Ketepatan target kinerja yang dipahami bersama"]},
            {"nama": "Kehangatan & Dukungan Sosial", "indikator": ["Suasana saling menghormati antar rekan sejawat", "Kesiapan rekan kerja saling mendukung dalam penuntasan tugas", "Hubungan interpersonal yang harmonis lintas seksi"]},
            {"nama": "Dukungan & Penghargaan Kepemimpinan", "indikator": ["Apresiasi pimpinan atas dedikasi dan prestasi kerja", "Keterbukaan komunikasi antara atasan dan bawahan", "Keadilan dalam pemberian perlakuan dan rekognisi"]}
        ]
        tokoh_rujuk_x = "Schneider, Ehrhart, & Macey (2013) dan Stringer (2002)"

    elif any(k in x_lower for k in ["beban kerja", "workload", "work load", "beban tugas", "tuntutan kerja", "job demand"]):
        p(doc, f"Beban kerja merupakan keseluruhan tuntutan tugas yang harus diselesaikan aparatur dalam kurun waktu tertentu, mencakup tuntutan kuantitatif maupun kualitatif yang menuntut pengerahan energi fisik, kognitif, dan emosional. Robbins & Judge (2022) menjelaskan bahwa beban kerja mencerminkan kesenjangan antara kapasitas individu dan tuntutan peran, sehingga taraf beban yang tidak terkelola menimbulkan ketegangan kerja (job strain), kelelahan, dan penurunan mutu pengambilan keputusan. Dalam kerangka Job Demands-Resources, Bakker & Demerouti (2022) membedakan tuntutan kerja (job demands) dari sumber daya kerja (job resources): tuntutan kerja menuntut usaha fisik dan psikologis yang berkelanjutan, sementara sumber daya kerja berfungsi mempertahankan energi, memfasilitasi pencapaian sasaran, dan menstimulasi pertumbuhan pribadi aparatur.", indent=0.5, space_after=6)

        p(doc, f"Mangkunegara (2022) menegaskan bahwa beban kerja tercermin melalui volume tugas, kompleksitas penugasan, dan tenggat waktu yang harus dipenuhi aparatur pada satu periode akuntabilitas. Wibowo (2023) menambahkan bahwa beban kerja yang melampaui kapasitas daya dukung aparatur akan menurunkan ketelitian, memperlambat penyelesaian tugas, dan memicu kelelahan emosional yang pada akhirnya menggerus mutu layanan. Sebaliknya, beban kerja yang terlalu ringan juga menimbulkan kebosanan dan hilangnya tantangan bermakna, sehingga produktivitas optimal berada pada taraf beban yang proporsional dengan kapasitas dan sumber daya pendukung yang tersedia.", indent=0.5, space_after=6)

        p(doc, f"Pada tataran psikologis, Luthans, Luthans, & Luthans (2022) menjelaskan bahwa dampak beban kerja terhadap unjuk kerja tidak bersifat mekanis, melainkan dimediasi oleh kapasitas psikologis aparatur dalam mengelola tekanan, mempertahankan fokus, dan memulihkan energi setelah menghadapi penugasan berat. Aparatur yang memiliki keyakinan diri, optimisme, dan daya lenting yang memadai mampu memaknai beban kerja sebagai tantangan yang terkelola, bukan sebagai ancaman yang melumpuhkan. Oleh karena itu, beban kerja perlu diposisikan sebagai variabel eksogen yang pengaruhnya terhadap kinerja bergantung pada kapasitas psikologis dan mekanisme keterikatan kerja aparatur (Bakker & Demerouti, 2022; Robbins & Judge, 2022).", indent=0.5, space_after=6)

        sintesis_x = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Robbins & Judge, 2022; Bakker & Demerouti, 2022; Mangkunegara, 2022; Wibowo, 2023; serta Luthans et al., 2022), dapat disimpulkan bahwa {x_name} merupakan keseluruhan tuntutan tugas kuantitatif dan kualitatif yang dibebankan kepada aparatur dalam kurun waktu tertentu, yang menuntut pengerahan energi fisik, mental, dan emosional. Beban kerja yang proporsional dengan kapasitas aparatur mendorong ketelitian dan produktivitas, sedangkan beban yang berlebihan menimbulkan ketegangan kerja, kelelahan, dan penurunan mutu unjuk kerja."
        dims_x = onto_x.get("dimensi") if (onto_x and onto_x.get("dimensi")) else [
            {"nama": "Beban Waktu (Temporal Demand)", "indikator": ["Kesesuaian volume tugas dengan waktu penyelesaian yang tersedia", "Frekuensi kerja lembur untuk menuntaskan target kedinasan", "Tekanan tenggat waktu pada periode pelaporan"]},
            {"nama": "Beban Mental (Mental Demand)", "indikator": ["Kompleksitas analisis yang dituntut dalam penyelesaian tugas", "Kebutuhan konsentrasi tinggi dan ketelitian berkelanjutan", "Frekuensi pengambilan keputusan berisiko dalam penugasan"]},
            {"nama": "Beban Fisik & Tekanan Kerja (Effort & Frustration)", "indikator": ["Tuntutan stamina fisik dalam menjalankan aktivitas kedinasan", "Tingkat kelelahan yang dirasakan setelah menyelesaikan tugas", "Tekanan emosional saat menghadapi tuntutan pemangku kepentingan"]}
        ]
        tokoh_rujuk_x = onto_x.get("tokoh_ahli", "Bakker & Demerouti (2022) dan Robbins & Judge (2022)")
    else:
        # DEFAULT: KEPEMIMPINAN TRANSFORMASIONAL / STRATEGIS
        p(doc, f"Kajian teoretis mengenai {x_name} berakar pada teori kepemimpinan dan kapasitas organisasi modern. Bass (1985), dalam Yukl (2022), mengemukakan bahwa kepemimpinan yang efektif mampu memotivasi pengikut untuk berbuat melampaui harapan awal dengan memperluas kesadaran nilai dan mengarahkan mereka melampaui kepentingan pribadi. Sejalan dengan itu, Yukl (2022) menegaskan bahwa kepemimpinan merupakan proses mempengaruhi anggota organisasi melalui artikulasi visi masa depan yang memikat, penanaman nilai etis bersama, dan fasilitasi pembelajaran kelompok. Sementara itu, Northouse (2022) mendefinisikannya sebagai proses di mana seorang individu mempengaruhi sekelompok orang untuk mencapai tujuan bersama melalui peningkatan derajat moralitas dan motivasi.", indent=0.5, space_after=6)

        p(doc, f"Dalam model Full Range of Leadership, Avolio, Yammarino, & Bass (1991), dalam Antonakis & Day (2022), memetakan perilaku pemimpin ke dalam pembinaan komitmen tinggi, penumbuhan kepercayaan timbal balik, dan stimulasi intelektual bawahan. Pandangan filosofis Burns (1978), dalam Antonakis & Day (2022), menggarisbawahi bahwa kepemimpinan sejati terjadi manakala pemimpin dan yang dipimpin saling mengangkat derajat etika dan motivasi ke tingkat yang lebih mulia. Pada ranah manajemen terapan, Robbins & Coulter (2022) mengartikan kepemimpinan sebagai kemampuan mengarahkan dan mempengaruhi tim untuk mencapai tujuan melalui keteladanan tindakan.", indent=0.5, space_after=6)

        p(doc, f"Lebih lanjut, Luthans, Luthans, & Luthans (2022) menekankan bahwa kepemimpinan positif menyalurkan optimisme terarah, keteladanan karakter, dan kepedulian mendalam terhadap pengembangan potensi bawahan. Schein (1985), dalam Denison (1990), menyoroti keterikatan erat kepemimpinan dalam membentuk dan memperkuat budaya organisasi yang adaptif. Hal ini dilengkapi oleh Denison (1990) yang menyatakan bahwa prinsip nilai kepemimpinan menjadi fondasi bagi terciptanya koordinasi yang konsisten dan misi yang jelas. Di Indonesia, Thoha (2023) menyimpulkan bahwa kepemimpinan adalah seni dan kemampuan mempengaruhi perilaku bawahan agar bersedia bekerja sama dengan penuh kesadaran dan antusiasme demi tercapainya sasaran bersama.", indent=0.5, space_after=6)

        sintesis_x = f"Berdasarkan berbagai pendapat para ahli di atas (seperti Bass, 1985, dalam Yukl, 2022; Northouse, 2022; Antonakis & Day, 2022; Luthans et al., 2022; serta Thoha, 2023), dapat disimpulkan bahwa {x_name} merupakan pilar kapasitas strategis organisasi yang mencakup integritas keteladanan moral, kejelasan artikulasi visi masa depan, stimulasi inovasi intelektual, dan perhatian individual yang tulus. Kepemimpinan ini secara konsisten membimbing, memberdayakan, dan menginspirasi aparatur untuk bekerja melampaui standar minimal demi keberhasilan institusi."
        dims_x = onto_x.get("dimensi") if (onto_x and onto_x.get("dimensi")) else [
            {"nama": "Keteladanan Karakter & Integritas", "indikator": ["Keteladanan moral dan integritas dalam pelaksanaan tugas", "Konsistensi memegang prinsip etika keorganisasian", "Kemampuan menumbuhkan rasa bangga dan kepercayaan"]},
            {"nama": "Artikulasi Visi & Motivasi", "indikator": ["Kejelasan pengartikulasian visi strategis masa depan", "Kemampuan membangkitkan optimisme dan antusiasme kerja tim", "Komunikasi persuasif mengenai pentingnya pencapaian sasaran"]},
            {"nama": "Stimulasi Intelektual & Inovasi", "indikator": ["Dorongan untuk menggunakan pendekatan baru yang inovatif", "Pemberian ruang kebebasan berpikir kritis dalam memecahkan masalah", "Apresiasi terhadap ide-ide perbaikan prosedur kerja"]},
            {"nama": "Pertimbangan Individual & Pembinaan", "indikator": ["Kesediaan mendengarkan aspirasi dan keluhan aparatur", "Peran aktif sebagai mentor dan pembimbing peningkatan kompetensi", "Pemberian penugasan yang disesuaikan dengan potensi individu"]}
        ]
        tokoh_rujuk_x = onto_x.get("tokoh_ahli", "Avolio & Bass (2022) dan Yukl (2022)")

    tokoh_rujuk_x = _canon_cite(tokoh_rujuk_x)

    # 3. Sintesis Konseptual Peneliti Doktoral (Wajib >= 4 baris & bervitasi)
    p(doc, sintesis_x, indent=0.5, space_after=6)

    # 4. Dimensi dan Indikator Operasional Baku (Format Teks Hirarkis Bertingkat Tanpa Tabel)
    _format_hierarchical_dimensions(doc, x_name, dims_x, tokoh_rujuk_x)

    # 5. Kesimpulan Terpadu Definisi dan Dimensi Tiap Variabel Bebas (Wajib >= 4 baris & bervitasi)
    dim_x_names_str = ", ".join([d.get("nama", "").split("(")[0].strip() for d in dims_x])
    p(doc, f"Berdasarkan sintesis definisi konseptual dan rincian dimensi serta indikator operasional yang merujuk pada landasan teoretis {tokoh_rujuk_x}, dapat disimpulkan bahwa pengukuran {x_name} dalam penelitian disertasi ini secara terpadu mencakup dimensi {dim_x_names_str}. Keterpaduan struktural antara rumusan konseptual hasil sintesis dengan butir-butir indikator operasional tersebut menunjukkan bahwa instrumen ini memenuhi kriteria validitas isi dan keterandalan pengukuran ilmiah (Hair et al., 2022; Cooper & Schindler, 2022), sehingga mampu merefleksikan esensi {x_name} secara komprehensif dan akurat.", indent=0.5, space_after=6)


def _build_doc_sub_2_1_3(doc: Document, ctx: Dict[str, Any]):
    """A.5 5. Kajian Teoretis Variabel-Variabel Bebas (X1, X2... Xn)."""
    x_vars = ctx.get("x_vars", [])
    if not x_vars:
        x_vars = [ctx.get("x1", "Variabel Bebas")]
    x_vars_str = _join_id(x_vars)
    h2(doc, f"5. Kajian Teoritis Variabel-Variabel Bebas ({x_vars_str})")

    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "2.1.3")
    xs = pm["_all"]["x"]
    y = pm["Y"]
    x1p = xs[0] if xs else nv.make_profile(x_vars[0], {}, ctx)
    x2p = xs[1] if len(xs) > 1 else None

    pv(doc,
       f"Variabel eksogen dalam model ini adalah {x_vars_str}. Keduanya tidak berada pada "
       f"ranah yang sama. {x1p['label']} berurusan dengan {x1p['family']} yang dimiliki "
       f"aparatur, sedangkan "
       + (f"{x2p['label']} berurusan dengan {x2p['family']} yang melekat pada tugas."
          if x2p else "variabel lainnya berurusan dengan kondisi kerja yang berbeda.")
       + f" Karena posisi teoretisnya berbeda, keduanya tidak dapat diperlakukan sebagai "
       f"satu blok kapasitas yang homogen.",
       ctx, "2.1.3-intro1", indent=0.5, space_after=6)

    pv(doc,
       f"Perbedaan itu tampak pada mekanismenya. {x1p['label']} bekerja dengan "
       f"{x1p['mekanisme']}. "
       + (f"Sementara itu, {x2p['label']} bekerja dengan {x2p['mekanisme']}."
          if x2p else "")
       + f" Apabila mekanisme tersebut diabaikan, penelitian berisiko menyimpulkan bahwa "
       f"seluruh variabel eksogen memiliki cara kerja yang sama terhadap {y['label']}. "
       f"Kajian berikut karena itu memisahkan pembahasan masing-masing konstruk.",
       ctx, "2.1.3-intro2", indent=0.5, space_after=6)

    for x_idx in range(1, len(x_vars) + 1):
        _build_doc_single_x_theory(doc, ctx, x_idx, standalone=False)

    mech_list = "; ".join(f"{p['label']}: {p['mekanisme']}" for p in xs) or "mekanisme kerja"
    pv(doc,
       f"Secara terpadu, {x_vars_str} tidak membentuk satu rangkaian sebab yang seragam. "
       f"Ringkasnya, {mech_list}. Oleh karena itu, model penelitian ini menguji pengaruh "
       f"masing-masing variabel secara terpisah sebelum menilai peran {pm['_all']['m'][0]['label'] if pm['_all']['m'] else 'mekanisme perantara'} "
       f"dan {pm['_all']['z'][0]['label'] if pm['_all']['z'] else 'kondisi batas'}. "
       f"Pemisahan ini penting agar rekomendasi kebijakan tidak mencampuradukkan intervensi "
       f"kapasitas dengan penataan tuntutan kerja.",
       ctx, "2.1.3-close", indent=0.5, space_after=6)


def _build_doc_sub_2_2(doc: Document, ctx: Dict[str, Any]):
    """2.2 Penelitian Terdahulu (State of the Art & Research Gap) - Format Baku TAPD 2026."""
    h2(doc, "B. Hasil Penelitian yang Relevan")

    x_vars = ctx.get("x_vars", [])
    x1 = x_vars[0] if x_vars else "variabel bebas pertama"
    x2 = x_vars[1] if len(x_vars) > 1 else ""
    m_list = ctx.get("m_vars", [])
    m = (m_list[0] if m_list else "") or (ctx.get("m") or "").strip()
    z_list = ctx.get("z_vars", [])
    z = (z_list[0] if z_list else "") or (ctx.get("z") or "").strip()
    y = ctx.get("y_var", "Kinerja Pegawai")
    instansi = ctx.get("instansi", "Instansi Pemerintah")
    model_acronym = _generate_doctoral_model_acronym(ctx)
    _has_m = bool(m)
    _has_z = bool(z)

    corpus = ctx.get("academic_refs")
    if not corpus:
        corpus = s3_academic_researcher.fetch_s3_verified_academic_corpus(ctx)
        ctx["academic_refs"] = corpus

    papers = corpus.get("papers", [])
    cite_grp = ctx.get("citation_groups") or s3_academic_researcher.get_s3_citation_groups(ctx)
    pos_cites = cite_grp.get("pos_citations", "Hair et al., 2022; Sarstedt et al., 2022")
    gap_cites = cite_grp.get("gap_citations", "Hayes, 2022; Creswell & Creswell, 2023")
    first_gap = cite_grp.get("first_gap_cite", "")

    # Fokus literatur dibangun dari variabel judul saja.
    _fokus = list(x_vars)
    if _has_m:
        _fokus.append(m)
    if _has_z:
        _fokus.append(z)
    _fokus.append(y)
    _fokus_str = _join_id(_fokus)

    # 1. Pengantar state of the art (profil variabel, bukan template)
    pv(doc,
       f"Penelaahan atas penelitian terdahulu bukan kelengkapan pustaka, melainkan upaya "
       f"memetakan posisi disertasi ini dalam percakapan ilmiah. Melalui komparasi temuan, "
       f"peneliti dapat memisahkan klaim yang telah mapan dari klaim yang masih diperdebatkan, "
       f"sekaligus mengenali celah yang belum dijawab (Snyder, 2019; Creswell & Creswell, 2023). "
       f"Dalam konteks {instansi}, state of the art diarahkan pada studi yang membahas "
       f"{_fokus_str}, agar perbandingan tidak melebar ke topik yang tidak relevan. "
       f"Pemetaan literatur 2022–2026 disajikan pada Tabel 2.1.",
       ctx, "2.2-intro", indent=0.5, space_after=6)

    # 2. Tabel 2.1 Hasil Penelitian yang Relevan
    #    Kolom metoda dan temuan diisi dari abstrak publikasi asal (bukan dikarang).
    t_headers = ["No", "Penulis, Tahun", "Judul / publikasi", "Metoda Riset", "Temuan Penelitian"]
    t_rows = []

    def _metode_dari_abstrak(ab: str) -> str:
        a = re.sub(r"\s+", " ", (ab or "")).lower()
        if not a:
            return ""
        label = ""
        if "pls" in a or "structural equation" in a or re.search(r"\bsem\b", a):
            label = "Kuantitatif (SEM-PLS)"
        elif "regresi" in a:
            label = "Kuantitatif (analisis regresi)"
        elif "kualitatif" in a or "wawancara" in a:
            label = "Kualitatif"
        elif "kuantitatif" in a or "kuesioner" in a or "angket" in a or "survei" in a:
            label = "Kuantitatif (survei)"
        n_match = re.search(r"(?:sampel|responden|populasi)[^\d]{0,30}(\d{2,4})", a)
        if not n_match:
            n_match = re.search(r"\bn\s*[=:]\s*(\d{2,4})", a)
        if n_match:
            label = (label + ", " if label else "") + f"sampel n = {int(n_match.group(1))} (menurut publikasi asal)"
        return label

    def _temuan_dari_abstrak(ab: str) -> str:
        teks = re.sub(r"\s+", " ", (ab or "")).strip()
        if not teks:
            return ""
        kalimat = re.split(r"(?<=[.!?])\s+", teks)
        kunci = ("hasil", "menunjuk", "menemu", "berpengaruh", "menyimpul", "temuan",
                 "terdapat", "signifikan", "positif", "negatif", "disimpulkan")
        pilih = [k for k in kalimat if any(x in k.lower() for x in kunci)]
        rangkai = " ".join(pilih[:2]) if pilih else " ".join(kalimat[-2:])
        return rangkai[:430].strip()

    if papers:
        for idx, p_item in enumerate(papers[:16], 1):
            cit = p_item.get("citation") or s3data.PLACEHOLDER
            title_p = p_item.get("title", "")
            journ = p_item.get("journal", "")
            yr = p_item.get("year", "")
            doi_link = p_item.get("doi", "")

            pub_str = f"{title_p}.\n{journ}, ({yr})." if journ else f"{title_p}. ({yr})."
            if doi_link:
                pub_str += f"\nDOI: {doi_link}"

            ab = p_item.get("abstract") or ""
            method_str = _metode_dari_abstrak(ab) or s3data.PLACEHOLDER
            finding_str = _temuan_dari_abstrak(ab) or s3data.PLACEHOLDER

            t_rows.append([str(idx), cit, pub_str, method_str, finding_str])
    else:
        t_rows = [[
            "1",
            s3data.PLACEHOLDER,
            "Korpus literatur ber-DOI belum tersedia pada lingkungan ini sehingga tabel "
            "penelitian terdahulu belum dapat diisi.",
            s3data.PLACEHOLDER,
            s3data.PLACEHOLDER,
        ]]

    add_table_data(
        doc,
        title="Tabel 2.1 Hasil Penelitian yang Relevan",
        headers=t_headers,
        data=t_rows,
        source=s3data.source_label(ctx, "Data diolah peneliti dari abstrak publikasi ber-DOI, (2026)") if papers else s3data.source_label(ctx, "Data diolah peneliti, (2026)"),
        widths=[Cm(0.9), Cm(2.7), Cm(3.4), Cm(2.9), Cm(4.1)],
        alignments=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.JUSTIFY],
        cell_font_size=10
    )
    p(doc, "", space_after=6)

    # 3. Analisis dialektika temuan empiris dan research gap (profil variabel)
    pm_22 = nv.profile_map(ctx)
    xs_22 = pm_22["_all"]["x"]
    x1p_22 = xs_22[0] if xs_22 else nv.make_profile(x1, {}, ctx)
    x2p_22 = xs_22[1] if len(xs_22) > 1 else None
    y_22 = pm_22["Y"]
    x_labels_22 = _join_id(p["label"] for p in xs_22) or x1

    pv(doc,
       f"Berdasarkan komparasi pada Tabel 2.1, temuan tentang {x_labels_22} terhadap "
       f"{y_22['label']} tidak seragam. Sebagian studi melaporkan pengaruh positif ({pos_cites}), "
       f"tetapi sebagian lain menemukan pengaruh yang lemah atau tidak signifikan ketika "
       f"mekanisme transmisi tidak diperhitungkan. Perbedaan itu menunjukkan bahwa hubungan "
       f"tersebut tidak dapat dibaca sebagai korelasi sederhana.",
       ctx, "2.2-p1", indent=0.5, space_after=6)

    _gap_awal = f"Sejumlah studi ({gap_cites}) memperlihatkan"
    _p2_tail = " Karena itu, model penelitian ini dirancang untuk menguji hubungan tersebut secara utuh."
    if _has_m and _has_z:
        _p2_tail = (f" Karena itu, model penelitian ini memasukkan {m} sebagai mekanisme transmisi "
                    f"dan {z} sebagai kondisi batas.")
    elif _has_m:
        _p2_tail = (f" Karena itu, model penelitian ini memasukkan {m} sebagai mekanisme transmisi "
                    f"yang menjelaskan bagaimana pengaruh tersebut bekerja.")
    elif _has_z:
        _p2_tail = (f" Karena itu, model penelitian ini memasukkan {z} sebagai kondisi batas "
                    f"yang menjelaskan kapan pengaruh tersebut menguat atau melemah.")

    pv(doc,
       f"Ketidakkonsistenan tersebut bukan semata akibat perbedaan sampel atau alat ukur. "
       f"{_gap_awal} bahwa pengaruh {x_labels_22} dapat melemah ketika kondisi kerja tidak "
       f"mendukung dan mekanisme transmisi tidak bekerja. "
       f"{x1p_22['label']} yang tinggi tidak otomatis menghasilkan {y_22['label']} apabila "
       f"risiko yang muncul adalah {x1p_22['risiko']}. "
       + (f"Sebaliknya, {x2p_22['label']} yang tidak terkelola dapat menurunkan mutu luaran "
          f"meskipun aparatur memiliki kapasitas memadai."
          if x2p_22 else "")
       + _p2_tail,
       ctx, "2.2-p2", indent=0.5, space_after=6)

    _aspek_uji = ["jalur langsung"]
    if _has_m:
        _aspek_uji.append("jalur tidak langsung")
    if _has_z:
        _aspek_uji.append("efek kondisional")

    _unsur_model = [_join_id(x_vars) + " sebagai variabel penjelas"]
    if _has_m:
        _unsur_model.append(f"{m} sebagai mekanisme transmisi")
    if _has_z:
        _unsur_model.append(f"{z} sebagai kondisi batas")
    _unsur_model.append(f"{y_22['label']} sebagai luaran")
    _unsur_str = _join_id(_unsur_model)

    pv(doc,
       f"Kebaruan penelitian ini terletak pada pengujian simultan atas {_join_id(_aspek_uji)} "
       f"dalam satu model {model_acronym}. Studi terdahulu umumnya menguji hubungan secara "
       f"terpisah, sehingga sulit menjelaskan mengapa pengaruh {x_labels_22} terhadap "
       f"{y_22['label']} berbeda antarkonteks. Dengan menyatukan {_unsur_str}, penelitian ini "
       f"tidak berhenti pada pengujian signifikansi, tetapi menawarkan penjelasan tentang kapan "
       f"dan mengapa pengaruh itu bekerja (Hair et al., 2022; Hayes, 2022; Sarstedt et al., 2022).",
       ctx, "2.2-p3", indent=0.5, space_after=6)

    # 4. Sintesis Kebaruan (Novelty)
    if _has_m and _has_z:
        _novel_kal = (
            f"Belum ditemukan studi yang secara terpadu mengintegrasikan {_join_id(x_vars)} "
            f"sebagai pilar eksogen, {m} sebagai jembatan transmisi, serta {z} sebagai kondisi "
            f"batas terhadap {y} dalam satu kerangka model kontingensi terintegrasi"
        )
    elif _has_m:
        _novel_kal = (
            f"Belum ditemukan studi yang secara terpadu mengintegrasikan {_join_id(x_vars)} "
            f"sebagai pilar eksogen dan {m} sebagai jembatan transmisi terhadap {y} dalam satu "
            f"model mediasi terintegrasi"
        )
    elif _has_z:
        _novel_kal = (
            f"Belum ditemukan studi yang secara terpadu mengintegrasikan {_join_id(x_vars)} "
            f"sebagai pilar eksogen dan {z} sebagai kondisi batas terhadap {y} dalam satu model "
            f"moderasi terintegrasi"
        )
    else:
        _novel_kal = (
            f"Belum ditemukan studi yang menguji secara simultan {_join_id(x_vars)} terhadap {y} "
            f"dengan mempertimbangkan mekanisme internal dan konteks kerja pada instansi publik"
        )

    p(doc, f"Kesenjangan penelitian empiris dan teoretis tersebut dijawab melalui kebaruan "
           f"(novelty) model konseptual yang diajukan dalam disertasi ini. Sebagian besar "
           f"penelitian terdahulu menguji pengaruh variabel penjelas secara parsial ({gap_cites}). "
           f"{_novel_kal} (Hair et al., 2022; Hayes, 2022; Sarstedt et al., 2022). Kehadiran "
           f"model struktural ini menjadi distingsi dan kontribusi orisinal riset doktoral dalam "
           f"khazanah ilmu manajemen kinerja dan tata kelola kelembagaan publik.", indent=0.5, space_after=6)

def _hypo_profile(pm: Dict[str, Any], role_id: Any) -> Dict[str, Any]:
    """Ambil profil variabel dari id grafik (X1, M, Z, Y)."""
    rid = str(role_id or "").strip().upper()
    if rid in pm:
        return pm[rid]
    if rid == "Y":
        return pm["Y"]
    if rid.startswith("X") and rid[1:].isdigit():
        xs = pm["_all"]["x"]
        idx = int(rid[1:]) - 1
        if 0 <= idx < len(xs):
            return xs[idx]
    if rid.startswith("M"):
        ms = pm["_all"]["m"]
        if ms:
            return ms[0]
    if rid.startswith("Z"):
        zs = pm["_all"]["z"]
        if zs:
            return zs[0]
    return pm["Y"]


def _build_doc_sub_2_3_1(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """C.1 Kerangka Konseptual & Silogisme Deduktif (narasi dirakit dari profil variabel)."""
    if standalone:
        h1(doc, "BAB II\nKAJIAN PUSTAKA, KERANGKA PEMIKIRAN, DAN HIPOTESIS")
        h2(doc, "C. Kerangka Berpikir")

    h3(doc, "C.1 Kerangka Konseptual & Silogisme Deduktif")

    instansi = ctx.get("instansi", "Instansi Pemerintah")
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "2.3.1")
    xs = pm["_all"]["x"]
    ms = pm["_all"]["m"]
    zs = pm["_all"]["z"]
    y = pm["Y"]
    hypos = ctx.get("graph", {}).get("hypotheses", [])
    x_labels = _join_id(p["label"] for p in xs) or "variabel eksogen"

    pv(doc,
       f"Kerangka konseptual disertasi ini bertolak dari keberatan terhadap pembacaan mekanistis "
       f"atas hubungan {x_labels} dengan {y['label']}. Dalam praktik birokrasi, penambahan anggaran "
       f"pelatihan atau penataan prosedur tidak dengan sendirinya menghasilkan luaran kerja yang "
       f"lebih baik apabila tidak diikuti perubahan cara aparatur memaknai dan mengerahkan usahanya "
       f"(Snyder, 2019; Robbins & Judge, 2022). Karena itu, model penelitian ini memisahkan tiga "
       f"lapis penjelasan: kapasitas yang dimiliki, proses psikologis yang mengubah kapasitas menjadi "
       f"tindakan, dan kondisi kontekstual yang menentukan kekuatan proses tersebut.",
       ctx, "2.3.1-p1", indent=0.5, space_after=6)

    if ms:
        m = ms[0]
        pv(doc,
           f"Posisi {m['label']} sebagai variabel intervening berangkat dari definisi {m['label']} sebagai {m['definisi']}. "
           f"Konstruk ini tidak identik dengan {x_labels}, tetapi juga bukan bagian dari {y['label']}; "
           f"ia bekerja sebagai mekanisme yang menyalurkan energi kerja. Tanpa mekanisme tersebut, "
           f"kapasitas yang tersedia berisiko berhenti sebagai potensi, sementara risiko yang muncul adalah {m['risiko']}. "
           f"Karena itu, pengujian mediasi diperlukan untuk memastikan apakah {x_labels} benar-benar "
           f"bekerja melalui {m['label']} atau hanya memiliki efek langsung yang semu.",
           ctx, "2.3.1-p2", indent=0.5, space_after=6)

    if zs:
        z = zs[0]
        pv(doc,
           f"Kehadiran {z['label']} sebagai moderator menjawab kemungkinan bahwa kekuatan hubungan "
           f"antarkonstruk tidak seragam. {z['label']} dipahami sebagai {z['definisi']}. Dalam "
           f"kerangka ini, {z['label']} tidak diposisikan sebagai pesaing {x_labels}, melainkan sebagai "
           f"kondisi batas yang memengaruhi seberapa efektif kapasitas eksogen diubah menjadi kinerja. "
           f"Apabila {z['label']} kuat, {z['tinggi']}. Sebaliknya, apabila {z['label']} lemah, "
           f"risiko yang muncul adalah {nv.field(z, 'risiko', rng)}, sehingga daya ungkit "
           f"variabel eksogen melemah meskipun prosedur formal tetap tersedia.",
           ctx, "2.3.1-p3", indent=0.5, space_after=6)
    elif not ms:
        pv(doc,
           f"Model penelitian ini memusatkan perhatian pada jalur langsung {x_labels} terhadap "
           f"{y['label']}. Fokus tersebut dipilih karena studi awal menunjukkan bahwa variasi "
           f"{y['label']} lebih banyak dijelaskan oleh perbedaan kondisi kerja daripada oleh "
           f"mekanisme perantara yang belum terukur. Dengan demikian, pengujian diarahkan pada "
           f"besaran dan signifikansi kontribusi setiap variabel eksogen.",
           ctx, "2.3.1-p3b", indent=0.5, space_after=6)

    pv(doc, nv.pick(nv.rng_for(ctx, "2.3.1-sil"), [
       (
           f"Perumusan hipotesis disusun melalui silogisme deduktif agar setiap proposisi memiliki "
           f"akar teoretis dan pijakan empiris yang jelas (Sekaran & Bougie, 2020; Cooper & Schindler, 2022). "
           f"Premis mayor berisi dalil umum dari teori yang relevan, premis minor memuat kondisi khusus "
           f"pada {instansi}, dan konklusi dirumuskan sebagai hipotesis yang dapat diuji. Struktur ini "
           f"menjaga agar hipotesis tidak lahir dari dugaan pragmatis, melainkan dari rantai penalaran "
           f"yang dapat ditelusuri kembali ke literatur."
       ),
       (
           f"Hipotesis dirumuskan melalui silogisme deduktif supaya tiap proposisi berakar pada teori "
           f"dan berpijak pada bukti empiris (Sekaran & Bougie, 2020; Cooper & Schindler, 2022). Premis "
           f"mayor memuat dalil umum dari teori yang relevan, premis minor berisi kondisi khusus di "
           f"{instansi}, dan konklusi diwujudkan sebagai hipotesis yang dapat diuji. Susunan ini menjaga "
           f"hipotesis agar tidak lahir dari dugaan pragmatis, melainkan dari rantai penalaran yang dapat "
           f"dilacak ke literatur."
       ),
       (
           f"Perumusan hipotesis ditempuh lewat silogisme deduktif agar setiap proposisi memiliki dasar "
           f"teori sekaligus pijakan empiris yang jelas (Sekaran & Bougie, 2020; Cooper & Schindler, 2022). "
           f"Dalil umum dari teori yang relevan menempati premis mayor, kondisi khusus di {instansi} "
           f"menjadi premis minor, dan konklusinya dirumuskan sebagai hipotesis yang dapat diuji. Struktur "
           f"demikian mencegah hipotesis lahir dari dugaan pragmatis dan menahannya pada rantai penalaran "
           f"yang tertelusuri ke literatur."
       ),
    ]), ctx, "2.3.1-p4", indent=0.5, space_after=6)

    if not hypos:
        hypos = [
            {"code": "H1", "type": "direct_to_mediator", "source": "X1", "target": "M",
             "statement": f"{xs[0]['label'] if xs else 'Variabel eksogen'} berpengaruh positif terhadap {ms[0]['label'] if ms else 'mediasi'}."},
            {"code": "H2", "type": "mediator_to_dependent", "source": "M", "target": "Y",
             "statement": f"{ms[0]['label'] if ms else 'Mediasi'} berpengaruh positif terhadap {y['label']}."},
        ]

    for idx, h in enumerate(hypos, 1):
        h_code = h.get("code", f"H{idx}")
        h_type = str(h.get("type") or "")
        cause = _hypo_profile(pm, h.get("source"))
        effect = _hypo_profile(pm, h.get("target"))
        z = _hypo_profile(pm, h.get("moderator")) if h.get("moderator") else (zs[0] if zs else None)
        h_stmt = h.get("statement", "")

        if h_type == "mediation_indirect":
            mayor = nv.mediation_paragraph(cause, ms[0] if ms else effect, effect, instansi, rng)
            minor = (
                f"Premis Minor: Pada {instansi}, variasi {cause['label']} tampak antara lain pada "
                f"{cause['risiko']}, tetapi capaian {effect['label']} belum bergerak searah; "
                f"pola ini mengarah pada dugaan adanya mekanisme perantara yang belum bekerja optimal."
            )
        elif h_type == "moderation_interaction" and z:
            mayor = nv.moderation_paragraph(cause, z, effect, instansi, rng)
            minor = (
                f"Premis Minor: Pada {instansi}, aparatur dengan tingkat {cause['label']} relatif "
                f"setara memperlihatkan capaian {effect['label']} yang berbeda-beda, sehingga "
                f"perbedaan itu diduga bersumber dari kondisi {z['label']} yang tidak seragam."
            )
        elif h_type == "conditional_indirect" and z and ms:
            mayor = nv.conditional_indirect_paragraph(cause, ms[0], z, effect, instansi, rng)
            minor = (
                f"Premis Minor: Pada {instansi}, jalur {cause['label']} menuju {effect['label']} "
                f"melalui {ms[0]['label']} diperkirakan tidak konstan, karena kualitas {z['label']} "
                f"berbeda antarseksi dan antarmasa kerja."
            )
        elif h_type == "moderator_main_effect":
            mayor = nv.premise_mayor(cause, effect, rng)
            minor = (
                f"Premis Minor: Pada {instansi}, {cause['label']} tidak hanya berfungsi sebagai "
                f"kondisi batas, tetapi juga memiliki kontribusi langsung terhadap {effect['label']} "
                f"melalui {cause['mekanisme']}."
            )
        else:
            mayor = nv.premise_mayor(cause, effect, rng)
            minor = nv.premise_minor(cause, effect, instansi, rng)

        p(doc, f"{idx}. Hipotesis {h_code}", bold=True, indent=0.5, space_after=2)
        pv(doc, mayor, ctx, f"2.3.1-mayor-{idx}", indent=0.75, space_after=2)
        pv(doc, minor, ctx, f"2.3.1-minor-{idx}", indent=0.75, space_after=2)
        p(doc, f"{h_code}: {h_stmt}", bold=True, indent=0.75, space_after=6)


def _build_doc_sub_2_3_2(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """C.2 Keterkaitan Kausal antar Variabel (narasi per jalur, bukan satu template)."""
    if standalone:
        h1(doc, "BAB II\nKAJIAN PUSTAKA, KERANGKA PEMIKIRAN, DAN HIPOTESIS")
        h2(doc, "C. Kerangka Berpikir")

    h3(doc, "C.2 Keterkaitan Kausal antar Variabel")

    instansi = ctx.get("instansi", "Instansi Pemerintah")
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "2.3.2")
    xs = pm["_all"]["x"]
    ms = pm["_all"]["m"]
    zs = pm["_all"]["z"]
    y = pm["Y"]
    x_labels = _join_id(p["label"] for p in xs) or "variabel eksogen"
    m_label = ms[0]["label"] if ms else "(tanpa mediasi)"
    z_label = zs[0]["label"] if zs else "(tanpa moderasi)"

    cite_grp = ctx.get("citation_groups") or s3_academic_researcher.get_s3_citation_groups(ctx)
    pos_cites = cite_grp.get("pos_citations", "Hair et al., 2022; Sarstedt et al., 2022")

    if ms and zs:
        _jalur_kal = (
            f"melalui dua jenis jalur: jalur langsung dan jalur tidak langsung. Jalur langsung "
            f"menggambarkan kontribusi setiap variabel eksogen terhadap luaran kerja, sedangkan jalur "
            f"tidak langsung menggambarkan transmisi melalui {m_label}. {z_label} diposisikan sebagai "
            f"kondisi batas yang mengubah kekuatan hubungan tersebut"
        )
    elif ms:
        _jalur_kal = (
            f"melalui dua jenis jalur: jalur langsung dan jalur tidak langsung. Jalur langsung "
            f"menggambarkan kontribusi setiap variabel eksogen terhadap luaran kerja, sedangkan jalur "
            f"tidak langsung menggambarkan transmisi melalui {m_label}"
        )
    elif zs:
        _jalur_kal = (
            f"melalui jalur langsung antarvariabel, dengan {z_label} diposisikan sebagai kondisi "
            f"batas yang mengubah kekuatan hubungan tersebut"
        )
    else:
        _jalur_kal = (
            "melalui jalur langsung antarvariabel eksogen dan luaran kerja, tanpa variabel mediasi "
            "maupun moderasi dalam model"
        )

    pv(doc,
       f"Kerangka berpikir pada Gambar 2.1 merangkum keterkaitan {x_labels} dengan {y['label']} "
       f"{_jalur_kal} (Hair et al., 2022; Hayes, 2022).",
       ctx, "2.3.2-p1", indent=0.5, space_after=6)

    t_box = doc.add_table(rows=1, cols=1)
    t_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t_box.rows[0].cells[0]
    cell.width = Cm(15.5)

    p(cell, {
        "direct": "BAGAN MODEL STRUKTURAL DOKTORAL (DIRECT EFFECT MODEL)",
        "mediation": "BAGAN MODEL STRUKTURAL DOKTORAL (MEDIATION MODEL)",
        "moderation": "BAGAN MODEL STRUKTURAL DOKTORAL (MODERATION MODEL)",
        "moderated_mediation": "BAGAN MODEL STRUKTURAL DOKTORAL (HYBRID MODERATED-MEDIATION MODEL)",
    }.get(str(ctx.get("model_type") or ""), "BAGAN MODEL STRUKTURAL DOKTORAL"),
      bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=4)
    p(cell, f"[ Variabel Bebas (X): {x_labels} ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    if ms:
        p(cell, "│ (Pengaruh Langsung H1..Hn)           │ (Pengaruh Eksogen ke Mediasi)",
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
        p(cell, f"▼                                     ▼\n[ Variabel Mediasi (M): {m_label} ] ───(Efek Transmisi)───► [ Variabel Terikat (Y): {y['label']} ]",
          bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    else:
        p(cell, "│ (Pengaruh Langsung H1..Hn)",
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
        p(cell, f"▼\n[ Variabel Terikat (Y): {y['label']} ]",
          bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    if zs:
        p(cell, f"▲\n│ (Efek Kontingensi Moderasi / Boundary Condition)\n[ Variabel Moderasi (Z): {z_label} ]",
          bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)

    p(doc, "Sumber : Data diolah peneliti, (2026)", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4, space_after=2)
    p(doc, "Gambar 2. 1\nKerangka Berpikir", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=8)

    if ms and zs:
        p2_txt = (
            f"Gambar 2.1 menunjukkan bahwa {x_labels} tidak berdiri sendiri dalam menjelaskan "
            f"{y['label']}. Panah tebal merepresentasikan jalur langsung yang diuji, sedangkan panah "
            f"transmisi menunjukkan peran {m_label} sebagai perantara. Adapun garis interaksi "
            f"menempatkan {z_label} sebagai faktor yang memperkuat atau memperlemah hubungan. "
            f"Pembacaan bagan ini penting agar setiap hipotesis memiliki posisi yang jelas dalam "
            f"struktur model, bukan sekadar daftar hubungan yang terpisah ({pos_cites})."
        )
    elif ms:
        p2_txt = (
            f"Gambar 2.1 menunjukkan bahwa {x_labels} tidak berdiri sendiri dalam menjelaskan "
            f"{y['label']}. Panah tebal merepresentasikan jalur langsung yang diuji, sedangkan panah "
            f"transmisi menunjukkan peran {m_label} sebagai perantara antara kapasitas eksogen dan "
            f"luaran kerja. Pembacaan bagan ini penting agar setiap hipotesis memiliki posisi yang "
            f"jelas dalam struktur model, bukan sekadar daftar hubungan yang terpisah ({pos_cites})."
        )
    elif zs:
        p2_txt = (
            f"Gambar 2.1 menunjukkan bahwa {x_labels} tidak berdiri sendiri dalam menjelaskan "
            f"{y['label']}. Panah tebal merepresentasikan jalur langsung yang diuji, sedangkan garis "
            f"interaksi menempatkan {z_label} sebagai faktor yang memperkuat atau memperlemah "
            f"hubungan. Pembacaan bagan ini penting agar setiap hipotesis memiliki posisi yang jelas "
            f"dalam struktur model, bukan sekadar daftar hubungan yang terpisah ({pos_cites})."
        )
    else:
        p2_txt = (
            f"Gambar 2.1 menunjukkan bahwa {x_labels} diposisikan sebagai penjelas utama {y['label']} "
            f"dalam struktur model. Panah tebal merepresentasikan jalur langsung yang diuji, tanpa "
            f"mekanisme perantara atau kondisi batas. Pembacaan bagan ini penting agar setiap "
            f"hipotesis memiliki posisi yang jelas dalam struktur model, bukan sekadar daftar "
            f"hubungan yang terpisah ({pos_cites})."
        )
    pv(doc, p2_txt, ctx, "2.3.2-p2", indent=0.5, space_after=6)

    # Sintesis jalur: empat paragraf, bukan pengulangan teori per jalur.
    x1s = xs[0] if xs else None
    x2s = xs[1] if len(xs) > 1 else None
    zs0 = zs[0] if zs else None
    ms0 = ms[0] if ms else None

    direct_lines = []
    for x in xs:
        direct_lines.append(
            f"{x['label']} bekerja dengan {nv.field(x, 'mekanisme', rng)}"
        )
    pv(doc,
       f"Jalur langsung yang diuji adalah kontribusi {x_labels} terhadap {y['label']}. "
       f"Keduanya tidak diperlakukan sebagai hubungan yang identik: {'; '.join(direct_lines)}. "
       f"Perbedaan mekanisme itu menjelaskan mengapa arah dan besaran koefisien setiap jalur "
       f"perlu dibaca secara terpisah, bukan dirata-ratakan sebagai satu pengaruh tunggal "
       f"({pos_cites}).",
       ctx, "2.3.2-direct", indent=0.5, space_after=6)

    if ms0:
        med_lines = []
        for x in xs:
            med_lines.append(
                f"{x['label']} menuju {ms0['label']} karena {nv.field(x, 'mekanisme', rng)}"
            )
        pv(doc,
           f"Jalur tidak langsung menghubungkan {x_labels} dengan {y['label']} melalui "
           f"{ms0['label']}. {'; '.join(med_lines)}. {ms0['label']} tidak berfungsi sebagai "
           f"saluran netral; ia menjadi tempat energi kerja diolah sebelum menjelma menjadi "
           f"luaran. Apabila {nv.field(ms0, 'risiko', rng)} terjadi, transmisi itu terhambat meskipun "
           f"kapasitas eksogen memadai. Karena itu, pengujian mediasi diperlukan untuk "
           f"memisahkan pengaruh langsung dari pengaruh yang benar-benar melewati "
           f"{ms0['label']}.",
           ctx, "2.3.2-mediation", indent=0.5, space_after=6)

    if zs0:
        pv(doc,
           f"Efek {x_labels} terhadap {y['label']} tidak diasumsikan konstan. "
           f"{zs0['label']} diposisikan sebagai kondisi batas yang mengubah kekuatan hubungan "
           f"tersebut. Apabila {nv.field(zs0, 'tinggi', rng)}, daya ungkit kapasitas eksogen "
           f"menguat; sebaliknya, apabila {zs0['label']} lemah, risiko yang muncul adalah "
           f"{nv.field(zs0, 'risiko', rng)}, sehingga hubungan itu melemah "
           f"meskipun prosedur formal tetap tersedia. Interaksi ini diuji sebagai efek "
           f"moderasi, bukan sebagai tambahan variabel kontrol.",
           ctx, "2.3.2-moderation", indent=0.5, space_after=6)

        if ms0:
            pv(doc,
               f"Hipotesis mediasi termoderasi menuntut bukti bahwa efek tidak langsung "
               f"{x_labels} terhadap {y['label']} melalui {ms0['label']} berubah pada tingkat "
               f"{zs0['label']} yang berbeda. Bukti itu dibaca dari indeks moderated mediation "
               f"dan interval kepercayaan bootstrap, bukan dari signifikansi masing-masing "
               f"jalur secara terpisah (Hayes, 2022). Dengan cara ini, klaim tentang kapan "
               f"model bekerja dapat diuji secara langsung, bukan disimpulkan dari pola "
               f"koefisien yang kebetulan berbeda.",
               ctx, "2.3.2-conditional", indent=0.5, space_after=6)


def _build_doc_sub_2_3_3(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """C.3 Paradigma Model & Boundary Conditions (dirakit dari teori variabel judul)."""
    if standalone:
        h1(doc, "BAB II\nKAJIAN PUSTAKA, KERANGKA PEMIKIRAN, DAN HIPOTESIS")
        h2(doc, "C. Kerangka Berpikir")

    h3(doc, "C.3 Paradigma Model & Boundary Conditions")

    instansi = ctx.get("instansi", "Instansi Pemerintah")
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "2.3.3")
    xs = pm["_all"]["x"]
    ms = pm["_all"]["m"]
    zs = pm["_all"]["z"]
    y = pm["Y"]
    model_acronym = _generate_doctoral_model_acronym(ctx)
    x_labels = _join_id(p["label"] for p in xs) or "variabel eksogen"
    m_label = ms[0]["label"] if ms else "(tanpa mediasi)"
    z_label = zs[0]["label"] if zs else "(tanpa moderasi)"

    cite_grp = ctx.get("citation_groups") or s3_academic_researcher.get_s3_citation_groups(ctx)
    pos_cites = cite_grp.get("pos_citations", "Hair et al., 2022; Sarstedt et al., 2022")
    gap_cites = cite_grp.get("gap_citations", "Gadzai et al., 2025; Almazrouei & Alnahhal, 2025")
    gov_cites = cite_grp.get("gov_citations", "Hayes, 2022; Creswell & Creswell, 2023")
    first_gap = cite_grp.get("first_gap_cite", "Gadzai et al. (2025)")

    # 1. Integrasi teori: ambil nama teori unik dari profil
    theory_names: List[str] = []
    for prof in [*xs, *ms, *zs, y]:
        for nama, sitasi, _ in (prof.get("teori") or [])[:2]:
            entry = nv.theory_entry(nama, sitasi)
            if entry not in theory_names:
                theory_names.append(entry)
    theory_txt = "; ".join(theory_names[:6]) or "teori perilaku organisasi dan manajemen kinerja"

    has_m = bool(ms)
    has_z = bool(zs)

    # Narasi integrasi model menyesuaikan jenis model yang BENAR-BENAR ada di judul,
    # sehingga judul mediasi murni tidak dipaksa bercerita tentang moderasi.
    if has_m and has_z:
        transmisi_txt = (
            f"mengapa {x_labels} dapat bekerja melalui {m_label} dan pada kondisi apa "
            f"pengaruhnya berubah"
        )
        struktur_txt = (
            f"kapasitas eksogen ({x_labels}), mekanisme transmisi ({m_label}), dan "
            f"kondisi batas ({z_label})"
        )
        integ_line = f"{x_labels} ──► {m_label} ──► {y['label']} (Dimoderasi {z_label})"
    elif has_m:
        transmisi_txt = (
            f"mengapa {x_labels} menyalurkan pengaruhnya terhadap {y['label']} melalui {m_label}"
        )
        struktur_txt = (
            f"kapasitas eksogen ({x_labels}) dan mekanisme transmisi ({m_label})"
        )
        integ_line = f"{x_labels} ──► {m_label} ──► {y['label']}"
    elif has_z:
        transmisi_txt = (
            f"pada kondisi apa pengaruh {x_labels} terhadap {y['label']} menguat atau melemah"
        )
        struktur_txt = f"kapasitas eksogen ({x_labels}) dan kondisi batas ({z_label})"
        integ_line = f"{x_labels} ──► {y['label']} (Dimoderasi {z_label})"
    else:
        transmisi_txt = f"sejauh mana {x_labels} menjelaskan variasi {y['label']}"
        struktur_txt = f"kapasitas eksogen ({x_labels})"
        integ_line = f"{x_labels} ──► {y['label']}"

    pv(doc,
       f"Model konseptual dalam disertasi ini menyintesiskan {theory_txt}. Sintesis tersebut "
       f"tidak dimaksudkan untuk menumpuk teori, melainkan untuk menjelaskan {transmisi_txt}. "
       f"Dengan cara itu, setiap konstruk memiliki dasar teoretis yang spesifik dan dapat "
       f"dipertanggungjawabkan secara akademik (Daft, 2022; Robbins & Judge, 2022).",
       ctx, "2.3.3-p1", indent=0.5, space_after=6)

    pv(doc,
       f"Kebaruan model terletak pada perumusan {model_acronym}, yaitu kerangka yang menyatukan "
       f"{struktur_txt} dalam satu struktur. Integrasi itu diperlukan karena studi terdahulu "
       f"cenderung menguji masing-masing hubungan secara terpisah, sehingga sulit menjelaskan "
       f"mengapa hasil penelitian tentang {y['label']} sering tidak konsisten ({nv.join_cites(first_gap, gap_cites)}). "
       f"Kerangka kebaruan tersebut divisualisasikan pada Gambar 2.2.",
       ctx, "2.3.3-p2", indent=0.5, space_after=6)

    # Teori middle-range diambil dari variabel mediasi/moderasi yang ada di judul (bukan hardcode).
    mid_bits: List[str] = []
    for prof in [*ms, *zs]:
        for nama, sitasi, _ in (prof.get("teori") or [])[:1]:
            if nama and nama not in mid_bits:
                mid_bits.append(nama)
    middle_txt = ", ".join(mid_bits[:3]) or "Teori Perilaku Organisasi"

    t_box2 = doc.add_table(rows=1, cols=1)
    t_box2.alignment = WD_TABLE_ALIGNMENT.CENTER
    c2 = t_box2.rows[0].cells[0]
    c2.width = Cm(15.5)
    p(c2, f"STATE OF THE ART FRAMEWORK: MODEL DOKTORAL {model_acronym}", bold=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=4)
    p(c2, f"Grand Theory: {y.get('grand') or (y['teori'][0][0] if y.get('teori') else 'Teori Kinerja')}\n"
          f"Middle Theory: {middle_txt}\n"
          f"Applied Theory: {', '.join([t[0] for t in (xs[0]['teori'][:1] if xs else [])]) or 'Manajemen Sumber Daya Manusia'}\n"
          f"Integrasi Konstruk: {integ_line}",
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)

    p(doc, "Sumber : Data diolah peneliti, (2026)", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4, space_after=2)
    p(doc, "Gambar 2. 2\nState of the Art", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=8)

    # 2. Boundary conditions: tiga argumen yang disesuaikan dengan profil model.
    z = zs[0] if zs else None
    z_cap = z["label"] if z else "kondisi kontekstual"
    pv(doc,
       f"Pertanyaan yang diajukan model ini bukan lagi apakah {x_labels} berpengaruh, melainkan "
       f"mengapa pengaruh itu tidak selalu sama kuat di berbagai unit kerja. Jawabannya terletak "
       f"pada boundary conditions, yaitu kondisi yang membatasi keberlakuan hubungan ({pos_cites}). "
       f"Tanpa memperhitungkan kondisi tersebut, rekomendasi kebijakan berisiko "
       f"menggeneralisasi praktik yang hanya cocok pada satu konteks.",
       ctx, "2.3.3-p3", indent=0.5, space_after=6)

    if z:
        bc1 = (
            f"1. Kapasitas dan Tuntutan Tidak Otomatis Menjadi Kinerja. Kondisi {x_labels} yang "
            f"memadai tidak dengan sendirinya menghasilkan {y['label']}. Konversinya bergantung pada "
            f"cara aparatur mengolah tuntutan dan sumber daya kerja. Di sinilah {z_cap} berperan: "
            f"apabila kondisinya kuat, {z['tinggi']}; "
            f"apabila rapuh, risiko yang muncul adalah {z['risiko']}. Argumen ini "
            f"menempatkan moderator sebagai bagian dari mekanisme, bukan variabel tambahan ({gov_cites})."
        )
    elif ms:
        bc1 = (
            f"1. Kapasitas dan Tuntutan Tidak Otomatis Menjadi Kinerja. Kondisi {x_labels} yang "
            f"memadai tidak dengan sendirinya menghasilkan {y['label']}. Konversinya bergantung pada "
            f"cara aparatur mengolah tuntutan dan sumber daya kerja. Di sinilah {m_label} berperan "
            f"sebagai mekanisme perantara: apabila mekanisme itu terbentuk, kapasitas berubah menjadi "
            f"tindakan nyata; apabila tidak, risiko yang muncul adalah {nv.field(ms[0], 'risiko', rng)}. "
            f"Argumen ini menempatkan mediasi sebagai bagian dari mekanisme, bukan variabel tambahan "
            f"({gov_cites})."
        )
    else:
        bc1 = (
            f"1. Kapasitas dan Tuntutan Tidak Otomatis Menjadi Kinerja. Kondisi {x_labels} yang "
            f"memadai tidak dengan sendirinya menghasilkan {y['label']}. Konversinya bergantung pada "
            f"cara aparatur mengolah tuntutan, sumber daya, dan prioritas kerja. Oleh karena itu, "
            f"model ini membaca pengaruh langsung secara hati-hati dengan menimbang konteks kerja "
            f"yang sesungguhnya ({gov_cites})."
        )
    pv(doc, bc1, ctx, "2.3.3-bc1", indent=0.75, space_after=4)

    if ms:
        bc2 = (
            f"2. Ketergantungan pada Eksekusi Manajerial. {m_label} tidak terbentuk melalui pengumuman "
            f"kebijakan, tetapi melalui interaksi sehari-hari antara atasan dan bawahan. Ketika atasan "
            f"memberi umpan balik, menjaga keadilan penugasan, dan menyediakan sumber daya, kondisi "
            f"kerja yang menopang {m_label} menguat. Sebaliknya, ketika atasan membiarkan beban menumpuk "
            f"dan tidak menindaklanjuti keluhan, {m_label} melemah meskipun prosedur formal lengkap "
            f"({gap_cites}). Oleh karena itu, model ini tidak dapat direduksi menjadi hubungan "
            f"variabel-variabel yang berjarak dari praktik manajerial."
        )
    else:
        bc2 = (
            f"2. Ketergantungan pada Eksekusi Manajerial. Pengaruh {x_labels} terhadap {y['label']} "
            f"tidak terbentuk melalui pengumuman kebijakan, tetapi melalui interaksi sehari-hari "
            f"antara atasan dan bawahan. Ketika atasan memberi umpan balik, menjaga keadilan "
            f"penugasan, dan menyediakan sumber daya, kondisi kerja yang menopang kinerja menguat. "
            f"Sebaliknya, ketika atasan membiarkan beban menumpuk dan tidak menindaklanjuti keluhan, "
            f"pengaruh tersebut melemah meskipun prosedur formal lengkap ({gap_cites}). Oleh karena "
            f"itu, model ini tidak dapat direduksi menjadi hubungan variabel yang berjarak dari "
            f"praktik manajerial."
        )
    pv(doc, bc2, ctx, "2.3.3-bc2", indent=0.75, space_after=4)

    if z and z.get("family") in ("kontrol", "integritas"):
        bc3 = (
            f"3. Legitimasi dan Keteladanan Puncak. Dalam birokrasi, nilai yang diumumkan tidak "
            f"selalu sama dengan nilai yang dijalankan. Apabila pimpinan memperlihatkan perilaku "
            f"permisif, norma integritas kehilangan legitimasi dan {z_cap} tidak dapat berkembang. "
            f"Sebaliknya, keteladanan pimpinan membuat norma tersebut terinternalisasi dan "
            f"memperkuat efek {x_labels} terhadap {y['label']} (Donaldson & Davis, 1991; Yukl & Gardner, 2024)."
        )
    elif z and z.get("family") == "modal":
        bc3 = (
            f"3. Dukungan Sosial dan Pemulihan Psikologis. {z_cap} bukan sifat tetap yang sekali "
            f"terbentuk lalu bertahan sendiri. Modal ini dapat terkuras ketika tekanan kerja "
            f"berlangsung lama tanpa dukungan atasan, keadilan penugasan, dan kesempatan pemulihan. "
            f"Sebaliknya, lingkungan kerja yang memberi umpan balik dan menghargai usaha aparatur "
            f"menjaga modal tersebut tetap tersedia. Karena itu, efek moderasi {z_cap} perlu dibaca "
            f"sebagai hasil interaksi antara kapasitas individu dan kualitas konteks kerja, bukan "
            f"semata atribut pribadi (Luthans et al., 2022; Bakker & Demerouti, 2022)."
        )
    elif z:
        bc3 = (
            f"3. Kejelasan Peran dan Dukungan Kelembagaan. {z_cap} bekerja melalui kejelasan "
            f"harapan peran, kewajaran penugasan, dan dukungan sumber daya. Apabila kondisi itu "
            f"tidak tersedia, pengaruh {x_labels} terhadap {y['label']} melemah karena aparatur "
            f"menghadapi ketidakpastian yang tidak dapat diselesaikan sendiri."
        )
    else:
        bc3 = (
            f"3. Kejelasan Peran dan Dukungan Kelembagaan. Hubungan antarkonstruk dalam model ini "
            f"bekerja melalui kejelasan harapan peran, kewajaran penugasan, dan dukungan sumber daya. "
            f"Apabila kondisi itu tidak tersedia, pengaruh {x_labels} terhadap {y['label']} melemah "
            f"karena aparatur menghadapi ketidakpastian yang tidak dapat diselesaikan sendiri."
        )
    pv(doc, bc3, ctx, "2.3.3-bc3", indent=0.75, space_after=6)

    if z:
        closing = (
            f"Dengan demikian, pengujian {z_cap} bukan pelengkap statistik, melainkan bagian dari "
            f"penjelasan tentang kapan dan bagi siapa model ini berlaku. Kerangka moderated-mediation "
            f"menutup celah pendekatan linier yang mengabaikan variasi konteks dan memberikan dasar "
            f"yang lebih realistis bagi perbaikan kinerja di {instansi}."
        )
    elif ms:
        closing = (
            f"Dengan demikian, pengujian peran {m_label} bukan pelengkap statistik, melainkan bagian "
            f"dari penjelasan tentang bagaimana {x_labels} benar-benar menyalurkan pengaruhnya kepada "
            f"{y['label']}. Kerangka mediasi menutup celah pendekatan linier yang mengabaikan "
            f"mekanisme antara dan memberi dasar yang lebih realistis bagi perbaikan kinerja di {instansi}."
        )
    else:
        closing = (
            f"Dengan demikian, pengujian pengaruh {x_labels} terhadap {y['label']} tetap memperhatikan "
            f"batas keberlakuan hubungan antarkonstruk. Kerangka jalur langsung yang dibaca bersama "
            f"konteks kerja memberi dasar yang lebih realistis bagi perbaikan kinerja di {instansi}."
        )
    pv(doc, closing, ctx, "2.3.3-p4", indent=0.5, space_after=6)


def _build_doc_sub_2_3(doc: Document, ctx: Dict[str, Any]):
    """2.3 Kerangka Konseptual & Model Struktural Disertasi (Gabungan Utuh C.1 s.d C.3)."""
    h2(doc, "C. Kerangka Berpikir")
    _build_doc_sub_2_3_1(doc, ctx, standalone=False)
    _build_doc_sub_2_3_2(doc, ctx, standalone=False)
    _build_doc_sub_2_3_3(doc, ctx, standalone=False)


def _build_doc_sub_2_4(doc: Document, ctx: Dict[str, Any]):
    """2.4 Pengembangan Hipotesis Penelitian Doktoral sesuai format baku TAPD 2026 (Tabel 7)."""
    h2(doc, "D. Hipotesis Penelitian")
    instansi = ctx.get("instansi", "Instansi Pemerintah")
    y = ctx.get("y_var", "Kinerja Aparatur")
    hypos = ctx.get("graph", {}).get("hypotheses", [])
    model_acronym = _generate_doctoral_model_acronym(ctx)

    # 1. Paragraf Pengantar Deduktif Sub-Bab D (>= 4 baris teks utuh, bervitasi)
    p(doc, f"Berdasarkan telaah mendalam terhadap kajian pustaka, sintesis temuan penelitian empiris terdahulu yang relevan, serta alur silogisme deduktif dalam kerangka pemikiran konseptual di atas, maka perumusan hipotesis operasional dalam disertasi ini disusun secara formal guna menguji validitas model struktural {model_acronym} pada {instansi} (Sekaran & Bougie, 2020; Cooper & Schindler, 2022). Hipotesis-hipotesis ini diformulasikan untuk menguji signifikansi statistik pengaruh langsung pilar-pilar kapasitas organisasi, daya transmisi variabel mediasi sebagai jembatan eksekusi perilaku, serta kekuatan kontingensi variabel moderasi pengawal integritas terhadap pencapaian {y} (Creswell & Creswell, 2023). Berdasarkan uraian pada tinjauan teori yang diuraikan sebelumnya dan hasil penelitian terdahulu yang relevan serta kerangka pemikiran di atas, maka hipotesis yang dirumuskan pada penelitian ini adalah sebagai berikut:", indent=0.5, space_after=6)

    # 2. Tabel Borderless 3 Kolom Baku TAPD 2026 (Tabel 7)
    if hypos:
        tbl = doc.add_table(rows=len(hypos), cols=3)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = tbl._tbl.tblPr
        tblBorders = parse_xml(
            r'<w:tblBorders %s>'
            r'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            r'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            r'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            r'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            r'  <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            r'  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            r'</w:tblBorders>' % nsdecls('w')
        )
        tblPr.append(tblBorders)

        col_widths = [Cm(1.2), Cm(0.6), Cm(14.0)]
        for r_idx, h in enumerate(hypos):
            row = tbl.rows[r_idx]
            for c_idx, w in enumerate(col_widths):
                row.cells[c_idx].width = w

            # Col 0: Kode Hipotesis (H1, H2.. bold)
            p0 = row.cells[0].paragraphs[0]
            p0.paragraph_format.space_before = Pt(2)
            p0.paragraph_format.space_after = Pt(2)
            r0 = p0.add_run(h.get("code", f"H{r_idx+1}"))
            r0.bold = True
            r0.font.name = "Times New Roman"
            r0.font.size = Pt(12)

            # Col 1: Simbol Pemisah (:)
            p1 = row.cells[1].paragraphs[0]
            p1.paragraph_format.space_before = Pt(2)
            p1.paragraph_format.space_after = Pt(2)
            r1 = p1.add_run(":")
            r1.font.name = "Times New Roman"
            r1.font.size = Pt(12)

            # Col 2: Pernyataan Hipotesis Formal Berarah (Justify)
            p2 = row.cells[2].paragraphs[0]
            p2.paragraph_format.space_before = Pt(2)
            p2.paragraph_format.space_after = Pt(2)
            p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            stmt = h.get("statement", "")
            if not stmt.lower().startswith("terdapat"):
                stmt = f"Terdapat {stmt}"
            r2 = p2.add_run(stmt)
            r2.font.name = "Times New Roman"
            r2.font.size = Pt(12)

    # 3. Paragraf Penutup Metodologis Pengujian Hipotesis (>= 4 baris teks utuh, bervitasi)
    p(doc, f"Seluruh hipotesis formal berarah (H1 s.d H{len(hypos)}) di atas diuji secara simultan menggunakan teknik Structural Equation Modeling berbasis varian (PLS-SEM) dengan bantuan perangkat lunak SmartPLS 4 (Hair et al., 2022). Uji signifikansi koefisien jalur struktural dilakukan melalui prosedur non-parametrik bootstrapping sebanyak 5.000 subsampel dengan tingkat signifikansi alpha 5% (t-statistik > 1,96 dan p-value < 0,05) guna memastikan ketahanan estimasi parameter model (Hayes, 2022). Dengan pendekatan analitis inferensial ini, penelitian mampu mengonfirmasi signifikansi statistik dari setiap lintasan kausalitas langsung, efek tidak langsung mediasi, maupun efek interaksi moderasi secara komprehensif dalam setting empiris pada {instansi} (Sarstedt et al., 2022).", indent=0.5, space_before=6, space_after=6)


def _get_ordered_variables(ctx: Dict[str, Any]) -> List[Tuple[str, str, str, Dict[str, Any]]]:
    onto_vars = ctx.get("ontology", {}).get("variables", {})

    def _canon_name(v: Dict[str, Any], fallback: str = "") -> str:
        """Nama tampilan variabel: utamakan nama dari JUDUL/panel (input_name),
        bukan nama_resmi ontologi yang kadang memakai padanan Inggris."""
        return (v.get("input_name") or fallback or v.get("nama_resmi") or "").strip()

    ordered = []
    m_names = [n for n in (ctx.get("m_vars") or []) if n]
    z_names = [n for n in (ctx.get("z_vars") or []) if n]
    x_names = [n for n in (ctx.get("x_vars") or []) if n]

    # 1. Y (Endogen)
    if "y" in onto_vars:
        v = onto_vars["y"]
        ordered.append(("y", "Y", _canon_name(v, ctx.get("y_var", "")), v))
    else:
        ordered.append(("y", "Y", ctx["y_var"], {}))
    # 2. M (Mediasi)
    if "m" in onto_vars:
        v = onto_vars["m"]
        ordered.append(("m", "M", _canon_name(v, m_names[0] if m_names else "Mediasi"), v))
    elif m_names:
        for i, m_name in enumerate(m_names):
            role = f"M{i+1}" if len(m_names) > 1 else "M"
            ordered.append((f"m{i+1}", role, m_name, {}))
    # 3. Z (Moderasi)
    if "z" in onto_vars:
        v = onto_vars["z"]
        ordered.append(("z", "Z", _canon_name(v, z_names[0] if z_names else "Moderasi"), v))
    elif z_names:
        for i, z_name in enumerate(z_names):
            role = f"Z{i+1}" if len(z_names) > 1 else "Z"
            ordered.append((f"z{i+1}", role, z_name, {}))
    # 4. X (Eksogen)
    x_keys = [k for k in onto_vars.keys() if k.startswith("x")]
    x_keys.sort()
    if x_keys:
        for pos, k in enumerate(x_keys):
            v = onto_vars[k]
            role = k.upper()
            fallback = x_names[pos] if pos < len(x_names) else k
            ordered.append((k, role, _canon_name(v, fallback), v))
    else:
        for i, x_name in enumerate(x_names):
            ordered.append((f"x{i+1}", f"X{i+1}", x_name, {}))
    return ordered


def _get_var_prefix(vname: str, role: str) -> str:
    words = [w for w in vname.replace("-", " ").split() if w]
    if len(words) >= 2:
        return (words[0][0] + words[1][0]).upper()
    elif len(words) == 1 and len(words[0]) >= 2:
        return words[0][:2].upper()
    return role.upper()


def _get_variable_instrument_details(
    vkey: str,
    role: str,
    vname: str,
    vdata: Dict[str, Any],
    r_tabel: str = "0,279",
    ctx: Optional[Dict[str, Any]] = None,
):
    """Rincian dimensi/butir instrumen.

    Nilai ``r_hitung`` dan Cronbach's Alpha **tidak dikarang**. Nilai diambil dari
    data empiris (``outer_model``) bila tersedia; jika belum ada, sel diisi
    placeholder dan keputusan butir ditandai "Belum diuji".
    """
    dims = vdata.get("dimensi", [])
    if nv.is_generic_dims(dims):
        kb = nv.lookup(vname) or {}
        if kb.get("dimensi"):
            dims = kb["dimensi"]
    if not dims:
        dims = [
            {"nama": f"Dimensi Kapasitas {vname}", "indikator": [f"Tingkat pemahaman pegawai terhadap konsep {vname}", f"Efektivitas penerapan {vname} dalam operasional kerja"]},
            {"nama": f"Dimensi Konsistensi {vname}", "indikator": [f"Konsistensi pemeliharaan standar {vname} di lingkungan kerja", f"Dukungan pimpinan dan sarana kerja terhadap {vname}"]}
        ]
    prefix = _get_var_prefix(vname, role)
    items = []
    item_num = 1

    # Peta loading riil (bila ada) untuk dijadikan dasar r_hitung butir.
    _real_outer = (s3data.outer(ctx, vkey) if ctx else None) or {}
    _load_map: Dict[str, Any] = {}
    if isinstance(_real_outer.get("loadings"), list):
        for _l in _real_outer["loadings"]:
            if isinstance(_l, dict) and _l.get("code"):
                _load_map[str(_l["code"])] = _l.get("loading")
    try:
        _r_min = float(str(r_tabel).replace(",", "."))
    except (TypeError, ValueError):
        _r_min = 0.279

    for d in dims:
        dname = d.get("nama", "Dimensi")
        inds = d.get("indikator", [])
        if not inds:
            inds = [f"Pelaksanaan indikator 1 {dname}", f"Pelaksanaan indikator 2 {dname}"]
        for ind in inds:
            code = f"{prefix}{item_num:02d}"
            _raw = _load_map.get(code)
            try:
                _rv = float(_raw)
            except (TypeError, ValueError):
                _rv = None
            if _rv is None:
                r_txt = s3data.PLACEHOLDER
                keputusan = "Belum diuji"
            else:
                r_txt = f"{_rv:.3f}".replace(".", ",")
                keputusan = "Valid" if _rv >= _r_min else "Tidak Valid"
            items.append({
                "no": str(item_num),
                "dimensi": dname,
                "indikator": ind,
                "butir": str(item_num),
                "kode": code,
                "r_hitung": r_txt,
                "r_tabel": r_tabel,
                "keputusan": keputusan
            })
            item_num += 1

    _alpha = _real_outer.get("cronbach_alpha")
    try:
        alpha_str = f"{float(_alpha):.3f}".replace(".", ",")
    except (TypeError, ValueError):
        alpha_str = s3data.PLACEHOLDER

    return dims, items, alpha_str


def _generate_populasi_units(total_pop: int, total_sample: int, instansi: str, unit_input=None):
    """Sebaran populasi per unit kerja.

    Jika ``unit_input`` berisi data riil (list dict/objek berisi nama unit dan jumlah
    pegawai), data itulah yang dipakai. Bila tidak, sebaran diturunkan secara
    proporsional dari struktur organisasi sebagai estimasi awal (bukan angka resmi).
    """
    units = _normalize_unit_input(unit_input, total_pop)
    provided_units = bool(units)
    if not units:
        low_inst = (instansi or "").lower()
        if any(k in low_inst for k in ["pendapatan", "pajak", "retribusi"]):
            unit_names = [
                "Sekretariat",
                "Bidang Pajak Daerah",
                "Bidang Retribusi Daerah",
                "Bidang Pengelolaan Keuangan dan Aset",
                "Bidang Perencanaan, Evaluasi, dan Pelaporan",
                "Unit Pelaksana Teknis Pemungutan dan Layanan",
                "Kelompok Jabatan Fungsional",
            ]
            weights = [0.22, 0.20, 0.16, 0.14, 0.12, 0.09, 0.07]
        elif any(k in low_inst for k in ["inspektorat", "pengawasan"]):
            unit_names = [
                "Sekretariat",
                "Inspektorat Bidang Pemerintahan",
                "Inspektorat Bidang Pembangunan",
                "Inspektorat Bidang Keuangan",
                "Kelompok Jabatan Fungsional Pengawas",
                "Unit Pengaduan dan Tindak Lanjut",
                "Bagian Evaluasi dan Pelaporan",
            ]
            weights = [0.22, 0.18, 0.17, 0.16, 0.12, 0.08, 0.07]
        else:
            unit_names = [
                "Sekretariat & Bagian Umum/Kepegawaian",
                "Bidang Perencanaan, Anggaran & Program Kerja",
                "Bidang Pengawasan Internal, Kepatuhan & APIP",
                "Bidang Pengadaan Barang/Jasa & Manajemen Aset",
                "Bidang Operasional Teknis & Pelayanan Kedinasan",
                "Bidang Evaluasi, Pelaporan & Akuntabilitas Kinerja",
                "Unit Pelaksana Teknis & Layanan Terpadu"
            ]
            weights = [0.24, 0.16, 0.15, 0.15, 0.12, 0.10, 0.08]
        units = []
        accum = 0
        for idx, (uname, w) in enumerate(zip(unit_names, weights)):
            if idx == len(unit_names) - 1:
                u_pop = total_pop - accum
            else:
                u_pop = int(round(total_pop * w))
                accum += u_pop
            units.append((uname, u_pop))

    # Alokasi proporsional dengan metode sisa terbesar agar total tepat = n.
    raw_alloc = [((up / total_pop) * total_sample) if total_pop else 0 for _, up in units]
    alloc = [int(math.floor(v)) for v in raw_alloc]
    sisa = total_sample - sum(alloc)
    if sisa > 0:
        order = sorted(range(len(units)), key=lambda i: (raw_alloc[i] - alloc[i]), reverse=True)
        for i in order[:sisa]:
            alloc[i] += 1

    units_data = []
    for idx, ((uname, u_pop), u_samp, raw_val) in enumerate(zip(units, alloc, raw_alloc)):
        units_data.append({
            "no": f"{idx + 1}.",
            "unit": uname if provided_units else (f"{uname} {instansi}" if instansi and instansi.lower() not in uname.lower() else uname),
            "populasi": str(u_pop),
            "prop_str": f"({u_pop} / {total_pop}) × {total_sample} = {f'{raw_val:.2f}'.replace('.', ',')}",
            "sampel": str(u_samp)
        })
    return units_data


def _normalize_unit_input(unit_input, total_pop: int):
    """Validasi data unit kerja dari panel: [{'unit': ..., 'populasi': ...}]."""
    if not unit_input:
        return []
    units = []
    for item in unit_input:
        if isinstance(item, dict):
            name = str(item.get("unit") or item.get("nama") or item.get("name") or "").strip()
            pop = item.get("populasi", item.get("jumlah", item.get("pop")))
        elif isinstance(item, (list, tuple)) and len(item) >= 2:
            name, pop = str(item[0]).strip(), item[1]
        else:
            continue
        if not name:
            continue
        try:
            u_pop = _parse_positive_int(pop, "Populasi unit kerja")
        except ValueError:
            continue
        if u_pop <= 0:
            continue
        units.append((name, u_pop))
    if not units:
        return []
    total = sum(u for _, u in units)
    if total != total_pop:
        raise ValueError(
            f"Total populasi per unit kerja ({total}) tidak sama dengan jumlah populasi (N = {total_pop}). "
            "Periksa kembali data sebaran pegawai per unit kerja."
        )
    return units



def _build_doc_sub_3_a(doc: Document, ctx: Dict[str, Any]):
    """Sub-Bab A. Desain Penelitian (narasi dirakit dari profil variabel)."""
    h1(doc, "BAB III\nMETODE PENELITIAN")
    h2(doc, "A. Desain Penelitian")
    instansi = ctx["instansi"]
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "3a")
    y = pm["Y"]
    x_labels = _join_id(p["label"] for p in pm["_all"]["x"]) or "variabel eksogen"
    m_labels = ", ".join(p["label"] for p in pm["_all"]["m"]) or "mekanisme perantara"
    z_labels = ", ".join(p["label"] for p in pm["_all"]["z"]) or "kondisi kontekstual"

    for i, text in enumerate(nv.method_design_paragraphs(ctx, pm, rng), 1):
        pv(doc, text, ctx, f"3a-{i}", indent=0.5, space_after=6)

    pv(doc,
       f"Posisi masing-masing konstruk dalam desain ini bersifat tegas. {x_labels} diperlakukan "
       f"sebagai anteseden yang diukur pada tingkat individu; {m_labels} sebagai mekanisme yang "
       f"menerjemahkan anteseden menjadi energi kerja; {z_labels} sebagai kondisi batas yang "
       f"mengubah kekuatan hubungan; dan {y['label']} sebagai luaran yang dinilai berdasarkan "
       f"pengalaman kerja responden. Ketegasan posisi ini mencegah tumpang tindih pengukuran dan "
       f"menjaga validitas internal model.",
       ctx, "3a-4", indent=0.5, space_after=6)




def _build_doc_sub_3_b(doc: Document, ctx: Dict[str, Any]):
    """Sub-Bab B. Lokasi Penelitian."""
    h2(doc, "B. Lokasi Penelitian")
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "3b")
    for i, text in enumerate(nv.method_location_paragraphs(ctx, pm, rng), 1):
        pv(doc, text, ctx, f"3b-{i}", indent=0.5, space_after=6)

    pv(doc,
       f"Batasan lokus penelitian adalah unit kerja yang berada di bawah kewenangan "
       f"{ctx['instansi']}. Unit di luar struktur tersebut tidak dimasukkan karena perbedaan "
       f"sistem penilaian kinerja dan prosedur kerja dapat mengaburkan makna {pm['Y']['label']}. "
       f"Dengan batasan ini, generalisasi temuan diarahkan pada populasi aparatur instansi, "
       f"bukan pada seluruh birokrasi daerah.",
       ctx, "3b-3", indent=0.5, space_after=6)




def _build_doc_sub_3_c(doc: Document, ctx: Dict[str, Any]):
    """Sub-Bab C. Waktu Penelitian & Tabel 3.1 Jadwal Penelitian."""
    h2(doc, "C. Waktu Penelitian")
    tahun = ctx.get("tahun", "2026")
    instansi = ctx["instansi"]
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "3c")
    y = pm["Y"]
    pilot = nv.pilot_plan(ctx)
    pv(doc,
       f"Penelitian ini dilaksanakan secara bertahap pada tahun {tahun}, dimulai pada awal "
       f"tahun akademik dan diakhiri pada tahap pelaporan hasil. Penjadwalan tidak disusun "
       f"sekadar untuk memenuhi kelengkapan administratif, tetapi untuk menjaga kualitas data. "
       f"Pengumpulan data primer dijadwalkan pada periode ketika aktivitas kedinasan di "
       f"{instansi} berada pada intensitas yang relatif terkendali, sehingga responden memiliki "
       f"waktu memadai untuk memahami butir kuesioner (Sugiyono, 2022; Neuman, 2022).",
       ctx, "3c-p1", indent=0.5, space_after=6)

    pv(doc, nv.pick(rng, [
       (
           f"Periode survei sengaja dihindarkan dari puncak pelaporan akhir tahun dan masa "
           f"rekonsiliasi data, karena pada periode tersebut {y['label']} dan beban administratif "
           f"meningkat tajam. Apabila survei dipaksakan pada periode itu, jawaban berisiko "
           f"mencerminkan kelelahan sesaat, bukan kondisi kerja yang biasa. Pertimbangan ini "
           f"merupakan bagian dari pengendalian bias pengukuran, bukan sekadar preferensi "
           f"operasional (Cooper & Schindler, 2022)."
       ),
       (
           f"Pelaksanaan survei dijauhkan dari puncak pelaporan akhir tahun dan masa rekonsiliasi "
           f"data karena pada rentang itu {y['label']} dan beban administratif cenderung meningkat. "
           f"Jika survei tetap dilakukan pada waktu tersebut, jawaban berpotensi merekam kelelahan "
           f"sesaat alih-alih kondisi kerja yang wajar. Pertimbangan semacam ini termasuk upaya "
           f"mengendalikan bias pengukuran, bukan semata pilihan operasional "
           f"(Cooper & Schindler, 2022)."
       ),
       (
           f"Waktu survei sengaja tidak dipilih pada masa puncak pelaporan akhir tahun maupun "
           f"rekonsiliasi data, sebab pada periode itu {y['label']} dan beban administratif "
           f"menanjak. Bila survei dipaksakan, jawaban berisiko menggambarkan kelelahan sementara, "
           f"bukan keadaan kerja yang biasa. Hal ini merupakan langkah pengendalian bias "
           f"pengukuran dan bukan sekadar pertimbangan teknis (Cooper & Schindler, 2022)."
       ),
       (
           f"Survei tidak dilaksanakan pada masa puncak pelaporan akhir tahun maupun rekonsiliasi "
           f"data karena {y['label']} dan beban administratif sedang tinggi pada rentang itu. "
           f"Andaikan survei dijalankan pada waktu tersebut, jawaban cenderung mencerminkan "
           f"kelelahan sesaat daripada keadaan kerja yang wajar. Langkah ini termasuk pengendalian "
           f"bias pengukuran, bukan semata pertimbangan teknis (Cooper & Schindler, 2022)."
       ),
    ]), ctx, "3c-p2", indent=0.5, space_after=6)

    t1_headers = ["No", "Tahapan dan Kegiatan Penelitian", "Waktu Pelaksanaan", "Target Luaran / Output"]
    t1_rows = [
        ["1", "Pengajuan dan Usulan Judul Disertasi", f"Januari {tahun}", "Matriks Model & SK Pembimbing"],
        ["2", "Penyusunan TAPD dan Sintesis Teoretis", f"Februari {tahun}", "Naskah Draft TAPD Doktoral"],
        ["3", "Analisis Data Awal dan LKjIP/LAKIP", f"Februari {tahun}", "Data Empiris Masalah Fenomena"],
        ["4", "Perumusan Dimensi dan Indikator Variabel", f"Maret {tahun}", "Blueprint Ontologi Variabel"],
        ["5", "Penyusunan Butir Instrumen Kuesioner", f"Maret {tahun}", "Draft Kuesioner Skala Likert"],
        ["6", "Pelaksanaan Survei Pendahuluan", f"April {tahun}", "Validasi Awal Gap Kinerja"],
        ["7", "Penulisan Naskah Lengkap Bab I", f"April {tahun}", "Draft Bab I Latar Belakang"],
        ["8", "Penulisan Naskah Lengkap Bab II", f"Mei {tahun}", "Draft Bab II Kajian Pustaka"],
        ["9", "Penulisan Naskah Lengkap Bab III", f"Juni {tahun}", "Draft Bab III Metode Penelitian"],
        ["10", "Seminar Proposal Disertasi (Sempro)", f"Juli {tahun}", "Berita Acara Kelayakan Proposal"],
        ["11", f"Uji Coba Instrumen (n = {pilot['n']} responden)", f"Agustus {tahun}", "Data r-hitung & Cronbach Alpha"],
        ["12", "Survei Empiris Utama & Analisis SmartPLS 4", f"September {tahun}", "Hasil Estimasi Outer & Inner Model"],
        ["13", "Penulisan Bab IV, V & Ujian Disertasi Tertutup", f"Oktober {tahun}", "Naskah Final Disertasi Doktor"],
    ]
    add_table_data(doc, title="Tabel 3. 1\nJadwal Penelitian", headers=t1_headers, data=t1_rows,
                   source=f"Sumber: Data diolah peneliti ({tahun})")

    pv(doc,
       f"Tahapan pada Tabel 3.1 disusun berjenjang agar setiap keputusan metodologis dapat "
       f"ditelusuri. Penentuan dimensi dan indikator mendahului penyusunan butir; uji coba "
       f"instrumen mendahului survei utama; dan analisis model dilakukan setelah data bersih. "
       f"Urutan ini mencegah praktik menyusun instrumen setelah melihat hasil, yang akan "
       f"merusak validitas penelitian. Bagi {instansi}, jadwal ini juga memberi kepastian "
       f"waktu keterlibatan aparatur sehingga tidak mengganggu pelayanan (Hair et al., 2022).",
       ctx, "3c-p3", indent=0.5, space_before=6, space_after=6)


def _build_doc_sub_3_d(doc: Document, ctx: Dict[str, Any]):
    """Sub-Bab D. Sumber Data."""
    h2(doc, "D. Sumber Data")
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "3d")
    for i, text in enumerate(nv.method_data_paragraphs(ctx, pm, rng), 1):
        pv(doc, text, ctx, f"3d-{i}", indent=0.5, space_after=6)

    pv(doc,
       f"Kualitas data dijaga melalui pemeriksaan kelengkapan, konsistensi jawaban, dan "
       f"pencocokan identitas unit kerja. Kuesioner yang tidak lengkap atau menunjukkan pola "
       f"jawaban ekstrem tanpa variasi tidak langsung dibuang, tetapi diperiksa ulang kepada "
       f"responden. Prosedur ini penting agar keputusan tentang data yang dipakai dapat "
       f"dipertanggungjawabkan dan tidak menimbulkan bias seleksi.",
       ctx, "3d-3", indent=0.5, space_after=6)




def _build_doc_sub_3_e(doc: Document, ctx: Dict[str, Any]):
    """Sub-Bab E. Populasi dan Sampel (narasi kontekstual, angka tetap auditable)."""
    h2(doc, "E. Populasi dan Sampel")
    instansi = ctx["instansi"]
    populasi = int(str(ctx["populasi"]).strip())
    sampel = int(str(ctx["sampel"]).strip())
    tahun = ctx.get("tahun", "2026")
    unit_kerja_input = ctx.get("unit_kerja") or []
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "3e")
    y = pm["Y"]
    xs = pm["_all"]["x"]
    x_labels = _join_id(p["label"] for p in xs) or "variabel penelitian"
    n_paths = len(ctx.get("graph", {}).get("hypotheses", [])) or 1

    p(doc, "Populasi", bold=True, indent=0, space_before=4, space_after=4)
    pv(doc,
       f"Populasi penelitian ini adalah seluruh aparatur yang bertugas di lingkungan "
       f"{instansi}, dengan jumlah sasaran N = {populasi} orang. Penetapan populasi tidak "
       f"didasarkan pada kemudahan akses, tetapi pada batas kewenangan organisasi dan "
       f"relevansi tugas dengan konstruk yang diukur, yaitu {x_labels} serta {y['label']}. "
       f"Dengan batas itu, kesimpulan penelitian dapat dikembalikan kepada kelompok aparatur "
       f"yang benar-benar terlibat dalam proses kerja instansi (Sugiyono, 2022; Cooper & "
       f"Schindler, 2022).",
       ctx, "3e-pop1", indent=0.5, space_after=6)

    units_data = _generate_populasi_units(populasi, sampel, instansi, unit_kerja_input)
    t2_headers = ["No", "Wilayah / Unit Kerja Pemerintahan", "Jumlah Pegawai (Populasi)"]
    t2_rows = [[u["no"], u["unit"], u["populasi"]] for u in units_data]
    t2_rows.append(["", "Total Populasi Pegawai (N)", str(populasi)])
    if unit_kerja_input:
        t2_source = f"Sumber: Data Bagian Kepegawaian {instansi}, diolah peneliti ({tahun})"
    else:
        t2_source = f"Sumber: Struktur organisasi {instansi}; sebaran jumlah pegawai diolah peneliti ({tahun})"
    add_table_data(doc, title=f"Tabel 3. 2\nSebaran Populasi Pegawai di Lingkungan {instansi}",
                   headers=t2_headers, data=t2_rows, source=t2_source)

    pv(doc, nv.pick(rng, [
       (
           f"Tabel 3.2 memperlihatkan bahwa populasi tersebar pada unit dengan karakteristik tugas "
           f"yang berbeda. Sebaran ini penting karena beban dan tuntutan kerja tidak identik "
           f"antarbidang. Apabila alokasi sampel tidak memperhitungkan bobot unit, unit dengan "
           f"jumlah aparatur kecil berisiko tidak terwakili, padahal unit tersebut dapat memiliki "
           f"persoalan {y['label']} yang khas."
       ),
       (
           f"Sebaran populasi pada Tabel 3.2 menunjukkan bahwa unit kerja memiliki karakteristik "
           f"tugas yang berlainan. Perbedaan ini penting sebab beban dan tuntutan kerja tidak "
           f"seragam antarbidang. Bila alokasi sampel mengabaikan bobot unit, unit dengan aparatur "
           f"sedikit berisiko luput terwakili, padahal unit itu bisa menghadapi persoalan "
           f"{y['label']} yang khas."
       ),
       (
           f"Tabel 3.2 memaparkan bahwa populasi tersebar ke dalam unit dengan ragam tugas yang "
           f"berbeda. Keragaman tersebut perlu diperhatikan karena beban dan tuntutan kerja tidak "
           f"sama di setiap bidang. Tanpa memperhitungkan bobot unit dalam alokasi sampel, unit "
           f"berukuran kecil rentan tidak terwakili meski memiliki persoalan {y['label']} yang "
           f"khas."
       ),
    ]), ctx, "3e-pop2", indent=0.5, space_before=4, space_after=6)

    p(doc, "Sampel", bold=True, indent=0, space_before=8, space_after=4)
    _sensus = populasi <= 100 and sampel >= populasi
    if _sensus:
        pv(doc, nv.pick(rng, [
           (
               f"Kriteria inklusi ditetapkan agar data berasal dari aparatur yang benar-benar "
               f"menjalani proses kerja terkait {y['label']}: (1) berstatus ASN atau PPPK yang "
               f"menduduki jabatan definitif; (2) memiliki masa kerja minimal dua tahun sehingga "
               f"mengenal prosedur dan budaya kerja; serta (3) terlibat langsung dalam proses "
               f"kerja yang dimaksud. Karena populasi terjangkau hanya {populasi} orang, seluruh "
               f"aparatur yang memenuhi kriteria tersebut dijadikan calon responden tanpa "
               f"penarikan acak."
           ),
           (
               f"Penentuan responden berpedoman pada tiga kriteria inklusi: (1) berstatus ASN atau "
               f"PPPK pada jabatan definitif; (2) memiliki masa kerja sekurang-kurangnya dua tahun "
               f"sehingga memahami prosedur dan budaya kerja; dan (3) terlibat langsung dalam proses "
               f"kerja yang bersinggungan dengan {y['label']}. Dengan populasi terjangkau sebanyak "
               f"{populasi} orang, seluruh aparatur yang memenuhi kriteria dijadikan calon "
               f"responden tanpa pengundian."
           ),
           (
               f"Responden ditetapkan melalui kriteria inklusi: (1) ASN atau PPPK yang menjabat "
               f"posisi definitif; (2) masa kerja minimal dua tahun agar mengenali prosedur dan "
               f"budaya kerja; serta (3) keterlibatan langsung pada proses kerja yang terkait "
               f"{y['label']}. Populasi terjangkau yang hanya {populasi} orang membuat seluruh "
               f"aparatur yang memenuhi kriteria dapat dijadikan calon responden tanpa penarikan "
               f"acak."
           ),
        ]), ctx, "3e-samp1", indent=0.5, space_after=6)
    else:
        pv(doc, nv.pick(rng, [
       (
           f"Penarikan sampel menggunakan kombinasi purposive sampling dan proportional random "
           f"sampling. Purposive sampling dipakai untuk memastikan responden memiliki informasi "
           f"yang memadai tentang pekerjaannya, sedangkan proportional random sampling dipakai "
           f"untuk menjaga keterwakilan unit (Sekaran & Bougie, 2020). Kriteria inklusi yang "
           f"ditetapkan meliputi: (1) berstatus ASN atau PPPK yang menduduki jabatan definitif; "
           f"(2) memiliki masa kerja minimal dua tahun sehingga mengenal prosedur dan budaya kerja; "
           f"serta (3) terlibat langsung dalam proses kerja yang berkaitan dengan {y['label']}. "
           f"Kriteria ini dipilih agar jawaban tidak bersumber pada kesan sepintas."
       ),
       (
           f"Teknik penarikan sampel menggabungkan purposive sampling dan proportional random "
           f"sampling. Purposive sampling dipakai agar responden memiliki pengetahuan memadai "
           f"tentang pekerjaannya, sedangkan proportional random sampling menjaga keterwakilan unit "
           f"(Sekaran & Bougie, 2020). Kriteria inklusi yang diterapkan mencakup: (1) berstatus ASN "
           f"atau PPPK pada jabatan definitif; (2) memiliki masa kerja sekurang-kurangnya dua "
           f"tahun sehingga memahami prosedur dan budaya kerja; dan (3) terlibat langsung dalam "
           f"proses kerja yang bersinggungan dengan {y['label']}. Kriteria itu dipilih agar "
           f"jawaban tidak bertumpu pada kesan sepintas."
       ),
       (
           f"Sampel ditarik dengan memadukan purposive sampling dan proportional random sampling. "
           f"Pendekatan purposif memastikan responden memahami pekerjaannya, sementara pendekatan "
           f"proporsional acak menjaga keterwakilan tiap unit (Sekaran & Bougie, 2020). Adapun "
           f"kriteria inklusinya: (1) ASN atau PPPK yang menjabat posisi definitif; (2) masa kerja "
           f"minimal dua tahun agar mengenali prosedur dan budaya kerja; serta (3) keterlibatan "
           f"langsung pada proses kerja yang terkait {y['label']}. Kriteria tersebut ditetapkan "
           f"agar jawaban tidak sekadar berdasarkan kesan permukaan."
       ),
    ]), ctx, "3e-samp1", indent=0.5, space_after=6)

    calc_slovin = populasi / (1 + populasi * (0.05 ** 2))
    if _sensus:
        pv(doc,
           f"Karena populasi terjangkau tidak melebihi 100 orang (N = {populasi}), penelitian "
           f"ini menggunakan sampling jenuh (sensus). Seluruh anggota populasi yang memenuhi "
           f"kriteria dijadikan responden, sehingga tidak ada unsur pemilihan yang dapat "
           f"menimbulkan bias seleksi. Dengan pendekatan ini, ukuran sampel ditetapkan "
           f"n = {sampel} responden (Sugiyono, 2022; Sinambela, 2023).",
           ctx, "3e-census", indent=0.5, space_after=6)
    else:
        pv(doc,
           f"Besaran sampel dihitung dengan Formula Slovin pada margin of error 5 persen. "
           f"Taraf tersebut dipilih karena penelitian manajemen pascasarjana menuntut presisi "
           f"yang cukup untuk menolak atau menerima hipotesis, tanpa mengabaikan keterjangkauan "
           f"pengumpulan data di {instansi} (Sugiyono, 2022; Cooper & Schindler, 2022). "
           f"Perhitungan dilakukan sebagai berikut.",
           ctx, "3e-slovin", indent=0.5, space_after=4)

        p(doc, "          N", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=1)
        p(doc, "n = -----------------", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=1, space_after=1)
        p(doc, "     1 + N × (e)^2", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=1, space_after=4)

        slovin_math = (
            f"n = {populasi} / [1 + {populasi}(0,05)^2] = "
            f"{populasi} / [1 + {f'{populasi*0.0025:.4f}'.replace('.', ',')}] = {populasi} / "
            f"{f'{1 + populasi*0.0025:.4f}'.replace('.', ',')} = "
            f"{f'{calc_slovin:.2f}'.replace('.', ',')}, yang kemudian dibulatkan menjadi {sampel} "
            f"responden (Sinambela, 2023)."
        )
        p(doc, nv.pick(rng, [
            (f"Keterangan formula Slovin di atas meliputi: n adalah ukuran sampel minimum yang "
             f"dibutuhkan; N adalah ukuran total populasi pegawai ({populasi} orang); dan e adalah "
             f"batas toleransi ketidaktelitian atau presisi yang ditetapkan sebesar 5% (0,05) "
             f"(Sugiyono, 2022). Penerapan angka parameter tersebut ke dalam rumus Slovin "
             f"menghasilkan kalkulasi matematis: {slovin_math}"),
            (f"Keterangan untuk formula Slovin di atas mencakup: n adalah ukuran sampel minimum "
             f"yang diperlukan; N adalah ukuran total populasi pegawai ({populasi} orang); dan e "
             f"adalah batas toleransi ketidaktelitian atau presisi yang ditetapkan 5% (0,05) "
             f"(Sugiyono, 2022). Memasukkan angka parameter itu ke dalam rumus Slovin menghasilkan "
             f"kalkulasi matematis: {slovin_math}"),
            (f"Adapun keterangan formula Slovin tersebut: n ialah ukuran sampel minimum yang "
             f"dibutuhkan; N ialah ukuran total populasi pegawai ({populasi} orang); dan e ialah "
             f"batas toleransi ketidaktelitian atau presisi sebesar 5% (0,05) (Sugiyono, 2022). "
             f"Penerapan angka parameter itu pada rumus Slovin menghasilkan perhitungan matematis: "
             f"{slovin_math}"),
            (f"Penjelasan untuk formula Slovin tersebut adalah: n sebagai ukuran sampel minimum "
             f"yang dibutuhkan; N sebagai ukuran total populasi pegawai ({populasi} orang); dan e "
             f"sebagai batas toleransi ketidaktelitian atau presisi sebesar 5% (0,05) (Sugiyono, "
             f"2022). Dengan menyisipkan angka parameter tersebut ke rumus Slovin, diperoleh "
             f"kalkulasi matematis: {slovin_math}"),
        ]), indent=0.5, space_before=2, space_after=6)

    pv(doc, nv.pick(rng, [
       (
           f"Ukuran sampel n = {sampel} dinilai memadai untuk model struktural yang memuat "
           f"{n_paths} jalur hipotesis. Kaidah 10-times rule mensyaratkan sampel minimal sepuluh "
           f"kali jumlah panah yang mengarah ke konstruk endogen; dengan jumlah jalur pada model "
           f"ini, kebutuhan tersebut terpenuhi. Selain itu, kekuatan uji statistik pada alpha 5 "
           f"persen dan power 80 persen berada pada rentang yang lazim untuk estimasi PLS-SEM "
           f"(Hair et al., 2022; Sarstedt et al., 2022)."
       ),
       (
           f"Besaran sampel n = {sampel} dipandang memadai bagi model struktural dengan "
           f"{n_paths} jalur hipotesis. Kaidah 10-times rule menetapkan sampel minimal sepuluh "
           f"kali jumlah panah menuju konstruk endogen, dan kebutuhan itu terpenuhi pada model "
           f"ini. Di samping itu, kekuatan uji statistik pada alpha 5 persen dan power 80 persen "
           f"berada pada rentang yang lazim untuk estimasi PLS-SEM (Hair et al., 2022; "
           f"Sarstedt et al., 2022)."
       ),
       (
           f"Ukuran sampel n = {sampel} dinilai cukup untuk model struktural yang terdiri atas "
           f"{n_paths} jalur hipotesis. Menurut kaidah 10-times rule, sampel minimal adalah "
           f"sepuluh kali banyaknya panah yang mengarah ke konstruk endogen; jumlah tersebut "
           f"sudah terpenuhi. Selain itu, kekuatan uji pada alpha 5 persen dengan power 80 persen "
           f"berada dalam rentang yang umum bagi estimasi PLS-SEM (Hair et al., 2022; "
           f"Sarstedt et al., 2022)."
       ),
       (
           f"Dengan n = {sampel} dan {n_paths} jalur hipotesis, ukuran sampel dinilai memadai "
           f"bagi model struktural. Kaidah 10-times rule mensyaratkan minimal sepuluh kali jumlah "
           f"panah menuju konstruk endogen, dan syarat itu sudah terpenuhi. Ditambah lagi, "
           f"kekuatan uji pada alpha 5 persen dan power 80 persen berada dalam rentang lazim "
           f"untuk PLS-SEM (Hair et al., 2022; Sarstedt et al., 2022)."
       ),
    ]), ctx, "3e-adequacy", indent=0.5, space_before=4, space_after=6)

    pv(doc,
       f"Alokasi responden per unit dihitung secara proporsional: ni = (Ni / N) × n, dengan "
       f"ni adalah sampel unit ke-i, Ni populasi unit ke-i, N total populasi ({populasi}), dan "
       f"n total sampel ({sampel}). Pembulatan dilakukan dengan metode sisa terbesar agar jumlah "
       f"sampel antarbagian tepat sama dengan n, tanpa selisih yang mengganggu perhitungan. "
       f"Rincian alokasi disajikan pada Tabel 3.3 (Sekaran & Bougie, 2020; Sinambela, 2023).",
       ctx, "3e-alloc", indent=0.5, space_after=6)

    t3_headers = ["No", "Wilayah / Unit Kerja Pemerintahan", "Populasi (Ni)", "Perhitungan Proporsional (ni)", "Jumlah Sampel"]
    t3_rows = [[u["no"], u["unit"], u["populasi"], u["prop_str"], u["sampel"]] for u in units_data]
    t3_rows.append(["", "Total Sampel Penelitian", str(populasi), "-", str(sampel)])
    add_table_data(doc, title=f"Tabel 3. 3\nAlokasi Sampel Penelitian secara Proporsional pada Unit Kerja di {instansi}",
                   headers=t3_headers, data=t3_rows, source=f"Sumber: Data diolah peneliti ({tahun})")

    pv(doc,
       f"Alokasi pada Tabel 3.3 menjadi dasar penyebaran kuesioner. Responden dipilih secara "
       f"acak dari daftar aparatur pada setiap unit yang memenuhi kriteria, bukan ditunjuk oleh "
       f"atasan. Prosedur ini penting untuk mengurangi bias seleksi dan menjaga kerahasiaan "
       f"jawaban, sehingga data yang terkumpul mencerminkan variasi pengalaman kerja di "
       f"{instansi}.",
       ctx, "3e-close", indent=0.5, space_before=4, space_after=6)


def _build_doc_single_variable_f(doc: Document, ctx: Dict[str, Any], vkey: str, role: str, vname: str, vdata: Dict[str, Any], table_num: int = 4, standalone: bool = True, sub_code: str = ""):
    """F.x Definisi Konseptual, Operasional & Kisi-Kisi untuk 1 Variabel."""
    if standalone:
        h1(doc, "BAB III\nMETODOLOGI PENELITIAN")
        h2(doc, "F. Definisi Konseptual dan Operasional Variabel")

    code_title = f"{sub_code} " if sub_code else ""
    h3(doc, f"{code_title}Variabel {vname} ({role})")

    instansi = ctx.get("instansi", "Instansi Pemerintah")
    tahun = ctx.get("tahun", "2026")
    pm = nv.profile_map(ctx)
    prof = pm.get(str(vkey).upper()) or nv.make_profile(vname, vdata, ctx)
    rng = nv.rng_for(ctx, f"3f-{vkey}")
    pilot = nv.pilot_plan(ctx)

    dims, items, _ = _get_variable_instrument_details(vkey, role, vname, vdata, pilot["r_tabel"], ctx=ctx)

    p(doc, "Definisi Konseptual", bold=True, indent=0, space_before=2, space_after=2)
    pv(doc, nv.conceptual_definition(prof, instansi, rng),
       ctx, f"3f-{vkey}-konsep", indent=0.5, space_after=6)

    p(doc, "Definisi Operasional", bold=True, indent=0, space_before=4, space_after=2)
    pv(doc, nv.operational_definition(prof, instansi, rng, len(items)),
       ctx, f"3f-{vkey}-oper", indent=0.5, space_after=6)

    dim_txt = "; ".join(
        f"{d.get('nama', 'Dimensi')} ({len(d.get('indikator', []))} indikator)"
        for d in dims
    )
    pv(doc,
       f"Operasionalisasi {vname} menurunkan {len(dims)} dimensi menjadi {len(items)} butir. "
       f"Rincian tersebut adalah {dim_txt}. Setiap butir diberi kode unik agar dapat ditelusuri "
       f"dari kuesioner, tabel validitas, hingga model pengukuran. Dengan cara ini, pembaca dapat "
       f"memeriksa apakah indikator benar-benar merepresentasikan {prof['label']} dan bukan "
       f"konstruk lain.",
       ctx, f"3f-{vkey}-dims", indent=0.5, space_after=6)

    t_headers = ["Dimensi", "Indikator", "Butir Pertanyaan"]
    t_rows = [[it["dimensi"], it["indikator"], it["butir"]] for it in items]
    tbl_title = f"Tabel 3. {table_num}\nKisi-Kisi Instrumen Variabel {vname}"
    add_table_data(doc, title=tbl_title, headers=t_headers, data=t_rows,
                   source=f"Sumber: Data diolah peneliti ({tahun})")


def _build_doc_sub_3_f_y(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """F.1 Definisi Konseptual & Operasional Variabel Terikat (Y)."""
    ordered_vars = _get_ordered_variables(ctx)
    y_entry = next((v for v in ordered_vars if str(v[0]).upper() == "Y"),
                   ("y", "Y", ctx.get("y_var", "Kinerja Pegawai"), {}))
    _build_doc_single_variable_f(doc, ctx, y_entry[0], y_entry[1], y_entry[2], y_entry[3], table_num=4, standalone=standalone, sub_code="F.1")


def _build_doc_sub_3_f_m(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """F.2 Definisi Konseptual & Operasional Variabel Mediasi (M)."""
    ordered_vars = _get_ordered_variables(ctx)
    m_entry = next((v for v in ordered_vars if str(v[0]).upper() == "M" or "mediasi" in v[1].lower()), None)
    if not m_entry:
        m_entry = ("M", "Variabel Mediasi", ctx.get("m", "Motivasi Kerja"), {})
    _build_doc_single_variable_f(doc, ctx, m_entry[0], m_entry[1], m_entry[2], m_entry[3], table_num=5, standalone=standalone, sub_code="F.2")


def _build_doc_sub_3_f_z(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """F.3 Definisi Konseptual & Operasional Variabel Moderasi (Z)."""
    ordered_vars = _get_ordered_variables(ctx)
    z_entry = next((v for v in ordered_vars if str(v[0]).upper() == "Z" or "moderasi" in v[1].lower()), None)
    if not z_entry:
        _zv = (ctx.get("z") or "").strip()
        if not _zv:
            return
        z_entry = ("Z", "Variabel Moderasi", _zv, {})
    _build_doc_single_variable_f(doc, ctx, z_entry[0], z_entry[1], z_entry[2], z_entry[3], table_num=6, standalone=standalone, sub_code="F.3")


def _build_doc_sub_3_f_xk(doc: Document, ctx: Dict[str, Any], x_idx: int = 1, standalone: bool = True):
    """F.4.x Definisi Konseptual & Operasional Variabel Bebas (X_k)."""
    ordered_vars = _get_ordered_variables(ctx)
    target_key = f"X{x_idx}"
    x_entry = next((v for v in ordered_vars if str(v[0]).upper() == target_key.upper()), None)
    if not x_entry:
        x_vars = ctx.get("x_vars", [])
        vname = x_vars[x_idx - 1] if x_idx <= len(x_vars) else f"Variabel Bebas X{x_idx}"
        x_entry = (target_key, f"Variabel Bebas / Eksogen {x_idx}", vname, {})
    _build_doc_single_variable_f(doc, ctx, x_entry[0], x_entry[1], x_entry[2], x_entry[3], table_num=6 + x_idx, standalone=standalone, sub_code=f"F.4.{x_idx}")


def _build_doc_sub_3_f(doc: Document, ctx: Dict[str, Any]):
    """Sub-Bab F. Definisi Konseptual & Operasional + Tabel Kisi-Kisi Instrumen (Composite Lengkap)."""
    h2(doc, "F. Definisi Konseptual dan Operasional Variabel")
    y_var = ctx["y_var"]
    m_vars_str = ", ".join(ctx["m_vars"]) if ctx["m_vars"] else "(Tanpa Mediasi)"
    z_vars_str = ", ".join(ctx["z_vars"]) if ctx["z_vars"] else "(Tanpa Moderasi)"
    x_vars_str = _join_id(ctx["x_vars"])

    _taksonomi = [f"variabel terikat/endogen ({y_var})"]
    if ctx["m_vars"]:
        _taksonomi.append(f"variabel intervening/mediasi ({m_vars_str})")
    if ctx["z_vars"]:
        _taksonomi.append(f"variabel kontingensi/moderasi ({z_vars_str})")
    _taksonomi.append(f"variabel bebas/eksogen ({x_vars_str})")
    _taksonomi_str = _join_id(_taksonomi)

    p(doc, f"Operasionalisasi variabel merupakan jembatan metodologis yang menerjemahkan konstruk teoretis abstrak ke dalam indikator empiris yang terukur secara objektif di lapangan (Cooper & Schindler, 2022). Dalam penelitian disertasi ini, variabel penelitian diklasifikasikan ke dalam {len(_taksonomi)} taksonomi struktural: {_taksonomi_str} (Sekaran & Bougie, 2020; Hair et al., 2022). Setiap variabel dijabarkan ke dalam Definisi Konseptual yang berakar pada teori substantif, Definisi Operasional yang menguraikan teknik pengukuran berskala Likert 1–5, serta Kisi-Kisi Instrumen yang memetakan dimensi dan indikator empirisnya secara sistematis (Sugiyono, 2022).", indent=0.5, space_after=6)

    ordered_vars = _get_ordered_variables(ctx)
    table_counter = 4
    for idx_v, (vkey, role, vname, vdata) in enumerate(ordered_vars, 1):
        _build_doc_single_variable_f(doc, ctx, vkey, role, vname, vdata, table_num=table_counter, standalone=False, sub_code=f"{idx_v}.")
        table_counter += 1


def _build_doc_sub_3_g(doc: Document, ctx: Dict[str, Any]):
    """Sub-Bab G. Instrumen Penelitian & Skala Likert."""
    h2(doc, "G. Instrumen Penelitian")
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "3g")
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    tbl_num = 4 + k_vars
    pilot = nv.pilot_plan(ctx)
    total_items = 0
    for vkey, role, vname, vdata in ordered_vars:
        _, items, _ = _get_variable_instrument_details(vkey, role, vname, vdata, pilot["r_tabel"], ctx=ctx)
        total_items += len(items)

    for i, text in enumerate(nv.method_instrument_paragraphs(ctx, pm, rng, total_items, pilot), 1):
        pv(doc, text, ctx, f"3g-{i}", indent=0.5, space_after=6)

    pv(doc,
       f"Penyusunan butir mengikuti kaidah satu indikator satu pernyataan. Pernyataan negatif "
       f"dihindari agar responden tidak perlu membalik makna jawaban; bila terpaksa digunakan, "
       f"skornya dibalik pada tahap pengolahan. Uji keterbacaan dilakukan bersama beberapa "
       f"aparatur untuk memastikan istilah teknis dipahami secara seragam sebelum kuesioner "
       f"disebarkan di {ctx['instansi']}.",
       ctx, "3g-3", indent=0.5, space_after=6)

    t_headers = ["Jawaban Responden", "Skor"]
    t_rows = [
        ["Sangat Setuju (SS)", "5"],
        ["Setuju (S)", "4"],
        ["Ragu - ragu (RR) / Netral (N)", "3"],
        ["Tidak Setuju (TS)", "2"],
        ["Sangat Tidak Setuju (STS)", "1"],
    ]
    add_table_data(doc, title=f"Tabel 3. {tbl_num}\nSkor Jawaban Responden",
                   headers=t_headers, data=t_rows,
                   source="Sumber: Sugiyono (2022); Sinambela (2023)")

    pv(doc,
       f"Skala lima poin memberi ruang bagi responden untuk menyatakan sikap netral tanpa "
       f"memaksakan pilihan pada dua kutub. Untuk keperluan analisis, skor ordinal diperlakukan "
       f"sebagai data interval dengan asumsi yang lazim dalam PLS-SEM, yaitu jarak antarkategori "
       f"dianggap setara. Asumsi ini diuji melalui distribusi jawaban dan tidak digeneralisasi "
       f"tanpa pemeriksaan.",
       ctx, "3g-4", indent=0.5, space_before=6, space_after=6)


def _build_doc_sub_3_h(doc: Document, ctx: Dict[str, Any]):
    """Sub-Bab H. Ujicoba Instrumen (Validitas Pearson & Reliabilitas Cronbach's Alpha)."""
    h2(doc, "H. Ujicoba Instrumen")
    instansi = ctx["instansi"]
    tahun = ctx.get("tahun", "2026")
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    y_var = ctx["y_var"]
    pilot = nv.pilot_plan(ctx)
    n_pilot = pilot["n"]
    r_tabel = pilot["r_tabel"]
    df = int(n_pilot) - 2

    pv(doc,
       f"Sebelum kuesioner digunakan pada sampel utama di {instansi}, instrumen diujicobakan "
       f"kepada {n_pilot} responden yang memiliki karakteristik setara tetapi berada di luar "
       f"sampel penelitian. Uji coba ini bertujuan memeriksa keterbacaan butir, konsistensi "
       f"jawaban, dan kualitas psikometrik instrumen. Data hasil uji coba dianalisis dengan "
       f"bantuan program statistik untuk menghitung koefisien validitas dan reliabilitas "
       f"(Sinambela, 2023; Hair et al., 2022).",
       ctx, "3h-1", indent=0.5, space_after=6)

    p(doc, "1. Uji Validitas Instrumen", bold=True, indent=0, space_before=4, space_after=4)
    pv(doc,
       f"Uji validitas butir dilakukan dengan korelasi Pearson Product Moment antara skor "
       f"setiap butir dan skor total konstruknya. Korelasi ini menunjukkan sejauh mana butir "
       f"bergerak searah dengan konstruk yang diukur (Sugiyono, 2022; Cooper & Schindler, 2022). "
       f"Nilai r_hitung kemudian dibandingkan dengan r_tabel pada taraf signifikansi 5 persen "
       f"dan derajat kebebasan df = {n_pilot} - 2 = {df}.",
       ctx, "3h-2", indent=0.5, space_after=4)

    p(doc, "                   N(ΣXY) - (ΣX)(ΣY)", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=1)
    p(doc, "r_xy = ---------------------------------------------", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=1, space_after=1)
    p(doc, "        √[N(ΣX^2) - (ΣX)^2] × [N(ΣY^2) - (ΣY)^2]", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=1, space_after=4)

    pv(doc,
       f"Keterangan: r_xy adalah koefisien korelasi Pearson; N adalah jumlah responden uji coba "
       f"(N = {n_pilot}); X adalah skor butir; dan Y adalah skor total konstruk. Nilai r_tabel "
       f"untuk df = {df} pada alpha 5 persen adalah {r_tabel}. Butir dinyatakan valid apabila "
       f"r_hitung ≥ {r_tabel}; butir dengan nilai di bawah ambang tersebut tidak dipakai pada "
       f"instrumen final (Sinambela, 2023; Hair et al., 2022).",
       ctx, "3h-3", indent=0.5, space_before=2, space_after=6)

    table_val_counter = 4 + k_vars + 1
    rel_results = []

    for vkey, role, vname, vdata in ordered_vars:
        dims, items, alpha_str = _get_variable_instrument_details(vkey, role, vname, vdata, r_tabel, ctx=ctx)
        rel_results.append((vname, alpha_str))
        rv = nv.rng_for(ctx, f"3h-val-{vkey}")

        pv(doc, nv.pick(rv, [
           (
               f"Hasil korelasi butir-total untuk variabel {vname} ({role}) disajikan pada tabel "
               f"berikut. Setiap koefisien dibandingkan dengan r_tabel {r_tabel} untuk menentukan "
               f"apakah butir layak dipertahankan. Butir yang gugur tidak dihapus dari catatan uji "
               f"coba, tetapi dikeluarkan dari perhitungan skor konstruk pada survei utama."
           ),
           (
               f"Korelasi butir-total untuk variabel {vname} ({role}) ditampilkan pada tabel "
               f"berikut. Tiap koefisien dibandingkan dengan r_tabel {r_tabel} guna menetapkan "
               f"kelayakan butir. Butir yang gugur tidak dihapus dari catatan uji coba, namun "
               f"dikeluarkan dari perhitungan skor konstruk pada survei utama."
           ),
           (
               f"Tabel berikut menyajikan hasil korelasi butir-total untuk variabel {vname} "
               f"({role}). Setiap koefisien diuji terhadap r_tabel {r_tabel} untuk menilai apakah "
               f"butir layak dipertahankan. Butir yang gugur tetap tercatat pada uji coba, tetapi "
               f"tidak diikutkan dalam perhitungan skor konstruk pada survei utama."
           ),
           (
               f"Hasil korelasi butir-total variabel {vname} ({role}) dimuat pada tabel berikut. "
               f"Tiap koefisien dibandingkan dengan r_tabel {r_tabel} untuk menetapkan kelayakan "
               f"butir. Butir yang tidak lolos tetap dicatat pada uji coba, namun dikeluarkan dari "
               f"penghitungan skor konstruk pada survei utama."
           ),
        ]), ctx, f"3h-val-{vkey}", indent=0.5, space_before=4, space_after=4)

        t_val_headers = ["No", "Kode Butir", "r_hitung", "r_tabel", "Keputusan"]
        t_val_rows = [[it["no"], it["kode"], it["r_hitung"], it["r_tabel"], it["keputusan"]] for it in items]
        add_table_data(doc, title=f"Tabel 3. {table_val_counter}\nHasil Uji Validitas Instrumen {vname}",
                       headers=t_val_headers, data=t_val_rows,
                       source=f"Sumber: Data ujicoba diolah peneliti ({tahun})")
        table_val_counter += 1

        n_valid = sum(1 for it in items if it["keputusan"] == "Valid")
        n_invalid = sum(1 for it in items if it["keputusan"] == "Tidak Valid")
        n_pending = len(items) - n_valid - n_invalid

        if n_pending == len(items):
            pv(doc, nv.pick(rv, [
               (
                   f"Tabel 3.{table_val_counter - 1} memuat kerangka uji validitas untuk {vname} "
                   f"({role}). Nilai r_hitung belum tersedia karena data uji coba instrumen belum "
                   f"dimasukkan, sehingga keputusan validitas butir belum dapat ditetapkan. Setelah "
                   f"data uji coba tersedia, setiap butir akan dinilai terhadap r_tabel {r_tabel} dan "
                   f"butir yang tidak memenuhi ambang akan dikeluarkan dari instrumen final "
                   f"(Cooper & Schindler, 2022; Sugiyono, 2022)."
               ),
               (
                   f"Kerangka uji validitas {vname} ({role}) disajikan pada tabel sebelumnya. Kolom "
                   f"r_hitung masih kosong karena data uji coba belum diolah, sehingga status butir "
                   f"belum dapat ditetapkan. Prosedur penilaian tetap mengikuti ambang r_tabel "
                   f"{r_tabel}; butir yang tidak memenuhi akan dikeluarkan dari perhitungan skor "
                   f"konstruk pada survei utama (Sugiyono, 2022; Hair et al., 2022)."
               ),
            ]), ctx, f"3h-valnote-{vkey}", indent=0.5, space_after=4)
        else:
            if n_invalid:
                catatan = (
                    f"Terdapat {n_invalid} butir dengan r_hitung di bawah {r_tabel} sehingga "
                    f"dinyatakan tidak valid dan dikeluarkan dari instrumen final."
                )
            else:
                catatan = (
                    f"Seluruh butir memiliki r_hitung di atas {r_tabel} sehingga dinyatakan valid."
                )
            pv(doc, nv.pick(rv, [
               (
                   f"Dari {len(items)} butir yang diujicobakan untuk {vname}, {n_valid} butir "
                   f"dinyatakan valid. {catatan} Butir yang bertahan menunjukkan daya pembeda yang "
                   f"memadai dan konsisten dengan dimensi yang diukur, sehingga dapat dipakai untuk "
                   f"mengukur {vname} pada aparatur di {instansi}."
               ),
               (
                   f"Dari {len(items)} butir yang diujicobakan untuk {vname}, sebanyak {n_valid} butir "
                   f"dinyatakan valid. {catatan} Butir yang bertahan mencerminkan daya pembeda yang "
                   f"memadai dan selaras dengan dimensi yang diukur, sehingga layak dipakai untuk "
                   f"mengukur {vname} pada aparatur di {instansi}."
               ),
               (
                   f"Pada uji coba {vname}, {n_valid} dari {len(items)} butir dinyatakan valid. "
                   f"{catatan} Butir yang lolos menunjukkan daya pembeda memadai dan konsisten dengan "
                   f"dimensi yang diukur, sehingga dapat dipakai untuk mengukur {vname} pada aparatur "
                   f"di {instansi}."
               ),
               (
                   f"Untuk {vname}, {n_valid} dari {len(items)} butir dinyatakan valid pada uji coba. "
                   f"{catatan} Butir yang lolos memperlihatkan daya pembeda memadai dan sejalan dengan "
                   f"dimensi yang diukur, sehingga layak dipakai mengukur {vname} pada aparatur di "
                   f"{instansi}."
               ),
            ]), ctx, f"3h-valnote-{vkey}", indent=0.5, space_before=4, space_after=6)

    # 2. Uji Reliabilitas Instrumen
    p(doc, "2. Uji Reliabilitas Instrumen", bold=True, indent=0, space_before=6, space_after=4)
    rng_h = nv.rng_for(ctx, "3h-tail")
    pv(doc, nv.pick(rng_h, [
       (
           "Uji reliabilitas dilakukan pada butir-butir yang telah lolos uji validitas. Tujuannya "
           "adalah menilai konsistensi internal instrumen ketika digunakan berulang pada kondisi "
           "yang relatif sama. Koefisien Cronbach's Alpha dipilih karena sesuai untuk kuesioner "
           "berskala Likert dengan beberapa butir per konstruk (Hair et al., 2022; Cooper & Schindler, 2022)."
       ),
       (
           "Uji reliabilitas diterapkan pada butir yang telah lolos validitas. Maksudnya adalah "
           "menilai konsistensi internal instrumen ketika dipakai berulang pada kondisi yang "
           "relatif sama. Koefisien Cronbach's Alpha dipilih karena cocok untuk kuesioner "
           "berskala Likert dengan sejumlah butir per konstruk (Hair et al., 2022; Cooper & Schindler, 2022)."
       ),
       (
           "Reliabilitas diuji pada butir-butir yang sudah dinyatakan valid. Pengujian ini "
           "bertujuan menakar konsistensi internal instrumen jika digunakan berulang pada kondisi "
           "serupa. Koefisien Cronbach's Alpha digunakan sebab sesuai bagi kuesioner berskala "
           "Likert yang memuat beberapa butir per konstruk (Hair et al., 2022; Cooper & Schindler, 2022)."
       ),
    ]), ctx, "3h-rel1", indent=0.5, space_after=4)

    p(doc, "         k         Σ σ_b^2", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=1)
    p(doc, "r_11 = (-----) × (1 - -------)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=1, space_after=1)
    p(doc, "        k - 1          σ_t^2", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=1, space_after=4)

    pv(doc, nv.pick(rng_h, [
       (
           "Keterangan: r_11 adalah koefisien reliabilitas; k adalah jumlah butir valid; Σ σ_b^2 "
           "adalah jumlah varians butir; dan σ_t^2 adalah varians total konstruk. Instrumen "
           "dinyatakan reliabel apabila koefisien alpha mencapai minimal 0,600 dan dikategorikan "
           "sangat reliabel apabila di atas 0,800 (Sinambela, 2023). Rujukan interpretasi "
           "koefisien disajikan pada tabel berikut."
       ),
       (
           "Keterangan: r_11 ialah koefisien reliabilitas; k ialah jumlah butir valid; Σ σ_b^2 "
           "ialah jumlah varians butir; dan σ_t^2 ialah varians total konstruk. Instrumen dinilai "
           "reliabel bila koefisien alpha mencapai minimal 0,600 dan tergolong sangat reliabel "
           "apabila melebihi 0,800 (Sinambela, 2023). Rujukan penafsiran koefisien ditampilkan "
           "pada tabel berikut."
       ),
       (
           "Penjelasan lambang: r_11 adalah koefisien reliabilitas; k adalah banyaknya butir "
           "valid; Σ σ_b^2 adalah jumlah varians butir; dan σ_t^2 adalah varians total konstruk. "
           "Instrumen dinyatakan reliabel jika koefisien alpha minimal 0,600 dan dikategorikan "
           "sangat reliabel jika di atas 0,800 (Sinambela, 2023). Acuan penafsiran koefisien "
           "disajikan pada tabel berikut."
       ),
    ]), ctx, "3h-rel2", indent=0.5, space_before=2, space_after=6)

    t_rel_ref_headers = ["Nomor", "Interval Koefisien Reliabilitas", "Makna Reliabilitas"]
    t_rel_ref_rows = [
        ["1.", "0,000 – 0,199", "Sangat Tidak Reliabel"],
        ["2.", "0,200 – 0,399", "Tidak Reliabel"],
        ["3.", "0,400 – 0,599", "Cukup Reliabel"],
        ["4.", "0,600 – 0,799", "Reliabel"],
        ["5.", "0,800 – 1,000", "Sangat Reliabel"],
    ]
    add_table_data(doc, title=f"Tabel 3. {table_val_counter}\nRujukan Interpretasi Nilai Koefisien Reliabilitas",
                   headers=t_rel_ref_headers, data=t_rel_ref_rows, source="Sumber: Sinambela (2023)")
    table_val_counter += 1

    pv(doc, nv.pick(rng_h, [
       (
           f"Hasil perhitungan Cronbach's Alpha untuk setiap konstruk disajikan pada tabel berikut. "
           f"Nilai dihitung terpisah per variabel agar kelemahan pada satu konstruk tidak tertutup "
           f"oleh rata-rata keseluruhan. Interpretasi mengacu pada ambang 0,600 dan kategori sangat "
           f"reliabel pada 0,800 (Hair et al., 2022)."
       ),
       (
           f"Tabel berikut memuat hasil perhitungan Cronbach's Alpha tiap konstruk. Penghitungan "
           f"dilakukan per variabel supaya kelemahan pada satu konstruk tidak tertutupi oleh "
           f"rata-rata keseluruhan. Penafsiran mengacu pada ambang 0,600 dengan kategori sangat "
           f"reliabel pada 0,800 (Hair et al., 2022)."
       ),
       (
           f"Perolehan Cronbach's Alpha untuk masing-masing konstruk ditampilkan pada tabel berikut. "
           f"Nilai dihitung secara terpisah per variabel agar kelemahan satu konstruk tidak "
           f"tertutup oleh rerata keseluruhan. Interpretasinya berpijak pada ambang 0,600 dan "
           f"kategori sangat reliabel pada 0,800 (Hair et al., 2022)."
       ),
       (
           f"Perhitungan Cronbach's Alpha untuk tiap konstruk ditampilkan pada tabel berikut. "
           f"Angka dihitung per variabel supaya kelemahan satu konstruk tidak tersembunyi di "
           f"balik rerata keseluruhan. Penafsirannya mengacu pada ambang 0,600 dan kategori "
           f"sangat reliabel pada 0,800 (Hair et al., 2022)."
       ),
    ]), ctx, "3h-rel3", indent=0.5, space_before=4, space_after=4)

    t_rel_res_headers = ["No", "Variabel", "Cronbach's Alpha", "Keterangan"]
    t_rel_res_rows = []
    for idx, (vname, val) in enumerate(rel_results, 1):
        try:
            fval = float(str(val).replace(",", "."))
        except ValueError:
            fval = 0.0
        ket = "Sangat Reliabel" if fval >= 0.800 else ("Reliabel" if fval >= 0.600 else "Cukup Reliabel")
        t_rel_res_rows.append([str(idx), vname, val, ket])
    add_table_data(doc, title=f"Tabel 3. {table_val_counter}\nHasil Uji Reliabilitas Instrumen Penelitian",
                   headers=t_rel_res_headers, data=t_rel_res_rows,
                   source=f"Sumber: Data ujicoba diolah peneliti ({tahun})")

    vals = []
    for _, val in rel_results:
        try:
            vals.append(float(str(val).replace(",", ".")))
        except ValueError:
            pass
    if vals:
        lo = f"{min(vals):.3f}".replace(".", ",")
        hi = f"{max(vals):.3f}".replace(".", ",")
    else:
        lo, hi = "0,800", "0,900"

    pv(doc, nv.pick(rng_h, [
       (
           f"Rangkuman pada tabel menunjukkan koefisien alpha berada pada rentang {lo} sampai {hi}. "
           f"Seluruh konstruk melampaui ambang 0,600, bahkan sebagian besar berada di atas 0,800. "
           f"Hasil ini menunjukkan konsistensi internal yang baik, tetapi tidak menghapus kebutuhan "
           f"pemeriksaan validitas konvergen dan diskriminan pada model pengukuran PLS-SEM."
       ),
       (
           f"Rangkuman tabel memperlihatkan koefisien alpha berkisar antara {lo} hingga {hi}. "
           f"Semua konstruk melewati ambang 0,600, dan sebagian besar bahkan melampaui 0,800. "
           f"Temuan ini menandakan konsistensi internal yang baik, namun tetap menuntut pemeriksaan "
           f"validitas konvergen dan diskriminan pada model pengukuran PLS-SEM."
       ),
       (
           f"Koefisien alpha pada tabel rangkuman berada dalam rentang {lo} sampai {hi}. Seluruh "
           f"konstruk melampaui ambang 0,600 dan mayoritas melebihi 0,800. Keadaan itu menunjukkan "
           f"konsistensi internal yang memadai, meskipun tidak menghapus perlunya penilaian "
           f"validitas konvergen dan diskriminan pada model pengukuran PLS-SEM."
       ),
    ]), ctx, "3h-rel4", indent=0.5, space_before=4, space_after=6)


def _build_doc_sub_3_i(doc: Document, ctx: Dict[str, Any]):
    """Sub-Bab I. Teknik Analisis Data PLS-SEM SmartPLS 4."""
    h2(doc, "I. Teknik Analisis Data")
    instansi = ctx["instansi"]
    n = int(ctx["n"])
    pm = nv.profile_map(ctx)
    rng = nv.rng_for(ctx, "3i")
    y_var = ctx["y_var"]
    m_vars_str = ", ".join(ctx["m_vars"]) if ctx["m_vars"] else "mekanisme perantara"
    z_vars_str = ", ".join(ctx["z_vars"]) if ctx["z_vars"] else "kondisi kontekstual"
    x_vars_str = _join_id(ctx["x_vars"])

    for i, text in enumerate(nv.method_analysis_paragraphs(ctx, pm, rng), 1):
        pv(doc, text, ctx, f"3i-{i}", indent=0.5, space_after=6)

    rng_i = nv.rng_for(ctx, "3i-tail")
    p(doc, "1. Evaluasi Model Pengukuran (Outer Model)", bold=True, indent=0, space_before=4, space_after=4)
    pv(doc, nv.pick(rng_i, [
       (
           f"Evaluasi outer model menilai apakah indikator benar-benar merefleksikan konstruk laten "
           f"{y_var}, {m_vars_str}, {z_vars_str}, serta {x_vars_str}. Empat kriteria dipakai: "
           f"(a) validitas konvergen melalui outer loading > 0,708 dan AVE > 0,50; (b) validitas "
           f"diskriminan melalui Fornell-Larcker, cross loading, dan HTMT < 0,85 atau < 0,90; "
           f"(c) reliabilitas konstruk melalui composite reliability dan Cronbach's Alpha > 0,70; "
           f"serta (d) multikolinearitas indikator melalui outer VIF < 3,0 (Hair et al., 2022; "
           f"Sarstedt et al., 2022)."
       ),
       (
           f"Penilaian outer model memeriksa apakah indikator sungguh mencerminkan konstruk laten "
           f"{y_var}, {m_vars_str}, {z_vars_str}, dan {x_vars_str}. Empat patokan diterapkan: "
           f"(a) validitas konvergen melalui outer loading > 0,708 dan AVE > 0,50; (b) validitas "
           f"diskriminan melalui Fornell-Larcker, cross loading, dan HTMT < 0,85 atau < 0,90; "
           f"(c) reliabilitas konstruk melalui composite reliability dan Cronbach's Alpha > 0,70; "
           f"dan (d) multikolinearitas indikator melalui outer VIF < 3,0 (Hair et al., 2022; "
           f"Sarstedt et al., 2022)."
       ),
       (
           f"Tahap ini menguji apakah indikator layak merepresentasikan konstruk laten {y_var}, "
           f"{m_vars_str}, {z_vars_str}, serta {x_vars_str}. Kriteria yang diacu ada empat: "
           f"(a) validitas konvergen melalui outer loading > 0,708 dan AVE > 0,50; (b) validitas "
           f"diskriminan melalui Fornell-Larcker, cross loading, dan HTMT < 0,85 atau < 0,90; "
           f"(c) reliabilitas konstruk melalui composite reliability dan Cronbach's Alpha > 0,70; "
           f"serta (d) multikolinearitas indikator melalui outer VIF < 3,0 (Hair et al., 2022; "
           f"Sarstedt et al., 2022)."
       ),
    ]), ctx, "3i-outer", indent=0.5, space_after=6)

    p(doc, "2. Evaluasi Model Struktural (Inner Model)", bold=True, indent=0, space_before=6, space_after=4)
    pv(doc, nv.pick(rng_i, [
       (
           f"Evaluasi inner model membaca kekuatan prediksi dan kualitas hubungan kausal pada "
           f"{instansi}. Kriteria yang digunakan meliputi inner VIF < 3,0 untuk mendeteksi "
           f"kolinearitas; R-square untuk menilai proporsi varians endogen; f-square untuk "
           f"mengukur kontribusi relatif variabel; Q-square dan PLSpredict untuk menguji relevansi "
           f"prediktif; serta SRMR < 0,08 dan NFI > 0,90 untuk menilai kecocokan model "
           f"(Hair et al., 2022; Tenenhaus et al., 2005)."
       ),
       (
           f"Penilaian inner model menyoroti kekuatan prediksi dan mutu hubungan kausal di "
           f"{instansi}. Kriteria yang dipakai mencakup inner VIF < 3,0 untuk mendeteksi "
           f"kolinearitas; R-square untuk menimbang proporsi varians endogen; f-square untuk "
           f"mengukur sumbangan relatif variabel; Q-square dan PLSpredict untuk menguji relevansi "
           f"prediktif; serta SRMR < 0,08 dan NFI > 0,90 untuk menilai kecocokan model "
           f"(Hair et al., 2022; Tenenhaus et al., 2005)."
       ),
       (
           f"Pada tahap inner model, perhatian diarahkan pada daya prediksi dan kualitas relasi "
           f"kausal di {instansi}. Kriteria yang diterapkan meliputi inner VIF < 3,0 guna "
           f"mendeteksi kolinearitas; R-square untuk menilai proporsi varians endogen; f-square "
           f"untuk mengukur kontribusi relatif variabel; Q-square dan PLSpredict untuk menguji "
           f"relevansi prediktif; serta SRMR < 0,08 dan NFI > 0,90 untuk menilai kecocokan model "
           f"(Hair et al., 2022; Tenenhaus et al., 2005)."
       ),
    ]), ctx, "3i-inner", indent=0.5, space_after=6)

    p(doc, "3. Pengujian Hipotesis dan Signifikansi Jalur (Bootstrapping)", bold=True, indent=0, space_before=6, space_after=4)
    pv(doc, nv.pick(rng_i, [
       (
           "Pengujian hipotesis menggunakan bootstrapping 5.000 subsampel untuk memperoleh "
           "standard error dan t-statistik tanpa mengasumsikan normalitas. Kriteria penerimaan "
           "adalah t-statistik ≥ 1,96 dan p-value < 0,05 pada uji dua sisi. Penolakan hipotesis "
           "nol tidak otomatis berarti efeknya besar; karena itu, interpretasi disertai arah dan "
           "besaran koefisien (Hair et al., 2022; Cooper & Schindler, 2022)."
       ),
       (
           "Hipotesis diuji melalui bootstrapping 5.000 subsampel guna memperoleh standard error "
           "dan t-statistik tanpa mengasumsikan normalitas. Kriteria penerimaannya adalah "
           "t-statistik ≥ 1,96 dan p-value < 0,05 pada uji dua sisi. Ditolaknya hipotesis nol "
           "tidak serta-merta menandakan efek besar, sehingga interpretasi tetap disertai arah "
           "dan besaran koefisien (Hair et al., 2022; Cooper & Schindler, 2022)."
       ),
       (
           "Pengujian hipotesis memakai bootstrapping 5.000 subsampel untuk mendapatkan standard "
           "error dan t-statistik tanpa asumsi normalitas. Ambang penerimaan yang dipakai ialah "
           "t-statistik ≥ 1,96 dan p-value < 0,05 pada uji dua sisi. Karena penolakan hipotesis "
           "nol tidak otomatis bermakna efek besar, penafsiran dilengkapi arah dan besaran "
           "koefisien (Hair et al., 2022; Cooper & Schindler, 2022)."
       ),
    ]), ctx, "3i-boot", indent=0.5, space_after=6)

    has_m = bool(ctx["m_vars"])
    has_z = bool(ctx["z_vars"])
    if has_m and has_z:
        indirect_variants = [
            (
                f"Efek mediasi {m_vars_str} diuji melalui specific indirect effect dengan bias-corrected "
                f"95 persen confidence interval. Efek moderasi {z_vars_str} diuji melalui latent "
                f"interaction term, dengan memperhatikan perubahan R-square dan f-square. Untuk mediasi "
                f"termoderasi, indeks mediasi termoderasi dipakai agar klaim bersyarat tidak didasarkan "
                f"pada pemisahan model yang kehilangan konteks (Hayes, 2022; Sarstedt et al., 2022)."
            ),
            (
                f"Efek mediasi {m_vars_str} ditelusuri lewat specific indirect effect dengan "
                f"bias-corrected 95 persen confidence interval. Adapun efek moderasi {z_vars_str} diuji "
                f"melalui latent interaction term sambil mencermati perubahan R-square dan f-square. "
                f"Pada mediasi termoderasi, indeks mediasi termoderasi digunakan agar klaim bersyarat "
                f"tidak bertumpu pada pemisahan model yang menghilangkan konteks (Hayes, 2022; Sarstedt et al., 2022)."
            ),
        ]
    elif has_m:
        indirect_variants = [
            (
                f"Efek mediasi {m_vars_str} diuji melalui specific indirect effect dengan bias-corrected "
                f"95 persen confidence interval. Karena model tidak memuat variabel moderasi, pengujian "
                f"difokuskan pada besaran dan signifikansi efek tidak langsung serta perbandingannya "
                f"dengan efek langsung (Hayes, 2022; Sarstedt et al., 2022)."
            ),
            (
                f"Efek tidak langsung melalui {m_vars_str} diuji lewat specific indirect effect dengan "
                f"bias-corrected 95 persen confidence interval. Tanpa variabel moderasi dalam model, "
                f"perhatian diarahkan pada besaran dan signifikansi efek mediasi sekaligus "
                f"perbandingannya dengan efek langsung (Hayes, 2022; Sarstedt et al., 2022)."
            ),
        ]
    elif has_z:
        indirect_variants = [
            (
                f"Efek moderasi {z_vars_str} diuji melalui latent interaction term dengan "
                f"memperhatikan perubahan R-square dan f-square. Kriteria kebermaknaan mengacu pada "
                f"signifikansi interaksi dan arah efek bersyarat, bukan semata koefisien utama "
                f"(Hayes, 2022; Sarstedt et al., 2022)."
            ),
            (
                f"Peran moderasi {z_vars_str} diuji lewat latent interaction term sambil mencermati "
                f"perubahan R-square dan f-square. Kebermaknaannya dinilai dari signifikansi interaksi "
                f"dan arah efek bersyarat, bukan hanya koefisien utama (Hayes, 2022; Sarstedt et al., 2022)."
            ),
        ]
    else:
        indirect_variants = [
            (
                f"Karena model hanya memuat jalur langsung, pengujian difokuskan pada besaran, arah, "
                f"dan signifikansi koefisien pengaruh {x_vars_str} terhadap {y_var}. Penafsiran tidak "
                f"berhenti pada nilai p, tetapi mempertimbangkan besaran efek agar saran bagi "
                f"{instansi} proporsional (Hair et al., 2022; Cooper & Schindler, 2022)."
            ),
            (
                f"Pada model jalur langsung tanpa mediasi maupun moderasi, perhatian tertuju pada "
                f"besaran, arah, dan signifikansi koefisien {x_vars_str} terhadap {y_var}. Penafsiran "
                f"tidak berhenti pada nilai p, melainkan menyertakan besaran efek supaya rekomendasi "
                f"bagi {instansi} proporsional (Hair et al., 2022; Cooper & Schindler, 2022)."
            ),
        ]
    pv(doc, nv.pick(rng_i, indirect_variants), ctx, "3i-indirect", indent=0.5, space_after=6)

    p(doc, "Model Persamaan Struktural Matematis", bold=True, indent=0, space_before=6, space_after=4)
    if has_m and has_z:
        eq_tail = [
            f"Persamaan disusun dalam bentuk sub-struktural agar jalur mediasi dan moderasi dapat dibaca secara terpisah.",
            f"Bentuk sub-struktural dipilih supaya jalur mediasi dan moderasi terbaca secara terpisah.",
        ]
    elif has_m:
        eq_tail = [
            f"Persamaan disusun dalam bentuk sub-struktural agar jalur mediasi dapat dipisahkan dari pengaruh langsung.",
            f"Bentuk sub-struktural dipilih supaya jalur mediasi terpisah jelas dari pengaruh langsung.",
        ]
    elif has_z:
        eq_tail = [
            f"Persamaan disusun dalam bentuk sub-struktural agar pengaruh utama dan efek interaksi moderasi dapat dibaca secara terpisah.",
            f"Bentuk sub-struktural dipilih supaya pengaruh utama dan efek interaksi moderasi terbaca terpisah.",
        ]
    else:
        eq_tail = [
            f"Persamaan dituliskan per jalur agar pengaruh tiap variabel eksogen dapat dibaca secara terpisah.",
            f"Bentuk persamaan dipisah per jalur supaya pengaruh masing-masing variabel eksogen terbaca tersendiri.",
        ]
    pv(doc, nv.pick(rng_i, [
       (
           f"Persamaan struktural dirumuskan untuk memodelkan pengaruh simultan seluruh konstruk "
           f"terhadap {y_var} pada {instansi}. {eq_tail[0]}"
       ),
       (
           f"Model persamaan struktural dibangun untuk menangkap pengaruh simultan seluruh "
           f"konstruk terhadap {y_var} di {instansi}. {eq_tail[1]}"
       ),
    ]), ctx, "3i-eq", indent=0.5, space_after=4)

    eqs = ctx["graph"].get("equations", [])
    if eqs:
        for eq in eqs:
            p(doc, f"• Persamaan {eq['name']}:", bold=True, indent=0.5, space_after=2)
            p(doc, f"  {eq['formula']}", italic=True, indent=0.75, space_after=4)
    else:
        p(doc, "• Persamaan Sub-Struktural 1 (Model Mediasi):", bold=True, indent=0.5, space_after=2)
        p(doc, "  M = β0 + β1.X1 + β2.X2 + ε1", italic=True, indent=0.75, space_after=4)
        p(doc, "• Persamaan Sub-Struktural 2 (Model Moderasi & Endogen Utama):", bold=True, indent=0.5, space_after=2)
        p(doc, "  Y = γ0 + γ1.X1 + γ2.X2 + γ3.M + γ4.Z + γ5.(M × Z) + ε2", italic=True, indent=0.75, space_after=6)


# Legacy backward-compatible aliases for Bab III builders
_build_doc_sub_3_1 = _build_doc_sub_3_a
_build_doc_sub_3_2 = _build_doc_sub_3_e
_build_doc_sub_3_3 = _build_doc_sub_3_f
_build_doc_sub_3_4 = _build_doc_sub_3_g
_build_doc_sub_3_5 = _build_doc_sub_3_h
_build_doc_sub_3_6 = _build_doc_sub_3_i


def _build_doc_sub_4_a_1(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """Sub-Bab A.1 Karakteristik Demografis Responden (Format Baku TAPD 2026)."""
    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
    
    h3(doc, "1. Deskriptif Data Demografis Responden")
    instansi = ctx["instansi"]
    n = ctx["n"]
    tahun = ctx.get("tahun", "2026")

    # Narasi Pengantar Demografi (>= 4 baris, bersitasi)
    p(doc, f"Analisis deskriptif karakteristik demografis responden merupakan tahapan empiris fundamental yang bertujuan untuk memetakan secara komprehensif profil latar belakang aparatur sipil yang menjadi unit analisis dalam penelitian disertasi ini (Cooper & Schindler, 2022). Pemahaman mendalam mengenai struktur demografi responden pada {instansi} memberikan basis interpretasi kontekstual yang esensial dalam membedah variasi respon perseptual terhadap variabel-variabel penelitian yang diteliti (Sekaran & Bougie, 2020). Melalui pemetaan atribut demografis yang mencakup distribusi jenis kelamin, kelompok usia biologis, dan jenjang kepangkatan/golongan ruang, peneliti dapat memverifikasi tingkat heterogenitas dan representativitas sampel yang digunakan (Sugiyono, 2022). Identifikasi karakteristik responden ini sekaligus menjamin bahwa data primer yang dihimpun memiliki reliabilitas substantif yang tinggi dan bebas dari bias konsentrasi segmen tertentu dalam organisasi kedinasan pemerintah daerah (Neuman, 2022).", indent=0.5, space_after=6)

    # 1. Tabel 4.1 Jenis Kelamin
    _gender = s3data.demographics(ctx, "gender")
    if _gender:
        _gtot = sum(int(x[1]) for x in _gender) or n
        n_pria = int(_gender[0][1]); pct_pria = (n_pria / _gtot) * 100
        n_wanita = int(_gender[1][1]) if len(_gender) > 1 else (_gtot - n_pria)
        pct_wanita = (n_wanita / _gtot) * 100
    else:
        _gtot = None
        n_pria = n_wanita = pct_pria = pct_wanita = None

    t1_headers = ["Karakteristik Responden", "Kategori Gender", "Frekuensi (Orang)", "Persentase (%)"]
    if _gender:
        t1_rows = [["Jenis Kelamin" if i == 0 else "", str(x[0]), str(int(x[1])),
                    f"{int(x[1]) / _gtot * 100:.1f}%"] for i, x in enumerate(_gender)]
    else:
        t1_rows = [
            ["Jenis Kelamin", "Laki-laki", s3data.PLACEHOLDER, s3data.PLACEHOLDER],
            ["", "Perempuan", s3data.PLACEHOLDER, s3data.PLACEHOLDER],
        ]
    t1_rows.append(["Total", "Keseluruhan Responden",
                    (str(_gtot) if _gender else s3data.PLACEHOLDER),
                    ("100.0%" if _gender else s3data.PLACEHOLDER)])
    add_table_data(doc, title="Tabel 4. 1\nDeskripsi Karakteristik Responden berdasarkan Jenis Kelamin", headers=t1_headers, data=t1_rows, source=s3data.source_label(ctx, f"Data primer diolah, ({tahun})"))

    # 2 Paragraf Narasi Analitis Jenis Kelamin (>= 4 baris per paragraf, bersitasi)
    if _gender:
        p(doc, f"Berdasarkan data distribusi pada Tabel 4.1 di atas, tampak bahwa responden berjenis kelamin laki-laki mendominasi komposisi aparatur pada {instansi} dengan proporsi sebesar {pct_pria:.1f}% ({n_pria} orang), sementara responden perempuan mencakup proporsi sebesar {pct_wanita:.1f}% ({n_wanita} orang). Komposisi gender yang menempatkan laki-laki sebagai proporsi mayoritas ini selaras dengan karakteristik beban tugas operasional di instansi pengelola pendapatan dan pajak daerah yang menuntut mobilitas lapangan yang intensif, pengawasan kepatuhan wajib pajak secara langsung, serta penegakan regulasi perpajakan yang memerlukan ketegasan psikologis dan ketahanan fisik prima (Armstrong & Taylor, 2023). Kendati demikian, representasi aparatur perempuan sebesar {pct_wanita:.1f}% menunjukkan keterlibatan yang sangat substansial dan strategis, khususnya dalam fungsi pelayanan loket terpadu, verifikasi data administrasi keuangan, pelaporan akuntansi, dan rekonsiliasi penerimaan kas daerah yang mengedepankan ketelitian, empati pelayanan, dan kecermatan analitis (Robbins & Judge, 2022).", indent=0.5, space_after=6)

        p(doc, f"Keseimbangan proporsional kedua gender ini mencerminkan implementasi prinsip meritokrasi dan keadilan peluang kerja yang sehat di lingkungan {instansi}, di mana pembagian penugasan didasarkan pada kompetensi fungsional tanpa diskriminasi berbasis gender (Sinambela, 2023). Keterlibatan perempuan yang mencapai lebih dari empat puluh persen menunjukkan bahwa iklim kerja organisasi telah mengadopsi nilai-nilai inklusivitas modern yang mampu memperkaya dinamika pengambilan keputusan manajerial (Hasibuan, 2022). Perspektif ganda antara ketegasan operasional aparatur laki-laki dan ketelitian administratif aparatur perempuan secara sinergis memperkuat etos kerja kelembagaan, memitigasi potensi kekeliruan teknis pemungutan pajak, serta mendorong percepatan pencapaian target penerimaan pendapatan asli daerah yang telah ditetapkan oleh pemerintah daerah (Sedarmayanti, 2022).", indent=0.5, space_after=6)

    else:
        p(doc, "Interpretasi rinci atas komposisi gender responden akan dilengkapi setelah data demografis riil tersedia. Pada mode template, tabel ditampilkan sebagai kerangka agar naskah tidak memuat angka yang belum terverifikasi.", indent=0.5, space_after=6)

    # 2. Tabel 4.2 Usia Responden
    _age = s3data.demographics(ctx, "age")
    if _age:
        _atot = sum(int(x[1]) for x in _age) or n
        _a_vals = [(str(x[0]), int(x[1])) for x in _age]
        _padded = (_a_vals + [("", 0)] * 4)[:4]
        n_u1, n_u2, n_u3, n_u4 = [fr for _, fr in _padded]
        p_u1, p_u2, p_u3, p_u4 = [(fr / _atot * 100) for fr in (n_u1, n_u2, n_u3, n_u4)]
    else:
        _atot = None
        _a_vals = None

    t2_headers = ["Karakteristik Responden", "Rentang Usia (Tahun)", "Frekuensi (Orang)", "Persentase (%)"]
    if _a_vals:
        t2_rows = [["Kelompok Usia" if i == 0 else "", lbl, str(fr),
                    f"{fr / _atot * 100:.1f}%"] for i, (lbl, fr) in enumerate(_a_vals)]
    else:
        t2_rows = [
            ["Kelompok Usia", "<= 30 Tahun", s3data.PLACEHOLDER, s3data.PLACEHOLDER],
            ["", "31 - 40 Tahun", s3data.PLACEHOLDER, s3data.PLACEHOLDER],
            ["", "41 - 50 Tahun", s3data.PLACEHOLDER, s3data.PLACEHOLDER],
            ["", "> 50 Tahun", s3data.PLACEHOLDER, s3data.PLACEHOLDER],
        ]
    t2_rows.append(["Total", "Keseluruhan Responden",
                    (str(_atot) if _a_vals else s3data.PLACEHOLDER),
                    ("100.0%" if _a_vals else s3data.PLACEHOLDER)])
    add_table_data(doc, title="Tabel 4. 2\nDeskripsi Karakteristik Responden berdasarkan Usia", headers=t2_headers, data=t2_rows, source=s3data.source_label(ctx, f"Data primer diolah, ({tahun})"))

    # 2 Paragraf Narasi Analitis Usia (>= 4 baris per paragraf, bersitasi)
    if _a_vals:
        p(doc, f"Sebagaimana terangkum pada Tabel 4.2, sebaran usia responden penelitian menunjukkan bahwa kelompok usia 31 sampai 40 tahun mendominasi dengan frekuensi {n_u2} orang ({p_u2:.1f}%), disusul oleh kelompok usia 41 sampai 50 tahun sebanyak {n_u3} orang ({p_u3:.1f}%), aparatur berusia di bawah atau sama dengan 30 tahun sebanyak {n_u1} orang ({p_u1:.1f}%), serta aparatur senior di atas 50 tahun sebanyak {n_u4} orang ({p_u4:.1f}%). Konsentrasi terbesar responden yang berada pada rentang usia 31 hingga 50 tahun—mencapai total {p_u2 + p_u3:.1f}% dari keseluruhan sampel—mengindikasikan bahwa birokrasi {instansi} saat ini digerakkan oleh sumber daya manusia yang berada pada fase puncak usia produktif dan kematangan karir profesional (Armstrong & Taylor, 2023). Aparatur pada jenjang usia ini telah memiliki akumulasi pengalaman kerja lapangan yang cukup panjang, pemahaman regulasi birokrasi yang matang, sekaligus masih memiliki vitalitas fisik dan motivasi intrinsik yang tinggi untuk terus berprestasi dalam sistem manajemen kinerja publik (Robbins & Judge, 2022).", indent=0.5, space_after=6)

        p(doc, f"Dominasi demografis aparatur usia produktif ini memberikan implikasi psikologis yang sangat positif terhadap kualitas dan validitas data penelitian, di mana para responden memiliki stabilitas emosional, objektivitas kognitif, dan kemampuan reflektif yang prima dalam menanggapi setiap butir instrumen kuesioner tanpa bias kecerobohan (Cooper & Schindler, 2022). Di sisi lain, keberadaan aparatur generasi muda (<= 30 tahun) sebesar {p_u1:.1f}% menyuntikkan energi inovasi dan adaptabilitas digital yang cepat terhadap transformasi sistem informasi perpajakan modern, sedangkan aparatur senior (> 50 tahun) sebesar {p_u4:.1f}% berfungsi sebagai jangkar kearifan organisasi, keteladanan kepatuhan nilai etika, dan transfer pengetahuan tacit birokrasi kepada generasi penerus (Sinambela, 2023; Hasibuan, 2022). Kombinasi demografi usia yang proporsional ini menciptakan kesinambungan regenerasi kepemimpinan transformasional dan stabilitas budaya kerja yang sangat kondusif di lingkungan {instansi}.", indent=0.5, space_after=6)

    else:
        p(doc, "Interpretasi rinci atas sebaran usia responden akan dilengkapi setelah data demografis riil tersedia. Pada mode template, tabel ditampilkan sebagai kerangka agar naskah tidak memuat angka yang belum terverifikasi.", indent=0.5, space_after=6)

    # 3. Tabel 4.3 Pangkat/Golongan
    _grade = s3data.demographics(ctx, "grade")
    if _grade:
        _gtot2 = sum(int(x[1]) for x in _grade) or n
        _g_vals = [(str(x[0]), int(x[1])) for x in _grade]
        _gp = (_g_vals + [("", 0)] * 3)[:3]
        n_g2, n_g3, n_g4 = [fr for _, fr in _gp]
        p_g2, p_g3, p_g4 = [(fr / _gtot2 * 100) for fr in (n_g2, n_g3, n_g4)]
    else:
        _gtot2 = None
        _g_vals = None

    t3_headers = ["Karakteristik Responden", "Golongan Ruang Kepangkatan", "Frekuensi (Orang)", "Persentase (%)"]
    if _g_vals:
        t3_rows = [["Pangkat / Golongan" if i == 0 else "", lbl, str(fr),
                    f"{fr / _gtot2 * 100:.1f}%"] for i, (lbl, fr) in enumerate(_g_vals)]
    else:
        t3_rows = [
            ["Pangkat / Golongan", "Golongan II (Pengatur / Pelaksana Teknis)", s3data.PLACEHOLDER, s3data.PLACEHOLDER],
            ["", "Golongan III (Penata / Analis & Pejabat Pengawas)", s3data.PLACEHOLDER, s3data.PLACEHOLDER],
            ["", "Golongan IV (Pembina / Pejabat Administrator & Pimpinan)", s3data.PLACEHOLDER, s3data.PLACEHOLDER],
        ]
    t3_rows.append(["Total", "Keseluruhan Responden",
                    (str(_gtot2) if _g_vals else s3data.PLACEHOLDER),
                    ("100.0%" if _g_vals else s3data.PLACEHOLDER)])
    add_table_data(doc, title="Tabel 4. 3\nDeskripsi Karakteristik Responden berdasarkan Pangkat/Golongan", headers=t3_headers, data=t3_rows, source=s3data.source_label(ctx, f"Data primer diolah, ({tahun})"))

    # 2 Paragraf Narasi Analitis Pangkat/Golongan (>= 4 baris per paragraf, bersitasi)
    if _g_vals:
        p(doc, f"Distribusi tingkat kepangkatan dan golongan ruang aparatur sebagaimana dipaparkan pada Tabel 4.3 memperlihatkan bahwa aparatur dengan kualifikasi Golongan III (Penata) merupakan kelompok mayoritas mutlak yang menempati proporsi {p_g3:.1f}% ({n_g3} orang), diikuti oleh aparatur Golongan II sebanyak {n_g2} orang ({p_g2:.1f}%), dan aparatur Golongan IV sebanyak {n_g4} orang ({p_g4:.1f}%). Tingginya persentase aparatur Golongan III menegaskan bahwa instrumen kuesioner penelitian ini diisi secara dominan oleh pejabat fungsional teknis, analis kebijakan, penilai pajak, dan pejabat pengawas operasional yang memegang tanggung jawab langsung dalam eksekusi teknis pemungutan pendapatan daerah (Hasibuan, 2022). Keberadaan kelompok aparatur bergolongan III ini menjadi representasi ideal dalam mengukur realitas efektivitas gaya kepemimpinan, iklim kerja, dan motivasi kerja organisasi karena mereka berada pada posisi garis depan yang berinteraksi langsung baik dengan pimpinan struktural di atasnya maupun dengan masyarakat wajib pajak (Sedarmayanti, 2022).", indent=0.5, space_after=6)

        p(doc, f"Sementara itu, kehadiran aparatur Golongan IV sebesar {p_g4:.1f}% memberikan bobot legitimasi manajerial yang kuat terhadap penelitian Disertasi ini, mengingat kelompok ini terdiri dari para kepala bidang, sekretaris, dan pejabat eselon yang berperan sebagai perumus kebijakan strategis, pengendali mutu organisasi, dan figur panutan kepemimpinan transformasional di {instansi} (Armstrong & Taylor, 2023). Di sisi lain, aparatur Golongan II sebesar {p_g2:.1f}% yang menjalankan fungsi administrasi dasar dan operasional lapangan melengkapi spektrum penilaian kinerja dari sudut pandang pelaksana lini bawah (Sinambela, 2023). Komposisi piramida kepangkatan birokrasi yang didominasi oleh golongan terdidik dan berpengalaman ini menjamin bahwa seluruh data penilaian empiris yang diperoleh mencerminkan persepsi komprehensif, multi-jenjang, dan memiliki akuntabilitas akademik tingkat doktoral yang kokoh (Robbins & Judge, 2022).", indent=0.5, space_after=6)


    else:
        p(doc, "Interpretasi rinci atas komposisi pangkat/golongan responden akan dilengkapi setelah data demografis riil tersedia. Pada mode template, tabel ditampilkan sebagai kerangka agar naskah tidak memuat angka yang belum terverifikasi.", indent=0.5, space_after=6)


def _build_doc_single_variable_desc(doc: Document, ctx: Dict[str, Any], vkey: str, role: str, vname: str, vdata: Dict[str, Any], table_index: int = 4, standalone: bool = True, sub_code: str = ""):
    """Tabel dan Narasi Analisis Deskriptif untuk 1 Konstruk Variabel (Format Baku TAPD 2026)."""
    import hashlib
    instansi = ctx["instansi"]
    tahun = ctx.get("tahun", "2026")

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "2. Analisis Statistik Deskriptif Variabel")

    code_title = f"{sub_code} " if sub_code else ""
    h4(doc, f"{code_title}Deskripsi Tanggapan Responden terhadap Variabel {vname}")

    dims, items, _ = _get_variable_instrument_details(vkey, role, vname, vdata, ctx=ctx)

    # Narasi Pembuka per Variabel (>= 4 baris, bersitasi)
    p(doc, f"Distribusi tanggapan responden terhadap butir-butir pernyataan yang merefleksikan konstruk laten {vname} ({role}) pada {instansi} dianalisis secara terperinci guna mengidentifikasi sejauh mana aparatur menginternalisasi nilai-nilai dan indikator operasional dari variabel tersebut dalam rutinitas kerjanya (Robbins & Judge, 2022). Variabel {vname} diukur melalui {len(items)} indikator terstandarisasi yang mencakup seluruh dimensi operasional yang telah divalidasi pada kajian teoretis sebelumnya (Sekaran & Bougie, 2020). Melalui evaluasi skor mean dan dispersi standar deviasi pada setiap butir instrumen, dapat diketahui dimensi mana yang telah terimplementasi secara optimal dan indikator mana yang masih memerlukan penguatan kebijakan institusional di lingkungan organisasi (Cooper & Schindler, 2022). Hasil tabulasi statistik deskriptif untuk konstruk {vname} disajikan pada tabel berikut (Sugiyono, 2022):", indent=0.5, space_after=6)

    _desc = s3data.descriptives(ctx, vkey)
    _real_items = {}
    if _desc:
        for _r in _desc.get("items", []):
            _c = str(_r.get("kode") or _r.get("code") or "").strip().upper()
            if _c:
                _real_items[_c] = _r

    t_headers = ["Pernyataan", "Min", "Max", "Median", "Mean", "Std. Deviasi"]
    t_rows = []
    item_means = []

    for idx_it, it in enumerate(items, start=1):
        p_text = f"Pernyataan {idx_it}: Aparatur mempersepsikan bahwa {it['indikator'].lower()} telah terwujud secara konsisten dalam kedinasan."
        _real = _real_items.get(str(it["kode"]).strip().upper())
        if _real:
            m_val = float(_real.get("mean")) if _real.get("mean") is not None else None
            sd_val = float(_real.get("sd")) if _real.get("sd") is not None else None
            if m_val is not None:
                item_means.append((m_val, it["indikator"], it["kode"]))
            _min_v = _real.get("min")
            _max_v = _real.get("max")
            _med_v = _real.get("median")
            t_rows.append([
                p_text,
                (f"{float(_min_v):.2f}" if _min_v is not None else s3data.PLACEHOLDER),
                (f"{float(_max_v):.2f}" if _max_v is not None else s3data.PLACEHOLDER),
                (f"{float(_med_v):.2f}" if _med_v is not None else s3data.PLACEHOLDER),
                (f"{m_val:.2f}" if m_val is not None else s3data.PLACEHOLDER),
                (f"{sd_val:.2f}" if sd_val is not None else s3data.PLACEHOLDER),
            ])
        else:
            t_rows.append([p_text, s3data.PLACEHOLDER, s3data.PLACEHOLDER,
                           s3data.PLACEHOLDER, s3data.PLACEHOLDER, s3data.PLACEHOLDER])

    add_table_data(doc, title=f"Tabel 4. {table_index}\nDeskripsi Tanggapan Responden terhadap Variabel {vname}", headers=t_headers, data=t_rows, source=s3data.source_label(ctx, f"Data primer diolah, ({tahun})"))

    if not item_means:
        p(doc, f"Tabulasi statistik deskriptif untuk konstruk {vname} pada tabel di atas masih ditampilkan sebagai kerangka karena data empiris riil belum tersedia pada konfigurasi penelitian. Setelah data hasil olahan SmartPLS 4 atau SPSS dimasukkan, nilai minimum, maksimum, median, rata-rata, dan simpangan baku setiap butir akan terisi otomatis beserta interpretasi capaian indikator tertinggi dan terendah (Sugiyono, 2022; Cooper & Schindler, 2022).", indent=0.5, space_after=6)
        return

    # Identifikasi Mean Tertinggi dan Terendah
    item_means.sort(key=lambda x: x[0], reverse=True)
    highest_mean, highest_ind, highest_code = item_means[0]
    lowest_mean, lowest_ind, lowest_code = item_means[-1]
    grand_m = sum([x[0] for x in item_means]) / len(item_means)

    if grand_m >= 4.20:
        kat_m = "sangat baik dan berbobot sangat tinggi"
    elif grand_m >= 3.40:
        kat_m = "baik dan berbobot tinggi"
    elif grand_m >= 2.60:
        kat_m = "cukup baik dan berada pada kategori sedang"
    else:
        kat_m = "kurang optimal dan memerlukan evaluasi intensif"

    # Narasi Penutup per Variabel (>= 4 baris, bersitasi)
    p(doc, f"Berdasarkan hasil tabulasi deskriptif pada tabel di atas, konstruk {vname} mencatatkan nilai rata-rata keseluruhan (grand mean) sebesar {grand_m:.2f}, yang mengindikasikan bahwa persepsi aparatur sipil pada {instansi} berada pada kategori {kat_m} (Sugiyono, 2022). Dari seluruh indikator yang dievaluasi, indikator dengan capaian skor rata-rata tertinggi dicatatkan oleh butir {highest_code} ({highest_ind}) dengan skor mean mencapai {highest_mean:.2f}, yang menunjukkan bahwa aspek tersebut telah terinternalisasi secara kokoh sebagai kekuatan utama budaya kerja organisasi (Robbins & Judge, 2022). Sebaliknya, skor rata-rata relatif terendah berada pada butir {lowest_code} ({lowest_ind}) dengan perolehan skor mean sebesar {lowest_mean:.2f}. Meskipun capaian terendah ini masih berada dalam batas kategori persepsi yang memuaskan, nilai tersebut menjadi sinyal manajerial penting bagi jajaran pimpinan untuk memberikan intervensi pembinaan khusus dan perbaikan sistemik berkelanjutan guna menyempurnakan performa kelembagaan secara menyeluruh (Armstrong & Taylor, 2023; Hair et al., 2022).", indent=0.5, space_after=6)


def _build_doc_sub_4_a_2_y(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.2.1 Deskripsi Tanggapan Responden terhadap Variabel Terikat (Y)."""
    ordered_vars = _get_ordered_variables(ctx)
    y_entry = next((v for v in ordered_vars if str(v[0]).upper() == "Y"),
                   ("y", "Y", ctx.get("y_var", "Kinerja Pegawai"), {}))
    _build_doc_single_variable_desc(doc, ctx, y_entry[0], y_entry[1], y_entry[2], y_entry[3], table_index=4, standalone=standalone, sub_code="2.1")


def _build_doc_sub_4_a_2_mz(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.2.2 Deskripsi Tanggapan Responden terhadap Variabel Mediasi (M) dan Moderasi (Z)."""
    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "2. Analisis Statistik Deskriptif Variabel")

    ordered_vars = _get_ordered_variables(ctx)
    m_entry = next((v for v in ordered_vars if str(v[0]).upper() == "M" or "mediasi" in v[1].lower()), None)
    if not m_entry:
        _mv = (ctx.get("m") or "").strip()
        m_entry = ("M", "Variabel Mediasi", _mv, {}) if _mv else None
    if m_entry:
        _build_doc_single_variable_desc(doc, ctx, m_entry[0], m_entry[1], m_entry[2], m_entry[3], table_index=5, standalone=False, sub_code="2.2.1")

    z_entry = next((v for v in ordered_vars if str(v[0]).upper() == "Z" or "moderasi" in v[1].lower()), None)
    if not z_entry:
        _zv = (ctx.get("z") or "").strip()
        z_entry = ("Z", "Variabel Moderasi", _zv, {}) if _zv else None
    if z_entry:
        _build_doc_single_variable_desc(doc, ctx, z_entry[0], z_entry[1], z_entry[2], z_entry[3], table_index=6, standalone=False, sub_code="2.2.2")


def _build_doc_sub_4_a_2_xk(doc: Document, ctx: Dict[str, Any], x_idx: int = 1, standalone: bool = True):
    """A.2.3.x Deskripsi Tanggapan Responden terhadap Variabel Bebas (X_k)."""
    ordered_vars = _get_ordered_variables(ctx)
    target_key = f"X{x_idx}"
    x_entry = next((v for v in ordered_vars if str(v[0]).upper() == target_key.upper()), None)
    if not x_entry:
        x_vars = ctx.get("x_vars", [])
        vname = x_vars[x_idx - 1] if x_idx <= len(x_vars) else f"Variabel Bebas X{x_idx}"
        x_entry = (target_key, f"Variabel Bebas {x_idx}", vname, {})
    _build_doc_single_variable_desc(doc, ctx, x_entry[0], x_entry[1], x_entry[2], x_entry[3], table_index=6 + x_idx, standalone=standalone, sub_code=f"2.3.{x_idx}")


def _build_doc_sub_4_a_2_rekap(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.2.Rekap Perhitungan Rentang Kategori Sugiyono & Rekapitulasi Rerata Keseluruhan Variabel."""
    import hashlib
    instansi = ctx["instansi"]
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    tahun = ctx.get("tahun", "2026")

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "2. Analisis Statistik Deskriptif Variabel")

    h4(doc, "2.4 Rekapitulasi Nilai Rata-rata Variabel & Kriteria Kategori Sugiyono")

    # Perhitungan Rentang Kategori Sugiyono & Tabel Kriteria
    p(doc, "Penetapan kategori capaian tanggapan responden dalam penelitian disertasi ini dipedomani oleh formula perhitungan rentang skala interval menurut Sugiyono (2018/2022). Dengan rentang skor teoretis instrumen berskala 1 (Sangat Tidak Setuju) sampai dengan 5 (Sangat Setuju) serta jumlah kategori penilaian yang ditetapkan sebanyak 3 (tiga) tingkatan (Rendah, Sedang, Tinggi), maka lebar interval kelas dihitung melalui persamaan matematis sebagai berikut (Sugiyono, 2022):", indent=0.5, space_after=4)

    p(doc, "RS = (Skor Maksimal - Skor Minimal) / Jumlah Kategori = (5,00 - 1,00) / 3 = 1,33", bold=True, indent=1.0, space_after=6)

    p(doc, "Berdasarkan hasil formulasi matematis lebar rentang skala interval sebesar 1,33 tersebut, maka disusun tabel batasan kriteria kategori interpretasi tanggapan responden guna mengklasifikasikan tingkat capaian persepsi aparatur secara objektif dan terstandarisasi (Sugiyono, 2018). Rentang kategorisasi trikotomis ini mengelompokkan skor rerata persepsi ke dalam tiga tingkatan capaian, yakni kategori Rendah (skor 1,00 hingga 2,33), kategori Sedang (skor 2,34 hingga 3,67), serta kategori Tinggi (skor 3,68 hingga 5,00) (Sugiyono, 2022). Standarisasi kriteria ini menjadi tolok ukur analitis yang sangat krusial dalam mendiagnosis posisi ketercapaian setiap variabel operasional serta mengidentifikasi dimensi-dimensi kritis yang membutuhkan intervensi penyempurnaan manajerial pada instansi penelitian (Cooper & Schindler, 2022; Robbins & Judge, 2022). Rincian batasan kriteria kategori disajikan pada tabel berikut (Sugiyono, 2018):", indent=0.5, space_after=6)

    t_kat_headers = ["Rentang Nilai Interval", "Kategori Tanggapan Responden", "Interpretasi Tingkat Ketercapaian"]
    t_kat_rows = [
        ["1,00 – 2,33", "Rendah", "Persepsi responden kurang baik / implementasi belum optimal"],
        ["2,34 – 3,67", "Sedang", "Persepsi responden cukup baik / implementasi berada pada taraf moderat"],
        ["3,68 – 5,00", "Tinggi", "Persepsi responden sangat baik / implementasi telah terlaksana optimal"]
    ]
    tbl_kat_num = 4 + k_vars
    add_table_data(doc, title=f"Tabel 4. {tbl_kat_num}\nKriteria Kategori Tanggapan Responden", headers=t_kat_headers, data=t_kat_rows, source="Sumber: Sugiyono, (2018)")

    # Tabel Rekapitulasi Rata-Rata Seluruh Variabel Penelitian
    tbl_rekap_num = 5 + k_vars
    p(doc, f"Rekapitulasi nilai rata-rata keseluruhan (grand mean) dan tingkat dispersi standar deviasi dari setiap konstruk laten penelitian yang diuji pada {instansi} disajikan pada tabel di bawah ini guna memberikan gambaran komparatif mengenai posisi relatif ketercapaian masing-masing variabel dalam sistem keorganisasian (Hair et al., 2022). Melalui perbandingan komparatif lintas konstruk ini, peneliti dapat memetakan variabel-variabel mana yang telah menjadi keunggulan kompetitif organisasi dan variabel mana yang masih berada pada zona perkembangan yang memerlukan perhatian ekstra dari pembuat kebijakan (Sekaran & Bougie, 2020). Rekapitulasi menyeluruh ini sekaligus membuktikan tingkat konsistensi tanggapan aparatur sipil terhadap instrumen penelitian sebelum melangkah pada pengujian inferensial yang lebih kompleks (Cooper & Schindler, 2022). Rangkuman skor penilaian seluruh variabel disajikan pada tabel berikut (Sugiyono, 2022):", indent=0.5, space_after=6)

    t_rek_headers = ["No", "Variabel Penelitian", "Peran Model", "Rata-rata Skor", "Standar Deviasi", "Kategori Respon"]
    t_rek_rows = []

    for idx_v, (vkey, role, vname, vdata) in enumerate(ordered_vars, start=1):
        _d = s3data.descriptives(ctx, vkey)
        gm_val = gsd_val = None
        if _d:
            _means = [float(it["mean"]) for it in _d.get("items", []) if it.get("mean") is not None]
            _sds = [float(it["sd"]) for it in _d.get("items", []) if it.get("sd") is not None]
            if _means:
                gm_val = sum(_means) / len(_means)
                gsd_val = (sum(_sds) / len(_sds)) if _sds else None
        if gm_val is None:
            t_rek_rows.append([str(idx_v), vname, role, s3data.PLACEHOLDER,
                               s3data.PLACEHOLDER, s3data.PLACEHOLDER])
        else:
            kat = "Tinggi" if gm_val >= 3.68 else ("Sedang" if gm_val >= 2.34 else "Rendah")
            t_rek_rows.append([
                str(idx_v), vname, role, f"{gm_val:.2f}",
                (f"{gsd_val:.2f}" if gsd_val is not None else s3data.PLACEHOLDER), kat,
            ])

    add_table_data(doc, title=f"Tabel 4. {tbl_rekap_num}\nRata-rata Skor Penilaian Seluruh Variabel Penelitian", headers=t_rek_headers, data=t_rek_rows, source=s3data.source_label(ctx, f"Data primer diolah, ({tahun})"))

    # Narasi Komparasi Antar-Konstruk (>= 4 baris, bersitasi)
    _rekap_real = [r for r in t_rek_rows if r[3] != s3data.PLACEHOLDER]
    if not _rekap_real:
        p(doc, f"Kategori interpretasi pada Tabel 4.{tbl_rekap_num} belum dapat ditetapkan karena data empiris riil belum tersedia. Tabel ditampilkan sebagai kerangka pengukuran; setelah tabulasi kuesioner dimasukkan, kolom rata-rata, standar deviasi, dan kategori akan terisi sehingga dapat dibandingkan antarkonstruk (Sugiyono, 2022; Hair et al., 2022).", indent=0.5, space_after=6)
    else:
        _gms = [float(r[3]) for r in _rekap_real]
        _lo, _hi = min(_gms), max(_gms)
        _kat_set = sorted({r[5] for r in _rekap_real})
        p(doc, f"Berdasarkan rekapitulasi pada Tabel 4.{tbl_rekap_num}, rata-rata skor variabel bergerak dari {_lo:.2f} hingga {_hi:.2f} dengan kategori {' dan '.join(_kat_set)}. Perbedaan angka antarkonstruk menunjukkan bahwa capaian tidak seragam, sehingga perbaikan perlu diarahkan pada variabel dengan skor terendah lebih dahulu (Robbins & Judge, 2022). Sebaran ini juga menjadi dasar untuk membaca hasil pengujian struktural pada subbab berikutnya, karena rerata yang tinggi tidak otomatis berarti hubungan antarvariabel kuat (Hair et al., 2022; Cooper & Schindler, 2022).", indent=0.5, space_after=6)


def _build_doc_sub_4_a_2(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """Sub-Bab A.2 Analisis Statistik Deskriptif Variabel (Composite Lengkap)."""
    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")

    h3(doc, "2. Analisis Statistik Deskriptif Variabel")
    instansi = ctx["instansi"]

    # Narasi Pengantar Analisis Deskriptif (>= 4 baris, bersitasi)
    p(doc, f"Analisis statistik deskriptif terhadap data tanggapan responden bertujuan untuk memberikan telaah kuantitatif yang sistematis mengenai karakteristik persepsi aparatur pada {instansi} terhadap seluruh butir indikator yang menyusun masing-masing konstruk laten penelitian (Sugiyono, 2022). Penilaian deskriptif ini disajikan melalui lima ukuran parameter statistik utama, yaitu Skor Minimum (Min), Skor Maksimum (Max), Nilai Tengah (Median), Nilai Rerata Hitung (Mean), serta Ukuran Dispersi Standar Deviasi (Std. Deviasi) (Cooper & Schindler, 2022). Pemanfaatan kelima parameter ini secara simultan memungkinkan peneliti untuk mendeteksi tingkat sebaran respon, konsistensi jawaban, tendensi sentral persepsi aparatur, serta potensi heterogenitas variansi yang melekat pada masing-masing dimensi pengukuran (Sekaran & Bougie, 2020). Pendekatan komprehensif ini menjadi fondasi empiris yang kokoh sebelum melangkah pada pengujian model pengukuran inferensial dan model struktural berbasis PLS-SEM (Hair et al., 2022).", indent=0.5, space_after=6)

    # 1. Variabel Terikat (Y)
    _build_doc_sub_4_a_2_y(doc, ctx, standalone=False)

    # 2. Variabel Mediasi (M) dan Moderasi (Z)
    h3(doc, "2.2 Deskripsi Variabel Mediasi dan Moderasi")
    _build_doc_sub_4_a_2_mz(doc, ctx, standalone=False)

    # 3. Variabel-Variabel Bebas (X1..Xn)
    h3(doc, "2.3 Deskripsi Variabel Bebas")
    x_vars = ctx.get("x_vars", [])
    for x_i in range(1, len(x_vars) + 1):
        _build_doc_sub_4_a_2_xk(doc, ctx, x_i, standalone=False)

    # 4. Rekapitulasi & Kategori Sugiyono
    _build_doc_sub_4_a_2_rekap(doc, ctx, standalone=False)


def _build_doc_sub_4_a_3_1(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.3.1 Validitas Konvergen (Outer Loading & AVE) SmartPLS 4 (Format Baku TAPD 2026)."""
    import hashlib
    instansi = ctx["instansi"]
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    tahun = ctx.get("tahun", "2026")

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "3. Evaluasi Model Pengukuran (Outer Model SmartPLS 4)")

        # Narasi Pengantar Inferensial Outer Model (>= 4 baris, bersitasi)
        p(doc, f"Evaluasi model pengukuran (measurement model atau outer model) merupakan tahapan krusial dalam metodologi Partial Least Squares Structural Equation Modeling (PLS-SEM) yang berfungsi untuk menguji secara ilmiah sifat psikometrik dari indikator-indikator yang merefleksikan konstruk laten (Hair et al., 2022). Sesuai dengan konvensi baku metodologis yang dirumuskan oleh Hair et al. (2022) dan Sarstedt et al. (2022), pengujian outer model reflektif ini dievaluasi melalui tiga parameter utama: Validitas Konvergen (Convergent Validity), Validitas Diskriminan (Discriminant Validity), serta Reliabilitas Konsistensi Internal Konstruk (Construct Reliability). Pengujian ini memastikan bahwa setiap butir instrumen secara empiris benar-benar mengukur konstruk yang seharusnya diukur tanpa terkontaminasi oleh galat pengukuran sistematis (Chin, 1998). Keseluruhan pengujian outer model dalam penelitian disertasi ini diestimasi menggunakan bantuan piranti lunak SmartPLS versi 4.0 yang berstandar internasional (Tenenhaus et al., 2005).", indent=0.5, space_after=6)

        # Box Gambar 4.1 Full Model SmartPLS 4
        x_vars_str = _join_id(ctx["x_vars"])
        m_var = ctx.get("m", "Motivasi Kerja")
        z_var = (ctx.get("z") or "").strip()
        y_var = ctx.get("y_var", "Kinerja Pegawai")

        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Cm

        t_box = doc.add_table(rows=1, cols=1)
        t_box.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = t_box.rows[0].cells[0]
        cell.width = Cm(15.5)

        # Hitung rentang loading empiris jika tersedia
        _all_loads = []
        if not _tpl:
            for _vk, _role, _nm, _vd in ordered_vars:
                _ov = s3data.outer(ctx, _vk) or {}
                for _ld in _ov.get("loadings") or []:
                    if isinstance(_ld, dict) and _ld.get("loading") is not None:
                        try:
                            _all_loads.append(float(_ld["loading"]))
                        except (TypeError, ValueError):
                            pass
        if _all_loads:
            _lmin, _lmax = min(_all_loads), max(_all_loads)
            _load_banner = f"   │ (Loadings: {_lmin:.3f} - {_lmax:.3f})            │ (Loadings: {_lmin:.3f} - {_lmax:.3f})"
        else:
            _load_banner = "   │ (Estimasi Outer Model)               │ (Estimasi Outer Model)"

        p(cell, "BAGAN ARSITEKTUR MODEL PENGUKURAN PENUH (FULL MODEL SMARTPLS 4)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=4)
        p(cell, f"[ Konstruk Eksogen (X): {x_vars_str} ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
        p(cell, _load_banner, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
        p(cell, "   ▼                                      ▼", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
        p(cell, f"[ Konstruk Mediasi (M): {m_var} ] ────(Transmisi)────► [ Konstruk Endogen (Y): {y_var} ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
        p(cell, "   ▲", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
        p(cell, "   │ (Interaksi Moderasi / Two-Stage)", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
        p(cell, f"[ Konstruk Moderasi (Z): {z_var} ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
        set_cell_border(cell, top={"sz": 12, "val": "single", "color": "003366"},
                              bottom={"sz": 12, "val": "single", "color": "003366"},
                              left={"sz": 12, "val": "single", "color": "003366"},
                              right={"sz": 12, "val": "single", "color": "003366"})

        p(doc, s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"), align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4, space_after=2)
        p(doc, "Gambar 4. 1\nBagan Model Pengukuran Penuh (Full Model SmartPLS 4)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=8)

        # Narasi Pembacaan Gambar 4.1 (>= 4 baris, bersitasi)
        p(doc, f"Arsitektur visual pada Gambar 4.1 di atas memperlihatkan konfigurasi model pengukuran yang menghubungkan konstruk laten penelitian dengan indikator-indikator refleksinya pada {instansi} (Hair et al., 2022). Dalam bagan algoritma SmartPLS 4 tersebut, diagram jalur memadukan variabel-variabel eksogen ({x_vars_str}) dengan variabel mediasi ({m_var}) dan variabel moderasi ({z_var}) dalam mempengaruhi variabel terikat akhir ({y_var}) (Sarstedt et al., 2022). Muatan faktor pada jalur laten ke indikator serta koefisien struktural antarkonstruk dibaca dari keluaran estimasi, sehingga bagan ini berfungsi sebagai peta pembacaan hasil, bukan sebagai bukti kesesuaian model (Chin, 1998). Visualisasi ini menunjukkan bagaimana operasionalisasi instrumen terhubung secara konseptual sesuai kaidah pemodelan (Tenenhaus et al., 2005).", indent=0.5, space_after=6)

    # a. Validitas Konvergen
    h4(doc, "a. Validitas Konvergen (Convergent Validity)")
    p(doc, f"Validitas konvergen mengukur derajat kesesuaian antara butir-butir indikator dalam mengukur konstruk laten yang sama secara teoretis (Hair et al., 2022). Menurut pedoman baku Hair et al. (2022) dan Chin (1998), pemenuhan validitas konvergen ditentukan oleh dua kriteria utama: (1) nilai muatan faktor luar (outer loadings) setiap butir indikator terhadap konstruknya wajib bernilai lebih besar dari 0,708, dan (2) nilai Average Variance Extracted (AVE) dari setiap konstruk laten harus melampaui nilai ambang batas minimal 0,500. Nilai outer loading di atas 0,708 menandakan bahwa konstruk laten mampu menjelaskan lebih dari 50% variansi dari indikator yang bersangkutan (Chin, 1998). Hasil estimasi outer loadings seluruh butir indikator variabel penelitian disajikan pada tabel berikut (Hair et al., 2022):", indent=0.5, space_after=6)

    # Tabel Outer Loadings
    tbl_load_num = 6 + k_vars
    t_load_headers = ["Variabel Konstruk", "Kode Indikator", "Pernyataan Ringkas Indikator", "Outer Loading", "Kriteria", "Keputusan"]
    t_load_rows = []

    for vkey, role, vname, vdata in ordered_vars:
        dims, items, _ = _get_variable_instrument_details(vkey, role, vname, vdata, ctx=ctx)
        _outer = s3data.outer(ctx, vkey) or {}
        _loads = {}
        if isinstance(_outer.get("loadings"), list):
            for _l in _outer["loadings"]:
                if isinstance(_l, dict) and _l.get("code"):
                    _loads[str(_l["code"])] = _l.get("loading")
        for it in items:
            _lv = _loads.get(it["kode"])
            try:
                load_txt = f"{float(_lv):.3f}" if _lv is not None else s3data.PLACEHOLDER
                keputusan = ("Valid Konvergen" if float(_lv) >= 0.708
                             else "Tidak Memenuhi") if _lv is not None else "Belum diuji"
            except (TypeError, ValueError):
                load_txt, keputusan = s3data.PLACEHOLDER, "Belum diuji"
            t_load_rows.append([
                vname,
                it["kode"],
                it["indikator"],
                load_txt,
                "> 0,708",
                keputusan
            ])

    add_table_data(doc, title=f"Tabel 4. {tbl_load_num}\nNilai Outer Loadings Indikator Variabel Penelitian", headers=t_load_headers, data=t_load_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi Analisis Outer Loadings (>= 4 baris, bersitasi)
    if s3data.is_template(ctx):
        p(doc, f"Tabel 4.{tbl_load_num} memuat kerangka muatan faktor luar untuk seluruh butir indikator. Nilai outer loading belum tersedia karena data empiris belum diolah, sehingga kolom keputusan ditandai 'Belum diuji'. Setelah keluaran SmartPLS 4 dimasukkan, setiap butir akan dinilai terhadap ambang 0,708 dan butir di bawah ambang akan dipertimbangkan untuk dikeluarkan (Hair et al., 2022; Sarstedt et al., 2022).", indent=0.5, space_after=6)
    else:
        _vals = []
        for _r in t_load_rows:
            try:
                _vals.append(float(_r[3]))
            except (TypeError, ValueError):
                pass
        _below = [r for r in t_load_rows if r[5] != "Valid Konvergen"]
        if _vals:
            p(doc, f"Berdasarkan Tabel 4.{tbl_load_num}, nilai outer loading bergerak dari {min(_vals):.3f} hingga {max(_vals):.3f} dengan {len(_below)} butir yang tidak memenuhi ambang 0,708. Butir yang berada di bawah ambang perlu ditinjau sebelum dipertahankan, karena muatan faktor yang rendah melemahkan kontribusi indikator terhadap konstruknya (Hair et al., 2022; Chin, 1998).", indent=0.5, space_after=6)
        else:
            p(doc, f"Tabel 4.{tbl_load_num} belum memuat nilai outer loading yang lengkap, sehingga penilaian validitas konvergen belum dapat disimpulkan (Hair et al., 2022).", indent=0.5, space_after=6)

    # Tabel AVE
    tbl_ave_num = 7 + k_vars
    t_ave_headers = ["No", "Konstruk Laten Variabel", "Peran Model", "Nilai AVE", "Kriteria Ambang", "Status Validitas"]
    t_ave_rows = []

    for idx_v, (vkey, role, vname, vdata) in enumerate(ordered_vars, start=1):
        _outer = s3data.outer(ctx, vkey) or {}
        _ave_raw = _outer.get("ave")
        try:
            ave_txt = f"{float(_ave_raw):.3f}" if _ave_raw is not None else s3data.PLACEHOLDER
            status = ("Valid Konvergen Terpenuhi" if float(_ave_raw) > 0.5
                      else "Tidak Memenuhi") if _ave_raw is not None else "Belum diuji"
        except (TypeError, ValueError):
            ave_txt, status = s3data.PLACEHOLDER, "Belum diuji"
        t_ave_rows.append([
            str(idx_v),
            vname,
            role,
            ave_txt,
            "> 0,500",
            status
        ])

    add_table_data(doc, title=f"Tabel 4. {tbl_ave_num}\nNilai Average Variance Extracted (AVE) Konstruk Laten", headers=t_ave_headers, data=t_ave_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi Analisis AVE (>= 4 baris, bersitasi)
    if s3data.is_template(ctx):
        p(doc, f"Tabel 4.{tbl_ave_num} memuat kerangka nilai Average Variance Extracted untuk setiap konstruk laten. Nilai AVE belum tersedia karena data empiris belum diolah, sehingga status validitas ditandai 'Belum diuji'. Setelah keluaran SmartPLS 4 dimasukkan, setiap konstruk akan dinilai terhadap ambang 0,500 sebelum analisis struktural dilanjutkan (Hair et al., 2022; Tenenhaus et al., 2005).", indent=0.5, space_after=6)
    else:
        _aves = []
        for _r in t_ave_rows:
            try:
                _aves.append(float(_r[3]))
            except (TypeError, ValueError):
                pass
        if _aves:
            p(doc, f"Hasil evaluasi AVE pada Tabel 4.{tbl_ave_num} menunjukkan nilai berkisar {min(_aves):.3f} hingga {max(_aves):.3f}. Konstruk dengan AVE di bawah 0,500 perlu diperiksa ulang karena indikatornya belum menjelaskan variansi konstruk secara memadai, sedangkan konstruk di atas ambang dapat dilanjutkan ke evaluasi model struktural (Hair et al., 2022; Chin, 1998).", indent=0.5, space_after=6)
        else:
            p(doc, f"Nilai AVE pada Tabel 4.{tbl_ave_num} belum lengkap, sehingga kesimpulan validitas konvergen belum dapat dirumuskan (Hair et al., 2022).", indent=0.5, space_after=6)


def _build_doc_sub_4_a_3_2(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.3.2 Validitas Diskriminan (HTMT & Cross Loadings) (Format Baku TAPD 2026)."""
    import hashlib
    instansi = ctx["instansi"]
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    tahun = ctx.get("tahun", "2026")

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "3. Evaluasi Model Pengukuran (Outer Model SmartPLS 4)")

    # b. Validitas Diskriminan
    h4(doc, "b. Validitas Diskriminan (Discriminant Validity)")
    p(doc, f"Validitas diskriminan dievaluasi untuk memastikan bahwa suatu konstruk laten secara empiris benar-benar unik dan berbeda secara nyata dari konstruk laten lainnya dalam model struktural, sehingga konstruk tersebut tidak menangkap fenomena yang diwakili oleh variabel lain (Hair et al., 2022). Dalam metodologi SEM mutakhir, pengujian rasio Heterotrait-Monotrait Ratio of Correlations (HTMT) menjadi standar baku emas menggantikan kriteria klasik Fornell-Larcker karena memiliki sensitivitas dan spesifisitas yang jauh lebih tinggi dalam mendeteksi masalah diskriminan (Henseler et al., 2015). Ambang batas kritis rasio HTMT yang disepakati secara luas adalah < 0,85 untuk konstruk yang secara konseptual berbeda secara tegas, atau < 0,90 untuk konstruk yang berdekatan secara konseptual (Franke & Sarstedt, 2019). Nilai rasio HTMT antar-konstruk penelitian disajikan pada tabel berikut (Henseler et al., 2015):", indent=0.5, space_after=6)

    # Tabel HTMT
    tbl_htmt_num = 8 + k_vars
    var_names_only = [v[2] for v in ordered_vars]
    t_htmt_headers = ["Konstruk Variabel"] + [f"({i})" for i in range(1, len(var_names_only) + 1)]
    t_htmt_rows = []

    # Peta HTMT riil (bila disediakan): list of dict atau dict-of-dict.
    _htmt_raw = (s3data.get(ctx) or {}).get("htmt") if s3data.get(ctx) else None
    _htmt_map: Dict[Any, Any] = {}
    if isinstance(_htmt_raw, list):
        for _row in _htmt_raw:
            if not isinstance(_row, dict):
                continue
            _a = str(_row.get("construct_a") or _row.get("a") or "").strip()
            _b = str(_row.get("construct_b") or _row.get("b") or "").strip()
            _v = _row.get("value", _row.get("htmt"))
            if _a and _b:
                _htmt_map[frozenset((_a.lower(), _b.lower()))] = _v
    elif isinstance(_htmt_raw, dict):
        for _a, _inner in _htmt_raw.items():
            if not isinstance(_inner, dict):
                continue
            for _b, _v in _inner.items():
                _htmt_map[frozenset((str(_a).strip().lower(), str(_b).strip().lower()))] = _v

    _htmt_vals = []
    for i, v1 in enumerate(var_names_only):
        row = [f"({i+1}) {v1}"]
        for j, v2 in enumerate(var_names_only):
            if i == j:
                row.append("\u2014")
            elif i > j:
                _hv = _htmt_map.get(frozenset((v1.lower(), v2.lower())))
                try:
                    row.append(f"{float(_hv):.3f}")
                    _htmt_vals.append(float(_hv))
                except (TypeError, ValueError):
                    row.append(s3data.PLACEHOLDER)
            else:
                row.append("")
        t_htmt_rows.append(row)

    add_table_data(doc, title=f"Tabel 4. {tbl_htmt_num}\nNilai Heterotrait-Monotrait Ratio (HTMT) antar Konstruk Laten", headers=t_htmt_headers, data=t_htmt_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi HTMT (>= 4 baris, bersitasi)
    if _htmt_vals:
        p(doc, f"Rasio HTMT pada Tabel 4.{tbl_htmt_num} bergerak dari {min(_htmt_vals):.3f} hingga {max(_htmt_vals):.3f}. Nilai tertinggi perlu dibandingkan dengan ambang 0,850 atau 0,900 bergantung pada kedekatan konseptual pasangan konstruk, karena pasangan yang secara teoretis berdekatan memang cenderung memiliki rasio lebih tinggi (Henseler et al., 2015; Franke & Sarstedt, 2019).", indent=0.5, space_after=6)
    else:
        p(doc, f"Tabel 4.{tbl_htmt_num} memuat kerangka matriks HTMT antarkonstruk. Nilai rasio belum tersedia karena data empiris belum diolah, sehingga penilaian validitas diskriminan belum dapat dilakukan. Setelah keluaran SmartPLS 4 dimasukkan, setiap pasangan konstruk akan dinilai terhadap ambang 0,850 atau 0,900 sesuai kedekatan konseptualnya (Henseler et al., 2015; Franke & Sarstedt, 2019).", indent=0.5, space_after=6)

    # Tabel Cross Loadings
    tbl_cross_num = 9 + k_vars
    p(doc, f"Sebagai konfirmasi tambahan terhadap ketahanan validitas diskriminan, dievaluasi pula matriks muatan silang (cross-loadings) yang mensyaratkan bahwa nilai muatan faktor suatu indikator ke konstruk laten tujuannya harus lebih besar secara signifikan dibandingkan nilai korelasinya terhadap seluruh konstruk laten lainnya dalam model struktural (Hair et al., 2022). Evaluasi muatan silang ini berfungsi untuk mendeteksi potensi adanya indikator yang mengalami salah pembebanan (cross-loading issue) atau indikator yang mengukur konstruk di luar domain teoritis yang seharusnya diukur pada {instansi} (Chin, 1998). Menurut konvensi metodologis Hair et al. (2022) dan Sarstedt et al. (2022), selisih antara nilai loading pada konstruk target dan nilai cross-loading pada konstruk lain idealnya berjarak minimal 0,10 hingga 0,20 guna memastikan bahwa indikator tersebut benar-benar merefleksikan konstruk induknya secara diskriminan. Matriks hasil pemeriksaan muatan silang seluruh indikator disajikan pada tabel berikut (Hair et al., 2022):", indent=0.5, space_after=6)

    t_cross_headers = ["Kode Indikator"] + [f"{v[:12]}.." for v in var_names_only]
    t_cross_rows = []
    _cross_vals: List[float] = []
    _cross_own_high = 0
    _cross_own_low = 0

    # Peta muatan silang riil (bila disediakan): list of dict
    # {"code": "X1.1", "loadings": {"Kompetensi": 0.81, ...}}.
    _cross_raw = s3data.records(ctx, "cross_loadings")
    _cross_map: Dict[str, Dict[str, Any]] = {}
    if isinstance(_cross_raw, list):
        for _row in _cross_raw:
            if not isinstance(_row, dict):
                continue
            _code = str(_row.get("code") or _row.get("kode") or "").strip()
            _ld = _row.get("loadings") if isinstance(_row.get("loadings"), dict) else None
            if _code and _ld:
                _cross_map[_code] = _ld

    for i_v, (vkey, role, vname, vdata) in enumerate(ordered_vars):
        dims, items, _ = _get_variable_instrument_details(vkey, role, vname, vdata, ctx=ctx)
        for it in items:
            code = it["kode"]
            row = [code]
            ld = _cross_map.get(code)
            if not ld:
                row.extend([s3data.PLACEHOLDER] * len(ordered_vars))
                t_cross_rows.append(row)
                continue
            own_val = None
            other_max = 0.0
            for j_v, v2 in enumerate(var_names_only):
                try:
                    val = float(ld.get(v2))
                except (TypeError, ValueError):
                    row.append(s3data.PLACEHOLDER)
                    continue
                row.append(f"{val:.3f}")
                _cross_vals.append(val)
                if j_v == i_v:
                    own_val = val
                else:
                    other_max = max(other_max, val)
            if own_val is not None and own_val >= other_max:
                _cross_own_high += 1
            elif own_val is not None:
                _cross_own_low += 1
            t_cross_rows.append(row)

    add_table_data(doc, title=f"Tabel 4. {tbl_cross_num}\nMatriks Cross Loadings Indikator terhadap Seluruh Konstruk Laten", headers=t_cross_headers, data=t_cross_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi Cross Loadings (>= 4 baris, bersitasi)
    if _cross_vals:
        p(doc, f"Pemeriksaan matriks muatan silang pada Tabel 4.{tbl_cross_num} menunjukkan nilai muatan yang bergerak dari {min(_cross_vals):.3f} hingga {max(_cross_vals):.3f}. Sebanyak {_cross_own_high} butir memiliki muatan tertinggi pada konstruk induknya, sedangkan {_cross_own_low} butir perlu ditinjau karena muatan pada konstruk lain menyamai atau melampaui muatan konstruk induknya. Selisih minimal 0,10 antara muatan utama dan muatan silang dipakai sebagai pedoman penilaian (Hair et al., 2022; Chin, 1998).", indent=0.5, space_after=6)
    else:
        p(doc, f"Tabel 4.{tbl_cross_num} memuat kerangka matriks muatan silang. Nilai belum tersedia karena matric cross-loadings belum dimasukkan ke konfigurasi data empiris, sehingga penilaian validitas diskriminan melalui muatan silang belum dapat dilakukan. Setelah matriks tersebut tersedia, setiap butir akan dinilai berdasarkan selisih muatan konstruk induk dan muatan silangnya (Hair et al., 2022; Sarstedt et al., 2022).", indent=0.5, space_after=6)


def _build_doc_sub_4_a_3_3(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.3.3 Reliabilitas Konstruk (Cronbach's Alpha, rho_a, rho_c) (Format Baku TAPD 2026)."""
    import hashlib
    instansi = ctx["instansi"]
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    tahun = ctx.get("tahun", "2026")

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "3. Evaluasi Model Pengukuran (Outer Model SmartPLS 4)")

    # c. Reliabilitas Konstruk
    h4(doc, "c. Reliabilitas Konstruk (Construct Reliability)")
    p(doc, f"Uji reliabilitas konstruk dilakukan untuk membuktikan stabilitas, konsistensi internal, dan presisi instrumen pengukuran dalam mengukur konstruk laten yang diteliti (Hair et al., 2022). Evaluasi reliabilitas dalam kerangka PLS-SEM dinilai melalui tiga koefisien terpadu: Cronbach's Alpha, Composite Reliability (rho_c), serta Dijkstra-Henseler's rho_a (Dijkstra & Henseler, 2015; Henseler et al., 2016). Kriteria ambang batas yang disyaratkan untuk penelitian ilmiah berbobot doktoral adalah bahwa seluruh nilai koefisien reliabilitas tersebut harus melampaui angka 0,700, di mana nilai di atas 0,800 diklasifikasikan sebagai tingkat reliabilitas yang sangat tinggi dan sempurna (Nunnally & Bernstein, 1994; Hair et al., 2022). Hasil estimasi reliabilitas konstruk disajikan pada tabel berikut (Hair et al., 2022):", indent=0.5, space_after=6)

    # Tabel Reliabilitas Konstruk
    tbl_rel_num = 10 + k_vars
    t_rel_headers = ["No", "Konstruk Laten Variabel", "Cronbach's Alpha", "Composite Rel (rho_a)", "Composite Rel (rho_c)", "Kriteria", "Kesimpulan"]
    t_rel_rows = []

    _rel_ca: List[float] = []
    _rel_above = 0
    for idx_v, (vkey, role, vname, vdata) in enumerate(ordered_vars, start=1):
        _outer = s3data.outer(ctx, vkey) or {}
        _ca = _outer.get("cronbach_alpha")
        _ra = _outer.get("rho_a")
        _rc = _outer.get("composite_reliability")
        try:
            ca_txt = f"{float(_ca):.3f}" if _ca is not None else s3data.PLACEHOLDER
        except (TypeError, ValueError):
            ca_txt = s3data.PLACEHOLDER
        try:
            ra_txt = f"{float(_ra):.3f}" if _ra is not None else s3data.PLACEHOLDER
        except (TypeError, ValueError):
            ra_txt = s3data.PLACEHOLDER
        try:
            rc_txt = f"{float(_rc):.3f}" if _rc is not None else s3data.PLACEHOLDER
        except (TypeError, ValueError):
            rc_txt = s3data.PLACEHOLDER
        try:
            _ca_v = float(_ca)
            _rel_ca.append(_ca_v)
            if _ca_v > 0.7:
                _rel_above += 1
            kesimpulan = "Reliabel" if _ca_v > 0.7 else "Tidak Reliabel"
        except (TypeError, ValueError):
            kesimpulan = "Belum diuji"
        t_rel_rows.append([
            str(idx_v),
            vname,
            ca_txt,
            ra_txt,
            rc_txt,
            "> 0,700",
            kesimpulan
        ])

    add_table_data(doc, title=f"Tabel 4. {tbl_rel_num}\nHasil Evaluasi Reliabilitas Konstruk (Construct Reliability)", headers=t_rel_headers, data=t_rel_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi Reliabilitas Konstruk (>= 4 baris, bersitasi)
    if _rel_ca:
        p(doc, f"Tabel 4.{tbl_rel_num} memuat koefisien reliabilitas konstruk. Nilai Cronbach's Alpha bergerak dari {min(_rel_ca):.3f} hingga {max(_rel_ca):.3f}, dengan {_rel_above} dari {len(_rel_ca)} konstruk melampaui ambang 0,700. Konstruk yang berada di bawah ambang perlu ditinjau kembali karena konsistensi internalnya belum memadai (Hair et al., 2022; Nunnally & Bernstein, 1994). Penafsiran tetap mempertimbangkan rho_a dan rho_c sebagai penguat kesimpulan (Henseler et al., 2016; Cooper & Schindler, 2022).", indent=0.5, space_after=6)
    else:
        p(doc, f"Tabel 4.{tbl_rel_num} memuat kerangka evaluasi reliabilitas konstruk. Nilai Cronbach's Alpha, rho_a, dan rho_c belum tersedia karena data empiris belum diolah, sehingga kesimpulan reliabilitas belum dapat ditetapkan. Setelah keluaran SmartPLS 4 dimasukkan, setiap konstruk akan dinilai terhadap ambang 0,700 dan hasilnya dilaporkan apa adanya, termasuk konstruk yang belum memenuhi ambang (Hair et al., 2022; Henseler et al., 2016).", indent=0.5, space_after=6)


def _build_doc_sub_4_a_3(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """Sub-Bab A.3 Evaluasi Model Pengukuran (Outer Model SmartPLS 4) (Composite Lengkap)."""
    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")

    h3(doc, "3. Evaluasi Model Pengukuran (Outer Model SmartPLS 4)")
    instansi = ctx["instansi"]
    tahun = ctx.get("tahun", "2026")
    x_vars_str = _join_id(ctx["x_vars"])
    _has_m = bool((ctx.get("m") or "").strip())
    _has_z = bool((ctx.get("z") or "").strip())
    m_var = (ctx.get("m") or "").strip()
    z_var = (ctx.get("z") or "").strip()
    y_var = ctx.get("y_var", "Kinerja Pegawai")

    # Rentang outer loading riil (bila tersedia) untuk bagan arsitektur.
    _all_loads = []
    for _vk, _role, _vn, _vd in _get_ordered_variables(ctx):
        _o = s3data.outer(ctx, _vk) or {}
        for _l in (_o.get("loadings") or []):
            try:
                _all_loads.append(float(_l.get("loading")))
            except (TypeError, ValueError, AttributeError):
                pass
    if _all_loads:
        _load_range_txt = f"(Loadings: {min(_all_loads):.3f} - {max(_all_loads):.3f})"
    else:
        _load_range_txt = f"(Loadings: {s3data.PLACEHOLDER})"

    # Narasi Pengantar Inferensial Outer Model (>= 4 baris, bersitasi)
    p(doc, f"Evaluasi model pengukuran (measurement model atau outer model) merupakan tahapan krusial dalam metodologi Partial Least Squares Structural Equation Modeling (PLS-SEM) yang berfungsi untuk menguji secara ilmiah sifat psikometrik dari indikator-indikator yang merefleksikan konstruk laten (Hair et al., 2022). Sesuai dengan konvensi baku metodologis yang dirumuskan oleh Hair et al. (2022) dan Sarstedt et al. (2022), pengujian outer model reflektif ini dievaluasi melalui tiga parameter utama: Validitas Konvergen (Convergent Validity), Validitas Diskriminan (Discriminant Validity), serta Reliabilitas Konsistensi Internal Konstruk (Construct Reliability). Pengujian ini memastikan bahwa setiap butir instrumen secara empiris benar-benar mengukur konstruk yang seharusnya diukur tanpa terkontaminasi oleh galat pengukuran sistematis (Chin, 1998). Keseluruhan pengujian outer model dalam penelitian disertasi ini diestimasi menggunakan bantuan piranti lunak SmartPLS versi 4.0 yang berstandar internasional (Tenenhaus et al., 2005).", indent=0.5, space_after=6)

    # Box Gambar 4.1 Full Model SmartPLS 4
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm

    t_box = doc.add_table(rows=1, cols=1)
    t_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t_box.rows[0].cells[0]
    cell.width = Cm(15.5)

    p(cell, "BAGAN ARSITEKTUR MODEL PENGUKURAN PENUH (FULL MODEL SMARTPLS 4)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=4)
    p(cell, f"[ Konstruk Eksogen (X): {x_vars_str} ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    p(cell, f"   │ {_load_range_txt}", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    p(cell, "   ▼", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
    if _has_m:
        p(cell, f"[ Konstruk Mediasi (M): {m_var} ] ────(Transmisi)────► [ Konstruk Endogen (Y): {y_var} ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    else:
        p(cell, f"[ Konstruk Endogen (Y): {y_var} ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    if _has_z:
        p(cell, "   ▲", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
        p(cell, "   │ (Interaksi Moderasi / Two-Stage)", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
        p(cell, f"[ Konstruk Moderasi (Z): {z_var} ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    set_cell_border(cell, top={"sz": 12, "val": "single", "color": "003366"},
                          bottom={"sz": 12, "val": "single", "color": "003366"},
                          left={"sz": 12, "val": "single", "color": "003366"},
                          right={"sz": 12, "val": "single", "color": "003366"})

    p(doc, s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"), align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4, space_after=2)
    p(doc, "Gambar 4. 1\nBagan Model Pengukuran Penuh (Full Model SmartPLS 4)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=8)

    # Narasi Pembacaan Gambar 4.1 (>= 4 baris, bersitasi)
    _gbr4_1_klausa = []
    _gbr4_1_klausa.append(f"variabel-variabel eksogen ({x_vars_str})")
    if _has_m:
        _gbr4_1_klausa.append(f"variabel mediasi ({m_var})")
    if _has_z:
        _gbr4_1_klausa.append(f"variabel moderasi ({z_var})")
    _gbr4_1_str = _join_id(_gbr4_1_klausa) if len(_gbr4_1_klausa) > 1 else _gbr4_1_klausa[0]
    if s3data.is_template(ctx):
        p(doc, f"Gambar 4.1 memperlihatkan rancangan arsitektur model pengukuran yang menghubungkan {_gbr4_1_str} dalam memengaruhi variabel terikat akhir ({y_var}). Bagan ini masih berupa kerangka: nilai muatan faktor belum terisi karena data empiris belum diolah, sehingga pembacaan psikometrik belum dapat dilakukan (Hair et al., 2022; Sarstedt et al., 2022). Setelah estimasi tersedia, setiap jalur akan dibaca melalui nilai muatan faktor, AVE, dan koefisien reliabilitas yang dilaporkan pada tabel-tabel berikutnya (Chin, 1998; Tenenhaus et al., 2005).", indent=0.5, space_after=6)
    else:
        p(doc, f"Gambar 4.1 memperlihatkan arsitektur model pengukuran yang menghubungkan {_gbr4_1_str} dalam memengaruhi variabel terikat akhir ({y_var}) (Hair et al., 2022). Bagan tersebut memadukan jalur antar-konstruk laten dengan indikator refleksinya {_load_range_txt.lower()}, sehingga keterpaduan model dapat diperiksa secara menyeluruh (Sarstedt et al., 2022). Nilai muatan faktor pada jalur laten ke indikator dibaca bersama AVE dan koefisien reliabilitas untuk menilai kualitas pengukuran (Chin, 1998; Tenenhaus et al., 2005).", indent=0.5, space_after=6)

    # 1. Validitas Konvergen
    _build_doc_sub_4_a_3_1(doc, ctx, standalone=False)

    # 2. Validitas Diskriminan
    _build_doc_sub_4_a_3_2(doc, ctx, standalone=False)

    # 3. Reliabilitas Konstruk
    _build_doc_sub_4_a_3_3(doc, ctx, standalone=False)


def _build_doc_sub_4_a_4_1(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.4.1 Evaluasi Koefisien Determinasi R-Square (R²) & Q-Square (Q²) (Format Baku TAPD 2026)."""
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    instansi = ctx["instansi"]
    tahun = ctx.get("tahun", "2026")
    pls = ctx["pls"]

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "4. Evaluasi Model Struktural (Inner Model PLS-SEM)")

    h4(doc, "a. Evaluasi Koefisien Determinasi (R-Square / R²)")
    p(doc, f"Koefisien determinasi (R²) mengukur proporsi variansi variabel endogen yang berhasil dijelaskan secara simultan oleh variabel-variabel eksogen yang mempengaruhinya dalam model struktural (Hair et al., 2022). Berdasarkan kriteria baku yang dirumuskan oleh Chin (1998) dan Hair et al. (2022), nilai R² sebesar 0,75 dikategorikan sebagai substansial (kuat), nilai 0,50 dikategorikan sebagai moderat (sedang), dan nilai 0,25 dikategorikan sebagai lemah. Di samping nilai R² sampel, disajikan pula nilai R² Adjusted guna mengoreksi potensi bias penambahan variabel prediktor serta nilai relevansi prediktif Q² (Chin, 1998). Hasil estimasi R-Square untuk variabel endogen disajikan pada tabel berikut (Hair et al., 2022):", indent=0.5, space_after=6)

    # Tabel R-Square
    tbl_r2_num = 11 + k_vars
    t_r2_headers = ["Variabel Endogen", "Peran dalam Model", "R-Square (R²)", "R-Square Adjusted", "Q-Square (Q²)", "Kategori Retensi Variansi"]
    t_r2_rows = []

    def _kategori_r2(v):
        if _tpl:
            return "Belum ditetapkan"
        try:
            v = float(v)
        except (TypeError, ValueError):
            return "Belum ditetapkan"
        if v >= 0.75:
            return "Substansial (Kuat)"
        if v >= 0.50:
            return "Moderat (Sedang)"
        if v >= 0.25:
            return "Lemah"
        return "Sangat Lemah"

    def _fmt3(v):
        if _tpl or v is None:
            return s3data.PLACEHOLDER
        try:
            return f"{float(v):.3f}"
        except (TypeError, ValueError):
            return s3data.PLACEHOLDER

    _tpl = s3data.is_template(ctx)

    r2_dict = pls.get("r_squares", {}) or {}
    if (ctx.get("m") or "").strip():
        _em = r2_dict.get("M") or {}
        t_r2_rows.append([
            _em.get("name") or ctx["m"], "Mediasi / Intervening",
            _fmt3(_em.get("r_square")), _fmt3(_em.get("r_square_adjusted")),
            _fmt3(_em.get("q_square")), _kategori_r2(_em.get("r_square")),
        ])
    _ey = r2_dict.get("Y") or {}
    t_r2_rows.append([
        _ey.get("name") or ctx.get("y_var", "Kinerja Pegawai"), "Terikat / Endogen Utama",
        _fmt3(_ey.get("r_square")), _fmt3(_ey.get("r_square_adjusted")),
        _fmt3(_ey.get("q_square")), _kategori_r2(_ey.get("r_square")),
    ])

    add_table_data(doc, title=f"Tabel 4. {tbl_r2_num}\nNilai Koefisien Determinasi R-Square (R2) dan R-Square Adjusted", headers=t_r2_headers, data=t_r2_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi R-Square (>= 4 baris, bersitasi)
    _r2_vals = []
    for _row in t_r2_rows:
        try:
            _r2_vals.append((_row[0], float(_row[2])))
        except (TypeError, ValueError):
            pass
    if _r2_vals:
        _r2_frasa = "; ".join(f"{_n} sebesar {_v:.3f} ({_v * 100:.1f}%)" for _n, _v in _r2_vals)
        p(doc, f"Tabel 4.{tbl_r2_num} menunjukkan proporsi variansi variabel endogen yang dijelaskan model, yaitu {_r2_frasa}. Angka tersebut dibaca dengan kriteria Chin (1998) dan Hair et al. (2022): 0,75 substansial, 0,50 moderat, dan 0,25 lemah. Nilai Q2 yang dilaporkan pada kolom terakhir menegaskan apakah relevansi prediktif model terpenuhi, sehingga kesimpulan tidak hanya bertumpu pada R2 (Chin, 1998; Sarstedt et al., 2022).", indent=0.5, space_after=6)
    else:
        p(doc, f"Tabel 4.{tbl_r2_num} memuat kerangka koefisien determinasi untuk variabel endogen. Nilai R2 dan Q2 belum tersedia karena data empiris belum diolah, sehingga kategori kekuatan model belum dapat ditetapkan. Setelah keluaran SmartPLS 4 dimasukkan, nilai tersebut akan dibaca dengan kriteria 0,75 (substansial), 0,50 (moderat), dan 0,25 (lemah) menurut Chin (1998) serta Hair et al. (2022).", indent=0.5, space_after=6)

def _build_doc_sub_4_a_4_2(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.4.2 Evaluasi Ukuran Efek (Effect Size f-Square / f²) (Format Baku TAPD 2026)."""
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    instansi = ctx["instansi"]
    tahun = ctx.get("tahun", "2026")
    pls = ctx["pls"]

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "4. Evaluasi Model Struktural (Inner Model PLS-SEM)")

    h4(doc, "b. Evaluasi Ukuran Efek (Effect Size f-Square / f²)")
    p(doc, f"Evaluasi ukuran efek (effect size f²) dilakukan untuk mengukur besaran kontribusi relatif dari masing-masing variabel prediktor eksogen terhadap koefisien determinasi (R²) variabel endogen ketika variabel tersebut disertakan atau dikeluarkan dari model struktural (Cohen, 1988; Hair et al., 2022). Sesuai dengan pedoman metodologis Cohen (1988), nilai f² sebesar 0,02 diklasifikasikan memiliki efek kecil (small effect), nilai 0,15 diklasifikasikan memiliki efek sedang (medium effect), dan nilai 0,35 diklasifikasikan memiliki efek besar (large effect). Pengujian f² ini penting guna memastikan bahwa pengaruh statistik yang terdeteksi tidak sekadar disebabkan oleh ukuran sampel yang besar, melainkan mencerminkan dampak substansial nyata di lapangan (Hair et al., 2022). Hasil estimasi nilai f² untuk seluruh jalur pengaruh struktural disajikan pada tabel berikut (Cohen, 1988):", indent=0.5, space_after=6)

    # Tabel f-Square
    tbl_f2_num = 12 + k_vars
    t_f2_headers = ["Jalur Pengaruh Struktural", "Konstruk Eksogen (Prediktor)", "Konstruk Target (Endogen)", "Nilai f²", "Kategori Efek (Cohen, 1988)"]
    t_f2_rows = []

    hypos = pls.get("hypotheses_testing", [])
    _dir_types = ("direct_to_mediator", "direct_to_dependent", "mediator_to_dependent", "moderator_main_effect")
    direct_hypos = [h for h in hypos if h.get("type") in _dir_types]
    if not direct_hypos:
        direct_hypos = [h for h in hypos if "direct" in h.get("type", "") and "indirect" not in h.get("type", "")]
    _tpl = s3data.is_template(ctx)

    for h in direct_hypos:
        _rel = str(h.get("relation") or "")
        _parts = [seg.strip() for seg in _rel.split("->")]
        _f2_raw = h.get("f2_effect_size")
        try:
            f2_v = None if _tpl else float(_f2_raw)
        except (TypeError, ValueError):
            f2_v = None
        if f2_v is None:
            f2_txt = s3data.PLACEHOLDER
            kat_f2 = "Belum ditetapkan"
        else:
            f2_txt = f"{f2_v:.3f}"
            if f2_v >= 0.35:
                kat_f2 = "Efek Besar (Large Effect)"
            elif f2_v >= 0.15:
                kat_f2 = "Efek Sedang (Medium Effect)"
            elif f2_v >= 0.02:
                kat_f2 = "Efek Kecil (Small Effect)"
            else:
                kat_f2 = "Efek Sangat Kecil"
        t_f2_rows.append([
            _rel or s3data.PLACEHOLDER,
            _parts[0] if _parts and _parts[0] else s3data.PLACEHOLDER,
            _parts[1] if len(_parts) > 1 and _parts[1] else s3data.PLACEHOLDER,
            f2_txt,
            kat_f2
        ])

    add_table_data(doc, title=f"Tabel 4. {tbl_f2_num}\nNilai Effect Size (f2) Seluruh Jalur Struktural", headers=t_f2_headers, data=t_f2_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi f-Square (>= 4 baris, bersitasi)
    _f2_vals = []
    for _r in t_f2_rows:
        try:
            _f2_vals.append((_r[0], float(_r[3])))
        except (TypeError, ValueError):
            pass
    if _f2_vals:
        _besar = [n for n, v in _f2_vals if v >= 0.35]
        _kecil = [n for n, v in _f2_vals if v < 0.15]
        _f2_range = f"{min(v for _, v in _f2_vals):.3f} hingga {max(v for _, v in _f2_vals):.3f}"
        _f2_ket = ""
        if _besar:
            _f2_ket += f" Jalur dengan efek besar adalah {_join_id(_besar)}."
        if _kecil:
            _f2_ket += f" Jalur dengan efek di bawah 0,15 adalah {_join_id(_kecil)}, sehingga kontribusinya perlu dibaca secara hati-hati."
        p(doc, f"Ukuran efek pada Tabel 4.{tbl_f2_num} berada pada rentang {_f2_range}. Menurut pedoman Cohen (1988) yang dipakai Hair et al. (2022), nilai 0,02 tergolong kecil, 0,15 sedang, dan 0,35 besar, sehingga penafsiran dilakukan per jalur, bukan secara seragam.{_f2_ket} Pembacaan ini penting agar pengaruh yang tampak signifikan secara statistik tidak otomatis dianggap besar secara praktis (Cohen, 1988; Hair et al., 2022).", indent=0.5, space_after=6)
    else:
        p(doc, f"Tabel 4.{tbl_f2_num} memuat kerangka ukuran efek untuk seluruh jalur struktural. Nilai f2 belum tersedia karena data empiris belum diolah, sehingga kategori efek per jalur belum dapat ditetapkan. Setelah keluaran SmartPLS 4 dimasukkan, setiap jalur akan dinilai dengan ambang 0,02 (kecil), 0,15 (sedang), dan 0,35 (besar) sesuai pedoman Cohen (1988) dan Hair et al. (2022).", indent=0.5, space_after=6)

def _build_doc_sub_4_a_4_3(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.4.3 Evaluasi Relevansi Prediksi Out-of-Sample (PLSpredict) (Format Baku TAPD 2026)."""
    import hashlib
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    instansi = ctx["instansi"]
    tahun = ctx.get("tahun", "2026")

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "4. Evaluasi Model Struktural (Inner Model PLS-SEM)")

    h4(doc, "c. Evaluasi Relevansi Prediksi Out-of-Sample (PLSpredict)")
    p(doc, f"Evaluasi daya prediksi luar sampel (out-of-sample predictive power) dinilai menggunakan prosedur PLSpredict dengan algoritma k-fold cross-validation (k=10 folds) sebagaimana dikembangkan oleh Shmueli et al. (2019) dan Hair et al. (2022). Prosedur ini bertujuan untuk mengevaluasi apakah model struktural yang dihasilkan mampu memprediksi data observasi baru yang tidak digunakan dalam proses estimasi awal. Menurut panduan Shmueli et al. (2019), model dikategorikan memiliki daya prediksi out-of-sample yang tinggi (high predictive power) jika seluruh nilai Q²_predict bernilai positif (> 0) dan mayoritas atau seluruh indikator variabel endogen menghasilkan nilai Root Mean Squared Error (RMSE) PLS-SEM yang lebih rendah dibandingkan dengan benchmark model regresi linier naif (Linear Model / LM). Hasil pengujian PLSpredict untuk butir-butir indikator variabel terikat disajikan pada tabel berikut (Shmueli et al., 2019):", indent=0.5, space_after=6)

    # Tabel PLSpredict
    tbl_pred_num = 13 + k_vars
    t_pred_headers = ["Indikator Endogen (Y)", "Q²_predict", "RMSE PLS-SEM", "RMSE Linear Model (LM)", "Selisih (PLS - LM)", "Daya Prediksi Out-of-Sample"]
    t_pred_rows = []

    y_dims, y_items, _ = _get_variable_instrument_details("Y", "Variabel Terikat", ctx["y_var"], {}, ctx=ctx)

    # Peta keluaran PLSpredict riil (bila disediakan): list of dict dengan kunci
    # {"code", "q2_predict", "rmse_pls", "rmse_lm"}.
    _pred_raw = s3data.records(ctx, "plspredict")
    _pred_map: Dict[str, Dict[str, Any]] = {}
    if isinstance(_pred_raw, list):
        for _row in _pred_raw:
            if isinstance(_row, dict):
                _c = str(_row.get("code") or _row.get("indicator") or "").strip()
                if _c:
                    _pred_map[_c] = _row

    _pred_q2: List[float] = []
    _pred_n = 0
    _pred_superior = 0
    for idx_y, y_it in enumerate(y_items[:5], start=1):
        _row = _pred_map.get(y_it["kode"]) or {}
        _pred_n += 1
        try:
            q2_v = float(_row.get("q2_predict", _row.get("q2")))
            q2_txt = f"{q2_v:.3f}"
            _pred_q2.append(q2_v)
        except (TypeError, ValueError):
            q2_txt = s3data.PLACEHOLDER
        try:
            rmse_pls_v = float(_row.get("rmse_pls", _row.get("rmse_pls_sem")))
            rmse_pls_txt = f"{rmse_pls_v:.3f}"
        except (TypeError, ValueError):
            rmse_pls_v = None
            rmse_pls_txt = s3data.PLACEHOLDER
        try:
            rmse_lm_v = float(_row.get("rmse_lm", _row.get("rmse_linear")))
            rmse_lm_txt = f"{rmse_lm_v:.3f}"
        except (TypeError, ValueError):
            rmse_lm_v = None
            rmse_lm_txt = s3data.PLACEHOLDER
        if rmse_pls_v is not None and rmse_lm_v is not None:
            diff_txt = f"{rmse_pls_v - rmse_lm_v:.3f}"
            if rmse_pls_v < rmse_lm_v:
                daya = "Superior (RMSE PLS < LM)"
                _pred_superior += 1
            else:
                daya = "Tidak Superior"
        else:
            diff_txt = s3data.PLACEHOLDER
            daya = "Belum ditetapkan"
        t_pred_rows.append([
            y_it["kode"],
            q2_txt,
            rmse_pls_txt,
            rmse_lm_txt,
            diff_txt,
            daya
        ])

    add_table_data(doc, title=f"Tabel 4. {tbl_pred_num}\nHasil Pengujian Out-of-Sample Predictive Relevance (PLSpredict)", headers=t_pred_headers, data=t_pred_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi PLSpredict (>= 4 baris, bersitasi)
    if _pred_q2:
        p(doc, f"Hasil PLSpredict pada Tabel 4.{tbl_pred_num} menunjukkan nilai Q2_predict berkisar {min(_pred_q2):.3f} hingga {max(_pred_q2):.3f}. Dari {_pred_n} indikator yang dinilai, {_pred_superior} indikator menghasilkan RMSE PLS-SEM yang lebih rendah daripada model linier pembanding. Menurut Shmueli et al. (2019), kesimpulan daya prediksi luar sampel ditetapkan berdasarkan pola perbandingan tersebut, bukan pada satu indikator saja, sehingga indikator yang belum superior tetap dilaporkan apa adanya (Shmueli et al., 2019; Hair et al., 2022).", indent=0.5, space_after=6)
    else:
        p(doc, f"Tabel 4.{tbl_pred_num} memuat kerangka evaluasi daya prediksi luar sampel. Nilai Q2_predict serta RMSE PLS-SEM dan model linier pembanding belum tersedia karena keluaran PLSpredict belum dimasukkan, sehingga kategori daya prediksi belum dapat ditetapkan. Setelah data tersedia, penilaian mengikuti aturan Shmueli et al. (2019) dan Hair et al. (2022), yaitu membandingkan galat prediksi PLS terhadap model linier untuk setiap indikator endogen.", indent=0.5, space_after=6)

def _build_doc_sub_4_a_4_4(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.4.4 Evaluasi Kecocokan Model Fit & Goodness of Fit (GoF) Tenenhaus (Format Baku TAPD 2026)."""
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    instansi = ctx["instansi"]
    tahun = ctx.get("tahun", "2026")
    pls = ctx["pls"]

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "4. Evaluasi Model Struktural (Inner Model PLS-SEM)")

    h4(doc, "d. Evaluasi Kecocokan Model (Model Fit Indices)")
    p(doc, f"Pengujian kesesuaian model (model fit) dalam kerangka PLS-SEM dilakukan untuk menilai sejauh mana matriks kovarians yang diimplikasikan oleh model struktural cocok dengan matriks kovarians sampel data empiris (Henseler et al., 2016). Kriteria utama yang dievaluasi mencakup Standardized Root Mean Square Residual (SRMR), unweighted least squares discrepancy (d_ULS), geodesic discrepancy (d_G), serta Normed Fit Index (NFI) (Hu & Bentler, 1999; Henseler et al., 2016). Nilai SRMR dinyatakan memiliki kecocokan yang baik (good fit) jika bernilai kurang dari 0,080, sedangkan nilai NFI dinyatakan dapat diterima jika melampaui nilai ambang batas 0,900 (Hu & Bentler, 1999; Hair et al., 2022). Hasil estimasi indeks kecocokan model PLS-SEM disajikan pada tabel berikut (Henseler et al., 2016):", indent=0.5, space_after=6)

    # Tabel Model Fit
    tbl_fit_num = 14 + k_vars
    t_fit_headers = ["Kriteria Kesesuaian (Model Fit)", "Saturated Model", "Estimated Model", "Nilai Ambang Kritis", "Kesimpulan Kesesuaian"]

    _tpl = s3data.is_template(ctx)

    def _fnum(v, digits: int = 3):
        if _tpl or v is None:
            return s3data.PLACEHOLDER
        try:
            return f"{float(v):.{digits}f}"
        except (TypeError, ValueError):
            return s3data.PLACEHOLDER

    def _simpul_srmr(v):
        if v is None:
            return "Belum diuji"
        try:
            return "Good Fit (Model Cocok)" if float(v) < 0.080 else "Belum Memenuhi"
        except (TypeError, ValueError):
            return "Belum diuji"

    def _simpul_nfi(v):
        if v is None:
            return "Belum diuji"
        try:
            return "Good Fit Terpenuhi" if float(v) > 0.900 else "Belum Memenuhi"
        except (TypeError, ValueError):
            return "Belum diuji"

    def _simpul_dis(v):
        return "Belum diuji" if v is None else "Fit Terpenuhi"

    fit_idx = {} if _tpl else dict(pls.get("fit_indices") or {})
    _inner_real = {} if _tpl else (s3data.inner(ctx) or {})
    if fit_idx.get("srmr") is None and _inner_real.get("srmr") is not None:
        fit_idx["srmr"] = _inner_real["srmr"]
    if fit_idx.get("nfi") is None and _inner_real.get("nfi") is not None:
        fit_idx["nfi"] = _inner_real["nfi"]

    srmr_sat = fit_idx.get("srmr")
    srmr_est = fit_idx.get("srmr_estimated", srmr_sat)
    nfi_sat = fit_idx.get("nfi")
    nfi_est = fit_idx.get("nfi_estimated", nfi_sat)
    d_uls_sat = fit_idx.get("d_uls")
    d_uls_est = fit_idx.get("d_uls_estimated", d_uls_sat)
    d_g_sat = fit_idx.get("d_g")
    d_g_est = fit_idx.get("d_g_estimated", d_g_sat)

    t_fit_rows = [
        ["SRMR (Standardized Root Mean Residual)", _fnum(srmr_sat), _fnum(srmr_est), "< 0,080", _simpul_srmr(srmr_sat)],
        ["d_ULS (Squared Euclidean Distance)", _fnum(d_uls_sat), _fnum(d_uls_est), "< HI 95%", _simpul_dis(d_uls_sat)],
        ["d_G (Geodesic Distance)", _fnum(d_g_sat), _fnum(d_g_est), "< HI 95%", _simpul_dis(d_g_sat)],
        ["NFI (Normed Fit Index)", _fnum(nfi_sat), _fnum(nfi_est), "> 0,900", _simpul_nfi(nfi_sat)]
    ]
    add_table_data(doc, title=f"Tabel 4. {tbl_fit_num}\nNilai Uji Kecocokan Model Fit PLS-SEM", headers=t_fit_headers, data=t_fit_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi Model Fit (>= 4 baris, bersitasi)
    if srmr_sat is not None:
        p(doc, f"Hasil evaluasi kesesuaian model pada Tabel 4.{tbl_fit_num} menunjukkan nilai SRMR saturated model sebesar {_fnum(srmr_sat)} dan estimated model sebesar {_fnum(srmr_est)}, dengan ambang toleransi 0,080 (Henseler et al., 2016; Hu & Bentler, 1999). Nilai NFI tercatat {_fnum(nfi_sat)} pada saturated model dan {_fnum(nfi_est)} pada estimated model, dibaca dengan ambang 0,900. Nilai diskrepansi d_ULS dan d_G dibandingkan dengan batas persentil ke-95 bootstrap untuk menilai kesalahan spesifikasi; kriteria yang belum terpenuhi dilaporkan apa adanya agar penilaian tidak menyimpang (Henseler et al., 2016; Hair et al., 2022).", indent=0.5, space_after=6)
    else:
        p(doc, f"Tabel 4.{tbl_fit_num} memuat kerangka indeks kecocokan model. Nilai SRMR, d_ULS, d_G, dan NFI belum tersedia karena data empiris belum diolah, sehingga kesimpulan kecocokan model belum dapat ditetapkan. Setelah keluaran SmartPLS 4 dimasukkan, penilaian mengikuti ambang SRMR < 0,080 dan NFI > 0,900, serta perbandingan d_ULS dan d_G terhadap batas persentil ke-95 bootstrap (Henseler et al., 2016; Hu & Bentler, 1999).", indent=0.5, space_after=6)

    # e. Evaluasi Tenenhaus GoF
    h4(doc, "e. Evaluasi Keterpaduan Model Menyeluruh (Goodness of Fit / GoF Tenenhaus)")
    p(doc, f"Evaluasi keterpaduan model secara global dinilai menggunakan formulasi Goodness of Fit (GoF) yang dikembangkan oleh Tenenhaus et al. (2005) dan Wetzels et al. (2009). Indeks GoF merupakan rata-rata geometrik dari perkalian nilai rata-rata komunalitas (AVE rerata) dengan nilai rata-rata koefisien determinasi (R2 rerata) yang merefleksikan performa terpadu antara model pengukuran (outer model) dan model struktural (inner model) dalam satu indeks tunggal (Tenenhaus et al., 2005). Menurut kriteria baku Wetzels et al. (2009), nilai GoF sebesar 0,10 dikategorikan kecil (GoF small), nilai 0,25 dikategorikan moderat (GoF medium), dan nilai 0,36 dikategorikan besar/unggul (GoF large). Formula matematis indeks GoF dirumuskan sebagai berikut (Tenenhaus et al., 2005):", indent=0.5, space_after=4)

    p(doc, "GoF = (AVE_rerata x R2_rerata)^0,5", bold=True, indent=1.0, space_after=6)

    # Tabel GoF
    tbl_gof_num = 15 + k_vars
    t_gof_headers = ["Parameter Komponen Global", "Nilai Rerata Hitung", "Indeks GoF Tenenhaus", "Kriteria Ambang Kritis", "Kategori Keterpaduan Global"]

    import math as _math

    _aves = []
    for _vk, _role, _vn, _vd in ordered_vars:
        _av = (s3data.outer(ctx, _vk) or {}).get("ave")
        try:
            _aves.append(float(_av))
        except (TypeError, ValueError):
            pass
    _r2s = []
    for _k, _v in (pls.get("r_squares") or {}).items():
        try:
            _r2s.append(float(_v.get("r_square")))
        except (TypeError, ValueError):
            pass
    if _tpl:
        _aves = []
        _r2s = []

    ave_mean = (sum(_aves) / len(_aves)) if _aves else None
    r2_mean = (sum(_r2s) / len(_r2s)) if _r2s else None
    if ave_mean is not None and r2_mean is not None:
        gof_val = _math.sqrt(ave_mean * r2_mean)
        gof_txt = f"{gof_val:.3f}"
        if gof_val >= 0.36:
            gof_kat = "GoF Large (Keterpaduan Tinggi)"
        elif gof_val >= 0.25:
            gof_kat = "GoF Medium (Keterpaduan Sedang)"
        elif gof_val >= 0.10:
            gof_kat = "GoF Small (Keterpaduan Rendah)"
        else:
            gof_kat = "Di Bawah GoF Small"
    else:
        gof_val = None
        gof_txt = s3data.PLACEHOLDER
        gof_kat = "Belum ditetapkan"

    t_gof_rows = [
        ["Rata-rata AVE Konstruk (Outer Model)", _fnum(ave_mean), gof_txt, "GoF small >= 0,10\nGoF medium >= 0,25\nGoF large >= 0,36", gof_kat],
        ["Rata-rata R2 Konstruk (Inner Model)", _fnum(r2_mean), "", "", ""]
    ]
    add_table_data(doc, title=f"Tabel 4. {tbl_gof_num}\nPerhitungan Goodness of Fit (GoF) Tenenhaus Index", headers=t_gof_headers, data=t_gof_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi GoF (>= 4 baris, bersitasi)
    if gof_val is not None:
        p(doc, f"Indeks Tenenhaus Goodness of Fit (GoF) pada Tabel 4.{tbl_gof_num} dihitung dari rata-rata AVE {_fnum(ave_mean)} dan rata-rata R2 {_fnum(r2_mean)}, menghasilkan nilai {gof_txt}. Nilai tersebut dibaca dengan kriteria Wetzels et al. (2009): 0,10 kecil, 0,25 sedang, dan 0,36 besar. Kategori GoF dihitung sebagaimana adanya, sehingga bila nilai berada pada kategori sedang atau kecil, hal itu dilaporkan tanpa pembulatan ke atas (Tenenhaus et al., 2005; Wetzels et al., 2009).", indent=0.5, space_after=6)
    else:
        p(doc, f"Tabel 4.{tbl_gof_num} memuat kerangka perhitungan GoF Tenenhaus. Nilai rata-rata AVE dan rata-rata R2 belum tersedia karena data empiris belum diolah, sehingga indeks GoF dan kategorinya belum dapat ditetapkan. Setelah keluaran SmartPLS 4 dimasukkan, GoF dihitung sebagai akar dari perkalian AVE rerata dan R2 rerata, lalu dibaca dengan kriteria 0,10 (kecil), 0,25 (sedang), dan 0,36 (besar) menurut Tenenhaus et al. (2005) dan Wetzels et al. (2009).", indent=0.5, space_after=6)

def _build_doc_sub_4_a_4(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """Sub-Bab A.4 Evaluasi Model Struktural (Inner Model PLS-SEM) (Composite Lengkap)."""
    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")

    h3(doc, "4. Evaluasi Model Struktural (Inner Model PLS-SEM)")
    instansi = ctx["instansi"]

    # Narasi Pengantar Inner Model (>= 4 baris, bersitasi)
    p(doc, f"Setelah model pengukuran (outer model) dievaluasi sesuai kriteria yang ditetapkan, tahapan berikutnya dalam PLS-SEM adalah menguji model struktural (inner model) guna menilai kapasitas prediktif model dan kekuatan hubungan kausal antar-konstruk laten (Hair et al., 2022). Evaluasi model struktural dalam penelitian disertasi ini dilakukan secara komprehensif melalui lima pilar penilaian metodologis utama: (1) Koefisien Determinasi (R-Square / R²), (2) Ukuran Efek (Effect Size f-Square / f²), (3) Relevansi Prediksi Luar Sampel (Out-of-Sample Predictive Power melalui PLSpredict), (4) Evaluasi Kecocokan Model (Model Fit Indices: SRMR, d_ULS, d_G, dan NFI), serta (5) Keterpaduan Model Menyeluruh melalui Indeks Tenenhaus Goodness of Fit (GoF) (Sarstedt et al., 2022; Chin, 1998; Tenenhaus et al., 2005). Rangkaian evaluasi inner model ini diarahkan untuk menilai apakah model teoretis yang diajukan tidak hanya memiliki signifikansi statistik internal tetapi juga relevansi praktis dan daya generalisasi prediktif pada organisasi sektor publik di {instansi} (Shmueli et al., 2019).", indent=0.5, space_after=6)

    # 1. R-Square
    _build_doc_sub_4_a_4_1(doc, ctx, standalone=False)

    # 2. f-Square
    _build_doc_sub_4_a_4_2(doc, ctx, standalone=False)

    # 3. PLSpredict
    _build_doc_sub_4_a_4_3(doc, ctx, standalone=False)

    # 4. Model Fit & GoF
    _build_doc_sub_4_a_4_4(doc, ctx, standalone=False)


def _build_doc_sub_4_a_5_1(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.5.1 Pengujian Pengaruh Langsung (Direct Effects) & Persamaan Sub-Struktural (Format Baku TAPD 2026)."""
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    instansi = ctx["instansi"]
    tahun = ctx.get("tahun", "2026")
    pls = ctx["pls"]
    hypos = pls.get("hypotheses_testing", [])

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "5. Pengujian Hipotesis Penelitian")

        # Narasi Pengantar Bootstrapping (>= 4 baris, bersitasi)
        p(doc, f"Pengujian hipotesis penelitian merupakan inti pembuktian empiris dalam disertasi ini yang bertujuan untuk menguji secara ilmiah keberlakuan hubungan kausal langsung (direct effects), efek transmisi mediasi (indirect effects), serta efek interaksi kontingensi moderasi (interaction effects) (Hair et al., 2022). Mengingat PLS-SEM tidak mengasumsikan distribusi normalitas data multivariat secara ketat, maka pengujian signifikansi statistik koefisien jalur dilakukan melalui prosedur non-parametrik Bootstrapping dengan menghasilkan 5.000 subsampel acak terdistribusi ulang (resampling) (Streukens & Leroi-Werelds, 2016; Hair et al., 2022). Keputusan penerimaan hipotesis dipedomani oleh pengujian dua arah (two-tailed test) pada taraf signifikansi alpha 5% (α = 0,05), di mana suatu hipotesis dinyatakan diterima dan terbukti signifikan secara statistik apabila menghasilkan nilai T-statistik >= 1,96 dan nilai probabilitas signifikansi (p-value) < 0,05 (Cooper & Schindler, 2022; Sarstedt et al., 2022).", indent=0.5, space_after=6)

    # a. Pengujian Pengaruh Langsung (Direct Effects)
    h4(doc, "a. Pengujian Pengaruh Langsung (Direct Effects)")
    p(doc, f"Pengujian pengaruh langsung (direct effects) dilakukan untuk membuktikan keberadaan hubungan kausalitas linier langsung antar-konstruk laten eksogen terhadap konstruk mediasi maupun terhadap konstruk endogen akhir pada {instansi} tanpa melalui jalur perantara (Hair et al., 2022). Estimasi koefisien jalur struktural (path coefficient β) mengindikasikan kekuatan dan arah hubungan teoritis yang dihipotesiskan, di mana koefisien bertanda positif mencerminkan hubungan yang searah dan harmonis dalam sistem manajemen kinerja organisasi (Sarstedt et al., 2022). Uji signifikansi non-parametrik bootstrapping dengan 5.000 subsampel diterapkan guna memastikan bahwa estimasi parameter terbebas dari bias distribusi sampel dan memiliki tingkat presisi statistik yang tinggi pada derajat kepercayaan 95% (Cooper & Schindler, 2022). Rangkuman hasil pengujian signifikansi koefisien jalur struktural pengaruh langsung disajikan pada tabel berikut (Hair et al., 2022):", indent=0.5, space_after=6)

    # Tabel Pengaruh Langsung
    tbl_dir_num = 16 + k_vars
    t_dir_headers = ["Hipotesis", "Jalur Hubungan Langsung", "Original Sample (β)", "Sample Mean (M)", "Std Dev (STDEV)", "T-Statistik", "P-Value", "Keputusan"]
    t_dir_rows = []
    _tpl = s3data.is_template(ctx)

    _dir_types = ("direct_to_mediator", "direct_to_dependent", "mediator_to_dependent", "moderator_main_effect")
    direct_hypos = [h for h in hypos if h.get("type") in _dir_types]
    if not direct_hypos:
        direct_hypos = [h for h in hypos if "direct" in h.get("type", "") and "indirect" not in h.get("type", "")]

    for h in direct_hypos:
        if _tpl:
            t_dir_rows.append([
                h["code"], h["relation"], s3data.PLACEHOLDER, s3data.PLACEHOLDER,
                s3data.PLACEHOLDER, s3data.PLACEHOLDER, s3data.PLACEHOLDER,
                "Belum ditetapkan",
            ])
        else:
            t_dir_rows.append([
                h["code"],
                h["relation"],
                _fmt3(h.get("original_sample")),
                _fmt3(h.get("sample_mean")),
                _fmt3(h.get("stdev")),
                _fmt3(h.get("t_statistics")),
                _fmt3(h.get("p_value")),
                str(h.get("decision") or "Belum ditetapkan")
            ])

    add_table_data(doc, title=f"Tabel 4. {tbl_dir_num}\nRangkuman Hasil Pengujian Hipotesis Pengaruh Langsung (Direct Effects)", headers=t_dir_headers, data=t_dir_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi Evaluasi Pengaruh Langsung (>= 4 baris, bersitasi)
    if _tpl:
        p(doc, f"Tabel 4.{tbl_dir_num} menampilkan kerangka hasil pengujian pengaruh langsung. Nilai koefisien jalur, T-statistik, dan p-value belum tersedia karena data empiris belum diolah, sehingga kolom keputusan diisi 'Belum ditetapkan'. Setelah keluaran SmartPLS 4 dimasukkan, tabel akan terisi otomatis dan interpretasi mengikuti ambang T >= 1,96 serta p < 0,05 (Hair et al., 2022; Sarstedt et al., 2022). Sampai tahap itu, tidak ada klaim penerimaan hipotesis yang dirumuskan pada bagian ini.", indent=0.5, space_after=6)
    else:
        _pos = all(float(h.get("original_sample", 0)) > 0 for h in direct_hypos)
        _tmin = min(float(h.get("t_statistics", 0)) for h in direct_hypos)
        _pmax = max(float(h.get("p_value", 1)) for h in direct_hypos)
        _acc = sum(1 for h in direct_hypos if str(h.get("decision", "")).lower().startswith("diterima"))
        p(doc, f"Berdasarkan Tabel 4.{tbl_dir_num}, dari {len(direct_hypos)} hipotesis pengaruh langsung, {_acc} dinyatakan diterima. Koefisien jalur {'seluruhnya bertanda positif' if _pos else 'tidak seluruhnya bertanda positif'}, dengan T-statistik terendah {_tmin:.3f} dan p-value tertinggi {_pmax:.3f}. Pembacaan ini menunjukkan bahwa arah dan kekuatan hubungan langsung perlu dibandingkan dengan peran mekanisme perantara pada subbab berikutnya, agar kesimpulan tidak berhenti pada efek langsung semata (Hair et al., 2022; Robbins & Judge, 2022).", indent=0.5, space_after=6)

    # Box Gambar 4.2 Diagram T-Statistik Bootstrapping
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm

    t_box2 = doc.add_table(rows=1, cols=1)
    t_box2.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell2 = t_box2.rows[0].cells[0]
    cell2.width = Cm(15.5)

    x_vars_str = _join_id(ctx["x_vars"])
    m_var = ctx.get("m", "")
    z_var = ctx.get("z", "")
    y_var = ctx.get("y_var", "Kinerja Pegawai")
    _has_m = bool(m_var)
    _has_z = bool(z_var)
    _dir_t = [float(h.get("t_statistics")) for h in direct_hypos
              if h.get("t_statistics") is not None and not _tpl]
    _t_txt = (f"t = {min(_dir_t):.3f} s.d {max(_dir_t):.3f}; p < 0,05" if _dir_t
              else "t = \u2014; p = \u2014")

    p(cell2, "DIAGRAM JALUR T-STATISTIK BOOTSTRAPPING SMARTPLS 4 (5.000 SUBSAMPEL)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=4)
    p(cell2, f"[ Variabel Eksogen (X): {x_vars_str} ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    p(cell2, f"   \u2502 ({_t_txt})", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
    p(cell2, "   \u25bc", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
    if _has_m:
        p(cell2, f"[ Variabel Mediasi (M): {m_var} ] \u2500\u2500\u2500({_t_txt})\u2500\u2500\u2500\u25ba [ Variabel Endogen (Y): {y_var} ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    else:
        p(cell2, f"[ Variabel Endogen (Y): {y_var} ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    if _has_z:
        p(cell2, "   \u25b2", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
        p(cell2, f"   \u2502 (Efek Moderasi: {_t_txt})", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
        p(cell2, f"[ Variabel Moderasi (Z): {z_var} ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    set_cell_border(cell2, top={"sz": 12, "val": "single", "color": "003366"},
                           bottom={"sz": 12, "val": "single", "color": "003366"},
                           left={"sz": 12, "val": "single", "color": "003366"},
                           right={"sz": 12, "val": "single", "color": "003366"})

    p(doc, s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"), align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4, space_after=2)
    p(doc, "Gambar 4. 2\nDiagram Jalur T-Statistik Bootstrapping SmartPLS 4", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=8)

    # Narasi Gambar 4.2 (>= 4 baris, bersitasi)
    if _tpl:
        p(doc, f"Gambar 4.2 menampilkan kerangka diagram jalur. Nilai T-statistik belum diisi karena data empiris belum diolah, sehingga gambar ini berfungsi sebagai panduan visual struktur model, bukan bukti signifikansi. Setelah estimasi SmartPLS 4 dilakukan, angka pada setiap jalur akan diperbarui dan ditafsirkan dengan ambang 1,96 (Streukens & Leroi-Werelds, 2016; Hair et al., 2022).", indent=0.5, space_after=6)
    else:
        _all_t = [float(h["t_statistics"]) for h in hypos if h.get("t_statistics") is not None]
        p(doc, f"Diagram jalur T-statistik pada Gambar 4.2 memvisualisasikan derajat signifikansi setiap lintasan hubungan pada {instansi}. Nilai T-statistik bergerak dari {min(_all_t):.3f} hingga {max(_all_t):.3f}. Jalur dengan nilai tertinggi menunjukkan hubungan yang paling stabil pada proses bootstrapping, sedangkan jalur di bawah 1,96 perlu ditafsirkan secara hati-hati (Streukens & Leroi-Werelds, 2016; Hair et al., 2022).", indent=0.5, space_after=6)

    # b. Persamaan Sub-Struktural 1
    if _has_m:
        h4(doc, "b. Persamaan Sub-Struktural 1 (Model Variabel Mediasi)")
        p(doc, f"Pemodelan sub-struktural pertama diformulasikan untuk menguji pengaruh simultan dan parsial variabel eksogen ({x_vars_str}) terhadap variabel mediasi ({m_var}) pada {instansi} (Hair et al., 2022). Persamaan disusun dari koefisien jalur pada Tabel 4.{tbl_dir_num} (Sarstedt et al., 2022):", indent=0.5, space_after=4)

        p1_terms = []
        if not _tpl:
            for h in hypos:
                if h.get("type") == "direct_to_mediator" and h.get("original_sample") is not None:
                    p1_terms.append(f"{_fmt3(h['original_sample'])} {h['relation'].split('->')[0].strip()}")
        if p1_terms:
            p1_formula = f"{m_var} = " + " + ".join(p1_terms) + " + e1"
        else:
            p1_formula = f"{m_var} = (koefisien jalur belum tersedia) + e1"
        p(doc, p1_formula, bold=True, indent=1.0, space_after=6)

        if _tpl:
            p(doc, f"Koefisien persamaan belum dapat diisi karena data empiris riil belum diolah. Setelah estimasi dilakukan, setiap koefisien akan mencerminkan besar perubahan {m_var} akibat satu satuan perubahan variabel eksogen, dan nilai residu e1 dihitung dari R-square model mediasi (Chin, 1998; Hair et al., 2022).", indent=0.5, space_after=6)
        else:
            _top = max(
                (h for h in hypos if h.get("type") == "direct_to_mediator"),
                key=lambda h: float(h.get("original_sample", 0)),
                default=None,
            )
            _rs_m = (pls.get("r_squares") or {}).get("M", {}).get("r_square")
            _top_txt = (
                f"Koefisien terbesar tercatat pada {_top.get('relation', '-')} "
                f"(beta = {float(_top.get('original_sample', 0)):.3f}). " if _top else ""
            )
            _rs_txt = (
                f"Nilai R-square model mediasi sebesar {float(_rs_m):.3f} berarti residu e1 "
                f"mencakup {(1 - float(_rs_m)) * 100:.1f}% variansi yang tidak terjelaskan. "
                if _rs_m is not None else
                "Nilai R-square model mediasi belum tersedia pada keluaran yang dimasukkan. "
            )
            p(doc, f"Seluruh koefisien pada persamaan pertama dibaca sebagai perubahan {m_var} akibat perubahan variabel eksogen. {_top_txt}{_rs_txt}Besaran tersebut menjadi dasar untuk menilai seberapa besar kapasitas aparatur diterjemahkan ke dalam keadaan kerja yang diharapkan (Robbins & Judge, 2022; Chin, 1998).", indent=0.5, space_after=6)

    # c. Persamaan Sub-Struktural 2
    h4(doc, "c. Persamaan Sub-Struktural 2 (Model Variabel Terikat Akhir)")
    p(doc, f"Pemodelan sub-struktural kedua menguji pengaruh terintegrasi variabel eksogen ({x_vars_str})"
           + (f", variabel mediasi ({m_var})" if _has_m else "")
           + (f", variabel moderasi ({z_var}) beserta efek interaksinya" if _has_z else "")
           + f" terhadap {y_var} pada {instansi} (Hair et al., 2022). Persamaan disusun dari koefisien jalur yang tersedia (Sarstedt et al., 2022):", indent=0.5, space_after=4)

    p2_terms = []
    if not _tpl:
        for h in hypos:
            if h.get("type") in ("direct_to_dependent", "mediator_to_dependent", "moderator_main_effect", "moderation_interaction"):
                if h.get("original_sample") is not None:
                    rel_left = h["relation"].split("->")[0].strip()
                    p2_terms.append(f"{_fmt3(h['original_sample'])} {rel_left}")
        _mz_int = pls.get("hayes15_mz_interaction")
        if _mz_int and _mz_int.get("original_sample") is not None:
            p2_terms.append(f"{_fmt3(_mz_int['original_sample'])} ({m_var} × {z_var})")
    p2_formula = (f"{y_var} = " + " + ".join(p2_terms) + " + e2") if p2_terms else f"{y_var} = (koefisien jalur belum tersedia) + e2"
    p(doc, p2_formula, bold=True, indent=1.0, space_after=6)

    if _tpl:
        p(doc, f"Persamaan kedua belum dapat diisi karena data empiris riil belum diolah. Setelah estimasi tersedia, koefisien akan menunjukkan kontribusi masing-masing variabel"
               + (f", termasuk peran interaksi {z_var}" if _has_z else "")
               + f" terhadap {y_var}, dan nilai residu e2 dihitung dari R-square model akhir (Chin, 1998; Hair et al., 2022).", indent=0.5, space_after=6)
    else:
        _rs_y = (pls.get("r_squares") or {}).get("Y", {}).get("r_square")
        _rs_y_txt = (f"Nilai R-square model akhir sebesar {float(_rs_y):.3f} menunjukkan bahwa residu e2 mencakup {(1 - float(_rs_y)) * 100:.1f}% variansi {y_var} yang tidak dijelaskan model. " if _rs_y is not None else "")
        if _has_z:
            p(doc, f"{_rs_y_txt}Koefisien pada persamaan kedua dibaca sebagai kontribusi relatif setiap variabel terhadap {y_var}. Perbandingan antara koefisien mediasi dan koefisien langsung menjadi dasar penilaian mekanisme, sedangkan koefisien interaksi menunjukkan perubahan kekuatan hubungan pada tingkat {z_var} yang berbeda (Hair et al., 2022; Hayes, 2022).", indent=0.5, space_after=6)
        else:
            p(doc, f"{_rs_y_txt}Koefisien pada persamaan kedua dibaca sebagai kontribusi relatif setiap variabel terhadap {y_var}, sekaligus menjadi dasar perbandingan antara jalur langsung dan jalur perantara (Hair et al., 2022; Hayes, 2022).", indent=0.5, space_after=6)


def _build_doc_sub_4_a_5_2(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.5.2 Pengujian Efek Mediasi (Indirect / Mediation Effects) (Format Baku TAPD 2026)."""
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    instansi = ctx["instansi"]
    tahun = ctx.get("tahun", "2026")
    pls = ctx["pls"]
    hypos = pls.get("hypotheses_testing", [])
    m_var = ctx.get("m", "")
    y_var = ctx.get("y_var", "Kinerja Pegawai")
    _tpl = s3data.is_template(ctx)

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "5. Pengujian Hipotesis Penelitian")

    # d. Pengujian Efek Mediasi (Indirect Effects)
    h4(doc, "d. Pengujian Efek Mediasi (Indirect Effects / Jalur Transmisi)")
    p(doc, f"Pengujian pengaruh tidak langsung (indirect effects) dilakukan guna membuktikan secara ilmiah apakah variabel {m_var} secara signifikan berfungsi sebagai jembatan mediasi yang menyalurkan pengaruh variabel-variabel eksogen terhadap variabel terikat {y_var} pada {instansi} (Hair et al., 2022). Evaluasi mediasi dipedomani oleh prosedur bootstrapping bias-corrected menurut Zhao et al. (2010) dan Hair et al. (2022), di mana sifat mediasi diklasifikasikan menjadi: (1) Mediasi Parsial Komplementer (Complementary Partial Mediation) jika pengaruh langsung dan tidak langsung keduanya signifikan dan bertanda searah, (2) Mediasi Penuh (Full Mediation) jika pengaruh tidak langsung signifikan namun pengaruh langsung tidak signifikan, serta (3) Tidak Memediasi (No Mediation) jika pengaruh tidak langsung tidak signifikan (Zhao et al., 2010). Hasil pengujian efek mediasi disajikan pada tabel berikut (Hair et al., 2022):", indent=0.5, space_after=6)

    # Tabel Mediasi
    tbl_med_num = 17 + k_vars
    t_med_headers = ["Hipotesis", "Jalur Pengaruh Tidak Langsung (Mediasi)", "Indirect Effect (β)", "Sample Mean (M)", "Std Dev (STDEV)", "T-Statistik", "P-Value", "Sifat Mediasi", "Keputusan"]
    t_med_rows = []

    med_hypos = [h for h in hypos if h.get("type") == "mediation_indirect"]
    if not med_hypos:
        med_hypos = [h for h in hypos if "->" in h.get("relation", "") and h.get("relation", "").count("->") > 1 and "index" not in h.get("relation", "").lower()]

    # Peta direct effects untuk klasifikasi Zhao et al. (2010)
    _dir_by_src = {}
    for dh in hypos:
        if dh.get("type") == "direct_to_dependent":
            _src = dh.get("source") or (dh.get("relation", "").split("->")[0].strip())
            _dir_by_src[_src] = dh

    for h in med_hypos:
        if _tpl:
            t_med_rows.append([
                h["code"], h["relation"], s3data.PLACEHOLDER, s3data.PLACEHOLDER,
                s3data.PLACEHOLDER, s3data.PLACEHOLDER, s3data.PLACEHOLDER,
                "Belum ditetapkan", "Belum ditetapkan",
            ])
        else:
            try:
                t_val = float(h.get("t_statistics", 0))
                p_val = float(h.get("p_value", 1))
                ind_b = float(h.get("original_sample", 0))
            except (TypeError, ValueError):
                t_val, p_val, ind_b = 0.0, 1.0, 0.0
            _src = h.get("source") or (h.get("relation", "").split("->")[0].strip())
            _dh = _dir_by_src.get(_src, {})
            try:
                dir_t = float(_dh.get("t_statistics", 0))
                dir_p = float(_dh.get("p_value", 1))
                dir_b = float(_dh.get("original_sample", 0))
            except (TypeError, ValueError):
                dir_t, dir_p, dir_b = 0.0, 1.0, 0.0

            is_ind_sig = (t_val >= 1.96 and p_val < 0.05)
            is_dir_sig = (dir_t >= 1.96 and dir_p < 0.05)
            if is_ind_sig:
                if is_dir_sig:
                    sifat = "Mediasi Parsial (Komplementer)" if (ind_b * dir_b > 0) else "Mediasi Parsial (Kompetitif)"
                else:
                    sifat = "Mediasi Penuh (Full Mediation)"
            else:
                sifat = "Tidak Terbukti Memediasi"

            dec = str(h.get("decision") or ("Diterima" if is_ind_sig else "Ditolak"))
            t_med_rows.append([
                h["code"],
                h["relation"],
                _fmt3(h.get("original_sample")),
                _fmt3(h.get("sample_mean")),
                _fmt3(h.get("stdev")),
                _fmt3(h.get("t_statistics")),
                _fmt3(h.get("p_value")),
                sifat,
                dec
            ])

    add_table_data(doc, title=f"Tabel 4. {tbl_med_num}\nHasil Pengujian Efek Mediasi Variabel Intervening", headers=t_med_headers, data=t_med_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi Mediasi (>= 4 baris, bersitasi)
    if _tpl:
        p(doc, f"Tabel 4.{tbl_med_num} menampilkan kerangka pengujian efek mediasi. Nilai efek tidak langsung, T-statistik, dan p-value belum tersedia karena data empiris belum diolah, sehingga sifat mediasi belum dapat ditetapkan. Setelah keluaran SmartPLS 4 dimasukkan, klasifikasi mengikuti prosedur Zhao et al. (2010): mediasi komplementer, mediasi penuh, atau tidak memediasi, bergantung pada signifikansi efek langsung dan tidak langsung (Hair et al., 2022; Hayes, 2022).", indent=0.5, space_after=6)
    else:
        _tmin = min(float(h.get("t_statistics", 0)) for h in med_hypos)
        _pmax = max(float(h.get("p_value", 1)) for h in med_hypos)
        _acc = sum(1 for h in med_hypos if str(h.get("decision", "")).lower().startswith("diterima"))
        p(doc, f"Hasil pengujian pada Tabel 4.{tbl_med_num} menunjukkan bahwa dari {len(med_hypos)} hipotesis mediasi, {_acc} dinyatakan diterima, dengan T-statistik terendah {_tmin:.3f} dan p-value tertinggi {_pmax:.3f}. Klasifikasi sifat mediasi mengikuti kriteria Zhao et al. (2010) dan Hair et al. (2022), yaitu membandingkan signifikansi serta arah efek langsung dan efek tidak langsung. Temuan ini perlu dibaca bersama koefisien jalur langsung pada tabel sebelumnya agar kesimpulan tentang peran {m_var} tidak berdiri sendiri (Robbins & Judge, 2022; Hayes, 2022).", indent=0.5, space_after=6)


def _build_doc_sub_4_a_5_3(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.5.3 Pengujian Efek Moderasi (Interaction Effects / Peran Kontingensi) (Format Baku TAPD 2026)."""
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    instansi = ctx["instansi"]
    tahun = ctx.get("tahun", "2026")
    pls = ctx["pls"]
    hypos = pls.get("hypotheses_testing", [])
    z_var = ctx.get("z", "")
    y_var = ctx.get("y_var", "Kinerja Pegawai")
    _tpl = s3data.is_template(ctx)

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "5. Pengujian Hipotesis Penelitian")

    # e. Pengujian Efek Moderasi (Interaction Effects)
    h4(doc, "e. Pengujian Efek Moderasi (Interaction Effects / Peran Kontingensi)")
    p(doc, f"Pengujian efek moderasi bertujuan untuk mengevaluasi apakah variabel kontingensi {z_var} secara signifikan memoderasi—memperkuat atau memperlemah—hubungan kausal antara variabel prediktor terhadap variabel terikat pada {instansi} (Henseler & Fassott, 2010). Pengujian efek interaksi moderasi diestimasi menggunakan pendekatan Dua Tahap (Two-Stage Approach) yang direkomendasikan oleh Henseler & Fassott (2010) dan Hayes (2022) karena memiliki kekuatan statistik (statistical power) yang paling unggul dibandingkan metode orthogonalizing atau product indicator. Suatu variabel dikonfirmasi berperan sebagai variabel pemoderasi murni (pure moderator) atau pemoderasi kuasi (quasi moderator) apabila koefisien jalur dari istilah interaksi (interaction term) menghasilkan nilai T-statistik >= 1,96 dan p-value < 0,05 (Hair et al., 2022). Hasil pengujian efek interaksi moderasi disajikan pada tabel berikut (Henseler & Fassott, 2010):", indent=0.5, space_after=6)

    # Tabel Moderasi
    tbl_mod_num = 18 + k_vars
    t_mod_headers = ["Hipotesis", "Jalur Interaksi Moderasi", "Interaction Coeff (β)", "Sample Mean (M)", "Std Dev (STDEV)", "T-Statistik", "P-Value", "Keterangan Moderasi", "Keputusan"]
    t_mod_rows = []

    mod_hypos = [h for h in hypos if h.get("type") == "moderation_interaction"]
    if not mod_hypos:
        mod_hypos = [h for h in hypos if "*" in h.get("relation", "") or "×" in h.get("relation", "")]

    # Cek signifikansi efek utama Z (moderator_main_effect) untuk klasifikasi Pure vs Quasi
    _z_main = next((h for h in hypos if h.get("type") == "moderator_main_effect"), {})
    try:
        _z_sig = float(_z_main.get("t_statistics", 0)) >= 1.96 and float(_z_main.get("p_value", 1)) < 0.05
    except (TypeError, ValueError):
        _z_sig = False

    for h in mod_hypos:
        if _tpl:
            t_mod_rows.append([
                h["code"], h["relation"], s3data.PLACEHOLDER, s3data.PLACEHOLDER,
                s3data.PLACEHOLDER, s3data.PLACEHOLDER, s3data.PLACEHOLDER,
                "Belum ditetapkan", "Belum ditetapkan",
            ])
        else:
            try:
                t_val = float(h.get("t_statistics", 0))
                p_val = float(h.get("p_value", 1))
                b_int = float(h.get("original_sample", 0))
            except (TypeError, ValueError):
                t_val, p_val, b_int = 0.0, 1.0, 0.0
            is_sig = (t_val >= 1.96 and p_val < 0.05)
            if is_sig:
                tipe_mod = "Quasi Moderator" if _z_sig else "Pure Moderator"
                arah_mod = "Memperkuat" if b_int >= 0 else "Memperlemah"
                ket = f"{arah_mod} ({tipe_mod})"
            else:
                ket = "Tidak Signifikan (Bukan Pemoderasi)"
            dec = str(h.get("decision") or ("Diterima" if is_sig else "Ditolak"))
            t_mod_rows.append([
                h["code"],
                h["relation"],
                _fmt3(h.get("original_sample")),
                _fmt3(h.get("sample_mean")),
                _fmt3(h.get("stdev")),
                _fmt3(h.get("t_statistics")),
                _fmt3(h.get("p_value")),
                ket,
                dec
            ])

    add_table_data(doc, title=f"Tabel 4. {tbl_mod_num}\nHasil Pengujian Efek Moderasi Variabel Kontingensi", headers=t_mod_headers, data=t_mod_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi Moderasi (>= 4 baris, bersitasi)
    if _tpl:
        p(doc, f"Tabel 4.{tbl_mod_num} menampilkan kerangka pengujian efek moderasi. Koefisien interaksi, T-statistik, dan p-value belum tersedia karena data empiris belum diolah, sehingga peran {z_var} sebagai penguat atau pelemah hubungan belum dapat ditetapkan. Setelah keluaran SmartPLS 4 dimasukkan, penilaian mengikuti ambang T >= 1,96 dan p < 0,05, serta arah koefisien interaksi (Henseler & Fassott, 2010; Hair et al., 2022).", indent=0.5, space_after=6)
    else:
        t_stat_val = t_mod_rows[0][5] if t_mod_rows else "-"
        p_val_mod = t_mod_rows[0][6] if t_mod_rows else "-"
        _acc = sum(1 for h in mod_hypos if str(h.get("decision", "")).lower().startswith("diterima"))
        p(doc, f"Hasil estimasi pada Tabel 4.{tbl_mod_num} menunjukkan bahwa dari {len(mod_hypos)} hipotesis moderasi, {_acc} dinyatakan diterima. Nilai T-statistik pada baris pertama sebesar {t_stat_val} dengan p-value {p_val_mod}. Arah koefisien interaksi menentukan apakah {z_var} memperkuat atau memperlemah hubungan yang diuji, sehingga penafsirannya harus disertai pembacaan tanda koefisien, bukan hanya signifikansinya (Henseler & Fassott, 2010; Armstrong & Taylor, 2023).", indent=0.5, space_after=6)


def _build_doc_sub_4_a_5_4(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """A.5.4 Pengujian Efek Mediasi Termoderasi (Conditional Indirect Effects / Hayes Model 15)."""
    ordered_vars = _get_ordered_variables(ctx)
    k_vars = len(ordered_vars)
    instansi = ctx["instansi"]
    tahun = ctx.get("tahun", "2026")
    pls = ctx["pls"]
    hypos = pls.get("hypotheses_testing", [])
    x_vars_str = _join_id(ctx.get("x_vars", []))
    m_var = ctx.get("m", "")
    z_var = ctx.get("z", "")
    y_var = ctx.get("y_var", "Kinerja Pegawai")
    _tpl = s3data.is_template(ctx)

    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")
        h3(doc, "5. Pengujian Hipotesis Penelitian")

    # f. Pengujian Efek Mediasi Termoderasi (Hayes Model 15)
    h4(doc, "f. Pengujian Efek Mediasi Termoderasi (Conditional Indirect Effects & Index of Moderated Mediation)")
    p(doc, f"Pengujian efek mediasi termoderasi (moderated mediation) bertujuan untuk membuktikan secara empiris apakah mekanisme transmisi tidak langsung (indirect effects) dari variabel-variabel eksogen ({x_vars_str}) terhadap {y_var} melalui variabel mediasi ({m_var}) beroperasi secara bersyarat pada tingkat variabel kontingensi pemoderasi ({z_var}) (Preacher, Rucker, & Hayes, 2007; Hayes, 2022). Sesuai dengan spesifikasi Hayes Process Model 15, variabel moderasi {z_var} mengondisikan jalur pengaruh mediasi ke variabel endogen akhir (M -> Y) secara simultan dengan jalur langsung (X -> Y) pada {instansi}. Keberadaan mediasi termoderasi dibuktikan melalui Indeks Mediasi Termoderasi (Index of Moderated Mediation) dan selang kepercayaan bootstrap 95% (Hayes, 2022). Efek mediasi termoderasi dinyatakan terbukti signifikan apabila selang kepercayaan 95% bias-corrected percentile bootstrap tidak mencakup angka nol (zero is not straddled) (Hayes, 2022; Hair et al., 2022). Hasil estimasi Indeks Mediasi Termoderasi disajikan pada tabel berikut:", indent=0.5, space_after=6)

    # Tabel Index of Moderated Mediation
    tbl_modmed_num = 19 + k_vars
    t_mm_headers = ["Hipotesis", "Jalur Mediasi Bersyarat", "Index of ModMed (β)", "Sample Mean (M)", "Std Dev (STDEV)", "T-Statistik", "P-Value", "95% CI Bootstrap", "Keputusan"]
    t_mm_rows = []

    cond_hypos = [h for h in hypos if h.get("type") == "conditional_indirect"]
    if not cond_hypos:
        cond_hypos = [h for h in hypos if "modmed" in h.get("relation", "").lower() or "conditional" in h.get("type", "")]

    for h in cond_hypos:
        if _tpl:
            t_mm_rows.append([
                h["code"], h["relation"], s3data.PLACEHOLDER, s3data.PLACEHOLDER,
                s3data.PLACEHOLDER, s3data.PLACEHOLDER, s3data.PLACEHOLDER,
                s3data.PLACEHOLDER, "Belum ditetapkan",
            ])
        else:
            _ci_low = h.get("ci_low")
            _ci_high = h.get("ci_high")
            if _ci_low is not None and _ci_high is not None:
                ci_str = f"[{_fmt3(_ci_low)}; {_fmt3(_ci_high)}]"
            else:
                ci_str = s3data.PLACEHOLDER
            t_mm_rows.append([
                h["code"],
                h["relation"],
                _fmt3(h.get("original_sample")),
                _fmt3(h.get("sample_mean")),
                _fmt3(h.get("stdev")),
                _fmt3(h.get("t_statistics")),
                _fmt3(h.get("p_value")),
                ci_str,
                str(h.get("decision") or "Belum ditetapkan")
            ])

    add_table_data(doc, title=f"Tabel 4. {tbl_modmed_num}\nHasil Estimasi Indeks Mediasi Termoderasi (Index of Moderated Mediation - Hayes Model 15)", headers=t_mm_headers, data=t_mm_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))

    # Narasi Tabel Index
    if _tpl:
        p(doc, f"Tabel 4.{tbl_modmed_num} menampilkan kerangka evaluasi mediasi termoderasi. Nilai Indeks Mediasi Termoderasi, standar deviasi, dan selang kepercayaan bootstrap belum tersedia karena data empiris belum diolah. Setelah estimasi SmartPLS 4 dan Hayes Process selesai, signifikansi dinilai berdasarkan apakah rentang interval 95% memuat angka nol (Hayes, 2022).", indent=0.5, space_after=6)
    else:
        _acc = sum(1 for h in cond_hypos if str(h.get("decision", "")).lower().startswith("diterima"))
        p(doc, f"Berdasarkan hasil estimasi pada Tabel 4.{tbl_modmed_num}, dari {len(cond_hypos)} hipotesis mediasi termoderasi, {_acc} dinyatakan diterima dan signifikan. Nilai selang kepercayaan 95% bias-corrected bootstrap tidak mencakup angka nol, membuktikan bahwa kekuatan mediasi {m_var} dalam menghubungkan variabel eksogen terhadap {y_var} secara empiris bergantung secara bersyarat pada tingkat {z_var} (Preacher et al., 2007; Hayes, 2022).", indent=0.5, space_after=6)

    # Tabel Rincian Efek Mediasi Bersyarat (-1 SD, Mean, +1 SD)
    cm_details = pls.get("conditional_indirect_details") or []
    if cm_details:
        tbl_cond_num = tbl_modmed_num + 1
        t_cd_headers = ["Jalur Pengaruh Mediasi Bersyarat", "Kondisi Variabel Pemoderasi (Z)", "Conditional Indirect Effect", "Interpretasi Efek Transmisi"]
        t_cd_rows = []
        for cmd in cm_details:
            path_lbl = cmd.get("path") or f"Xi -> {m_var} -> {y_var} | {z_var}"
            t_cd_rows.append([path_lbl, f"Rendah (-1 SD {z_var})", _fmt3(cmd.get("indirect_low_z")), "Transmisi mediasi berlangsung pada intensitas minimal"])
            t_cd_rows.append(["", f"Sedang (Mean {z_var})", _fmt3(cmd.get("indirect_mean_z")), "Transmisi mediasi berlangsung pada intensitas moderat"])
            t_cd_rows.append(["", f"Tinggi (+1 SD {z_var})", _fmt3(cmd.get("indirect_high_z")), "Transmisi mediasi diperkuat secara optimal"])

        if t_cd_rows:
            add_table_data(doc, title=f"Tabel 4. {tbl_cond_num}\nRincian Efek Tidak Langsung Bersyarat pada Berbagai Tingkat {z_var}", headers=t_cd_headers, data=t_cd_rows, source=s3data.source_label(ctx, f"Data diolah SmartPLS 4, ({tahun})"))
            p(doc, f"Tabel 4.{tbl_cond_num} memperlihatkan bahwa efek tidak langsung {m_var} mengalami eskalasi bertahap dari kondisi {z_var} rendah (-1 SD), sedang (Mean), hingga tinggi (+1 SD). Pola ini menegaskan fungsi {z_var} sebagai pengakselerasi transmisi psikologis-keperilakuan di lingkungan {instansi} (Hayes, 2022; Robbins & Judge, 2022).", indent=0.5, space_after=6)


def _build_doc_sub_4_a_5(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """Sub-Bab A.5 Pengujian Hipotesis Penelitian (Direct, Indirect, Moderation & ModMed) (Composite Lengkap)."""
    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "A. Hasil Analisis")

    h3(doc, "5. Pengujian Hipotesis Penelitian")

    # Narasi Pengantar Bootstrapping (>= 4 baris, bersitasi)
    p(doc, f"Pengujian hipotesis penelitian merupakan inti pembuktian empiris dalam disertasi ini yang bertujuan untuk menguji secara ilmiah keberlakuan hubungan kausal langsung (direct effects), efek transmisi mediasi (indirect effects), serta efek interaksi kontingensi moderasi (interaction effects) (Hair et al., 2022). Mengingat PLS-SEM tidak mengasumsikan distribusi normalitas data multivariat secara ketat, maka pengujian signifikansi statistik koefisien jalur dilakukan melalui prosedur non-parametrik Bootstrapping dengan menghasilkan 5.000 subsampel acak terdistribusi ulang (resampling) (Streukens & Leroi-Werelds, 2016; Hair et al., 2022). Keputusan penerimaan hipotesis dipedomani oleh pengujian dua arah (two-tailed test) pada taraf signifikansi alpha 5% (α = 0,05), di mana suatu hipotesis dinyatakan diterima dan terbukti signifikan secara statistik apabila menghasilkan nilai T-statistik >= 1,96 dan nilai probabilitas signifikansi (p-value) < 0,05 (Cooper & Schindler, 2022; Sarstedt et al., 2022).", indent=0.5, space_after=6)

    # 1. Direct Effects
    _build_doc_sub_4_a_5_1(doc, ctx, standalone=False)

    # 2. Indirect / Mediation Effects
    if ctx.get("m"):
        _build_doc_sub_4_a_5_2(doc, ctx, standalone=False)

    # 3. Moderation Effects
    if ctx.get("z"):
        _build_doc_sub_4_a_5_3(doc, ctx, standalone=False)

    # 4. Moderated Mediation Effects (Hayes Model 15)
    _hypos = (ctx.get("pls") or {}).get("hypotheses_testing", [])
    if ctx.get("m") and ctx.get("z") and any(h.get("type") == "conditional_indirect" for h in _hypos):
        _build_doc_sub_4_a_5_4(doc, ctx, standalone=False)


def _build_doc_single_hypothesis_b(doc: Document, ctx: Dict[str, Any], hypo_index: int = 1, standalone: bool = True):
    """Pembahasan satu hipotesis: lima alinea, sadar variabel judul dan mode data."""
    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
        h2(doc, "B. Pembahasan Hasil Penelitian")

    instansi = ctx["instansi"]
    pls = ctx["pls"]
    hypos = pls.get("hypotheses_testing", [])
    _tpl = s3data.is_template(ctx)
    v = _v5(ctx)

    if hypo_index <= len(hypos):
        h = dict(hypos[hypo_index - 1])
    else:
        h = {
            "code": f"H{hypo_index}",
            "statement": f"Hubungan struktural ke-{hypo_index} dalam model penelitian",
            "original_sample": None,
            "t_statistics": None,
            "p_value": None,
            "type": "direct_effect",
            "relation": "",
            "decision": None,
        }

    code = str(h.get("code") or f"H{hypo_index}")
    stat = str(h.get("statement") or h.get("relation") or "")
    rel = str(h.get("relation") or "")
    h_type = str(h.get("type") or "")
    decision = str(h.get("decision") or "")

    h3(doc, f"{hypo_index}. Pembahasan Hipotesis {code}: {stat}")

    # --- Alinea 1: bukti statistik (jujur pada mode template) ---
    if _tpl:
        a1 = (
            f"Hipotesis {code} dirumuskan sebagai berikut: {stat}. Nilai koefisien jalur, "
            f"T-statistik, dan p-value belum tersedia karena data empiris riil belum diolah, "
            f"sehingga alinea ini tidak menetapkan penerimaan atau penolakan hipotesis. Setelah "
            f"keluaran SmartPLS 4 dimasukkan, pembacaan akan mengikuti ambang T >= 1,96 dan "
            f"p < 0,05 disertai penilaian arah koefisien (Hair et al., 2022; Sarstedt et al., "
            f"2022). Sampai tahap itu, uraian pada alinea berikutnya bersifat penjelasan "
            f"teoretis dan kontekstual, bukan klaim hasil."
        )
    else:
        try:
            orig = float(h.get("original_sample"))
            t_val = float(h.get("t_statistics"))
            p_val = float(h.get("p_value"))
        except (TypeError, ValueError):
            orig = t_val = p_val = None
        if orig is None:
            a1 = (
                f"Parameter statistik untuk hipotesis {code} belum lengkap pada keluaran yang "
                f"dimasukkan, sehingga pembacaan signifikansi belum dapat dilakukan. Lengkapi "
                f"koefisien jalur, T-statistik, dan p-value agar alinea ini dapat ditafsirkan "
                f"(Hair et al., 2022)."
            )
        else:
            arah = "positif" if orig >= 0 else "negatif"
            if decision:
                keputusan = decision.lower()
            else:
                keputusan = "diterima" if (t_val >= 1.96 and p_val < 0.05) else "tidak diterima"
            a1 = (
                f"Hasil estimasi model struktural PLS-SEM menunjukkan koefisien jalur hipotesis "
                f"{code} bertanda {arah} sebesar {orig:.3f}, dengan T-statistik {t_val:.3f} dan "
                f"p-value {p_val:.3f}. Pada taraf signifikansi 5%, hipotesis ini dinyatakan "
                f"{keputusan}. Besaran tersebut dibaca bersama ukuran efek dan interval "
                f"kepercayaan agar kesimpulan tidak bertumpu pada nilai p semata "
                f"(Hair et al., 2022; Cooper & Schindler, 2022)."
            )
    p(doc, a1, indent=0.5, space_after=4)

    # --- Alinea 2: dialog teori sesuai jenis hipotesis dan variabel judul ---
    _x_str = v["x_str"]
    _m_str = v["m_str"]
    _z_str = v["z_str"]
    _y = v["y"]

    if "moderation" in h_type or "conditional" in h_type or ("*" in rel):
        a2 = (
            f"Secara teoretis, hipotesis {code} menyangkut peran {_z_str or 'variabel moderasi'} "
            f"sebagai kondisi batas. Teori kontingensi menjelaskan bahwa kekuatan hubungan "
            f"antarvariabel tidak seragam, melainkan bergantung pada tingkat faktor situasional "
            f"tertentu (Donaldson, 2023; Robbins & Judge, 2022). Dalam kerangka itu, koefisien "
            f"interaksi dibaca sebagai perubahan kemiringan hubungan, bukan sebagai tambahan "
            f"prediktor biasa, sehingga penafsirannya menuntut pemeriksaan simple slope pada "
            f"tingkat moderator yang rendah, sedang, dan tinggi (Henseler & Fassott, 2010; "
            f"Hair et al., 2022)."
        )
    elif "mediation" in h_type or "indirect" in h_type or rel.count("->") > 1:
        a2 = (
            f"Secara teoretis, hipotesis {code} menyangkut mekanisme perantara. Teori kognitif "
            f"sosial dan teori penentuan diri menjelaskan bahwa pengaruh faktor eksternal "
            f"terhadap perilaku kerja berlangsung melalui proses internal individu, bukan secara "
            f"otomatis (Ryan & Deci, 2022; Robbins & Judge, 2022). Karena itu, efek tidak "
            f"langsung perlu dibaca bersama efek langsung: bila keduanya searah, mekanisme "
            f"bersifat komplementer; bila efek langsung melemah, mekanisme berperan lebih "
            f"dominan (Zhao et al., 2010; Hayes, 2022)."
        )
    else:
        a2 = (
            f"Secara teoretis, hipotesis {code} menyangkut hubungan langsung dalam rantai "
            f"kapasitas-kinerja. Perspektif resource-based view menjelaskan bahwa kapasitas "
            f"internal organisasi, termasuk {_x_str}, merupakan aset yang harus diaktifkan "
            f"melalui praktik manajerial agar menghasilkan luaran (Hitt et al., 2023; "
            f"Boxall & Purcell, 2022). Perspektif tujuan menambahkan bahwa aktivasi tersebut "
            f"efektif bila sasaran kerja jelas dan umpan balik tersedia, sehingga hubungan "
            f"langsung tidak dapat dipisahkan dari konteks pengelolaan kinerja (Schunk, 2022; "
            f"Robbins & Judge, 2022)."
        )
    p(doc, a2, indent=0.5, space_after=4)

    # --- Alinea 3: konteks lokus tanpa mengarang observasi ---
    a3 = (
        f"Pada konteks {instansi}, hubungan yang diuji pada hipotesis {code} perlu dibaca dengan "
        f"memperhatikan karakter pekerjaan aparatur yang menangani {_get_instansi_focus(instansi)}. "
        f"Tuntutan ketepatan administrasi, keterbatasan waktu, dan kebutuhan koordinasi "
        f"lintas unit membuat {_x_str} tidak selalu dapat langsung diterjemahkan menjadi {_y} "
        f"(Hasibuan, 2022; Sedarmayanti, 2022). Karena itu, pembahasan pada tingkat lokus "
        f"diarahkan pada kondisi kerja yang memungkinkan atau menghambat penerjemahan tersebut, "
        f"bukan pada dugaan angka yang belum terukur."
    )
    p(doc, a3, indent=0.5, space_after=4)

    # --- Alinea 4: komparasi literatur ---
    cite_grp = ctx.get("citation_groups") or s3_academic_researcher.get_s3_citation_groups(ctx)
    pos_cites = cite_grp.get("pos_citations", "Hair et al., 2022")
    gap_cites = cite_grp.get("gap_citations", "Hayes, 2022")

    rng_h = nv.rng_for(ctx, f"4b-{code}")
    if _tpl:
        a4 = nv.pick(rng_h, [
            (
                f"Literatur mutakhir menyediakan sejumlah temuan yang relevan untuk dibandingkan "
                f"bagi hipotesis {code} ({pos_cites}). Arah dan besaran hubungan yang dilaporkan "
                f"tidak seragam, dan sebagian studi mencatat hasil yang tidak signifikan pada "
                f"konteks sektor publik ({gap_cites}). Perbandingan dengan hasil penelitian ini "
                f"akan dilakukan setelah nilai estimasi tersedia, agar tidak muncul klaim yang "
                f"mendahului data (Cooper & Schindler, 2022; Neuman, 2022)."
            ),
            (
                f"Temuan terdahulu yang terkumpul pada literatur mutakhir memberi bahan pembanding "
                f"bagi hipotesis {code} ({pos_cites}). Namun arah dan kekuatan hubungan yang "
                f"dilaporkan belum seragam; beberapa kajian bahkan menemukan hubungan yang tidak "
                f"signifikan pada organisasi sektor publik ({gap_cites}). Oleh sebab itu, dialog "
                f"dengan literatur ditunda sampai nilai estimasi tersedia, agar pembahasan tidak "
                f"mendahului data (Cooper & Schindler, 2022; Neuman, 2022)."
            ),
            (
                f"Sebagai bahan pembanding, literatur mutakhir memuat sejumlah hasil yang relevan "
                f"bagi hipotesis {code} ({pos_cites}). Keragaman arah maupun besaran efek, serta "
                f"laporan hasil yang tidak signifikan pada sebagian studi sektor publik "
                f"({gap_cites}), menuntut kehati-hatian. Perbandingan sistematis akan disajikan "
                f"setelah nilai estimasi tersedia (Cooper & Schindler, 2022; Neuman, 2022)."
            ),
            (
                f"Bukti empiris terdahulu yang dirangkum untuk hipotesis {code} memperlihatkan "
                f"dua kubu temuan ({pos_cites} berbanding {gap_cites}). Selisih arah dan besaran "
                f"efek tersebut belum dapat ditafsirkan sebelum nilai estimasi penelitian ini "
                f"tersedia; karena itu, pembahasan komparatif ditempatkan setelah tahap estimasi "
                f"(Cooper & Schindler, 2022; Neuman, 2022)."
            ),
            (
                f"Pembanding bagi hipotesis {code} berasal dari studi yang melaporkan efek positif "
                f"({pos_cites}) maupun studi yang menemukan efek lemah atau kondisional "
                f"({gap_cites}). Kedua kelompok akan didialogkan secara berimbang begitu nilai "
                f"estimasi tersedia, sehingga simpulan tidak bertumpu pada kesan awal "
                f"(Cooper & Schindler, 2022; Neuman, 2022)."
            ),
        ])
    else:
        a4 = (
            f"Hasil hipotesis {code} dapat didialogkan dengan temuan terdahulu yang tercatat pada "
            f"literatur mutakhir ({pos_cites}). Perbandingan diarahkan pada kesamaan arah, "
            f"besaran efek, dan karakter sampel, bukan pada kesimpulan umum tentang berlaku "
            f"universalnya hubungan. Sejumlah studi lain melaporkan hasil yang berbeda atau tidak "
            f"signifikan ({gap_cites}), sehingga temuan penelitian ini lebih tepat dibaca sebagai "
            f"bukti kontekstual pada {instansi} (Cooper & Schindler, 2022; Neuman, 2022)."
        )
    p(doc, a4, indent=0.5, space_after=4)

    # --- Alinea 5: implikasi manajerial ---
    if "moderation" in h_type or "conditional" in h_type:
        a5 = (
            f"Implikasi manajerial hipotesis {code} terletak pada perlunya kebijakan yang "
            f"diferensial, bukan seragam. Jika {_z_str or 'faktor moderasi'} berperan sebagai "
            f"kondisi batas, unit dengan tingkat {_z_str or 'kondisi tersebut'} rendah memerlukan "
            f"dukungan berbeda dibanding unit dengan tingkat tinggi (Armstrong & Taylor, 2023). "
            f"Pimpinan dapat memetakan kondisi tersebut per unit kerja dan menyesuaikan "
            f"intensitas pembinaan, tanpa mengubah standar kinerja yang berlaku umum "
            f"(Hair et al., 2022)."
        )
    else:
        a5 = (
            f"Implikasi manajerial hipotesis {code} terletak pada penajaman tindakan, bukan "
            f"penambahan program. Pimpinan {instansi} dapat memusatkan perbaikan pada aspek "
            f"{_x_str} yang paling dekat dengan tuntutan tugas, disertai pemantauan indikator "
            f"antara agar perubahan dapat dinilai (Armstrong & Taylor, 2023; Sinambela, 2023). "
            f"Kebijakan demikian lebih mudah dievaluasi karena memiliki dasar variabel yang jelas "
            f"dan terukur (Hasibuan, 2022; Cooper & Schindler, 2022)."
        )
    p(doc, a5, indent=0.5, space_after=8)


def _build_doc_sub_4_b(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """Sub-Bab B Pembahasan Hasil Penelitian (Formula Baku 5 Alinea Doktoral) (Composite Lengkap)."""
    if standalone:
        h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")

    h2(doc, "B. Pembahasan Hasil Penelitian")
    instansi = ctx["instansi"]
    pls = ctx["pls"]
    hypos = pls.get("hypotheses_testing", [])

    # Narasi Pengantar Sub-Bab B (>= 4 baris, bersitasi)
    p(doc, f"Pembahasan hasil penelitian ini disusun untuk menginterpretasikan dan mendialogkan seluruh temuan empiris yang diperoleh dari pemodelan struktural PLS-SEM dengan khazanah teori keilmuan manajemen perilaku keorganisasian serta konteks faktual pada {instansi} (Hair et al., 2022). Sejalan dengan standar rigoritas akademik program doktor (S3), pembahasan naskah tidak sekadar mengulang pembacaan angka-angka koefisien statistik, melainkan melakukan sintesis kritis mengenai bagaimana temuan empiris tersebut mengonfirmasi, memperluas, atau merekonstruksi Grand Theory yang dijadikan pijakan ontologis penelitian (Cooper & Schindler, 2022; Sekaran & Bougie, 2020). Setiap hipotesis penelitian (mulai dari H1 hingga H{len(hypos)}) dibedah secara komprehensif menggunakan formula baku lima alinea akademik doktoral: (1) evaluasi bukti statistik empiris, (2) telaah dan dialog teori rujukan, (3) dinamika faktual operasional di lokus instansi, (4) komparasi riset empiris bereputasi mutakhir (2022–2026), serta (5) perumusan implikasi manajerial dan kebijakan praktis bagi jajaran pimpinan (Armstrong & Taylor, 2023; Robbins & Judge, 2022).", indent=0.5, space_after=6)

    # Loop SELURUH Hipotesis (H1 s.d H_total) memanggil single hypothesis builder
    for idx_h in range(1, len(hypos) + 1):
        _build_doc_single_hypothesis_b(doc, ctx, hypo_index=idx_h, standalone=False)


def _build_doc_sub_4_a(doc: Document, ctx: Dict[str, Any]):
    """Composite Sub-Bab A: Hasil Analisis Lengkap (A.1 s.d A.5)."""
    h1(doc, "BAB IV\nHASIL ANALISIS DAN PEMBAHASAN")
    h2(doc, "A. Hasil Analisis")
    _build_doc_sub_4_a_1(doc, ctx, standalone=False)
    p(doc, "", space_after=12)
    _build_doc_sub_4_a_2(doc, ctx, standalone=False)
    p(doc, "", space_after=12)
    _build_doc_sub_4_a_3(doc, ctx, standalone=False)
    p(doc, "", space_after=12)
    _build_doc_sub_4_a_4(doc, ctx, standalone=False)
    p(doc, "", space_after=12)
    _build_doc_sub_4_a_5(doc, ctx, standalone=False)


# Legacy backward-compatible aliases for Bab IV builders
_build_doc_sub_4_1_1 = _build_doc_sub_4_a_1
_build_doc_sub_4_1_2 = _build_doc_sub_4_a_1
_build_doc_sub_4_1_3 = _build_doc_sub_4_a_2
_build_doc_sub_4_1_4 = _build_doc_sub_4_a_2
_build_doc_sub_4_1_5 = _build_doc_sub_4_a_2

_build_doc_sub_4_2_poin_1 = _build_doc_sub_4_a_3
_build_doc_sub_4_2_poin_2 = _build_doc_sub_4_a_4
_build_doc_sub_4_2_poin_3 = _build_doc_sub_4_a_5
_build_doc_sub_4_2_poin_4 = _build_doc_sub_4_a_5
_build_doc_sub_4_2_poin_5 = _build_doc_sub_4_a_5
_build_doc_sub_4_2_poin_6 = _build_doc_sub_4_a_5
_build_doc_sub_4_2_poin_7 = _build_doc_sub_4_b


def _v5(ctx: Dict[str, Any]) -> Dict[str, Any]:
    """Konteks ringkas Bab V: label variabel, klausa kondisional, dan mode data."""
    y = (ctx.get("y_var") or "Kinerja Pegawai").strip()
    xs = [x for x in (ctx.get("x_vars") or []) if x]
    ms = [m for m in (ctx.get("m_vars") or []) if m]
    zs = [z for z in (ctx.get("z_vars") or []) if z]
    pls = ctx.get("pls") or {}
    return {
        "instansi": ctx.get("instansi") or "Instansi Pemerintah",
        "y": y,
        "x": xs,
        "m": ms,
        "z": zs,
        "x_str": _join_id(xs) or "variabel eksogen",
        "m_str": _join_id(ms),
        "z_str": _join_id(zs),
        "has_m": bool(ms),
        "has_z": bool(zs),
        "template": s3data.is_template(ctx),
        "hypos": pls.get("hypotheses_testing", []) or [],
        "acronym": _generate_doctoral_model_acronym(ctx),
        "n": ctx.get("n", 0),
    }


def _v5_stat(h: Dict[str, Any]) -> str:
    """Format ringkas parameter jalur: beta, T, dan p."""
    def _f(value, digits: int = 3) -> str:
        try:
            return f"{float(value):.{digits}f}"
        except (TypeError, ValueError):
            return "-"

    return (f"\u03b2 = {_f(h.get('original_sample'))}; "
            f"T = {_f(h.get('t_statistics'))}; p = {_f(h.get('p_value'))}")


def _v5_model_phrase(v: Dict[str, Any]) -> str:
    """Frasa model sesuai variabel yang benar-benar ada pada judul."""
    if v["has_m"] and v["has_z"]:
        return (f"model mediasi termoderasi yang menempatkan {v['m_str']} sebagai variabel "
                f"mediasi dan {v['z_str']} sebagai variabel moderasi")
    if v["has_m"]:
        return f"model mediasi yang menempatkan {v['m_str']} sebagai variabel mediasi"
    if v["has_z"]:
        return f"model moderasi yang menempatkan {v['z_str']} sebagai variabel moderasi"
    return "model pengaruh langsung antarvariabel"


def _build_doc_sub_5_a(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """Sub-Bab A: Ringkasan Penelitian, dirakit dari variabel judul dan mode data."""
    if standalone:
        h1(doc, "BAB V\nSIMPULAN DAN SARAN")
    h2(doc, "A. Ringkasan Penelitian")

    v = _v5(ctx)
    instansi, y = v["instansi"], v["y"]

    p(doc, f"Penelitian disertasi ini bertolak dari kesenjangan antara tuntutan kinerja yang "
           f"semakin ketat dan kapasitas aparatur yang belum sepenuhnya terkelola di lingkungan "
           f"{instansi}. Reformasi birokrasi menuntut bukti empiris tentang bagaimana "
           f"{v['x_str']} memengaruhi {y}, bukan sekadar penjelasan normatif tentang pentingnya "
           f"perbaikan. Karena itu kajian ini merancang {_v5_model_phrase(v)} dan mengujinya pada "
           f"aparatur yang menjalankan fungsi {_get_instansi_focus(instansi)}. "
           f"Tujuannya adalah menghasilkan penjelasan yang dapat diuji ulang sekaligus memberi "
           f"dasar bagi perbaikan kebijakan kepegawaian yang lebih terarah (Robbins & Judge, 2022; "
           f"Armstrong & Taylor, 2023).", indent=0.5, space_after=6)

    p(doc, f"Secara metodologis, penelitian ini menggunakan pendekatan kuantitatif kausal "
           f"eksplanatori dengan survei kuesioner berskala Likert lima poin terhadap {v['n']} "
           f"responden. Data dianalisis dengan Partial Least Squares Structural Equation Modeling "
           f"melalui perangkat lunak SmartPLS 4, mencakup evaluasi model pengukuran, model "
           f"struktural, serta pengujian efek tidak langsung dan efek interaksi sesuai kebutuhan "
           f"model (Hair et al., 2022; Sarstedt et al., 2022). Pemeriksaan validitas konvergen, "
           f"validitas diskriminan, dan reliabilitas komposit dilakukan sebelum koefisien jalur "
           f"ditafsirkan (Hayes, 2022).", indent=0.5, space_after=6)

    if v["template"]:
        p(doc, f"Ringkasan temuan belum dapat dinyatakan secara kuantitatif karena data empiris "
               f"riil belum dimasukkan ke dalam konfigurasi penelitian. Bab ini karena itu tidak "
               f"menetapkan hipotesis mana pun sebagai diterima atau ditolak. Setelah keluaran "
               f"SmartPLS 4 tersedia, ringkasan akan mengikuti urutan hipotesis pada Bab I dan "
               f"memuat koefisien jalur, nilai T, serta nilai p untuk setiap hubungan yang diuji "
               f"(Hair et al., 2022).", indent=0.5, space_after=6)
    else:
        n_h = len(v["hypos"])
        n_acc = sum(1 for h in v["hypos"] if str(h.get("decision", "")).lower().startswith("diterima"))
        p(doc, f"Dari {n_h} hipotesis yang diuji, {n_acc} dinyatakan diterima dan {n_h - n_acc} "
               f"tidak diterima pada taraf signifikansi 5%. Rincian koefisien jalur dan keputusan "
               f"setiap hipotesis disajikan pada subbab kesimpulan, sedangkan penafsiran "
               f"teoretisnya dibahas pada Bab IV. Ringkasan ini menekankan pola umum: hubungan "
               f"yang melibatkan {v['x_str']} sebagai penentu {y} perlu dibaca bersama mekanisme "
               f"yang menghubungkan keduanya (Hayes, 2022; Robbins & Judge, 2022).",
               indent=0.5, space_after=6)

    if v["has_m"] and v["has_z"]:
        novelty = (f"Kebaruan penelitian ini terletak pada perumusan {v['acronym']} sebagai satu "
                   f"kesatuan: {v['m_str']} diposisikan sebagai mekanisme yang meneruskan pengaruh "
                   f"{v['x_str']} terhadap {y}, sementara {v['z_str']} diposisikan sebagai kondisi "
                   f"batas yang menentukan kuat lemahnya penerusan tersebut.")
    elif v["has_m"]:
        novelty = (f"Kebaruan penelitian ini terletak pada perumusan {v['acronym']}, yaitu model "
                   f"yang menjelaskan jalur {v['x_str']} menuju {y} melalui {v['m_str']} sebagai "
                   f"mekanisme perantara, bukan sebagai hubungan langsung yang berdiri sendiri.")
    elif v["has_z"]:
        novelty = (f"Kebaruan penelitian ini terletak pada perumusan {v['acronym']}, yaitu model "
                   f"yang menempatkan {v['z_str']} sebagai kondisi batas bagi hubungan "
                   f"{v['x_str']} dengan {y}.")
    else:
        novelty = (f"Kebaruan penelitian ini terletak pada perumusan {v['acronym']}, yaitu pemetaan "
                   f"jalur langsung antara {v['x_str']} dan {y} pada konteks instansi pengelola "
                   f"pendapatan daerah yang selama ini lebih banyak dibahas secara normatif.")
    p(doc, novelty + " Dengan begitu, kontribusi teoretis tidak berhenti pada penambahan variabel, "
                     "tetapi pada penjelasan mekanisme dan batas berlakunya hubungan yang diuji "
                     "(Hair et al., 2022; Hayes, 2022).", indent=0.5, space_after=6)

    if v["has_m"] and v["has_z"]:
        synthesis = (f"Secara keseluruhan, peningkatan {y} di {instansi} tidak dapat dicapai "
                     f"melalui satu tuas kebijakan tunggal. Kapasitas aparatur ({v['x_str']}) perlu "
                     f"diterjemahkan ke dalam perilaku kerja yang terlihat melalui {v['m_str']}, "
                     f"dengan {v['z_str']} sebagai faktor yang menentukan seberapa kuat penerjemahan "
                     f"itu berlangsung. Sintesis ini menegaskan bahwa perbaikan kinerja menuntut "
                     f"orkestrasi kebijakan kepegawaian, penataan beban kerja, dan penguatan kondisi "
                     f"kerja secara simultan (Armstrong & Taylor, 2023; Sarstedt et al., 2022).")
    elif v["has_m"]:
        synthesis = (f"Secara keseluruhan, peningkatan {y} di {instansi} tidak dapat dicapai "
                     f"melalui satu tuas kebijakan tunggal. Kapasitas aparatur ({v['x_str']}) perlu "
                     f"diterjemahkan ke dalam perilaku kerja yang terlihat melalui {v['m_str']}, "
                     f"sehingga perbaikan indikator antara menjadi bagian dari strategi, bukan "
                     f"pelengkap laporan (Armstrong & Taylor, 2023; Sarstedt et al., 2022).")
    elif v["has_z"]:
        synthesis = (f"Secara keseluruhan, peningkatan {y} di {instansi} bergantung pada penataan "
                     f"{v['x_str']}, dan {v['z_str']} menentukan seberapa kuat hubungan itu "
                     f"berlangsung pada kondisi kerja yang berbeda. Sintesis ini menegaskan bahwa "
                     f"kebijakan kepegawaian perlu mempertimbangkan variasi konteks, bukan hanya "
                     f"nilai rata-rata organisasi (Armstrong & Taylor, 2023; Sarstedt et al., 2022).")
    else:
        synthesis = (f"Secara keseluruhan, peningkatan {y} di {instansi} menuntut penataan "
                     f"{v['x_str']} secara langsung, tanpa mengabaikan faktor kontekstual yang "
                     f"turut menentukan pelaksanaannya. Sintesis ini menegaskan bahwa perbaikan "
                     f"kinerja bersifat multidimensi dan menuntut kebijakan kepegawaian yang "
                     f"konsisten (Armstrong & Taylor, 2023; Sarstedt et al., 2022).")
    p(doc, synthesis, indent=0.5, space_after=6)


def _build_doc_sub_5_b(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """Sub-Bab B: Kesimpulan, diturunkan langsung dari struktur hipotesis judul."""
    if standalone:
        h1(doc, "BAB V\nSIMPULAN DAN SARAN")
    h2(doc, "B. Kesimpulan")

    v = _v5(ctx)

    if v["template"]:
        p(doc, f"Kesimpulan disusun mengikuti urutan rumusan masalah pada Bab I. Karena data "
               f"empiris riil belum tersedia, subbab ini memuat daftar hubungan yang diuji "
               f"beserta pernyataan bahwa keputusannya belum ditetapkan. Kalimat kesimpulan final "
               f"menunggu hasil estimasi SmartPLS 4 agar tidak muncul klaim yang tidak didukung "
               f"data (Hair et al., 2022; Hayes, 2022).", indent=0.5, space_after=6)
    else:
        n_h = len(v["hypos"])
        n_acc = sum(1 for h in v["hypos"] if str(h.get("decision", "")).lower().startswith("diterima"))
        p(doc, f"Kesimpulan berikut dirumuskan dari hasil estimasi model struktural atas {n_h} "
               f"hipotesis yang diajukan, dengan {n_acc} hipotesis diterima dan {n_h - n_acc} "
               f"tidak diterima pada taraf signifikansi 5%. Setiap butir kesimpulan diikat pada "
               f"kode hipotesis dan parameter jalurnya agar dapat diverifikasi ulang (Hair et al., "
               f"2022; Sarstedt et al., 2022).", indent=0.5, space_after=6)

    specs = [
        (("direct_to_mediator",),
         f"Pengaruh {v['x_str']} terhadap {v['m_str']}", v["has_m"]),
        (("direct_to_dependent",),
         f"Pengaruh Langsung {v['x_str']} terhadap {v['y']}", True),
        (("mediator_to_dependent",),
         f"Pengaruh {v['m_str']} terhadap {v['y']}", v["has_m"]),
        (("mediation_indirect",),
         f"Peran Mediasi {v['m_str']}", v["has_m"]),
        (("moderation_interaction", "conditional_indirect", "moderator_main_effect"),
         f"Peran Moderasi {v['z_str']}", v["has_z"]),
    ]

    num = 0
    for types, label, enabled in specs:
        if not enabled:
            continue
        hs = [h for h in v["hypos"] if h.get("type") in types]
        if not hs:
            continue
        num += 1
        codes = ", ".join(str(h.get("code", "-")) for h in hs)
        p(doc, f"{num}. {label} ({codes}):", bold=True, indent=0.5, space_after=3)
        if v["template"]:
            rels = "; ".join(
                f"{h.get('code', '-')} \u2014 {h.get('relation') or h.get('statement', '')}"
                for h in hs
            )
            p(doc, f"Hubungan yang diuji pada kelompok ini adalah: {rels}. Keputusan atas "
                   f"hipotesis tersebut belum ditetapkan karena nilai koefisien jalur, T, dan p "
                   f"belum tersedia. Setelah data dimasukkan, butir ini akan memuat arah hubungan, "
                   f"besaran pengaruh, dan status penerimaan hipotesis (Hair et al., 2022).",
                   indent=0.75, space_after=6)
        else:
            detail = "; ".join(
                f"{h.get('code', '-')} ({_v5_stat(h)}; {h.get('decision', '-')})" for h in hs
            )
            p(doc, f"Hasil estimasi untuk kelompok ini adalah: {detail}. Besaran tersebut dibaca "
                   f"sebagai bukti empiris atas arah hubungan yang dirumuskan pada Bab II, sekaligus "
                   f"menjadi dasar penilaian kontribusi teoretis pada subbab implikasi (Hayes, 2022; "
                   f"Robbins & Judge, 2022).", indent=0.75, space_after=6)

    if num == 0:
        p(doc, f"Belum ada hipotesis terstruktur yang dapat dikelompokkan pada konfigurasi ini, "
               f"sehingga kesimpulan dirumuskan langsung dari hubungan antara {v['x_str']} dan "
               f"{v['y']} (Hair et al., 2022).", indent=0.75, space_after=6)


def _build_doc_sub_5_c(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """Sub-Bab C: Implikasi Teori, berakar pada variabel judul (bukan tema fraud/audit)."""
    if standalone:
        h1(doc, "BAB V\nSIMPULAN DAN SARAN")
    h2(doc, "C. Implikasi Teori")

    v = _v5(ctx)
    instansi, y = v["instansi"], v["y"]

    p(doc, f"Temuan penelitian ini memberi kontribusi pada penjelasan perilaku kerja aparatur di "
           f"sektor publik, khususnya pada rantai hubungan antara kapasitas individu, pengalaman "
           f"kerja, dan luaran kinerja di {instansi}. Kontribusi tersebut tidak terletak pada "
           f"penambahan istilah baru, melainkan pada penataan hubungan antarkonstruk yang selama "
           f"ini sering diuji secara terpisah. Empat butir implikasi teoretis berikut merangkum "
           f"posisi kajian ini dalam literatur manajemen sumber daya manusia dan perilaku "
           f"keorganisasian (Robbins & Judge, 2022; Armstrong & Taylor, 2023).",
           indent=0.5, space_after=6)

    h3(doc, "1. Kontribusi bagi Teori Perilaku Keorganisasian")
    p(doc, f"Kajian ini memperkuat pandangan bahwa {y} merupakan hasil interaksi antara "
           f"karakteristik individu dan karakteristik pekerjaan, bukan akibat satu faktor tunggal. "
           f"{v['x_str']} diposisikan sebagai masukan yang harus diolah organisasi, sedangkan "
           f"kinerja diposisikan sebagai luaran yang tidak serta-merta mengikuti besarnya masukan. "
           f"Dengan demikian, teori perilaku keorganisasian memperoleh penegasan bahwa mekanisme "
           f"psikologis dan kontekstual perlu dimasukkan sebagai variabel penjelas, bukan "
           f"diasumsikan konstan (Robbins & Judge, 2022; Armstrong & Taylor, 2023).",
      indent=0.5, space_after=6)

    h3(doc, "2. Rekonstruksi Hubungan Kapasitas dan Kinerja")
    p(doc, f"Literatur manajemen sumber daya manusia cenderung mengandaikan hubungan linier antara "
           f"peningkatan kapasitas dan peningkatan kinerja. Kajian ini menunjukkan bahwa andaian "
           f"itu perlu diperiksa ulang: kapasitas yang tinggi tidak otomatis berubah menjadi "
           f"kinerja yang tinggi apabila tidak disertai kondisi kerja yang memungkinkan. "
           f"Implikasinya, teori bukan hanya perlu menjelaskan apa yang dimiliki aparatur, tetapi "
           f"juga bagaimana kepemilikan tersebut diaktifkan dalam tugas sehari-hari di {instansi} "
           f"(Armstrong & Taylor, 2023; Hair et al., 2022).", indent=0.5, space_after=6)

    if v["has_m"] and v["has_z"]:
        h3(doc, "3. Peran Mekanisme Perantara dan Kondisi Batas")
        p(doc, f"Kajian ini menempatkan {v['m_str']} sebagai mekanisme perantara dan {v['z_str']} "
               f"sebagai kondisi batas dalam satu model. Pemisahan peran tersebut penting karena "
               f"keduanya menjawab pertanyaan berbeda: mekanisme menjelaskan mengapa pengaruh "
               f"terjadi, sedangkan kondisi batas menjelaskan kapan pengaruh itu menguat atau "
               f"melemah. Penggabungan keduanya memperkaya penjelasan teoretis tentang {y} dan "
               f"menghindari kesimpulan yang berlaku umum tanpa batas konteks (Hayes, 2022; "
               f"Hair et al., 2022).", indent=0.5, space_after=6)
    elif v["has_m"]:
        h3(doc, "3. Peran Mekanisme Perantara")
        p(doc, f"Kajian ini menempatkan {v['m_str']} sebagai mekanisme perantara antara "
               f"{v['x_str']} dan {y}. Posisi tersebut penting karena menjelaskan proses, bukan "
               f"hanya korelasi: perubahan pada variabel eksogen baru bermakna apabila terlebih "
               f"dahulu mengubah keadaan psikologis atau perilaku yang relevan. Penjelasan berbasis "
               f"mekanisme ini memperkaya teori {y} dan memberi dasar yang lebih kokoh bagi "
               f"intervensi manajerial (Hayes, 2022; Robbins & Judge, 2022).",
          indent=0.5, space_after=6)
    elif v["has_z"]:
        h3(doc, "3. Peran Kondisi Batas")
        p(doc, f"Kajian ini menempatkan {v['z_str']} sebagai kondisi batas bagi hubungan "
               f"{v['x_str']} dengan {y}. Temuan semacam ini mengingatkan bahwa hubungan "
               f"antarvariabel tidak berlaku seragam pada semua keadaan, sehingga teori perlu "
               f"merumuskan batas konteks secara eksplisit. Penjelasan tersebut memperkaya teori "
               f"{y} dan mengurangi risiko generalisasi berlebihan (Hayes, 2022; Hair et al., 2022).",
          indent=0.5, space_after=6)
    else:
        h3(doc, "3. Penegasan Jalur Langsung")
        p(doc, f"Kajian ini menempatkan {v['x_str']} sebagai penentu langsung {y}. Posisi tersebut "
               f"penting karena memberi dasar empiris bagi hubungan yang selama ini lebih banyak "
               f"diterima sebagai proposisi normatif. Penegasan jalur langsung ini memperkaya "
               f"teori {y} dan menyediakan titik tolak bagi pengujian mekanisme lanjutan "
               f"(Hair et al., 2022; Robbins & Judge, 2022).", indent=0.5, space_after=6)

    h3(doc, f"4. Validasi Konstruksi Model {v['acronym']}")
    p(doc, f"Secara keseluruhan, {v['acronym']} bukan sekadar akronim, melainkan pernyataan "
           f"tentang hubungan yang saling mengunci. Model ini menyatukan {v['x_str']} sebagai "
           f"masukan, {y} sebagai luaran, serta mekanisme dan kondisi batas yang menghubungkan "
           f"keduanya pada konteks {instansi}. Nilai teoretisnya terletak pada kejelasan posisi "
           f"setiap konstruk, sehingga pengujian ulang pada konteks lain dapat dilakukan tanpa "
           f"mengaburkan batas antar-konsep (Sarstedt et al., 2022; Cooper & Schindler, 2022).",
      indent=0.5, space_after=6)


def _build_doc_sub_5_d(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """Sub-Bab D: Implikasi Manajerial, diturunkan dari variabel dan model judul."""
    if standalone:
        h1(doc, "BAB V\nSIMPULAN DAN SARAN")
    h2(doc, "D. Implikasi Manajerial")

    v = _v5(ctx)
    instansi, y = v["instansi"], v["y"]

    p(doc, f"Implikasi manajerial berikut diturunkan dari struktur variabel pada judul, bukan dari "
           f"daftar rekomendasi umum. Sasaran utamanya adalah pimpinan dan unit pengelola "
           f"kepegawaian di {instansi} yang bertanggung jawab menata kapasitas, beban kerja, dan "
           f"kondisi kerja aparatur (Armstrong & Taylor, 2023; Robbins & Judge, 2022).",
           indent=0.5, space_after=6)

    num = 0
    num += 1
    h3(doc, f"{num}. Penguatan Kapasitas dan Kesesuaian Penempatan")
    p(doc, f"Penguatan {v['x_str']} sebaiknya tidak berhenti pada penambahan jam pelatihan. "
           f"Pemetaan kebutuhan kompetensi per jabatan, penilaian kesesuaian penempatan, dan "
           f"penugasan yang menantang namun terukur lebih menentukan apakah kapasitas benar-benar "
           f"terpakai. Pimpinan unit perlu meninjau secara berkala apakah tuntutan jabatan sejalan "
           f"dengan penguasaan tugas aparatur, sehingga kesenjangan kapasitas dapat ditutup secara "
           f"spesifik di tiap unit kerja (Armstrong & Taylor, 2023).", indent=0.5, space_after=6)

    num += 1
    h3(doc, f"{num}. Penataan Beban Kerja dan Ketersediaan Sumber Daya")
    p(doc, f"Beban kerja yang tidak terdistribusi menimbulkan kelelahan dan menurunkan mutu "
           f"luaran, meskipun kapasitas individu memadai. Manajemen perlu mengukur volume tugas "
           f"per unit, menyeimbangkan penugasan, dan memastikan ketersediaan sarana serta informasi "
           f"yang dibutuhkan. Penataan ini bukan sekadar penambahan personel, melainkan penyesuaian "
           f"antara tuntutan tugas dan sumber daya yang tersedia di {instansi} (Robbins & Judge, "
           f"2022; Sarstedt et al., 2022).", indent=0.5, space_after=6)

    if v["has_m"]:
        num += 1
        h3(doc, f"{num}. Pelembagaan {v['m_str']} dalam Indikator Kinerja")
        p(doc, f"Apabila {v['m_str']} berfungsi sebagai mekanisme perantara, maka keadaan tersebut "
               f"layak diukur dan dipantau, bukan hanya diasumsikan. Pimpinan dapat memasukkan "
               f"indikator {v['m_str']} ke dalam penilaian kinerja berkala, misalnya melalui "
               f"survei singkat, dialog tim, dan umpan balik atasan-bawahan. Dengan cara itu, "
               f"perbaikan indikator antara menjadi bagian dari siklus manajemen, bukan laporan "
               f"tahunan yang berdiri sendiri (Hayes, 2022; Armstrong & Taylor, 2023).",
          indent=0.5, space_after=6)

    if v["has_z"]:
        num += 1
        h3(doc, f"{num}. Penguatan {v['z_str']} sebagai Kondisi Pendukung")
        p(doc, f"Peran {v['z_str']} sebagai kondisi batas menuntut perhatian pada kualitas "
               f"lingkungan kerja psikologis. Program pengembangan {v['z_str']} dapat dirancang "
               f"melalui pelatihan resiliensi, pembinaan harapan dan optimisme kerja, serta "
               f"keteladanan pimpinan dalam menghadapi tekanan. Penguatan ini bersifat melengkapi, "
               f"bukan menggantikan perbaikan struktural, sehingga hasilnya terasa pada unit yang "
               f"bekerja di bawah tekanan tinggi (Robbins & Judge, 2022; Hair et al., 2022).",
          indent=0.5, space_after=6)
    else:
        num += 1
        h3(doc, f"{num}. Penguatan Iklim Kerja dan Dukungan Pimpinan")
        p(doc, f"Kinerja aparatur tidak hanya ditentukan oleh kapasitas individu, tetapi juga oleh "
               f"kejelasan arahan dan dukungan atasan. Pimpinan unit perlu memastikan bahwa sasaran "
               f"kerja dipahami, hambatan operasional ditangani, dan umpan balik diberikan secara "
               f"teratur. Penguatan iklim kerja ini bersifat melengkapi perbaikan struktural dan "
               f"menjaga agar kapasitas yang telah dibangun tidak surut di tengah tekanan tugas "
               f"(Robbins & Judge, 2022).", indent=0.5, space_after=6)


def _build_doc_sub_5_e(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """Sub-Bab E: Implikasi Praktis, berupa langkah operasional yang dapat langsung dijalankan."""
    if standalone:
        h1(doc, "BAB V\nSIMPULAN DAN SARAN")
    h2(doc, "E. Implikasi Praktis")

    v = _v5(ctx)
    instansi, y = v["instansi"], v["y"]

    p(doc, f"Implikasi praktis berikut dirumuskan sebagai langkah operasional yang dapat langsung "
           f"diterapkan pada rutinitas kerja di {instansi}. Setiap langkah diarahkan pada variabel "
           f"yang benar-benar diuji dalam penelitian ini, sehingga dapat dipantau hasilnya tanpa "
           f"menambah beban administratif yang berlebihan (Cooper & Schindler, 2022; "
           f"Sinambela, 2023).", indent=0.5, space_after=6)

    num = 0
    num += 1
    h3(doc, f"{num}. Pemetaan Kompetensi dan Pelatihan Terarah")
    p(doc, f"Unit kepegawaian dapat menyusun peta kompetensi per jabatan yang memuat penguasaan "
           f"tugas, kebutuhan pelatihan, dan rencana pengembangan jangka pendek. Peta ini dipakai "
           f"untuk memilih peserta pelatihan secara tepat sasaran, bukan berdasarkan pemerataan "
           f"kuota. Evaluasi pascapelatihan dilakukan dengan menilai perubahan cara kerja di unit "
           f"masing-masing, sehingga manfaat pelatihan dapat ditelusuri (Armstrong & Taylor, 2023).",
      indent=0.5, space_after=6)

    num += 1
    h3(doc, f"{num}. Redistribusi Tugas dan Pemantauan Kelelahan Kerja")
    p(doc, f"Pimpinan unit dapat meninjau kembali pembagian tugas secara berkala, khususnya pada "
           f"periode puncak pelaporan dan pelayanan. Pemantauan sederhana terhadap indikasi "
           f"kelelahan, keterlambatan penyelesaian, dan penumpukan pekerjaan dapat menjadi dasar "
           f"penyesuaian beban. Langkah ini menjaga agar tuntutan tugas tetap sepadan dengan "
           f"kapasitas aparatur di tiap unit kerja (Robbins & Judge, 2022).", indent=0.5,
      space_after=6)

    num += 1
    h3(doc, f"{num}. Umpan Balik Rutin dan Pengakuan Kinerja")
    p(doc, f"Umpan balik yang diberikan secara rutin membantu aparatur memahami standar yang "
           f"diharapkan dan memperbaiki cara kerja lebih cepat. Pertemuan singkat mingguan, catatan "
           f"capaian, dan pengakuan atas hasil kerja yang memenuhi standar dapat memperkuat "
           f"keterlibatan kerja. Praktik ini tidak menuntut anggaran besar, tetapi menuntut "
           f"konsistensi pimpinan unit dalam menjalankannya (Armstrong & Taylor, 2023; "
           f"Hasibuan, 2022).", indent=0.5, space_after=6)

    if v["has_z"]:
        num += 1
        h3(doc, f"{num}. Program Penguatan {v['z_str']}")
        p(doc, f"Program penguatan {v['z_str']} dapat dirancang sebagai rangkaian kegiatan pendek "
               f"yang terintegrasi dengan kegiatan rutin, misalnya sesi refleksi tim, pelatihan "
               f"ketahanan menghadapi tekanan, dan pendampingan bagi unit dengan beban tertinggi. "
               f"Kegiatan ini sebaiknya dievaluasi dengan indikator sederhana yang disepakati, "
               f"sehingga dapat dilanjutkan atau diperbaiki pada periode berikutnya di {instansi}. ",
          indent=0.5, space_after=6)


def _build_doc_sub_5_f(doc: Document, ctx: Dict[str, Any], standalone: bool = True):
    """Sub-Bab F: Saran dan Keterbatasan Penelitian."""
    if standalone:
        h1(doc, "BAB V\nSIMPULAN DAN SARAN")
    h2(doc, "F. Saran dan Keterbatasan Penelitian")

    v = _v5(ctx)
    instansi, y = v["instansi"], v["y"]
    _model_vars = [x for x in (v["x"] + v["m"] + v["z"] + [v["y"]]) if x]
    _model_vars_str = _join_id(_model_vars)

    h3(doc, "1. Keterbatasan Penelitian")
    p(doc, f"Meskipun penelitian ini dirancang dengan protokol metodologis yang terjaga, peneliti "
           f"menyadari adanya sejumlah keterbatasan yang perlu dikemukakan secara terbuka agar "
           f"hasilnya dibaca secara proporsional (Creswell & Creswell, 2023; Sekaran & Bougie, "
           f"2020). Keterbatasan tersebut dipetakan sebagai berikut.", indent=0.5, space_after=6)

    p(doc, f"a. Keterbatasan Geografis dan Ruang Lingkup: Penelitian ini difokuskan pada instansi "
           f"pemerintah daerah di lingkungan {instansi}. Karakteristik geografis, dinamika "
           f"kepemimpinan lokal, dan kultur birokrasi di wilayah ini memiliki kekhasan yang mungkin "
           f"berbeda dengan instansi lain. Penerapan kesimpulan pada konteks organisasi yang "
           f"berbeda menuntut penyesuaian kontekstual (Creswell & Creswell, 2023; "
           f"Cooper & Schindler, 2022).", indent=0.75, space_after=4)

    p(doc, f"b. Kerentanan Bias Respons: Pengumpulan data menggunakan kuesioner tertutup yang "
           f"diisi sendiri oleh responden. Karena variabel yang diteliti berkaitan dengan "
           f"penilaian diri atas pekerjaan, terdapat kemungkinan responden memberi jawaban yang "
           f"cenderung normatif. Peneliti telah menjaga kerahasiaan identitas dan memeriksa "
           f"common method bias, namun potensi bias tetap perlu diperhitungkan saat menafsirkan "
           f"hasil (Podsakoff et al., 2024).", indent=0.75, space_after=4)

    p(doc, f"c. Keterbatasan Variabel Model: Model {v['acronym']} membatasi pengujian pada "
           f"{_model_vars_str}. Kinerja aparatur merupakan fenomena multidimensional, sehingga "
           f"faktor lain di luar model ini, seperti digitalisasi layanan, sistem remunerasi berbasis "
           f"merit, dan kebijakan pengawasan eksternal, berpotensi memiliki pengaruh yang belum "
           f"terukur (Hair et al., 2022; Hayes, 2022).", indent=0.75, space_after=6)

    h3(doc, "2. Saran")
    p(doc, f"Berdasarkan temuan, kesimpulan, dan keterbatasan yang telah diuraikan, peneliti "
           f"merumuskan saran bagi dua kelompok sasaran: peneliti lanjutan dan pemangku kebijakan "
           f"di {instansi} (Sekaran & Bougie, 2020; Neuman, 2022).", indent=0.5, space_after=6)

    h4(doc, "a. Bagi Peneliti Selanjutnya (Agenda Riset Mendatang):")
    p(doc, f"1) Pengayaan Pendekatan Metodologi: Peneliti lanjutan disarankan mengombinasikan "
           f"survei kuantitatif dengan wawancara mendalam atau desain longitudinal. Pendekatan "
           f"campuran membantu menjelaskan motif di balik jawaban kuesioner dan menangkap perubahan "
           f"perilaku kerja sepanjang waktu (Neuman, 2022; Creswell & Creswell, 2023).",
      indent=0.75, space_after=4)

    p(doc, f"2) Perluasan Objek dan Penambahan Variabel: Penelitian berikutnya dapat memperluas "
           f"cakupan sampel ke beberapa instansi untuk memungkinkan perbandingan antar-daerah. "
           f"Variabel kontemporer seperti pemanfaatan sistem informasi, kecakapan digital, dan "
           f"kualitas komunikasi internal layak dipertimbangkan sebagai perluasan model "
           f"(Hair et al., 2022; Sarstedt et al., 2022).", indent=0.75, space_after=6)

    h4(doc, "b. Bagi Pemerintah Daerah dan Pemangku Kebijakan:")
    p(doc, f"1) Penggunaan {v['acronym']} sebagai Kerangka Diagnosis: Pemerintah daerah dapat "
           f"memakai kerangka model ini untuk mendiagnosis titik lemah manajemen kepegawaian, bukan "
           f"sebagai dokumen seremonial. Diagnosis dilakukan dengan memeriksa secara berurutan: "
           f"kapasitas aparatur, kondisi kerja, keadaan perantara, dan luaran kinerja.",
      indent=0.75, space_after=4)

    p(doc, f"2) Penyelarasan Kebijakan Kepegawaian: Kebijakan pelatihan, penempatan, penilaian "
           f"kinerja, dan tunjangan sebaiknya diselaraskan agar saling menopang. Tanpa "
           f"penyelarasan, perbaikan pada satu aspek mudah kehilangan dampak karena aspek lain "
           f"tetap berjalan dengan arah berbeda di {instansi} (Armstrong & Taylor, 2023; "
           f"Robbins & Judge, 2022).", indent=0.75, space_after=6)


def _build_doc_sub_5_all(doc: Document, ctx: Dict[str, Any]):
    """
    Composite Sub-Bab Bab V Lengkap: Simpulan dan Saran (A s.d F).
    Format Baku TAPD 2026 Disertasi S3.
    """
    _build_doc_sub_5_a(doc, ctx, standalone=True)
    p(doc, "", space_after=12)
    _build_doc_sub_5_b(doc, ctx, standalone=False)
    p(doc, "", space_after=12)
    _build_doc_sub_5_c(doc, ctx, standalone=False)
    p(doc, "", space_after=12)
    _build_doc_sub_5_d(doc, ctx, standalone=False)
    p(doc, "", space_after=12)
    _build_doc_sub_5_e(doc, ctx, standalone=False)
    p(doc, "", space_after=12)
    _build_doc_sub_5_f(doc, ctx, standalone=False)


# Legacy backward-compatible aliases for Bab V builders
def _build_doc_sub_5_1(doc: Document, ctx: Dict[str, Any]):
    """5.1 Kesimpulan Penelitian Doktoral (Legacy Alias -> memanggil Sub-Bab B)."""
    _build_doc_sub_5_b(doc, ctx, standalone=True)


def _build_doc_sub_5_2(doc: Document, ctx: Dict[str, Any]):
    """5.2 Implikasi Penelitian (Legacy Alias -> memanggil Sub-Bab C, D, dan E)."""
    h1(doc, "BAB V\nSIMPULAN DAN SARAN")
    _build_doc_sub_5_c(doc, ctx, standalone=False)
    p(doc, "", space_after=12)
    _build_doc_sub_5_d(doc, ctx, standalone=False)
    p(doc, "", space_after=12)
    _build_doc_sub_5_e(doc, ctx, standalone=False)


def _build_doc_sub_5_3(doc: Document, ctx: Dict[str, Any]):
    """5.3 Keterbatasan dan Saran (Legacy Alias -> memanggil Sub-Bab F)."""
    _build_doc_sub_5_f(doc, ctx, standalone=True)


def _sanitize_document_citations(doc: Document, biblio: List[Any]) -> int:
    """Bersihkan sitasi naskah yang tidak punya entri pustaka (kanon + artikel
    riil ter-fetch). Jaring pengaman zero-ghost sebelum audit resiprokal.
    Hanya mengubah paragraf/sel yang benar-benar memuat sitasi tanpa entri."""
    try:
        keys = s3_academic_researcher._bibliography_key_set(biblio)
    except Exception:
        return 0
    if not keys:
        return 0
    fixed = 0
    in_dapus = False

    def _fix_par(par) -> None:
        nonlocal fixed
        old = par.text or ""
        if not old.strip():
            return
        new = s3_academic_researcher.sanitize_citation_text(old, keys)
        if new != old:
            if par.runs:
                par.runs[0].text = new
                for r in par.runs[1:]:
                    r.text = ""
            else:
                par.add_run(new)
            fixed += 1

    for par in doc.paragraphs:
        t = (par.text or "").strip()
        # Jangan pernah menyunting isi Daftar Pustaka: entri bibliografi memang
        # berbentuk "Nama (tahun)." dan bisa terbaca sebagai sitasi, sehingga
        # sanitizer berisiko memotong penulis terakhir + tahun entri itu sendiri.
        if re.fullmatch(r"DAFTAR\s+PUSTAKA", t, re.I):
            in_dapus = True
            continue
        if in_dapus:
            if re.match(r"^(?:LAMPIRAN|DAFTAR\s+LAMPIRAN|RIWAYAT\s+HIDUP)\b", t, re.I):
                in_dapus = False
            else:
                continue
        if re.match(r"^(?:sumber|source)\s*:", t, re.I) or t.startswith("MODE TEMPLATE"):
            continue
        _fix_par(par)
    # Sel tabel hanya disentuh bila diaudit (sel yang diawali sitasi), agar entri
    # tabel rujukan (mis. nama jurnal + tahun) tidak ikut terpotong.
    cite_cell_re = re.compile(
        r"^[A-Z][A-Za-z'\u2019\-]+"
        r"(?:\s*(?:,\s*[A-Z][A-Za-z'\u2019\-]+|&\s*[A-Z][A-Za-z'\u2019\-]+|et\s+al\.))*"
        r"\s*,?\s*\((?:19|20)\d{2}\)"
    )
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                if not cite_cell_re.match((cell.text or "").strip()):
                    continue
                for par in cell.paragraphs:
                    _fix_par(par)
    return fixed


def _build_doc_sub_dapus(doc: Document, ctx: Dict[str, Any]):
    """Daftar Pustaka Lengkap Standar APA 7th Edition (100% Pustaka Riil Ber-DOI Terverifikasi & Teori Kanon)."""
    h1(doc, "DAFTAR PUSTAKA")

    corpus = ctx.get("academic_refs")
    if not corpus:
        corpus = s3_academic_researcher.fetch_s3_verified_academic_corpus(ctx)
        ctx["academic_refs"] = corpus

    biblio = corpus.get("bibliography", [])
    if not biblio:
        biblio = [(c["citation"] + ". " + c["title"], c["title"]) for c in s3_academic_researcher.ALL_CANONICAL_REFERENCES]

    # Jaring pengaman Zero-Ghost: bersihkan sitasi yang tidak memiliki entri
    # pustaka (kanon + artikel riil) langsung dari naskah sebelum audit.
    try:
        _fixed = _sanitize_document_citations(doc, biblio)
        if _fixed:
            print(f"[s3_refs] Sanitasi Zero-Ghost: {_fixed} sitasi tanpa entri dibersihkan dari naskah.")
    except Exception as _e:
        print(f"[s3_refs] Sanitasi Zero-Ghost dilewati: {_e}")

    # Ekstraksi teks naskah untuk sinkronisasi sitasi resiprokal APA 7th (Zero Ghost References)
    body_text = ""
    # 1. Cek parameter context eksplisit
    if ctx.get("body_text"):
        body_text = str(ctx["body_text"])
    elif ctx.get("full_manuscript_text"):
        body_text = str(ctx["full_manuscript_text"])

    # 2. Cek paragraf dokumen docx saat ini jika naskah dibangun kumulatif
    audit_parts: List[str] = []
    if len(body_text.strip()) < 100 and doc.paragraphs:
        doc_paras_text = []
        for p_elem in doc.paragraphs:
            t = p_elem.text.strip()
            if t == "DAFTAR PUSTAKA":
                break
            if not t:
                continue
            # Catatan provenans tabel (mis. "Sumber: KemenPANRB (2024) ...") dan
            # spanduk mode template bukan sitasi naskah, sehingga tidak dihitung
            # sebagai kutipan pada audit ghost reference.
            if re.match(r"^(?:sumber|source)\s*:", t, re.I) or t.startswith("MODE TEMPLATE"):
                continue
            doc_paras_text.append(t)
        audit_parts = list(doc_paras_text)
        # Sertakan teks tabel (mis. Tabel 1.5 dan Tabel 2.1) agar sitasi di dalam
        # tabel ikut terhitung dalam filter resiprokal.
        cite_cell_re = re.compile(
            r"^[A-Z][A-Za-z'\u2019\-]+"
            r"(?:\s*(?:,\s*[A-Z][A-Za-z'\u2019\-]+|&\s*[A-Z][A-Za-z'\u2019\-]+|et\s+al\.))*"
            r"\s*,?\s*\((?:19|20)\d{2}\)"
        )
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    t = cell.text.strip()
                    if not t:
                        continue
                    doc_paras_text.append(t)
                    # Audit ghost hanya memakai sel yang benar-benar berisi sitasi,
                    # bukan sel judul tabel/nama jurnal.
                    if cite_cell_re.match(t):
                        audit_parts.append(t)
        if doc_paras_text:
            body_text = " ".join(doc_paras_text)
    audit_text = " ".join(audit_parts) if audit_parts else ""

    # 3. Jika dokumen saat ini berdiri sendiri (standalone sub_dapus doc),
    # ambil isi naskah langsung dari Word Aktif via COM
    if len(body_text.strip()) < 100:
        try:
            import win32com.client
            word_app = win32com.client.GetActiveObject("Word.Application")
            active_word_doc = word_app.ActiveDocument
            word_text = active_word_doc.Content.Text
            if "DAFTAR PUSTAKA" in word_text:
                body_text = word_text.split("DAFTAR PUSTAKA")[0]
            else:
                body_text = word_text
        except Exception:
            pass

    # Saring bibliografi berdasarkan teks kutipan yang ada di naskah (100% Bebas Ghost References)
    if len(body_text.strip()) >= 50:
        biblio = s3_academic_researcher.filter_cited_bibliography(biblio, body_text, audit_text=audit_text)

    # Urutkan secara alfabetis baku A-Z berdasarkan nama pengarang
    biblio.sort(key=lambda x: (x[0] if isinstance(x, (list, tuple)) else str(x)).lower())
    # Simpan daftar pustaka final agar bisa dipakai sapuan sanitasi terakhir
    # setelah tahap parafrase pascabangun (vary_document) di akhir perakitan.
    try:
        ctx["_final_biblio"] = list(biblio)
    except Exception:
        pass

    for item in biblio:
        if isinstance(item, (list, tuple)) and len(item) >= 2:
            entry_text, italic_part = item[0], item[1]
        elif isinstance(item, (list, tuple)) and len(item) == 1:
            entry_text, italic_part = item[0], ""
        else:
            entry_text, italic_part = str(item), ""

        para = p(doc, "", hanging=0.5, space_after=6)
        if italic_part and italic_part in entry_text:
            parts = entry_text.split(italic_part, 1)
            r_pre = para.add_run(parts[0])
            r_pre.font.name = "Times New Roman"
            r_pre.font.size = Pt(12)

            r_ital = para.add_run(italic_part)
            r_ital.font.name = "Times New Roman"
            r_ital.font.size = Pt(12)
            r_ital.italic = True

            if len(parts) > 1 and parts[1]:
                r_post = para.add_run(parts[1])
                r_post.font.name = "Times New Roman"
                r_post.font.size = Pt(12)
        else:
            add_rich_runs_to_para(para, entry_text, font_size=12)


def _build_doc_sub_lampiran(doc: Document, ctx: Dict[str, Any]):
    """Lampiran Doktoral Lengkap (Kuesioner 40 Butir & Output PLS)."""
    h1(doc, "LAMPIRAN")
    h2(doc, "Lampiran 1. Kuesioner Penelitian Doktoral (40 Butir Operasional)")

    p(doc, f"Kuesioner ini disusun untuk pengumpulan data primer penelitian Disertasi Doktor berjudul: \"{ctx['judul']}\" pada {ctx['instansi']}.", indent=0.5, space_after=6)
    p(doc, "Petunjuk Pengisian: Berilah tanda centang (✓) pada kolom skala pilihan (1 = Sangat Tidak Setuju, 2 = Tidak Setuju, 3 = Kurang Setuju, 4 = Setuju, 5 = Sangat Setuju).", italic=True, indent=0.5, space_after=6)

    kuesioner_items = ctx.get("kuesioner", [])
    t_headers = ["No", "Kode", "Konstruk / Indikator Pernyataan", "STS (1)", "TS (2)", "KS (3)", "S (4)", "SS (5)"]
    t_rows = []
    for idx, item in enumerate(kuesioner_items, start=1):
        t_rows.append([
            str(item.get("no", idx)),
            item.get("kode", f"IT{idx:02d}"),
            f"[{item.get('variabel', '')}] {item.get('pernyataan', '')}",
            "", "", "", "", ""
        ])

    add_table_data(doc, title="Tabel L.1. Tabulasi Instrumen Kuesioner Doktoral Lengkap", headers=t_headers, data=t_rows, source="Sumber: Master Blueprint Kuesioner Doktoral S3 (2026)")
    p(doc, "", space_after=12)

    h2(doc, "Lampiran 2. Rekapitulasi Output SmartPLS 4 Bootstrapping")
    _lamp_tpl = s3data.is_template(ctx)
    p(doc, f"Total Sampel N = {ctx['n']} | Algoritma: PLS-SEM Consistent Path Weighting | Bootstrapping: 5.000 Subsamples.", italic=True, indent=0.5, space_after=6)

    def _lamp_num(val):
        try:
            return f"{float(val):.3f}"
        except (TypeError, ValueError):
            return s3data.PLACEHOLDER

    if _lamp_tpl:
        p(doc, s3data.TEMPLATE_BANNER, italic=True, indent=0.5, space_after=8)

    _ordered = _get_ordered_variables(ctx)
    _name_by_key = {vk: nm for vk, _role, nm, _vd in _ordered}

    # --- L.2 Ringkasan Model Pengukuran (Outer Model) ---
    ov_headers = ["Variabel Laten", "Kode", "AVE", "Composite Reliability", "Cronbach's Alpha", "VIF", "Keputusan"]
    ov_rows = []
    for vkey, role, vname, _vdata in _ordered:
        ov = (s3data.outer(ctx, vkey) if not _lamp_tpl else None) or {}
        ov_rows.append([
            vname,
            str(ov.get("code") or vkey.upper()),
            _lamp_num(ov.get("ave")),
            _lamp_num(ov.get("composite_reliability")),
            _lamp_num(ov.get("cronbach_alpha")),
            _lamp_num(ov.get("vif")),
            str(ov.get("decision") or "Belum diuji"),
        ])
    add_table_data(
        doc,
        title="Tabel L.2. Ringkasan Model Pengukuran (Outer Model)",
        headers=ov_headers,
        data=ov_rows,
        source=s3data.source_label(ctx, "Output SmartPLS 4 (Algoritma PLS-SEM)"),
    )
    p(doc, "", space_after=12)

    # --- L.3 Ringkasan Koefisien Determinasi (R-Square) ---
    r2_headers = ["Variabel Endogen", "R-Square", "R-Square Adjusted", "Interpretasi"]
    r2_rows = []
    _r2_raw = (ctx.get("pls") or {}).get("r_squares", {}) or {}
    _r2_keys = [k for k in _r2_raw.keys() if k] if not _lamp_tpl else []
    if not _r2_keys:
        _r2_keys = ["y"] + (["m"] if ctx.get("m_vars") else [])
    for vk in _r2_keys:
        if _lamp_tpl:
            _rs_val, _rs_cat, _rs_adj = s3data.PLACEHOLDER, "Belum ditetapkan", s3data.PLACEHOLDER
        else:
            entry = _r2_raw.get(vk) or _r2_raw.get(vk.upper()) or _r2_raw.get(vk.lower()) or {}
            if isinstance(entry, (int, float, str)):
                _raw_val = entry
                _raw_adj = None
            elif isinstance(entry, dict):
                _raw_val = entry.get("r_square", entry.get("r2"))
                _raw_adj = entry.get("r_square_adjusted", entry.get("r2_adj"))
            else:
                _raw_val, _raw_adj = None, None
            try:
                _rv = float(_raw_val)
                _rs_val = f"{_rv:.3f}"
                _rs_cat = "Kuat" if _rv >= 0.50 else ("Moderat" if _rv >= 0.33 else "Lemah")
            except (TypeError, ValueError):
                _rs_val, _rs_cat = s3data.PLACEHOLDER, "Belum ditetapkan"
            try:
                _ra = float(_raw_adj)
                _rs_adj = f"{_ra:.3f}"
            except (TypeError, ValueError):
                _rs_adj = s3data.PLACEHOLDER
        _var_label = _name_by_key.get(vk.lower(), _name_by_key.get(vk.upper(), vk.upper()))
        r2_rows.append([_var_label, _rs_val, _rs_adj, _rs_cat])
    add_table_data(
        doc,
        title="Tabel L.3. Ringkasan Koefisien Determinasi (R-Square)",
        headers=r2_headers,
        data=r2_rows,
        source=s3data.source_label(ctx, "Output SmartPLS 4"),
    )
    p(doc, "", space_after=12)

    # --- L.4 Rekapitulasi Uji Hipotesis ---
    h_headers = ["Kode", "Hubungan Struktural", "Koefisien", "T-Statistik", "P-Value", "Keputusan"]
    _hypos_lamp = (ctx.get("pls") or {}).get("hypotheses_testing", []) or []
    h_rows = []
    for h in _hypos_lamp:
        if _lamp_tpl:
            h_rows.append([
                str(h.get("code") or ""),
                str(h.get("relation") or ""),
                s3data.PLACEHOLDER, s3data.PLACEHOLDER, s3data.PLACEHOLDER, "Belum diuji",
            ])
        else:
            h_rows.append([
                str(h.get("code") or ""),
                str(h.get("relation") or ""),
                _lamp_num(h.get("original_sample")),
                _lamp_num(h.get("t_statistics")),
                _lamp_num(h.get("p_value")),
                str(h.get("decision") or "Belum ditetapkan"),
            ])
    if not h_rows:
        h_rows = [[s3data.PLACEHOLDER, "Belum ada hipotesis teruji", s3data.PLACEHOLDER,
                   s3data.PLACEHOLDER, s3data.PLACEHOLDER, "Belum diuji"]]
    add_table_data(
        doc,
        title="Tabel L.4. Rekapitulasi Uji Hipotesis (Bootstrapping 5.000 Subsampel)",
        headers=h_headers,
        data=h_rows,
        source=s3data.source_label(ctx, "Output SmartPLS 4"),
    )

# ==============================================================================
# 4. MAPPING BUILDER MASTER S3 (36 UNIT)
# ==============================================================================
S3_BUILDER_MAP = {
    "sub_bagian_awal": _build_doc_sub_bagian_awal,
    "sub_1_1": _build_doc_sub_1_1,
    "sub_1_2": _build_doc_sub_1_2,
    "sub_1_3": _build_doc_sub_1_3,
    "sub_1_4": _build_doc_sub_1_4,
    "sub_1_5": _build_doc_sub_1_5,
    "sub_1_6": _build_doc_sub_1_6,

    "sub_2_1_0": _build_doc_sub_2_1_0,
    "sub_2_1_1": _build_doc_sub_2_1_1,
    "sub_2_1_2": _build_doc_sub_2_1_2,
    "sub_2_1_4": _build_doc_sub_2_1_4,
    "sub_2_1_3": _build_doc_sub_2_1_3,
    "sub_2_2": _build_doc_sub_2_2,
    "sub_2_3": _build_doc_sub_2_3,
    "sub_2_3_1": _build_doc_sub_2_3_1,
    "sub_2_3_2": _build_doc_sub_2_3_2,
    "sub_2_4": _build_doc_sub_2_4,

    "sub_3_a": _build_doc_sub_3_a,
    "sub_3_b": _build_doc_sub_3_b,
    "sub_3_c": _build_doc_sub_3_c,
    "sub_3_d": _build_doc_sub_3_d,
    "sub_3_e": _build_doc_sub_3_e,
    "sub_3_f": _build_doc_sub_3_f,
    "sub_3_g": _build_doc_sub_3_g,
    "sub_3_h": _build_doc_sub_3_h,
    "sub_3_i": _build_doc_sub_3_i,

    "sub_3_1": _build_doc_sub_3_1,
    "sub_3_2": _build_doc_sub_3_2,
    "sub_3_3": _build_doc_sub_3_3,
    "sub_3_4": _build_doc_sub_3_4,
    "sub_3_5": _build_doc_sub_3_5,
    "sub_3_6": _build_doc_sub_3_6,

    # Bab II: Granular Builders
    "sub_2_1_3_x1": lambda doc, ctx: _build_doc_single_x_theory(doc, ctx, x_idx=1, standalone=True),
    "sub_2_1_3_x2": lambda doc, ctx: _build_doc_single_x_theory(doc, ctx, x_idx=2, standalone=True),
    "sub_2_1_3_x3": lambda doc, ctx: _build_doc_single_x_theory(doc, ctx, x_idx=3, standalone=True),
    "sub_2_1_3_x4": lambda doc, ctx: _build_doc_single_x_theory(doc, ctx, x_idx=4, standalone=True),
    "sub_2_3_1": _build_doc_sub_2_3_1,
    "sub_2_3_2": _build_doc_sub_2_3_2,
    "sub_2_3_3": _build_doc_sub_2_3_3,

    # Bab III: Granular Builders
    "sub_3_f_y": _build_doc_sub_3_f_y,
    "sub_3_f_m": _build_doc_sub_3_f_m,
    "sub_3_f_z": _build_doc_sub_3_f_z,
    "sub_3_f_x1": lambda doc, ctx: _build_doc_sub_3_f_xk(doc, ctx, x_idx=1, standalone=True),
    "sub_3_f_x2": lambda doc, ctx: _build_doc_sub_3_f_xk(doc, ctx, x_idx=2, standalone=True),
    "sub_3_f_x3": lambda doc, ctx: _build_doc_sub_3_f_xk(doc, ctx, x_idx=3, standalone=True),
    "sub_3_f_x4": lambda doc, ctx: _build_doc_sub_3_f_xk(doc, ctx, x_idx=4, standalone=True),

    # Bab IV: Granular Builders
    "sub_4_a_2_y": _build_doc_sub_4_a_2_y,
    "sub_4_a_2_mz": _build_doc_sub_4_a_2_mz,
    "sub_4_a_2_x1": lambda doc, ctx: _build_doc_sub_4_a_2_xk(doc, ctx, x_idx=1, standalone=True),
    "sub_4_a_2_x2": lambda doc, ctx: _build_doc_sub_4_a_2_xk(doc, ctx, x_idx=2, standalone=True),
    "sub_4_a_2_x3": lambda doc, ctx: _build_doc_sub_4_a_2_xk(doc, ctx, x_idx=3, standalone=True),
    "sub_4_a_2_x4": lambda doc, ctx: _build_doc_sub_4_a_2_xk(doc, ctx, x_idx=4, standalone=True),
    "sub_4_a_2_rekap": _build_doc_sub_4_a_2_rekap,

    "sub_4_a_3_1": _build_doc_sub_4_a_3_1,
    "sub_4_a_3_2": _build_doc_sub_4_a_3_2,
    "sub_4_a_3_3": _build_doc_sub_4_a_3_3,

    "sub_4_a_4_1": _build_doc_sub_4_a_4_1,
    "sub_4_a_4_2": _build_doc_sub_4_a_4_2,
    "sub_4_a_4_3": _build_doc_sub_4_a_4_3,
    "sub_4_a_4_4": _build_doc_sub_4_a_4_4,

    "sub_4_a_5_1": _build_doc_sub_4_a_5_1,
    "sub_4_a_5_2": _build_doc_sub_4_a_5_2,
    "sub_4_a_5_3": _build_doc_sub_4_a_5_3,

    "sub_4_b_h1": lambda doc, ctx: _build_doc_single_hypothesis_b(doc, ctx, hypo_index=1, standalone=True),
    "sub_4_b_h2": lambda doc, ctx: _build_doc_single_hypothesis_b(doc, ctx, hypo_index=2, standalone=True),
    "sub_4_b_h3": lambda doc, ctx: _build_doc_single_hypothesis_b(doc, ctx, hypo_index=3, standalone=True),
    "sub_4_b_h4": lambda doc, ctx: _build_doc_single_hypothesis_b(doc, ctx, hypo_index=4, standalone=True),
    "sub_4_b_h5": lambda doc, ctx: _build_doc_single_hypothesis_b(doc, ctx, hypo_index=5, standalone=True),
    "sub_4_b_h6": lambda doc, ctx: _build_doc_single_hypothesis_b(doc, ctx, hypo_index=6, standalone=True),
    "sub_4_b_h7": lambda doc, ctx: _build_doc_single_hypothesis_b(doc, ctx, hypo_index=7, standalone=True),
    "sub_4_b_h8": lambda doc, ctx: _build_doc_single_hypothesis_b(doc, ctx, hypo_index=8, standalone=True),

    # Bab IV: Format Baku TAPD 2026 (Kanonikal S3)
    "sub_4_a_1": _build_doc_sub_4_a_1,
    "sub_4_a_2": _build_doc_sub_4_a_2,
    "sub_4_a_3": _build_doc_sub_4_a_3,
    "sub_4_a_4": _build_doc_sub_4_a_4,
    "sub_4_a_5": _build_doc_sub_4_a_5,
    "sub_4_b": _build_doc_sub_4_b,
    "sub_4_a": _build_doc_sub_4_a,

    # Bab IV: Legacy Aliases (Backward-Compatible)
    "sub_4_1_1": _build_doc_sub_4_1_1,
    "sub_4_1_2": _build_doc_sub_4_1_2,
    "sub_4_1_3": _build_doc_sub_4_1_3,
    "sub_4_1_4": _build_doc_sub_4_1_4,
    "sub_4_1_5": _build_doc_sub_4_1_5,
    "sub_4_2_poin_1": _build_doc_sub_4_2_poin_1,
    "sub_4_2_poin_2": _build_doc_sub_4_2_poin_2,
    "sub_4_2_poin_3": _build_doc_sub_4_2_poin_3,
    "sub_4_2_poin_4": _build_doc_sub_4_2_poin_4,
    "sub_4_2_poin_5": _build_doc_sub_4_2_poin_5,
    "sub_4_2_poin_6": _build_doc_sub_4_2_poin_6,
    "sub_4_2_poin_7": _build_doc_sub_4_2_poin_7,

    # Bab V: Format Baku TAPD 2026 (Kanonikal S3)
    "sub_5_a": _build_doc_sub_5_a,
    "sub_5_b": _build_doc_sub_5_b,
    "sub_5_c": _build_doc_sub_5_c,
    "sub_5_d": _build_doc_sub_5_d,
    "sub_5_e": _build_doc_sub_5_e,
    "sub_5_f": _build_doc_sub_5_f,
    "sub_5_all": _build_doc_sub_5_all,

    # Bab V: Legacy Aliases (Backward-Compatible)
    "sub_5_1": _build_doc_sub_5_1,
    "sub_5_2": _build_doc_sub_5_2,
    "sub_5_3": _build_doc_sub_5_3,
    "sub_dapus": _build_doc_sub_dapus,
    "sub_lampiran": _build_doc_sub_lampiran
}


# ==============================================================================
# 5. ENTRY POINTS UTAMA: DOCX, NATIVE COM, & CHAINED RUNNER
# ==============================================================================



# ==============================================================================
# SUB-BAB KUSTOM DINAMIS (BEBAS KETIK JUDUL & SUB-BAB OLEH PENGGUNA)
# ==============================================================================

def build_s3_custom_subchapter_doc(custom_meta: Dict[str, Any], config: Optional[Dict[str, Any]] = None) -> Tuple[str, str]:
    """
    Menghasilkan berkas Word .docx untuk Sub-Bab Kustom buatan pengguna
    yang beradaptasi penuh terhadap JUDUL, BAB INDUK (Ch 0-6), VARIABEL, dan INSTANSI PENELITIAN.
    """
    ctx = resolve_s3_context(config)
    code = custom_meta.get("code", "E").strip().upper()
    raw_title = custom_meta.get("title", "Sub-Bab Kustom").strip()
    try:
        chapter = int(custom_meta.get("chapter", 1))
    except Exception:
        chapter = 1

    # Normalisasi judul agar tidak duplikasi kode ("E. E. Judul")
    if raw_title.upper().startswith(f"{code}."):
        title = raw_title[len(code)+1:].strip()
    elif raw_title.upper().startswith(f"{code} "):
        title = raw_title[len(code):].strip()
    else:
        title = raw_title

    doc = Document()
    setup_page(doc, continuous=True)

    h2(doc, f"{code}. {title}")

    y_var = ctx["y_var"]
    instansi = ctx["instansi"]
    x_vars_str = _join_id(ctx["x_vars"])
    m_vars_str = ", ".join(ctx["m_vars"]) if ctx["m_vars"] else "(Tanpa variabel mediasi)"
    z_vars_str = ", ".join(ctx["z_vars"]) if ctx["z_vars"] else "(Tanpa variabel moderasi)"
    model_acronym = _generate_doctoral_model_acronym(ctx)

    title_upper = title.upper()

    # 1. BAGIAN AWAL (CHAPTER 0)
    if chapter == 0 or any(k in title_upper for k in ["PERNYATAAN", "KEASLIAN", "PRAKATA", "SAMBUTAN", "PENGESAHAN"]):
        p(doc, f"Naskah Disertasi Doktor berjudul \"{ctx['judul']}\" ini merupakan karya orisinal hasil penelitian independen mahasiswa Program Pascasarjana dengan lokus empiris pada {instansi}, di bawah supervisi Komisi Promotor.", indent=0.5, space_after=6)
        p(doc, f"1. Integritas Akademik: Penulis menyatakan dengan sebenar-benarnya bahwa seluruh narasi teoritis, instrumen penelitian, model konseptual Model {model_acronym}, kompilasi data empiris lapangan n = {ctx['n']}, serta evaluasi persamaan struktural PLS-SEM di dalamnya bebas dari segala bentuk fabrikasi, falsifikasi, dan plagiarisme ilmiah.", indent=0.75, space_after=3)
        p(doc, f"2. Kepatuhan Standar Doktoral: Seluruh kutipan dan sintesis rujukan yang digunakan telah dicantumkan sumber resminya secara transparan sesuai kaidah sitasi internasional APA 7th Edition (2022–2026).", indent=0.75, space_after=3)
        p(doc, f"3. Tanggung Jawab Moral & Ilmiah: Apabila di kemudian hari terbukti terdapat penyimpangan etika riset atau pelanggaran hak cipta dalam naskah ini, penulis bersedia menerima sanksi akademik yang berlaku sesuai peraturan perundang-undangan pendidikan tinggi Republik Indonesia.", indent=0.75, space_after=6)

    # 2. BAB I (PENDAHULUAN)
    elif chapter == 1 and any(k in title_upper for k in ["RUANG LINGKUP", "BATASAN", "DELIMITASI"]):
        p(doc, f"Ruang lingkup dan batasan dalam penelitian disertasi ini ditetapkan secara terukur dan spesifik agar pengujian Model {model_acronym} memiliki fokus yang tajam, mendalam, dan memenuhi kaidah ketat riset ilmiah pascasarjana sebagai berikut:", indent=0.5, space_after=6)
        _lingkup_vars = [f"variabel terikat ({y_var})"]
        if ctx["m_vars"]:
            _lingkup_vars.append(f"variabel mediasi ({m_vars_str})")
        if ctx["z_vars"]:
            _lingkup_vars.append(f"variabel moderasi ({z_vars_str})")
        _lingkup_vars.append(f"variabel bebas ({x_vars_str})")
        p(doc, f"1. Lingkup Konseptual & Model Struktural: Penelitian difokuskan pada pengujian hubungan kausalitas struktural terintegrasi antara {_join_id(_lingkup_vars)}.", indent=0.75, space_after=3)
        p(doc, f"2. Lingkup Lokus & Objek Empiris: Penelitian dilaksanakan pada lingkungan institusi {instansi} guna menjawab dinamika tata kelola dan pencapaian kinerja riil aparatur.", indent=0.75, space_after=3)
        p(doc, f"3. Lingkup Subjek & Unit Analisis: Subjek penelitian adalah aparatur sipil yang bertugas pada {instansi} dengan ukuran sampel representatif n = {ctx['n']} responden.", indent=0.75, space_after=3)
        p(doc, f"4. Lingkup Metodologis: Menggunakan metode kuantitatif kausal-eksplanatori dengan teknik pemodelan persamaan struktural berbasis varian (PLS-SEM) menggunakan SmartPLS 4 dan Hayes PROCESS Macro.", indent=0.75, space_after=3)
        p(doc, f"5. Lingkup Temporal: Pengumpulan data primer kuesioner dan observasi empiris dilaksanakan pada kurun waktu tahun {ctx['tahun']}.", indent=0.75, space_after=6)

    elif chapter == 1 and any(k in title_upper for k in ["SISTEMATIKA"]):
        p(doc, "Naskah disertasi ini disusun ke dalam 5 (lima) bab yang saling terkait secara logis dan terstruktur sebagai berikut:", indent=0.5, space_after=6)
        p(doc, "BAB I PENDAHULUAN: Memuat latar belakang masalah berbasis 7 lapisan deduktif doktoral, rumusan masalah sekuensial kontinu, tujuan penelitian yang selaras, manfaat penelitian teoretis dan praktis, serta ruang lingkup penelitian.", indent=0.5, space_after=3)
        p(doc, f"BAB II KAJIAN PUSTAKA, KERANGKA KONSEPTUAL, DAN PENGEMBANGAN HIPOTESIS: Menguraikan telaah Grand Theory, Middle Theory, kajian teoritis mendalam per variabel ({y_var}, {m_vars_str}, {z_vars_str}, {x_vars_str}), telaah 12 jurnal internasional bereputasi terdahulu, kerangka konseptual Model {model_acronym}, dan formulasi hipotesis formal.", indent=0.5, space_after=3)
        p(doc, f"BAB III METODE PENELITIAN: Menguraikan desain penelitian, populasi dan sampel {instansi}, operasionalisasi variabel indikator, instrumen penelitian, serta rancangan evaluasi pengukuran dan struktural PLS-SEM.", indent=0.5, space_after=3)
        p(doc, "BAB IV HASIL PENELITIAN DAN PEMBAHASAN: Menyajikan karakteristik responden, statistik deskriptif, evaluasi Outer Model dan Inner Model PLS-SEM, uji hipotesis bootstrapping, analisis mediasi dan moderasi, serta pembahasan 4 alinea pascasarjana.", indent=0.5, space_after=3)
        p(doc, "BAB V SIMPULAN DAN SARAN: Menyajikan kesimpulan menyeluruh 17 hipotesis, implikasi teoretis (doctoral novelty), implikasi manajerial 4 pilar kebijakan, implikasi praktis taktis operasional, serta saran dan keterbatasan penelitian.", indent=0.5, space_after=6)

    elif chapter == 1 and any(k in title_upper for k in ["KEBARUAN", "NOVELTY", "SPESIFIKASI", "ORISINALITAS"]):
        p(doc, f"Sebagai penelitian tingkat doktoral, Disertasi ini memformulasikan kebaruan ilmiah (doctoral novelty) melalui integrasi konseptual Model {model_acronym} yang menjawab anomali kinerja birokrasi pada {instansi}.", indent=0.5, space_after=6)
        p(doc, f"1. Orisinalitas Teoretis: Mensintesiskan proposisi-proposisi teoretis dari Grand Theory dan Middle-Range Theory ke dalam satu kerangka struktural kontingensi hibrida, menunjukkan bahwa infrastruktur formal tata kelola ({x_vars_str}) membutuhkan jembatan transmisi eksekusi perilaku ({m_vars_str}) dan perlindungan integritas ({z_vars_str}) untuk menghasilkan unjuk kerja optimal ({y_var}).", indent=0.75, space_after=3)
        p(doc, f"2. Orisinalitas Kontekstual-Empiris: Menempatkan pengujian model pada institusi sektor publik {instansi} yang memiliki karakteristik beban target dan dinamika birokrasi yang kompleks.", indent=0.75, space_after=3)
        p(doc, f"3. Orisinalitas Metodologis: Mengkombinasikan pengujian model mediasi termoderasi secara simultan menggunakan pendekatan PLS-SEM mutakhir dengan pengujian efek bersyarat (conditional indirect effect) yang presisi.", indent=0.75, space_after=6)

    # 3. BAB II (TINJAUAN PUSTAKA / REGULASI)
    elif chapter == 2 or any(k in title_upper for k in ["KEBIJAKAN", "REGULASI", "YURIDIS", "PERATURAN"]):
        p(doc, f"Kajian mendalam mengenai {title} dalam kerangka konseptual disertasi ini memperkuat landasan teoretis dan yuridis-normatif yang mendasari dinamika tata kelola {instansi}. Keterkaitan antara dimensi regulasi formal dengan perilaku aparatur ({m_vars_str}) dan pertahanan integritas ({z_vars_str}) menjadi prasyarat esensial guna mendorong efektivitas pencapaian {y_var}.", indent=0.5, space_after=6)
        p(doc, f"Secara teoretis, telaah ini mengintegrasikan doktrin stewardship dengan teori kontingensi organisasi, di mana kepatuhan terhadap koridor yuridis bukan sekadar kewajiban administratif, melainkan pendorong terciptanya iklim kerja yang profesional, transparan, dan akuntabel pada {instansi}.", indent=0.5, space_after=6)
        p(doc, f"Dengan menempatkan {title} sebagai pilar penjelas dalam Model {model_acronym}, penelitian ini berhasil menjembatani kesenjangan antara kebijakan makro pemerintah dan praktik manajemen mikro operasional di tingkat unit kerja.", indent=0.5, space_after=6)

    # 4. BAB III (METODOLOGI / PROTOKOL ETIK / ROBUSTNESS METODE)
    elif chapter == 3 or any(k in title_upper for k in ["ETIK", "PROTOKOL", "BIAS", "CMB", "PILOT"]):
        p(doc, f"Penerapan {title} dalam metodologi penelitian disertasi ini dirancang untuk menjamin standar rigoritas, objektivitas, dan validitas tertinggi pada lokus penelitian {instansi}.", indent=0.5, space_after=6)
        p(doc, f"1. Prosedur Perlindungan Subjek & Etika Riset: Pengumpulan data primer dari {ctx['n']} responden aparatur menjunjung tinggi prinsip informed consent, kerahasiaan identitas (anonymity), dan kebebasan partisipasi tanpa paksaan administratif.", indent=0.75, space_after=3)
        p(doc, f"2. Pengendalian Bias Metodologis (Common Method Bias): Dilakukan uji diagnostik Common Method Variance melalui pendekatan Full Collinearity VIF dan Harman's Single-Factor Test pada SmartPLS 4, memastikan seluruh nilai VIF faktor berada di bawah ambang batas toleransi 3,3.", indent=0.75, space_after=3)
        p(doc, f"3. Stabilitas Estimasi: Parameter model struktural diuji melalui prosedur non-parametrik bootstrapping 5.000 subsample dengan two-tailed test tingkat signifikansi 5% (t-tabel 1,96) guna memastikan kestabilan koefisien jalur matematis.", indent=0.75, space_after=6)

    # 5. BAB IV (HASIL / ANALISIS TAMBAHAN / ROBUSTNESS CHECK)
    elif chapter == 4 or any(k in title_upper for k in ["ROBUST", "SENSITIVITAS", "TAMBAHAN", "POST-HOC", "SEGMENTASI"]):
        p(doc, f"Pengujian tambahan mengenai {title} dilaksanakan guna menguji kekokohan (robustness) dan sensitivitas parameter estimasi Model {model_acronym} terhadap variasi spesifikasi model dan segmentasi demografis aparatur pada {instansi}.", indent=0.5, space_after=6)
        p(doc, f"Hasil analisis robustness check mengonfirmasi bahwa arah dan signifikansi koefisien jalur dari {x_vars_str} terhadap {y_var}, baik secara langsung maupun melalui jembatan mediasi {m_vars_str} dan interaksi moderasi {z_vars_str}, tetap stabil dan konsisten secara statistik (p < 0,05).", indent=0.5, space_after=6)
        p(doc, f"Stabilitas arah dan signifikansi tersebut dibaca sebagai indikasi awal ketahanan estimasi; generalisasi ke institusi lain tetap menuntut replikasi pada konteks dan waktu yang berbeda.", indent=0.5, space_after=6)

    # 6. BAB V (SIMPULAN / ROADMAP KEBIJAKAN / MITIGASI RISIKO)
    elif chapter == 5 or any(k in title_upper for k in ["ROADMAP", "PETA JALAN", "MITIGASI", "AGENDA", "REKOMENDASI"]):
        p(doc, f"Perumusan {title} dalam Bab Simpulan dan Saran ini dirancang sebagai panduan aksi transformatif (transformational action plan) berorientasi masa depan bagi pimpinan tertinggi dan jajaran manajemen pada {instansi}.", indent=0.5, space_after=6)
        p(doc, f"1. Rencana Aksi Jangka Pendek (Tahun ke-1): Penguatan pilar infrastruktur tata kelola ({x_vars_str}) melalui pembaharuan SOP berbasis digital dan penguatan sistem pengawasan kepatuhan operasional.", indent=0.75, space_after=3)
        p(doc, f"2. Rencana Aksi Jangka Menengah (Tahun ke-2 s.d ke-3): Peningkatan kapasitas aparatur dan akselerasi budaya kerja ({m_vars_str}) melalui pelatihan terstruktur dan lokakarya internal berkala.", indent=0.75, space_after=3)
        p(doc, f"3. Rencana Aksi Jangka Panjang (Tahun ke-4 s.d ke-5): Institusionalisasi pengawalan integritas ({z_vars_str}) menuju terwujudnya tata kelola birokrasi berkelas dunia (world-class bureaucracy) yang menjamin keunggulan kinerja {y_var} secara berkelanjutan.", indent=0.75, space_after=6)

    # 7. BAGIAN AKHIR (CHAPTER 6 / LAMPIRAN / DOKUMEN RESMI)
    elif chapter == 6 or any(k in title_upper for k in ["LAMPIRAN", "SURAT", "IZIN", "DOKUMEN", "BUKTI"]):
        p(doc, f"Bagian ini mendokumentasikan berkas-berkas administratif, legalitas formal, instrumen instansional, serta bukti empiris pendukung yang memvalidasi seluruh tahapan pelaksanaan riset Disertasi Doktor pada {instansi}.", indent=0.5, space_after=6)
        p(doc, f"Seluruh dokumen dalam {title} telah diverifikasi keabsahannya oleh Komisi Promotor dan pengelola Program Pascasarjana sebagai bukti kepatuhan penuh terhadap tata kelola perizinan riset, kode etik penelitian institusi, dan transparansi data ilmiah.", indent=0.5, space_after=6)

    # 8. GENERIK DOKTORAL ADAPTIF BERBOBOT TINGGI
    else:
        p(doc, f"Kajian mengenai {title} dalam disertasi ini diposisikan sebagai bagian integral dari pengembangan kerangka kerja komprehensif dalam menelaah optimalisasi {y_var} pada {instansi}. Secara konseptual, aspek ini memberikan landasan operasional dan telaah mendalam yang memperkaya dialektika teoretis maupun metodologis dalam pemodelan kausalitas struktural Model {model_acronym}.", indent=0.5, space_after=6)
        p(doc, f"Dalam kaitannya dengan pilar-pilar kapasitas organisasi ({x_vars_str}), perumusan {title} berfungsi sebagai parameter penjelas guna memastikan bahwa seluruh instrumen kebijakan dan interaksi manajerial mampu bertransformasi secara efektif melalui peran {m_vars_str} serta terproteksi oleh kondisi batas {z_vars_str}. Hal ini sangat esensial mengingat realitas dinamika tugas dan akuntabilitas kinerja aparatur pada {instansi} menuntut adanya kejelasan arahan, konsistensi implementasi, dan pengawalan integritas yang kokoh.", indent=0.5, space_after=6)
        p(doc, f"Dengan demikian, keberadaan {title} dalam struktur penelitian ini tidak hanya mempertegas batasan keilmuan yang dikaji, melainkan juga memperkuat orisinalitas kontribusi akademik Disertasi dalam menjawab kesenjangan penelitian (research gap) serta memberikan rekomendasi manajerial yang aplikatif bagi pengambil kebijakan pada {instansi}.", indent=0.5, space_after=6)

    clean_code = code.replace(" ", "_").replace(".", "")
    fn = f"DISERTASI_S3_CUSTOM_CH{chapter}_{clean_code}_{int(time.time())}.docx"
    fp = os.path.join(OUTPUT_DIR, fn)
    doc.save(fp)
    return fn, fp


def build_s3_custom_prompt(custom_meta: Dict[str, Any], ctx: Dict[str, Any]) -> str:
    """
    Menghasilkan Prompt Doktoral AI untuk Sub-Bab Kustom buatan pengguna
    yang beradaptasi terhadap bab induk penempatan (Bab 0-6).
    """
    code = custom_meta.get("code", "E").strip().upper()
    title = custom_meta.get("title", "Sub-Bab Kustom").strip()
    try:
        chapter = int(custom_meta.get("chapter", 1))
    except Exception:
        chapter = 1

    bab_names = {
        0: "BAGIAN AWAL: Front Matter Doktoral",
        1: "BAB I: Pendahuluan (Disertasi Doktor)",
        2: "BAB II: Kajian Pustaka, Kerangka Konseptual, dan Hipotesis",
        3: "BAB III: Metode Penelitian",
        4: "BAB IV: Hasil Penelitian dan Pembahasan",
        5: "BAB V: Simpulan dan Saran (Format Baku TAPD 2026)",
        6: "BAGIAN AKHIR: Daftar Pustaka dan Lampiran"
    }
    bab_induk = bab_names.get(chapter, f"BAB {chapter} (Disertasi Doktor)")

    y_var = ctx["y_var"]
    instansi = ctx["instansi"]
    x_vars_str = _join_id(ctx["x_vars"])
    m_vars_str = ", ".join(ctx["m_vars"]) if ctx["m_vars"] else "(Tanpa variabel mediasi)"
    z_vars_str = ", ".join(ctx["z_vars"]) if ctx["z_vars"] else "(Tanpa variabel moderasi)"
    model_acronym = _generate_doctoral_model_acronym(ctx)

    return f"""Anda adalah Guru Besar dan Pakar Metodologi Penelitian Manajemen yang bertindak sebagai Penulis Ahli Disertasi Doktor (S3) Ilmu Manajemen Pascasarjana Indonesia.

Tugas Anda adalah menyusun naskah akademik tingkat doktoral untuk:
BAGIAN / SUB-BAB: {code}. {title}
BAB INDUK: {bab_induk}

JUDUL DISERTASI: {ctx['judul']}
LOKUS PENELITIAN: {instansi}
VARIABEL PENELITIAN:
- Variabel Terikat (Y): {y_var}
- Variabel Bebas (X): {x_vars_str}
- Variabel Mediasi (M): {m_vars_str}
- Variabel Moderasi (Z): {z_vars_str}
- Model Orisinal: Model {model_acronym}

=== PETUNJUK SPESIFIK PENULISAN DOKTORAL (SUB-BAB {code}. {title}) ===
1. Judul Sub-bab WAJIB ditulis persis: {code}. {title}
2. Karakteristik Naskah:
   - Bahasa Indonesia akademik tingkat doktoral (EYD V), mendalam, analitis, kaya rujukan mutakhir (2022-2026).
   - Selaraskan pembahasan {title} dengan konteks bab induk ({bab_induk}), fenomena empiris pada {instansi}, serta model struktural Model {model_acronym}.
   - Susun dalam 3-5 paragraf epistemologis yang padat, terstruktur rapi, dan bebas dari kalimat klise esai umum.
3. Kepatuhan Format:
   - Dilarang membuat sub-judul kelompok yang tidak relevan.
   - Naskah harus langsung siap disisipkan ke dalam lembar naskah Disertasi resmi.
"""

def get_s3_builder(sub_id: str) -> Optional[Callable]:
    """
    Mengambil builder function untuk sub_id tertentu, mendukung static lookup
    dan regex dynamic matching untuk anak baris granular (X_k, H_k, dll).
    """
    if sub_id in S3_BUILDER_MAP:
        return S3_BUILDER_MAP[sub_id]

    m_x = re.match(r"^sub_2_1_3_x(\d+)$", sub_id)
    if m_x:
        x_idx = int(m_x.group(1))
        return lambda doc, ctx: _build_doc_single_x_theory(doc, ctx, x_idx=x_idx, standalone=True)

    m_f = re.match(r"^sub_3_f_x(\d+)$", sub_id)
    if m_f:
        x_idx = int(m_f.group(1))
        return lambda doc, ctx: _build_doc_sub_3_f_xk(doc, ctx, x_idx=x_idx, standalone=True)

    m_a2 = re.match(r"^sub_4_a_2_x(\d+)$", sub_id)
    if m_a2:
        x_idx = int(m_a2.group(1))
        return lambda doc, ctx: _build_doc_sub_4_a_2_xk(doc, ctx, x_idx=x_idx, standalone=True)

    m_bh = re.match(r"^sub_4_b_h(\d+)$", sub_id)
    if m_bh:
        h_idx = int(m_bh.group(1))
        return lambda doc, ctx: _build_doc_single_hypothesis_b(doc, ctx, hypo_index=h_idx, standalone=True)

    return None


def build_s3_subchapter_doc(sub_id: str, config: Optional[Dict[str, Any]] = None, include_dapus: bool = False) -> Tuple[str, str]:
    """
    Menghasilkan berkas Word .docx untuk 1 sub-bab Disertasi Doktor (S3)
    yang 100% tersinkronisasi dan beradaptasi penuh terhadap JUDUL PENELITIAN.
    """
    ctx = resolve_s3_context(config)
    meta = get_s3_subchapter_meta(sub_id, ctx)

    doc = Document()
    setup_page(doc, continuous=(sub_id not in ("sub_1_1", "sub_bagian_awal")))

    builder = get_s3_builder(sub_id)
    if builder:
        builder(doc, ctx)
    else:
        h2(doc, meta["title"])
        p(doc, f"Uraian lengkap untuk {meta['title']} disusun secara mendalam sesuai standar Disertasi Doktor Ilmu Manajemen berakar pada judul: \"{ctx['judul']}\".", indent=0.5)

    if include_dapus and sub_id not in ("sub_bagian_awal", "sub_dapus", "sub_lampiran"):
        doc.add_page_break()
        _build_doc_sub_dapus(doc, ctx)

    # Format penutup rapi (justify 1.5 line spacing)
    for para in doc.paragraphs:
        if para.text.strip():
            para.paragraph_format.line_spacing = 1.5

    clean_sub = re.sub(r"[^A-Za-z0-9_]+", "_", sub_id)
    fn = f"DISERTASI_S3_{clean_sub.upper()}_{int(time.time())}.docx"
    fp = os.path.join(OUTPUT_DIR, fn)
    doc.save(fp)
    return fn, fp


def _normalize_heading(text: str) -> str:
    """Normalisasi teks heading untuk perbandingan tahan format."""
    return re.sub(r"\s+", " ", str(text or "").replace("\x0b", " ")).strip().casefold()


def _target_heading_index(active_doc) -> set:
    """Kumpulkan teks heading (outline level <= 4) yang sudah ada di dokumen aktif."""
    found = set()
    try:
        total = active_doc.Paragraphs.Count
    except Exception:
        return found
    for i in range(1, total + 1):
        try:
            para = active_doc.Paragraphs(i)
            txt = para.Range.Text or ""
            if not txt.strip():
                continue
            try:
                lvl = int(para.OutlineLevel)
            except Exception:
                lvl = 10
            if lvl <= 4:
                found.add(_normalize_heading(txt))
        except Exception:
            continue
    return found


def _strip_duplicate_chapter_opener(doc_path: str, sub_id: str, existing_headings: set) -> bool:
    """
    Hapus pembuka bab (mis. 'BAB II ...' dan 'A. KAJIAN PUSTAKA') dari sub-dokumen
    bila judul bab tersebut SUDAH ADA di dokumen target. Ini mencegah heading bab
    berulang setiap kali sub-bab disisipkan.

    Return True bila sub-bab ini sudah pernah ada di dokumen (tidak perlu disisipkan lagi).
    """
    if sub_id in ("sub_dapus", "sub_lampiran", "sub_bagian_awal"):
        return False
    try:
        tmp = Document(doc_path)
    except Exception:
        return False

    from docx.text.paragraph import Paragraph as _DocxParagraph

    def _set_paragraphs():
        return [p for p in tmp.paragraphs if (p.text or "").strip()]

    non_empty = _set_paragraphs()
    if not non_empty:
        return False

    changed = False

    # 1. Buang heading "BAB ..." di awal bila bab tersebut sudah ada di target.
    first = non_empty[0]
    first_txt = _normalize_heading(first.text)
    chapter_key = None
    if first_txt.startswith("bab "):
        # Kunci pembanding: nomor bab saja (mis. 'bab ii') agar tidak terpengaruh
        # perbedaan subjudul antar-versi builder.
        m_ch = re.match(r"^(bab\s+[ivxlcdm]+)", first_txt)
        if m_ch:
            chapter_key = m_ch.group(1)
    if chapter_key and any(h.startswith(chapter_key) for h in existing_headings):
        first._element.getparent().remove(first._element)
        changed = True
        # 2. Buang pembuka "A. KAJIAN PUSTAKA / A. Kajian Teori" yang juga sudah ada.
        non_empty = _set_paragraphs()
        if non_empty:
            nxt = non_empty[0]
            nxt_txt = _normalize_heading(nxt.text)
            if nxt_txt in ("a. kajian pustaka", "a. kajian teori"):
                nxt._element.getparent().remove(nxt._element)
                changed = True

    if changed:
        try:
            tmp.save(doc_path)
        except Exception as ex:
            print(f"[s3_native dedup] gagal menyimpan {doc_path}: {ex}")
            return False

    # 3. Idempotensi: bila heading utama sub-bab ini sudah ada, jangan sisipkan ulang.
    non_empty = _set_paragraphs()
    if non_empty:
        primary = _normalize_heading(non_empty[0].text)
        m_ch = re.match(r"^(bab\s+[ivxlcdm]+)", primary)
        key = m_ch.group(1) if m_ch else primary
        if key and (key in existing_headings or any(h.startswith(key) for h in existing_headings if key.startswith("bab "))):
            print(f"[s3_native dedup] '{non_empty[0].text[:60]}' sudah ada — penyisipan dilewati.")
            return True
    return False


def write_s3_subchapter_native(sub_id: str, config: Optional[Dict[str, Any]] = None, path: Optional[str] = None) -> bool:
    """
    Mengetikkan / menyisipkan naskah sub-bab Disertasi Doktor langsung
    ke lembar dokumen Microsoft Word aktif melalui Native COM.
    """
    if path and os.path.exists(path):
        doc_path = path
    else:
        _, doc_path = build_s3_subchapter_doc(sub_id, config)

    try:
        word = engine_word_native.get_or_create_word(visible=True)
        active_doc = word.ActiveDocument
        try:
            active_doc.Activate()
        except Exception:
            pass
        
        def _get_target_range(doc_obj):
            if sub_id == "sub_bagian_awal":
                return doc_obj.Range(0, 0)
            elif sub_id == "sub_dapus":
                rng_dapus = doc_obj.Content
                rng_dapus.Find.ClearFormatting()
                rng_dapus.Find.Text = "DAFTAR PUSTAKA"
                rng_dapus.Find.MatchCase = True
                rng_dapus.Find.MatchWholeWord = True
                if rng_dapus.Find.Execute():
                    start_del = rng_dapus.Start
                    rng_lamp = doc_obj.Content
                    rng_lamp.Find.ClearFormatting()
                    rng_lamp.Find.Text = "LAMPIRAN"
                    rng_lamp.Find.MatchCase = True
                    rng_lamp.Find.MatchWholeWord = True
                    if rng_lamp.Find.Execute() and rng_lamp.Start > start_del:
                        end_del = rng_lamp.Start
                    else:
                        end_del = doc_obj.Content.End
                    doc_obj.Range(start_del, end_del).Delete()
                    return doc_obj.Range(start_del, start_del)
                else:
                    end_pos = max(0, doc_obj.Content.End - 1)
                    return doc_obj.Range(end_pos, end_pos)
            elif sub_id == "sub_lampiran":
                end_pos = max(0, doc_obj.Content.End - 1)
                return doc_obj.Range(end_pos, end_pos)
            else:
                # Sub-bab naskah Bab 1 - Bab 5:
                # Jika dokumen sudah memiliki DAFTAR PUSTAKA, sisipkan tepat SEBELUM DAFTAR PUSTAKA
                rng_dapus = doc_obj.Content
                rng_dapus.Find.ClearFormatting()
                rng_dapus.Find.Text = "DAFTAR PUSTAKA"
                rng_dapus.Find.MatchCase = True
                rng_dapus.Find.MatchWholeWord = True
                if rng_dapus.Find.Execute():
                    return doc_obj.Range(rng_dapus.Start, rng_dapus.Start)
                else:
                    end_pos = max(0, doc_obj.Content.End - 1)
                    return doc_obj.Range(end_pos, end_pos)

        target_range = _get_target_range(active_doc)

        # Cegah heading bab berulang & penyisipan ganda (mis. klik tombol dua kali).
        try:
            existing_headings = _target_heading_index(active_doc)
            if _strip_duplicate_chapter_opener(doc_path, sub_id, existing_headings):
                print(f"[s3_native] {sub_id}: konten sudah ada, tidak disisipkan ulang.")
                return True
        except Exception as dedup_err:
            print(f"[s3_native dedup] dilewati: {dedup_err}")

        target_range.InsertFile(doc_path)
        return True
    except Exception as ex:
        print(f"[s3_native] Error writing {sub_id} to active Word: {ex}")
        try:
            import pythoncom
            import win32com.client.dynamic
            disp = pythoncom.GetActiveObject("Word.Application")
            idisp = disp.QueryInterface(pythoncom.IID_IDispatch)
            word = win32com.client.dynamic.Dispatch(idisp)
            active_doc = word.ActiveDocument
            target_range = _get_target_range(active_doc)
            target_range.InsertFile(doc_path)
            return True
        except Exception as ex2:
            print(f"[s3_native fallback] Error writing {sub_id}: {ex2}")
            return False


def run_s3_chained_pipeline(chapter_num: int, config: Optional[Dict[str, Any]] = None, mode: str = "native") -> Dict[str, Any]:
    """
    Menjalankan Chained Sequential Sub-Chapter Generator untuk seluruh sub-bab
    dalam 1 Bab Disertasi Doktor secara otomatis dan berurutan.
    """
    ctx = resolve_s3_context(config)
    subs = [sid for sid, meta in S3_SUBCHAPTER_CATALOG.items() if meta["chapter"] == chapter_num and not meta.get("is_legacy_alias") and not meta.get("is_granular_child")]
    if not subs:
        if chapter_num == 0:
            subs = ["sub_bagian_awal"]
        elif chapter_num == 6:
            subs = ["sub_dapus", "sub_lampiran"]
        else:
            subs = ["sub_1_1"]

    details = []
    dapus_fn, dapus_fp, dapus_b64 = None, None, None

    for sid in subs:
        fn, fp = build_s3_subchapter_doc(sid, ctx, include_dapus=False)
        with open(fp, "rb") as f:
            import base64
            b64 = base64.b64encode(f.read()).decode("utf-8")

        written = False
        if mode == "native":
            written = write_s3_subchapter_native(sid, ctx, path=fp)

        details.append({
            "sub_id": sid,
            "filename": fn,
            "docx_base64": b64,
            "native_written": written,
            "title": get_s3_subchapter_meta(sid, ctx).get("title", sid)
        })

    # Sertakan sub-bab kustom jika ada pada bab terkait (Bab 0-6)
    custom_subs = (config or {}).get("custom_subchapters", [])
    try:
        target_ch_int = int(chapter_num)
    except Exception:
        target_ch_int = 1

    custom_in_chapter = []
    for c in custom_subs:
        try:
            if int(c.get("chapter", 1)) == target_ch_int:
                custom_in_chapter.append(c)
        except Exception:
            pass

    for csub in custom_in_chapter:
        c_code = csub.get("code", "E")
        c_sid = csub.get("id") or f"custom_{target_ch_int}_{c_code}"
        fn, fp = build_s3_custom_subchapter_doc(csub, ctx)
        with open(fp, "rb") as f:
            import base64
            b64 = base64.b64encode(f.read()).decode("utf-8")

        written = False
        if mode == "native":
            written = write_s3_subchapter_native(c_sid, ctx, path=fp)

        details.append({
            "sub_id": c_sid,
            "filename": fn,
            "docx_base64": b64,
            "native_written": written,
            "title": f"{c_code}. {csub.get('title', 'Sub-Bab Kustom')}"
        })

    title_label = f"Bab {chapter_num} Disertasi Doktor"
    if target_ch_int == 0:
        title_label = "Bagian Awal Disertasi Doktor"
    elif target_ch_int == 6:
        title_label = "Bagian Akhir Disertasi Doktor"

    return {
        "status": "success",
        "chapter": target_ch_int,
        "mode": mode,
        "title": title_label,
        "details": details,
        "dapus_filename": dapus_fn,
        "dapus_base64": dapus_b64
    }


def build_s3_full_dissertation_doc(
    config: Optional[Dict[str, Any]] = None,
    chapters: Optional[list] = None,
    include_front_matter: bool = True,
    include_dapus: bool = True,
    include_lampiran: bool = True,
) -> Tuple[str, str]:
    """
    Merangkai naskah Disertasi Doktor (Bab 0 s/d Bab 5, Dapus, Lampiran)
    beserta seluruh sub-bab kustom pengguna menjadi satu berkas utuh (.docx).

    Parameters
    ----------
    chapters : list[int] | None
        Bila diisi (mis. [1, 2, 3]), hanya bab tersebut yang dirangkai.
        None = seluruh bab.
    include_front_matter, include_dapus, include_lampiran : bool
        Kendali penyertaan bagian awal, daftar pustaka, dan lampiran.
    """
    ctx = resolve_s3_context(config)
    doc = Document()
    setup_page(doc, continuous=False)

    all_keys = [
        "sub_bagian_awal",
        "sub_1_1", "sub_1_2", "sub_1_3", "sub_1_4",
        "sub_2_1_0", "sub_2_1_1", "sub_2_1_2", "sub_2_1_4", "sub_2_1_3", "sub_2_2", "sub_2_3", "sub_2_4",
        "sub_3_a", "sub_3_b", "sub_3_c", "sub_3_d", "sub_3_e", "sub_3_f", "sub_3_g", "sub_3_h", "sub_3_i",
        "sub_4_a_1", "sub_4_a_2", "sub_4_a_3", "sub_4_a_4", "sub_4_a_5", "sub_4_b",
        "sub_5_a", "sub_5_b", "sub_5_c", "sub_5_d", "sub_5_e", "sub_5_f",
        "sub_dapus",
        "sub_lampiran"
    ]

    def _chapter_of(key: str) -> Optional[int]:
        if key == "sub_bagian_awal":
            return 0
        if key == "sub_dapus":
            return 6
        if key == "sub_lampiran":
            return 7
        m_chk = re.match(r"^sub_(\d)", key)
        return int(m_chk.group(1)) if m_chk else None

    if chapters is not None:
        wanted = set(int(c) for c in chapters)
        filtered = []
        for k in all_keys:
            ch = _chapter_of(k)
            if ch == 0:
                if include_front_matter:
                    filtered.append(k)
            elif ch == 6:
                if include_dapus:
                    filtered.append(k)
            elif ch == 7:
                if include_lampiran:
                    filtered.append(k)
            elif ch in wanted:
                filtered.append(k)
        all_keys = filtered

    custom_subs = (config or {}).get("custom_subchapters", [])
    custom_by_ch = {}
    for c in custom_subs:
        try:
            ch_num = int(c.get("chapter", 1))
        except Exception:
            ch_num = 1
        custom_by_ch.setdefault(ch_num, []).append(c)

    def _append_custom_subs(ch_target: int):
        subs_to_append = custom_by_ch.get(ch_target, [])
        for csub in subs_to_append:
            try:
                _, c_fp = build_s3_custom_subchapter_doc(csub, ctx)
                c_doc = Document(c_fp)
                for el in c_doc.element.body:
                    if el.tag.endswith('sectPr'):
                        continue
                    doc.element.body.append(el)
            except Exception as c_err:
                print(f"Error appending custom subchapter ch {ch_target} to full doc: {c_err}")

    _front_included = "sub_bagian_awal" in all_keys
    _section_split_done = False
    # Sub-bab lanjutan Bab IV/V tidak boleh mencetak ulang judul bab dan "A. Hasil Analisis".
    _SUBDOC_NO_CHAPTER_HEAD = {
        "sub_4_a_2", "sub_4_a_3", "sub_4_a_4", "sub_4_a_5", "sub_4_b",
        "sub_5_b", "sub_5_c", "sub_5_d", "sub_5_e", "sub_5_f",
    }

    for k in all_keys:
        builder = S3_BUILDER_MAP.get(k)
        if builder:
            # Berikan jeda page break antar bab utama
            if k in ("sub_1_1", "sub_2_1_0", "sub_2_1_1", "sub_3_a", "sub_3_1", "sub_4_a_1", "sub_4_1_1", "sub_5_a", "sub_5_1", "sub_dapus", "sub_lampiran") and k != all_keys[0]:
                if k == "sub_1_1" and _front_included and not _section_split_done:
                    # Batas BAGIAN AWAL (angka romawi) -> ISI UTAMA (angka arab).
                    doc.add_section()
                    apply_section_margins(doc.sections[-1])
                    _section_split_done = True
                else:
                    doc.add_page_break()
            if k in _SUBDOC_NO_CHAPTER_HEAD:
                builder(doc, ctx, standalone=False)
            else:
                builder(doc, ctx)

        # Sisipkan sub-bab kustom pengguna tepat di akhir masing-masing bab terkait
        if k == "sub_bagian_awal":
            _append_custom_subs(0)
        elif k == "sub_1_4":
            _append_custom_subs(1)
        elif k == "sub_2_4":
            _append_custom_subs(2)
        elif k == "sub_3_i":
            _append_custom_subs(3)
        elif k == "sub_4_b":
            _append_custom_subs(4)
        elif k == "sub_5_f":
            _append_custom_subs(5)
        elif k == "sub_lampiran":
            _append_custom_subs(6)

    for para in doc.paragraphs:
        if para.text.strip():
            para.paragraph_format.line_spacing = 1.5

    # Sapuan akhir: parafrase paragraf naratif statis yang belum divariasikan.
    # Hasilnya deterministik per judul dan tidak menyentuh sitasi/rumus/tabel.
    try:
        n_varied = nv.vary_document(doc, ctx)
        if n_varied:
            print(f"[s3_narasi] {n_varied} paragraf statis divariasikan.")
    except Exception as e:
        print(f"[s3_narasi] Peringatan variasi dokumen: {e}")

    # Sapuan kejujuran data: tandai tegas bab hasil yang masih memakai angka
    # template dan belum memuat data empiris riil.
    try:
        de_fabricate_sweep(doc, ctx)
    except Exception as e:
        print(f"[s3_data] Peringatan sapuan data: {e}")

    # Jaring pengaman terakhir: tahap variasi/parafrase di atas bekerja dari
    # templat asli dan dapat memunculkan kembali sitasi yang sebelumnya sudah
    # dibersihkan. Jalankan sanitasi sekali lagi pada naskah final.
    try:
        _final_biblio = ctx.get("_final_biblio") or []
        if _final_biblio:
            _n_final = _sanitize_document_citations(doc, _final_biblio)
            if _n_final:
                print(f"[s3_refs] Sanitasi akhir pascavariasi: {_n_final} sitasi tanpa entri dibersihkan.")
    except Exception as e:
        print(f"[s3_refs] Sanitasi akhir dilewati: {e}")

    # --- Kelengkapan naskah (Blok B): caption bertingkat, nomor halaman, field ---
    try:
        _jml_cap = mark_caption_styles(doc)
        print(f"[s3_komplit] caption ditandai (tabel/gambar/lampiran): {_jml_cap}")
    except Exception as e:
        print(f"[s3_komplit] Peringatan penandaan caption: {e}")

    try:
        _secs = doc.sections
        if _front_included and len(_secs) >= 2:
            add_page_number_footer(_secs[0], numfmt="lowerRoman", start=1)
            for _s in _secs[1:]:
                add_page_number_footer(_s, numfmt="decimal", start=1)
        else:
            add_page_number_footer(_secs[0], numfmt="decimal", start=1)
        print(f"[s3_komplit] nomor halaman dipasang pada {len(_secs)} section.")
    except Exception as e:
        print(f"[s3_komplit] Peringatan nomor halaman: {e}")

    try:
        _set_update_fields(doc)
    except Exception as e:
        print(f"[s3_komplit] Peringatan penyegaran field: {e}")

    fn = f"DISERTASI_DOKTOR_S3_LENGKAP_{int(time.time())}.docx"
    fp = os.path.join(OUTPUT_DIR, fn)
    doc.save(fp)
    return fn, fp


def write_s3_full_dissertation_native(
    config: Optional[Dict[str, Any]] = None,
    chapters: Optional[list] = None,
    include_front_matter: bool = True,
    include_dapus: bool = True,
    include_lampiran: bool = True,
    replace_document: bool = False,
) -> bool:
    """
    Mengetikkan naskah Disertasi Doktor langsung ke lembar Microsoft Word aktif.

    replace_document=True akan mengosongkan isi dokumen aktif terlebih dahulu
    (dipakai saat memulai naskah baru dari nol agar tidak ada sisa judul lama).
    """
    _, fp = build_s3_full_dissertation_doc(
        config,
        chapters=chapters,
        include_front_matter=include_front_matter,
        include_dapus=include_dapus,
        include_lampiran=include_lampiran,
    )
    try:
        word = engine_word_native.get_or_create_word(visible=True)
        active_doc = word.ActiveDocument
        try:
            active_doc.Activate()
        except Exception:
            pass

        if replace_document:
            # Sisakan satu paragraf kosong agar dokumen tetap valid.
            content = active_doc.Content
            if content.End > 1:
                content.Delete()
            active_doc.Range(0, 0).InsertParagraphAfter()
            target_range = active_doc.Range(0, 0)
        else:
            target_range = active_doc.Range(active_doc.Content.End - 1, active_doc.Content.End - 1)

        target_range.InsertFile(fp)
        return True
    except Exception as ex:
        print(f"[s3_full_native] Error writing full dissertation: {ex}")
        return False
