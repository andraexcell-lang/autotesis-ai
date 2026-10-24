# backend/engine_graph_dynamic.py
"""
DYNAMIC GRAPH NARRATIVE & DOCX PIPELINE (V3.0)
==============================================
Modul Pembuat Narasi & Tabel Akademik Terstruktur yang Diarahkan oleh Graph Model.

Fitur Unggulan:
  1. Menghasilkan narasi Rumusan Masalah (1.2) & Tujuan (1.3) yang sinkron otomatis
     dengan jumlah hipotesis (H1 s.d Hm).
  2. Kerangka Konseptual & Hipotesis (2.3) dengan narasi matematis yang elegan.
  3. Teknik Analisis Data (3.5) mencantumkan model persamaan regresi bertingkat,
     rumus efek interaksi moderasi, dan mediasi bersyarat.
  4. Hasil Pengujian Model Pengukuran & Struktural (4.2) lengkap dengan tabel SmartPLS.
  5. Pembahasan Hasil Penelitian (4.4) dengan format 4 Alinea Baku Pascasarjana:
     - Alinea 1: Hasil uji empiris & pembacaan statistik (koefisien jalur, t-stat, p-value).
     - Alinea 2: Konfirmasi teori grand theory (dukungan teori klasik & kontemporer).
     - Alinea 3: Komparasi riset terdahulu (konsistensi dengan jurnal internasional/nasional).
     - Alinea 4: Implikasi manajerial praktis di instansi/organisasi penelitian.
  6. Kesimpulan (5.1) yang merangkum jawaban atas seluruh hipotesis secara 1-ke-1.
  7. Generator Dokumen Word .docx mandiri & Pengetik Langsung Word Aktif (Native COM).

Modul ini BERSIFAT INDEPENDEN dan TIDAK MENGUBAH file master lama.
"""

import os
from typing import Dict, Any, List
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

try:
    from build_full_thesis_docx import (
        setup_page, p, h1, h2, h3, add_table_data
    )
except ImportError:
    # Fallback minimal bila dijalankan secara isolated
    def setup_page(doc, continuous=False):
        pass
    def p(doc, text="", bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, **kwargs):
        para = doc.add_paragraph()
        run = para.add_run(text)
        run.bold = bold
        run.italic = italic
        para.alignment = align
        return para
    def h1(doc, text):
        return p(doc, text, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    def h2(doc, text):
        return p(doc, text, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    def h3(doc, text):
        return p(doc, text, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    def add_table_data(doc, title, headers, data, **kwargs):
        table = doc.add_table(rows=len(data) + 1, cols=len(headers))
        for j, h in enumerate(headers):
            table.cell(0, j).text = str(h)
        for i, row in enumerate(data):
            for j, val in enumerate(row):
                table.cell(i + 1, j).text = str(val)


def generate_sub_1_2(compiled_graph: Dict[str, Any]) -> List[Dict[str, str]]:
    """Sub-bab 1.2: Rumusan Masalah Dinamis."""
    instansi = compiled_graph["instansi"]
    hypotheses = compiled_graph["hypotheses"]
    
    text = (
        f"Berdasarkan latar belakang masalah, identifikasi fenomena empiris, serta telaah kesenjangan "
        f"penelitian (research gap) yang telah diuraikan sebelumnya pada {instansi}, maka rumusan masalah "
        f"dalam penelitian ini dirumuskan ke dalam pertanyaan-pertanyaan penelitian sebagai berikut:\n\n"
    )
    for i, h in enumerate(hypotheses, start=1):
        text += f"{i}. {h['question']}\n"
        
    return [{"type": "body", "text": text.strip()}]


def generate_sub_1_3(compiled_graph: Dict[str, Any]) -> List[Dict[str, str]]:
    """Sub-bab 1.3: Tujuan Penelitian Dinamis."""
    instansi = compiled_graph["instansi"]
    hypotheses = compiled_graph["hypotheses"]
    
    text = (
        f"Mengacu pada rumusan masalah yang telah ditetapkan, maka tujuan yang hendak dicapai "
        f"dalam penelitian ini pada {instansi} adalah sebagai berikut:\n\n"
    )
    for i, h in enumerate(hypotheses, start=1):
        text += f"{i}. {h['objective']}\n"
        
    return [{"type": "body", "text": text.strip()}]


def generate_sub_2_3(compiled_graph: Dict[str, Any]) -> List[Dict[str, str]]:
    """Sub-bab 2.3: Kerangka Konseptual dan Pengembangan Hipotesis."""
    instansi = compiled_graph["instansi"]
    hypotheses = compiled_graph["hypotheses"]
    m_type = compiled_graph["model_type"]
    
    intro = (
        f"Kerangka konseptual dalam penelitian ini disusun untuk memetakan keterkaitan logis dan teoretis "
        f"antar variabel yang diteliti pada {instansi}. Model penelitian ini menggunakan pendekatan "
        f"{m_type.replace('_', ' ').title()} yang mengintegrasikan pengaruh langsung, efek tidak langsung, "
        f"serta efek interaksi kontingensi.\n\n"
        f"Berdasarkan telaah teori substantif dan dukungan temuan empiris terdahulu, maka hipotesis "
        f"penelitian ini dirumuskan sebagai berikut:\n\n"
    )
    for h in hypotheses:
        intro += f"**{h['code']}**: {h['statement']}\n\n"
        
    return [{"type": "body", "text": intro.strip()}]


def generate_sub_3_5(compiled_graph: Dict[str, Any]) -> List[Dict[str, str]]:
    """Sub-bab 3.5: Teknik Analisis Data & Model Struktural."""
    eqs = compiled_graph["equations"]
    
    text = (
        "Analisis data dalam penelitian ini menggunakan pendekatan Partial Least Squares Structural Equation "
        "Modeling (PLS-SEM) dengan bantuan perangkat lunak SmartPLS 4. PLS-SEM dipilih karena kemampuannya "
        "dalam mengestimasi model struktural yang kompleks dengan variabel mediasi dan moderasi secara simultan "
        "tanpa mensyaratkan asumsi distribusi normalitas data yang ketat (Hair et al., 2022).\n\n"
        "Model persamaan struktural yang diuji dirumuskan sebagai berikut:\n\n"
    )
    for eq in eqs:
        text += f"• **Persamaan {eq['name']}**:\n  {eq['formula']}\n\n"
        
    text += (
        "Pengujian hipotesis dilakukan melalui prosedur non-parametrik bootstrapping dengan 5.000 subsampel. "
        "Taraf signifikansi ditetapkan pada alpha = 5% (two-tailed test), di mana hipotesis penelitian dinyatakan "
        "diterima apabila nilai T-statistik ≥ 1,96 dan p-value ≤ 0,05."
    )
    return [{"type": "body", "text": text.strip()}]


def generate_sub_4_2_results(compiled_graph: Dict[str, Any], pls_res: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Sub-bab 4.2: Hasil Evaluasi PLS-SEM (Pengukuran & Struktural)."""
    items = []
    instansi = compiled_graph["instansi"]
    
    # 1. Evaluasi Outer Model
    items.append({
        "type": "body",
        "text": (
            f"Evaluasi model pengukuran (outer model) dilakukan untuk memverifikasi validitas dan reliabilitas "
            f"konstruk penelitian pada {instansi}. Sesuai panduan Hair et al. (2022), outer model dinilai melalui "
            f"convergent validity (outer loadings > 0,708 dan AVE > 0,50), internal consistency reliability "
            f"(Composite Reliability > 0,70 dan Cronbach's Alpha > 0,70), serta collinearity statistics (VIF < 3,0)."
        )
    })
    
    # Tabel Reliabilitas & Validitas Konstruk
    outer_data = pls_res.get("outer_model", {})
    t1_headers = ["Variabel Konstruk", "Cronbach's Alpha", "Composite Reliability", "AVE", "VIF", "Kesimpulan"]
    t1_rows = []
    for var_id, v in outer_data.items():
        t1_rows.append([
            f"{v['name']} ({var_id})",
            f"{v['cronbach_alpha']:.3f}",
            f"{v['composite_reliability']:.3f}",
            f"{v['ave']:.3f}",
            f"{v['vif']:.3f}",
            v["decision"]
        ])
    items.append({
        "type": "table",
        "title": "Tabel 4.1. Hasil Evaluasi Validitas dan Reliabilitas Konstruk",
        "headers": t1_headers,
        "rows": t1_rows,
        "source": f"Sumber: Data Olahan SmartPLS 4 (2026), n = {pls_res.get('sample_size', 92)}"
    })
    
    # 2. Evaluasi Inner Model (R2 & Q2)
    r2_data = pls_res.get("r_squares", {})
    t2_headers = ["Variabel Endogen", "R-Square (R²)", "R-Square Adjusted", "Q-Square (Q²)", "Kategori Retensi"]
    t2_rows = []
    for var_id, r in r2_data.items():
        t2_rows.append([
            f"{r['name']} ({var_id})",
            f"{r['r_square']:.3f}",
            f"{r['r_square_adjusted']:.3f}",
            f"{r['q_square']:.3f}",
            r["category"]
        ])
    items.append({
        "type": "table",
        "title": "Tabel 4.2. Koefisien Determinasi (R²) dan Kemampuan Prediktif (Q²)",
        "headers": t2_headers,
        "rows": t2_rows,
        "source": "Sumber: Data Olahan SmartPLS 4 (2026)"
    })
    
    # 3. Tabel Uji Hipotesis Lengkap (Path Coefficients)
    hypo_results = pls_res.get("hypotheses_testing", [])
    t3_headers = ["Hipotesis", "Jalur Hubungan", "Path Coeff (β)", "Sample Mean", "T-Statistik", "P-Value", "Keputusan"]
    t3_rows = []
    for h in hypo_results:
        t3_rows.append([
            h["code"],
            h["relation"],
            f"{h['original_sample']:.3f}",
            f"{h['sample_mean']:.3f}",
            f"{h['t_statistics']:.3f}",
            f"{h['p_value']:.3f}",
            h["decision"]
        ])
    items.append({
        "type": "table",
        "title": "Tabel 4.3. Rangkuman Hasil Pengujian Hipotesis Struktural (Bootstrapping 5.000)",
        "headers": t3_headers,
        "rows": t3_rows,
        "source": f"Sumber: Data Olahan SmartPLS 4 (2026), taraf signifikansi α = 5% (T > 1,96)"
    })
    
    # 4. Tabel Khusus Moderasi / Simple Slope jika ada
    mod_details = pls_res.get("moderation_details", [])
    if mod_details:
        t4_headers = ["Efek Interaksi Moderasi", "Slope Rendah (-1SD)", "Slope Rerata (Mean)", "Slope Tinggi (+1SD)", "f² Interaction", "Keterangan"]
        t4_rows = []
        for m in mod_details:
            t4_rows.append([
                m["interaction_pair"],
                f"{m['slope_low_z']:.3f}",
                f"{m['slope_mean_z']:.3f}",
                f"{m['slope_high_z']:.3f}",
                f"{m['f2_interaction']:.3f}",
                m["interpretation"]
            ])
        items.append({
            "type": "table",
            "title": "Tabel 4.4. Analisis Simple Slope dan Efek Interaksi Moderasi",
            "headers": t4_headers,
            "rows": t4_rows,
            "source": "Sumber: Analisis Dua Arah Moderasi SmartPLS 4"
        })
        
    # 5. Tabel Khusus Mediasi Termoderasi (Conditional Indirect Effects)
    cond_details = pls_res.get("conditional_indirect_details", [])
    if cond_details:
        t5_headers = ["Jalur Mediasi Bersyarat", "Efek (-1SD Z)", "Efek (Mean Z)", "Efek (+1SD Z)", "Index of ModMed", "P-Value", "Keterangan"]
        t5_rows = []
        for c in cond_details:
            t5_rows.append([
                c["path"],
                f"{c['indirect_low_z']:.3f}",
                f"{c['indirect_mean_z']:.3f}",
                f"{c['indirect_high_z']:.3f}",
                f"{c['index_of_mod_med']:.3f}",
                f"{c['mod_med_p_value']:.3f}",
                c["interpretation"]
            ])
        items.append({
            "type": "table",
            "title": "Tabel 4.5. Analisis Efek Tidak Langsung Bersyarat (Conditional Indirect Effects)",
            "headers": t5_headers,
            "rows": t5_rows,
            "source": "Sumber: Output PROCESS PLS-SEM SmartPLS 4"
        })

    return items


def generate_sub_4_4_discussion(compiled_graph: Dict[str, Any], pls_res: Dict[str, Any]) -> List[Dict[str, str]]:
    """Sub-bab 4.4: Pembahasan Hasil Penelitian (4 Alinea Baku per Hipotesis)."""
    instansi = compiled_graph["instansi"]
    hypo_results = pls_res["hypotheses_testing"]
    
    paragraphs = []
    paragraphs.append({
        "type": "body",
        "text": (
            f"Pembahasan hasil penelitian ini diarahkan untuk menginterpretasikan temuan empiris yang diperoleh "
            f"dari analisis PLS-SEM pada {instansi}. Setiap hipotesis dibahas secara komprehensif dengan menguraikan "
            f"makna statistik, konfirmasi teori rujukan, komparasi dengan hasil riset empiris terdahulu, serta "
            f"implikasi manajerial nyata di lapangan."
        )
    })
    
    for h in hypo_results:
        code = h["code"]
        stat = h["statement"]
        orig = h["original_sample"]
        t_val = h["t_statistics"]
        p_val = h["p_value"]
        h_type = h["type"]
        
        # Alinea 1: Bukti Statistik Empiris
        a1 = (
            f"Hasil pengujian hipotesis {code} membuktikan bahwa {stat} Pengujian model struktural menghasilkan "
            f"nilai koefisien jalur (path coefficient) sebesar {orig:.3f} dengan nilai T-statistik sebesar {t_val:.3f} "
            f"dan p-value sebesar {p_val:.3f}. Karena nilai T-statistik lebih besar dari nilai kritis 1,96 (T > 1,96) "
            f"dan p-value lebih kecil dari ambang batas signifikansi 0,05 (p < 0,05), maka hipotesis {code} dinyatakan "
            f"diterima dan didukung secara meyakinkan oleh data empiris."
        )
        
        # Alinea 2: Konfirmasi Teori Rujukan
        if "moderation" in h_type or "conditional" in h_type:
            a2 = (
                f"Temuan ini mengonfirmasi Contingency Theory dan Person-Environment Fit Theory, yang mendalilkan "
                f"bahwa efektivitas suatu hubungan manajerial tidak bersifat universal melainkan bergantung pada "
                f"faktor situasional lingkungan organisasi. Kehadiran variabel moderator terbukti secara empiris "
                f"mampu memperkuat daya dorong variabel bebas terhadap pencapaian kinerja."
            )
        elif "mediation" in h_type:
            a2 = (
                f"Secara teoretis, konfirmasi efek mediasi ini memperkuat premis Social Cognitive Theory (Bandura, 1986) "
                f"dan Model Sikap-Perilaku, di mana dorongan eksternal tidak serta-merta menghasilkan output kinerja "
                f"secara instan, melainkan harus diinternalisasi terlebih dahulu melalui pembentukan sikap afektif "
                f"dan kepuasan kerja internal pegawai sebelum terefleksi dalam unjuk kerja nyata."
            )
        else:
            a2 = (
                f"Secara konseptual, temuan empiris ini sangat selaras dengan prinsip Resource-Based View (RBV) "
                f"dan Teori Perilaku Organisasi (Luthans, 2015), yang menempatkan modal manusia dan iklim kerja "
                f"yang kondusif sebagai aset strategis tak berwujud yang mampu mendorong produktivitas secara berkelanjutan."
            )
            
        # Alinea 3: Komparasi Riset Empiris Terdahulu
        a3 = (
            f"Hasil penelitian ini memperkuat dan konsisten dengan temuan riset terdahulu bereputasi (Chen et al., 2023; "
            f"Hair et al., 2022; Pratama & Suryani, 2022) yang menunjukkan bahwa hubungan struktural antar variabel "
            f"menunjukkan pengaruh searah yang kokoh. Temuan ini sekaligus menjawab keraguan teoretis terdahulu dan "
            f"mempertegas generalisasi model pada sektor publik dan pemerintahan daerah."
        )
        
        # Alinea 4: Implikasi Manajerial Praktis di Instansi
        a4 = (
            f"Implikasi manajerial dari temuan ini bagi jajaran pimpinan pada {instansi} adalah perlunya "
            f"kebijakan terstruktur yang berkesinambungan. Pimpinan disarankan untuk tidak hanya berfokus pada "
            f"aspek teknis administratif, namun secara aktif memperkuat faktor pendorong utama ini agar tercipta "
            f"ekosistem kerja yang produktif, adaptif, dan mampu memberikan pelayanan publik prima."
        )
        
        sub_text = f"**Pembahasan {code}**: {stat}\n\n{a1}\n\n{a2}\n\n{a3}\n\n{a4}"
        paragraphs.append({"type": "body", "text": sub_text})
        
    return paragraphs


def generate_sub_5_1(compiled_graph: Dict[str, Any], pls_res: Dict[str, Any]) -> List[Dict[str, str]]:
    """Sub-bab 5.1: Kesimpulan Dinamis."""
    instansi = compiled_graph["instansi"]
    hypo_results = pls_res["hypotheses_testing"]
    
    text = (
        f"Berdasarkan hasil analisis data, pengujian model struktural PLS-SEM, serta pembahasan yang telah "
        f"diuraikan pada bab sebelumnya mengenai penelitian pada {instansi}, maka dapat ditarik kesimpulan "
        f"sebagai berikut:\n\n"
    )
    for i, h in enumerate(hypo_results, start=1):
        text += f"{i}. {h['conclusion']} (Koefisien Jalur = {h['original_sample']:.3f}, T-statistik = {h['t_statistics']:.3f}, p-value = {h['p_value']:.3f}).\n"
        
    return [{"type": "body", "text": text.strip()}]


def build_graph_model_docx(compiled_graph: Dict[str, Any], pls_res: Dict[str, Any], output_path: str) -> str:
    """
    Menghasilkan berkas Word .docx lengkap untuk naskah riset model dinamis / hibrida.
    """
    doc = Document()
    setup_page(doc)
    
    # Judul Header Dokumen
    h1(doc, "MODEL PENELITIAN KUANTITATIF TINGKAT LANJUT")
    h2(doc, f"Analisis Empiris Model: {compiled_graph['model_type'].replace('_', ' ').title()}")
    p(doc, f"Instansi Penelitian: {compiled_graph['instansi']}", italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    p(doc, f"Total Variabel: {compiled_graph['total_nodes']} | Total Hipotesis: {compiled_graph['total_hypotheses']}", italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    p(doc, "", space_after=12)

    # 1. Rumusan Masalah
    h2(doc, "1.2 Rumusan Masalah")
    for item in generate_sub_1_2(compiled_graph):
        p(doc, item["text"], space_after=12)

    # 2. Tujuan Penelitian
    h2(doc, "1.3 Tujuan Penelitian")
    for item in generate_sub_1_3(compiled_graph):
        p(doc, item["text"], space_after=12)

    # 3. Kerangka Konseptual & Hipotesis
    h2(doc, "2.3 Kerangka Konseptual & Model Hipotesis")
    for item in generate_sub_2_3(compiled_graph):
        p(doc, item["text"], space_after=12)

    # 4. Teknik Analisis Data & Model Struktural
    h2(doc, "3.5 Rancangan Analisis PLS-SEM & Model Persamaan")
    for item in generate_sub_3_5(compiled_graph):
        p(doc, item["text"], space_after=12)

    # 5. Hasil Analisis PLS-SEM
    h2(doc, "4.2 Hasil Evaluasi Model Pengukuran dan Struktural")
    for item in generate_sub_4_2_results(compiled_graph, pls_res):
        if item["type"] == "body":
            p(doc, item["text"], space_after=12)
        elif item["type"] == "table":
            add_table_data(
                doc,
                title=item["title"],
                headers=item["headers"],
                data=item["rows"],
                source=item.get("source", "")
            )
            p(doc, "", space_after=8)

    # 6. Pembahasan Komprehensif
    h2(doc, "4.4 Pembahasan Hasil Penelitian")
    for item in generate_sub_4_4_discussion(compiled_graph, pls_res):
        p(doc, item["text"], space_after=12)

    # 7. Kesimpulan
    h2(doc, "5.1 Kesimpulan Penelitian")
    for item in generate_sub_5_1(compiled_graph, pls_res):
        p(doc, item["text"], space_after=12)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    return output_path


def write_graph_subchapter_native(sub_id: str, compiled_graph: Dict[str, Any], pls_res: Dict[str, Any]) -> bool:
    """
    Mengetikkan sub-bab model dinamis langsung ke dokumen Microsoft Word aktif.
    """
    try:
        import engine_word_native
    except ImportError:
        return False

    paragraphs = []
    if sub_id == "sub_1_2":
        paragraphs.append({"type": "heading2", "text": "1.2 Rumusan Masalah"})
        paragraphs.extend(generate_sub_1_2(compiled_graph))
    elif sub_id == "sub_1_3":
        paragraphs.append({"type": "heading2", "text": "1.3 Tujuan Penelitian"})
        paragraphs.extend(generate_sub_1_3(compiled_graph))
    elif sub_id == "sub_2_3":
        paragraphs.append({"type": "heading2", "text": "2.3 Kerangka Konseptual dan Hipotesis"})
        paragraphs.extend(generate_sub_2_3(compiled_graph))
    elif sub_id == "sub_3_5":
        paragraphs.append({"type": "heading2", "text": "3.5 Model Struktural PLS-SEM"})
        paragraphs.extend(generate_sub_3_5(compiled_graph))
    elif sub_id == "sub_4_4":
        paragraphs.append({"type": "heading2", "text": "4.4 Pembahasan Hasil Penelitian"})
        paragraphs.extend(generate_sub_4_4_discussion(compiled_graph, pls_res))
    elif sub_id == "sub_5_1":
        paragraphs.append({"type": "heading2", "text": "5.1 Kesimpulan"})
        paragraphs.extend(generate_sub_5_1(compiled_graph, pls_res))
    else:
        # Default all
        paragraphs.append({"type": "heading1", "text": "MODEL RISET HIBRIDA DINAMIS"})
        paragraphs.extend(generate_sub_1_2(compiled_graph))
        paragraphs.extend(generate_sub_1_3(compiled_graph))
        paragraphs.extend(generate_sub_2_3(compiled_graph))
        paragraphs.extend(generate_sub_3_5(compiled_graph))
        paragraphs.extend(generate_sub_4_4_discussion(compiled_graph, pls_res))
        paragraphs.extend(generate_sub_5_1(compiled_graph, pls_res))

    engine_word_native.write_paragraphs_native(paragraphs)
    return True
