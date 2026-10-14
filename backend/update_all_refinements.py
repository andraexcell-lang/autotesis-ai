# update_all_refinements.py
import re
import sys

# 1. Update academic_connector.py
acad_path = "e:/SKILL UNTUK AI/system_prompts_leaks-main/word-addin-tesismagister/backend/academic_connector.py"
with open(acad_path, "r", encoding="utf-8") as f:
    acad_code = f.read()

# Expand makro_sdm and mediasi pools
new_makro_sdm = '''    # 17. Rumpun Makro Tata Kelola & Manajemen SDM Sektor Publik
    "makro_sdm": [
        {
            "authors": ["A. Suharyanto", "B. Setiawan"],
            "year": 2023,
            "title": "Evaluasi Reformasi Birokrasi Dalam Meningkatkan Kinerja Pegawai dan Pelayanan Publik",
            "journal": "Jurnal Ekonomi dan Manajemen",
            "volume": "24",
            "issue": "3",
            "pages": "45–58",
            "doi": "https://doi.org/10.62951/jem.v24i3.46"
        },
        {
            "authors": ["R. Pratama"],
            "year": 2025,
            "title": "Karakteristik dan Metode Manajemen SDM dalam Reformasi Pelayanan Sektor Publik",
            "journal": "Jurnal Ilmu Sosial dan Kebijakan Publik",
            "volume": "2",
            "issue": "1",
            "pages": "114–128",
            "doi": "https://doi.org/10.24929/semnasfisip.v2i1.4114"
        },
        {
            "authors": ["J. R. Asthenu"],
            "year": 2025,
            "title": "Pengaruh Motivasi dan Iklim Kerja Terhadap Kinerja Pegawai",
            "journal": "Jurnal Ilmiah Poli Bisnis",
            "volume": "16",
            "issue": "2",
            "pages": "82–93",
            "doi": "https://doi.org/10.30630/jipb.v16i2.1570"
        },
        {
            "authors": ["H. Kurniawan", "E. Kusumawardani"],
            "year": 2024,
            "title": "Transformasi Birokrasi Publik dan Produktivitas Kerja Aparatur Sipil Negara di Daerah",
            "journal": "Jurnal Administrasi Publik Indonesia",
            "volume": "12",
            "issue": "1",
            "pages": "45–58",
            "doi": "https://doi.org/10.24198/japi.v12i1.51234"
        },
        {
            "authors": ["Y. C. Simatupang", "W. Safitri"],
            "year": 2023,
            "title": "Akuntabilitas Tata Kelola Publik dan Efektivitas Kinerja Kelembagaan Sektor Publik",
            "journal": "Jurnal Kebijakan dan Administrasi Publik",
            "volume": "27",
            "issue": "2",
            "pages": "112–126",
            "doi": "https://doi.org/10.22146/jkap.78120"
        },
        {
            "authors": ["R. Hartono", "S. Wulandari"],
            "year": 2024,
            "title": "Implementasi Sistem Akuntabilitas Kinerja Instansi Pemerintah dalam Mengoptimalkan Pelayanan Publik",
            "journal": "Jurnal Ilmu Pemerintahan Suara Khatulistiwa",
            "volume": "9",
            "issue": "1",
            "pages": "67–80",
            "doi": "https://doi.org/10.33701/jipsk.v9i1.3512"
        },
        {
            "authors": ["A. Fauzi", "L. Hakim", "M. Nugroho"],
            "year": 2023,
            "title": "Kapabilitas Dinamis Sumber Daya Manusia Sektor Publik di Era Transformasi Digital",
            "journal": "Jurnal Borneo Administrator",
            "volume": "19",
            "issue": "3",
            "pages": "315–328",
            "doi": "https://doi.org/10.24258/jba.v19i3.1205"
        }
    ]'''

# Replace makro_sdm pool
makro_pattern = re.compile(r'    # 17\. Rumpun Makro Tata Kelola & Manajemen SDM Sektor Publik.*?    \]', re.DOTALL)
acad_code = makro_pattern.sub(new_makro_sdm, acad_code)

# Replace mediasi selection logic so it uses topic_m
old_med_selection = '''    # Slot 10 & 11: Artikel Jurnal Research Gap Mediasi M (Segmen 9) - 2 Paper Berbeda
    p_med_list = search_openalex_verified(f"pengaruh {x1} terhadap {y} mediasi {m}", limit=4)
    paper_med_1 = select_unique_paper(p_med_list, TOPIC_JOURNAL_POOLS["mediasi"])
    paper_med_2 = select_unique_paper(p_med_list, TOPIC_JOURNAL_POOLS["mediasi"])'''

new_med_selection = '''    # Slot 10 & 11: Artikel Jurnal Research Gap Mediasi M (Segmen 9) - 2 Paper Berbeda & Spesifik Topik M
    p_med_list = search_openalex_verified(f"pengaruh {x1} terhadap {y} mediasi {m}", limit=4)
    # Gunakan pool spesifik topik M jika tersedia agar sitasi mediasi 100% selaras variabel M
    m_pool = TOPIC_JOURNAL_POOLS.get(topic_m, TOPIC_JOURNAL_POOLS["mediasi"])
    paper_med_1 = select_unique_paper(p_med_list, m_pool)
    paper_med_2 = select_unique_paper(p_med_list, TOPIC_JOURNAL_POOLS["mediasi"] if m_pool != TOPIC_JOURNAL_POOLS["mediasi"] else TOPIC_JOURNAL_POOLS["makro_sdm"])'''

acad_code = acad_code.replace(old_med_selection, new_med_selection)

# Bump cache version prefix to v4
acad_code = acad_code.replace('cache_key = f"v3_', 'cache_key = f"v4_')

with open(acad_path, "w", encoding="utf-8") as f:
    f.write(acad_code)
print("Updated academic_connector.py (v4) successfully!")

# 2. Update engine_hierarchical_pipeline.py macro paradigm selector
engine_path = "e:/SKILL UNTUK AI/system_prompts_leaks-main/word-addin-tesismagister/backend/engine_hierarchical_pipeline.py"
with open(engine_path, "r", encoding="utf-8") as f:
    engine_code = f.read()

old_macro_head = '''def build_narrative_segments_macro(sektor, peran_strategis, urgensi_kinerja, instansi, y_title, y_lower, acad, seed_val):
    """
    Menghasilkan narasi pembuka Segmen 1 (Urgensi Makro) dan Segmen 2 (Kedudukan Strategis & Y)
    secara dinamis berpijak pada 4 paradigma teoritis yang berbeda (Good Governance, RBV/Human Capital,
    Agile Bureaucracy, SPBE/Digital Transformation) sehingga setiap judul & instansi menghasilkan
    narasi yang unik, organik, dan tidak pernah kloningan satu sama lain.
    """
    mode = seed_val % 4'''

new_macro_head = '''def build_narrative_segments_macro(sektor, peran_strategis, urgensi_kinerja, instansi, y_title, y_lower, acad, seed_val):
    """
    Menghasilkan narasi pembuka Segmen 1 (Urgensi Makro) dan Segmen 2 (Kedudukan Strategis & Y)
    secara dinamis berpijak pada 4 paradigma teoritis yang berbeda (Good Governance, RBV/Human Capital,
    Agile Bureaucracy, SPBE/Digital Transformation) sehingga setiap judul & instansi menghasilkan
    narasi yang unik, organik, dan tidak pernah kloningan satu sama lain.
    """
    s_low = (sektor or "").lower()
    if any(k in s_low for k in ["keuangan", "aset", "perbendaharaan"]):
        mode = 0  # Paradigma 1: Good Public Governance & Public Value (Akuntabilitas fiskal, WTP)
    elif any(k in s_low for k in ["riset", "inovasi", "litbang", "kelitbangan"]):
        mode = 2  # Paradigma 3: Agile Bureaucracy & Thematic Public Service Reform (Kebijakan berbasis bukti, IID)
    elif any(k in s_low for k in ["pajak", "pendapatan", "retribusi"]):
        mode = 1  # Paradigma 2: Resource-Based View (RBV) & Human Capital (Modal manusia sebagai aset strategis PAD)
    elif any(k in s_low for k in ["perbankan", "finansial", "komersial"]):
        mode = 3  # Paradigma 4: Digital Agility & Modern Performance Governance
    elif any(k in s_low for k in ["perizinan", "ptsp", "investasi"]):
        mode = 2  # Paradigma 3: Agile Bureaucracy & Lincah OSS
    else:
        mode = seed_val % 4'''

engine_code = engine_code.replace(old_macro_head, new_macro_head)

with open(engine_path, "w", encoding="utf-8") as f:
    f.write(engine_code)
print("Updated engine_hierarchical_pipeline.py macro paradigm selector successfully!")
