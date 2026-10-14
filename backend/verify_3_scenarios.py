# verify_3_scenarios.py
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
import engine_hierarchical_pipeline

scenarios = [
    {
        "id": "UJI_1_BPKAD",
        "cfg": {
            "judul": "Pengaruh Organizational Citizenship Behavior dan Motivasi Kerja Terhadap Kinerja Pegawai Melalui Kepuasan Kerja Sebagai Variabel Intervening Pada BPKAD Kabupaten OKI",
            "instansi": "BPKAD Kabupaten OKI",
            "x1": "Organizational Citizenship Behavior (OCB)",
            "x2": "Motivasi Kerja",
            "m": "Kepuasan Kerja",
            "y": "Kinerja Pegawai",
            "penulis": "PENELITI_1"
        }
    },
    {
        "id": "UJI_2_BRIDA",
        "cfg": {
            "judul": "Pengaruh Promosi Jabatan dan Mutasi Kerja Terhadap Kinerja Pegawai Melalui Motivasi Kerja Sebagai Variabel Intervening Pada BRIDA Kabupaten OKI",
            "instansi": "BRIDA Kabupaten OKI",
            "x1": "Promosi Jabatan",
            "x2": "Mutasi Kerja",
            "m": "Motivasi Kerja",
            "y": "Kinerja Pegawai",
            "penulis": "PENELITI_2"
        }
    },
    {
        "id": "UJI_3_BPPD",
        "cfg": {
            "judul": "Pengaruh Pengembangan Karier dan Budaya Organisasi Terhadap Kinerja Pegawai Melalui Kepuasan Kerja Sebagai Variabel Intervening Pada BPPD Kabupaten OKI",
            "instansi": "BPPD Kabupaten OKI",
            "x1": "Pengembangan Karier",
            "x2": "Budaya Organisasi",
            "m": "Kepuasan Kerja",
            "y": "Kinerja Pegawai",
            "penulis": "PENELITI_3"
        }
    }
]

for s in scenarios:
    print(f"\n==================== {s['id']} ====================")
    doc_fn, doc_fp = engine_hierarchical_pipeline.build_subchapter_doc("sub_1_1", s["cfg"], include_dapus=False)
    dapus_fn, dapus_fp = engine_hierarchical_pipeline.build_standalone_dapus_doc(s["cfg"])
    print(f"Sub-1.1 Doc: {doc_fn}")
    print(f"Dapus Doc: {dapus_fn}")

    # Inspect generated paragraphs
    doc = Document(doc_fp)
    p_texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    
    # Check opening paragraph 1 (Segmen 1)
    seg1 = p_texts[2] if len(p_texts) > 2 else ""
    print(f"\n[Paragraf 1]:\n{seg1[:160]}...")
    
    # Check Segmen 5 (M)
    seg_m = [t for t in p_texts if any(k in t.lower() for k in ["kepuasan", "ocb", "promosi", "mutasi", "karier", "motivasi"]) and "diindikasikan" in t.lower()]
    if seg_m:
        print(f"\n[Segmen M]:\n{seg_m[0][:180]}...")
    
    # Check Analisis Tabel 1.2
    seg_t2 = [t for t in p_texts if "Fakta empiris di atas membuktikan" in t]
    if seg_t2:
        print(f"\n[Analisis Tabel 1.2]:\n{seg_t2[0]}")
        # Verify no '. di mana' or '. sebanyak'
        if ". di mana" in seg_t2[0] or ". sebanyak" in seg_t2[0] or ". diperparah" in seg_t2[0]:
            print(">>> ERROR: Still contains punctuation bug!")
        else:
            print(">>> Punctuation check: PASSED (Clean Indonesian commas intact)")

    # Check Segmen 9 (Mediasi)
    seg_med = [t for t in p_texts if "Guna memecahkan kebuntuan" in t]
    if seg_med:
        print(f"\n[Segmen 9 Mediasi]:\n{seg_med[0][:200]}...")
        if "kepemimpinan" in seg_med[0] and "kepemimpinan" not in s["cfg"]["x1"].lower() and "kepemimpinan" not in s["cfg"]["x2"].lower():
            print(">>> ERROR: Still contains hardcoded 'kepemimpinan'!")
        else:
            print(">>> Mediation logic check: PASSED (No hardcoded fake terms)")

    # Inspect Dapus
    dapus_doc = Document(dapus_fp)
    dapus_entries = [p.text.strip() for p in dapus_doc.paragraphs if p.text.strip() and p.text.strip() != "DAFTAR PUSTAKA"]
    print(f"\n[Daftar Pustaka Entries Count: {len(dapus_entries)}]")
    for e in dapus_entries[:3]:
        print(f" - {e[:120]}...")
