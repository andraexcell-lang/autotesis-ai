# backend/compile_andre_thesis.py
"""
Master Compiler Naskah Tesis Lengkap dan Dokumen Per-Bab Andre Juliansya
Format Standar Tesis Pascasarjana Magister Manajemen Universitas Tridinanti Palembang (2026).
"""

import os
from docx import Document
from build_full_thesis_docx import setup_page, build_bagian_awal
from thesis_bab1_bab2 import build_bab_1, build_bab_2
from thesis_bab3_bab4 import build_bab_3, build_bab_4
from thesis_bab5_lampiran import build_bab_5, build_daftar_pustaka, build_lampiran

OUTPUT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

def compile_all():
    print("Memulai kompilasi naskah master Tesis Andre Juliansya...")
    
    # 1. TESIS LENGKAP UTUH
    doc_full = Document()
    setup_page(doc_full)
    build_bagian_awal(doc_full)
    build_bab_1(doc_full)
    build_bab_2(doc_full)
    build_bab_3(doc_full)
    build_bab_4(doc_full)
    build_bab_5(doc_full)
    build_daftar_pustaka(doc_full)
    build_lampiran(doc_full)
    
    full_path = os.path.join(OUTPUT_DIR, "TESIS_LENGKAP_ANDRE_JULIANSYA.docx")
    doc_full.save(full_path)
    print(f"--> Berhasil membuat: {full_path}")

    # 2. DOKUMEN PER-BAB
    # Bagian Awal
    doc_awal = Document()
    setup_page(doc_awal)
    build_bagian_awal(doc_awal)
    doc_awal.save(os.path.join(OUTPUT_DIR, "BAGIAN_AWAL_ANDRE_JULIANSYA.docx"))
    print("--> Berhasil membuat: BAGIAN_AWAL_ANDRE_JULIANSYA.docx")

    # Bab 1
    doc_b1 = Document()
    setup_page(doc_b1)
    build_bab_1(doc_b1)
    doc_b1.save(os.path.join(OUTPUT_DIR, "BAB_1_PENDAHULUAN_ANDRE_JULIANSYA.docx"))
    print("--> Berhasil membuat: BAB_1_PENDAHULUAN_ANDRE_JULIANSYA.docx")

    # Bab 2
    doc_b2 = Document()
    setup_page(doc_b2)
    build_bab_2(doc_b2)
    doc_b2.save(os.path.join(OUTPUT_DIR, "BAB_2_KAJIAN_PUSTAKA_ANDRE_JULIANSYA.docx"))
    print("--> Berhasil membuat: BAB_2_KAJIAN_PUSTAKA_ANDRE_JULIANSYA.docx")

    # Bab 3
    doc_b3 = Document()
    setup_page(doc_b3)
    build_bab_3(doc_b3)
    doc_b3.save(os.path.join(OUTPUT_DIR, "BAB_3_METODE_PENELITIAN_ANDRE_JULIANSYA.docx"))
    print("--> Berhasil membuat: BAB_3_METODE_PENELITIAN_ANDRE_JULIANSYA.docx")

    # Bab 4
    doc_b4 = Document()
    setup_page(doc_b4)
    build_bab_4(doc_b4)
    doc_b4.save(os.path.join(OUTPUT_DIR, "BAB_4_HASIL_DAN_PEMBAHASAN_ANDRE_JULIANSYA.docx"))
    print("--> Berhasil membuat: BAB_4_HASIL_DAN_PEMBAHASAN_ANDRE_JULIANSYA.docx")

    # Bab 5 + Dapus
    doc_b5 = Document()
    setup_page(doc_b5)
    build_bab_5(doc_b5)
    build_daftar_pustaka(doc_b5)
    doc_b5.save(os.path.join(OUTPUT_DIR, "BAB_5_KESIMPULAN_DAN_DAPUS_ANDRE_JULIANSYA.docx"))
    print("--> Berhasil membuat: BAB_5_KESIMPULAN_DAN_DAPUS_ANDRE_JULIANSYA.docx")

    # Lampiran
    doc_lamp = Document()
    setup_page(doc_lamp)
    build_lampiran(doc_lamp)
    doc_lamp.save(os.path.join(OUTPUT_DIR, "LAMPIRAN_LENGKAP_ANDRE_JULIANSYA.docx"))
    print("--> Berhasil membuat: LAMPIRAN_LENGKAP_ANDRE_JULIANSYA.docx")

    print("\nSeluruh berkas naskah Word Tesis Magister berhasil dikompilasi 100% sempurna!")

if __name__ == "__main__":
    compile_all()
