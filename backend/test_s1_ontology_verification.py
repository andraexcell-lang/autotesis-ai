# backend/test_s1_ontology_verification.py
"""
VERIFIKASI INTEGRASI RISET DIMENSI & INDIKATOR MASTER S1
=========================================================
Menguji:
1. Riset ontologi S1 (s1_ontology_researcher) untuk:
   - Variabel bawaan (Kepemimpinan Transformasional, Budaya Organisasi)
   - Variabel dari variable_library (Work-Life Balance, Beban Kerja)
   - Model Multi-X (3 Variabel: X1, X2, X3 -> Y)
   - Model Mediasi (X1, X2 -> M -> Y)
2. Pendaftaran dinamis ke s1_constructs & konsistensi butir Likert.
3. Eksekusi pipeline SPSS S1 dengan variabel hasil riset ontologi (100% Lolos).
4. Pembuatan dokumen DOCX Sub-Bab 3.5 (Tabel Kisi-kisi Instrumen) & Bab IV.
"""

import os
import sys
import json

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import s1_ontology_researcher
import s1_constructs
import s1_data_pipeline
import s1_hierarchical_engine

def run_tests():
    print("=" * 70)
    print("TEST 1: Riset Ontologi S1 dengan Variabel Kompleks (Work-Life Balance & Beban Kerja)")
    print("=" * 70)
    
    title = "PENGARUH WORK-LIFE BALANCE, BEBAN KERJA, DAN BUDAYA ORGANISASI TERHADAP KINERJA PEGAWAI PADA DINAS KESEHATAN KOTA PALEMBANG"
    x_vars = ["Work-Life Balance", "Beban Kerja", "Budaya Organisasi"]
    y_var = "Kinerja Pegawai"
    instansi = "Dinas Kesehatan Kota Palembang"
    
    onto = s1_ontology_researcher.research_s1_ontology(
        title=title,
        x_vars=x_vars,
        y_var=y_var,
        model_type="regression",
        instansi=instansi,
        force_refresh=True
    )
    
    assert onto["status"] == "success", "Status ontologi harus success"
    assert onto["metadata"]["n_x"] == 3, "Jumlah X harus 3"
    assert "x1" in onto["variables"], "X1 harus ada dalam ontologi"
    assert "x2" in onto["variables"], "X2 harus ada dalam ontologi"
    assert "x3" in onto["variables"], "X3 harus ada dalam ontologi"
    assert "y" in onto["variables"], "Y harus ada dalam ontologi"
    
    x1_data = onto["variables"]["x1"]
    print(f"✓ X1 Nama: {x1_data['nama_resmi']}")
    print(f"✓ X1 Ahli/Teori: {x1_data['tokoh_ahli']}")
    print(f"✓ X1 Dimensi ({len(x1_data['dimensi'])}): {[d['nama'] for d in x1_data['dimensi']]}")
    print(f"✓ X1 Engine: {x1_data['engine']}")
    assert len(x1_data["dimensi"]) >= 3, "Harus ada minimal 3 dimensi untuk X1"
    
    x2_data = onto["variables"]["x2"]
    print(f"\n✓ X2 Nama: {x2_data['nama_resmi']}")
    print(f"✓ X2 Ahli/Teori: {x2_data['tokoh_ahli']}")
    print(f"✓ X2 Dimensi ({len(x2_data['dimensi'])}): {[d['nama'] for d in x2_data['dimensi']]}")
    print(f"✓ X2 Engine: {x2_data['engine']}")
    assert len(x2_data["dimensi"]) >= 3, "Harus ada minimal 3 dimensi untuk X2"
    
    print("\n" + "=" * 70)
    print("TEST 2: Verifikasi Memori Dinamis s1_constructs")
    print("=" * 70)
    
    prof_wlb = s1_constructs.get_construct_profile("Work-Life Balance")
    assert prof_wlb["name"] == "Work Life Balance" or "Work-Life Balance" in prof_wlb["name"], "Profil dinamis WLB harus terbaca"
    
    items_wlb = s1_constructs.construct_items("X1", "Work-Life Balance", max_dims=3)
    print(f"✓ Jumlah butir X1 (WLB): {len(items_wlb)}")
    for it in items_wlb[:2]:
        print(f"   [{it['code']}] ({it['dim']}) -> {it['text']}")
    assert len(items_wlb) == 6, "Total butir X1 harus 6 (3 dimensi x 2 butir)"
    
    items_beban = s1_constructs.construct_items("X2", "Beban Kerja", max_dims=3)
    print(f"✓ Jumlah butir X2 (Beban Kerja): {len(items_beban)}")
    for it in items_beban[:2]:
        print(f"   [{it['code']}] ({it['dim']}) -> {it['text']}")
    assert len(items_beban) == 6, "Total butir X2 harus 6 (3 dimensi x 2 butir)"
    
    print("\n" + "=" * 70)
    print("TEST 3: Eksekusi Pipeline Olah Data SPSS S1 dengan Variabel Hasil Riset")
    print("=" * 70)
    
    cfg = {
        "jenjang": "S1",
        "judul": title,
        "instansi": instansi,
        "model_type": "regression",
        "n_x": 3,
        "x_names": x_vars,
        "x_vars": [{"code": f"X{i+1}", "name": v} for i, v in enumerate(x_vars)],
        "y_name": y_var,
        "n": 60,
        "populasi": 75,
        "sampel": 60
    }
    
    res_spss = s1_data_pipeline.run_s1_full_pipeline(cfg)
    summary = res_spss["summary"]
    print(f"✓ Sampel N: {res_spss['n']}")
    print(f"✓ Total Butir: {res_spss['total_butir']} (harus 24: 3*6 X + 6 Y)")
    assert res_spss['total_butir'] == 24, "Total butir harus 24 butir"
    print(f"✓ R2: {summary['r2']}")
    print(f"✓ Uji F: F={summary['f_stat']} (Sig. {summary['f_sig']})")
    print(f"✓ Normalitas Residual Sig: {summary['normality_sig']}")
    print(f"✓ Multikolinearitas Max VIF: {summary['vif_max']}")
    print(f"✓ Keputusan: {summary['decision']}")
    assert summary["f_sig"] < 0.05, "Uji F harus signifikan"
    assert summary["normality_sig"] > 0.05, "Residual harus normal"
    assert summary["vif_max"] < 10.0, "VIF harus bebas multikolinearitas"
    
    print("\n" + "=" * 70)
    print("TEST 4: Pembuatan Sub-Bab 3.5 (Kisi-kisi Instrumen) & Bab 4.2 (Validitas/Reliabilitas)")
    print("=" * 70)
    
    fname_35, path_35 = s1_hierarchical_engine.build_s1_subchapter_doc("sub_3_5", cfg)
    print(f"✓ Dokumen Kisi-kisi Sub 3.5 dibuat: {fname_35} ({os.path.getsize(path_35)} bytes)")
    assert os.path.exists(path_35) and os.path.getsize(path_35) > 1000, "Dokumen sub_3_5 harus berhasil dibuat"
    
    fname_42, path_42 = s1_hierarchical_engine.build_s1_subchapter_doc("sub_4_2", cfg)
    print(f"✓ Dokumen Uji Kualitas Sub 4.2 dibuat: {fname_42} ({os.path.getsize(path_42)} bytes)")
    assert os.path.exists(path_42) and os.path.getsize(path_42) > 1000, "Dokumen sub_4_2 harus berhasil dibuat"
    
    print("\n" + "=" * 70)
    print("TEST 5: Riset Ontologi Model Mediasi (X1, X2 -> M -> Y)")
    print("=" * 70)
    
    title_med = "PENGARUH KEPEMIMPINAN TRANSFORMASIONAL DAN LINGKUNGAN KERJA TERHADAP KINERJA PEGAWAI MELALUI MOTIVASI KERJA SEBAGAI VARIABEL MEDIASI PADA BADAN PENDAPATAN DAERAH"
    onto_med = s1_ontology_researcher.research_s1_ontology(
        title=title_med,
        x_vars=["Kepemimpinan Transformasional", "Lingkungan Kerja"],
        m_var="Motivasi Kerja",
        y_var="Kinerja Pegawai",
        model_type="mediation",
        instansi="Badan Pendapatan Daerah",
        force_refresh=True
    )
    assert onto_med["status"] == "success", "Ontologi mediasi harus sukses"
    assert "m" in onto_med["variables"], "M harus ada dalam ontologi mediasi"
    print(f"✓ M Nama: {onto_med['variables']['m']['nama_resmi']}")
    print(f"✓ M Tokoh: {onto_med['variables']['m']['tokoh_ahli']}")
    print(f"✓ M Total Dimensi: {len(onto_med['variables']['m']['dimensi'])}")
    print(f"✓ Total Indikator Blueprint Mediasi: {onto_med['metadata']['total_indikator']} butir")
    assert onto_med['metadata']['total_indikator'] == 24, "Total indikator mediasi 2X + 1M + 1Y harus 24 butir (4 x 6)"

    print("\n" + "=" * 70)
    print("🎉 SELURUH PENGUJIAN RISET DIMENSI & INDIKATOR MASTER S1 100% SUKSES!")
    print("=" * 70)

if __name__ == "__main__":
    run_tests()
