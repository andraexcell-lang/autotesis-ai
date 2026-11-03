import os
import sys
import json
import urllib.request

BASE_DIR = r"E:\SKILL UNTUK AI\system_prompts_leaks-main\word-addin-tesismagister\backend"
sys.path.insert(0, BASE_DIR)

import s1_hierarchical_engine
import s1_data_pipeline

def test_api_call(endpoint, payload):
    url = f"http://127.0.0.1:5050{endpoint}"
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))

def verify_multi_x():
    print("=" * 60)
    print("STARTING MULTI-X S1 VERIFICATION")
    print("=" * 60)

    test_cases = [
        {
            "n_x": 2,
            "title": "PENGARUH KEPEMIMPINAN TRANSFORMASIONAL DAN BUDAYA ORGANISASI TERHADAP KINERJA PEGAWAI PADA DINAS PENDIDIKAN KOTA PALEMBANG",
            "x_names": ["Kepemimpinan Transformasional", "Budaya Organisasi"],
            "instansi": "Dinas Pendidikan Kota Palembang"
        },
        {
            "n_x": 3,
            "title": "PENGARUH KEPEMIMPINAN TRANSFORMASIONAL, BUDAYA ORGANISASI, DAN MOTIVASI KERJA TERHADAP KINERJA PEGAWAI PADA DINAS PENDIDIKAN KOTA PALEMBANG",
            "x_names": ["Kepemimpinan Transformasional", "Budaya Organisasi", "Motivasi Kerja"],
            "instansi": "Dinas Pendidikan Kota Palembang"
        },
        {
            "n_x": 4,
            "title": "PENGARUH KEPEMIMPINAN TRANSFORMASIONAL, BUDAYA ORGANISASI, MOTIVASI KERJA, DAN DISIPLIN KERJA TERHADAP KINERJA PEGAWAI PADA DINAS PENDIDIKAN KOTA PALEMBANG",
            "x_names": ["Kepemimpinan Transformasional", "Budaya Organisasi", "Motivasi Kerja", "Disiplin Kerja"],
            "instansi": "Dinas Pendidikan Kota Palembang"
        },
        {
            "n_x": 5,
            "title": "PENGARUH KEPEMIMPINAN TRANSFORMASIONAL, BUDAYA ORGANISASI, MOTIVASI KERJA, DISIPLIN KERJA, DAN KOMPETENSI TERHADAP KINERJA PEGAWAI PADA DINAS PENDIDIKAN KOTA PALEMBANG",
            "x_names": ["Kepemimpinan Transformasional", "Budaya Organisasi", "Motivasi Kerja", "Disiplin Kerja", "Kompetensi Pegawai"],
            "instansi": "Dinas Pendidikan Kota Palembang"
        }
    ]

    for tc in test_cases:
        k = tc["n_x"]
        print(f"\n>>> TESTING {k}-X VARIABLES: {tc['x_names']}")
        
        cfg = {
            "jenjang": "S1",
            "degree": "Sarjana",
            "judul": tc["title"],
            "title": tc["title"],
            "instansi": tc["instansi"],
            "n_x": k,
            "x_names": tc["x_names"],
            "x_vars": [{"code": f"X{i+1}", "name": name} for i, name in enumerate(tc["x_names"])],
            "y_name": "Kinerja Pegawai",
            "populasi": 75,
            "sampel": 60,
            "n": 60,
            "tahun": "2026",
            "kota": "Palembang"
        }
        for i, name in enumerate(tc["x_names"]):
            cfg[f"x{i+1}_name"] = name

        # 1. Test Endpoint Generate Respondents
        print(f"  [1] Calling /api/s1/data/generate-respondents for k={k}...")
        gen_res = test_api_call("/api/s1/data/generate-respondents", {"config": cfg})
        assert gen_res["status"] == "success", f"generate-respondents failed: {gen_res}"
        print(f"      OK -> File: {gen_res['data_file']}, Sample: {gen_res['n']}")

        # 2. Test Endpoint Olah Data
        print(f"  [2] Calling /api/s1/data/olah-data for k={k}...")
        olah_res = test_api_call("/api/s1/data/olah-data", {"config": cfg})
        assert olah_res["status"] == "success", f"olah-data failed: {olah_res}"
        spss = olah_res["spss"]
        summ = olah_res["summary"]
        
        # Verify Statistical Properties
        print(f"      R2: {summ['r2']}")
        print(f"      F-stat: {summ['f_stat']}, Sig: {summ['f_sig']}")
        print(f"      Normality Sig: {summ['normality_sig']}")
        print(f"      Max VIF: {summ['vif_max']}")
        
        # Check F-test
        assert float(spss["regression"]["anova_f"]["f_hitung"]) > float(spss["regression"]["anova_f"]["f_tabel"]), "F-stat must be > F-table"
        assert float(spss["regression"]["anova_f"]["sig"]) < 0.05, "F-sig must be < 0.05"
        
        # Check t-test for each X
        coefs = spss["regression"]["coefficients"]
        assert len(coefs) == k + 1, f"Expected {k+1} regression coefficients, got {len(coefs)}"
        for idx in range(1, k + 1):
            c = coefs[idx]
            t_val = float(c["t_hitung"])
            t_tab = float(spss["t_table"])
            p_val = float(c["sig"])
            print(f"      Coef {c['model']}: t={t_val} > {t_tab}, p={p_val}")
            assert t_val > t_tab, f"t-statistic for {c['model']} must be > t-table ({t_val} <= {t_tab})"
            assert p_val < 0.05, f"p-value for {c['model']} must be < 0.05 ({p_val})"
            
        # Check Multicollinearity
        vif_items = spss["classic_assumptions"]["multicollinearity"]["items"]
        assert len(vif_items) == k, f"Expected {k} VIF items, got {len(vif_items)}"
        for vi in vif_items:
            assert vi["vif"] < 10.0, f"VIF for {vi['var_name']} must be < 10 ({vi['vif']})"
            assert vi["tolerance"] > 0.10, f"Tolerance for {vi['var_name']} must be > 0.10 ({vi['tolerance']})"

        # Check Heteroscedasticity (Glejser)
        het_items = spss["classic_assumptions"]["heteroscedasticity"]["items"]
        assert len(het_items) == k, f"Expected {k} Glejser items, got {len(het_items)}"
        for hi in het_items:
            assert hi["sig"] > 0.05, f"Glejser Sig for {hi['var_name']} must be > 0.05 ({hi['sig']})"

        # 3. Test Full DOCX Build
        print(f"  [3] Building full skripsi DOCX for k={k}...")
        fname, doc_path = s1_hierarchical_engine.build_s1_full_skripsi_doc(cfg)
        assert os.path.exists(doc_path), f"Generated file not found: {doc_path}"
        file_sz = os.path.getsize(doc_path)
        print(f"      DOCX Success -> {fname} ({file_sz:,} bytes)")
        assert file_sz > 50000, "DOCX file suspiciously small"

        # 4. Check Subchapter Generation (Bab I, II, III, IV, V)
        print(f"  [4] Verifying individual chapter builds for k={k}...")
        for ch in ["bab1", "bab2", "bab3", "bab4", "bab5"]:
            ch_name, ch_doc = s1_hierarchical_engine.build_s1_subchapter_doc(ch, cfg)
            assert os.path.exists(ch_doc), f"Chapter {ch} doc not found: {ch_doc}"
            print(f"      Chapter {ch} -> {ch_name} ({os.path.getsize(ch_doc):,} bytes)")

    print("\n" + "=" * 60)
    print("ALL MULTI-X S1 TESTS PASSED SUCCESSFULLY!")
    print("=" * 60)

if __name__ == "__main__":
    verify_multi_x()
