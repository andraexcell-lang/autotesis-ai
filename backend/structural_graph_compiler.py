# backend/structural_graph_compiler.py
"""
STRUCTURAL GRAPH COMPILER (SGC) - V3.0
=======================================
Modul Kompiler Model Riset Kuantitatif Dinamis.

Mendukung:
  1. Jumlah variabel bebas (X) dinamis (N = 1, 2, 3, 4, ... n).
  2. Model Mediasi Murni (X -> M -> Y).
  3. Model Moderasi Murni (Z memoderasi X -> Y).
  4. Model Mediasi Termoderasi / Hibrida (Moderated Mediation - Hayes Process Model).
  5. Model Mediasi Paralel (X -> M1, M2 -> Y).

MENGHASILKAN:
  - Daftar Node & Edge Terstruktur.
  - Daftar Hipotesis Lengkap (H1 s.d Hm) beserta narasi baku akademik.
  - Persamaan Matematis Regresi Struktural (Persamaan 1, 2, dst).
  - Butir Rumusan Masalah (Sub 1.2) & Tujuan Penelitian (Sub 1.3).
  - Butir Kesimpulan (Sub 5.1).
  - Skema Diagram Jalur SmartPLS.

Modul ini BERSIFAT INDEPENDEN dan TIDAK MENGUBAH file master lama.
"""

from typing import List, Dict, Any, Optional
import json


class ModelType:
    MEDIATION = "mediation"
    MODERATION = "moderation"
    MODERATED_MEDIATION = "moderated_mediation"
    PARALLEL_MEDIATION = "parallel_mediation"


def compile_research_graph(
    model_type: str = ModelType.MEDIATION,
    x_vars: Optional[List[str]] = None,
    m_vars: Optional[List[str]] = None,
    z_vars: Optional[List[str]] = None,
    y_var: str = "Kinerja Pegawai",
    instansi: str = "Instansi Penelitian",
    focus_topic: str = "Manajemen Sumber Daya Manusia"
) -> Dict[str, Any]:
    """
    Mengompilasi spesifikasi variabel menjadi Graph Model terpadu.
    """
    x_vars = [x.strip() for x in (x_vars or ["Kecerdasan Emosional", "Motivasi Kerja"]) if x.strip()]
    m_vars = [m.strip() for m in (m_vars or []) if m.strip()]
    z_vars = [z.strip() for z in (z_vars or []) if z.strip()]
    y_var = y_var.strip() or "Kinerja Pegawai"
    instansi = instansi.strip() or "Instansi Terkait"

    # Default fallback bila user memilih model tertentu tapi list variabel kosong
    if model_type in (ModelType.MEDIATION, ModelType.MODERATED_MEDIATION) and not m_vars:
        m_vars = ["Kepuasan Kerja"]
    if model_type in (ModelType.MODERATION, ModelType.MODERATED_MEDIATION) and not z_vars:
        z_vars = ["Budaya Organisasi"]

    nodes = []
    # Node X
    for i, x in enumerate(x_vars, start=1):
        nodes.append({"id": f"X{i}", "label": f"X{i}", "name": x, "role": "exogenous", "kind": "independen"})
    
    # Node M
    for j, m in enumerate(m_vars, start=1):
        nodes.append({"id": f"M{j}" if len(m_vars) > 1 else "M", "label": f"M{j}" if len(m_vars) > 1 else "M", "name": m, "role": "mediator", "kind": "intervening"})
    
    # Node Z
    for k, z in enumerate(z_vars, start=1):
        nodes.append({"id": f"Z{k}" if len(z_vars) > 1 else "Z", "label": f"Z{k}" if len(z_vars) > 1 else "Z", "name": z, "role": "moderator", "kind": "moderasi"})
    
    # Node Y
    nodes.append({"id": "Y", "label": "Y", "name": y_var, "role": "endogenous", "kind": "dependen"})

    # Kompilasi Hipotesis
    hypotheses = []
    h_idx = 1

    # 1. Jalur Pengaruh Langsung ke Mediator (jika ada M)
    for m_node in [n for n in nodes if n["role"] == "mediator"]:
        for x_node in [n for n in nodes if n["role"] == "exogenous"]:
            hypotheses.append({
                "code": f"H{h_idx}",
                "type": "direct_to_mediator",
                "source": x_node["id"],
                "target": m_node["id"],
                "source_name": x_node["name"],
                "target_name": m_node["name"],
                "relation": f"{x_node['id']} -> {m_node['id']}",
                "statement": f"{x_node['name']} berpengaruh positif dan signifikan terhadap {m_node['name']} pada {instansi}.",
                "question": f"Apakah {x_node['name']} berpengaruh positif dan signifikan terhadap {m_node['name']} pada {instansi}?",
                "objective": f"Untuk menganalisis dan membuktikan pengaruh {x_node['name']} terhadap {m_node['name']} pada {instansi}.",
                "conclusion": f"{x_node['name']} terbukti berpengaruh positif dan signifikan terhadap {m_node['name']} pada {instansi}."
            })
            h_idx += 1

    # 2. Jalur Pengaruh Langsung ke Dependen (Y)
    for x_node in [n for n in nodes if n["role"] == "exogenous"]:
        hypotheses.append({
            "code": f"H{h_idx}",
            "type": "direct_to_dependent",
            "source": x_node["id"],
            "target": "Y",
            "source_name": x_node["name"],
            "target_name": y_var,
            "relation": f"{x_node['id']} -> Y",
            "statement": f"{x_node['name']} berpengaruh positif dan signifikan terhadap {y_var} pada {instansi}.",
            "question": f"Apakah {x_node['name']} berpengaruh positif dan signifikan terhadap {y_var} pada {instansi}?",
            "objective": f"Untuk menganalisis dan membuktikan pengaruh {x_node['name']} terhadap {y_var} pada {instansi}.",
            "conclusion": f"{x_node['name']} terbukti berpengaruh positif dan signifikan terhadap {y_var} pada {instansi}."
        })
        h_idx += 1

    # Jalur Mediator (M) ke Dependen (Y)
    for m_node in [n for n in nodes if n["role"] == "mediator"]:
        hypotheses.append({
            "code": f"H{h_idx}",
            "type": "mediator_to_dependent",
            "source": m_node["id"],
            "target": "Y",
            "source_name": m_node["name"],
            "target_name": y_var,
            "relation": f"{m_node['id']} -> Y",
            "statement": f"{m_node['name']} berpengaruh positif dan signifikan terhadap {y_var} pada {instansi}.",
            "question": f"Apakah {m_node['name']} berpengaruh positif dan signifikan terhadap {y_var} pada {instansi}?",
            "objective": f"Untuk menganalisis dan membuktikan pengaruh {m_node['name']} terhadap {y_var} pada {instansi}.",
            "conclusion": f"{m_node['name']} terbukti berpengaruh positif dan signifikan terhadap {y_var} pada {instansi}."
        })
        h_idx += 1

    # Jalur Moderator (Z) langsung ke Y (Direct main effect of moderator)
    for z_node in [n for n in nodes if n["role"] == "moderator"]:
        hypotheses.append({
            "code": f"H{h_idx}",
            "type": "moderator_main_effect",
            "source": z_node["id"],
            "target": "Y",
            "source_name": z_node["name"],
            "target_name": y_var,
            "relation": f"{z_node['id']} -> Y",
            "statement": f"{z_node['name']} berpengaruh positif dan signifikan terhadap {y_var} pada {instansi}.",
            "question": f"Apakah {z_node['name']} berpengaruh positif dan signifikan terhadap {y_var} pada {instansi}?",
            "objective": f"Untuk menganalisis dan membuktikan pengaruh {z_node['name']} terhadap {y_var} pada {instansi}.",
            "conclusion": f"{z_node['name']} terbukti berpengaruh positif dan signifikan terhadap {y_var} pada {instansi}."
        })
        h_idx += 1

    # 3. Jalur Mediasi Tidak Langsung (Indirect Effect)
    for m_node in [n for n in nodes if n["role"] == "mediator"]:
        for x_node in [n for n in nodes if n["role"] == "exogenous"]:
            hypotheses.append({
                "code": f"H{h_idx}",
                "type": "mediation_indirect",
                "source": x_node["id"],
                "mediator": m_node["id"],
                "target": "Y",
                "source_name": x_node["name"],
                "mediator_name": m_node["name"],
                "target_name": y_var,
                "relation": f"{x_node['id']} -> {m_node['id']} -> Y",
                "statement": f"{x_node['name']} berpengaruh positif dan signifikan terhadap {y_var} melalui {m_node['name']} sebagai variabel intervening pada {instansi}.",
                "question": f"Apakah {x_node['name']} berpengaruh terhadap {y_var} melalui {m_node['name']} sebagai variabel intervening pada {instansi}?",
                "objective": f"Untuk menganalisis dan membuktikan peran mediasi {m_node['name']} dalam pengaruh {x_node['name']} terhadap {y_var} pada {instansi}.",
                "conclusion": f"{m_node['name']} terbukti memediasi secara signifikan pengaruh {x_node['name']} terhadap {y_var} pada {instansi}."
            })
            h_idx += 1

    # 4. Jalur Moderasi (Interaction Effect)
    for z_node in [n for n in nodes if n["role"] == "moderator"]:
        for x_node in [n for n in nodes if n["role"] == "exogenous"]:
            hypotheses.append({
                "code": f"H{h_idx}",
                "type": "moderation_interaction",
                "source": x_node["id"],
                "moderator": z_node["id"],
                "target": "Y",
                "source_name": x_node["name"],
                "moderator_name": z_node["name"],
                "target_name": y_var,
                "relation": f"({x_node['id']} * {z_node['id']}) -> Y",
                "statement": f"{z_node['name']} memperkuat pengaruh positif {x_node['name']} terhadap {y_var} pada {instansi}.",
                "question": f"Apakah {z_node['name']} memoderasi (memperkuat) pengaruh {x_node['name']} terhadap {y_var} pada {instansi}?",
                "objective": f"Untuk menganalisis dan membuktikan peran moderasi {z_node['name']} dalam hubungan antara {x_node['name']} dan {y_var} pada {instansi}.",
                "conclusion": f"{z_node['name']} terbukti memoderasi dan memperkuat pengaruh {x_node['name']} terhadap {y_var} pada {instansi}."
            })
            h_idx += 1

    # 5. Jalur Mediasi Termoderasi Khusus (Conditional Indirect Effect) jika Moderated Mediation
    if model_type == ModelType.MODERATED_MEDIATION:
        for z_node in [n for n in nodes if n["role"] == "moderator"]:
            for m_node in [n for n in nodes if n["role"] == "mediator"]:
                for x_node in [n for n in nodes if n["role"] == "exogenous"]:
                    hypotheses.append({
                        "code": f"H{h_idx}",
                        "type": "conditional_indirect",
                        "source": x_node["id"],
                        "mediator": m_node["id"],
                        "moderator": z_node["id"],
                        "target": "Y",
                        "source_name": x_node["name"],
                        "mediator_name": m_node["name"],
                        "moderator_name": z_node["name"],
                        "target_name": y_var,
                        "relation": f"Index of ModMed ({x_node['id']} -> {m_node['id']} -> Y | {z_node['id']})",
                        "statement": f"Efek tidak langsung {x_node['name']} terhadap {y_var} melalui {m_node['name']} signifikan bersyarat pada tingkat {z_node['name']} (Mediasi Termoderasi) pada {instansi}.",
                        "question": f"Apakah efek mediasi {m_node['name']} dalam hubungan {x_node['name']} terhadap {y_var} dimoderasi secara bersyarat oleh {z_node['name']} pada {instansi}?",
                        "objective": f"Untuk membuktikan efek mediasi bersyarat (conditional indirect effect) dari {m_node['name']} pada berbagai tingkat {z_node['name']} di {instansi}.",
                        "conclusion": f"Terbukti terdapat efek mediasi termoderasi di mana kekuatan mediasi {m_node['name']} meningkat secara signifikan pada kondisi {z_node['name']} yang tinggi."
                    })
                    h_idx += 1

    # Kompilasi Persamaan Matematis Struktural
    equations = []
    # Persamaan Sub-struktural 1 (jika ada M)
    for m_node in [n for n in nodes if n["role"] == "mediator"]:
        terms = [f"γ_{m_node['id']}{x['id']} {x['id']}" for x in nodes if x["role"] == "exogenous"]
        eq1 = f"{m_node['id']} = " + " + ".join(terms) + f" + ζ_{m_node['id']}"
        equations.append({"name": f"Sub-struktural {m_node['id']}", "formula": eq1})

    # Persamaan Struktural Utama (ke Y)
    y_terms = [f"γ_Y{x['id']} {x['id']}" for x in nodes if x["role"] == "exogenous"]
    for m_node in [n for n in nodes if n["role"] == "mediator"]:
        y_terms.append(f"β_Y{m_node['id']} {m_node['id']}")
    for z_node in [n for n in nodes if n["role"] == "moderator"]:
        y_terms.append(f"β_Y{z_node['id']} {z_node['id']}")
        for x_node in [n for n in nodes if n["role"] == "exogenous"]:
            y_terms.append(f"β_interaksi ({x_node['id']} × {z_node['id']})")
    
    eq_main = "Y = " + " + ".join(y_terms) + " + ζ_Y"
    equations.append({"name": "Struktural Utama Y", "formula": eq_main})

    return {
        "model_type": model_type,
        "instansi": instansi,
        "y_var": y_var,
        "x_vars": x_vars,
        "m_vars": m_vars,
        "z_vars": z_vars,
        "num_x": len(x_vars),
        "num_m": len(m_vars),
        "num_z": len(z_vars),
        "total_nodes": len(nodes),
        "total_hypotheses": len(hypotheses),
        "nodes": nodes,
        "hypotheses": hypotheses,
        "equations": equations,
        "meta": {
            "version": "SGC-3.0",
            "compatible_standard": "By SkripsiPlg & SmartPLS 4",
            "compiler": "StructuralGraphCompiler"
        }
    }


if __name__ == "__main__":
    # Test kompilasi berbagai skenario
    print("=== TEST 1: Model 3X Mediasi (10 Hipotesis) ===")
    g1 = compile_research_graph(
        model_type=ModelType.MEDIATION,
        x_vars=["Disiplin Kerja", "Motivasi Kerja", "Kompetensi Pegawai"],
        m_vars=["Kepuasan Kerja"],
        y_var="Kinerja Pegawai",
        instansi="Dinas Perhubungan"
    )
    print(f"Total Hipotesis: {g1['total_hypotheses']}")
    for h in g1["hypotheses"]:
        print(f"  {h['code']}: {h['statement']}")

    print("\n=== TEST 2: Model Moderasi 2X + 1Z + 1Y (6 Hipotesis) ===")
    g2 = compile_research_graph(
        model_type=ModelType.MODERATION,
        x_vars=["Beban Kerja", "Work-Life Balance"],
        z_vars=["Budaya Organisasi"],
        y_var="Kinerja Pegawai",
        instansi="Rumah Sakit Daerah"
    )
    print(f"Total Hipotesis: {g2['total_hypotheses']}")
    for h in g2["hypotheses"]:
        print(f"  {h['code']}: {h['statement']}")

    print("\n=== TEST 3: Model Mediasi Termoderasi (2X + 1M + 1Z + 1Y) ===")
    g3 = compile_research_graph(
        model_type=ModelType.MODERATED_MEDIATION,
        x_vars=["Kepemimpinan Transformasional", "Lingkungan Kerja"],
        m_vars=["Etos Kerja"],
        z_vars=["Budaya Organisasi"],
        y_var="Kinerja Pegawai",
        instansi="BPKAD Kabupaten PALI"
    )
    print(f"Total Hipotesis: {g3['total_hypotheses']}")
    for h in g3["hypotheses"]:
        print(f"  {h['code']}: {h['statement']}")
