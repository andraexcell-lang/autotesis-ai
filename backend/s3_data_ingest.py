# -*- coding: utf-8 -*-
"""Ingesti data empiris riil (SmartPLS/SPSS) untuk Bab IV, Bab V, dan Lampiran S3.

Prinsip utama modul ini adalah **tidak mengarang data**. Mesin hanya menampilkan
angka empiris apabila pengguna benar-benar menyediakannya. Bila data belum ada,
Bab IV dirender dalam ``mode template``: sel angka diisi placeholder dan sumber
tabel ditandai jelas sebagai belum diolah, sehingga naskah tidak pernah
menyajikan angka fiktif seolah-olah hasil penelitian.

Kontrak data (semua kunci opsional; yang tersedia saja yang dipakai)::

    {
      "meta": {"n": 110, "source": "Output SmartPLS 4, diolah peneliti", "collected": "2026-03"},
      "demographics": {
        "gender": [["Laki-laki", 64], ["Perempuan", 46]],
        "age":    [["<= 30 Tahun", 20], ["31 - 40 Tahun", 48]],
        "grade":  [["Golongan II", 18], ["Golongan III", 68]]
      },
      "descriptives": {
        "X1": {"items": [
          {"kode": "X1.1", "indikator": "kecakapan teknis", "min": 1, "max": 5,
           "median": 4, "mean": 4.21, "sd": 0.63}
        ]}
      },
      "outer_model": {
        "X1": {"loadings": [{"code": "X1.1", "loading": 0.812}],
               "ave": 0.63, "composite_reliability": 0.87,
               "cronbach_alpha": 0.81, "vif": 1.60}
      },
      "inner_model": {
        "r_squares": {"M": {"r_square": 0.55, "q_square": 0.38}},
        "srmr": 0.061, "nfi": 0.923
      },
      "hypotheses": [
        {"code": "H1", "original_sample": 0.318, "sample_mean": 0.316,
         "stdev": 0.098, "t_statistics": 3.245, "p_value": 0.001,
         "f2_effect_size": 0.142, "decision": "Diterima"}
      ],
      "mediation": [{"code": "H7", "indirect": 0.121, "ci_low": 0.032, "ci_high": 0.226}],
      "moderation": [{"code": "H9", "interaction": 0.184, "t_statistics": 2.61, "p_value": 0.009}]
    }

Sumber data dapat diberikan sebagai ``dict`` (kunci konfigurasi ``empirical_data``)
atau sebagai berkas JSON/CSV (kunci ``empirical_data_file``).
"""

from __future__ import annotations

import csv
import hashlib
import json
import os
import re
from typing import Any, Dict, List, Optional

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, "..", "output"))

PLACEHOLDER = "\u2014"  # em dash

TEMPLATE_BANNER = (
    "MODE TEMPLATE \u2014 angka pada tabel dan narasi bab ini belum berasal dari "
    "data empiris riil. Sel angka ditampilkan sebagai placeholder (" + PLACEHOLDER + ") "
    "sampai konfigurasi penelitian memuat kunci \"empirical_data\" atau "
    "\"empirical_data_file\" berisi hasil olahan SmartPLS 4 / SPSS yang sebenarnya."
)

GENERATED_BANNER = (
    "CATATAN KEJUJURAN DATA: Seluruh tabel hasil analisis, pengujian hipotesis, dan pembahasan "
    "pada Bab IV dan Bab V ini disusun menggunakan data simulasi sintetis terstandarisasi "
    "(hasil pembangkitan mesin internal, bukan data empiris lapangan) guna memverifikasi "
    "keseluruhan alur metodologis dan model struktural doktoral (Hayes Process Model 15)."
)

SOURCE_TEMPLATE = (
    "Sumber: mode template \u2014 menunggu data empiris riil (belum diolah peneliti)"
)
SOURCE_EMPIRICAL_FMT = "Sumber: {src}"


# ---------------------------------------------------------------------------
# Pemuatan & normalisasi
# ---------------------------------------------------------------------------

def _read_structured_file(path: str) -> Dict[str, Any]:
    """Baca berkas JSON/CSV menjadi struktur dict sederhana."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".json":
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)

    if ext in (".csv", ".tsv"):
        delim = "\t" if ext == ".tsv" else ","
        with open(path, "r", encoding="utf-8-sig", newline="") as fh:
            rows = list(csv.DictReader(fh, delimiter=delim))
        return _normalize_csv_rows(rows)

    # XLSX bersifat opsional: hanya dipakai bila openpyxl tersedia.
    if ext in (".xlsx", ".xlsm"):
        try:
            import openpyxl  # type: ignore
        except Exception as exc:  # pragma: no cover - bergantung lingkungan
            raise RuntimeError(
                "Berkas Excel memerlukan paket openpyxl. Gunakan JSON/CSV atau pasang openpyxl."
            ) from exc
        wb = openpyxl.load_workbook(path, data_only=True)
        out: Dict[str, Any] = {}
        for ws in wb.worksheets:
            header = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
            rows = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row is None or all(v is None for v in row):
                    continue
                rows.append({header[i]: row[i] for i in range(len(header)) if i < len(row)})
            out[ws.title.lower()] = _normalize_csv_rows(rows)
        return out

    raise RuntimeError(f"Format berkas data tidak didukung: {ext}")


def _normalize_csv_rows(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Terima CSV datar dan kelompokkan berdasarkan kolom ``table``.

    Kolom yang dikenali: table, code, label, value, mean, sd, min, max,
    median, loading, ave, cr, cronbach, vif, t, p, decision, source, n.
    """
    grouped: Dict[str, List[Dict[str, Any]]] = {}
    for r in rows:
        table = str(r.get("table") or r.get("tabel") or "descriptives").strip().lower()
        grouped.setdefault(table, []).append(r)
    return grouped


def load_empirical_data(config: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """Ambil data empiris dari konfigurasi (inline dict, berkas terstruktur, atau cache analisis)."""
    cfg = config or {}
    raw = cfg.get("empirical_data")

    # 1. Bila belum ada di config, coba baca dari berkas yang ditentukan
    if raw is None:
        path = cfg.get("empirical_data_file") or cfg.get("data_file")
        if path:
            cand_paths = [str(path)]
            if not os.path.isabs(str(path)):
                cand_paths.append(os.path.join(OUTPUT_DIR, str(path)))
                cand_paths.append(os.path.join(os.path.dirname(OUTPUT_DIR), str(path)))
            for cp in cand_paths:
                if os.path.exists(cp):
                    try:
                        read_res = _read_structured_file(cp)
                        # Validasi apakah berkas memuat struktur analisis lengkap
                        if isinstance(read_res, dict) and ("outer_model" in read_res or "hypotheses" in read_res):
                            raw = read_res
                            break
                    except Exception:
                        pass

    # 2. Bila belum ada atau raw belum memuat hasil analisis (mis. berkas mentah responden),
    # cari cache hasil olah data empiris S3 yang tersimpan di output/
    if not raw or (isinstance(raw, dict) and not raw.get("outer_model") and not raw.get("hypotheses")):
        raw_title = str(cfg.get("judul") or cfg.get("title") or "").strip()

        # 2a. Cache spesifik hash judul
        if raw_title:
            title_hash = hashlib.sha256(raw_title.encode("utf-8")).hexdigest()[:12]
            hash_cache = os.path.join(OUTPUT_DIR, f"s3_analysis_cache_{title_hash}.json")
            if os.path.exists(hash_cache):
                try:
                    with open(hash_cache, "r", encoding="utf-8") as f:
                        raw = json.load(f)
                except Exception:
                    pass

        # 2b. Cache utama s3_analysis_cache.json
        if not raw or (isinstance(raw, dict) and not raw.get("outer_model") and not raw.get("hypotheses")):
            gen_cache = os.path.join(OUTPUT_DIR, "s3_analysis_cache.json")
            if os.path.exists(gen_cache):
                try:
                    with open(gen_cache, "r", encoding="utf-8") as f:
                        cand = json.load(f)
                        if isinstance(cand, dict) and (cand.get("outer_model") or cand.get("hypotheses")):
                            cand_title = str((cand.get("meta") or {}).get("title") or "").strip()
                            clean_cand = re.sub(r"[^a-zA-Z0-9]+", "", cand_title.lower())
                            clean_raw = re.sub(r"[^a-zA-Z0-9]+", "", raw_title.lower())
                            # Muat bila judul cocok atau bila tidak ada judul spesifik yang dibedakan
                            if not clean_cand or not clean_raw or clean_cand == clean_raw:
                                raw = cand
                except Exception:
                    pass

    if not raw:
        return None
    return normalize_empirical_data(raw)


def normalize_empirical_data(raw: Any) -> Dict[str, Any]:
    """Pastikan struktur data memiliki bentuk yang dapat dipakai builder."""
    if not isinstance(raw, dict):
        return {}
    data: Dict[str, Any] = dict(raw)
    data.setdefault("meta", {})
    for key in ("demographics", "descriptives", "outer_model", "inner_model"):
        if key in data and not isinstance(data[key], dict):
            data[key] = {}
    for key in ("hypotheses", "mediation", "moderation"):
        if key in data and isinstance(data[key], dict):
            data[key] = list(data[key].values())
    return data


def get(ctx: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    return ctx.get("empirical_data")


def is_template(ctx: Dict[str, Any]) -> bool:
    """True bila naskah belum memakai data empiris riil."""
    return not ctx.get("empirical_data")


def source_label(ctx: Dict[str, Any], default: str = "Data primer diolah peneliti") -> str:
    """Label sumber tabel yang jujur sesuai mode data."""
    if is_template(ctx):
        return SOURCE_TEMPLATE
    meta = (get(ctx) or {}).get("meta") or {}
    origin = str(meta.get("origin") or "").strip().lower()
    if origin == "generated":
        col = meta.get("collected") or meta.get("year") or "2026"
        return f"Sumber: Data simulasi sintetis (uji mesin internal, bukan data empiris lapangan), {col}"
    src = str(meta.get("source") or default).strip() or default
    return SOURCE_EMPIRICAL_FMT.format(src=src)


# ---------------------------------------------------------------------------
# Aksesor per tabel
# ---------------------------------------------------------------------------

def _as_pairs(seq: Any) -> Optional[List[List[Any]]]:
    if not seq:
        return None
    out = []
    for row in seq:
        if isinstance(row, dict):
            label = row.get("label") or row.get("kategori") or row.get("name")
            freq = row.get("freq", row.get("value", row.get("jumlah")))
        elif isinstance(row, (list, tuple)) and len(row) >= 2:
            label, freq = row[0], row[1]
        else:
            continue
        if label is None or freq is None:
            continue
        out.append([str(label), freq])
    return out or None


def demographics(ctx: Dict[str, Any], key: str) -> Optional[List[List[Any]]]:
    """Frekuensi demografis riil (gender/age/grade) atau None bila belum ada."""
    data = get(ctx)
    if not data:
        return None
    return _as_pairs((data.get("demographics") or {}).get(key))


def descriptives(ctx: Dict[str, Any], key: str) -> Optional[Dict[str, Any]]:
    """Statistik deskriptif butir riil untuk satu konstruk atau None."""
    data = get(ctx)
    if not data:
        return None
    payload = (data.get("descriptives") or {}).get(key)
    if isinstance(payload, dict) and payload.get("items"):
        return payload
    return None


def outer(ctx: Dict[str, Any], key: str) -> Optional[Dict[str, Any]]:
    data = get(ctx)
    if not data:
        return None
    return (data.get("outer_model") or {}).get(key)


def inner(ctx: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    data = get(ctx)
    if not data:
        return None
    return data.get("inner_model") or None


def hypotheses_by_code(ctx: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    data = get(ctx)
    if not data:
        return {}
    out: Dict[str, Dict[str, Any]] = {}
    for h in data.get("hypotheses") or []:
        if isinstance(h, dict) and h.get("code"):
            out[str(h["code"]).strip().upper()] = h
    return out


def records(ctx: Dict[str, Any], key: str) -> List[Dict[str, Any]]:
    data = get(ctx)
    if not data:
        return []
    seq = data.get(key) or []
    return [r for r in seq if isinstance(r, dict)]


def _derive_decision(h: Dict[str, Any]) -> Optional[str]:
    """Turunkan keputusan hipotesis dari t/p bila pengguna tidak mengisinya eksplisit."""
    if h.get("decision"):
        return None
    try:
        t = float(h.get("t_statistics"))
        p = float(h.get("p_value"))
    except (TypeError, ValueError):
        return None
    if h.get("ci_low") is not None and h.get("ci_high") is not None:
        try:
            lo, hi = float(h["ci_low"]), float(h["ci_high"])
            return "Diterima" if (lo > 0 and hi > 0) or (lo < 0 and hi < 0) else "Tidak Diterima"
        except (TypeError, ValueError):
            pass
    return "Diterima" if (t >= 1.96 and p < 0.05) else "Tidak Diterima"


def merge_into_pls(base_pls: Dict[str, Any], data: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """Timpa hasil simulasi dengan nilai riil yang tersedia (tanpa mengubah skema)."""
    if not data or not isinstance(base_pls, dict):
        return base_pls
    merged = dict(base_pls)

    # Outer model: pertahankan struktur, ganti angka yang disediakan.
    real_outer = data.get("outer_model") or {}
    if real_outer:
        sim_outer = dict(merged.get("outer_model") or {})
        for k, v in real_outer.items():
            if isinstance(v, dict) and k in sim_outer and isinstance(sim_outer[k], dict):
                sim_outer[k].update(v)
            else:
                sim_outer[k] = v
        merged["outer_model"] = sim_outer

    # Inner model (R-square/Q-square/SRMR/NFI).
    real_inner = data.get("inner_model") or {}
    if real_inner:
        merged["inner_model"] = {**(merged.get("inner_model") or {}), **real_inner}
        if isinstance(real_inner.get("r_squares"), dict):
            merged["r_squares"] = {
                **(merged.get("r_squares") or {}), **real_inner["r_squares"]
            }

    # Pengujian hipotesis: cocokkan berdasarkan kode.
    real_h = hypotheses_by_code({"empirical_data": data})
    if real_h:
        sim_h = list(merged.get("hypotheses_testing") or [])
        for h in sim_h:
            code = str(h.get("code") or "").strip().upper()
            if code in real_h:
                h.update({k: v for k, v in real_h[code].items() if v is not None})
                derived = _derive_decision(h)
                if derived:
                    h["decision"] = derived
            else:
                # Cegah kebocoran angka simulasi acak ke naskah empiris
                for num_key in (
                    "original_sample", "sample_mean", "stdev",
                    "t_statistics", "p_value", "ci_low", "ci_high", "f2_effect_size"
                ):
                    h[num_key] = None
                h["decision"] = "Belum diuji"
        seen = {str(h.get("code") or "").strip().upper() for h in sim_h}
        for code, h in real_h.items():
            if code not in seen:
                if h.get("decision") is None:
                    d = _derive_decision(h)
                    if d:
                        h = {**h, "decision": d}
                sim_h.append(h)
        merged["hypotheses_testing"] = sim_h
    else:
        # Data empiris disediakan tetapi tidak ada blok hypotheses
        sim_h = list(merged.get("hypotheses_testing") or [])
        for h in sim_h:
            for num_key in (
                "original_sample", "sample_mean", "stdev",
                "t_statistics", "p_value", "ci_low", "ci_high", "f2_effect_size"
            ):
                h[num_key] = None
            h["decision"] = "Belum diuji"
        merged["hypotheses_testing"] = sim_h

    # Bersihkan sisa angka acak fit_indices jika empiris tidak menyediakannya
    if not data.get("fit_indices") and not (data.get("inner_model") or {}).get("srmr"):
        merged["fit_indices"] = {}

    # Blok tambahan: salin apa adanya bila disediakan.
    for key in (
        "mediation", "moderation", "model_fit", "htmT", "htmt",
        "moderation_details", "conditional_indirect_details",
        "fit_indices", "specific_indirect", "total_indirect",
        "hypotheses_testing",
    ):
        if data.get(key):
            if key == "hypotheses_testing":
                continue
            merged[key] = data[key]

    merged["_empirical"] = True
    return merged
