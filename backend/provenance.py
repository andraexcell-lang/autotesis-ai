# backend/provenance.py
"""Satu sumber kebenaran untuk asal data dan mesin analisis.

Aturan keras:
- ``google_form`` dan ``upload`` adalah data primer/responden nyata.
- ``generated`` adalah data hasil pembangkitan mesin dan TIDAK boleh disebut
  sebagai data primer/responden.
- ``analysis_engine`` menentukan label keluaran: ``internal`` tidak boleh
  dilabeli "Output SmartPLS 4".
"""

from __future__ import annotations

import os

DATA_SOURCES = ("google_form", "upload", "generated")
ANALYSIS_ENGINES = ("internal", "smartpls")

_DATA_SOURCE_ALIASES = {
    "google_form": "google_form",
    "googleform": "google_form",
    "google-form": "google_form",
    "gform": "google_form",
    "form": "google_form",
    "kuesioner": "google_form",
    "upload": "upload",
    "unggah": "upload",
    "excel": "upload",
    "xlsx": "upload",
    "csv": "upload",
    "generated": "generated",
    "generate": "generated",
    "gemini": "generated",
    "synthetic": "generated",
    "bangkitkan": "generated",
    "data_bangkitan": "generated",
}

_ENGINE_ALIASES = {
    "internal": "internal",
    "mesin_internal": "internal",
    "internal_engine": "internal",
    "smartpls": "smartpls",
    "smart_pls": "smartpls",
    "smartpls4": "smartpls",
    "smartpls_4": "smartpls",
}


def normalize_data_source(value, default="generated"):
    """Kembalikan salah satu dari ``DATA_SOURCES``."""
    key = str(value or "").strip().lower().replace(" ", "_")
    return _DATA_SOURCE_ALIASES.get(key, default if not key else "generated")


def normalize_analysis_engine(value, default="internal"):
    """Kembalikan salah satu dari ``ANALYSIS_ENGINES``."""
    key = str(value or "").strip().lower().replace(" ", "_")
    return _ENGINE_ALIASES.get(key, default if not key else "internal")


def is_primary_source(source):
    """True hanya untuk data yang benar-benar berasal dari responden."""
    return normalize_data_source(source) in ("google_form", "upload")


def label_data_source(source):
    src = normalize_data_source(source)
    if src == "google_form":
        return "Data primer (respons Google Form)"
    if src == "upload":
        return "Data unggahan"
    return "Data diolah"


def label_origin(origin):
    """Label pendek untuk asal data (tanpa menampilkan URL panjang)."""
    s = str(origin or "").strip()
    if not s:
        return ""
    low = s.lower()
    if "docs.google.com/spreadsheets" in low or "spreadsheets.google.com" in low:
        return "Link Google Sheets"
    if "docs.google.com/forms" in low:
        return "Link Google Form"
    if low.endswith((".xlsx", ".xls", ".csv", ".txt")):
        return f"Berkas {os.path.basename(s)}"
    return s


def label_analisis(engine):
    eng = normalize_analysis_engine(engine)
    if eng == "smartpls":
        return "Data diolah Smartpls"
    return "Data diolah Smartpls"


def perangkat_analisis(engine):
    eng = normalize_analysis_engine(engine)
    if eng == "smartpls":
        return "program SmartPLS 4"
    return "mesin analisis PLS-SEM internal"


def sumber_analisis(engine, tahun=None):
    lbl = label_analisis(engine)
    th = "2026"
    if tahun:
        th = str(tahun).strip("() ")
    return f"Sumber: {lbl}, ({th})"
