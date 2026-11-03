# backend/import_references.py
"""
PIPELINE IMPOR REFERENSI MASSAL (SUMBER DIMENSI & INDIKATOR S1).

Mengimpor banyak file jurnal/buku (PDF) ke mesin S1 dengan cara yang SAMA
persis dengan 4 PDF awal: file PDF disalin ke references/, teksnya diekstrak,
dan entri pustaka ditulis ulang dengan source.url = /references/<file> sehingga
kartu "Sumber" di blueprint meresolusi (tidak error DOI) dan konten dimensi/
indikator berbasis PDF yang benar-benar tersedia di server.

PEMETAAN: references/reference_map.json
  primers[]   -> 1 per variabel; sumber utama dimensi & indikator.
  pendukung[] -> jurnal/buku pembanding Bab II (tidak menimpa sumber primer).

MODE per entri primer:
  develop  -> dimensi dari sumber (entri kurasi yang sudah ada dipertahankan;
              bila belum ada dan API aktif, dikembangkan mesin). Indikator
              DIPERTAHANKAN (blueprint-siap); audit literal tetap dicetak.
  literal  -> SETIAP indikator wajib kutipan LITERAL dari teks PDF
              (anti-halusinasi). Indikator yang tidak cocok DIBUANG.
  auto     -> ekstraksi otomatis penuh oleh mesin variable_tracer (LLM).

PENGGUNAAN:
  python import_references.py import [--map references/reference_map.json]
                                    [--clear-cache] [--dry-run]
  python import_references.py status
  python import_references.py --clear-cache
"""

import argparse
import io
import json
import os
import shutil
import sys
import datetime

BACKEND_DIR = os.path.dirname(os.path.abspath(__file__))
if BACKEND_DIR not in sys.path:
    sys.path.insert(0, BACKEND_DIR)

REFERENCES_DIR = os.path.join(BACKEND_DIR, "references")
PRIMER_DIR = os.path.join(REFERENCES_DIR, "01_primers")
PENDUKUNG_DIR = os.path.join(REFERENCES_DIR, "02_pendukung")
PENDUKUNG_REGISTRY = os.path.join(REFERENCES_DIR, "pendukung.json")
OUTPUT_DIR = os.path.abspath(os.path.join(BACKEND_DIR, "..", "output"))

CACHE_FILES = [
    os.path.join(OUTPUT_DIR, "s1_ontology_cache.json"),
    os.path.join(BACKEND_DIR, "_extraction_cache.json"),
    os.path.join(BACKEND_DIR, "_dimension_source_cache.json"),
]

VALID_MODES = ("develop", "literal", "auto")
VALID_ROLES = ("X", "M", "Y")


def _utf8_stdout():
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Bantuan I/O
# ---------------------------------------------------------------------------

def load_map(map_path):
    with open(map_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError("reference_map.json harus objek JSON.")
    primers = data.get("primers", []) or []
    pendukung = data.get("pendukung", []) or []
    return primers, pendukung


def validate_primers(primers):
    """Satu variabel = satu primer. Menolak duplikat sebelum kerja apa pun."""
    seen = {}
    for i, p in enumerate(primers):
        var = (p.get("variable") or "").strip()
        if not var:
            raise ValueError(f"primer[{i}]: field 'variable' wajib.")
        if var.lower() in seen:
            raise ValueError(
                f"primer[{i}]: variabel '{var}' punya >1 primer "
                f"(duplikat dengan primer[{seen[var.lower()]}]). "
                "Satu variabel = 1 sumber primer.")
        seen[var.lower()] = i
        mode = p.get("mode", "develop")
        if mode not in VALID_MODES:
            raise ValueError(f"primer '{var}': mode '{mode}' tak dikenal. "
                             f"Gunakan {VALID_MODES}.")


def target_pdf_path(entry, base_dir):
    """Path tujuan di dalam references/ untuk sebuah entri map."""
    rel = (entry.get("file") or "").strip().lstrip("/\\")
    if not rel:
        return None
    rel = rel.replace("\\", "/")
    # file yang sudah bersubfolder relatif ke references/
    if "/" in rel:
        return os.path.normpath(os.path.join(REFERENCES_DIR, rel))
    return os.path.join(base_dir, rel)


def ensure_copy(entry, target, dry_run):
    """Salin source -> target bila ada field 'source'; lapor bila target hilang."""
    src = (entry.get("source") or "").strip()
    if src:
        src_path = os.path.normpath(os.path.join(BACKEND_DIR, src))
        if not os.path.exists(src_path) and not os.path.isabs(src):
            # coba relatif ke references/
            src_path = os.path.normpath(os.path.join(REFERENCES_DIR, src))
        if not os.path.exists(src_path):
            raise FileNotFoundError(
                f"source '{src}' tidak ditemukan (dicari di {src_path}).")
        if dry_run:
            print(f"    [dry-run] salin {src} -> {os.path.relpath(target, BACKEND_DIR)}")
            return False
        os.makedirs(os.path.dirname(target), exist_ok=True)
        shutil.copy2(src_path, target)
        return True
    if not os.path.exists(target):
        raise FileNotFoundError(
            f"PDF '{entry.get('file')}' tidak ditemukan dan tidak ada "
            "'source' untuk disalin.")
    return False


def rel_ref_url(entry, base_dir):
    """URL relatif /references/<file> yang dipakai di kartu Sumber."""
    rel = (entry.get("file") or "").strip().lstrip("/\\").replace("\\", "/")
    if not rel:
        return ""
    if "/" not in rel and base_dir is PENDUKUNG_DIR:
        rel = "02_pendukung/" + rel
    if "/" not in rel and base_dir is PRIMER_DIR:
        rel = "01_primers/" + rel
    return "/references/" + rel


def http_ok(url):
    """Proxy cek resolusi: server me-layani /references/ sebagai static, jadi
    cek berkas di disk sudah cukup untuk memastikan HTTP 200 nanti."""
    rel = url.replace("/references/", "", 1) if url.startswith("/references/") else url
    full = os.path.normpath(os.path.join(REFERENCES_DIR, rel))
    return os.path.isfile(full)


# ---------------------------------------------------------------------------
# Audit indikator terhadap teks PDF
# ---------------------------------------------------------------------------

def audit_literal(dimensi, text):
    """Audit kutipan literal: kembalikan (clean, report) memakai mesin yang ada."""
    import variable_library as VL
    clean, report = VL.verify_indicators(dimensi, text)
    return clean, report


# ---------------------------------------------------------------------------
# Impor primer
# ---------------------------------------------------------------------------

def build_source_metadata(entry, url, prior=None):
    prior = prior or {}
    return {
        "type": "pdf_reference",
        "label": entry.get("label") or "Sumber Referensi",
        "apa": entry.get("apa") or entry.get("label") or "Sumber Referensi",
        "year": entry.get("year"),
        "doc_kind": entry.get("doc_kind") or "artikel",
        "journal_title": entry.get("journal") or entry.get("apa") or "",
        "doi": entry.get("doi") or prior.get("doi") or "",
        "url": url,
    }


def import_primer(p, dry_run):
    """Impor satu primer. Mengembalikan dict laporan."""
    var = (p.get("variable") or "").strip()
    mode = p.get("mode", "develop")
    target = target_pdf_path(p, REFERENCES_DIR)  # primer default di root references/
    report = {
        "variable": var, "mode": mode, "file": p.get("file", ""),
        "status": "skip", "dims": 0, "indicators": 0,
        "literal_checked": 0, "literal_passed": 0, "literal_dropped": 0,
        "url": "", "url_ok": False, "saved": "", "note": "",
    }
    if target is None:
        report["status"] = "error"
        report["note"] = "field 'file' kosong."
        return report

    # 1. Salin / cek berkas
    try:
        copied = ensure_copy(p, target, dry_run)
    except FileNotFoundError as e:
        report["status"] = "error"
        report["note"] = str(e)
        return report

    report["url"] = rel_ref_url(p, REFERENCES_DIR)
    report["url_ok"] = http_ok(report["url"])
    if not report["url_ok"]:
        report["status"] = "error"
        report["note"] = "PDF tujuan belum ada di references/."
        return report

    # 2. Ekstrak teks
    from doc_text_extract import extract_pdf_text
    text = extract_pdf_text(target) or ""
    if not text or len(text.strip()) < 300:
        report["status"] = "error"
        report["note"] = f"teks PDF terlalu pendek ({len(text.strip())} char)."
        return report

    # 3. Entri lama
    import variable_library as VL
    existing = VL.load_entry(var)

    # 4. Pilih dimensi sesuai mode
    if mode == "auto" and not (existing or {}).get("dimensi"):
        import variable_tracer as VT
        role = p.get("role", "X") if p.get("role", "X") in VALID_ROLES else "X"
        res = VT.trace_from_upload(target, var, role=role,
                                   judul=p.get("judul", ""),
                                   instansi=p.get("instansi", ""),
                                   display_name=p.get("label") or os.path.basename(target))
        if res.get("status") != "ok" or not (res.get("entry") or {}).get("dimensi"):
            report["status"] = "needs_review"
            report["note"] = (res.get("report", {}).get("note")
                              or "ekstraksi otomatis tidak menghasilkan dimensi.")
            return report
        dims = res["entry"]["dimensi"]
        existing = res["entry"]  # supaya metadata (definisi, teori) ikut terbawa
        audit = {"checked": 0, "passed": 0, "dropped": []}
    elif mode == "literal":
        base = (existing or {}).get("dimensi") or (p.get("dimensi") or [])
        if not base:
            report["status"] = "skip"
            report["note"] = ("mode literal tanpa dimensi: isi field 'dimensi' "
                              "di map atau impor dulu mode develop.")
            return report
        dims, audit = audit_literal(base, text)
    else:  # develop
        base = (existing or {}).get("dimensi") or (p.get("dimensi") or [])
        if base:
            dims = base
            _, audit = audit_literal(base, text)
        else:
            # belum ada kurasi: coba mesin (LLM) sebagai "develop from source"
            import variable_tracer as VT
            role = p.get("role", "X") if p.get("role", "X") in VALID_ROLES else "X"
            res = VT.trace_from_upload(target, var, role=role,
                                       judul=p.get("judul", ""),
                                       instansi=p.get("instansi", ""),
                                       display_name=p.get("label") or os.path.basename(target))
            if res.get("status") == "ok" and (res.get("entry") or {}).get("dimensi"):
                dims = res["entry"]["dimensi"]
                existing = res["entry"]
                audit = {"checked": 0, "passed": 0, "dropped": []}
            else:
                report["status"] = "needs_review"
                report["note"] = ("belum ada kurasi & ekstraksi otomatis tidak "
                                  "mencukupi. Impor ditunda agar entri lama tidak "
                                  "ditimpa data kosong.")
                return report

    if not dims:
        report["status"] = "skip"
        report["note"] = "tidak ada dimensi yang layak; entri lama dipertahankan."
        return report

    n_ind = sum(len(d.get("indikator", []) or []) for d in dims)
    report["dims"] = len(dims)
    report["indicators"] = n_ind
    report["literal_checked"] = audit.get("checked", 0)
    report["literal_passed"] = audit.get("passed", 0)
    report["literal_dropped"] = len(audit.get("dropped", []))
    if mode == "literal" and report["literal_passed"] < 1:
        report["status"] = "skip"
        report["note"] = ("tidak ada indikator literal yang lolos verifikasi "
                          "terhadap teks PDF; entri lama dipertahankan.")
        return report

    # 5. Bangun entri final (pertahankan metadata kurasi lama yang tak ada di map)
    entry = dict(existing or {})
    entry["variable"] = var
    if not entry.get("synonyms"):
        entry["synonyms"] = [var.lower()]
    prior_src = (existing or {}).get("source") or {}
    entry["source"] = build_source_metadata(p, report["url"], prior_src)
    for k in ("grand_theory", "tokoh_ahli", "definisi"):
        if p.get(k):
            entry[k] = p[k]
    entry["dimensi"] = dims
    entry.setdefault("engine", "s1-pdf-reference")
    entry["verified_at"] = datetime.datetime.now().isoformat(timespec="seconds")

    # 6. Simpan
    try:
        path = VL.save_entry(entry)
        report["saved"] = os.path.relpath(path, BACKEND_DIR)
        report["status"] = "ok"
        if mode == "literal" and report["literal_dropped"]:
            report["note"] = (f"{report['literal_dropped']} indikator non-literal "
                              "dibuang (anti-halusinasi).")
    except Exception as e:
        report["status"] = "error"
        report["note"] = f"gagal simpan: {e}"
    return report


# ---------------------------------------------------------------------------
# Impor pendukung (tidak menimpa primer)
# ---------------------------------------------------------------------------

def import_pendukung(p, dry_run):
    """Salin + daftarkan referensi pendukung Bab II ke references/pendukung.json."""
    target = target_pdf_path(p, PENDUKUNG_DIR)
    report = {
        "variable": p.get("variable", ""), "file": p.get("file", ""),
        "status": "skip", "url": "", "url_ok": False, "words": 0, "note": "",
    }
    if target is None:
        report["status"] = "error"
        report["note"] = "field 'file' kosong."
        return report
    try:
        ensure_copy(p, target, dry_run)
    except FileNotFoundError as e:
        report["status"] = "error"
        report["note"] = str(e)
        return report

    report["url"] = rel_ref_url(p, PENDUKUNG_DIR)
    report["url_ok"] = http_ok(report["url"])
    if not report["url_ok"]:
        report["status"] = "error"
        report["note"] = "PDF tujuan belum ada."
        return report

    from doc_text_extract import extract_pdf_text
    text = extract_pdf_text(target) or ""
    report["words"] = len((text or "").split())

    registry = []
    if os.path.exists(PENDUKUNG_REGISTRY):
        try:
            with open(PENDUKUNG_REGISTRY, "r", encoding="utf-8") as f:
                registry = (json.load(f) or {}).get("entries", []) or []
        except Exception:
            registry = []

    registry = [e for e in registry if e.get("url") != report["url"]]
    registry.append({
        "variable": p.get("variable", ""),
        "label": p.get("label", ""),
        "apa": p.get("apa", p.get("label", "")),
        "year": p.get("year"),
        "doc_kind": p.get("doc_kind", ""),
        "journal_title": p.get("journal", ""),
        "doi": p.get("doi", ""),
        "url": report["url"],
        "words": report["words"],
        "catatan": p.get("catatan", ""),
        "imported_at": datetime.datetime.now().isoformat(timespec="seconds"),
    })

    if dry_run:
        report["status"] = "ok"
        report["note"] = "[dry-run] registry belum ditulis."
        return report

    with open(PENDUKUNG_REGISTRY, "w", encoding="utf-8") as f:
        json.dump({"updated_at": datetime.datetime.now().isoformat(timespec="seconds"),
                   "entries": registry}, f, indent=2, ensure_ascii=False)
    report["status"] = "ok"
    return report


# ---------------------------------------------------------------------------
# Cache
# ---------------------------------------------------------------------------

def clear_cache(dry_run=False):
    """Bersihkan cache S1 + cache ekstraksi agar blueprint diregenerasi."""
    removed = []
    for p in CACHE_FILES:
        full = os.path.normpath(p)
        if os.path.isfile(full):
            if not dry_run:
                try:
                    os.remove(full)
                except OSError as e:
                    print(f"  gagal hapus {full}: {e}")
                    continue
            removed.append(os.path.relpath(full, BACKEND_DIR))
    if os.path.isdir(OUTPUT_DIR):
        for fn in os.listdir(OUTPUT_DIR):
            if fn.startswith("s1_onto_"):
                full = os.path.join(OUTPUT_DIR, fn)
                if not dry_run:
                    try:
                        os.remove(full)
                    except OSError:
                        continue
                removed.append(os.path.relpath(full, BACKEND_DIR))
    return removed


# ---------------------------------------------------------------------------
# Laporan
# ---------------------------------------------------------------------------

def print_primers_report(reports):
    ok = [r for r in reports if r["status"] == "ok"]
    n_err = len([r for r in reports if r["status"] == "error"])
    n_skip = len([r for r in reports if r["status"] in ("skip", "needs_review")])
    print("\n=== LAPORAN IMPOR PRIMER ===")
    for r in reports:
        flag = {"ok": "OK", "error": "ERROR", "skip": "SKIP",
                "needs_review": "REVIEW"}.get(r["status"], r["status"])
        line = (f"[{flag}] {r['variable']}  (mode={r['mode']})  "
                f"{r['dims']} dim / {r['indicators']} ind")
        if r.get("literal_checked"):
            line += (f"  | audit literal: {r['literal_passed']}/"
                     f"{r['literal_checked']} lolos, {r['literal_dropped']} dibuang")
        if r.get("url"):
            line += f"\n        sumber: {r['url']} (HTTP200={r.get('url_ok')})"
        if r.get("saved"):
            line += f"\n        entri : {r['saved']}"
        if r.get("note"):
            line += f"\n        catatan: {r['note']}"
        print(line)
    print(f"\nRingkasan: {len(ok)} sukses, {n_err} error, {n_skip} skip/review "
          f"dari {len(reports)} primer.")


def print_pendukung_report(reports):
    ok = [r for r in reports if r["status"] == "ok"]
    n_err = len([r for r in reports if r["status"] == "error"])
    print("\n=== LAPORAN IMPOR PENDUKUNG (BAB II) ===")
    for r in reports:
        flag = "OK" if r["status"] == "ok" else r["status"].upper()
        line = (f"[{flag}] {r['variable']}  {r['file']}  "
                f"({r['words']} kata)  {r['url']} (HTTP200={r.get('url_ok')})")
        if r.get("note"):
            line += f"\n        catatan: {r['note']}"
        print(line)
    print(f"\nRingkasan: {len(ok)} terdaftar, {n_err} error dari {len(reports)} pendukung.")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def cmd_status(map_path):
    primers, pendukung = load_map(map_path)
    import variable_library as VL
    print(f"Peta: {len(primers)} primer, {len(pendukung)} pendukung "
          f"(berkas: {os.path.relpath(map_path, BACKEND_DIR)})")
    print("\nPrimer:")
    for p in primers:
        var = p["variable"]
        target = target_pdf_path(p, REFERENCES_DIR)
        exists = bool(target) and os.path.isfile(target)
        lib = VL.load_entry(var)
        n_ind = sum(len(d.get("indikator", [])) for d in (lib or {}).get("dimensi", []))
        src = (lib or {}).get("source", {}).get("url", "-")
        print(f"  - {var:<22} mode={p.get('mode','develop'):<8} "
              f"pdf={'ADA' if exists else 'HILANG':<5} "
              f"pustaka={'ada' if lib else '-':<4} {len((lib or {}).get('dimensi', []))}dim/"
              f"{n_ind}ind  sumber={src}")
    print("\nPendukung:")
    for p in pendukung:
        target = target_pdf_path(p, PENDUKUNG_DIR)
        exists = bool(target) and os.path.isfile(target)
        print(f"  - {p.get('variable',''):<22} {p.get('file',''):<50} "
              f"pdf={'ADA' if exists else 'HILANG'}")
    if os.path.exists(PENDUKUNG_REGISTRY):
        try:
            with open(PENDUKUNG_REGISTRY, "r", encoding="utf-8") as f:
                reg = (json.load(f) or {}).get("entries", [])
            print(f"\nRegistry pendukung: {len(reg)} entri "
                  f"({os.path.relpath(PENDUKUNG_REGISTRY, BACKEND_DIR)})")
        except Exception:
            pass


def cmd_import(args):
    primers, pendukung = load_map(args.map)
    validate_primers(primers)
    os.makedirs(PRIMER_DIR, exist_ok=True)
    os.makedirs(PENDUKUNG_DIR, exist_ok=True)

    print(f"Memuat peta {os.path.relpath(args.map, BACKEND_DIR)}: "
          f"{len(primers)} primer, {len(pendukung)} pendukung "
          f"(dry-run={args.dry_run})")

    reports = [import_primer(p, args.dry_run) for p in primers]
    print_primers_report(reports)

    pend_reports = [import_pendukung(p, args.dry_run) for p in pendukung]
    print_pendukung_report(pend_reports)

    if args.clear_cache:
        removed = clear_cache(dry_run=args.dry_run)
        if removed:
            print("\nCache dibersihkan:")
            for r in removed:
                print(f"  - {r}")
        else:
            print("\nTidak ada cache yang perlu dibersihkan.")

    bad = [r for r in reports if r["status"] == "error"]
    return 1 if bad else 0


def main():
    _utf8_stdout()
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="cmd")
    p_import = sub.add_parser("import", help="Proses reference_map.json.")
    p_import.add_argument("--map", default=os.path.join(REFERENCES_DIR, "reference_map.json"))
    p_import.add_argument("--clear-cache", action="store_true",
                          help="Hapus cache S1 & ekstraksi setelah impor.")
    p_import.add_argument("--dry-run", action="store_true",
                          help="Jangan menulis apa pun; hanya simulasi.")
    sub.add_parser("status", help="Ringkasan peta vs pustaka vs berkas.")

    ap.add_argument("--clear-cache", action="store_true",
                    help="Hanya bersihkan cache S1 & ekstraksi.")
    args = ap.parse_args()

    if args.cmd == "import":
        sys.exit(cmd_import(args))
    if args.cmd == "status":
        sys.exit(cmd_status(args.map if hasattr(args, "map") else os.path.join(REFERENCES_DIR, "reference_map.json")))
    if args.clear_cache:
        removed = clear_cache()
        for r in removed:
            print(f"hapus: {r}")
        print("Cache dibersihkan." if removed else "Tidak ada cache.")
        sys.exit(0)
    ap.print_help()
    sys.exit(1)


if __name__ == "__main__":
    main()
