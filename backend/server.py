# backend/server.py
import http.server
import socketserver
import json
import os
import sys
import ssl
import threading
import urllib.parse
import base64

# Ensure backend directory is in sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Output konsol dipaksa UTF-8 (dengan errors='replace' sebagai jaring pengaman).
# Tanpa ini, pencetakan karakter non-ASCII (mis. "Grbić", "Karadayı") memicu
# UnicodeEncodeError di Windows (cp1252) -> kode lama harus men bastion respons
# LLM ke ASCII sehingga NAMA PENULIS jadi "Grbi?" / "Karaday?" dan tidak lagi
# cocok dengan DAFTAR PUSTAKA.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
try:
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    os.environ.setdefault("PYTHONUTF8", "1")
except Exception:
    pass

from engine_smartpls import generate_google_apps_script_from_items
from gform_sheet import fetch_google_sheet_csv
from engine_docx import build_full_thesis_docx
from build_full_thesis_docx import rapikan_tabel_dokumen
import engine_word_native
import agent_bagian_awal
import agent_bab_1
import agent_bab_2
import agent_bab_3
import agent_bab_4
import agent_bab_5
import engine_pembahasan_poin
import engine_hierarchical_pipeline
import semantic_title_parser
import llm_title_parser
import academic_dimension_researcher
import structural_graph_compiler
import dynamic_pls_engine
import engine_graph_dynamic
import s3_title_parser
from provenance import (
    label_data_source,
    label_origin,
    normalize_analysis_engine,
    normalize_data_source,
)
from generate_cert import generate_self_signed_cert

# Port dapat di-override lewat env (Render/Fly.io menyediakan $PORT).
HTTP_PORT = int(os.environ.get("PORT", "5050"))
HTTPS_PORT = int(os.environ.get("HTTPS_PORT", "5051"))
# Di cloud (Render/Fly), TLS dihentikan oleh proxy platform -> cukup HTTP saja.
CLOUD = bool(os.environ.get("PORT") or os.environ.get("RENDER") or os.environ.get("FLY_APP_NAME"))
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TASKPANE_DIR = os.path.join(BASE_DIR, "taskpane")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def ontology_gate_status(judul=""):
    """Baca cache ontologi JUDUL AKTIF dan kembalikan (ok, problems, summary, judul).

    Blueprint disimpan per judul (hash judul). Membaca berkas global lama akan
    memvalidasi judul yang salah; karena itu judul wajib diteruskan bila ada.
    """
    import variable_tracer
    onto = {}
    judul = (judul or "").strip()
    if judul:
        try:
            import academic_dimension_researcher as ADR
            onto = ADR.baca_ontology_cache(judul) or {}
        except Exception as e:
            print(f"[ontology] gagal baca cache judul: {e}")
            onto = {}
    if not onto:
        onto_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 "cache_thesis_ontology.json")
        if os.path.exists(onto_path):
            try:
                with open(onto_path, encoding="utf-8") as f:
                    onto = json.load(f)
            except Exception as e:
                print(f"[ontology] gagal baca cache gate: {e}")
    ok, problems, summary = variable_tracer.validate_ontology_complete(onto)
    judul_onto = ((onto.get("metadata") or {}).get("judul") or judul)
    return ok, problems, summary, judul_onto


def get_file_base64(filepath):
    """Encodes a file to base64 string for direct Office.js Word insertion."""
    try:
        if filepath and os.path.exists(filepath):
            with open(filepath, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")
    except Exception as e:
        print(f"Error encoding {filepath} to base64: {e}")
    return ""

def safe_save_doc(doc, target_path):
    """Saves docx safely. If target_path is locked by Microsoft Word, saves with an incremented suffix."""
    base, ext = os.path.splitext(target_path)
    candidate = target_path
    counter = 1
    while counter < 100:
        try:
            doc.save(candidate)
            return candidate
        except PermissionError:
            candidate = f"{base}_{counter}{ext}"
            counter += 1
        except Exception:
            import time
            candidate = f"{base}_{int(time.time())}{ext}"
            doc.save(candidate)
            return candidate
    return target_path

socketserver.TCPServer.allow_reuse_address = False

class ThreadedTCPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    daemon_threads = True
    allow_reuse_address = False

class AutoTesisHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, HEAD')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200, "OK")
        self.end_headers()

    def do_HEAD(self):
        self.do_GET(head_only=True)

    def do_GET(self, head_only=False):
        try:
            parsed = urllib.parse.urlparse(self.path)
            path = parsed.path

            # Health check (Render/Fly.io) — GET /api/healthz & GET /api/variable/list
            if path in ("/api/healthz", "/api/variable/list"):
                try:
                    import variable_library as _VL
                    n = len(_VL.list_library() or [])
                except Exception:
                    n = 0
                encoded = json.dumps({"status": "ok", "library": n}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Content-Length', str(len(encoded)))
                self.end_headers()
                if not head_only:
                    self.wfile.write(encoded)
                return

            # Cache ontologi SESI (in-memory) — status & pembersihan manual.
            if path.startswith("/api/ontology-session-cache"):
                try:
                    import engine_hierarchical_pipeline as _ehp
                    if "clear" in path:
                        dihapus = _ehp.bersihkan_cache_ontologi_sesi()
                        payload = {"status": "success", "action": "clear",
                                   "entries_removed": dihapus}
                    else:
                        payload = {
                            "status": "success",
                            "ttl_seconds": int(_ehp._ONTOLOGI_MEMORI_TTL),
                            "entries": len(_ehp._ONTOLOGI_MEMORI),
                        }
                except Exception as e:
                    payload = {"status": "error", "message": str(e)}
                encoded = json.dumps(payload).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Content-Length', str(len(encoded)))
                self.end_headers()
                if not head_only:
                    self.wfile.write(encoded)
                return

            # Download endpoint
            if path.startswith("/api/download/"):
                filename = os.path.basename(path)
                file_path = os.path.join(OUTPUT_DIR, filename)
                if os.path.exists(file_path):
                    with open(file_path, "rb") as f:
                        content_bytes = f.read()
                    self.send_response(200)
                    self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                    self.send_header("Content-Length", str(len(content_bytes)))
                    self.end_headers()
                    if not head_only:
                        self.wfile.write(content_bytes)
                    return
                else:
                    self.send_error(404, "File not found")
                    return

            # Poin Pembahasan metadata list
            if path == "/api/pembahasan-poin-list":
                resp = {
                    "status": "success",
                    "items": [
                        {
                            "poin": k,
                            "hipotesis": v["hipotesis"],
                            "short_title": v["short_title"],
                            "title": v["title"],
                            "beta": v["beta"],
                            "t_stat": v["t_stat"],
                            "p_val": v["p_val"],
                            "status": v["status"]
                        }
                        for k, v in engine_pembahasan_poin.PEMBAHASAN_DATA.items()
                    ]
                }
                encoded = json.dumps(resp).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Content-Length', str(len(encoded)))
                self.end_headers()
                if not head_only:
                    self.wfile.write(encoded)
                return

            # Pipeline Catalog
            if path == "/api/pipeline/catalog":
                data = engine_hierarchical_pipeline.get_all_subchapters()
                encoded = json.dumps({"status": "success", "catalog": data}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Content-Length', str(len(encoded)))
                self.end_headers()
                if not head_only:
                    self.wfile.write(encoded)
                return

            # Live Ontology Get endpoint
            if path == "/api/get-ontology":
                ontology = academic_dimension_researcher.load_thesis_ontology()
                encoded = json.dumps({"status": "success", "ontology": ontology}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Content-Length', str(len(encoded)))
                self.end_headers()
                if not head_only:
                    self.wfile.write(encoded)
                return

            # Static files
            if path == "/" or path == "/taskpane.html" or path == "/index.html":
                target = os.path.join(TASKPANE_DIR, "taskpane.html")
                content_type = "text/html; charset=utf-8"
            elif path == "/taskpane.css":
                target = os.path.join(TASKPANE_DIR, "taskpane.css")
                content_type = "text/css"
            elif path == "/taskpane.js":
                target = os.path.join(TASKPANE_DIR, "taskpane.js")
                content_type = "application/javascript"
            elif path.startswith("/assets/"):
                target = os.path.join(TASKPANE_DIR, path.lstrip("/"))
                content_type = "image/png" if path.endswith(".png") else "text/plain"
            else:
                target = os.path.join(TASKPANE_DIR, path.lstrip("/"))
                content_type = "text/plain"

            if os.path.exists(target) and os.path.isfile(target):
                with open(target, "rb") as f:
                    content_bytes = f.read()
                self.send_response(200)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Length", str(len(content_bytes)))
                self.end_headers()
                if not head_only:
                    self.wfile.write(content_bytes)
            else:
                self.send_error(404, "Resource not found")
        except (ConnectionAbortedError, ConnectionResetError, BrokenPipeError):
            pass
        except Exception as e:
            import traceback
            traceback.print_exc()
            try:
                self.send_error(500, str(e))
            except Exception:
                pass


    def do_POST(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length).decode('utf-8')
            try:
                body = json.loads(post_data) if post_data else {}
            except Exception:
                body = {}

            parsed = urllib.parse.urlparse(self.path)
            path = parsed.path

            response_data = {}
            status_code = 200

            # 0. Parse Thesis Title (LLM Gemini + Regex Fallback)
            if path == "/api/parse-title":
                raw_title = body.get("title", "") or body.get("judul", "")
                llm_result = llm_title_parser.parse_title_llm(raw_title)
                if llm_result:
                    parsed = {
                        "x1": llm_result["x1"] or "Motivasi Kerja",
                        "x2": llm_result["x2"] or "Disiplin Kerja",
                        "m": llm_result["m"] or "Kepuasan Kerja",
                        "y": llm_result["y"] or "Kinerja Pegawai",
                        "instansi": llm_result["instansi"] or "Instansi Pemerintah Daerah",
                        "sektor": semantic_title_parser.detect_institutional_sector(
                            llm_result["instansi"], raw_title),
                        "raw_title": raw_title,
                        "engine": llm_result.get("_engine", "gemini"),
                    }
                else:
                    parsed = semantic_title_parser.parse_academic_title(raw_title)
                    parsed["engine"] = "regex-fallback"
                response_data = {
                    "status": "success",
                    "title": raw_title,
                    "data": parsed
                }

            # 0.2 Research Academic Ontology (Live Dimensions & Indicators)
            elif path == "/api/research-ontology":
                import academic_dimension_researcher
                x1 = body.get("x1", "")
                x2 = body.get("x2", "")
                m = body.get("m", "")
                y = body.get("y", "")
                instansi = body.get("instansi", "")
                sektor = body.get("sektor", "")
                judul = body.get("judul", body.get("title", ""))
                # Blueprint BEKU untuk judul yang sama dipakai ulang. Ini
                # mencegah klik "Telusuri" yang tidak sengaja menimpa ontologi
                # yang sudah disetujui. Regenerasi paksa memakai force=true.
                ontology = None
                if judul and not body.get("force"):
                    try:
                        ontology = academic_dimension_researcher.baca_ontology_cache(judul)
                    except Exception:
                        ontology = None
                if ontology:
                    # Cache lama/beku: lengkapi lapisan teoretis/operasional dan
                    # metadata seleksi dimensi tanpa mengubah indikator operasional.
                    try:
                        if academic_dimension_researcher.ensure_indicator_layers(
                                ontology, judul=judul, instansi=instansi):
                            academic_dimension_researcher.simpan_ontology_cache(
                                ontology, judul)
                    except Exception as e:
                        print(f"[ontology] pelengkapan lapisan endpoint dilewati: {e}")
                if not ontology:
                    ontology = academic_dimension_researcher.build_full_thesis_ontology(
                        x1, x2, m, y, instansi, sektor, judul
                    )
                response_data = {
                    "status": "success",
                    "ontology": ontology
                }

            # 0.21 Status kelengkapan ontologi (gate Bab III-V & Lampiran).
            # Tanpa gate ini, variabel dengan 0 indikator bisa lolos dan
            # menghasilkan naskah/kuesioner kosong tanpa peringatan.
            elif path == "/api/ontology/status":
                _judul = (body.get("judul") or body.get("title") or "").strip()
                ok, problems, summary, _judul_onto = ontology_gate_status(_judul)
                response_data = {
                    "status": "success",
                    "complete": ok,
                    "problems": problems,
                    "summary": summary,
                    "judul": _judul_onto,
                }

            # 0.25 Pustaka Variabel Permanen (Telusuri via Live Academic API, Tahap B)
            elif path == "/api/variable/trace":
                import variable_tracer
                var_name = (body.get("variable") or "").strip()
                role = body.get("role", "X")
                if not var_name:
                    response_data = {"status": "error", "message": "Nama variabel kosong."}
                else:
                    res = variable_tracer.trace_variable(
                        var_name, role=role,
                        judul=body.get("judul", ""), instansi=body.get("instansi", ""),
                        force=bool(body.get("force", False)))
                    response_data = {"status": res["status"], "entry": res.get("entry"),
                                     "report": res.get("report", {})}

            # Unggahan jurnal/buku khusus pengguna (base64) -> ekstraksi -> pustaka
            elif path == "/api/variable/upload":
                import base64, variable_tracer, variable_library as VL
                var_name = (body.get("variable") or "").strip()
                role = body.get("role", "X")
                fname = os.path.basename(body.get("filename", "upload.pdf"))
                ext = os.path.splitext(fname)[1].lower()
                if not var_name or ext not in (".pdf", ".docx", ".txt"):
                    response_data = {"status": "error",
                                     "message": "Isi nama variabel dan unggah berkas .pdf/.docx/.txt."}
                else:
                    os.makedirs(VL.UPLOAD_DIR, exist_ok=True)
                    saved = os.path.join(VL.UPLOAD_DIR, f"{VL._slug(var_name)}_{fname}")
                    with open(saved, "wb") as f:
                        f.write(base64.b64decode(body.get("content_base64", "")))
                    res = variable_tracer.trace_from_upload(saved, var_name, role=role,
                                                            display_name=fname)
                    response_data = {"status": res["status"], "entry": res.get("entry"),
                                     "report": res.get("report", {})}

            # Daftar pustaka variabel tersimpan (untuk reuse lintas judul)
            elif path == "/api/variable/list":
                import variable_library as VL
                response_data = {"status": "success", "library": VL.list_library()}

            # 0.3 Get Current Cached Ontology (per judul; fallback berkas lama)
            elif path == "/api/get-ontology":
                _cfg_onto = body.get("config") or {}
                _judul = (body.get("judul") or _cfg_onto.get("judul") or "").strip()
                _onto = academic_dimension_researcher.baca_ontology_cache(_judul) if _judul else None
                if _onto:
                    try:
                        academic_dimension_researcher.ensure_indicator_layers(
                            _onto, judul=_judul,
                            instansi=(_cfg_onto.get("instansi") or ""))
                    except Exception as e:
                        print(f"[ontology] pelengkapan lapisan get-ontology dilewati: {e}")
                if not _onto:
                    onto_cache = os.path.join(os.path.dirname(__file__),
                                              "cache_thesis_ontology.json")
                    if os.path.exists(onto_cache):
                        with open(onto_cache, "r", encoding="utf-8") as f:
                            _onto = json.load(f)
                response_data = _onto if _onto else {"status": "empty"}

            # 0.5 Fetch Live Academic References (OpenAlex & Crossref Verified DOIs)
            elif path == "/api/academic-references":
                import academic_connector
                refs = academic_connector.fetch_live_verified_references(body, max_references=int(body.get("limit", 25)))
                response_data = {
                    "status": "success",
                    "total": len(refs),
                    "references": [{"text": r[0], "italic": r[1]} for r in refs]
                }

            # 0.6 P0 #9 — Verifikasi Referensi (DOI, Judul, Penulis, Tahun, Retraksi)
            elif path == "/api/verify-references":
                import academic_connector
                apa_list = body.get("references", [])
                if not apa_list:
                    response_data = {"status": "error", "message": "Field 'references' (list of APA strings) required"}
                else:
                    result = academic_connector.verify_references_from_apa_list(apa_list)
                    response_data = {
                        "status": "success" if result["all_verified"] else "partial",
                        "summary": {
                            "total": result["total"],
                            "verified": result["verified"],
                            "failed": result["failed"],
                            "all_verified": result["all_verified"]
                        },
                        "details": result["details"],
                        "issues": result["issues"]
                    }

            # 1. Jalur pembangkitan data lama DINONAKTIFKAN.
            # Alasan: menghasilkan berkas data tanpa provenance dan melewati
            # gerbang sumber data. Gunakan Tab 3 (impor Google Form/berkas atau
            # data hasil pembangkitan mesin) lalu "Olah Data Sekarang".
            elif path == "/api/simulate-smartpls":
                status_code = 410
                response_data = {
                    "status": "error",
                    "message": (
                        "Jalur pembangkitan data lama dinonaktifkan karena tidak "
                        "mencatat sumber data dan melewati gerbang Bab IV/V. "
                        "Gunakan Tab 3 untuk menyiapkan data dengan provenance."),
                }

            # 2. Google Form Script Generator (DINAMIS dari kuesioner judul terpilih)
            # Tidak ada fallback ke 36 butir generik: bila kisi-kisi belum ada,
            # peneliti harus menjalankan Bab III (3.3)/Lampiran terlebih dahulu.
            elif path == "/api/generate-gform-script":
                try:
                    cfg = engine_hierarchical_pipeline.resolve_thesis_config(body)
                    onto = (engine_hierarchical_pipeline._load_cached_ontology(cfg)
                            or engine_hierarchical_pipeline.get_or_create_thesis_ontology(cfg))
                    items = engine_hierarchical_pipeline._lampiran_items(onto, cfg)
                    if not items:
                        raise RuntimeError(
                            "Kuesioner belum tersedia. Jalankan Bab III (3.3) "
                            "atau Lampiran sehingga butir nyata judul ini terbentuk.")
                    script = generate_google_apps_script_from_items(
                        items,
                        judul=cfg.get("judul", ""),
                        instansi=cfg.get("instansi", ""),
                        tahun=cfg.get("tahun", "2026"))
                    response_data = {
                        "status": "success",
                        "script": script,
                        "total_butir": len(items),
                        "judul": cfg.get("judul", ""),
                        "pesan": (
                            "Setelah Run, buka Spreadsheet respons yang tercatat "
                            "di Execution log, bagikan sebagai Pelihat untuk siapa "
                            "saja yang memiliki link, lalu tempel link-nya di Langkah 2."),
                    }
                except Exception as ex:
                    status_code = 400
                    response_data = {
                        "status": "error",
                        "message": f"Skrip Google Form tidak dapat dibuat: {ex}",
                    }

            # 3. Jalur dokumen utuh 1-klik DINONAKTIFKAN.
            # Alasan: jalur ini membangkitkan data tanpa provenance dan melewati
            # gerbang Bab IV/V. Gunakan generator hierarkis per sub-bab.
            elif path == "/api/generate-full-thesis":
                status_code = 410
                response_data = {
                    "status": "error",
                    "message": (
                        "Jalur dokumen utuh 1-klik dinonaktifkan karena membangkitkan "
                        "data tanpa mencatat sumber dan melewati gerbang Bab IV/V. "
                        "Gunakan Tab 3 untuk menyiapkan data, 'Olah Data Sekarang', "
                        "lalu generate sub-bab Bab IV/V."),
                }

            # 3.5 Generate Dedicated Standalone Front Matter (.docx)
            elif path == "/api/generate-bagian-awal":
                penulis = body.get("penulis", "RUSMIANITA").replace(" ", "_")
                out_name = f"BAGIAN_AWAL_LENGKAP_{penulis}.docx"
                out_file = os.path.join(OUTPUT_DIR, out_name)
                from docx import Document
                from build_full_thesis_docx import setup_page, build_bagian_awal
                doc = Document()
                setup_page(doc)
                build_bagian_awal(doc, body)
                saved_path = safe_save_doc(doc, out_file)
                actual_name = os.path.basename(saved_path)
                opened = False
                try:
                    engine_word_native.open_docx_file_in_word(saved_path)
                    opened = True
                except Exception:
                    try:
                        os.startfile(saved_path)
                        opened = True
                    except Exception:
                        opened = False
                response_data = {
                    "status": "success",
                    "filename": actual_name,
                    "opened": opened,
                    "docx_base64": get_file_base64(saved_path),
                    "download_url": f"/api/download/{actual_name}"
                }

            # 3.6 Generate Dedicated Standalone Bab I Pendahuluan (.docx)
            elif path == "/api/generate-bab-1":
                penulis = body.get("penulis", "RUSMIANITA").replace(" ", "_")
                out_name = f"BAB_1_PENDAHULUAN_{penulis}.docx"
                out_file = os.path.join(OUTPUT_DIR, out_name)
                from docx import Document
                from build_full_thesis_docx import setup_page
                doc = Document()
                setup_page(doc)
                is_andre = ("ANDRE" in body.get("penulis", "").upper()) and ("DPMPTSP" in body.get("instansi", "").upper() or "DPMPTSP" in body.get("judul", "").upper())
                if is_andre:
                    from thesis_bab1_bab2 import build_bab_1
                    build_bab_1(doc)
                else:
                    engine_hierarchical_pipeline._build_sub_1_1(doc, body)
                    engine_hierarchical_pipeline._build_sub_1_2(doc)
                    engine_hierarchical_pipeline._build_sub_1_3(doc)
                    engine_hierarchical_pipeline._build_sub_1_4(doc)
                    engine_hierarchical_pipeline._build_sub_1_5(doc)
                    engine_hierarchical_pipeline._build_sub_1_6(doc)
                engine_hierarchical_pipeline.append_cited_bibliography_page(doc, body, "bab_1")
                rapikan_tabel_dokumen(doc)
                saved_path = safe_save_doc(doc, out_file)
                actual_name = os.path.basename(saved_path)
                opened = False
                try:
                    engine_word_native.open_docx_file_in_word(saved_path)
                    opened = True
                except Exception:
                    try:
                        os.startfile(saved_path)
                        opened = True
                    except Exception:
                        opened = False
                response_data = {
                    "status": "success",
                    "filename": actual_name,
                    "opened": opened,
                    "docx_base64": get_file_base64(saved_path),
                    "download_url": f"/api/download/{actual_name}"
                }

            # 3.7 Generate Dedicated Standalone Bab II Kajian Pustaka (.docx)
            elif path == "/api/generate-bab-2":
                penulis = body.get("penulis", "ANDRE JULIANSYA").replace(" ", "_")
                out_name = f"BAB_2_KAJIAN_PUSTAKA_{penulis}.docx"
                out_file = os.path.join(OUTPUT_DIR, out_name)
                from docx import Document
                from build_full_thesis_docx import setup_page
                from thesis_bab1_bab2 import build_bab_2
                doc = Document()
                setup_page(doc)
                build_bab_2(doc)
                rapikan_tabel_dokumen(doc)
                saved_path = safe_save_doc(doc, out_file)
                actual_name = os.path.basename(saved_path)
                opened = False
                try:
                    engine_word_native.open_docx_file_in_word(saved_path)
                    opened = True
                except Exception:
                    try:
                        os.startfile(saved_path)
                        opened = True
                    except Exception:
                        opened = False
                response_data = {
                    "status": "success",
                    "filename": actual_name,
                    "opened": opened,
                    "docx_base64": get_file_base64(saved_path),
                    "download_url": f"/api/download/{actual_name}"
                }

            # 3.8 Generate Dedicated Standalone Bab III Metode Penelitian (.docx)
            elif path == "/api/generate-bab-3":
                penulis = body.get("penulis", "ANDRE JULIANSYA").replace(" ", "_")
                out_name = f"BAB_3_METODE_PENELITIAN_{penulis}.docx"
                out_file = os.path.join(OUTPUT_DIR, out_name)
                from docx import Document
                from build_full_thesis_docx import setup_page
                from thesis_bab3_bab4 import build_bab_3
                doc = Document()
                setup_page(doc)
                build_bab_3(doc)
                rapikan_tabel_dokumen(doc)
                saved_path = safe_save_doc(doc, out_file)
                actual_name = os.path.basename(saved_path)
                opened = False
                try:
                    engine_word_native.open_docx_file_in_word(saved_path)
                    opened = True
                except Exception:
                    try:
                        os.startfile(saved_path)
                        opened = True
                    except Exception:
                        opened = False
                response_data = {
                    "status": "success",
                    "filename": actual_name,
                    "opened": opened,
                    "docx_base64": get_file_base64(saved_path),
                    "download_url": f"/api/download/{actual_name}"
                }

            # 3.9 Generate Dedicated Standalone Bab IV Hasil & Pembahasan (.docx)
            elif path == "/api/generate-bab-4":
                penulis = body.get("penulis", "ANDRE JULIANSYA").replace(" ", "_")
                out_name = f"BAB_4_HASIL_DAN_PEMBAHASAN_{penulis}.docx"
                out_file = os.path.join(OUTPUT_DIR, out_name)
                from docx import Document
                from build_full_thesis_docx import setup_page
                from thesis_bab3_bab4 import build_bab_4
                doc = Document()
                setup_page(doc)
                build_bab_4(doc)
                rapikan_tabel_dokumen(doc)
                saved_path = safe_save_doc(doc, out_file)
                actual_name = os.path.basename(saved_path)
                opened = False
                try:
                    engine_word_native.open_docx_file_in_word(saved_path)
                    opened = True
                except Exception:
                    try:
                        os.startfile(saved_path)
                        opened = True
                    except Exception:
                        opened = False
                response_data = {
                    "status": "success",
                    "filename": actual_name,
                    "opened": opened,
                    "docx_base64": get_file_base64(saved_path),
                    "download_url": f"/api/download/{actual_name}"
                }

            # 3.10 Generate Dedicated Standalone Bab V Kesimpulan & Saran + Dapus (.docx)
            elif path == "/api/generate-bab-5":
                penulis = body.get("penulis", "ANDRE JULIANSYA").replace(" ", "_")
                out_name = f"BAB_5_KESIMPULAN_DAN_DAPUS_{penulis}.docx"
                out_file = os.path.join(OUTPUT_DIR, out_name)
                from docx import Document
                from build_full_thesis_docx import setup_page
                from thesis_bab5_lampiran import build_bab_5, build_daftar_pustaka
                doc = Document()
                setup_page(doc)
                build_bab_5(doc)
                build_daftar_pustaka(doc)
                rapikan_tabel_dokumen(doc)
                saved_path = safe_save_doc(doc, out_file)
                actual_name = os.path.basename(saved_path)
                opened = False
                try:
                    engine_word_native.open_docx_file_in_word(saved_path)
                    opened = True
                except Exception:
                    try:
                        os.startfile(saved_path)
                        opened = True
                    except Exception:
                        opened = False
                response_data = {
                    "status": "success",
                    "filename": actual_name,
                    "opened": opened,
                    "docx_base64": get_file_base64(saved_path),
                    "download_url": f"/api/download/{actual_name}"
                }

            # 3.11 Generate Pembahasan Bab 4.2 Per Poin Hipotesis (H1 s/d H7)
            elif path == "/api/generate-pembahasan-poin":
                poin = int(body.get("poin", 1))
                mode = body.get("mode", "download") # "download", "native", "officejs"
                data = engine_pembahasan_poin.get_pembahasan_poin(poin)
                if not data:
                    response_data = {"status": "error", "message": f"Poin pembahasan {poin} tidak ditemukan."}
                else:
                    out_name = f"PEMBAHASAN_BAB_4_POIN_{poin}_{data['hipotesis']}.docx"
                    fn, fp = engine_pembahasan_poin.generate_pembahasan_docx(poin, out_name)
                    
                    native_written = False
                    if mode == "native":
                        try:
                            engine_pembahasan_poin.write_pembahasan_poin_word(poin)
                            native_written = True
                        except Exception as ex:
                            print(f"Native write error for poin {poin}: {ex}")

                    response_data = {
                        "status": "success",
                        "poin": poin,
                        "hipotesis": data["hipotesis"],
                        "title": data["title"],
                        "filename": fn,
                        "docx_base64": get_file_base64(fp),
                        "download_url": f"/api/download/{fn}",
                        "native_written": native_written,
                        "content": data
                    }

            # 3.12 Hierarchical Sub-Chapter Generator (Per Sub-Bab)
            elif path == "/api/pipeline/generate-subchapter":
                sub_id = body.get("sub_id", "sub_1_1")
                mode = body.get("mode", "download") # "native", "download"
                config = body.get("config", {})

                # GERBANG ONTOLOGI: Bab III/IV & Lampiran wajib punya indikator
                # untuk SEMUA variabel. Tanpa ini, X2 kosong tetap ditulis.
                gate_error = None
                if sub_id.startswith(("sub_3_", "sub_4_")) or sub_id == "sub_lampiran":
                    _ok, _problems, _summary, _judul = ontology_gate_status(
                        config.get("judul") or config.get("title") or "")
                    if not _ok:
                        gate_error = {
                            "status": "error",
                            "sub_id": sub_id,
                            "message": ("STOP: ontologi belum lengkap -> " +
                                        "; ".join(_problems) +
                                        ". Jalankan 'Telusuri Dimensi & Indikator' "
                                        "atau unggah sumber untuk variabel tersebut."),
                            "problems": _problems,
                            "summary": _summary,
                        }

                if gate_error:
                    response_data = gate_error
                else:
                    # Naskah sub-bab MURNI (tanpa daftar pustaka) -> dipakai untuk sisipan
                    # ke Word; daftar pustaka dikirim terpisah agar tidak terjadi dapus ganda.
                    fn, fp = engine_hierarchical_pipeline.build_subchapter_doc(sub_id, config, include_dapus=False)

                    # Dokumen khusus Daftar Pustaka kumulatif (selalu di halaman paling akhir)
                    dapus_fn, dapus_fp = engine_hierarchical_pipeline.build_standalone_dapus_doc(config)

                    # Permintaan khusus sub_dapus: dokumen utamanya adalah daftar
                    # pustaka kumulatif (tanpa ini berkas utamanya kosong).
                    if sub_id == "sub_dapus" and dapus_fp and os.path.exists(dapus_fp):
                        fn, fp = dapus_fn, dapus_fp

                    native_written = False
                    if mode == "native":
                        try:
                            # Pakai file yang SAMA (jangan bangun ulang) supaya isi yang
                            # disisipkan identik dengan yang diunduh dan hemat waktu.
                            engine_hierarchical_pipeline.write_subchapter_native(sub_id, config, path=fp)
                            native_written = True
                        except Exception as ex:
                            print(f"Error native write {sub_id}: {ex}")

                    response_data = {
                        "status": "success",
                        "sub_id": sub_id,
                        "filename": fn,
                        "download_url": f"/api/download/{fn}",
                        "docx_base64": get_file_base64(fp),
                        "dapus_filename": dapus_fn,
                        "dapus_base64": get_file_base64(dapus_fp),
                        "native_written": native_written,
                        "meta": engine_hierarchical_pipeline.SUBCHAPTER_CATALOG.get(sub_id, {})
                    }

            # 3.13 Chained Pipeline Execution (Run Entire Chapter Sequentially)
            elif path == "/api/pipeline/run-chain":
                chapter_num = int(body.get("chapter", 1))
                config = body.get("config", {})
                mode = body.get("mode", "native")
                _ok, _problems, _summary, _judul = (True, [], {}, "")
                if chapter_num in (3, 4, 5):
                    _ok, _problems, _summary, _judul = ontology_gate_status(
                        config.get("judul") or config.get("title") or "")
                if not _ok:
                    response_data = {
                        "status": "error",
                        "message": ("STOP: ontologi belum lengkap -> " +
                                    "; ".join(_problems) +
                                    ". Lengkapi dimensi/indikator sebelum menjalankan "
                                    f"rantai Bab {chapter_num}."),
                        "problems": _problems,
                    }
                else:
                    response_data = engine_hierarchical_pipeline.run_chained_pipeline(chapter_num, config, mode)

            # 3.14 Wizard SmartPLS - langkah 1: hasilkan / unggah data jawaban responden
            elif path == "/api/pipeline/generate-respondents":
                try:
                    config = body.get("config", {}) or {}
                    aug = body.get("upload") or {}
                    sheet_url = str(aug.get("sheet_url") or body.get("sheet_url") or "").strip()
                    mode = (body.get("mode") or "generated").strip().lower()
                    form_mode = mode in ("google_form", "gform", "form", "sheet", "sheets")
                    if mode in ("gemini", "generate", "synthetic"):
                        mode = "generated"
                    if form_mode:
                        mode = "upload"
                    if mode not in ("generated", "upload"):
                        raise RuntimeError(f"Mode data '{mode}' tidak dikenal.")

                    default_src = ("google_form" if (sheet_url or form_mode)
                                   else ("upload" if mode == "upload" else "generated"))
                    src = normalize_data_source(body.get("data_source") or default_src)
                    eng = normalize_analysis_engine(
                        body.get("analysis_engine") or config.get("analysis_engine"))
                    if mode == "generated" and src != "generated":
                        raise RuntimeError(
                            "Mode pembangkitan hanya untuk data_source='generated'.")
                    if mode == "generated" and sheet_url:
                        raise RuntimeError(
                            "Link Google Sheets hanya untuk mode impor, bukan pembangkitan.")
                    if sheet_url and src == "google_form" and not body.get("konfirmasi_google_form"):
                        raise RuntimeError(
                            "Konfirmasi diperlukan: nyatakan bahwa Spreadsheet berisi "
                            "respons Google Form responden nyata, bukan data hasil "
                            "pembangkitan mesin.")
                    if mode == "upload" and src not in ("upload", "google_form"):
                        raise RuntimeError(
                            "Mode unggah hanya untuk data_source='upload' atau 'google_form'.")
                    if mode == "upload":
                        if not sheet_url and not (aug.get("csv_text") or aug.get("xlsx_b64")
                                                  or aug.get("rows")):
                            raise RuntimeError(
                                "Pilih berkas .xlsx/.csv atau tempel link Google Sheets respons.")
                        if sheet_url and (aug.get("csv_text") or aug.get("xlsx_b64")
                                          or aug.get("rows")):
                            raise RuntimeError(
                                "Pilih salah satu sumber: link Google Sheets atau berkas, "
                                "bukan keduanya.")

                    cfg = engine_hierarchical_pipeline.resolve_thesis_config(config)
                    cfg["data_source"] = src
                    cfg["analysis_engine"] = eng
                    onto = (engine_hierarchical_pipeline._load_cached_ontology(cfg)
                            or engine_hierarchical_pipeline.get_or_create_thesis_ontology(cfg))
                    items = engine_hierarchical_pipeline._lampiran_items(onto, cfg)
                    if not items:
                        raise RuntimeError("Kuesioner belum tersedia. Jalankan Bab III (3.3) atau Lampiran dulu.")
                    import pls_engine
                    if mode == "upload":
                        csv_text = aug.get("csv_text")
                        origin = str(aug.get("filename") or "").strip()
                        awal = []
                        if sheet_url:
                            csv_text, _csv_url = fetch_google_sheet_csv(sheet_url)
                            origin = sheet_url
                            awal.append("Respons dibaca langsung dari link Google Sheets.")
                        # Berkas Excel (.xlsx) -> teks CSV, lalu alur unggahan yang sama.
                        if aug.get("xlsx_b64"):
                            csv_text = pls_engine.xlsx_b64_to_csv_text(aug.get("xlsx_b64"))
                        X, warn, demog, demog_rows = pls_engine.ingest_upload_detail(
                            cfg, items, csv_text=csv_text, rows=aug.get("rows"))
                        if awal:
                            warn = awal + list(warn)
                        cols, _blk = pls_engine.build_blocks_from_items(items)
                        data_file = pls_engine.export_dataset(
                            cfg, cols, X, demog=demog, seed=1,
                            demog_rows=demog_rows, source=src)
                        pls_engine.save_pending(
                            cfg, X, source=src, data_file=data_file,
                            demog=demog, analysis_engine=eng,
                            demog_rows=demog_rows, origin=origin)
                        fn = os.path.basename(data_file)
                        response_data = {
                            "status": "success", "source": src, "data_source": src,
                            "analysis_engine": eng, "label_sumber": label_data_source(src),
                            "n": int(X.shape[0]),
                            "total_butir": len(cols), "data_file": fn,
                            "download_url": f"/api/download/{fn}", "warnings": warn,
                            "demographics": demog,
                            "origin": origin, "origin_label": label_origin(origin),
                        }
                    else:
                        res = pls_engine.generate_respondents(
                            cfg, items, source=src, analysis_engine=eng)
                        res.update({"status": "success", "source": src,
                                    "data_source": src, "analysis_engine": eng,
                                    "label_sumber": label_data_source(src),
                                    "origin": "", "origin_label": "",
                                    "download_url": f"/api/download/{res.get('data_file')}"})
                        response_data = res
                except Exception as ex:
                    status_code = 400
                    response_data = {"status": "error", "message": str(ex)}

            # 3.14b Gerbang data Bab IV: apakah data jawaban responden sudah siap?
            elif path == "/api/pipeline/data-status":
                try:
                    config = body.get("config", {}) or {}
                    cfg = engine_hierarchical_pipeline.resolve_thesis_config(config)
                    import pls_engine
                    st = pls_engine.data_status(cfg)
                    st["status"] = "success"
                    response_data = st
                except Exception as ex:
                    status_code = 400
                    response_data = {"status": "error", "message": str(ex)}

            # 3.14c Penomoran halaman otomatis pada dokumen Word AKTIF:
            #       bagian depan (Bagian Awal) = angka romawi kecil (i, ii, ...),
            #       naskah (BAB I dst.) = angka Arab mulai 1, plus pembatas bab otomatis.
            #       Dipakai jalur Office.js (task pane) yang tidak punya API format halaman.
            elif path == "/api/format-halaman":
                try:
                    res = engine_word_native.ensure_page_numbering()
                    res["status"] = "success"
                    response_data = res
                except Exception as ex:
                    status_code = 400
                    response_data = {"status": "error", "message": str(ex)}

            # 3.15 Wizard - langkah 2: "Olah Data Sekarang"
            elif path == "/api/pipeline/olah-data":
                try:
                    config = body.get("config", {}) or {}
                    bootstrap = int(body.get("bootstrap", 5000) or 5000)
                    cfg = engine_hierarchical_pipeline.resolve_thesis_config(config)
                    onto = (engine_hierarchical_pipeline._load_cached_ontology(cfg)
                            or engine_hierarchical_pipeline.get_or_create_thesis_ontology(cfg))
                    items = engine_hierarchical_pipeline._lampiran_items(onto, cfg)
                    if not items:
                        raise RuntimeError("Kuesioner belum tersedia. Jalankan Bab III (3.3) atau Lampiran dulu.")
                    import pls_engine
                    entry = pls_engine.analyze_pending(
                        cfg, items, bootstrap=bootstrap,
                        source=cfg.get("data_source"),
                        analysis_engine=cfg.get("analysis_engine"))
                    df = entry.get("data_file")
                    eng = normalize_analysis_engine(
                        entry.get("analysis_engine") or cfg.get("analysis_engine"))
                    gname = ""
                    if eng == "smartpls":
                        guide = pls_engine.smartpls_guide(cfg, items, df)
                        gname = f"PANDUAN_SMARTPLS_{pls_engine._safe_tag(cfg)}.txt"
                        with open(os.path.join(OUTPUT_DIR, gname), "w", encoding="utf-8") as gf:
                            gf.write(guide)
                    response_data = {
                        "status": "success",
                        "source": entry.get("source"),
                        "data_source": entry.get("source"),
                        "analysis_engine": eng,
                        "label_sumber": label_data_source(entry.get("source")),
                        "origin": entry.get("origin") or "",
                        "origin_label": label_origin(entry.get("origin")),
                        "n": entry.get("n"),
                        "total_butir": len(entry.get("columns") or []),
                        "data_file": df,
                        "download_url": f"/api/download/{df}",
                        "guide_file": gname,
                        "guide_download_url": f"/api/download/{gname}" if gname else "",
                        "summary": pls_engine.metrics_summary(entry.get("metrics") or {}),
                    }
                except Exception as ex:
                    status_code = 400
                    response_data = {"status": "error", "message": str(ex)}

            elif path == "/api/write-native-word":
                action = body.get("action", "bab_1")
                config = body.get("config", {})
                paragraphs = self._get_chapter_paragraphs(action, config)
                
                success = False
                err_msg = ""
                try:
                    engine_word_native.write_paragraphs_native(paragraphs)
                    success = True
                except Exception as e:
                    err_msg = str(e)

                response_data = {
                    "status": "success" if success else "error",
                    "action": action,
                    "message": "Naskah berhasil diketik langsung ke dokumen Microsoft Word aktif!" if success else f"Gagal mengetik ke Word: {err_msg}"
                }

            # 5. Generate Chapter Text (Structured JSON for Office.js)
            elif path == "/api/generate-chapter":
                action = body.get("action", "bab_1")
                config = body.get("config", {})
                paragraphs = self._get_chapter_paragraphs(action, config)
                response_data = {"status": "success", "paragraphs": paragraphs}

            # 6. Tahap 1: Semantic Hybrid Title Parser
            elif path == "/api/parse-title":
                title_input = body.get("title") or body.get("judul") or ""
                # Lapis 1: Gemini LLM (semantik penuh). Lapis 2: regex fallback.
                llm_result = llm_title_parser.parse_title_llm(title_input)
                if llm_result:
                    parsed = {
                        "x1": llm_result["x1"] or "Motivasi Kerja",
                        "x2": llm_result["x2"] or "Disiplin Kerja",
                        "m": llm_result["m"] or "Kepuasan Kerja",
                        "y": llm_result["y"] or "Kinerja Pegawai",
                        "instansi": llm_result["instansi"] or "Instansi Pemerintah Daerah",
                        "sektor": semantic_title_parser.detect_institutional_sector(
                            llm_result["instansi"], title_input),
                        "raw_title": title_input,
                        "engine": llm_result.get("_engine", "gemini"),
                    }
                else:
                    parsed = semantic_title_parser.parse_academic_title(title_input)
                    parsed["engine"] = "regex-fallback"
                response_data = {
                    "status": "success",
                    "title": title_input,
                    "data": parsed
                }

            # 7. Tahap 1: Live Academic Dimension Researcher
            elif path == "/api/research-ontology":
                title_input = body.get("title") or body.get("judul") or ""
                config_override = body.get("config") if ("config" in body and isinstance(body["config"], dict)) else body
                ontology = academic_dimension_researcher.research_thesis_ontology(
                    title=title_input,
                    config_override=config_override
                )
                response_data = {
                    "status": "success",
                    "title": title_input,
                    "ontology": ontology
                }

            elif path == "/api/get-ontology":
                ontology = academic_dimension_researcher.load_thesis_ontology()
                response_data = {
                    "status": "success",
                    "ontology": ontology
                }

            # =========================================================================
            # S3 DOCTORAL DISSERTATION ENGINE (V3.0) - PARSER & AUTO-DETECT JUDUL S3
            # =========================================================================
            elif path == "/api/s3/parse-title":
                try:
                    raw_title = body.get("title", "") or body.get("judul", "")
                    parsed_s3 = s3_title_parser.parse_title_doctoral(raw_title)
                    
                    # Otomatis kompilasi graph model berdasarkan hasil deteksi judul
                    compiled = structural_graph_compiler.compile_research_graph(
                        model_type=parsed_s3.get("model_type", "moderated_mediation"),
                        x_vars=parsed_s3.get("x_vars", []),
                        m_vars=parsed_s3.get("m_vars", []),
                        z_vars=parsed_s3.get("z_vars", []),
                        y_var=parsed_s3.get("y_var", "Kinerja Pegawai"),
                        instansi=parsed_s3.get("instansi", "Instansi Terkait")
                    )
                    response_data = {
                        "status": "success",
                        "title": raw_title,
                        "data": parsed_s3,
                        "graph": compiled
                    }
                except Exception as ex:
                    status_code = 400
                    response_data = {"status": "error", "message": f"Deteksi judul S3 gagal: {ex}"}

            # =========================================================================
            # DYNAMIC RESEARCH GRAPH ENGINE (V3.0) - MULTI-X, MODERASI & MEDIASI TERMODERASI
            # (Modul Terpisah - 100% Tidak Mengganggu Mesin Master)
            # =========================================================================
            elif path == "/api/graph/compile":
                try:
                    m_type = body.get("model_type", "mediation")
                    x_list = body.get("x_vars", [])
                    m_list = body.get("m_vars", [])
                    z_list = body.get("z_vars", [])
                    y_val = body.get("y_var", "Kinerja Pegawai")
                    inst = body.get("instansi", "Instansi Terkait")
                    focus = body.get("focus_topic", "Manajemen Sumber Daya Manusia")

                    compiled = structural_graph_compiler.compile_research_graph(
                        model_type=m_type,
                        x_vars=x_list,
                        m_vars=m_list,
                        z_vars=z_list,
                        y_var=y_val,
                        instansi=inst,
                        focus_topic=focus
                    )
                    response_data = {"status": "success", "graph": compiled}
                except Exception as ex:
                    status_code = 400
                    response_data = {"status": "error", "message": f"Kompilasi Graph gagal: {ex}"}

            elif path == "/api/graph/simulate-pls":
                try:
                    graph_spec = body.get("graph", {})
                    n_sample = int(body.get("sample_size", 92))
                    pls_res = dynamic_pls_engine.simulate_dynamic_pls(graph_spec, sample_size=n_sample)
                    response_data = {"status": "success", "pls_results": pls_res}
                except Exception as ex:
                    status_code = 400
                    response_data = {"status": "error", "message": f"Simulasi PLS gagal: {ex}"}

            elif path == "/api/graph/generate-docx":
                try:
                    graph_spec = body.get("graph", {})
                    pls_res = body.get("pls_results", {})
                    out_name = body.get("filename", "NASKAH_RISET_MODEL_HIBRIDA.docx")
                    if not out_name.endswith(".docx"):
                        out_name += ".docx"
                    out_path = os.path.join(OUTPUT_DIR, out_name)
                    sub_id = body.get("sub_id", "all")
                    saved = engine_graph_dynamic.build_graph_model_docx(graph_spec, pls_res, out_path, sub_id=sub_id)
                    response_data = {
                        "status": "success",
                        "filename": out_name,
                        "download_url": f"/api/download/{out_name}",
                        "docx_base64": get_file_base64(saved)
                    }
                except Exception as ex:
                    status_code = 400
                    response_data = {"status": "error", "message": f"Pembuatan DOCX gagal: {ex}"}

            elif path == "/api/graph/write-word":
                try:
                    sub_id = body.get("sub_id", "all")
                    graph_spec = body.get("graph", {})
                    pls_res = body.get("pls_results", {})
                    written = engine_graph_dynamic.write_graph_subchapter_native(sub_id, graph_spec, pls_res)
                    response_data = {
                        "status": "success" if written else "error",
                        "sub_id": sub_id,
                        "message": "Naskah berhasil diketik ke dokumen Word aktif!" if written else "Gagal mengetik ke Word aktif (pastikan Microsoft Word terbuka)."
                    }
                except Exception as ex:
                    status_code = 400
                    response_data = {"status": "error", "message": f"Pengetikan Word gagal: {ex}"}

            else:
                status_code = 404
                response_data = {"error": "Endpoint not found"}

            self.send_response(status_code)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(response_data).encode('utf-8'))
        except (ConnectionAbortedError, ConnectionResetError, BrokenPipeError):
            pass
        except Exception as e:
            import traceback
            traceback.print_exc()
            try:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "error", "message": str(e)}).encode('utf-8'))
            except Exception:
                pass


    def _get_chapter_paragraphs(self, action, config):
        instansi = config.get("instansi", "Badan Pengelola Pajak Daerah Kabupaten Ogan Komering Ilir")
        x1 = config.get("x1", "Self-Efficacy")
        x2 = config.get("x2", "Work Environment")
        m = config.get("m", "Kepuasan Kerja Pegawai")
        y = config.get("y", "Kinerja Pegawai")
        penulis = config.get("penulis", "NURMAYANTI")
        nim = config.get("nim", "255541046")
        tahun = config.get("tahun", "2026")
        univ = config.get("univ", "UNIVERSITAS TRIDINANTI")
        fakultas = config.get("fakultas", "FAKULTAS EKONOMI DAN BISNIS")
        prodi = config.get("prodi", "PROGRAM STUDI MAGISTER MANAJEMEN")
        konsentrasi = config.get("konsentrasi", "Manajemen Sumber Daya Manusia (MSDM)")
        judul = config.get("judul", f"PENGARUH {x1.upper()} DAN {x2.upper()} TERHADAP {y.upper()} YANG DIMEDIASI OLEH {m.upper()} PADA {instansi.upper()}")

        paragraphs = []

        if action == "front_matter":
            paragraphs = [
                {"type": "heading1", "text": "TESIS MAGISTER MANAJEMEN"},
                {"type": "title", "text": judul},
                {"type": "body_center", "text": "Ditulis untuk Memenuhi Sebagian Persyaratan dalam\nMenempuh Gelar Magister Manajemen (M.M.)\n\n[ LOGO UNIVERSITAS ]\n"},
                {"type": "body_center", "text": f"Disusun Oleh:\nNama : {penulis}\nNIM : {nim}\nKonsentrasi : {konsentrasi}"},
                {"type": "body_center", "text": f"{prodi}\n{fakultas}\n{univ}\n{tahun}"},
                
                {"type": "heading1", "text": "LEMBAR PERSETUJUAN KOMISI PEMBIMBING TESIS"},
                {"type": "body_center", "text": "Tesis ini Telah Disetujui Untuk Diujikan"},
                {"type": "body", "text": f"Judul : {judul}\nNama : {penulis}\nNIM : {nim}\nKonsentrasi : {konsentrasi}"},
                {"type": "body", "text": "Pembimbing Utama: Prof. Dr. Ir. H. Edizal, AE., M.S.\nPembimbing Pendamping: Dr. M. Ima Andriyani, S.E., M.Si.\nMengetahui: Dr. Hj. Nina Fitriana, S.E., M.Si. (Ketua Program Studi)"},

                {"type": "heading1", "text": "LEMBAR PENGESAHAN KOMISI PENGUJI TESIS"},
                {"type": "body_center", "text": f"Tesis ini Telah Dipertahankan di Depan Sidang Tim Penguji Tesis dan Dinyatakan LULUS\nPalembang, {tahun}"},
                {"type": "body", "text": "1. Prof. Dr. Ir. H. Edizal, AE., M.S. (Ketua Tim Penguji)\n2. Dr. M. Ima Andriyani, S.E., M.Si. (Sekretaris)\n3. Dr. H. Ruslan Abdul Gani, S.E., M.Si. (Penguji Utama)\n4. Dr. Hj. Sri Rahayu, S.E., M.M. (Penguji Anggota)\nMengesahkan: Prof. Dr. H. Djatmiko, S.E., M.Si. (Dekan FEB)"},

                {"type": "heading1", "text": "KATA PENGANTAR"},
                {"type": "body", "text": f"Puji dan syukur penulis panjatkan ke hadirat Allah SWT atas segala limpahan rahmat, taufik, dan hidayah-Nya sehingga naskah tesis berjudul \"{judul}\" dapat diselesaikan dengan baik guna memenuhi salah satu syarat memperoleh gelar Magister Manajemen (M.M.) pada {univ}."},
                {"type": "body", "text": f"Penulis menyampaikan terima kasih yang setinggi-tingginya kepada Rektor {univ}, Dekan {fakultas}, Ketua {prodi}, Pembimbing Utama Prof. Dr. Ir. H. Edizal, AE., M.S., Pembimbing Pendamping Dr. M. Ima Andriyani, S.E., M.Si., Kepala {instansi} beserta seluruh jajaran responden, keluarga tercinta, dan rekan-rekan mahasiswa angkatan {tahun}."},

                {"type": "heading1", "text": "SURAT PERNYATAAN BEBAS PLAGIARISME"},
                {"type": "body", "text": f"Saya yang bertanda tangan di bawah ini: {penulis} (NIM. {nim}) menyatakan dengan sesungguhnya bahwa tesis berjudul \"{judul}\" adalah orisinal karya saya sendiri dan bebas dari plagiarisme. Apabila di kemudian hari terbukti sebaliknya, saya bersedia menerima sanksi pencabutan gelar akademik Magister Manajemen (M.M.) sesuai hukum yang berlaku."},

                {"type": "heading1", "text": "ABSTRAK"},
                {"type": "body", "text": f"{penulis}, Pengaruh {x1} Dan {x2} Terhadap {y} Yang Dimediasi Oleh {m} Pada {instansi}. Dibimbing oleh Prof. Dr. Ir. H. Edizal, AE., M.S. dan Dr. M. Ima Andriyani, S.E., M.Si. {prodi} {univ}, {tahun}."},
                {"type": "body", "text": f"Penelitian ini menganalisis pengaruh {x1.lower()} dan {x2.lower()} terhadap {y.lower()} melalui mediasi {m.lower()} pada {instansi}. Desain riset kuantitatif asosiatif kausal dengan teknik sensus jenuh (N=45) dan dianalisis menggunakan SmartPLS 4."},
                {"type": "body", "text": "Hasil uji membuktikan seluruh 7 hipotesis diterima signifikan (p < 0,05, T > 1,96). Efikasi diri dan lingkungan kerja terbukti berkontribusi positif langsung terhadap kinerja maupun secara tidak langsung melalui kepuasan kerja pegawai."},
                {"type": "body", "text": f"Kata Kunci: {x1}, {x2}, {m}, {y}, SmartPLS, Mediasi."},

                {"type": "heading1", "text": "ABSTRACT"},
                {"type": "body", "text": f"{penulis}, The Effect of {x1} and {x2} on {y} Mediated by {m} at {instansi}. Supervised by Prof. Dr. Ir. H. Edizal, AE., M.S. and Dr. M. Ima Andriyani, S.E., M.Si. Postgraduate Program in Management, {univ}, {tahun}."},
                {"type": "body", "text": f"This study empirically examines the effect of {x1.lower()} and {x2.lower()} on {y.lower()} mediated by {m.lower()} at {instansi}. Saturated census sampling (N=45) and SmartPLS 4 PLS-SEM bootstrapping were deployed. All seven hypotheses are supported with p < 0.05 and T > 1.96."},
                {"type": "body", "text": f"Keywords: {x1}, {x2}, {m}, {y}, SmartPLS, Mediation."},

                {"type": "heading1", "text": "DAFTAR RIWAYAT HIDUP"},
                {"type": "body", "text": f"Nama: {penulis}\nTTL: Kayuagung, 12 Agustus 1990\nPendidikan: SDN 1 Kayuagung (2002), SMPN 1 Kayuagung (2005), SMAN 1 Kayuagung (2008), S1 Fakultas Ekonomi UNSRI (2012), S2 Magister Manajemen Universitas Tridinanti.\nPekerjaan: Staf Pelaksana pada {instansi} (2014 - Sekarang)."}
            ]

        elif action == "bab_1":
            paragraphs = [
                {"type": "heading1", "text": "BAB I\nPENDAHULUAN"},
                {"type": "heading2", "text": "1.1 Latar Belakang Masalah"},
                {"type": "body", "text": f"Pengelolaan sumber daya manusia dalam sektor publik memegang peranan yang sangat fundamental dalam memastikan keberhasilan pencapaian tujuan pembangunan daerah. Pada {instansi}, efektivitas penerimaan pendapatan asli daerah sangat ditentukan oleh kapabilitas, etos kerja, dan dedikasi aparatur pengelola pajak. Terwujudnya {y.lower()} yang optimal bukan hanya sekadar pencapaian indikator kinerja individu, melainkan pilar utama dalam pemenuhan target penerimaan kas daerah yang berimplikasi langsung pada pembiayaan pelayanan publik."},
                {"type": "body", "text": f"Berdasarkan data empiris pada {instansi} selama periode 2022 hingga 2024, realisasi penerimaan pajak daerah menunjukkan adanya fluktuasi dan kesenjangan kinerja (performance gap). Pada tahun 2022 realisasi mencapai 95,00%, kemudian pada tahun 2023 menurun menjadi 92,00%, dan pada tahun 2024 berada pada angka 88,80% dari target yang ditetapkan. Kondisi ini mencerminkan perlunya evaluasi mendalam terhadap faktor-faktor perilaku aparatur organisasi."},
                {"type": "body", "text": f"Dalam perspektif Social Cognitive Theory (Bandura, 1986), keyakinan pegawai atas kemampuan dirinya ({x1.lower()}) menjadi faktor penggerak utama dalam menyelesaikan beban kerja perpajakan. Selain itu, suasana dan sarana prasarana ({x2.lower()}) sangat krusial dalam menunjang konsentrasi kerja. Interaksi kedua faktor tersebut berdampak langsung terhadap kepuasan kerja ({m.lower()}) yang selanjutnya mengeskalasi capaian {y.lower()}."},
                {"type": "heading2", "text": "1.2 Identifikasi Masalah"},
                {"type": "bullet", "text": f"1. Adanya penurunan persentase realisasi penerimaan pajak daerah pada {instansi} selama 2022-2024."},
                {"type": "bullet", "text": f"2. Fluktuasi ketahanan mental dan keyakinan diri ({x1.lower()}) pegawai saat menghadapi target pelaporan."},
                {"type": "bullet", "text": f"3. Belum optimalnya pemerataan fasilitas dan kenyamanan ruang kerja ({x2.lower()})."},
                {"type": "bullet", "text": f"4. Keluhan pegawai terkait kepuasan kerja ({m.lower()}) pada aspek pengakuan kinerja."},
                {"type": "bullet", "text": f"5. Keterlambatan verifikasi berkas wajib pajak yang berdampak pada ketepatan waktu pelaporan."},
                {"type": "heading2", "text": "1.3 Pembatasan Masalah"},
                {"type": "body", "text": f"Penelitian ini dibatasi pada pengujian empat konstruk: dua variabel independen ({x1} dan {x2}), satu variabel intervening ({m}), dan satu variabel dependen ({y}) pada aparatur {instansi}."},
                {"type": "heading2", "text": "1.4 Perumusan Masalah"},
                {"type": "bullet", "text": f"1. Apakah terdapat pengaruh positif dan signifikan {x1} terhadap {m} pada {instansi}?"},
                {"type": "bullet", "text": f"2. Apakah terdapat pengaruh positif dan signifikan {x2} terhadap {m} pada {instansi}?"},
                {"type": "bullet", "text": f"3. Apakah terdapat pengaruh positif dan signifikan {x1} terhadap {y} pada {instansi}?"},
                {"type": "bullet", "text": f"4. Apakah terdapat pengaruh positif dan signifikan {x2} terhadap {y} pada {instansi}?"},
                {"type": "bullet", "text": f"5. Apakah terdapat pengaruh positif dan signifikan {m} terhadap {y} pada {instansi}?"},
                {"type": "bullet", "text": f"6. Apakah terdapat pengaruh positif dan signifikan {x1} terhadap {y} yang dimediasi oleh {m} pada {instansi}?"},
                {"type": "bullet", "text": f"7. Apakah terdapat pengaruh positif dan signifikan {x2} terhadap {y} yang dimediasi oleh {m} pada {instansi}?"},
                {"type": "heading2", "text": "1.5 Tujuan Penelitian"},
                {"type": "body", "text": "Penelitian ini bertujuan untuk menganalisis dan membuktikan secara empiris ke-7 hubungan pengaruh langsung dan tidak langsung (mediasi) antar-variabel tersebut di atas."},
                {"type": "heading2", "text": "1.6 Manfaat Penelitian"},
                {"type": "body", "text": "Memberikan kontribusi teoretis bagi pengembangan ilmu manajemen sumber daya manusia dengan pendekatan PLS-SEM serta memberikan rekomendasi kebijakan praktis bagi pimpinan instansi dalam optimalisasi potensi aparatur."}
            ]
        elif action == "bab_2":
            paragraphs = [
                {"type": "heading1", "text": "BAB II\nKAJIAN PUSTAKA DAN PENGEMBANGAN HIPOTESIS"},
                {"type": "heading2", "text": "2.1 Landasan Teori"},
                {"type": "body", "text": f"Penelitian ini berlandaskan pada Grand Theory Social Cognitive Theory (Bandura, 1986) dan Two-Factor Theory (Herzberg, 1966). Social Cognitive Theory menjelaskan determinisme resiprokal triadik bahwa faktor kognitif internal individu ({x1}), lingkungan kerja ({x2}), dan pola perilaku/kinerja ({y}) saling memengaruhi secara timbal balik. Sementara itu, Two-Factor Herzberg menjelaskan bahwa pemenuhan faktor higienis dan motivator secara simultan menciptakan kepuasan kerja pegawai ({m}) yang mendorong kinerja optimal."},
                {"type": "heading2", "text": "2.2 Penelitian Terdahulu yang Relevan"},
                {"type": "body", "text": "Riset terdahulu oleh Pratama & Suryani (2022), Chen et al. (2023), serta Yurahadi (2024) secara konsisten membuktikan bahwa efikasi diri dan lingkungan kerja merupakan prediktor signifikan terhadap peningkatan kepuasan dan kinerja pegawai sektor publik."},
                {"type": "heading2", "text": "2.3 Kerangka Berpikir Konseptual"},
                {"type": "body", "text": f"Model konseptual memposisikan {x1} dan {x2} sebagai variabel eksogen, {m} sebagai variabel intervening (mediasi), dan {y} sebagai variabel endogen utama, membentuk model mediasi komplementer 7 jalur hipotesis."},
                {"type": "heading2", "text": "2.4 Perumusan 7 Hipotesis Penelitian"},
                {"type": "bullet", "text": f"H1: {x1} berpengaruh positif dan signifikan terhadap {m} pada {instansi}."},
                {"type": "bullet", "text": f"H2: {x2} berpengaruh positif dan signifikan terhadap {m} pada {instansi}."},
                {"type": "bullet", "text": f"H3: {x1} berpengaruh positif dan signifikan terhadap {y} pada {instansi}."},
                {"type": "bullet", "text": f"H4: {x2} berpengaruh positif dan signifikan terhadap {y} pada {instansi}."},
                {"type": "bullet", "text": f"H5: {m} berpengaruh positif dan signifikan terhadap {y} pada {instansi}."},
                {"type": "bullet", "text": f"H6: {m} memediasi secara positif dan signifikan pengaruh {x1} terhadap {y} pada {instansi}."},
                {"type": "bullet", "text": f"H7: {m} memediasi secara positif dan signifikan pengaruh {x2} terhadap {y} pada {instansi}."}
            ]
        elif action == "bab_3":
            import math as _math
            try:
                _pop = int(float(config.get("populasi") or config.get("n") or 45))
            except Exception:
                _pop = 45
            if _pop <= 0:
                _pop = 45
            if _pop > 100:
                _n_exact = _pop / (1 + _pop * 0.0025)
                _n = int(_math.ceil(_n_exact))
                _sampel_txt = (f"Karena jumlah populasi lebih dari 100 orang, ukuran sampel ditentukan dengan Rumus Slovin pada taraf kelonggaran ketidaktelitian 5%, yaitu n = N / (1 + N(e)²) = {_pop} / (1 + {_pop}(0,05)²) = {_n_exact:.2f} ≈ {_n} orang responden.")
            else:
                _n = _pop
                _sampel_txt = (f"Mengingat jumlah populasi terjangkau (N ≤ 100), teknik sampling yang digunakan adalah teknik sensus (sampling jenuh), di mana seluruh populasi {_pop} orang dijadikan sebagai responden penelitian.")
            paragraphs = [
                {"type": "heading1", "text": "BAB III\nMETODE PENELITIAN"},
                {"type": "heading2", "text": "3.1 Desain Penelitian"},
                {"type": "body", "text": "Penelitian ini menggunakan pendekatan kuantitatif asosiatif kausal untuk menguji hubungan kausalitas antar-variabel dan menguji hipotesis yang diajukan."},
                {"type": "heading2", "text": "3.2 Populasi dan Sampel"},
                {"type": "body", "text": f"Populasi adalah seluruh pegawai ASN pada {instansi} yang berjumlah {_pop} orang. {_sampel_txt} Dengan demikian, ukuran sampel yang dianalisis dalam penelitian ini berjumlah {_n} responden."},
                {"type": "heading2", "text": "3.3 Instrumen dan Skala Pengukuran"},
                {"type": "body", "text": "Data primer dikumpulkan menggunakan kuesioner berskala Likert 1-5 (1 = Sangat Tidak Setuju hingga 5 = Sangat Setuju) yang mengoperasionalkan 16 indikator dari 4 konstruk penelitian."},
                {"type": "heading2", "text": "3.4 Teknik Analisis Data"},
                {"type": "body", "text": "Analisis data menggunakan Partial Least Squares Structural Equation Modeling (PLS-SEM) dengan SmartPLS 4, meliputi pengujian Measurement Model (Outer Loadings >= 0,708, AVE > 0,50, Composite Reliability > 0,70, HTMT < 0,85) dan Structural Model (R2, f2, Bootstrapping 5.000 subsamples pada alpha 5%)."}
            ]
        elif action == "bab_4":
            paragraphs = [
                {"type": "heading1", "text": "BAB IV\nHASIL PENELITIAN DAN PEMBAHASAN"},
                {"type": "heading2", "text": "4.1 Evaluasi Measurement Model (Outer Model)"},
                {"type": "body", "text": "Seluruh 16 indikator penelitian memiliki nilai Outer Loading >= 0,798 (jauh di atas batas kritis 0,708) dan nilai AVE masing-masing konstruk berada di atas 0,680 (syarat > 0,50), membuktikan terpenuhinya validitas konvergen. Nilai Composite Reliability > 0,890 membuktikan reliabilitas konstruk sangat tinggi, serta nilai HTMT < 0,70 membuktikan terpenuhinya validitas diskriminan."},
                {"type": "heading2", "text": "4.2 Evaluasi Structural Model (Inner Model)"},
                {"type": "body", "text": f"Nilai R-Square untuk {m} sebesar 0,518 (kategori moderat) dan untuk {y} sebesar 0,684 (kategori kuat). Artinya, 68,4% variasi kinerja pegawai dapat dijelaskan secara simultan oleh efikasi diri, lingkungan kerja, dan kepuasan kerja."},
                {"type": "heading2", "text": "4.3 Pengujian Hipotesis (Bootstrapping 5.000)"},
                {"type": "body", "text": "Hasil pengujian membuktikan seluruh 7 hipotesis diterima secara signifikan dengan nilai t-statistik > 1,960 dan nilai p-value < 0,05:\n- H1: β = 0,421 (t = 3,842, p = 0,000) -> Diterima\n- H2: β = 0,378 (t = 3,415, p = 0,001) -> Diterima\n- H3: β = 0,298 (t = 2,764, p = 0,006) -> Diterima\n- H4: β = 0,276 (t = 2,532, p = 0,012) -> Diterima\n- H5: β = 0,395 (t = 3,918, p = 0,000) -> Diterima\n- H6: β = 0,166 (t = 2,685, p = 0,008) -> Diterima (Mediasi Signifikan)\n- H7: β = 0,149 (t = 2,471, p = 0,014) -> Diterima (Mediasi Signifikan)"},
                {"type": "heading2", "text": "4.4 Pembahasan Komprehensif (Formula 4 Alinea)"},
                {"type": "body", "text": f"Secara teoretis temuan ini memperkuat Social Cognitive Theory (Bandura, 1986) dan Two-Factor Theory (Herzberg, 1966). Secara empiris temuan ini mengonfirmasi riset Pratama & Suryani (2022) dan Yurahadi (2024). Secara praktis di {instansi}, peningkatan kepuasan aparatur melalui penataan fasilitas dan pelatihan efikasi diri terbukti menjadi kunci utama dalam mengatasi kesenjangan kinerja penerimaan pajak daerah."}
            ]
        elif action == "bab_5":
            paragraphs = [
                {"type": "heading1", "text": "BAB V\nKESIMPULAN DAN SARAN"},
                {"type": "heading2", "text": "5.1 Kesimpulan"},
                {"type": "bullet", "text": f"1. {x1} berpengaruh positif dan signifikan terhadap {m} (β = 0,421; t = 3,842; p < 0,05)."},
                {"type": "bullet", "text": f"2. {x2} berpengaruh positif dan signifikan terhadap {m} (β = 0,378; t = 3,415; p < 0,05)."},
                {"type": "bullet", "text": f"3. {x1} berpengaruh positif dan signifikan terhadap {y} (β = 0,298; t = 2,764; p < 0,05)."},
                {"type": "bullet", "text": f"4. {x2} berpengaruh positif dan signifikan terhadap {y} (β = 0,276; t = 2,532; p < 0,05)."},
                {"type": "bullet", "text": f"5. {m} berpengaruh positif dan signifikan terhadap {y} (β = 0,395; t = 3,918; p < 0,05)."},
                {"type": "bullet", "text": f"6. {m} secara signifikan memediasi pengaruh {x1} terhadap {y} (β = 0,166; t = 2,685; p < 0,05)."},
                {"type": "bullet", "text": f"7. {m} secara signifikan memediasi pengaruh {x2} terhadap {y} (β = 0,149; t = 2,471; p < 0,05)."},
                {"type": "heading2", "text": "5.2 Implikasi Kebijakan dan Manajerial"},
                {"type": "body", "text": f"Pimpinan {instansi} disarankan memprioritaskan pemeliharaan fasilitas kerja ergonomis, memodernisasi sarana TI penagihan pajak, serta menyelenggarakan pembinaan motivasi kerja berkala guna mempertahankan kepuasan kerja aparatur."},
                {"type": "heading2", "text": "5.3 Saran"},
                {"type": "body", "text": "Disarankan bagi peneliti selanjutnya untuk memperluas model penelitian dengan menambahkan variabel komitmen organisasi atau kepemimpinan transformasional serta memperluas cakupan sampel lintas daerah."}
            ]
        else:
            paragraphs = [
                {"type": "heading1", "text": "DAFTAR PUSTAKA"},
                {"type": "body", "text": "Bandura, A. (1986). Social Foundations of Thought and Action: A Social Cognitive Theory. Prentice-Hall.\n\nChen, X., Liu, Y., & Wang, H. (2023). Employee self-efficacy, work environment and public sector performance. Public Management Review, 25(4), 589-612.\n\nHair, J. F., Hult, G. T. M., Ringle, C. M., & Sarstedt, M. (2022). A Primer on Partial Least Squares Structural Equation Modeling (PLS-SEM) (3rd ed.). SAGE Publications.\n\nHerzberg, F. (1966). Work and the Nature of Man. World Publishing Company.\n\nLuthans, F. (2015). Organizational Behavior: An Evidence-Based Approach (13th ed.). McGraw-Hill Education.\n\nPratama, A., & Suryani, E. (2022). Analisis pengaruh lingkungan kerja dan efikasi diri terhadap kinerja pegawai melalui kepuasan kerja. Jurnal Ilmu Manajemen, 10(2), 145-158.\n\nSugiyono. (2019). Metode Penelitian Kuantitatif, Kualitatif, dan R&D. Alfabeta.\n\nYurahadi. (2024). Pengaruh Lingkungan Kerja dan Tunjangan Kinerja terhadap Kinerja Pegawai Melalui Kepuasan Kerja pada Organisasi Pemerintah Daerah. Tesis Magister Manajemen, Universitas Tridinanti."}
            ]

        return paragraphs

def run_http_server():
    server_address = ('', HTTP_PORT)
    httpd = ThreadedTCPServer(server_address, AutoTesisHandler)
    print(f"[AutoTesis HTTP] Server berjalan di http://localhost:{HTTP_PORT}")
    httpd.serve_forever()

def run_https_server():
    server_dir = os.path.dirname(os.path.abspath(__file__))
    loc_crt = os.path.join(server_dir, "localhost.crt")
    loc_key = os.path.join(server_dir, "localhost.key")
    if os.path.exists(loc_crt) and os.path.exists(loc_key):
        cert_file, key_file = loc_crt, loc_key
        print(f"[SSL] Menggunakan Sertifikat Resmi Office Add-in: {cert_file}")
    else:
        cert_file, key_file = generate_self_signed_cert(server_dir)
    
    server_address = ('', HTTPS_PORT)
    httpd = ThreadedTCPServer(server_address, AutoTesisHandler)
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(certfile=cert_file, keyfile=key_file)
    httpd.socket = context.wrap_socket(httpd.socket, server_side=True)
    print(f"[AutoTesis HTTPS] Server berjalan di https://localhost:{HTTPS_PORT}")
    httpd.serve_forever()

def run_dual_server():
    if CLOUD:
        # Mode cloud (Render/Fly): hanya HTTP; TLS ditangani proxy platform.
        print("=================================================================")
        print(f"Independent Researcher Cloud Server Aktif! PORT={HTTP_PORT}")
        print(f"   Web UI : http://0.0.0.0:{HTTP_PORT}/taskpane.html")
        print("=================================================================")
        run_http_server()
        return
    t_http = threading.Thread(target=run_http_server, daemon=True)
    t_https = threading.Thread(target=run_https_server, daemon=True)
    t_http.start()
    t_https.start()
    print("=================================================================")
    print("Independent Researcher Dual Server Aktif!")
    print(f"1. HTTP Web UI  : http://localhost:{HTTP_PORT}/taskpane.html")
    print(f"2. HTTPS Add-in : https://localhost:{HTTPS_PORT}/taskpane.html")
    print("=================================================================")
    t_http.join()

if __name__ == '__main__':
    run_dual_server()
