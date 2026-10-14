# RESUME & CATATAN STATUS PROYEK AUTOTESIS AI
**Terakhir Diperbarui:** 15 Oktober 2026, 00:20 WIB  
**Status:** ✅ TELUSUR MULTI-SUMBER (OPENALEX + CROSSREF + SEMANTIC SCHOLAR) & MULTI-KEY (15 KEY GEMINI) AKTIF — 11 ENTRI PUSTAKA SEMUA BER-DOI TERVERIFIKASI; RANTAI LLM DIBERSIHKAN DARI 2 MODEL YANG SUDAH DIHAPUS GOOGLE DAN KINI MENANG DI 0,9 DETIK; FALLBACK NVIDIA NEMOTRON 3 ULTRA TETAP UTUH

---

## 1. Peningkatan Utama yang Telah Selesai Dieksekusi

### A. Solusi Masalah WebView2 & Pengetikan Word (100% Tuntas):
1. **Engine Otomasi Native Word COM (`backend/engine_word_native.py`):**
   - Menggunakan API COM resmi Microsoft Word (`win32com.client`).
   - Dapat mengetikkan naskah terformat per bab langsung ke lembar kerja Word yang aktif, mengatur margin A4 (4-4-3-3 cm), Times New Roman 12 pt, 1.5 spasi, bold heading, dan tabel terbuka akademik tanpa bergantung pada batasan keamanan browser WebView2!
2. **Dual Server HTTP (5050) & HTTPS (5051) (`backend/server.py`):**
   - Menghasilkan sertifikat SSL mandiri otomatis untuk localhost (`backend/generate_cert.py`).
   - Port `5050`: HTTP untuk Web Browser (Chrome/Edge) dan API eksternal.
   - Port `5051`: HTTPS resmi untuk Add-in Microsoft Word Taskpane.
   - Mengatasi isu mixed content / HTTP block pada Office Desktop.

### B. Penyempurnaan Engine Dokumen Tesis 5 Bab Penuh (`backend/engine_docx.py`):
Naskah Tesis Lengkap (`output/TESIS_LENGKAP_NURMAYANTI.docx`, ~60 KB) kini merangkum secara utuh standar Tesis Yurahadi:
1. **Bagian Awal (Front Matter):** Cover Luar, Cover Dalam, Lembar Persetujuan Pembimbing, Lembar Pengesahan Penguji, Surat Bebas Plagiasi bermaterai, Kata Pengantar, Abstrak Indonesia, Abstract Inggris, Daftar Riwayat Hidup, dan Daftar Isi.
2. **BAB I: Pendahuluan:** Latar Belakang dengan **Tabel 1.1 Kesenjangan Kinerja Empiris/IKU 2022-2024**, Identifikasi Masalah, Batasan Masalah, 7 Perumusan Masalah, 7 Tujuan Penelitian, dan Manfaat Penelitian.
3. **BAB II: Kajian Pustaka & Hipotesis:** Grand Theory (Social Cognitive Bandura & Two-Factor Herzberg), Kajian Pustaka Variabel X1, X2, M, Y, **Tabel 2.1 Matriks Komparasi 10 Jurnal Relevan**, Kerangka Konseptual Mediasi, dan **7 Rumusan Hipotesis**.
4. **BAB III: Metode Penelitian:** Desain Kausal, Sensus Sampling ($N=45$), **Tabel 3.2 Matriks Kisi-Kisi Operasional Variabel (36 Indikator Likert 1-5, 3 Dimensi x 3 Indikator)**, dan Evaluasi PLS-SEM SmartPLS 4.
5. **BAB IV: Hasil & Pembahasan:** Gambaran Objek, Karakteristik Demografi Responden, **Tabel Outer Loadings 36 Indikator Running 1 & 2**, **Tabel Validitas Konvergen & Reliabilitas (AVE, CR, Cronbach)**, **Tabel HTMT**, **Tabel R-Square & f-Square**, **Tabel Pengujian 7 Hipotesis Bootstrapping 5.000**, serta **Pembahasan Formula Baku 4 Alinea untuk masing-masing ke-7 Hipotesis**.
6. **BAB V: Kesimpulan & Saran:** Kesimpulan menjawab 7 hipotesis, Implikasi Teoretis, Implikasi Manajerial Praktis, Keterbatasan, dan Saran.
7. **Daftar Pustaka:** Format APA 7th Edition (35+ referensi mutakhir, hanging indent 1.27 cm).

### C. Antarmuka Pengguna & Launcher Terpadu:
- **`taskpane/taskpane.html` & `taskpane.js`:**
  - Dilengkapi pemilih mode: `⚡ Native COM (Direct Word - Rekomendasi)` vs `📄 Office.js (Taskpane)`.
  - Tombol unduh langsung Dataset Excel SmartPLS (`output/DATASET_SMARTPLS_45_RESPONDEN.xlsx`).
  - Tombol 1-klik salin script otomatis Google Form (`.gs`).
  - Tombol khusus ekspor mandiri (.docx) per bab: Bagian Awal & Bab I.
  - Tombol 1-klik buka dokumen tesis lengkap di Word.
- **`run_plugin.bat`:** Launcher 1-klik yang otomatis memicu pembuatan sertifikat SSL, menyalakan server, dan membuka antarmuka dashboard.
- **`manifest.xml`:** Diperbarui ke `https://localhost:5051` dan disinkronkan ke `C:\Addins\autotesis_manifest.xml`.

---

## 2. Progres Modul Agen & Skill Spesialis Per Bab

1. ✅ **Modul 1: Bagian Awal (Front Matter):**
   - Skill: `tesis-bagian-awal` (Global & Workspace).
   - Engine: `agent_bagian_awal.py`.
   - Output: `output/BAGIAN_AWAL_LENGKAP_NURMAYANTI.docx` (43.5 KB, 12 komponen lengkap).
2. ✅ **Modul 2: BAB I Pendahuluan:**
   - Skill: `tesis-bab-1-pendahuluan` (Global & Workspace).
   - Engine: `agent_bab_1.py`.
   - Output: `output/BAB_1_PENDAHULUAN_NURMAYANTI.docx` (41.9 KB, Tabel 1.1 Kriteria IKU, Tabel 1.2 Capaian 2022-2024, 7 Rumusan, 7 Tujuan, Manfaat).
3. ✅ **Modul 3: BAB II Kajian Pustaka & Hipotesis:**
   - Skill: `tesis-bab-2-kajian-pustaka` (Global & Workspace).
   - Engine: `agent_bab_2.py`.
   - Output: `output/BAB_2_KAJIAN_PUSTAKA_NURMAYANTI.docx` (43.9 KB, Grand Theory, Matriks 10 Jurnal, 7 Alur Logika, Gambar 2.1, 7 Hipotesis Ilmiah).
4. ✅ **Modul 4: BAB III Metode Penelitian:**
   - Skill: `tesis-bab-3-metode-penelitian` (Global & Workspace).
   - Engine: `agent_bab_3.py`.
   - Endpoint: `/api/generate-bab-3`.
   - UI: Tombol `✍️ Tulis Word` & `📄 File .docx` di `taskpane.html` & `taskpane.js`.
   - Output: `output/BAB_3_METODE_PENELITIAN_NURMAYANTI.docx` (46.8 KB, Jadwal Riset Tabel 3.1, Sensus N=45, Matriks 36 Indikator Tabel 3.2, PLS-SEM Outer/Inner Model, Bootstrapping 5.000 Tabel 3.3).
5. ✅ **Modul 5: BAB IV Hasil Penelitian & Pembahasan:**
   - Skill: `tesis-bab-4-hasil-dan-pembahasan` (Global & Workspace).
   - Engine: `agent_bab_4.py`.
   - Endpoint: `/api/generate-bab-4`.
   - UI: Tombol `✍️ Tulis Word` & `📄 File .docx` di `taskpane.html` & `taskpane.js`.
   - Output: `output/BAB_4_HASIL_DAN_PEMBAHASAN_NURMAYANTI.docx` (46.1 KB, Profil Responden N=45 Tabel 4.1, Deskriptif Tabel 4.2, Outer Loadings 36 Indikator Running 1 & 2 Tabel 4.3, Construct Reliability & AVE Tabel 4.4, HTMT Tabel 4.5, R2 & f2 Tabel 4.6 & 4.7, Bootstrapping 5.000 7 Hipotesis Tabel 4.8, Pembahasan Formula Baku 4 Alinea per Hipotesis).
6. ✅ **Modul 6: BAB V Kesimpulan, Saran & Daftar Pustaka:**
   - Skill: `tesis-bab-5-dan-dapus` (Global & Workspace).
   - Engine: `agent_bab_5.py`.
   - Endpoint: `/api/generate-bab-5`.
   - UI: Tombol `✍️ Tulis Word` & `📄 File .docx` di `taskpane.html` & `taskpane.js`.
   - Output: `output/BAB_5_KESIMPULAN_DAN_DAPUS_NURMAYANTI.docx` (42.1 KB, Kesimpulan menjawab 7 Hipotesis, Implikasi Teoretis & Manajerial, Keterbatasan, Saran Praktis & Akademis, serta Daftar Pustaka APA 7th Edition 35+ Referensi bereputasi).

---

## 3. Cara Menjalankan & Menggunakan Proyek:
1. **Jalankan:** Klik ganda file [run_plugin.bat](file:///e:/SKILL%20UNTUK%20AI/system_prompts_leaks-main/word-addin-tesismagister/run_plugin.bat) atau jalankan `python backend\server.py`.
2. **Buka di Browser:** Kunjungi `http://localhost:5050/taskpane.html` (Bisa digunakan berdampingan langsung dengan Microsoft Word).
3. **Pilih Bagian/Bab:** Klik tombol bab yang diinginkan, naskah akan langsung terketik secara otomatis ke dalam dokumen Word yang sedang Anda buka!
4. **Atau 1-Klik Tesis Utuh:** Klik tombol hijau di tab ke-4, file Word utuh 5 Bab akan otomatis dibuat dan langsung dibuka di Microsoft Word.

---

## 4. Fitur Baru: Telusuri Dimensi & Indikator via Live Academic API (Dimensi Terverifikasi)

**Inti keputusan pengguna:** dimensi dan indikator harus benar-benar nyata (jadi alat ukur kuesioner), cukup 1 sumber per variabel, format kisi-kisi seperti contoh, dan hasilnya disimpan permanen agar bisa dipakai ulang untuk judul baru.

### Modul Baru:
1. **`backend/doc_text_extract.py`** — ekstraksi teks PDF (PyMuPDF) / DOCX / TXT, unduh PDF Open Access, dan fungsi `quote_in_source()` (verifikasi kutipan indikator ada di teks sumber).
2. **`backend/variable_library.py`** — pustaka permanen `backend/variable_library/<slug>.json`; satu variabel = satu sumber aktif; `verify_indicators()` membuang indikator yang tidak punya bukti kutipan; `entry_needs_review()` menolak sumber di luar 2022-2026 untuk hasil pencarian.
3. **`backend/variable_tracer.py`** — alur penelusuran: cek pustaka → **pencarian multi-sumber paralel** (OpenAlex + Crossref + Semantic Scholar, 2022-2026, has_doi) → ungkap PDF via Unpaywall → unduh PDF → ekstraksi LLM HANYA dari teks sumber dengan kutipan literal → verifikasi kode → simpan permanen. Ambang diterima: minimal 3 dimensi & 8 indikator lolos verifikasi. Jalur `trace_from_upload()` untuk berkas pengguna (`.pdf/.docx/.txt`).

### Pencarian Multi-Sumber (baru, 14 Okt 2026 malam):

**Masalah awal:** semua kandidat berasal dari **satu titik tunggal (OpenAlex)**. Begitu OpenAlex memberi `429 Retry-After` berjam-jam, seluruh telusur mandek dan semua variabel masuk `needs_review` — persis yang terjadi pada uji *Employee Engagement*.

**Solusi:** tambah sumber berbayar-gratis tanpa API key, dijalankan **paralel** dengan rate limit terpisah, sehingga bila satu sumber diblokir yang lain tetap menyuplai kandidat.

| Sumber | Peran | Butuh key |
|---|---|---|
| OpenAlex | kandidat + PDF Open Access langsung | tidak |
| **Crossref** | registry DOI terbesar (~150 jt) → cakupan | tidak |
| **Semantic Scholar** | PDF Open Access langsung | tidak |
| **Unpaywall** | DOI → PDF Open Access (menyelamatkan kandidat Crossref yang belum punya PDF) | tidak |
| ~~DOAJ~~ | **diuji lalu dibuang** — latensi terukur 24–64 dtk per kueri | — |

**Hasil pengukuran terkontrol (kueri "Employee Engagement dimensi indikator"):**

| Kondisi | OpenAlex saja | Multi-sumber |
|---|---|---|
| OpenAlex **mati total** (429) | **44,5 dtk → 0 kandidat** | **5,3 dtk → 8 kandidat** |
| OpenAlex sehat | 20,5 dtk → 4 kandidat | ~21 dtk → **27 kandidat unik** |
| Unduh PDF kandidat | — | **6/8 berhasil**, teks 29rb–48rb karakter |

**Detail implementasi:**
- `_search_all()` — orkestrator paralel (`ThreadPoolExecutor`, satu worker per pekerjaan agar tak ada yang mengantre di belakang sumber yang menahan backoff).
- **Profil cepat** saat multi-sumber: cuma 2 percobaan + backoff pendek (2 dtk). Prinsipnya: *dengan adanya redundansi, sumber yang tertahan sebaiknya menyerah cepat* — dulu backoff 5/10/20/40 dtk justru menahan hasil sumber yang sehat.
- **Dua pengaman kecepatan:** batas waktu total `SEARCH_WALL_MAX=25s`, dan pelepasan awal begitu kandidat mencukupi (`SEARCH_ENOUGH` = jumlah slot percobaan) setelah `SEARCH_MIN_WAIT=8s` berlalu.
- **Unduh PDF paralel** (`_download_all`, 4 worker) — tahap I/O-bound terlama yang dulu serial.
- **Pengurutan relevansi** `_relevance()`: slot percobaan diisi makalah yang paling mungkin memuat instrumen (nama variabel di judul/abstrak + kata *dimensi/indikator*), bukan sekadar terbaru. Tanpa ini, makalah "Green HRM"/"Talent Management" menghabiskan 8 slot dan ekstraksi gagal walaupun LLM sehat.
- `source.type` kini **mencatat sumber asli** (`crossref`/`openalex`/`semantic_scholar`), bukan lagi selalu ditulis `openalex`.

4. **`backend/tokoh_verifier.py`** — validasi tokoh ahli melalui OpenAlex `/authors`; nama tokoh yang tidak ditemukan (fiktif/karangan) otomatis dibuang dari `tokoh_ahli`. Bukti (`tokoh_proofs`) disimpan ke pustaka.

### Integrasi:
- `llm_ontology_researcher.research_variable_llm()` kini: **pustaka → telusur live → None** (gagal = `needs_review`, tidak pernah mengarang dimensi). Versi lama yang mengarang dimensi dari memori dinonaktifkan (`_legacy_free_research_variable_llm`).
- `academic_dimension_researcher.research_variable_ontology()` — pesan `needs_review` kini mengarahkan pengguna mengunggah jurnal/buku khusus.
- **`backend/server.py` endpoint baru:** `POST /api/variable/trace`, `POST /api/variable/upload` (base64), `POST /api/variable/list`.
- **`taskpane.html` + `taskpane.js`:** panel unggah jurnal/buku khusus (dropdown variabel X1/X2/M/Y, pilih berkas, tombol ekstrak & simpan).

### Hasil Uji Terakhir (14 Okt 2026, 22:00–23:00 WIB — end-to-end di browser):
- **Uji Tab 1 (judul BPKAD Penukal Abab Lematang Ilir, penulis AKBAR BILLY ADAM / NPM 245241059):** deteksi otomatis menghasilkan X1 Employee Engagement, X2 Work Life Balance, M **Kepuasan Kerja**, Y Kinerja Pegawai, instansi benar. Tombol **Telusuri** dijalankan → **keempat variabel (4/4) akhirnya tersimpan**: Work Life Balance (Badarudin et al. 2025, DOI 10.31004/joecy.v5i3.5057), Kepuasan Kerja (Sari et al. 2025), Kinerja Pegawai (dipakai ulang dari pustaka), dan **Employee Engagement (Puspitasari & Aulia 2025, DOI 10.22441/indikator.v9i2.32922, 3×3 indikator)**.
  - *Employee Engagement* **awalnya GAGAL** (`needs_review`) karena OpenAlex memberi 429 `Retry-After: 28666` (≈8 jam). **Setelah lapisan multi-sumber dipasang, ia berhasil** — sumber pemenangnya justru **Crossref** (`type=crossref`), di saat OpenAlex dan Semantic Scholar mati total. DOI pemenang diverifikasi resolusi ke doi.org → **HTTP 200**.
- **Uji Tab 2 (generate sub-bab 1.1 Latar Belakang):** `POST /api/pipeline/generate-subchapter` → **HTTP 200**; 13 paragraf utuh (nasional → teori tiap variabel → profil objek → fenomena → research gap → penutup), sitasi terpakai tepat, paragraf penutup menyisipkan judul lengkap. Kualitas prosa Indonesia teruji baik.
- **Audit pustaka: 11 entri, semua lolos** — tahun 2022–2025, semua ≥3 dimensi & ≥8 indikator. Semua entri lama `source.type = openalex` (sumber dari API, bukan karangan LLM); entri baru *Employee Engagement* `type = crossref`. **DOI uji acak diresolusi ke doi.org → HTTP 200** (nyata & terbuka).
- **Uji kritis ekstraksi Nemotron berbahasa Indonesia:** 5/5 indikator menghasilkan kutipan literal yang cocok persis di teks sumber (syarat `quote_in_source` terpenuhi).
- Parse judul via endpoint `/api/parse-title` mengembalikan `data.m = "Kepuasan Kerja"` benar dengan `engine: gemini:nvidia/nemotron-3-ultra-550b-a55b` — membuktikan jalur fallback Nemotron bekerja nyata.
- End-to-end HTTP `/api/variable/list|trace|upload` teruji OK. Panel unggah taskpane teruji di browser.
- Pustaka kini **11 entri nyata ber-DOI**: Employee Engagement, ETOS KERJA, Gaya Kepemimpinan, Kecerdasan Emosional, Kepuasan Kerja, Kesejahteraan Pegawai, Kinerja, Kinerja Pegawai, Motivasi Kerja, Produktivitas Kerja, Work Life Balance.
- Server dual 5050/5051 sudah di-restart dengan kode final; seluruh skrip uji sementara (`_uji_*.py`) telah dihapus.

---

## 5. Arsitektur Rantai LLM & Fallback (Baru)

**Keputusan pengguna:** Gemini tetap utama; bila Gemini gagal, fallback terakhir adalah **NVIDIA Nemotron 3 Ultra** (`nvidia/nemotron-3-ultra-550b-a55b`).

1. **Rantai model di `backend/llm_config.json`** (disusun ulang 15 Okt 2026 dari pengukuran nyata): `gemini-flash-lite-latest` → `gemini-3.5-flash-lite` → `gemini-3.1-flash-lite` → `gemini-flash-latest` → **`nvidia/nemotron-3-ultra-550b-a55b`** (urutan terakhir).
   - **Dibuang: `gemini-2.5-flash` dan `gemini-2.5-flash-lite`.** Uji 15 Okt 2026 membuktikan keduanya **SUDAH DIHAPUS Google** — membalas `404 This model ... is no longer available` pada **semua 15 key**. Versi sebelumnya masih memuat keduanya, sehingga **setiap rantai membuang 2 percobaan penuh** (termasuk jeda backoff) untuk sesuatu yang pasti gagal. Inilah asal kegagalan `404` yang dulu disangka "nama model salah" — ternyata modelnya memang sudah mati.
   - Urutan disusung dari hasil uji: `gemini-flash-lite-latest` dan `gemini-3.5-flash-lite` terbukti **200** dan menang pada uji ekstraksi; `gemini-3.1-flash-lite` saat itu sering **503** sehingga dimundurkan.
2. **Dispatcher `_call_llm_json()` di `backend/llm_ontology_researcher.py`:** model berawalan `nvidia/` dialihkan ke endpoint NIM OpenAI-compatible (`POST https://integrate.api.nvidia.com/v1/chat/completions`, auth `Bearer`); sisanya ke Gemini. Alias `_call_gemini_json = _call_llm_json` sehingga 4 pemanggil lama (`llm_background_writer`, tracer, parser judul, riset ontologi) ikut otomatis tanpa mengubah kode.
3. **Adapter:** `_call_nvidia_json()` + `_extract_json()` (parsing longgar + bukti blok JSON). Parser judul (`llm_title_parser._call_gemini`) ikut rute NVIDIA via lazy import.
4. **Kunci di `llm_config.json`:** **`api_keys` (daftar, 15 key)** dan `api_key` (tunggal, untuk config lama) untuk Gemini, plus `nvidia_api_key` — keduanya dilindungi `.gitignore`. Pembaca `gemini_keys(cfg)` menyatukan keduanya dan membuang placeholder `GANTI_...`.

### Multi-Key: 15 API Key Gemini (Diuji 15 Okt 2026)

Pengguna menyerahkan **14 API key** tambahan untuk dipakai sekaligus. Hasil pengujian — dan ini penting agar tidak salah harap:

| Pertanyaan | Hasil uji |
|---|---|
| Berapa model per 1 key? | **63 model**, di antaranya **46 mendukung `generateContent`** |
| Apakah 14 key = 14× kuota? | **TIDAK.** Ke-14 key memberi **429 kuota IDENTIK** pada **dua model berbeda** (`gemini-pro-latest`, `gemini-3.1-pro-preview`) → kuotanya **berbagi satu proyek Google** |

Konsekuensinya jujur: **rotasi antar-key TIDAK melipatgandakan kuota.** Yang membuat kuota tersedia justru **pemilihan model** — tier `pro` terbukti habis (429) sedangkan `flash-lite` masih hidup (200).

Rotasi tetap dipasang, tetapi untuk **ketahanan**, bukan kapasitas: bila sebuah key **ditolak** (HTTP `401`/`403` — dicabut atau salah proyek), pemanggil otomatis berpindah ke key berikutnya dari daftar melingkar (`_next_gemini_key`, batas `GEMINI_KEY_ROTATE=3`). Diterapkan di `_call_gemini_raw` dan `llm_title_parser._call_gemini`. URL dibentuk **di dalam** loop agar kunci dapat diganti tanpa menyusun ulang panggilan.

**Peringatan keamanan:** ke-15 key kini tersimpan di `backend/llm_config.json` (terlindungi `.gitignore`). Key-key itu sempat tertempel di percakapan — bila ingin ketenangan penuh, buat key baru di AI Studio lalu cabut yang lama.

### Pengendali Laju & Ketahanan:
- **Gemini:** `GEMINI_MIN_INTERVAL=4.5s`, `GEMINI_MAX_RETRY=1`, `GEMINI_BACKOFF=6.0`.
- **Circuit breaker kuota Gemini (PER MODEL):** saat 429 berbunyi *"exceeded your current quota"*, **model itu saja** ditandai mati 600 dtk (`_gemini_quota_dead_until = {model: epoch}`) dan dilewati instan. **Perbaikan penting:** versi awalnya global — kuota habis pada `gemini-2.5-flash` ikut memblokir `gemini-3.1-flash-lite` yang justru sehat, sehingga memaksa jatuh ke Nemotron padahal tidak perlu. Kini tiap model punya kuota sendiri.
- **Circuit breaker kesibukan (503) — baru 15 Okt 2026:** pemutus kuota tadi **hanya** bereaksi terhadap 429-kuota, sehingga model yang 503-menahun tetap membebankan **~22 dtk backoff pada SETIAP panggilan** sepanjang rantai. Kini bila 503 bertahan **setelah** dicoba ulang, model ditandai sibuk 120 dtk (`_gemini_busy_until`, `GEMINI_BUSY_DEAD_SEC`) dan dilewati dalam **0,000 dtk**. Terukur pada uji: `gemini-3.1-flash-lite` 503 dua kali → otomatis ditandai sibuk.
- **Rotasi key saat ditolak:** lihat bagian *Multi-Key* di atas — aktif pada HTTP 401/403, tidak pada 429 (karena kuota memang dibagi).
- **OpenAlex:** masuk *polite pool* via `&mailto=tesismagister@research.ac.id`, jeda antar pencarian, exponential backoff 429 (5/10/20/40 dtk, 4 percobaan).
- **Nemotron:** retry 1× pada 503 dengan jeda 6 dtk.

### Dampak Terukur dari Pembersihan Rantai (15 Okt 2026):
Simulasi rantai produksi dengan prompt ekstraksi yang sama, dijalankan sebelum dan sesudah:

| | Sebelum | Sesudah |
|---|---|---|
| Waktu rantai | **138,8 dtk** | **0,9 dtk** |
| Penyebab | model #1 (`gemini-3.1-flash-lite`) 503 → 22 dtk terbuang; pemenang baru muncul di #2 setelah 131,8 dtk | model #1 (`gemini-flash-lite-latest`) langsung menang |
| Model 503-mentahun | membakar 22 dtk tiap panggilan | dilewati **0,000 dtk** |
| Model 404 (dihapus Google) | 2 percobaan terbuang sia-sia | tidak ada di rantai |
| Parser judul | — | **1,1 dtk**, hasil benar |

### Catatan Risiko Jujur:
- **Bahasa Indonesia tidak tercantum** dalam daftar bahasa didukung Nemotron → berisiko pada mutu prosa & keakuratan kutipan literal. **Hasil uji membuktikan aman** (5/5 kutipan literal cocok), tetapi ini tetap risiko yang perlu dipantau. Gemini tetap diprioritaskan.
- Endpoint gratis Nemotron **memberi 503 bila dibebani banyak panggilan konkuren** (terjadi saat tracer mengekstrak banyak kandidat PDF sekaligus); dari proses bersih respons hanya 0,8 dtk.
- Endpoint gratis punya batas laju → sebaiknya panggilan ekstraksi besar diserialisasi.

### Catatan / Pekerjaan Tersisa:
1. **Verifikasi tokoh memakai nama lengkap** (prompt meminta nama depan + belakang; `tokoh_verifier.py` cocokkan nama belakang + minimal satu nama depan). Uji: "Stephen P. Robbins" match ke Stephen P. Robbins (141 karya), bukan Trevor Robbins.
2. **API key Gemini lama (`AIzaSy...`) bocor di chat → sudah di-rotate** ke key baru. **Konfirmasi definitif (15 Okt 2026):** kuota free-tier dihitung **per proyek Google/AI Studio, bukan per key** — dibuktikan dengan 14 key yang semuanya membalas 429 kuota identik pada dua model berbeda. Jadi **menambah/memutar key tidak menambah kuota**; yang menolong adalah memakai model yang belum habis (`flash-lite`) plus fallback Nemotron. Untuk beban produksi disarankan proyek/API key berbayar.
3. ~~**Employee Engagement** perlu diulang setelah kuota OpenAlex reset~~ — **SELESAI.** Dengan lapisan pencarian multi-sumber, variabel ini berhasil tersimpan (Puspitasari & Aulia 2025, `type=crossref`, DOI resolusi HTTP 200) justru saat OpenAlex dan Semantic Scholar sedang mati total. Jalur `trace_from_upload()` tetap tersedia untuk berkas milik pengguna.
4. Waktu telusur 4 variabel bisa puluhan menit (unduh PDF + ekstraksi per variabel) dan rawan 429.
5. **Mengurangi beban:** `llm_background_writer` masih mengulang `time.sleep(5)` untuk tiap model Gemini yang mati (loop `for attempt in range(2)`), sehingga *circuit breaker* hanya mempercepat kegagalan, tidak menghilangkan jeda itu. Perlu disesuaikan agar tidak ikut tidur saat model sudah ditandai mati.

