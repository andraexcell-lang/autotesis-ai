# DOKUMENTASI LENGKAP PROYEK: AUTOTESIS AI
> **Handover Dossier & Developer Guide untuk OpenCode**  
> **Terakhir Diperbarui:** 14 Oktober 2026  
> **Platform Target:** Microsoft Word Add-in (Taskpane) & Autonomous Thesis Drafting Engine  
> **Standar Mutu:** Tesis Magister Manajemen (S2) Pascasarjana Universitas Tridinanti Palembang (Model Yurahadi / Andre Juliansya / Muhamad Yusi)  
> **Model Riset:** Kausal Asosiatif Model Mediasi 4 Variabel ($X_1, X_2 \rightarrow M \rightarrow Y$) dengan 7 Hipotesis Berarah & Analisis PLS-SEM (SmartPLS 4).

---

## 1. Ringkasan Eksekutif & Tujuan Proyek

**AutoTesis AI** adalah sistem kecerdasan buatan berbasis *hybrid agent* yang dirancang khusus untuk memfasilitasi asistensi dan penulisan naskah Tesis Magister Manajemen (S2) secara otomatis, komprehensif, dan presisi tinggi langsung ke dalam lembar kerja Microsoft Word.

### Keunggulan Teknis Utama:
1. **Solusi Anti-Truncation (Hierarchical Sub-Chapter Pipeline):** Memecah naskah tesis menjadi **24 unit sub-bab/sub-sub-bab mandiri**, membebaskan LLM dari batasan limit token dan *summarization bias*, sehingga mampu menghasilkan karya ilmiah utuh berstandar tebal tesis pascasarjana (**110–150 halaman**).
2. **Arsitektur "Benang Merah Baku" (Golden Thread Architecture):** Menjamin seluruh komponen dari Judul $\rightarrow$ Bab I $\rightarrow$ Bab II $\rightarrow$ Bab III $\rightarrow$ Bab IV $\rightarrow$ Bab V $\rightarrow$ Lampiran saling terikat secara mutlak:
   - Indikator Variabel $Y \rightarrow$ Mengunci parameter baris Tabel 1.1 (Capaian Kinerja 3 Tahun).
   - Dimensi Variabel $M \rightarrow$ Mengunci fokus masalah Tabel 1.2 (Rekapitulasi Kendala 12 Bulan).
   - Indikator $X_1$ dan $X_2 \rightarrow$ Mengunci telaah fenomena lapangan dan *research gap* di Bab I.
   - Grand Theory, dimensi, dan 10 komparasi jurnal $\rightarrow$ Tersinkronisasi di Bab II.
   - Tabel Kisi-kisi Operasional Variabel (Tabel 3.2 s/d 3.5) $\rightarrow$ Mengunci indikator yang sama persis di Bab III.
   - Pernyataan kuesioner angket Likert 1–5 $\rightarrow$ Terbentuk otomatis di Lampiran.
3. **Dual Writing Transport (Office.js & Native COM):**
   - **Mode Office.js:** Menyisipkan naskah berbasis Base64 langsung ke dokumen aktif Word via API Office.js.
   - **Mode Native COM (`pywin32`):** Mengetik dan menyisipkan naskah melalui antarmuka Windows COM Word (`Selection.InsertFile`) tanpa hambatan keamanan browser WebView2.
4. **Unified Cumulative Bibliography (Daftar Pustaka Cerdas):** Algoritma yang secara otomatis mendeteksi blok Daftar Pustaka lama saat sub-bab baru diinjeksi, menghapusnya, menyambungkan teks baru di halaman yang sama, lalu menaruh Daftar Pustaka kumulatif di halaman paling akhir.

---

## 2. Topologi & Diagram Alir Sistem

```
┌────────────────────────────────────────────────────────────────────────┐
│                        MICROSOFT WORD DESKTOP / BROWSER                │
│                                                                        │
│   ┌────────────────────────────────┐    ┌──────────────────────────┐   │
│   │   DOKUMEN WORD AKTIF (.DOCX)   │    │  TASKPANE ADD-IN UI      │   │
│   │                                │    │  (taskpane.html & .js)   │   │
│   │  • Format Kertas: A4           │◄───┤  • Input Judul Cerdas    │   │
│   │  • Margin: 4-4-3-3 cm          │    │  • Live Title Parser     │   │
│   │  • Font: TNR 12 pt, 1.5 Spasi  │    │  • Card Ontologi Riset   │   │
│   │  • Tabel Terbuka Akademik      │    │  • 24 Tombol Sub-Bab     │   │
│   │  • Unified Daftar Pustaka APA7 │    │  • Tombol Run Chain      │   │
│   └────────────────────────────────┘    └────────────┬─────────────┘   │
└──────────────────────────────────────────────────────┼─────────────────┘
                                                       │ Dual Transport (5050/5051)
                                                       ▼
┌────────────────────────────────────────────────────────────────────────┐
│                     AUTOTESIS DUAL SERVER (server.py)                  │
│                                                                        │
│   [Port 5050: HTTP Web Browser]    [Port 5051: HTTPS Word Add-in SSL]  │
│                                                                        │
│   ┌─────────────────────────────────┐ ┌────────────────────────────┐   │
│   │ 1. Semantic Title Parser        │ │ 2. Academic Dimension Res. │   │
│   │ • 10+ Pola Gramatikal Judul     │ │ • 12+ Ontologi Baku MSDM   │   │
│   │ • Isolasi Lokus & 10 Sektor     │ │ • Indikator Likert 1-5     │   │
│   │ • Pertahankan Akronim Daerah    │ │ • Jurnal Bereputasi & DOI  │   │
│   └────────────────┬────────────────┘ └─────────────┬──────────────┘   │
│                    │                                │                  │
│                    └────────────────┬───────────────┘                  │
│                                     ▼                                  │
│   ┌────────────────────────────────────────────────────────────────┐   │
│   │ 3. Hierarchical Sub-Chapter Pipeline (24 Unit Mandiri)         │   │
│   │ • Anti-Token Truncation & Anti-Summarization Bias (120+ Hal)   │   │
│   │ • Penguncian "Benang Merah":                                   │   │
│   │   - Indikator Y  --> Parameter Baris Tabel 1.1 (Capaian 3 Thn) │   │
│   │   - Dimensi M    --> Isu Operasional Tabel 1.2 (Rekap 12 Bln)  │   │
│   │   - Indikator X1 & X2 --> Narasi Fenomena & Research Gap Bab I │   │
│   │   - Grand Theory, Dimensi, 10 Jurnal, Kerangka di Bab II       │   │
│   │   - Tabel 3.2 s/d 3.5 Kisi-Kisi Operasional Instrumen di Bab III│  │
│   │   - Pembahasan Formula Baku 4 Alinea di Bab IV (7 Poin H1-H7)  │   │
│   │   - Kesimpulan Menjawab Angka Statistik di Bab V               │   │
│   └────────────────────────────────┬───────────────────────────────┘   │
│                                    ▼                                   │
│   ┌────────────────────────────────────────────────────────────────┐   │
│   │ 4. Output & Injection Drivers                                  │   │
│   │ • Office.js Driver (Base64 file insertion ke dokumen aktif)    │   │
│   │ • Native Word COM Driver (pywin32 Selection.InsertFile)        │   │
│   │ • SmartPLS 4 Dataset Simulator (Excel 100% Lolos Uji Statistik)│   │
│   │ • Google Form Generator (.gs Google Apps Script otomatis)      │   │
│   └────────────────────────────────────────────────────────────────┘   │
└────────────────────────────────────────────────────────────────────────┘
```

---

## 3. Struktur Direktori & File Proyek

```
e:/SKILL UNTUK AI/system_prompts_leaks-main/word-addin-tesismagister/
├── manifest.xml                     <- Manifest resmi Add-in Microsoft Word (HTTPS 5051)
├── run_plugin.bat                   <- Launcher 1-klik untuk menyalakan server dan antarmuka
├── install_sertifikat_word.bat      <- Script instalasi sertifikat SSL ke Trusted Root Windows
├── BLUEPRINT_ARSITEKTUR.md          <- Catatan blueprint arsitektur awal
├── RESUME_PROYEK_AUTOTESIS.md       <- Ringkasan status dan histori pengembangan
├── DOKUMENTASI_PROYEK_OPENCODE.md   <- File dokumentasi teknis ini
│
├── backend/
│   ├── server.py                    <- Dual HTTP (5050) & HTTPS (5051) Threaded Server
│   ├── semantic_title_parser.py     <- Parser semantik judul tesis (X1, X2, M, Y, Lokus)
│   ├── academic_dimension_researcher.py <- Riset teori, dimensi, indikator & kisi-kisi
│   ├── cache_thesis_ontology.json   <- Cache ontologi aktif yang mengunci Benang Merah
│   ├── engine_hierarchical_pipeline.py <- Engine utama 24 unit sub-bab mandiri
│   ├── engine_docx.py               <- Helper python-docx (margin 4-4-3-3, tabel, spasi)
│   ├── engine_word_native.py        <- Driver otomatisasi Microsoft Word via Windows COM (pywin32)
│   ├── engine_smartpls.py           <- Simulator data responden Likert 1-5 & lolos kriteria Hair et al.
│   ├── engine_pembahasan_poin.py    <- Generator 4 alinea pembahasan per hipotesis (H1-H7)
│   ├── academic_connector.py        <- Fetcher referensi jurnal OpenAlex & Crossref dengan DOI
│   ├── build_full_thesis_docx.py    <- Builder dokumen utuh gabungan 5 Bab
│   ├── generate_cert.py             <- Pembuat sertifikat SSL localhost mandiri
│   ├── localhost.crt / key / cer    <- Sertifikat SSL resmi Office Add-in
│   └── output/                      <- Folder penyimpanan file .docx dan .xlsx hasil generate
│
└── taskpane/
    ├── taskpane.html                <- Tampilan antarmuka sidebar Office Add-in
    ├── taskpane.css                 <- Styling modern responsif (Emerald & Slate Glassmorphism)
    └── taskpane.js                  <- Controller antarmuka, live parsing, dan pemanggil API
```

---

## 4. Penjelasan Detail Modul Inti

### A. Semantic Title Parser (`backend/semantic_title_parser.py`)
Mesin pemroses bahasa alami (*Rule-Based NLP*) yang mampu membedah judul penelitian kuantitatif model mediasi:
- **Pola Sintaksis yang Didukung:**
  1. *Standar:* `PENGARUH X1 DAN X2 TERHADAP Y MELALUI M SEBAGAI VARIABEL INTERVENING PADA [INSTANSI]`
  2. *Inversi:* `PENGARUH X1 DAN X2 TERHADAP M SERTA DAMPAKNYA PADA/TERHADAP Y PADA [INSTANSI]`
  3. *Kalimat Pasif:* `PERAN M DALAM MEMEDIASI PENGARUH X1 DAN X2 TERHADAP Y DI [INSTANSI]`
  4. *Relasional:* `PENGARUH X1 DAN X2 TERHADAP Y YANG DIMEDIASI OLEH M PADA [INSTANSI]`
- **Preservasi Akronim:** Mempertahankan akronim huruf besar seperti `BPKAD`, `BRIDA`, `DPMPTSP`, `DPRD`, `RSUD`, `PALI`, `OKI`, `ASN`, `PNS`, `SLA`, `OSS`, `SPBE`, `IKU`.
- **Deteksi 10 Sektor Organisasi:** Mengidentifikasi sektor objek penelitian (Pengelolaan Keuangan, Kesehatan, Perbankan, Perizinan, Pendidikan, Riset, Legislatif, Korporasi, dll.).

### B. Academic Dimension Researcher (`backend/academic_dimension_researcher.py`)
Basis pengetahuan ontologi baku MSDM S2 lintas **12+ variabel penelitian populer**:
- `kinerja` $\rightarrow$ Goal-Setting Theory (Locke & Latham, 1990)
- `kepuasan_kerja` $\rightarrow$ Value-Percept Theory (Locke, 1976) & Two-Factor (Herzberg)
- `ocb` $\rightarrow$ Social Exchange Theory (Blau, 1964) & Organ (1988)
- `motivasi` $\rightarrow$ Need Theory (McClelland, 1987) & PSM (Perry & Wise, 1990)
- `disiplin` $\rightarrow$ Theory of Planned Behavior (Ajzen, 1991)
- `budaya_organisasi` $\rightarrow$ Organizational Culture Theory (Schein, 1985)
- `engagement_komitmen` $\rightarrow$ Job Demands-Resources / JD-R (Bakker & Demerouti)
- `promosi_jabatan` $\rightarrow$ Expectancy Theory (Vroom) & Meritokrasi
- `mutasi_kerja` $\rightarrow$ Person-Environment Fit Theory (Edwards)
- `pengembangan_karier` $\rightarrow$ Career Development Theory (Hall; Dessler)
- `kecerdasan_emosional` $\rightarrow$ Emotional Intelligence Framework (Goleman)
- `beban_kerja` $\rightarrow$ Cognitive Load Theory & NASA-TLX
- **Fungsi Utama:**
  - `research_thesis_ontology(title, config_override)`: Menghasilkan master blueprint ontologi.
  - `load_thesis_ontology()`: Mengambil data ontologi dari cache lokal `cache_thesis_ontology.json`.

### C. Hierarchical Pipeline Engine (`backend/engine_hierarchical_pipeline.py`)
Engine utama pembuat 24 unit naskah sub-bab mandiri. Katalog unit terdiri dari:
- **Bab I:** Sub 1.1 (Latar Belakang), Sub 1.2 (Identifikasi Masalah), Sub 1.3 (Pembatasan Masalah), Sub 1.4 (Rumusan Masalah 7 butir), Sub 1.5 (Tujuan Penelitian 7 butir), Sub 1.6 (Manfaat Penelitian).
- **Bab II:** Sub 2.1.1 s/d 2.1.4 (Kajian Teori Y, M, X1, X2), Sub 2.2 (Matriks Komparasi 10 Jurnal), Sub 2.3 (Kerangka Berpikir 7 Alur Mediasi & Gambar 2.1), Sub 2.4 (Perumusan 7 Hipotesis).
- **Bab III:** Sub 3.1 (Jadwal Riset), Sub 3.2 (Populasi & Sensus Sampel), Sub 3.3 (Tabel 3.2–3.5 Kisi-Kisi Operasional), Sub 3.4 (Skala Likert), Sub 3.5 (Teknik Analisis PLS-SEM), Sub 3.6 (Kriteria Bootstrapping).
- **Bab IV:** Sub 4.1.1 (Profil Responden), Sub 4.1.2 (Outer Model & Validitas/Reliabilitas), Sub 4.1.3 (Inner Model & 7 Hipotesis Bootstrapping), Sub 4.2 Poin 1 s/d 7 (Pembahasan Formula Baku 4 Alinea per Hipotesis).
- **Bab V:** Sub 5.1 (Kesimpulan Menjawab Angka Statistik), Sub 5.2 (Implikasi Kebijakan), Sub 5.3 (Saran Praktis & Akademis), Sub Dapus (Daftar Pustaka APA 7th Edition).

---

## 5. Standar Format Dokumen Microsoft Word (Yurahadi / Tridinanti)

| Atribut Format | Standar Baku yang Diterapkan |
| :--- | :--- |
| **Ukuran Kertas** | A4 (21 cm x 29.7 cm) |
| **Margin Halaman** | Atas: 4 cm, Kiri: 4 cm, Bawah: 3 cm, Kanan: 3 cm |
| **Jenis Huruf** | Times New Roman |
| **Ukuran Huruf** | Judul Bab: 12 pt (Bold, Center, Uppercase)<br>Sub-Judul: 12 pt (Bold, Left)<br>Isi Teks: 12 pt (Regular, Justify)<br>Tabel & Sumber: 10–11 pt (Regular) |
| **Spasi Teks** | 1.5 Spasi (atau 2.0 spasi untuk naskah penuh) |
| **Format Paragraf** | Menjorok ke dalam (*First Line Indent*) 1.27 cm (0.5 inci) |
| **Format Tabel** | Garis batas terlihat (semua batas), *AutoFit Window*, spasi 1.0 di dalam tabel, teks sumber di bawah tabel tanpa huruf miring (*regular*) |
| **Daftar Pustaka** | APA 7th Edition, urut abjad A–Z, *hanging indent* 1.27 cm (0.5 inci), spasi 1.0, selalu berada di halaman paling akhir |

---

## 6. Katalog REST API Backend

Dual server berjalan pada port:
- `http://localhost:5050` (HTTP Browser & External API)
- `https://localhost:5051` (HTTPS Word Add-in SSL)

### A. Endpoint Pemrosesan Semantik & Ontologi
```http
POST /api/parse-title
Content-Type: application/json

{
  "title": "PENGARUH DISIPLIN KERJA DAN MOTIVASI KERJA TERHADAP KINERJA PEGAWAI DENGAN KEPUASAN KERJA SEBAGAI VARIABEL INTERVENING PADA DINAS PENDIDIKAN KABUPATEN OGAN ILIR"
}
```
**Response JSON:**
```json
{
  "status": "success",
  "title": "...",
  "data": {
    "x1": "Disiplin Kerja",
    "x2": "Motivasi Kerja",
    "m": "Kepuasan Kerja",
    "y": "Kinerja Pegawai",
    "instansi": "Dinas Pendidikan Kabupaten Ogan Ilir",
    "sektor": "Penyelenggaraan Pendidikan dan Peningkatan Mutu Pembelajaran"
  }
}
```

```http
POST /api/research-ontology
Content-Type: application/json

{
  "title": "...",
  "config": {
    "x1": "Disiplin Kerja",
    "x2": "Motivasi Kerja",
    "m": "Kepuasan Kerja",
    "y": "Kinerja Pegawai",
    "instansi": "Dinas Pendidikan Kabupaten Ogan Ilir"
  }
}
```
**Response JSON:**
```json
{
  "status": "success",
  "ontology": {
    "metadata": { "total_indikator": 26, "skala": "Likert 1-5" },
    "variables": { "x1": {...}, "x2": {...}, "m": {...}, "y": {...} },
    "phenomenon_anchors": {
      "tabel_1_1_params": ["Kualitas Kinerja (...)", "Kuantitas Output (...)", "Ketepatan Waktu (...)"],
      "tabel_1_2_issues": ["Kepuasan terhadap Pekerjaan", "Kepuasan terhadap Apresiasi", "..."]
    },
    "kisi_kisi_bab3": [...],
    "kuesioner_lampiran": [...]
  }
}
```

```http
GET /api/get-ontology
```
Mengambil ontologi tesis yang sedang aktif di memori/cache.

---

### B. Endpoint Hierarchical Sub-Chapter Pipeline
```http
GET /api/pipeline/catalog
```
Mengembalikan daftar 24 sub-bab yang dikelompokkan ke dalam 5 Bab.

```http
POST /api/pipeline/generate-subchapter
Content-Type: application/json

{
  "sub_id": "sub_1_1",
  "mode": "officejs",
  "config": { "judul": "...", "penulis": "ANDRE" }
}
```
**Response JSON:**
```json
{
  "status": "success",
  "sub_id": "sub_1_1",
  "filename": "SUB_1_1_1_1_ANDRE.docx",
  "download_url": "/api/download/SUB_1_1_1_1_ANDRE.docx",
  "docx_base64": "UEsDBBQABgAIAAAAIQD...",
  "dapus_filename": "DAFTAR_PUSTAKA_KUMULATIF_ANDRE.docx",
  "dapus_base64": "UEsDBBQABgAIAAAAIQD...",
  "native_written": false
}
```

```http
POST /api/pipeline/run-chain
Content-Type: application/json

{
  "chapter": 1,
  "mode": "officejs",
  "config": { ... }
}
```
Mengeksekusi seluruh sub-bab dalam satu bab berurutan.

---

### C. Endpoint Ekspor Dokumen Mandiri & Olah Data
- `POST /api/generate-bagian-awal`: Membuat dokumen Front Matter (Cover, Lembar Pengesahan, Abstrak, Daftar Isi).
- `POST /api/generate-bab-1` s/d `/api/generate-bab-5`: Membuat file .docx per Bab.
- `POST /api/generate-full-thesis`: Menggabungkan seluruh 5 Bab utuh menjadi satu file master.
- `POST /api/generate-pembahasan-poin`: Menghasilkan 4 alinea pembahasan untuk Poin H1 s/d H7.
- `POST /api/simulate-smartpls`: Menghasilkan dataset Excel responden yang 100% lolos kriteria SmartPLS.
- `POST /api/generate-gform-script`: Menghasilkan script Google Apps Script (`.gs`) kuesioner Google Form.
- `GET /api/download/{filename}`: Mengunduh file dokumen Word atau Excel.

---

## 7. Panduan Menjalankan & Melanjutkan di OpenCode

### Prasyarat Sistem
- **Sistem Operasi:** Windows 10 / 11 (Mendukung COM Automation `win32com`).
- **Python:** Python 3.10+ (Direkomendasikan Python 3.11 – 3.13).
- **Microsoft Word:** Microsoft Office Word 2016, 2019, 2021, atau Office 365.
- **Pustaka Python:**
  ```bash
  pip install python-docx pandas openpyxl pywin32 cryptography
  ```

### Menjalankan Server di OpenCode
Buka terminal di OpenCode dan jalankan perintah:
```powershell
cd "e:\SKILL UNTUK AI\system_prompts_leaks-main\word-addin-tesismagister\backend"
python -u server.py
```
Output yang menandakan server siap:
```
AutoTesis AI Dual Server Aktif!
1. HTTP Web UI  : http://localhost:5050/taskpane.html
2. HTTPS Add-in : https://localhost:5051/taskpane.html
```

### Membuka Antarmuka Pengguna
1. Buka browser: `http://localhost:5050/taskpane.html`.
2. Masukkan judul tesis penelitian Anda di kotak input judul.
3. Sistem secara otomatis memilah 4 variabel ($X_1, X_2, M, Y$) dan objek instansi.
4. Klik tombol **Tulis** pada sub-bab yang diinginkan:
   - Jika mode **📄 Office.js (Taskpane)** dipilih dan add-in terpasang di Word, naskah akan disisipkan ke Word.
   - Jika mode **⚡ Native COM** dipilih, naskah akan langsung terketik otomatis ke lembar kerja Word yang aktif di desktop Anda!

---

## 8. Roadmap Pengembangan Selanjutnya (Untuk OpenCode)

Berikut modul-modul yang dapat dikembangkan lebih lanjut di OpenCode:
1. **Tahap 2: Live Journal Scraper & Citation Validator**
   - Menghubungkan secara langsung ke Semantic Scholar / OpenAlex API untuk mengambil abstrak jurnal Indonesia mutakhir (2022–2026) secara live saat variabel baru dimasukkan.
2. **Tahap 3: Generator Tabulasi Data Mentah Responden Kuesioner (CSV SmartPLS)**
   - Menyediakan tombol 1-klik untuk mengekspor tabulasi data mentah responden dalam format `.csv` yang langsung kompatibel di-import ke SmartPLS 3 dan SmartPLS 4.
3. **Tahap 4: Turnitin Plagiarism Shielding (Paraphrasing Engine)**
   - Menambahkan pengacak gaya bahasa dan sinonim akademik berstandar EYD V guna memastikan skor kemiripan (*similarity index*) Turnitin selalu di bawah 15%.
