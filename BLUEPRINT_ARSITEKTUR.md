# BLUEPRINT ARSITEKTUR: AI THESIS GENERATOR FOR MICROSOFT WORD
**Nama Proyek:** `AutoTesis AI (Office Taskpane & Engine)`
**Standar Mutu:** Tesis Magister Manajemen Yurahadi (5 Bab, Model Mediasi 7 Hipotesis, SmartPLS 4, Margin 4-4-3-3 cm).

---

## 1. Topologi & Diagram Alir Sistem

```
┌────────────────────────────────────────────────────────────────────────┐
│                        MICROSOFT WORD APPLICATION                      │
│                                                                        │
│   ┌────────────────────────────────┐    ┌──────────────────────────┐   │
│   │   DOKUMEN WORD (.DOCX)         │    │ TASKPANE SIDEBAR (UI)    │   │
│   │                                │    │ [Office.js]              │   │
│   │  • Margin: 4-4-3-3 cm          │◄───┤ • Input Judul & Objek    │   │
│   │  • Font: TNR 12 pt, 1.5 Spasi  │    │ • Pemetaan Variabel      │   │
│   │  • Tabel Terbuka Akademik      │    │ • Tombol Generate Bab    │   │
│   │  • Page Numbering Romawi/Angka │    │ • Progress Bar Real-Time │   │
│   └────────────────────────────────┘    └────────────┬─────────────┘   │
└──────────────────────────────────────────────────────┼─────────────────┘
                                                       │ HTTP / REST API
                                                       ▼ (localhost:5050)
┌────────────────────────────────────────────────────────────────────────┐
│                     AUTOTESIS LOCAL BACKEND SERVER                     │
│                                                                        │
│   ┌─────────────────────────────┐   ┌──────────────────────────────┐   │
│   │ 1. Content Synthesizer      │   │ 2. SmartPLS Simulator        │   │
│   │ • Bab I s.d Bab V Prompts   │   │ • Likert 1-5 Non-straight    │   │
│   │ • 4 Alinea Pembahasan       │   │ • Outer Loading >= 0.76      │   │
│   │ • 7 Hipotesis Konsisten     │   │ • T > 1.96, p < 0.05 Sig.    │   │
│   └──────────────┬──────────────┘   └──────────────┬───────────────┘   │
│                  │                                 │                   │
│                  └────────────────┬────────────────┘                   │
│                                   ▼                                    │
│   ┌────────────────────────────────────────────────────────────────┐   │
│   │ 3. OOXML & python-docx Engine (Anthropic Standards)            │   │
│   │ • Dual Width Tables (DXA)                                      │   │
│   │ • PositionalTab Dot-Leaders (Daftar Isi Otomatis)              │   │
│   │ • Open-Border Academic Tables Tanpa Garis Vertikal             │   │
│   └───────────────────────────────┬────────────────────────────────┘   │
│                                   ▼                                    │
│             LLM Connector (9router / Gemini 3.1 Pro / Claude)          │
└────────────────────────────────────────────────────────────────────────┘
```

---

## 2. Struktur Modul & Folder Proyek

Folder proyek: `e:\SKILL UNTUK AI\system_prompts_leaks-main\word-addin-tesismagister\`

```
word-addin-tesismagister/
├── manifest.xml                     <- Manifest resmi Add-in Microsoft Word
├── backend/
│   ├── app.py                       <- Server REST API FastAPI/Flask (port 5050)
│   ├── engine_docx.py               <- Engine pembuat dokumen Word presisi
│   ├── engine_smartpls.py           <- Simulator data & tabel statistik PLS-SEM
│   ├── prompts_tesis.py             <- Prompt chaining untuk 5 Bab lengkap
│   └── requirements.txt             <- python-docx, fastapi, uvicorn, pandas
├── taskpane/
│   ├── index.html                   <- Antarmuka sidebar modern (Dark/Light mode)
│   ├── taskpane.css                 <- Styling glassmorphism premium & responsif
│   └── taskpane.js                  <- Interaksi Office.js ke backend AI
└── run_plugin.bat                   <- Launcher 1-klik untuk menyalakan server & Word
```

---

## 3. Matriks Alur Pembuatan Naskah (5 Bab Otomatis)

| Tahap | Modul | Luaran (*Deliverable*) yang Dihasilkan |
| :--- | :--- | :--- |
| **0. Front Matter** | `engine_docx.create_front_matter` | Cover luar/dalam, Lembar Persetujuan, Pernyataan Bebas Plagiasi, Kata Pengantar, Abstrak (ID & EN), Riwayat Hidup, Daftar Isi, Tabel, & Gambar. |
| **1. Bab I** | `prompts_tesis.bab_1` | Latar Belakang dengan **Tabel Gap Kinerja Empiris/IKU**, Identifikasi, Batasan Masalah, 7 Rumusan, 7 Tujuan, dan Manfaat. |
| **2. Bab II** | `prompts_tesis.bab_2` | Grand Theory, Kajian Variabel, **Tabel Komparasi 10 Jurnal Terdahulu**, Kerangka Berpikir Konseptual, dan **7 Rumusan Hipotesis**. |
| **3. Bab III** | `prompts_tesis.bab_3` | Desain Riset, Populasi & Sampel Sensus, **Tabel Matriks Operasional Variabel (Indikator & Skala Likert 1-5)**, Teknik Analisis PLS-SEM. |
| **4. Olah Data** | `engine_smartpls.generate` | Dataset Excel responden, tabel Karakteristik Responden, Outer Loadings Running 1 & 2, Construct Reliability & Validity (AVE, CR), HTMT, R-Square, f-Square, dan Bootstrapping Path Coefficients (T-Stat & P-Values). |
| **5. Bab IV** | `prompts_tesis.bab_4` | Narasi Hasil Olah Data lengkap + **Pembahasan 4 Alinea per Hipotesis** (Hasil uji empiris, dukungan teori, konfirmasi jurnal terdahulu, implikasi riil lapangan). |
| **6. Bab V** | `prompts_tesis.bab_5` | Kesimpulan menjawab 7 hipotesis, Implikasi Kebijakan/Manajerial, Keterbatasan, dan Saran Rekomendasi. |
| **7. Referensi** | `prompts_tesis.daftar_pustaka` | Daftar Pustaka format APA 7th Edition (urut abjad, gantung 1.27 cm). |

---

## 4. Keunggulan Teknis
1. **Bisa Menyisipkan Langsung ke Word:** Lewat `Office.js`, teks bisa langsung diketikkan ke dokumen yang sedang aktif di Word pengguna.
2. **Bisa Mengekspor File Lengkap:** Menyediakan tombol unduh file naskah utuh `.docx` yang sudah diformat sempurna (margin, heading, nomor halaman berbeda).
3. **Multi-Model LLM:** Bisa menggunakan Gemini via 9router lokal port 20128, Gemini API langsung, atau Claude.
