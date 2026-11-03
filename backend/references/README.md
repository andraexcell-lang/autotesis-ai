# references/ — Sumber Dimensi & Indikator S1

Folder ini menyimpan PDF sumber rujukan yang dipakai mesin S1 untuk mengisi
kartu **Sumber** di blueprint, dengan konten dimensi/indikator yang benar-benar
berasal dari PDF (bukan karangan).

```
references/
├── 01_primers/                     ← sumber UTAMA dimensi & indikator (1 per variabel)
├── 02_pendukung/                   ← jurnal/buku PEMBANDING untuk Bab II
├── reference_map.json              ← peta: PDF -> variabel -> peran -> sitasi APA
└── pendukung.json                  ← (dibuat tool) registry referensi pendukung
```

## Peran

| Peran | Fungsi | Aturan |
|---|---|---|
| `primer` | Sumber utama dimensi & indikator, menjadi "Sumber" di blueprint | **1 variabel = 1 primer** (tool menolak duplikat) |
| `pendukung` | Jurnal/buku pembanding, definisi, grand theory untuk Bab II | Tidak menimpa sumber primer mana pun |

## Mode impor (per entri primer)

| Mode | Arti |
|---|---|
| `develop` | Dimensi dari sumber + indikator dikembangkan (blueprint-siap). Entri kurasi yang sudah ada **dipertahankan**; bila belum ada, mesin mengembangkannya dari PDF (API). Audit kutipan literal tetap dicetak. |
| `literal` | **Setiap** indikator wajib kutipan literal dari teks PDF (anti-halusinasi). Indikator yang tidak cocok dibuang. |
| `auto` | Ekstraksi otomatis penuh oleh mesin (`variable_tracer`, LLM) dari teks PDF. |

## Menambah referensi baru

> **Impor massal dari folder PDF:** `_build_map_entries.py` di `backend/` membaca
> `_PETA_PROPOSAL.csv` (hasil review dengan kolom `keputusan`), mengekstrak metadata
> terbaik dari tiap PDF (judul/penulis/tahun/jurnal/DOI), lalu menghasilkan
> `_MAP_PENDUKUNG.json`; `_merge_map.py` menggabungkannya ke `pendukung[]` pada
> `reference_map.json`. Setiap entri memakai `"source": "<path absolut>"` dan tool
> menyalin PDF-nya otomatis ke `references/02_pendukung/` — `source` absolut
> didukung di Windows. Setelah itu jalankan perintah impor di bawah.

### 1) Sumber primer (variabel baru / ganti sumber)

1. Taruh PDF di `references/` (atau `references/01_primers/`).
2. Tambah satu objek ke `primers[]` di `reference_map.json`:
   ```json
   {
     "file": "komitmen_organisasi_jurnal_2025.pdf",
     "variable": "Komitmen Organisasi",
     "role": "primer",
     "mode": "develop",
     "label": "Penulis (2025)",
     "apa": "Penulis, A. B. (2025). Judul Artikel. Nama Jurnal, 12(3), 100-110.",
     "year": 2025,
     "doc_kind": "artikel",
     "journal": "Nama Jurnal, 12(3), 100-110",
     "doi": "https://doi.org/xxxx (bila ada, opsional)",
     "grand_theory": "Nama Grand Theory",
     "tokoh_ahli": "Ahli (Tahun)"
   }
   ```
   > Bila PDF masih di folder lain, isi `"source": "lokasi/asal.pdf"` dan tool akan
   > menyalinnya otomatis.
3. Jalankan:
   ```
   python import_references.py import --clear-cache
   ```

### 2) Sumber pendukung (pembanding Bab II)

1. Taruh PDF di `references/02_pendukung/`.
2. Tambah satu objek ke `pendukung[]` di `reference_map.json`:
   ```json
   {
     "file": "02_pendukung/jurnal_pembanding_2024.pdf",
     "source": "folder_asal/jurnal_pembanding_2024.pdf",
     "variable": "Komitmen Organisasi",
     "role": "pendukung",
     "label": "Penulis (2024)",
     "apa": "Penulis, C. D. (2024). Judul. Nama Jurnal.",
     "year": 2024,
     "doc_kind": "artikel"
   }
   ```
3. Jalankan perintah yang sama. Referensi terdaftar di `pendukung.json`.

## Perintah

```
python import_references.py import [--map references/reference_map.json] [--clear-cache] [--dry-run]
python import_references.py status
python import_references.py --clear-cache
```

`--clear-cache` menghapus cache blueprint S1 agar hasil baru langsung terlihat;
`--dry-run` hanya simulasi tanpa menulis apa pun.

## Jaminan kualitas

- **Anti-halusinasi:** mode `literal` membuang indikator yang tidak ditemukan
  verbatim di teks PDF; mode lain tetap mencetak laporan audit literal.
- **Tidak menimpa entri lama dengan data kosong:** bila ekstraksi gagal/terlalu
  pendek, entri pustaka yang ada dipertahankan.
- **Satu variabel = satu primer:** duplikat ditolak sebelum kerja apa pun.
- **Tautan pasti resolve:** `source.url` memakai `/references/<file>` (PDF lokal),
  bukan DOI yang bisa mati. DOI yang terdaftar tetap tampil sebagai tombol tambahan.
