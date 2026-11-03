# Deploy Gratis: GitHub + Cloudflare Pages + Backend

Panduan ini menjelaskan cara mempublikasikan **Independent Researcher (AutoTesis AI)**
sebagai Word Add-in yang bisa diakses dari mana saja, memakai layanan gratis.

---

## 0. Fakta arsitektur (baca dulu)

| Komponen | Isi | Bisa di Cloudflare Pages? |
|---|---|---|
| Frontend | `taskpane/taskpane.html`, `.js`, `.css`, `config.js` | ✅ Ya (statis, HTTPS gratis) |
| Backend | `backend/server.py` (Python stdlib `http.server`) + PyMuPDF, python-docx, lxml, requests | ❌ Tidak |

Alasannya:

- **Cloudflare Pages** hanya menyajikan berkas statis (plus Pages Functions/Workers).
- **Cloudflare Workers tidak bisa menjalankan backend ini**: runtime Python Workers
  masih beta (Pyodide), batas memori 128 MB, dan paket seperti **PyMuPDF (fitz),
  lxml, python-docx** tidak tersedia. Batas CPU 10 ms/permintaan juga jauh di bawah
  kebutuhan generate dokumen.
- Karena itu pola yang benar: **frontend di Cloudflare Pages + backend Python di
  Render (gratis)**. Frontend memanggil API lintas-origin; CORS sudah diaktifkan
  di `server.py` (`Access-Control-Allow-Origin: *`).

```
Word  →  https://<proyek>.pages.dev/taskpane.html      (Cloudflare Pages, statis)
                 │  fetch https://<api>.onrender.com/api/...
                 ▼
        https://<api>.onrender.com                    (Render, backend Python)
```

---

## 1. Push kode ke GitHub

Repo sudah terhubung: `https://github.com/andraexcell-lang/autotesis-ai.git`.

```powershell
cd "E:\SKILL UNTUK AI\system_prompts_leaks-main\word-addin-tesismagister"
git add -A
git commit -m "feat(deploy): dukungan frontend Cloudflare Pages + backend Render"
git push origin main
```

Yang **tidak** ikut terkirim (sudah di `.gitignore`, jangan diubah):
`backend/llm_config.json`, `.env`, `_render_env_values.txt`, sertifikat `*.pfx/*.crt`,
`cloudflared.exe`, `output/`, `dist/`.

> Jangan pernah menaruh API key di file yang di-commit. Repo ini publik.

---

## 2. Backend ke Render (gratis)

1. Buka <https://dashboard.render.com> → **New** → **Blueprint**.
2. Pilih repo `autotesis-ai`. Render membaca `render.yaml` otomatis.
3. Saat diminta, isi Environment Variables:
   - `GEMINI_API_KEYS` → JSON array, mis. `["AIza...","AIza..."]` (atau dipisah koma)
   - `NVIDIA_API_KEY` → `nvapi-...` (opsional)
   - `LLM_MODELS`, `EXTRACT_MODELS` → opsional
4. Deploy. Tunggu build `pip install -r requirements.render.txt` (±3–5 menit).
5. Cek: `https://<nama-service>.onrender.com/api/healthz` → `200`.
6. Salin origin-nya, mis. `https://autotesis-api.onrender.com`.

**Batas gratis Render:** RAM 512 MB, tidur setelah ±15 menit idle, cold start
30–60 detik. Dampaknya: permintaan pertama setelah lama tidak dipakai terasa lambat
(tapi tidak gagal). Mitigasi:

- ping `/api/healthz` tiap 10 menit dari cron eksternal (cron-job.org, UptimeRobot), atau
- Render Starter $7/bln agar selalu hidup, atau
- pakai Cloudflare Tunnel ke PC sendiri (bagian 5).

Alternatif backend gratis lain: **Koyeb** (512 MB, scale-to-zero) atau **Hugging
Face Spaces** (Docker). `Dockerfile` dan `fly.toml` sudah tersedia bila mau pindah;
catatan: Fly.io kini tidak lagi punya tier gratis penuh.

---

## 3. Frontend ke Cloudflare Pages (gratis)

### Cara A – dashboard (tanpa token, paling mudah)

1. <https://dash.cloudflare.com> → **Workers & Pages** → **Create** → **Pages** →
   **Connect to Git** → pilih repo `autotesis-ai`.
2. Isi:
   - **Framework preset**: `None`
   - **Build command**: `node deploy/build_pages.mjs`
   - **Build output directory**: `dist`
   - **Environment variable**: `AUTOTESIS_BACKEND_URL` = `https://autotesis-api.onrender.com`
3. **Save and Deploy** → URL jadi `https://autotesis.pages.dev`.
4. Uji: `https://autotesis.pages.dev/taskpane.html` harus menampilkan panel, dan badge
   Pengaturan Sitasi terisi (artinya frontend berhasil memanggil backend).

### Cara B – dari komputer (wrangler)

```powershell
npm i -D wrangler
$env:AUTOTESIS_BACKEND_URL="https://autotesis-api.onrender.com"
node deploy/build_pages.mjs
npx wrangler pages deploy dist --project-name=autotesis
```

### Cara C – GitHub Actions

`.github/workflows/deploy-pages.yml` sudah disiapkan (trigger manual). Isi
**Settings → Secrets and variables → Actions**:

- `CLOUDFLARE_API_TOKEN` (izin *Cloudflare Pages: Edit*)
- `CLOUDFLARE_ACCOUNT_ID`
- variable `AUTOTESIS_BACKEND_URL`

lalu jalankan workflow **Deploy taskpane ke Cloudflare Pages**.

### Apa yang dilakukan `deploy/build_pages.mjs`

- menyalin hanya `taskpane.html`, `taskpane.css`, `taskpane.js`, `config.js`
  (file `.bak` tidak ikut terbit),
- menulis `config.js` berisi `window.AUTOTESIS_API_BASE = "<backend>"`,
- membuat `index.html` (redirect ke `taskpane.html`) dan `_headers` (nosniff, no-cache).

`taskpane.js` membaca `window.AUTOTESIS_API_BASE`; bila kosong, ia memakai origin
yang sama seperti sebelumnya (mode lokal/Render-saja tetap jalan).

---

## 4. Manifest Word Add-in

Manifest wajib memakai **https** dan URL absolut.

```powershell
python deploy/make_manifest.py --base-url https://autotesis.pages.dev
```

Perintah ini membuat `manifest_autotesis_pages_dev.xml` dengan semua URL (SourceLocation,
ikon, AppDomain, Taskpane.Url) menunjuk ke Pages.

Sideload:

1. Word → **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.
2. Tambahkan folder berisi file manifest, centang **Show in Menu**, restart Word.
3. **Insert → My Add-ins → Shared Folder → Independent Researcher**.

Untuk publikasi luas: upload manifest ke **Microsoft 365 admin center → Integrated apps**
(butuh akun tenant; tidak wajib untuk pemakaian sendiri).

> Catatan: ikon `assets/icon-16/32/64/80.png` belum ada di repo. Word tetap bisa
> memuat add-in, tetapi tombol ribbon memakai ikon default. Tambahkan PNG ukuran
> tersebut ke `taskpane/assets/` agar manifest sepenuhnya valid.

---

## 5. Alternatif: Cloudflare Tunnel (tanpa backend cloud)

Kalau ingin **backend tetap di PC sendiri** tetapi URL-nya permanen:

1. Punya domain di Cloudflare (boleh domain murah).
2. `cloudflared tunnel login` → `cloudflared tunnel create autotesis`.
3. Arahkan `autotesis.domainanda.com` ke `http://localhost:5050`.
4. Jalankan `cloudflared tunnel run autotesis` (atau sebagai service Windows).

Kelebihan: selalu hidup, tanpa cold start, data tetap lokal.
Kekurangan: PC harus menyala dan online.

Quick tunnel (`run_tunnel.py`) yang dipakai sekarang **tidak stabil** — URL
`*.trycloudflare.com` berubah setiap restart, sehingga `manifest_online.xml`
menjadi kedaluwarsa. Pakai named tunnel bila menginginkan URL tetap.

---

## 6. Ringkasan biaya

| Layanan | Paket | Biaya | Catatan |
|---|---|---|---|
| GitHub | Free | Rp 0 | repo publik/privat |
| Cloudflare Pages | Free | Rp 0 | 500 build/bulan, bandwidth tak terbatas |
| Cloudflare DNS/Tunnel | Free | Rp 0 | butuh domain sendiri untuk named tunnel |
| Render Web Service | Free | Rp 0 | tidur saat idle (cold start) |
| Render Starter | — | ~$7/bln | selalu hidup (opsional) |

---

## 7. Checklist verifikasi setelah deploy

- [ ] `https://<api>.onrender.com/api/healthz` → 200
- [ ] `https://<api>.onrender.com/api/s1/cite-settings` (POST `{}`) → JSON `status: success`
- [ ] `https://<proyek>.pages.dev/taskpane.html` tampil tanpa error console
- [ ] Badge Pengaturan Sitasi terisi (bukti CORS + API base benar)
- [ ] Ganti tahun/gaya sitasi → muat ulang halaman → nilai tersimpan
- [ ] Manifest menunjuk `pages.dev` (bukan `trycloudflare.com`)
- [ ] Add-in muncul di Word dan panel terbuka

Sudah diuji lokal: frontend statis di `:8080` + backend `:5050` (beda origin) →
`POST /api/s1/cite-settings` 200, badge terisi, 0 error console.
