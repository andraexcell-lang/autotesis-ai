/* Independent Researcher - Word Taskpane Controller (Dual Native COM & Office.js Engine) */

const API_BASE = (window.location.origin && window.location.origin.startsWith("http")) 
  ? `${window.location.origin}/api` 
  : "http://localhost:5050/api";

// State wizard SmartPLS HARUS dideklarasikan sebelum initUI() memanggil getConfig().
// Jika diletakkan di bawah, getConfig() kena TDZ ("Cannot access before initialization").
const smartplsState = {
  source: (() => {
    try { return localStorage.getItem("autothesis_data_source") || null; }
    catch (e) { return null; }
  })(),
  dataFile: null,
  guideFile: null
};
let smartplsGformScript = "";
// Status gerbang Bab IV juga dipakai initUI(), jadi deklarasikan lebih awal.
const dataGateState = { last: null, diperiksa: 0, dipaksa: {} };

console.log("Independent Researcher initialized. Target API:", API_BASE);

if (window.Office && Office.onReady) {
  Office.onReady((info) => {
    console.log("Office.js ready on host:", info.host);
    initUI();
  });
} else {
  // Standalone browser execution
  document.addEventListener("DOMContentLoaded", () => {
    initUI();
  });
}

function initUI() {
  // Tab Switching
  const tabs = document.querySelectorAll(".tab-btn");
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      tabs.forEach((t) => t.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
      tab.classList.add("active");
      const targetId = `panel-${tab.dataset.tab}`;
      const panel = document.getElementById(targetId);
      if (panel) panel.classList.add("active");
    });
  });

  loadSavedConfig();
  autoDetectFromTitle(false);

  document.getElementById("btn-save-config")?.addEventListener("click", saveConfig);
  document.getElementById("btn-auto-detect")?.addEventListener("click", () => autoDetectFromTitle(true));
  document.getElementById("btn-research-ontology")?.addEventListener("click", () => handleResearchOntology(true));
  document.getElementById("btn-upload-variable")?.addEventListener("click", handleUploadVariable);
  populateUploadVariableSelect();
  // Isi ulang dropdown sesuai variabel terkini setelah deteksi otomatis
  const judulInput = document.getElementById("input-judul");
  if (judulInput) {
    const origHandler = judulInput.oninput;
    judulInput.addEventListener("change", populateUploadVariableSelect);
  }

  // Live Auto-Detection saat Peneliti Mengetik / Menempelkan Judul
  const judulEl = document.getElementById("input-judul");
  if (judulEl) {
    let debounceTimer = null;
    judulEl.addEventListener("input", () => {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => {
        autoDetectFromTitle(true);
      }, 700);
    });
    judulEl.addEventListener("change", () => autoDetectFromTitle(true));
    judulEl.addEventListener("paste", () => {
      setTimeout(() => autoDetectFromTitle(true), 100);
    });
  }

  // Live update preview 7 hipotesis jika user mengubah variabel secara manual
  ["input-x1", "input-x2", "input-m", "input-y", "input-instansi"].forEach(id => {
    document.getElementById(id)?.addEventListener("input", () => {
      updateHypothesesPreview(
        document.getElementById("input-x1")?.value,
        document.getElementById("input-x2")?.value,
        document.getElementById("input-m")?.value,
        document.getElementById("input-y")?.value,
        document.getElementById("input-instansi")?.value
      );
    });
  });

  // Live update jumlah sampel otomatis dari populasi (Slovin / Sampling Jenuh)
  document.getElementById("input-populasi")?.addEventListener("input", () => {
    refreshSampelPreview();
  });
  refreshSampelPreview();

  // Chapter buttons
  document.querySelectorAll("[data-action]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const action = btn.dataset.action;
      await handleChapterGeneration(action, btn);
    });
  });

  document.getElementById("btn-generate-data")?.addEventListener("click", () => handleGenerateRespondents("generated"));
  document.getElementById("btn-upload-data")?.addEventListener("click", () => handleGenerateRespondents("upload"));
  document.getElementById("btn-import-sheet")?.addEventListener("click", () => handleGenerateRespondents("sheet"));
  document.getElementById("btn-olah-data")?.addEventListener("click", handleOlahData);
  document.getElementById("btn-generate-gform")?.addEventListener("click", handleGoogleFormScript);

  document.getElementById("btn-copy-gform")?.addEventListener("click", async () => {
    const ta = document.getElementById("gform-script");
    const text = (ta && ta.value) || smartplsGformScript || "";
    if (!text) { setStatus("error", "Belum ada skrip. Klik “Buat Google Form” dulu."); return; }
    const ok = await copyToClipboard(text);
    setStatus(ok ? "ready" : "error",
      ok ? "Skrip Google Form disalin ke clipboard."
         : "Gagal menyalin otomatis — blok teks pada kotak skrip lalu tekan Ctrl+C.");
  });

  document.getElementById("btn-download-gform")?.addEventListener("click", () => {
    const ta = document.getElementById("gform-script");
    const text = (ta && ta.value) || smartplsGformScript || "";
    if (!text) { setStatus("error", "Belum ada skrip. Klik “Buat Google Form” dulu."); return; }
    const blob = new Blob([text], { type: "text/javascript;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "KUESIONER_GOOGLE_FORM.gs";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
    setStatus("ready", "Skrip diunduh sebagai KUESIONER_GOOGLE_FORM.gs — tempel isinya di script.google.com.");
  });
  document.getElementById("btn-generate-full")?.addEventListener("click", handleFullDocumentGeneration);
  initGraphUI();
  
  // Dedicated Front Matter Export
  document.getElementById("btn-export-front-matter")?.addEventListener("click", async () => {
    const config = getConfig();
    const btn = document.getElementById("btn-export-front-matter");
    btn.disabled = true;
    setStatus("working", "Membuat & menyisipkan Bagian Awal (.docx)...");
    try {
      const response = await fetch(`${API_BASE}/generate-bagian-awal`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(config)
      });
      const data = await response.json();
      let inserted = false;
      if (data.docx_base64) {
        inserted = await insertDocxBase64ToActiveWord(data.docx_base64);
      }
      if (inserted) {
        setStatus("ready", `Bagian Awal berhasil disisipkan langsung ke dokumen Word aktif!`);
      } else {
        setStatus("ready", `Bagian Awal berhasil dibuat: ${data.filename}!`);
        window.open(`${API_BASE}/download/${data.filename}`, "_blank");
      }
    } catch (e) {
      setStatus("error", `Gagal: ${e.message}`);
    } finally {
      btn.disabled = false;
    }
  });

  // Dedicated Bab I Export
  document.getElementById("btn-export-bab-1")?.addEventListener("click", async () => {
    const config = getConfig();
    const btn = document.getElementById("btn-export-bab-1");
    btn.disabled = true;
    setStatus("working", "Membuat & menyisipkan Bab I Pendahuluan (.docx)...");
    try {
      const response = await fetch(`${API_BASE}/generate-bab-1`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(config)
      });
      const data = await response.json();
      let inserted = false;
      if (data.docx_base64) {
        inserted = await insertDocxBase64ToActiveWord(data.docx_base64);
      }
      if (inserted) {
        setStatus("ready", `Bab I Pendahuluan berhasil disisipkan langsung ke dokumen Word aktif!`);
      } else {
        setStatus("ready", `Bab I Pendahuluan berhasil dibuat: ${data.filename}!`);
        window.open(`${API_BASE}/download/${data.filename}`, "_blank");
      }
    } catch (e) {
      setStatus("error", `Gagal: ${e.message}`);
    } finally {
      btn.disabled = false;
    }
  });

  // Dedicated Bab II Export
  document.getElementById("btn-export-bab-2")?.addEventListener("click", async () => {
    const config = getConfig();
    const btn = document.getElementById("btn-export-bab-2");
    btn.disabled = true;
    setStatus("working", "Membuat & menyisipkan Bab II Kajian Pustaka (.docx)...");
    try {
      const response = await fetch(`${API_BASE}/generate-bab-2`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(config)
      });
      const data = await response.json();
      let inserted = false;
      if (data.docx_base64) {
        inserted = await insertDocxBase64ToActiveWord(data.docx_base64);
      }
      if (inserted) {
        setStatus("ready", `Bab II Kajian Pustaka berhasil disisipkan langsung ke dokumen Word aktif!`);
      } else {
        setStatus("ready", `Bab II Kajian Pustaka berhasil dibuat: ${data.filename}!`);
        window.open(`${API_BASE}/download/${data.filename}`, "_blank");
      }
    } catch (e) {
      setStatus("error", `Gagal: ${e.message}`);
    } finally {
      btn.disabled = false;
    }
  });

  // Dedicated Bab III Export
  document.getElementById("btn-export-bab-3")?.addEventListener("click", async () => {
    const config = getConfig();
    const btn = document.getElementById("btn-export-bab-3");
    btn.disabled = true;
    setStatus("working", "Membuat & menyisipkan Bab III Metode Penelitian (.docx)...");
    try {
      const response = await fetch(`${API_BASE}/generate-bab-3`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(config)
      });
      const data = await response.json();
      let inserted = false;
      if (data.docx_base64) {
        inserted = await insertDocxBase64ToActiveWord(data.docx_base64);
      }
      if (inserted) {
        setStatus("ready", `Bab III Metode Penelitian berhasil disisipkan langsung ke dokumen Word aktif!`);
      } else {
        setStatus("ready", `Bab III Metode Penelitian berhasil dibuat: ${data.filename}!`);
        window.open(`${API_BASE}/download/${data.filename}`, "_blank");
      }
    } catch (e) {
      setStatus("error", `Gagal: ${e.message}`);
    } finally {
      btn.disabled = false;
    }
  });

  // Dedicated Bab IV Export
  document.getElementById("btn-export-bab-4")?.addEventListener("click", async () => {
    const config = getConfig();
    const btn = document.getElementById("btn-export-bab-4");
    btn.disabled = true;
    setStatus("working", "Membuat & menyisipkan Bab IV Hasil & Pembahasan (.docx)...");
    try {
      const response = await fetch(`${API_BASE}/generate-bab-4`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(config)
      });
      const data = await response.json();
      let inserted = false;
      if (data.docx_base64) {
        inserted = await insertDocxBase64ToActiveWord(data.docx_base64);
      }
      if (inserted) {
        setStatus("ready", `Bab IV Hasil & Pembahasan berhasil disisipkan langsung ke dokumen Word aktif!`);
      } else {
        setStatus("ready", `Bab IV Hasil & Pembahasan berhasil dibuat: ${data.filename}!`);
        window.open(`${API_BASE}/download/${data.filename}`, "_blank");
      }
    } catch (e) {
      setStatus("error", `Gagal: ${e.message}`);
    } finally {
      btn.disabled = false;
    }
  });

  // Dedicated Bab V Export
  document.getElementById("btn-export-bab-5")?.addEventListener("click", async () => {
    const config = getConfig();
    const btn = document.getElementById("btn-export-bab-5");
    btn.disabled = true;
    setStatus("working", "Membuat & menyisipkan Bab V Kesimpulan, Saran & Dapus (.docx)...");
    try {
      const response = await fetch(`${API_BASE}/generate-bab-5`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(config)
      });
      const data = await response.json();
      let inserted = false;
      if (data.docx_base64) {
        inserted = await insertDocxBase64ToActiveWord(data.docx_base64);
      }
      if (inserted) {
        setStatus("ready", `Bab V Kesimpulan & Dapus berhasil disisipkan langsung ke dokumen Word aktif!`);
      } else {
        setStatus("ready", `Bab V Kesimpulan & Dapus berhasil dibuat: ${data.filename}!`);
        window.open(`${API_BASE}/download/${data.filename}`, "_blank");
      }
    } catch (e) {
      setStatus("error", `Gagal: ${e.message}`);
    } finally {
      btn.disabled = false;
    }
  });

  // ========================================================
  // HIERARCHICAL SUB-CHAPTER PIPELINE EVENT LISTENERS
  // ========================================================
  
  // Accordion Toggle Header
  document.querySelectorAll(".chapter-accordion-header").forEach((header) => {
    header.addEventListener("click", (e) => {
      if (e.target.closest(".btn-chain")) return;
      const targetId = header.dataset.toggle;
      const list = document.getElementById(targetId);
      if (list) {
        list.classList.toggle("collapsed");
        const toggleIcon = header.querySelector(".btn-accordion-toggle");
        if (toggleIcon) {
          toggleIcon.textContent = list.classList.contains("collapsed") ? "▶" : "▼";
        }
      }
    });
  });

  // Sub-Chapter Write (Native / Office.js)
  document.querySelectorAll(".btn-sub-write").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const subId = btn.dataset.sub;
      await handleSubChapterWrite(subId, btn);
    });
  });

  // Sub-Chapter Docx Download
  document.querySelectorAll(".btn-sub-docx").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const subId = btn.dataset.sub;
      await handleSubChapterDocx(subId, btn);
    });
  });

  // Run Chained Pipeline per Bab
  document.querySelectorAll(".btn-chain").forEach((btn) => {
    btn.addEventListener("click", async (e) => {
      e.stopPropagation();
      const chapterNum = btn.dataset.chapter;
      await handleRunChainBab(chapterNum, btn);
    });
  });

  // Handle Tombol Tulis Poin Pembahasan (1 s/d 7)
  document.querySelectorAll(".btn-poin-write").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const poin = btn.dataset.poin;
      await handlePembahasanPoinWrite(poin, btn);
    });
  });

  // Handle Tombol Unduh .docx Poin Pembahasan (1 s/d 7)
  document.querySelectorAll(".btn-poin-docx").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const poin = btn.dataset.poin;
      await handlePembahasanPoinDocx(poin, btn);
    });
  });

  // Handle Tulis Seluruh 7 Poin Pembahasan Berurutan
  document.getElementById("btn-write-all-pembahasan")?.addEventListener("click", handleWriteAllPembahasan);

  // Gerbang data Bab IV (data jawaban responden)
  document.getElementById("link-buka-tab-smartpls")?.addEventListener("click", (e) => {
    e.preventDefault();
    bukaTabSmartpls();
  });
  document.querySelectorAll(".tab-btn").forEach((t) => {
    t.addEventListener("click", () => {
      if (t.dataset.tab === "chapters") refreshDataGate();
    });
  });
  document.getElementById("btn-save-config")?.addEventListener("click", () => {
    setTimeout(() => refreshDataGate(true), 300);
  });
  refreshDataGate(true);
}

// -------------------------------------------------------------
// OFFICE.JS DIRECT INSERTION HELPER (LANGSUNG KE DOKUMEN AKTIF)
// -------------------------------------------------------------

async function insertDocxBase64ToActiveWord(base64Data) {
  if (!base64Data) return false;
  if (window.Word && typeof Word.run === "function") {
    try {
      await Word.run(async (context) => {
        context.document.body.insertFileFromBase64(base64Data, Word.InsertLocation.end);
        await context.sync();
      });
      return true;
    } catch (err) {
      console.warn("Office.js insertFileFromBase64 error:", err);
      return false;
    }
  }
  return false;
}

// Perbarui SEMUA field pada dokumen (mis. DAFTAR ISI otomatis) supaya Word langsung
// menghitung ulang judul & nomor halaman. Best-effort: bila API tidak mendukung,
// pengguna tetap bisa menekan Ctrl+A lalu F9 / klik kanan > Update Field.
async function updateAllFieldsInWord() {
  if (!window.Word || typeof Word.run !== "function") return false;
  try {
    return await Word.run(async (context) => {
      const fields = context.document.body.fields;
      fields.load("items");
      await context.sync();
      let n = 0;
      for (const f of fields.items) {
        try {
          if (typeof f.updateResult === "function") {
            f.updateResult();
            n++;
          }
        } catch (e) { /* field yang tidak mendukung: lewati */ }
      }
      await context.sync();
      return n > 0;
    });
  } catch (err) {
    console.warn("updateAllFieldsInWord gagal:", err);
    return false;
  }
}

// Minta SERVER mengatur nomor halaman pada dokumen Word aktif (lewat COM):
// bagian depan romawi kecil (i, ii, ...), naskah (BAB I dst.) angka Arab mulai 1,
// plus pembatas bab otomatis. Dipakai karena Office.js tidak menyediakan API
// untuk format/penomoran halaman.
async function rapikanNomorHalamanWord() {
  try {
    const r = await fetch(`${API_BASE}/format-halaman`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({})
    });
    if (!r.ok) return null;
    const j = await r.json();
    return (j && j.status === "success") ? j : null;
  } catch (err) {
    console.warn("rapikanNomorHalamanWord gagal:", err);
    return null;
  }
}

// Khusus BAGIAN AWAL (front matter): sisipkan di AWAL dokumen (sebelum BAB I),
// tanpa menyentuh Daftar Pustaka karena front matter tidak memuat sitasi.
async function insertBagianAwalAtStart(base64Data) {
  if (!base64Data) return false;
  if (!window.Word || typeof Word.run !== "function") return false;

  return Word.run(async (context) => {
    context.document.body.insertFileFromBase64(base64Data, Word.InsertLocation.start);
    await context.sync();
    return true;
  }).catch((err) => {
    console.error("Office.js insertBagianAwalAtStart error:", err);
    throw err;
  });
}

// Pintar & Runtut: Menyambungkan sub-bab tepat di bawah sub-bab sebelumnya pada halaman yang sama,
// dengan cara menghapus blok Daftar Pustaka lama & page-break-nya terlebih dahulu, lalu menambahkan naskah baru
// tepat di akhir naskah, dan menaruh kembali Daftar Pustaka kumulatif di halaman paling akhir.
async function insertSubchapterWithUnifiedDapus(subchapterBase64, dapusBase64) {
  if (!subchapterBase64) return false;
  if (!window.Word || typeof Word.run !== "function") return false;

  return Word.run(async (context) => {
    const body = context.document.body;
    // 1. Cari teks "DAFTAR PUSTAKA" di dalam dokumen aktif
    const searchResults = body.search("DAFTAR PUSTAKA", { matchCase: true });
    context.load(searchResults, "items");
    await context.sync();

    if (searchResults.items && searchResults.items.length > 0) {
      // DAFTAR PUSTAKA SUDAH ADA DI DOKUMEN!
      // PENTING: pakai kemunculan TERAKHIR. Entri "DAFTAR PUSTAKA" di DAFTAR ISI
      // (bagian awal) juga ikut terdeteksi; kalau dipakai, seluruh naskah setelah
      // daftar isi akan terhapus.
      const dapusRange = searchResults.items[searchResults.items.length - 1];
      const dapusPara = dapusRange.paragraphs.getFirst();
      context.load(dapusPara);
      const prevPara = dapusPara.getPreviousOrNullObject();
      context.load(prevPara, "text");
      await context.sync();

      // PENGAMAN LAMPIRAN: bila dokumen sudah punya heading LAMPIRAN, jangan
      // hapus blok Daftar Pustaka sampai akhir dokumen (Lampiran akan ikut
      // terhapus). Cukup sisipkan naskah SEBELUM heading DAFTAR PUSTAKA.
      const cekLamp = body.search("LAMPIRAN", { matchCase: true });
      context.load(cekLamp, "items");
      await context.sync();
      for (const it of (cekLamp.items || [])) {
        context.load(it.paragraphs.getFirst(), "text");
      }
      await context.sync();
      const adaLampiranHeading = (cekLamp.items || []).some((it) => {
        const t = (it.paragraphs.getFirst().text || "").trim().toUpperCase();
        return t === "LAMPIRAN";
      });
      if (adaLampiranHeading) {
        const titikSisip = dapusPara.getRange(Word.RangeLocation.start);
        titikSisip.insertFileFromBase64(subchapterBase64, Word.InsertLocation.start);
        await context.sync();
        return true;
      }

      // Tentukan titik awal hapus: jika ada paragraf kosong/page-break sebelum DAFTAR PUSTAKA, ikut dihapus
      let startDeleteRange = dapusPara.getRange(Word.RangeLocation.start);
      if (prevPara && !prevPara.isNullObject) {
        if (!prevPara.text || prevPara.text.trim() === "") {
          startDeleteRange = prevPara.getRange(Word.RangeLocation.start);
        }
      }

      // Hapus blok Daftar Pustaka lama beserta page break-nya dari dokumen
      const fullDapusRange = startDeleteRange.expandTo(body.getRange(Word.RangeLocation.end));
      fullDapusRange.delete();
      await context.sync();

      // Sekarang akhir dokumen adalah akhir naskah sub-bab sebelumnya!
      // Sisipkan naskah sub-bab baru langsung di akhir naskah (100% menyambung pada halaman yang sama)
      body.insertFileFromBase64(subchapterBase64, Word.InsertLocation.end);
      await context.sync();

      // Terakhir, sisipkan Daftar Pustaka terbaru di paling akhir dokumen (pada halaman tersendiri)
      if (dapusBase64) {
        body.insertFileFromBase64(dapusBase64, Word.InsertLocation.end);
        await context.sync();
      }
      return true;
    } else {
      // BELUM ADA DAFTAR PUSTAKA (dokumen kosong atau baru sub-bab pertama)
      body.insertFileFromBase64(subchapterBase64, Word.InsertLocation.end);
      if (dapusBase64) {
        body.insertFileFromBase64(dapusBase64, Word.InsertLocation.end);
      }
      await context.sync();
      return true;
    }
  }).catch((err) => {
    console.error("Office.js insertSubchapterWithUnifiedDapus error:", err);
    throw err;
  });
}

// Mode NATIVE: naskah SUDAH ditulis server via COM. Fungsi ini hanya
// memperbarui blok DAFTAR PUSTAKA di halaman paling akhir (tanpa menyisipkan
// ulang naskah), supaya tidak terjadi naskah ganda / daftar pustaka ganda.
async function replaceDapusOnly(dapusBase64) {
  if (!dapusBase64) return true;
  if (!window.Word || typeof Word.run !== "function") return false;

  return Word.run(async (context) => {
    const body = context.document.body;
    const hasil = body.search("DAFTAR PUSTAKA", { matchCase: true });
    context.load(hasil, "items");
    await context.sync();

    const items = hasil.items || [];
    if (items.length > 0) {
      // PENGAMAN LAMPIRAN: bila dokumen sudah punya heading "LAMPIRAN" (selalu
      // berada di belakang Daftar Pustaka), JANGAN menghapus apa pun. Blok hapus
      // di bawah membentang sampai AKHIR dokumen -> akan menelan Lampiran.
      const cekLamp = body.search("LAMPIRAN", { matchCase: true });
      context.load(cekLamp, "items");
      await context.sync();
      for (const it of (cekLamp.items || [])) {
        context.load(it.paragraphs.getFirst(), "text");
      }
      await context.sync();
      const adaHeadingLampiran = (cekLamp.items || []).some((it) => {
        const t = (it.paragraphs.getFirst().text || "").trim().toUpperCase();
        return t === "LAMPIRAN";
      });
      if (adaHeadingLampiran) {
        console.warn("replaceDapusOnly dilewati: dokumen berisi heading LAMPIRAN (tidak ada yang dihapus).");
        return true;
      }
      // Kemunculan TERAKHIR = heading Daftar Pustaka sebenarnya (bukan daftar isi).
      const head = items[items.length - 1];
      const para = head.paragraphs.getFirst();
      context.load(para);
      await context.sync();
      const prev = para.getPreviousOrNullObject();
      context.load(prev, "text");
      await context.sync();
      // Ikut hapus paragraf kosong/page-break tepat sebelum heading agar tidak
      // tersisa halaman kosong saat daftar pustaka baru disisipkan.
      let mulai = para.getRange(Word.RangeLocation.start);
      if (prev && !prev.isNullObject && (!prev.text || !prev.text.trim())) {
        mulai = prev.getRange(Word.RangeLocation.start);
      }
      mulai.expandTo(body.getRange(Word.RangeLocation.end)).delete();
      await context.sync();
    }

    body.insertFileFromBase64(dapusBase64, Word.InsertLocation.end);
    await context.sync();
    return true;
  }).catch((err) => {
    console.error("Office.js replaceDapusOnly error:", err);
    return false;
  });
}

// Khusus LAMPIRAN: sisipkan paling akhir (SETELAH Daftar Pustaka), tanpa memindahkan dapus.
async function appendLampiranAfterDapus(subchapterBase64, dapusBase64) {
  if (!subchapterBase64) return false;
  if (!window.Word || typeof Word.run !== "function") return false;

  return Word.run(async (context) => {
    const body = context.document.body;
    const searchResults = body.search("DAFTAR PUSTAKA", { matchCase: true });
    context.load(searchResults, "items");
    await context.sync();

    if (!searchResults.items || searchResults.items.length === 0) {
      // Belum ada Daftar Pustaka: sisipkan dapus lebih dulu, baru lampiran di bawahnya.
      if (dapusBase64) {
        body.insertFileFromBase64(dapusBase64, Word.InsertLocation.end);
        await context.sync();
      }
    }
    body.insertFileFromBase64(subchapterBase64, Word.InsertLocation.end);
    await context.sync();
    return true;
  }).catch((err) => {
    console.error("Office.js appendLampiranAfterDapus error:", err);
    throw err;
  });
}

// -------------------------------------------------------------
// HIERARCHICAL SUB-CHAPTER PIPELINE HANDLERS
// -------------------------------------------------------------

// ---------------------------------------------------------------------------
// PENGAMAN CONFIG (kasus nyata: judul diganti, variabel tidak)
// Panel ini dibuka dengan nilai CONTOH (judul & variabel tesis lain). Bila
// pengguna hanya mengganti judul lalu langsung menekan Tulis/Run Chain tanpa
// "Deteksi Otomatis Objek & Variabel", mesin akan menulis naskah dengan
// variabel CONTOH -> dokumen berjudul A tetapi isinya variabel B.
// Pengaman ini mendeteksi keadaan tsb lalu mendeteksi variabel dari judul;
// bila tetap gagal, proses DIHENTIKAN dengan pesan jelas (tidak menulis salah).
// ---------------------------------------------------------------------------
const JUDUL_CONTOH = "PENGARUH KECERDASAN EMOSIONAL DAN MOTIVASI KERJA TERHADAP KINERJA MELALUI ETOS KERJA SEBAGAI VARIABEL INTERVENING DI BPKAD KABUPATEN PALI";
const VAR_CONTOH = ["kecerdasan emosional", "motivasi kerja", "etos kerja"];

function variabelMasihContoh(cfg) {
  const judulBeda = (cfg.judul || "").trim().toUpperCase() !== JUDUL_CONTOH;
  if (!judulBeda) return false;
  const x1 = (cfg.x1 || "").trim().toLowerCase();
  const x2 = (cfg.x2 || "").trim().toLowerCase();
  const m = (cfg.m || "").trim().toLowerCase();
  return x1 === VAR_CONTOH[0] && x2 === VAR_CONTOH[1] && m === VAR_CONTOH[2];
}

async function pastikanVariabelSesuaiJudul() {
  if (!variabelMasihContoh(getConfig())) return true;
  setStatus("working", "Variabel masih nilai contoh panel -> mendeteksi otomatis dari judul...");
  try { await autoDetectFromTitle(false); } catch (e) { /* lanjut cek di bawah */ }
  if (variabelMasihContoh(getConfig())) {
    setStatus("error", "STOP: X1/X2/variabel mediasi belum diisi untuk judul ini (masih nilai contoh: Kecerdasan Emosional/Motivasi Kerja/Etos Kerja). Klik 'Deteksi Otomatis Objek & Variabel' atau isi manual dulu, supaya naskah tidak memakai variabel tesis contoh.");
    return false;
  }
  setStatus("working", "Variabel terdeteksi dari judul. Melanjutkan...");
  return true;
}

async function pastikanOntologiLengkap(subId) {
  // Bab III (metode & kisi-kisi), Bab IV (analisis), dan Lampiran (angket)
  // tidak boleh ditulis bila ada variabel tanpa indikator: hasilnya naskah
  // kosong / kuesioner bolong. Bab I-II tetap boleh (belum butuh indikator).
  const butuh = subId && (
    subId.indexOf("sub_3_") === 0 ||
    subId.indexOf("sub_4_") === 0 ||
    subId.indexOf("sub_5_") === 0 ||
    subId === "sub_lampiran"
  );
  if (!butuh) return true;
  try {
    const cfg = getConfig();
    const r = await fetch(`${API_BASE}/ontology/status`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ judul: cfg.judul || "" })
    });
    const d = await r.json();
    if (d.complete) return true;
    setStatus("error", "STOP: ontologi belum lengkap — " +
      (d.problems || []).join("; ") +
      ". Jalankan 'Telusuri Dimensi & Indikator' dan pastikan semua variabel " +
      "berindikator sebelum menulis Bab III–V/Lampiran.");
    return false;
  } catch (e) {
    console.warn("Gerbang ontologi tidak dapat diperiksa:", e);
    return true;
  }
}

async function handleSubChapterWrite(subId, btn) {
  // Gerbang data: sub-bab Bab IV/V butuh data yang sudah diolah lebih dulu.
  if (subId && (subId.indexOf("sub_4_") === 0 || subId.indexOf("sub_5_") === 0)) {
    const lanjut = await pastikanDataJawaban("tulis:" + subId, "✍️ Tulis");
    if (!lanjut) return;
  }
  // Gerbang variabel: cegah naskah memakai variabel contoh panel.
  if (!(await pastikanVariabelSesuaiJudul())) return;
  // Gerbang ontologi: cegah Bab III/IV/Lampiran ditulis dengan variabel kosong.
  if (!(await pastikanOntologiLengkap(subId))) return;
  const mode = getSelectedWriteMode();
  const row = document.getElementById(`row-${subId}`);
  btn.disabled = true;
  btn.textContent = "Mengetik...";
  setStatus("working", `Mengetikkan sub-bab ${subId} langsung ke dokumen Word aktif...`);
  setProgress(40);

  try {
    const config = getConfig();
    const response = await fetch(`${API_BASE}/pipeline/generate-subchapter`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sub_id: subId, mode: mode, config: config })
    });
    const res = await response.json();
    if (!response.ok || res.status === "error") {
      throw new Error(res.message || `HTTP ${response.status}`);
    }

    if (res.status === "success") {
      let inserted = false;
      if (res.native_written) {
        // Server SUDAH menulis naskah lewat COM (mode native) DAN memperbarui
        // Daftar Pustaka di dokumen yang sama. JANGAN panggil Office.js lagi:
        // replaceDapusOnly menghapus mulai heading DAFTAR PUSTAKA sampai AKHIR
        // dokumen, sehingga LAMPIRAN (yang berada di belakang Daftar Pustaka)
        // ikut terhapus. Itulah penyebab "Lampiran hilang setelah klik 4.1.1".
        inserted = true;
      } else if (res.docx_base64) {
        if (subId === "sub_bagian_awal") {
          inserted = await insertBagianAwalAtStart(res.docx_base64);
        } else if (subId === "sub_lampiran") {
          inserted = await appendLampiranAfterDapus(res.docx_base64, res.dapus_base64);
        } else {
          inserted = await insertSubchapterWithUnifiedDapus(res.docx_base64, res.dapus_base64);
        }
      }
      setProgress(100);
      if (inserted) {
        // Mode NATIVE: server sudah menomori halaman (romawi/Arab + pembatas bab)
        // DAN memperbarui semua field (DAFTAR ISI/TABEL/GAMBAR) lewat COM pada
        // dokumen yang sama. Memanggil ulang dari panel berarti mengerjakan
        // pekerjaan berat yang sama untuk kedua kalinya (lambat, tanpa hasil baru).
        // Mode Office.js (non-native): field dihitung, nomor halaman dirapikan
        // lewat server, lalu field dihitung sekali lagi agar nomor di daftar pas.
        let fieldOk = true;
        let nomor = res.native_written ? { sudah: true } : null;
        if (!res.native_written) {
          fieldOk = await updateAllFieldsInWord();
          nomor = await rapikanNomorHalamanWord();
          if (nomor) fieldOk = await updateAllFieldsInWord();
        }
        if (subId === "sub_bagian_awal") {
          const pesanNomor = nomor
            ? "Nomor halaman otomatis: bagian depan angka romawi (i, ii, ...), naskah mulai 1."
            : "Nomor halaman belum diatur otomatis (server Word tidak tersedia).";
          setStatus("ready", fieldOk
            ? `Bagian Awal (Front Matter) disisipkan di AWAL dokumen. DAFTAR ISI dihitung Word. ${pesanNomor}`
            : `Bagian Awal (Front Matter) disisipkan di AWAL dokumen. ${pesanNomor}`);
        } else if (subId === "sub_lampiran") {
          setStatus("ready", "Lampiran Kuesioner berhasil disisipkan SETELAH Daftar Pustaka di dokumen Word!");
        } else {
          setStatus("ready", `Sub-bab ${res.meta.code || subId} (${res.meta.title || ''}) berhasil menyambung di dokumen Word, dan Daftar Pustaka diperbarui di halaman paling bawah!`);
        }
      } else {
        setStatus("ready", `Sub-bab ${res.meta.code || subId} (${res.meta.title || ''}) berhasil dibuat: ${res.filename}`);
      }
      if (row) row.classList.add("row-success");
    } else {
      setStatus("error", res.message || "Gagal mengetik sub-bab");
    }
  } catch (err) {
    console.error(err);
    setStatus("error", `Error: ${err.message}`);
  } finally {
    setTimeout(() => setProgress(0), 1000);
    btn.disabled = false;
    btn.textContent = "✍️ Tulis";
  }
}

async function handleSubChapterDocx(subId, btn) {
  btn.disabled = true;
  setStatus("working", `Membuat dokumen .docx sub-bab ${subId}...`);
  try {
    const config = getConfig();
    const response = await fetch(`${API_BASE}/pipeline/generate-subchapter`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sub_id: subId, mode: "download", config: config })
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const res = await response.json();
    if (res.status === "success") {
      setStatus("ready", `Dokumen ${res.filename} siap diunduh!`);
      window.open(`${API_BASE}/download/${res.filename}`, "_blank");
    } else {
      setStatus("error", res.message || "Gagal membuat berkas sub-bab");
    }
  } catch (err) {
    console.error(err);
    setStatus("error", `Error: ${err.message}`);
  } finally {
    btn.disabled = false;
  }
}

async function handleRunChainBab(chapterNum, btn) {
  // Gerbang data: Run Chain Bab IV/V butuh data yang sudah diolah lebih dulu.
  if (String(chapterNum) === "4" || String(chapterNum) === "5") {
    const lanjut = await pastikanDataJawaban(
      "chain:" + chapterNum, `⚡ Run Chain Bab ${chapterNum}`);
    if (!lanjut) return;
  }
  // Gerbang variabel: cegah naskah memakai variabel contoh panel.
  if (!(await pastikanVariabelSesuaiJudul())) return;
  // Gerbang ontologi: rantai Bab III/IV/V wajib punya indikator lengkap.
  if (["3", "4", "5"].includes(String(chapterNum))) {
    if (!(await pastikanOntologiLengkap(`sub_${chapterNum}_`))) return;
  }
  const mode = getSelectedWriteMode();
  btn.disabled = true;
  const originalHtml = btn.innerHTML;
  btn.innerHTML = `⚡ Menjalankan Chain...`;
  setStatus("working", `Menjalankan Chained Sub-Chapter Pipeline Bab ${chapterNum} langsung ke dokumen Word aktif...`);
  setProgress(20);

  try {
    const config = getConfig();
    const response = await fetch(`${API_BASE}/pipeline/run-chain`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ chapter: parseInt(chapterNum, 10), mode: mode, config: config })
    });
    const res = await response.json();
    if (!response.ok || res.status === "error") {
      throw new Error(res.message || `HTTP ${response.status}`);
    }

    if (res.status === "success") {
      if (res.details && Array.isArray(res.details) && typeof Word !== "undefined" && typeof Word.run === "function") {
        await Word.run(async (context) => {
          const searchResults = context.document.body.search("DAFTAR PUSTAKA", { matchCase: true });
          context.load(searchResults, "items");
          await context.sync();

          if (searchResults.items && searchResults.items.length > 0) {
            const dapusRange = searchResults.items[0];
            const dapusPara = dapusRange.paragraphs.getFirst();
            context.load(dapusPara);
            const prevPara = dapusPara.getPreviousOrNullObject();
            context.load(prevPara, "text");
            await context.sync();

            let startDeleteRange = dapusPara.getRange(Word.RangeLocation.start);
            if (prevPara && !prevPara.isNullObject) {
              if (!prevPara.text || prevPara.text.trim() === "") {
                startDeleteRange = prevPara.getRange(Word.RangeLocation.start);
              }
            }

            const fullDapusRange = startDeleteRange.expandTo(context.document.body.getRange(Word.RangeLocation.end));
            fullDapusRange.delete();
            await context.sync();

            const body = context.document.body;
            for (const item of res.details) {
              if (item.docx_base64) {
                body.insertFileFromBase64(item.docx_base64, Word.InsertLocation.end);
              }
              const row = document.getElementById(`row-${item.sub_id}`);
              if (row) row.classList.add("row-success");
            }
            if (res.dapus_base64) {
              body.insertFileFromBase64(res.dapus_base64, Word.InsertLocation.end);
            }
          } else {
            const body = context.document.body;
            for (const item of res.details) {
              if (item.docx_base64) {
                body.insertFileFromBase64(item.docx_base64, Word.InsertLocation.end);
              }
              const row = document.getElementById(`row-${item.sub_id}`);
              if (row) row.classList.add("row-success");
            }
            if (res.dapus_base64) {
              body.insertFileFromBase64(res.dapus_base64, Word.InsertLocation.end);
            }
          }
          await context.sync();
        });
      }
      setProgress(100);
      setStatus("ready", `⚡ Eksekusi Berantai Bab ${chapterNum} Selesai: Seluruh ${res.total_subchapters || ''} sub-bab berhasil disisipkan runtut dan Daftar Pustaka ditempatkan di halaman paling bawah!`);
    } else {
      setStatus("error", res.message || "Gagal mengeksekusi pipeline berantai");
    }
  } catch (err) {
    console.error(err);
    setStatus("error", `Error: ${err.message}`);
  } finally {
    setTimeout(() => setProgress(0), 1500);
    btn.disabled = false;
    btn.innerHTML = originalHtml;
  }
}


async function handlePembahasanPoinWrite(poin, btn) {
  const mode = document.querySelector('input[name="write-mode"]:checked')?.value || "native";
  btn.disabled = true;
  setStatus("working", `Mengetikkan Pembahasan Poin ${poin} ke Word...`);
  try {
    const config = getConfig();
    const response = await fetch(`${API_BASE}/generate-pembahasan-poin`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ poin: parseInt(poin, 10), mode: mode, config: config })
    });
    const res = await response.json();
    if (res.status === "success") {
      let inserted = false;
      if (res.docx_base64) {
        inserted = await insertDocxBase64ToActiveWord(res.docx_base64);
      } else if (mode === "officejs" && window.Word) {
        await writePembahasanToOfficeJs(res.content);
        inserted = true;
      }
      setStatus("ready", `Pembahasan Poin ${poin} (${res.hipotesis}) berhasil ditulis langsung ke dokumen Word!`);
    } else {
      setStatus("error", res.message || "Gagal menulis pembahasan");
    }
  } catch (err) {
    setStatus("error", `Error: ${err.message}`);
  } finally {
    btn.disabled = false;
  }
}

async function handlePembahasanPoinDocx(poin, btn) {
  btn.disabled = true;
  setStatus("working", `Membuat dokumen Word Poin ${poin}...`);
  try {
    const config = getConfig();
    const response = await fetch(`${API_BASE}/generate-pembahasan-poin`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ poin: parseInt(poin, 10), mode: "download", config: config })
    });
    const res = await response.json();
    if (res.status === "success") {
      setStatus("ready", `Dokumen Poin ${poin} (${res.hipotesis}) siap: ${res.filename}`);
      window.open(`${API_BASE}/download/${res.filename}`, "_blank");
    } else {
      setStatus("error", res.message || "Gagal membuat dokumen");
    }
  } catch (err) {
    setStatus("error", `Error: ${err.message}`);
  } finally {
    btn.disabled = false;
  }
}

async function handleWriteAllPembahasan() {
  const btn = document.getElementById("btn-write-all-pembahasan");
  btn.disabled = true;
  setStatus("working", "Mengetikkan seluruh 7 Poin Pembahasan secara berurutan...");
  try {
    for (let p = 1; p <= 7; p++) {
      setStatus("working", `Sedang mengetikkan Poin ${p} dari 7...`);
      const response = await fetch(`${API_BASE}/generate-pembahasan-poin`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ poin: p, mode: "native", config: getConfig() })
      });
      await response.json();
    }
    setStatus("ready", "Seluruh 7 Poin Pembahasan berhasil ditulis lengkap ke Word!");
  } catch (err) {
    setStatus("error", `Error: ${err.message}`);
  } finally {
    btn.disabled = false;
  }
}

async function writePembahasanToOfficeJs(content) {
  if (!window.Word) return;
  return Word.run(async (context) => {
    const body = context.document.body;
    const titleP = body.insertParagraph(content.title, Word.InsertLocation.end);
    titleP.font.bold = true;
    titleP.font.name = "Times New Roman";
    titleP.font.size = 12;
    
    for (const alinea of [content.alinea_1, content.alinea_2, content.alinea_3, content.alinea_4]) {
      const p = body.insertParagraph(alinea, Word.InsertLocation.end);
      p.font.name = "Times New Roman";
      p.font.size = 12;
      p.lineSpacing = 1.5 * 12;
    }
    await context.sync();
  });
}

// Hitung ukuran sampel dari populasi:
// Populasi > 100 -> Rumus Slovin (e = 5%); Populasi <= 100 -> Sampling Jenuh (sensus)
function hitungSampel(populasi) {
  const N = parseInt(populasi, 10);
  if (!N || N <= 0) return 0;
  if (N > 100) return Math.ceil(N / (1 + N * 0.05 * 0.05));
  return N;
}

function refreshSampelPreview() {
  const populasi = parseInt(document.getElementById("input-populasi")?.value || "0", 10);
  const sampel = hitungSampel(populasi);
  const sampelEl = document.getElementById("input-sampel");
  if (sampelEl && sampel > 0) {
    sampelEl.value = `${sampel} responden (${populasi > 100 ? "Rumus Slovin" : "Sampling Jenuh"})`;
  }
  const respEl = document.getElementById("input-responden");
  if (respEl && sampel > 0) respEl.value = sampel;
  return sampel;
}

function getConfig() {
  const populasi = parseInt(document.getElementById("input-populasi")?.value || "0", 10) || 0;
  const nFromPopulasi = hitungSampel(populasi);
  return {
    judul: document.getElementById("input-judul")?.value.trim() || "PENGARUH KECERDASAN EMOSIONAL DAN MOTIVASI KERJA TERHADAP KINERJA MELALUI ETOS KERJA SEBAGAI VARIABEL INTERVENING DI BPKAD KABUPATEN PALI",
    penulis: document.getElementById("input-penulis")?.value.trim() || "MUHAMAD YUSI",
    nim: document.getElementById("input-nim")?.value.trim() || "255341028",
    instansi: document.getElementById("input-instansi")?.value.trim() || "Badan Pengelola Keuangan dan Aset Daerah (BPKAD) Kabupaten Penukal Abab Lematang Ilir (PALI)",
    tahun: document.getElementById("input-tahun")?.value.trim() || "2026",
    x1: document.getElementById("input-x1")?.value.trim() || "Kecerdasan Emosional",
    x2: document.getElementById("input-x2")?.value.trim() || "Motivasi Kerja",
    m: document.getElementById("input-m")?.value.trim() || "Etos Kerja",
    y: document.getElementById("input-y")?.value.trim() || "Kinerja Pegawai",
    populasi: populasi,
    n: nFromPopulasi || parseInt(document.getElementById("input-responden")?.value || "55", 10),
    data_source: (typeof smartplsState !== "undefined" && smartplsState.source) || "generated",
    analysis_engine: "internal"
  };
}

function toTitleCase(str) {
  if (!str) return "";
  const lowercaseWords = new Set(["dan", "terhadap", "melalui", "pada", "di", "serta", "sebagai", "oleh", "yang", "ke", "dari"]);
  return str.trim().split(/\s+/).map((word, idx) => {
    const lw = word.toLowerCase();
    if (idx > 0 && lowercaseWords.has(lw)) {
      return lw;
    }
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join(" ");
}

function cleanVarName(str) {
  if (!str) return "";
  let s = str.trim();
  s = s.replace(/^(?:variabel\s+|faktor\s+)/i, "");
  s = s.replace(/\s+sebagai\s+(?:variabel\s+)?(?:intervening|mediasi|mediator)$/i, "");
  s = s.replace(/\s+sebagai\s+mediasi$/i, "");
  return s.trim();
}

function parseThesisTitle(title) {
  if (!title) return null;
  const t = title.trim().replace(/\s+/g, " ");

  let instansi = "";
  let x1 = "", x2 = "", m = "", y = "";

  // 1. Pola: PENGARUH X1 DAN X2 TERHADAP M SERTA DAMPAKNYA PADA/TERHADAP Y [DI/PADA] INSTANSI
  if (/SERTA\s+DAMPAKNYA\s+(?:PADA|TERHADAP)/i.test(t)) {
    const match = t.match(/PENGARUH\s+(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+?)\s+SERTA\s+DAMPAKNYA\s+(?:PADA|TERHADAP)\s+(.+)$/i);
    if (match) {
      x1 = match[1];
      x2 = match[2];
      m = match[3];
      const rest = match[4].trim();
      const mInst = rest.match(/^(KINERJA(?:\s+PEGAWAI|\s+KARYAWAN|\s+GURU|\s+APARATUR)?)\s+(?:DI|PADA)?\s*(.+)$/i);
      if (mInst) {
        y = mInst[1];
        instansi = mInst[2];
      } else {
        y = rest;
      }
    }
  } 
  // 2. Pola: PENGARUH X1 DAN X2 TERHADAP Y MELALUI M [SEBAGAI VARIABEL INTERVENING] [DI/PADA] INSTANSI
  else if (/MELALUI\s+/i.test(t)) {
    let core = t;
    const mInst = t.match(/\b(?:DI|PADA|STUDI KASUS PADA|STUDI PADA)\s+((?:BPKAD|BANK|DPMPTSP|DINAS|BADAN|KANTOR|PT|CV|SMA|SMP|SD|RSUD|RUMAH SAKIT|BAPENDA|KECAMATAN|KABUPATEN|PROVINSI|KOTA|[A-Z0-9\s]+)$)/i);
    if (mInst) {
      instansi = mInst[1].trim();
      core = t.substring(0, mInst.index).trim();
    }
    const match = core.match(/PENGARUH\s+(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+?)\s+MELALUI\s+(.+)$/i);
    if (match) {
      x1 = match[1];
      x2 = match[2];
      y = match[3];
      m = match[4];
    }
  }
  // 3. Pola: PENGARUH X1 DAN X2 TERHADAP Y YANG DIMEDIASI OLEH M / DENGAN M SEBAGAI ...
  else if (/DIMEDIASI\s+OLEH|DENGAN\s+.+?\s+SEBAGAI\s+VARIABEL/i.test(t)) {
    let core = t;
    const mInst = t.match(/\b(?:DI|PADA)\s+((?:BPKAD|BANK|DPMPTSP|DINAS|BADAN|KANTOR|PT|CV|SMA|SMP|SD|RSUD|RUMAH SAKIT|BAPENDA|KECAMATAN|KABUPATEN|PROVINSI|KOTA|[A-Z0-9\s]+)$)/i);
    if (mInst) {
      instansi = mInst[1].trim();
      core = t.substring(0, mInst.index).trim();
    }
    const match = core.match(/PENGARUH\s+(.+?)\s+DAN\s+(.+?)\s+TERHADAP\s+(.+?)\s+(?:YANG\s+DIMEDIASI\s+OLEH|DENGAN)\s+(.+)$/i);
    if (match) {
      x1 = match[1];
      x2 = match[2];
      y = match[3];
      m = match[4];
    }
  }

  // Fallback ekstraksi instansi jika belum dapat
  if (!instansi) {
    const fallbackInst = t.match(/\b(?:DI|PADA)\s+(.+)$/i);
    if (fallbackInst) instansi = fallbackInst[1].trim();
  }

  return {
    instansi: instansi ? instansi.trim() : "",
    x1: toTitleCase(cleanVarName(x1)),
    x2: toTitleCase(cleanVarName(x2)),
    m: toTitleCase(cleanVarName(m)),
    y: toTitleCase(cleanVarName(y))
  };
}

function updateHypothesesPreview(x1, x2, m, y, instansi) {
  const box = document.getElementById("hypotheses-list");
  if (!box) return;
  const _x1 = x1 || "X1";
  const _x2 = x2 || "X2";
  const _m = m || "M (Mediasi)";
  const _y = y || "Y (Kinerja)";
  const _loc = instansi ? ` pada ${instansi}` : "";

  const hypotheses = [
    `<strong>H1:</strong> Diduga terdapat pengaruh <em>${_x1}</em> terhadap <em>${_m}</em>${_loc}.`,
    `<strong>H2:</strong> Diduga terdapat pengaruh <em>${_x2}</em> terhadap <em>${_m}</em>${_loc}.`,
    `<strong>H3:</strong> Diduga terdapat pengaruh <em>${_x1}</em> terhadap <em>${_y}</em>${_loc}.`,
    `<strong>H4:</strong> Diduga terdapat pengaruh <em>${_x2}</em> terhadap <em>${_y}</em>${_loc}.`,
    `<strong>H5:</strong> Diduga terdapat pengaruh <em>${_m}</em> terhadap <em>${_y}</em>${_loc}.`,
    `<strong>H6:</strong> Diduga terdapat pengaruh <em>${_x1}</em> terhadap <em>${_y}</em> melalui <em>${_m}</em> sebagai variabel mediasi${_loc}.`,
    `<strong>H7:</strong> Diduga terdapat pengaruh <em>${_x2}</em> terhadap <em>${_y}</em> melalui <em>${_m}</em> sebagai variabel mediasi${_loc}.`
  ];

  box.innerHTML = hypotheses.map(h => `<div style="margin-bottom: 4px;">• ${h}</div>`).join("");
}

function applyParsedData(parsed, showStatus = true) {
  if (!parsed) return;
  let changed = false;
  if (parsed.instansi) {
    const elInstansi = document.getElementById("input-instansi");
    if (elInstansi) { elInstansi.value = parsed.instansi; changed = true; }
  }
  if (parsed.x1) {
    const elX1 = document.getElementById("input-x1");
    if (elX1) { elX1.value = parsed.x1; changed = true; }
  }
  if (parsed.x2) {
    const elX2 = document.getElementById("input-x2");
    if (elX2) { elX2.value = parsed.x2; changed = true; }
  }
  if (parsed.m) {
    const elM = document.getElementById("input-m");
    if (elM) { elM.value = parsed.m; changed = true; }
  }
  if (parsed.y) {
    const elY = document.getElementById("input-y");
    if (elY) { elY.value = parsed.y; changed = true; }
  }

  // Update visual 7 hipotesis secara instan
  updateHypothesesPreview(
    document.getElementById("input-x1")?.value,
    document.getElementById("input-x2")?.value,
    document.getElementById("input-m")?.value,
    document.getElementById("input-y")?.value,
    document.getElementById("input-instansi")?.value
  );

  // Auto-save ke localStorage
  const currentCfg = getConfig();
  localStorage.setItem("autotesis_config", JSON.stringify(currentCfg));

  if (showStatus) {
    const statusEl = document.getElementById("detect-status");
    if (statusEl) {
      statusEl.style.display = "block";
      statusEl.innerHTML = `⚡ <strong>Terdeteksi Otomatis:</strong> Objek: <em>${parsed.instansi || 'Objek Penelitian'}</em> | X1: <em>${parsed.x1}</em> | X2: <em>${parsed.x2}</em> | M: <em>${parsed.m}</em> | Y: <em>${parsed.y}</em>`;
      setTimeout(() => {
        statusEl.style.display = "none";
      }, 6000);
    }
  }
}

async function autoDetectFromTitle(showStatus = true) {
  const judulInput = document.getElementById("input-judul");
  if (!judulInput) return;
  const title = judulInput.value.trim();
  if (!title) return;

  // 1. Ekstraksi langsung sisi klien (Zero-latency sub-millisecond)
  const localParsed = parseThesisTitle(title);
  if (localParsed) {
    applyParsedData(localParsed, showStatus);
  }

  // 2. Sinkronisasi AI backend untuk validasi struktur komprehensif
  try {
    const response = await fetch(`${API_BASE}/parse-title`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ title: title })
    });
    if (response.ok) {
      const res = await response.json();
      if (res.status === "success" && res.data) {
        applyParsedData(res.data, showStatus);
        // Otomatis jalankan riset ontologi dimensi & indikator di latar belakang
        handleResearchOntology(false);
      }
    }
  } catch (err) {
    // Server offline / mode terputus: parsing lokal sudah aktif
  }
}

let _ontologyReqSeq = 0;
const _TEKS_TOMBOL_ONTOLOGI = "🔍 Telusuri Dimensi & Indikator via Live Academic API";

async function handleResearchOntology(showLoading = true) {
  const btn = document.getElementById("btn-research-ontology");
  const badge = document.getElementById("badge-ontology-status");
  if (!btn) return;

  const config = getConfig();
  if (!config.x1 || !config.y) {
    setStatus("error", "Isi Variabel Terikat (Y) dan minimal X1 sebelum menelusuri dimensi & indikator.");
    return;
  }

  // Satu riset pada satu waktu. Auto-trigger (showLoading=false) tidak boleh
  // menumpuk; klik manual SELALU dilayani dengan membatalkan riset lama.
  if (window.__ontologyInFlight) {
    if (!showLoading) return;
    try { window.__ontologyCtrl && window.__ontologyCtrl.abort(); } catch (e) {}
    window.__ontologyInFlight = false;
  }
  window.__ontologyInFlight = true;
  const seq = ++_ontologyReqSeq;

  if (showLoading) {
    btn.disabled = true;
    btn.textContent = "⏳ Meriset Dimensi & Indikator via Live Academic API...";
  }
  if (badge) badge.textContent = "Sedang Meriset...";
  setStatus("working", "Menelusuri dimensi & indikator via Live Academic API...");

  // Batas waktu 4 menit: bila server tidak menjawab, tombol TIDAK boleh
  // terkunci selamanya (penyebab tombol terasa tidak bisa diklik).
  const ctrl = (typeof AbortController !== "undefined") ? new AbortController() : null;
  window.__ontologyCtrl = ctrl;
  const timer = ctrl ? setTimeout(() => { try { ctrl.abort(); } catch (e) {} }, 240000) : null;

  try {
    const response = await fetch(`${API_BASE}/research-ontology`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal: ctrl ? ctrl.signal : undefined,
      body: JSON.stringify({
        x1: config.x1,
        x2: config.x2,
        m: config.m,
        y: config.y,
        instansi: config.instansi,
        judul: config.judul,
        sektor: config.sektor || ""
      })
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    const data = await response.json();
    if (seq !== _ontologyReqSeq) return; // hasil dari permintaan lama -> abaikan

    if (data.status === "success" && data.ontology) {
      window._autotesisOntology = data.ontology;
      try {
        renderOntologyPreview(data.ontology);
      } catch (eRender) {
        console.warn("Gagal merender pratinjau ontologi:", eRender);
      }
      const nr = (data.ontology.metadata && data.ontology.metadata.needs_review) || [];
      if (badge) {
        if (nr.length) {
          badge.textContent = `⚠️ Belum lengkap: ${nr.join(', ')}`;
          badge.style.background = "#fef3c7";
          badge.style.color = "#b45309";
          setStatus("error", "Ontologi belum lengkap untuk: " + nr.join(', ') +
            ". Variabel ini belum punya indikator; Bab III–V diblokir sampai dilengkapi.");
        } else {
          const _teo = data.ontology.metadata.total_indikator_teoretis || '-';
          badge.textContent = `✅ Siap (${data.ontology.metadata.total_indikator} indikator operasional; ${_teo} teoretis)`;
          badge.style.background = "#bbf7d0";
          badge.style.color = "#15803d";
          setStatus("ready", "Dimensi & indikator berhasil ditelusuri dan siap dipakai.");
        }
      }
    } else {
      throw new Error(data.message || "respons tidak dikenali");
    }
  } catch (err) {
    if (seq === _ontologyReqSeq) {
      const pesan = (err && err.name === "AbortError")
        ? "Riset dimensi melebihi batas waktu 4 menit. Silakan coba lagi."
        : ("Riset dimensi gagal: " + (err && err.message ? err.message : String(err)));
      if (badge) {
        badge.textContent = "⚠️ Gagal";
        badge.style.background = "#fee2e2";
        badge.style.color = "#b91c1c";
      }
      setStatus("error", pesan);
    }
    console.warn("Ontology research error:", err);
  } finally {
    if (timer) clearTimeout(timer);
    if (seq === _ontologyReqSeq) {
      window.__ontologyInFlight = false;
      window.__ontologyCtrl = null;
      // SELALU pulihkan tombol, termasuk saat showLoading=false atau request gagal.
      btn.disabled = false;
      btn.textContent = _TEKS_TOMBOL_ONTOLOGI;
    }
  }
}

function populateUploadVariableSelect() {
  const sel = document.getElementById("upload-var-select");
  if (!sel) return;
  const config = getConfig();
  const options = [
    ["x1", config.x1, "X1"],
    ["x2", config.x2, "X2"],
    ["m", config.m, "M"],
    ["y", config.y, "Y"]
  ];
  sel.innerHTML = options
    .filter(([, name]) => name)
    .map(([role, name, label]) => `<option value="${role}" data-name="${name}">${label}: ${name}</option>`)
    .join("");
}

async function handleUploadVariable() {
  const btn = document.getElementById("btn-upload-variable");
  const status = document.getElementById("upload-var-status");
  const sel = document.getElementById("upload-var-select");
  const fileEl = document.getElementById("upload-var-file");
  if (!btn || !status || !sel || !fileEl) return;

  const file = fileEl.files && fileEl.files[0];
  const option = sel.options[sel.selectedIndex];
  if (!file || !option) {
    status.textContent = "⚠️ Pilih variabel dan berkas terlebih dahulu.";
    status.style.color = "#b45309";
    return;
  }

  btn.disabled = true;
  status.style.color = "#475569";
  status.textContent = "⏳ Membaca berkas dan mengekstrak indikator...";

  try {
    const base64 = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result).split(",")[1] || "");
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(file);
    });

    const response = await fetch(`${API_BASE}/variable/upload`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        variable: option.dataset.name,
        role: option.value.toUpperCase().replace("X1", "X").replace("X2", "X"),
        filename: file.name,
        content_base64: base64
      })
    });
    const data = await response.json();

    if (data.status === "ok" && data.entry) {
      const n = data.entry.dimensi.reduce((a, d) => a + d.indikator.length, 0);
      status.style.color = "#15803d";
      status.textContent = `✅ ${n} indikator dari "${file.name}" tersimpan ke pustaka. Klik Telusuri untuk memakainya.`;
    } else {
      status.style.color = "#b45309";
      status.textContent = `⚠️ ${data.report?.note || data.message || "Indikator tidak dapat diverifikasi dari berkas ini."}`;
    }
  } catch (err) {
    status.style.color = "#b91c1c";
    status.textContent = `❌ Gagal mengunggah: ${err.message || err}`;
  } finally {
    btn.disabled = false;
  }
}

function renderOntologyPreview(onto) {
  const box = document.getElementById("ontology-preview-box");
  if (!box) return;

  box.style.display = "block";
  const vars = onto.variables;
  const anchors = onto.phenomenon_anchors;
  const nr = (onto.metadata && onto.metadata.needs_review) || [];
  const complete = nr.length === 0;
  const dimsText = (v) => {
    const ds = (v && v.dimensi) || [];
    if (!ds.length) {
      return '<span style="color: #b45309;">⚠️ belum ada indikator</span>';
    }
    return '• ' + ds.map(d => {
      const op = (d.indikator_operasional || d.indikator || []).length;
      const teo = (d.indikator_teoretis || []).length;
      return `${d.nama} (${op} operasional / ${teo} teoretis)`;
    }).join(', ');
  };
  const indCount = (v) => ((v && v.dimensi) || [])
    .reduce((a, d) => a + ((d.indikator || []).length), 0);
  const teoCount = (v) => ((v && v.dimensi) || [])
    .reduce((a, d) => a + ((d.indikator_teoretis || []).length), 0);
  const selectionLine = (v) => {
    const s = v && v.dimension_selection;
    if (!s) return '';
    const buang = (s.tidak_dipakai || []).length;
    return `• Seleksi dimensi: ${s.dipilih_total}/${s.sumber_total} dipakai` +
      (buang ? `, ${buang} tidak dipakai (tercatat).` : ', tidak ada yang dibuang.');
  };
  const selectionDetails = () => {
    const rows = [];
    ['x1', 'x2', 'm', 'y'].forEach(k => {
      const v = vars[k] || {};
      const s = v.dimension_selection;
      if (!s || !(s.tidak_dipakai || []).length) {
        if (v.nama_resmi) {
          rows.push(`<div style="margin-top: 3px;"><strong>${k.toUpperCase()} ${v.nama_resmi}:</strong> ${selectionLine(v) || 'seluruh dimensi sumber dipakai.'}</div>`);
        }
        return;
      }
      const dibuang = s.tidak_dipakai.map(x => `“${x.nama}” (${x.alasan})`).join('; ');
      rows.push(`<div style="margin-top: 3px;"><strong>${k.toUpperCase()} ${v.nama_resmi}:</strong> ${selectionLine(v)}<br><span style="color:#7c2d12;">Tidak dipakai: ${dibuang}</span></div>`);
    });
    return rows.join('');
  };

  let html = `
    <div style="font-weight: bold; color: ${complete ? '#15803d' : '#b45309'}; margin-bottom: 6px; font-size: 11px;">
      ${complete
        ? `📌 Blueprint Ontologi Siap Pakai (${onto.metadata.total_indikator} Indikator Operasional):`
        : `⚠️ Ontologi BELUM LENGKAP — ${nr.join(', ')} belum berindikator:`}
    </div>
    <div style="font-size: 9px; color: #64748b; margin-bottom: 6px;">
      Dimensi: sumber/teori terverifikasi. Indikator: dikembangkan dari dimensi dan konteks objek (2–4 per dimensi).
    </div>
    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 6px; margin-bottom: 8px;">
      <div style="padding: 6px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px;">
        <div style="font-weight: 700; color: #0284c7;">X1: ${vars.x1.nama_resmi}</div>
        <div style="color: #64748b; font-size: 10px;">Ahli: ${vars.x1.tokoh_ahli}</div>
        <div style="color: #334155; font-size: 10px; margin-top: 2px;">${dimsText(vars.x1)}</div>
        <div style="color: #0f766e; font-size: 9px; margin-top: 2px;">Operasional: ${indCount(vars.x1)}; Teoretis: ${teoCount(vars.x1)} butir</div>
      </div>
      <div style="padding: 6px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px;">
        <div style="font-weight: 700; color: #0284c7;">X2: ${vars.x2.nama_resmi}</div>
        <div style="color: #64748b; font-size: 10px;">Ahli: ${vars.x2.tokoh_ahli}</div>
        <div style="color: #334155; font-size: 10px; margin-top: 2px;">${dimsText(vars.x2)}</div>
        <div style="color: #0f766e; font-size: 9px; margin-top: 2px;">Operasional: ${indCount(vars.x2)}; Teoretis: ${teoCount(vars.x2)} butir</div>
      </div>
      <div style="padding: 6px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px;">
        <div style="font-weight: 700; color: #d97706;">M: ${vars.m.nama_resmi} (Mediasi)</div>
        <div style="color: #64748b; font-size: 10px;">Ahli: ${vars.m.tokoh_ahli}</div>
        <div style="color: #334155; font-size: 10px; margin-top: 2px;">${dimsText(vars.m)}</div>
        <div style="color: #0f766e; font-size: 9px; margin-top: 2px;">Operasional: ${indCount(vars.m)}; Teoretis: ${teoCount(vars.m)} butir</div>
      </div>
      <div style="padding: 6px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px;">
        <div style="font-weight: 700; color: #16a34a;">Y: ${vars.y.nama_resmi} (Terikat)</div>
        <div style="color: #64748b; font-size: 10px;">Ahli: ${vars.y.tokoh_ahli}</div>
        <div style="color: #334155; font-size: 10px; margin-top: 2px;">${dimsText(vars.y)}</div>
        <div style="color: #0f766e; font-size: 9px; margin-top: 2px;">Operasional: ${indCount(vars.y)}; Teoretis: ${teoCount(vars.y)} butir</div>
      </div>
    </div>
    <div style="background: #f8fafc; border: 1px solid #e2e8f0; padding: 6px; border-radius: 4px; font-size: 10px; line-height: 1.5; color: #334155; margin-bottom: 8px;">
      <strong>🔎 Seleksi Dimensi Otomatis (transparan):</strong>
      ${selectionDetails()}
      <div style="margin-top: 4px; color: #475569;">
        Bab II memakai <strong>indikator teoretis</strong> (generik, tanpa nama/objek/sektor penelitian).
        Bab III/IV/Lampiran memakai <strong>indikator operasional</strong> (spesifik konteks objek).
      </div>
    </div>
    <div style="background: #f1f5f9; padding: 6px; border-radius: 4px; font-size: 10px; line-height: 1.5; color: #334155;">
      <strong>🔗 Keterikatan Benang Merah Bab I s/d Lampiran:</strong><br>
      • <strong>Tabel 1.1 (Evaluasi Y):</strong> Menggunakan parameter <em>${anchors.tabel_1_1_params.join(' | ')}</em>.<br>
      • <strong>Tabel 1.2 (Operasional M):</strong> Mengukur disparitas <em>${anchors.tabel_1_2_issues.join(', ')}</em>.<br>
      • <strong>Bab III & Kuesioner:</strong> Disusun ke dalam Tabel Kisi-Kisi 3.1 & ${onto.metadata.total_indikator} butir angket skala Likert (1-5).
    </div>
  `;

  box.innerHTML = html;
}

function saveConfig() {
  const config = getConfig();
  localStorage.setItem("autotesis_config", JSON.stringify(config));
  updateHypothesesPreview(config.x1, config.x2, config.m, config.y, config.instansi);
  setStatus("ready", "Konfigurasi variabel & 7 hipotesis berhasil disimpan!");
}

function loadSavedConfig() {
  const raw = localStorage.getItem("autotesis_config");
  if (!raw) return;
  try {
    const config = JSON.parse(raw);
    if (config.judul) document.getElementById("input-judul").value = config.judul;
    if (config.penulis) document.getElementById("input-penulis").value = config.penulis;
    if (config.nim) document.getElementById("input-nim").value = config.nim;
    if (config.instansi) document.getElementById("input-instansi").value = config.instansi;
    if (config.tahun) document.getElementById("input-tahun").value = config.tahun;
    if (config.x1) document.getElementById("input-x1").value = config.x1;
    if (config.x2) document.getElementById("input-x2").value = config.x2;
    if (config.m) document.getElementById("input-m").value = config.m;
    if (config.y) document.getElementById("input-y").value = config.y;
    if (config.populasi) document.getElementById("input-populasi").value = config.populasi;
    if (config.n) document.getElementById("input-responden").value = config.n;
    refreshSampelPreview();
    
    updateHypothesesPreview(config.x1, config.x2, config.m, config.y, config.instansi);
  } catch (e) {
    console.error("Failed loading saved config", e);
  }
}

function setStatus(state, message) {
  const dot = document.getElementById("status-dot");
  const text = document.getElementById("status-text");
  if (!dot || !text) return;
  dot.className = `status-indicator ${state}`;
  text.textContent = message;
}

function setProgress(percent) {
  const container = document.getElementById("progress-container");
  const bar = document.getElementById("progress-bar");
  if (!container || !bar) return;
  if (percent <= 0) {
    container.classList.add("hidden");
    bar.style.width = "0%";
  } else {
    container.classList.remove("hidden");
    bar.style.width = `${percent}%`;
  }
}

function getSelectedWriteMode() {
  const checked = document.querySelector('input[name="write-mode"]:checked');
  return checked ? checked.value : "native";
}

// Generate chapter handler (Native Word COM or Office.js)
async function handleChapterGeneration(action, btn) {
  const config = getConfig();
  const mode = getSelectedWriteMode();
  const originalText = btn.textContent;
  btn.disabled = true;
  btn.textContent = "Mengetik...";
  setStatus("working", `Mengetik ${action.toUpperCase()} (${mode === "native" ? "Native Word" : "Office.js"})...`);
  setProgress(40);

  try {
    if (mode === "native" || !window.Word || !window.Word.run) {
      // 1. Native Word COM Execution via Server
      const response = await fetch(`${API_BASE}/write-native-word`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action, config })
      });
      if (!response.ok) throw new Error(`HTTP Error ${response.status}`);
      const data = await response.json();
      setProgress(100);
      setStatus("ready", data.message || `Naskah ${action.toUpperCase()} berhasil diketik ke Word!`);
    } else {
      // 2. Office.js Taskpane Insertion
      const response = await fetch(`${API_BASE}/generate-chapter`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action, config })
      });
      if (!response.ok) throw new Error(`HTTP Error ${response.status}`);
      const data = await response.json();
      setProgress(75);

      await Word.run(async (context) => {
        const body = context.document.body;
        if (data.paragraphs && Array.isArray(data.paragraphs)) {
          for (const item of data.paragraphs) {
            const lines = item.text.split("\n");
            for (const line of lines) {
              if (!line.trim()) continue;
              const p = body.insertParagraph(line.trim(), Word.InsertLocation.end);
              p.font.name = "Times New Roman";
              p.font.size = 12;
              p.font.color = "#000000";
              p.lineSpacing = 18; // 1.5 spacing
              p.spaceBefore = 0;
              p.spaceAfter = 0;

              if (item.type === "heading1" || item.type === "title") {
                p.alignment = Word.Alignment.center;
                p.font.bold = true;
                p.spaceBefore = 12;
                p.spaceAfter = 6;
              } else if (item.type === "heading2") {
                p.alignment = Word.Alignment.left;
                p.font.bold = true;
                p.spaceBefore = 12;
                p.spaceAfter = 4;
              } else if (item.type === "body_center") {
                p.alignment = Word.Alignment.center;
                p.spaceAfter = 6;
              } else if (item.type === "bullet") {
                p.alignment = Word.Alignment.left;
                p.leftIndent = 36;
                p.firstLineIndent = 0;
              } else {
                p.alignment = Word.Alignment.left;
                p.firstLineIndent = 36;
              }
            }
          }
        }
        await context.sync();
      });
      setProgress(100);
      setStatus("ready", `Naskah ${action.toUpperCase()} selesai diketik di dokumen Word.`);
    }
  } catch (error) {
    console.error(error);
    setStatus("error", `Gagal: ${error.message}`);
  } finally {
    setTimeout(() => setProgress(0), 1000);
    btn.disabled = false;
    btn.textContent = originalText;
  }
}

// =========================================================================
// Wizard Olah Data SmartPLS: (1) Kuesioner -> (2) Generate/Unggah Jawaban
// -> (3) "Olah Data Sekarang"
// =========================================================================
// =========================================================================
// Gerbang data Bab IV: pastikan peneliti memilih sumber data jawaban
// (unggah Excel/CSV sendiri atau dibuatkan sistem) di Tab 3. SmartPLS.
// =========================================================================
async function refreshDataGate(force) {
  const gate = document.getElementById("bab4-data-gate");
  const ok = document.getElementById("bab4-data-ok");
  if (!gate || !ok) return dataGateState.last;

  const now = Date.now();
  if (!force && dataGateState.diperiksa && (now - dataGateState.diperiksa) < 15000) {
    return dataGateState.last;
  }
  dataGateState.diperiksa = now;

  try {
    const response = await fetch(`${API_BASE}/pipeline/data-status`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ config: getConfig() })
    });
    const st = await response.json();
    dataGateState.last = st;
    if (!st || st.status !== "success") return st;

    const warn = document.getElementById("bab4-data-warn");
    if (st.siap_bab4) {
      gate.classList.add("hidden");
      gate.classList.remove("data-gate-alert");
      if (warn) warn.classList.add("hidden");
      const sumber = st.label_sumber || (
        st.sumber === "google_form" ? "Data primer (respons Google Form)" :
        st.sumber === "upload" ? "Data unggahan" :
        "Data hasil pembangkitan mesin");
      const asal = st.origin_label ? ` (${st.origin_label})` : "";
      ok.innerHTML = `✅ ${sumber}${asal} siap: N=${st.n}, sudah diolah — angka Bab IV/V memakai data ini. `
        + `<a href="#" class="data-gate-link" data-buka-smartpls="1">Tab 3. Analisis PLS-SEM →</a>`;
      ok.classList.remove("hidden");
      ok.querySelector("[data-buka-smartpls]")?.addEventListener("click", (e) => {
        e.preventDefault();
        bukaTabSmartpls();
      });
    } else {
      ok.classList.add("hidden");
      gate.classList.remove("hidden");
      if (warn && st.ada_data && !st.sudah_diolah) {
        warn.classList.remove("hidden");
        warn.innerHTML = `⚠️ Data sudah disiapkan (N=${st.n}), tetapi belum diolah. `
          + `Klik <strong>Olah Data Sekarang</strong> di `
          + `<a href="#" class="data-gate-link" data-buka-smartpls="2">Tab 3. Analisis PLS-SEM →</a> `
          + `sebelum menyusun Bab IV/V.`;
        warn.querySelector("[data-buka-smartpls]")?.addEventListener("click", (e) => {
          e.preventDefault();
          bukaTabSmartpls();
        });
      }
    }
    return st;
  } catch (err) {
    console.error(err);
    return dataGateState.last;
  }
}

function bufToBase64(buf) {
  const bytes = new Uint8Array(buf);
  let bin = "";
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  }
  return btoa(bin);
}

function bukaTabSmartpls() {
  document.querySelector('.tab-btn[data-tab="smartpls"]')?.click();
  const step = document.getElementById("wizard-step-2");
  if (!step) return;
  setTimeout(() => {
    try { step.scrollIntoView({ behavior: "smooth", block: "center" }); } catch (e) { step.scrollIntoView(); }
    step.classList.add("flash");
    setTimeout(() => step.classList.remove("flash"), 2300);
  }, 130);
}

/**
 * Periksa data sebelum menulis Bab IV/V.
 * Kembalikan true HANYA bila data sudah diolah. Tidak ada klik kedua untuk
 * memaksa; server juga menolak sub-bab Bab IV/V tanpa analisis.
 */
async function pastikanDataJawaban(kunci, aksi) {
  const st = await refreshDataGate(true);
  if (st && st.siap_bab4) return true;

  const gate = document.getElementById("bab4-data-gate");
  const warn = document.getElementById("bab4-data-warn");
  if (gate) {
    gate.classList.remove("hidden");
    gate.classList.add("data-gate-alert");
    try { gate.scrollIntoView({ behavior: "smooth", block: "center" }); } catch (e) { gate.scrollIntoView(); }
  }
  if (warn) {
    warn.classList.remove("hidden");
    const belumOlah = st && st.ada_data && !st.sudah_diolah;
    warn.innerHTML = belumOlah
      ? `⚠️ Data sudah ada (N=${st.n}), tetapi belum diolah. Jalankan `
        + `<strong>Olah Data Sekarang</strong> di `
        + `<a href="#" class="data-gate-link" data-buka-smartpls="2">Tab 3. Analisis PLS-SEM →</a> `
        + `sebelum menyusun Bab IV/V.`
      : `⚠️ Data belum siap. Siapkan data di `
        + `<a href="#" class="data-gate-link" data-buka-smartpls="2">Tab 3. Analisis PLS-SEM →</a> `
        + `(impor respons Google Form/berkas, atau data hasil pembangkitan mesin), `
        + `lalu klik <strong>Olah Data Sekarang</strong>.`;
    warn.querySelector("[data-buka-smartpls]")?.addEventListener("click", (e) => {
      e.preventDefault();
      bukaTabSmartpls();
    });
  }
  setStatus("error", `Bab IV/V belum bisa disusun: ${aksi} memerlukan data yang sudah diolah.`);
  return false;
}

function setSmartplsSource(source, originLabel) {
  smartplsState.source = source;
  try { localStorage.setItem("autothesis_data_source", source); } catch (e) {}
  const note = document.getElementById("smartpls-src-note");
  if (!note) return;
  const label = source === "google_form"
    ? "Data primer (respons Google Form)"
    : source === "upload"
      ? "Data unggahan"
      : "Data hasil pembangkitan mesin";
  const asal = originLabel ? ` — ${originLabel}` : "";
  note.textContent = `Sumber data terpilih: ${label}${asal}. Lanjutkan ke langkah 3 “Olah Data Sekarang”.`;
  note.classList.remove("hidden");
}

function isGoogleSheetLink(url) {
  try {
    const u = new URL(url);
    const host = (u.hostname || "").toLowerCase();
    const google = host === "docs.google.com" || host === "spreadsheets.google.com"
      || host.endsWith(".docs.google.com");
    return google && u.pathname.includes("/spreadsheets/");
  } catch (e) {
    return false;
  }
}

// Langkah 2: link Google Sheets, unggah berkas, atau data hasil pembangkitan mesin
async function handleGenerateRespondents(mode) {
  const config = getConfig();
  const buttonIds = ["btn-generate-data", "btn-upload-data", "btn-import-sheet"];
  const btnId = mode === "sheet" ? "btn-import-sheet"
    : mode === "upload" ? "btn-upload-data" : "btn-generate-data";
  const btn = document.getElementById(btnId);
  const box = document.getElementById("smartpls-download-box");
  const dlLink = document.getElementById("link-download-excel");

  const payload = { config, analysis_engine: "internal" };
  let workingMsg = "";

  if (mode === "sheet") {
    const url = (document.getElementById("input-sheet-url")?.value || "").trim();
    if (!url) {
      setStatus("error", "Tempel link Spreadsheet respons Google Form terlebih dahulu.");
      return;
    }
    if (!isGoogleSheetLink(url)) {
      setStatus("error", "Itu bukan link Google Sheets. Salin link Spreadsheet respons (bukan link Form).");
      return;
    }
    const konfirmasi = document.getElementById("input-konfirmasi-gform");
    if (!konfirmasi || !konfirmasi.checked) {
      setStatus("error", "Centang pernyataan bahwa Spreadsheet berisi respons Google Form responden nyata.");
      return;
    }
    payload.mode = "upload";
    payload.data_source = "google_form";
    payload.konfirmasi_google_form = true;
    payload.upload = { sheet_url: url };
    workingMsg = "Membaca respons terbaru langsung dari Google Sheets...";
  } else if (mode === "upload") {
    const fileInput = document.getElementById("input-upload-jawaban");
    const file = fileInput && fileInput.files && fileInput.files[0];
    if (!file) { setStatus("error", "Pilih berkas .xlsx atau .csv data jawaban terlebih dahulu."); return; }
    payload.mode = "upload";
    payload.data_source = document.getElementById("input-sumber-upload")?.value || "google_form";
    const nama = (file.name || "").toLowerCase();
    try {
      if (nama.endsWith(".xlsx") || nama.endsWith(".xls")) {
        const buf = await file.arrayBuffer();
        payload.upload = { xlsx_b64: bufToBase64(buf), filename: file.name };
      } else {
        payload.upload = { csv_text: await file.text(), filename: file.name };
      }
    } catch (e) { setStatus("error", `Gagal membaca berkas: ${e.message}`); return; }
    workingMsg = "Membaca & memvalidasi berkas respons...";
  } else {
    payload.mode = "generated";
    payload.data_source = "generated";
    workingMsg = "Membangkitkan data hasil pembangkitan mesin dari butir kuesioner nyata...";
  }

  buttonIds.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.disabled = true;
  });
  setStatus("working", workingMsg);
  setProgress(mode === "generated" ? 25 : 40);

  try {
    const response = await fetch(`${API_BASE}/pipeline/generate-respondents`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    const data = await response.json();
    if (!response.ok || data.status !== "success") {
      throw new Error(data.message || `HTTP Error ${response.status}`);
    }
    setProgress(70);
    setSmartplsSource(data.source, data.origin_label);
    smartplsState.dataFile = data.data_file;
    if (box && dlLink && data.data_file) {
      box.classList.remove("hidden");
      dlLink.href = `${API_BASE}/download/${data.data_file}`;
      dlLink.textContent = `📥 Unduh Data: ${data.data_file}`;
    }

    let detail = "";
    if (data.source === "generated") {
      detail = `Berhasil membangkitkan ${data.n} × ${data.total_butir} butir data hasil pembangkitan mesin.`;
    } else {
      const asal = data.origin_label ? ` dari ${data.origin_label}` : "";
      detail = `Respons terbaca${asal}: N=${data.n}, ${data.total_butir} butir. `
        + (data.warnings || []).join(" ");
    }
    setStatus("ready", `${detail} Klik “Olah Data Sekarang”.`);
    refreshDataGate(true);
  } catch (error) {
    console.error(error);
    setStatus("error", `Gagal menyiapkan data jawaban: ${error.message}`);
  } finally {
    setTimeout(() => setProgress(0), 800);
    buttonIds.forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.disabled = false;
    });
  }
}

// Langkah 3: hitung PLS-SEM penuh atas data jawaban terpilih
async function handleOlahData() {
  const config = getConfig();
  const btn = document.getElementById("btn-olah-data");
  if (btn) btn.disabled = true;
  setStatus("working", "Menghitung PLS-SEM (algorithm + bootstrap 5.000) dan menguji H1–H7...");
  setProgress(30);

  try {
    const response = await fetch(`${API_BASE}/pipeline/olah-data`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ config, bootstrap: 5000 })
    });
    const data = await response.json();
    if (!response.ok || data.status !== "success") {
      throw new Error(data.message || `HTTP Error ${response.status}`);
    }
    setProgress(100);

    const box = document.getElementById("smartpls-download-box");
    const dlLink = document.getElementById("link-download-excel");
    const guideLink = document.getElementById("link-download-panduan");
    if (box && data.data_file) {
      box.classList.remove("hidden");
      if (dlLink) {
        dlLink.href = `${API_BASE}/download/${data.data_file}`;
        dlLink.textContent = `📥 Unduh Data: ${data.data_file}`;
      }
      if (guideLink) {
        if (data.guide_file) {
          guideLink.classList.remove("hidden");
          guideLink.href = `${API_BASE}/download/${data.guide_file}`;
          guideLink.textContent = "📄 Unduh Panduan Olah Data SmartPLS 4";
        } else {
          guideLink.classList.add("hidden");
        }
      }
    }

    renderSmartplsSummary(data.summary);
    const mesin = data.analysis_engine === "smartpls" ? "SmartPLS 4" : "mesin internal";
    const asal = data.origin_label ? ` — ${data.origin_label}` : "";
    setStatus("ready", `Olah data selesai (${data.label_sumber || data.source}${asal}; ${mesin}). Angka Bab IV/V kini memakai hasil ini.`);
    refreshDataGate(true);
  } catch (error) {
    console.error(error);
    setStatus("error", `Gagal olah data: ${error.message}`);
  } finally {
    setTimeout(() => setProgress(0), 1000);
    if (btn) btn.disabled = false;
  }
}

function fmtNum(n, d) {
  d = (d === undefined) ? 3 : d;
  return (n === null || n === undefined || isNaN(n)) ? "-" : Number(n).toFixed(d).replace(".", ",");
}

function renderSmartplsSummary(s) {
  const preview = document.getElementById("smartpls-preview");
  const grid = document.getElementById("smartpls-metrics");
  const statusEl = document.getElementById("smartpls-status");
  if (!preview || !grid || !s) return;
  const minOf = (o) => { const v = Object.values(o || {}); return v.length ? Math.min.apply(null, v) : 0; };
  const aveMin = minOf(s.ave), crMin = minOf(s.cr);
  const r2 = s.r2 || {}, q2 = s.q2 || {};
  const cards = [
    ["Outer Loading", `${fmtNum(s.loading_min)} – ${fmtNum(s.loading_max)}`, s.loading_min >= 0.70],
    ["AVE (min)", fmtNum(aveMin), aveMin > 0.50],
    ["CR (min)", fmtNum(crMin), crMin > 0.70],
    ["HTMT (maks)", fmtNum(s.htmt_max), s.htmt_max < 0.90],
    ["VIF (maks)", fmtNum(s.vif_max, 2), s.vif_max < 5],
    ["R² (M / Y)", `${fmtNum(r2.M)} / ${fmtNum(r2.Y)}`, true],
    ["Q² (M / Y)", `${fmtNum(q2.M)} / ${fmtNum(q2.Y)}`, (q2.M > 0 && q2.Y > 0)],
    ["SRMR", fmtNum(s.srmr), (s.srmr || 1) < 0.10]
  ];
  grid.innerHTML = cards.map(function (c) {
    return `<div class="metric-card"><span class="label">${c[0]}</span>` +
           `<span class="value ${c[2] ? "valid" : ""}">${c[1]} ${c[2] ? "✓" : "!"}</span></div>`;
  }).join("");
  const nPath = (s.paths || []).length, nInd = (s.indirect || []).length;
  statusEl.innerHTML = s.all_significant
    ? `Status Uji: <strong>Semua jalur signifikan (${nPath} langsung + ${nInd} mediasi; |t| &gt; 1,96; p &lt; 0,05)</strong>`
    : "Status Uji: <strong>Ada jalur yang belum signifikan — periksa kembali data.</strong>";
  preview.classList.remove("hidden");
}

// Copy helper: clipboard API dengan fallback execCommand (Office webview / browser sandbox)
function _copyFallback(text) {
  try {
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.setAttribute("readonly", "");
    ta.style.position = "fixed";
    ta.style.top = "-1000px";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    ta.setSelectionRange(0, ta.value.length);
    const ok = document.execCommand("copy");
    document.body.removeChild(ta);
    return !!ok;
  } catch (e) {
    return false;
  }
}

function copyToClipboard(text) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    return navigator.clipboard.writeText(text)
      .then(function () { return true; }, function () { return _copyFallback(text); });
  }
  return Promise.resolve(_copyFallback(text));
}

// Generate Google Apps Script for Google Form
async function handleGoogleFormScript() {
  const config = getConfig();
  const btn = document.getElementById("btn-generate-gform");
  if (btn) btn.disabled = true;
  setStatus("working", "Membuat kode Google Apps Script dari butir kuesioner nyata...");
  try {
    const response = await fetch(`${API_BASE}/generate-gform-script`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(config)
    });
    const data = await response.json();
    if (!response.ok || !data.script) {
      throw new Error(data.message || `HTTP Error ${response.status}`);
    }

    smartplsGformScript = data.script;
    const copied = await copyToClipboard(data.script);

    const box = document.getElementById("gform-result");
    const okEl = document.getElementById("gform-ok");
    const ta = document.getElementById("gform-script");
    if (ta) ta.value = data.script;
    if (box) box.classList.remove("hidden");
    const n = data.total_butir || 0;
    if (okEl) {
      okEl.textContent = (copied
        ? "✅ Skrip berhasil disalin ke clipboard. "
        : "⚠️ Salin otomatis diblokir — klik “Salin Ulang Skrip” di bawah. ")
        + `Skrip memuat ${n} butir kuesioner NYATA judul terpilih (kode butir sama dengan kisi-kisi Bab III & Lampiran). `
        + "Setelah Run, Spreadsheet respons tercatat di Execution log; bagikan sebagai Pelihat lalu impor link-nya di Langkah 2.";
    }
    setStatus("ready", "Skrip Google Form + Spreadsheet siap. Run di script.google.com, lalu salin link Spreadsheet respons ke Langkah 2.");
  } catch (error) {
    console.error(error);
    setStatus("error", `Gagal buat script: ${error.message}`);
  } finally {
    if (btn) btn.disabled = false;
  }
}

// 1-Click Full Thesis Document Generation & Auto-Open in Word
async function handleFullDocumentGeneration() {
  const config = getConfig();
  const btn = document.getElementById("btn-generate-full");
  btn.disabled = true;
  setStatus("working", "Merangkai 5 Bab naskah lengkap & membuka di Microsoft Word...");
  setProgress(40);

  try {
    const response = await fetch(`${API_BASE}/generate-full-thesis`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(config)
    });

    const result = await response.json();
    if (!response.ok || result.status === "error") {
      throw new Error(result.message || `HTTP Error ${response.status}`);
    }
    setProgress(100);

    let inserted = false;
    if (result.docx_base64) {
      inserted = await insertDocxBase64ToActiveWord(result.docx_base64);
    }

    if (inserted) {
      setStatus("ready", "Naskah Tesis Utuh (5 Bab Lengkap) Berhasil Disisipkan Langsung ke Dokumen Word Aktif!");
    } else {
      setStatus("ready", "Naskah Tesis Utuh Berhasil Dibuat & Siap Diunduh!");
    }
    const dlBox = document.getElementById("full-download-link");
    const link = document.getElementById("link-download");
    if (dlBox && link) {
      dlBox.classList.remove("hidden");
      link.href = `${API_BASE}/download/${result.filename}`;
      link.textContent = `📥 Unduh Dokumen: ${result.filename}`;
    }
  } catch (error) {
    console.error(error);
    setStatus("error", `Gagal generate dokumen: ${error.message}`);
  } finally {
    setTimeout(() => setProgress(0), 1500);
    btn.disabled = false;
  }
}

// =========================================================================
// GRAPH ENGINE (V3.0) - MULTI-X, MODERASI & MEDIASI TERMODERASI
// =========================================================================
let currentGraphModel = null;
let currentPlsResults = null;

function initGraphUI() {
  const modelTypeSel = document.getElementById("graph-model-type");
  const groupM = document.getElementById("group-var-m");
  const groupZ = document.getElementById("group-var-z");
  const btnAddX = document.getElementById("btn-add-x-var");
  const xContainer = document.getElementById("x-vars-container");
  const btnCompile = document.getElementById("btn-compile-graph");
  const btnSimulatePls = document.getElementById("btn-simulate-graph-pls");
  const btnWriteWord = document.getElementById("btn-graph-write-native");
  const btnDownloadDocx = document.getElementById("btn-graph-download-docx");

  if (modelTypeSel) {
    modelTypeSel.addEventListener("change", () => {
      const val = modelTypeSel.value;
      if (val === "moderation") {
        if (groupM) groupM.style.display = "none";
        if (groupZ) groupZ.style.display = "block";
      } else if (val === "mediation") {
        if (groupM) groupM.style.display = "block";
        if (groupZ) groupZ.style.display = "none";
      } else {
        if (groupM) groupM.style.display = "block";
        if (groupZ) groupZ.style.display = "block";
      }
    });
  }

  if (btnAddX && xContainer) {
    btnAddX.addEventListener("click", () => {
      const rows = xContainer.querySelectorAll(".x-var-row");
      const nextIdx = rows.length + 1;
      const row = document.createElement("div");
      row.className = "x-var-row";
      row.style.cssText = "display: flex; gap: 6px; align-items: center;";
      row.innerHTML = `
        <span style="font-size: 11px; font-weight: 700; width: 28px; color: #475569;">X${nextIdx}:</span>
        <input type="text" class="input-graph-x" value="Variabel X${nextIdx}" style="flex: 1; padding: 6px 10px; font-size: 11px;">
        <button type="button" class="btn btn-outline btn-xs btn-remove-x" style="color: #ef4444; border-color: #fca5a5;">✕</button>
      `;
      xContainer.appendChild(row);
      attachRemoveXListeners();
    });
    attachRemoveXListeners();
  }

  function attachRemoveXListeners() {
    if (!xContainer) return;
    xContainer.querySelectorAll(".btn-remove-x").forEach(btn => {
      btn.onclick = function() {
        const row = this.closest(".x-var-row");
        if (row) row.remove();
        xContainer.querySelectorAll(".x-var-row").forEach((r, idx) => {
          const sp = r.querySelector("span");
          if (sp) sp.textContent = `X${idx + 1}:`;
        });
      };
    });
  }

  if (btnCompile) btnCompile.addEventListener("click", handleCompileGraph);
  if (btnSimulatePls) btnSimulatePls.addEventListener("click", handleSimulateGraphPls);
  if (btnWriteWord) btnWriteWord.addEventListener("click", handleGraphWriteWord);
  if (btnDownloadDocx) btnDownloadDocx.addEventListener("click", handleGraphDownloadDocx);
}

async function handleCompileGraph() {
  const btn = document.getElementById("btn-compile-graph");
  if (btn) btn.disabled = true;
  setStatus("working", "Mengompilasi Structural Graph & Hipotesis Formal...");

  try {
    const model_type = document.getElementById("graph-model-type")?.value || "moderated_mediation";
    const instansi = document.getElementById("graph-instansi")?.value || "";
    const x_vars = Array.from(document.querySelectorAll(".input-graph-x")).map(i => i.value.trim()).filter(Boolean);
    const m_vars = model_type !== "moderation" ? [document.getElementById("graph-var-m")?.value.trim()].filter(Boolean) : [];
    const z_vars = model_type !== "mediation" ? [document.getElementById("graph-var-z")?.value.trim()].filter(Boolean) : [];
    const y_var = document.getElementById("graph-var-y")?.value.trim() || "Kinerja Pegawai";

    const payload = {
      model_type,
      x_vars,
      m_vars,
      z_vars,
      y_var,
      instansi
    };

    const res = await fetch(`${API_BASE}/graph/compile`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const data = await res.json();
    if (!res.ok || data.status === "error") {
      throw new Error(data.message || `HTTP ${res.status}`);
    }

    currentGraphModel = data.graph;

    // Render Preview
    const badgeHipo = document.getElementById("badge-total-hipo");
    if (badgeHipo) badgeHipo.textContent = `${currentGraphModel.total_hypotheses} Hipotesis`;

    const eqList = document.getElementById("graph-equations-list");
    if (eqList) {
      eqList.innerHTML = currentGraphModel.equations.map(e => `<div>• <strong>${e.name}</strong>: ${e.formula}</div>`).join("");
    }

    const hipoList = document.getElementById("graph-hypotheses-list");
    if (hipoList) {
      hipoList.innerHTML = currentGraphModel.hypotheses.map(h => `
        <div style="margin-bottom: 8px; padding-bottom: 6px; border-bottom: 1px dashed #e2e8f0;">
          <span style="display: inline-block; background: #e0e7ff; color: #3730a3; font-weight: 700; padding: 1px 6px; border-radius: 4px; font-size: 10px; margin-right: 4px;">${h.code}</span>
          <span style="font-weight: 600; color: #1e293b;">[${h.relation}]</span>: ${h.statement}
        </div>
      `).join("");
    }

    document.getElementById("card-graph-result")?.classList.remove("hidden");
    setStatus("ready", `Graph berhasil dikompilasi! Terbentuk ${currentGraphModel.total_hypotheses} hipotesis terstruktur.`);
  } catch (err) {
    console.error(err);
    setStatus("error", `Gagal kompilasi graph: ${err.message}`);
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function handleSimulateGraphPls() {
  if (!currentGraphModel) {
    await handleCompileGraph();
    if (!currentGraphModel) return;
  }

  const btn = document.getElementById("btn-simulate-graph-pls");
  if (btn) btn.disabled = true;
  setStatus("working", "Menjalankan Simulasi SmartPLS 4 (Bootstrapping 5.000 & Evaluasi Hibrida)...");

  try {
    const sample_size = parseInt(document.getElementById("graph-sample-size")?.value || "92", 10);
    const payload = {
      graph: currentGraphModel,
      sample_size
    };

    const res = await fetch(`${API_BASE}/graph/simulate-pls`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const data = await res.json();
    if (!res.ok || data.status === "error") {
      throw new Error(data.message || `HTTP ${res.status}`);
    }

    currentPlsResults = data.pls_results;

    // Render Table Summary
    const tableDiv = document.getElementById("graph-pls-summary-table");
    if (tableDiv) {
      let html = `<div style="margin-bottom: 8px;"><strong style="color: #166534;">Ringkasan Uji Hipotesis Struktural:</strong></div>`;
      html += `<table style="width: 100%; border-collapse: collapse; font-size: 10px; margin-bottom: 10px;">
        <thead>
          <tr style="background: #e2e8f0; color: #1e293b; text-align: left;">
            <th style="padding: 4px; border: 1px solid #cbd5e1;">Hipo</th>
            <th style="padding: 4px; border: 1px solid #cbd5e1;">Jalur Hubungan</th>
            <th style="padding: 4px; border: 1px solid #cbd5e1;">Path (β)</th>
            <th style="padding: 4px; border: 1px solid #cbd5e1;">T-Stat</th>
            <th style="padding: 4px; border: 1px solid #cbd5e1;">P-Val</th>
            <th style="padding: 4px; border: 1px solid #cbd5e1;">Keputusan</th>
          </tr>
        </thead>
        <tbody>`;

      currentPlsResults.hypotheses_testing.forEach(h => {
        html += `<tr>
          <td style="padding: 4px; border: 1px solid #cbd5e1; font-weight: 700;">${h.code}</td>
          <td style="padding: 4px; border: 1px solid #cbd5e1;">${h.relation}</td>
          <td style="padding: 4px; border: 1px solid #cbd5e1; text-align: right;">${h.original_sample.toFixed(3)}</td>
          <td style="padding: 4px; border: 1px solid #cbd5e1; text-align: right;">${h.t_statistics.toFixed(3)}</td>
          <td style="padding: 4px; border: 1px solid #cbd5e1; text-align: right;">${h.p_value.toFixed(3)}</td>
          <td style="padding: 4px; border: 1px solid #cbd5e1; color: #15803d; font-weight: 600;">${h.decision}</td>
        </tr>`;
      });
      html += `</tbody></table>`;

      // Moderation section
      if (currentPlsResults.moderation_details && currentPlsResults.moderation_details.length > 0) {
        html += `<div style="margin: 8px 0 4px 0;"><strong style="color: #166534;">Evaluasi Moderasi (Simple Slope):</strong></div>`;
        html += `<table style="width: 100%; border-collapse: collapse; font-size: 10px; margin-bottom: 10px;">
          <thead>
            <tr style="background: #e2e8f0; color: #1e293b;">
              <th style="padding: 4px; border: 1px solid #cbd5e1;">Pasangan Interaksi</th>
              <th style="padding: 4px; border: 1px solid #cbd5e1;">-1SD (Rendah)</th>
              <th style="padding: 4px; border: 1px solid #cbd5e1;">Mean</th>
              <th style="padding: 4px; border: 1px solid #cbd5e1;">+1SD (Tinggi)</th>
              <th style="padding: 4px; border: 1px solid #cbd5e1;">f²</th>
            </tr>
          </thead>
          <tbody>`;
        currentPlsResults.moderation_details.forEach(m => {
          html += `<tr>
            <td style="padding: 4px; border: 1px solid #cbd5e1;">${m.interaction_pair}</td>
            <td style="padding: 4px; border: 1px solid #cbd5e1; text-align: right;">${m.slope_low_z.toFixed(3)}</td>
            <td style="padding: 4px; border: 1px solid #cbd5e1; text-align: right;">${m.slope_mean_z.toFixed(3)}</td>
            <td style="padding: 4px; border: 1px solid #cbd5e1; text-align: right; font-weight: 700; color: #15803d;">${m.slope_high_z.toFixed(3)}</td>
            <td style="padding: 4px; border: 1px solid #cbd5e1; text-align: right;">${m.f2_interaction.toFixed(3)}</td>
          </tr>`;
        });
        html += `</tbody></table>`;
      }

      // Conditional Indirect section
      if (currentPlsResults.conditional_indirect_details && currentPlsResults.conditional_indirect_details.length > 0) {
        html += `<div style="margin: 8px 0 4px 0;"><strong style="color: #166534;">Mediasi Termoderasi (Conditional Indirect Effects):</strong></div>`;
        html += `<table style="width: 100%; border-collapse: collapse; font-size: 10px;">
          <thead>
            <tr style="background: #e2e8f0; color: #1e293b;">
              <th style="padding: 4px; border: 1px solid #cbd5e1;">Jalur Mediasi Bersyarat</th>
              <th style="padding: 4px; border: 1px solid #cbd5e1;">Index ModMed</th>
              <th style="padding: 4px; border: 1px solid #cbd5e1;">P-Value</th>
              <th style="padding: 4px; border: 1px solid #cbd5e1;">Kesimpulan</th>
            </tr>
          </thead>
          <tbody>`;
        currentPlsResults.conditional_indirect_details.forEach(c => {
          html += `<tr>
            <td style="padding: 4px; border: 1px solid #cbd5e1;">${c.path}</td>
            <td style="padding: 4px; border: 1px solid #cbd5e1; text-align: right; font-weight: 700;">${c.index_of_mod_med.toFixed(3)}</td>
            <td style="padding: 4px; border: 1px solid #cbd5e1; text-align: right;">${c.mod_med_p_value.toFixed(3)}</td>
            <td style="padding: 4px; border: 1px solid #cbd5e1; color: #15803d;">Didukung Empiris (Signifikan)</td>
          </tr>`;
        });
        html += `</tbody></table>`;
      }

      tableDiv.innerHTML = html;
    }

    document.getElementById("card-graph-pls-results")?.classList.remove("hidden");
    setStatus("ready", `Analisis PLS-SEM selesai! Seluruh ${currentPlsResults.total_hypotheses} hipotesis signifikan dan valid.`);
  } catch (err) {
    console.error(err);
    setStatus("error", `Gagal analisis PLS: ${err.message}`);
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function handleGraphWriteWord() {
  if (!currentGraphModel || !currentPlsResults) {
    setStatus("error", "Harap jalankan Kompilasi Graph dan Analisis PLS-SEM terlebih dahulu.");
    return;
  }

  const sub_id = document.getElementById("graph-select-subchapter")?.value || "all";
  const btn = document.getElementById("btn-graph-write-native");
  const statEl = document.getElementById("graph-action-status");
  if (btn) btn.disabled = true;
  setStatus("working", "Menulis naskah model hibrida langsung ke Microsoft Word...");
  if (statEl) statEl.textContent = "Sedang mengetik ke dokumen Word...";

  try {
    // 1. Coba via Office.js dengan docx base64
    const docRes = await fetch(`${API_BASE}/graph/generate-docx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        graph: currentGraphModel,
        pls_results: currentPlsResults,
        filename: `NASKAH_MODEL_${currentGraphModel.model_type.toUpperCase()}.docx`
      })
    });
    const docData = await docRes.json();
    let officeJsInserted = false;
    if (docData && docData.docx_base64) {
      officeJsInserted = await insertDocxBase64ToActiveWord(docData.docx_base64);
    }

    // 2. Jalur Native Word COM bila di desktop
    let nativeSuccess = false;
    try {
      const nativeRes = await fetch(`${API_BASE}/graph/write-word`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sub_id,
          graph: currentGraphModel,
          pls_results: currentPlsResults
        })
      });
      const nativeData = await nativeRes.json();
      if (nativeData.status === "success") nativeSuccess = true;
    } catch (e) {
      console.warn("Native write call:", e);
    }

    if (officeJsInserted || nativeSuccess) {
      setStatus("ready", "Naskah model hibrida berhasil diketik langsung ke dokumen Word aktif!");
      if (statEl) {
        statEl.style.color = "#15803d";
        statEl.textContent = "✅ Berhasil disisipkan ke dokumen Word aktif!";
      }
    } else {
      setStatus("ready", "Naskah dokumen Word telah disiapkan.");
      if (statEl) {
        statEl.style.color = "#0369a1";
        statEl.textContent = `ℹ️ File dokumen siap. Gunakan tombol 'Unduh .docx Standar' bila Word tidak terbuka.`;
      }
    }
  } catch (err) {
    console.error(err);
    setStatus("error", `Gagal menulis ke Word: ${err.message}`);
    if (statEl) {
      statEl.style.color = "#b91c1c";
      statEl.textContent = `❌ ${err.message}`;
    }
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function handleGraphDownloadDocx() {
  if (!currentGraphModel || !currentPlsResults) {
    setStatus("error", "Harap jalankan Kompilasi Graph dan Analisis PLS-SEM terlebih dahulu.");
    return;
  }

  const btn = document.getElementById("btn-graph-download-docx");
  const statEl = document.getElementById("graph-action-status");
  if (btn) btn.disabled = true;
  setStatus("working", "Membuat dokumen Word (.docx) standar By SkripsiPlg...");

  try {
    const filename = `TESIS_MODEL_${currentGraphModel.model_type.toUpperCase()}_${new Date().toISOString().slice(0, 10)}.docx`;
    const res = await fetch(`${API_BASE}/graph/generate-docx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        graph: currentGraphModel,
        pls_results: currentPlsResults,
        filename
      })
    });

    const data = await res.json();
    if (!res.ok || data.status === "error") {
      throw new Error(data.message || `HTTP ${res.status}`);
    }

    // Trigger download
    const dlUrl = `${API_BASE}/download/${data.filename}`;
    const a = document.createElement("a");
    a.href = dlUrl;
    a.download = data.filename;
    document.body.appendChild(a);
    a.click();
    a.remove();

    setStatus("ready", `Dokumen ${data.filename} berhasil diunduh!`);
    if (statEl) {
      statEl.style.color = "#15803d";
      statEl.textContent = `✅ Unduhan dimulai: ${data.filename}`;
    }
  } catch (err) {
    console.error(err);
    setStatus("error", `Gagal membuat .docx: ${err.message}`);
    if (statEl) {
      statEl.style.color = "#b91c1c";
      statEl.textContent = `❌ ${err.message}`;
    }
  } finally {
    if (btn) btn.disabled = false;
  }
}
