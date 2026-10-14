/* AutoTesis AI - Word Taskpane Controller (Dual Native COM & Office.js Engine) */

const API_BASE = (window.location.origin && window.location.origin.startsWith("http")) 
  ? `${window.location.origin}/api` 
  : "http://localhost:5050/api";

console.log("AutoTesis AI initialized. Target API:", API_BASE);

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

  // Chapter buttons
  document.querySelectorAll("[data-action]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const action = btn.dataset.action;
      await handleChapterGeneration(action, btn);
    });
  });

  document.getElementById("btn-generate-data")?.addEventListener("click", handleSmartPLSData);
  document.getElementById("btn-generate-gform")?.addEventListener("click", handleGoogleFormScript);
  document.getElementById("btn-generate-full")?.addEventListener("click", handleFullDocumentGeneration);
  
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
      const dapusRange = searchResults.items[0];
      const dapusPara = dapusRange.paragraphs.getFirst();
      context.load(dapusPara);
      const prevPara = dapusPara.getPreviousOrNullObject();
      context.load(prevPara, "text");
      await context.sync();

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

// -------------------------------------------------------------
// HIERARCHICAL SUB-CHAPTER PIPELINE HANDLERS
// -------------------------------------------------------------

async function handleSubChapterWrite(subId, btn) {
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
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const res = await response.json();

    if (res.status === "success") {
      let inserted = false;
      if (res.docx_base64) {
        inserted = await insertSubchapterWithUnifiedDapus(res.docx_base64, res.dapus_base64);
      }
      setProgress(100);
      if (inserted) {
        setStatus("ready", `Sub-bab ${res.meta.code || subId} (${res.meta.title || ''}) berhasil menyambung di dokumen Word, dan Daftar Pustaka diperbarui di halaman paling bawah!`);
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
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const res = await response.json();

    if (res.status === "success") {
      if (res.details && Array.isArray(res.details) && window.Word && typeof Word.run === "function") {
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

function getConfig() {
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
    n: parseInt(document.getElementById("input-responden")?.value || "55", 10)
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

async function handleResearchOntology(showLoading = true) {
  const btn = document.getElementById("btn-research-ontology");
  const badge = document.getElementById("badge-ontology-status");
  const previewBox = document.getElementById("ontology-preview-box");
  if (!btn) return;

  const config = getConfig();
  if (!config.x1 || !config.y) return;

  if (showLoading) {
    btn.disabled = true;
    btn.textContent = "⏳ Meriset Dimensi & Indikator via Live Academic API...";
    if (badge) badge.textContent = "Sedang Meriset...";
  }

  try {
    const response = await fetch(`${API_BASE}/research-ontology`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
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

    if (response.ok) {
      const data = await response.json();
      if (data.status === "success" && data.ontology) {
        renderOntologyPreview(data.ontology);
        if (badge) {
          badge.textContent = `✅ Terkunci (${data.ontology.metadata.total_indikator} Indikator Likert)`;
          badge.style.background = "#bbf7d0";
          badge.style.color = "#15803d";
        }
      }
    }
  } catch (err) {
    console.warn("Ontology research error:", err);
  } finally {
    if (showLoading) {
      btn.disabled = false;
      btn.textContent = "🔍 Telusuri Dimensi & Indikator via Live Academic API";
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

  let html = `
    <div style="font-weight: bold; color: #15803d; margin-bottom: 6px; font-size: 11px;">
      📌 Blueprint Ontologi Terverifikasi (${onto.metadata.total_indikator} Indikator):
    </div>
    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 6px; margin-bottom: 8px;">
      <div style="padding: 6px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px;">
        <div style="font-weight: 700; color: #0284c7;">X1: ${vars.x1.nama_resmi}</div>
        <div style="color: #64748b; font-size: 10px;">Ahli: ${vars.x1.tokoh_ahli}</div>
        <div style="color: #334155; font-size: 10px; margin-top: 2px;">• ${vars.x1.dimensi.map(d => d.nama).join(', ')}</div>
      </div>
      <div style="padding: 6px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px;">
        <div style="font-weight: 700; color: #0284c7;">X2: ${vars.x2.nama_resmi}</div>
        <div style="color: #64748b; font-size: 10px;">Ahli: ${vars.x2.tokoh_ahli}</div>
        <div style="color: #334155; font-size: 10px; margin-top: 2px;">• ${vars.x2.dimensi.map(d => d.nama).join(', ')}</div>
      </div>
      <div style="padding: 6px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px;">
        <div style="font-weight: 700; color: #d97706;">M: ${vars.m.nama_resmi} (Mediasi)</div>
        <div style="color: #64748b; font-size: 10px;">Ahli: ${vars.m.tokoh_ahli}</div>
        <div style="color: #334155; font-size: 10px; margin-top: 2px;">• ${vars.m.dimensi.map(d => d.nama).join(', ')}</div>
      </div>
      <div style="padding: 6px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px;">
        <div style="font-weight: 700; color: #16a34a;">Y: ${vars.y.nama_resmi} (Terikat)</div>
        <div style="color: #64748b; font-size: 10px;">Ahli: ${vars.y.tokoh_ahli}</div>
        <div style="color: #334155; font-size: 10px; margin-top: 2px;">• ${vars.y.dimensi.map(d => d.nama).join(', ')}</div>
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
    if (config.n) document.getElementById("input-responden").value = config.n;
    
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

// Generate SmartPLS dataset and calculations
async function handleSmartPLSData() {
  const config = getConfig();
  const btn = document.getElementById("btn-generate-data");
  btn.disabled = true;
  setStatus("working", "Mensimulasikan dataset & evaluasi kriteria Hair et al...");
  setProgress(50);

  try {
    const response = await fetch(`${API_BASE}/simulate-smartpls`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(config)
    });

    if (!response.ok) throw new Error(`HTTP Error ${response.status}`);
    const data = await response.json();
    setProgress(100);

    const preview = document.getElementById("smartpls-preview");
    if (preview) preview.classList.remove("hidden");

    const dlBox = document.getElementById("smartpls-download-box");
    const dlLink = document.getElementById("link-download-excel");
    if (dlBox && dlLink && data.excel_file) {
      dlBox.classList.remove("hidden");
      dlLink.href = `${API_BASE}/download/${data.excel_file}`;
      dlLink.textContent = `📥 Unduh Excel: ${data.excel_file}`;
    }

    setStatus("ready", "Data SmartPLS 100% Lolos Uji & File Excel siap diunduh!");
  } catch (error) {
    console.error(error);
    setStatus("error", `Gagal olah data: ${error.message}`);
  } finally {
    setTimeout(() => setProgress(0), 1000);
    btn.disabled = false;
  }
}

// Generate Google Apps Script for Google Form
async function handleGoogleFormScript() {
  const config = getConfig();
  setStatus("working", "Membuat kode Google Apps Script...");
  try {
    const response = await fetch(`${API_BASE}/generate-gform-script`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(config)
    });
    const data = await response.json();
    await navigator.clipboard.writeText(data.script);
    setStatus("ready", "Script Google Form disalin ke Clipboard! Tempel di script.google.com");
    alert("Script Google Form otomatis berhasil disalin ke Clipboard!\n\nCara pakai:\n1. Buka script.google.com\n2. Buat proyek baru lalu Paste (Ctrl+V) kode ini.\n3. Klik Run/Jalankan, kuesioner akan otomatis terbentuk di Google Drive Anda!");
  } catch (error) {
    setStatus("error", `Gagal buat script: ${error.message}`);
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

    if (!response.ok) throw new Error(`HTTP Error ${response.status}`);
    const result = await response.json();
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
