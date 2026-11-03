#!/usr/bin/env node
/**
 * build_pages.mjs - Menyiapkan folder statis `dist/` untuk Cloudflare Pages
 * (atau GitHub Pages / Netlify). Tanpa dependency eksternal.
 *
 * Pakai:
 *   node deploy/build_pages.mjs
 *   AUTOTESIS_BACKEND_URL=https://autotesis-api.onrender.com node deploy/build_pages.mjs
 *   node deploy/build_pages.mjs --out public --backend https://xxx.onrender.com
 *
 * Hasil di dist/:
 *   taskpane.html, taskpane.css, taskpane.js, config.js (diisi API base),
 *   index.html (redirect ke taskpane.html), _headers.
 *
 * Catatan: backend Python TIDAK ikut ke Cloudflare Pages (Pages hanya statis).
 * Deploy backend ke Render/Fly, lalu isi AUTOTESIS_BACKEND_URL dengan URL-nya.
 */
import { cp, mkdir, readdir, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, "..");
const SRC = path.join(ROOT, "taskpane");

function arg(name, fallback = "") {
  const i = process.argv.indexOf(`--${name}`);
  return i > -1 && process.argv[i + 1] ? process.argv[i + 1] : fallback;
}

const outDir = path.resolve(ROOT, arg("out", "dist"));
const rawBackend = (arg("backend", process.env.AUTOTESIS_BACKEND_URL || "") || "").trim();

function normalizeBase(url) {
  if (!url) return "";
  const clean = url.replace(/\/+$/, "");
  if (!/^https?:\/\//i.test(clean)) {
    throw new Error(`AUTOTESIS_BACKEND_URL harus diawali http(s):// (dapat: "${url}")`);
  }
  if (clean.startsWith("http://") && !/^http:\/\/(localhost|127\.0\.0\.1)/i.test(clean)) {
    console.warn("[warn] backend http:// non-localhost: Word akan memblokir konten campuran. Pakai https://.");
  }
  return clean;
}

const backend = normalizeBase(rawBackend);

// File yang ikut terbit (file .bak / draf tidak boleh bocor ke publik).
const WHITELIST = ["taskpane.html", "taskpane.css", "taskpane.js", "config.js"];

async function main() {
  const entries = await readdir(SRC, { withFileTypes: true });
  const missing = WHITELIST.filter((f) => !entries.some((e) => e.isFile() && e.name === f));
  if (missing.length) throw new Error(`File wajib tidak ditemukan: ${missing.join(", ")}`);

  await rm(outDir, { recursive: true, force: true });
  await mkdir(outDir, { recursive: true });

  for (const f of WHITELIST) {
    await cp(path.join(SRC, f), path.join(outDir, f));
  }
  // Folder aset (ikon) bila ada.
  const assets = entries.find((e) => e.isDirectory() && e.name === "assets");
  if (assets) await cp(path.join(SRC, "assets"), path.join(outDir, "assets"), { recursive: true });

  // Manifest add-in ikut diterbitkan agar bisa diunduh dari URL publik.
  const manifest = path.join(ROOT, "manifest_pages.xml");
  if (await stat(manifest).then((s) => s.isFile(), () => false)) {
    await cp(manifest, path.join(outDir, "manifest.xml"));
  }

  const configJs = `// Digenerate otomatis oleh deploy/build_pages.mjs - jangan diedit manual.\nwindow.AUTOTESIS_API_BASE = ${JSON.stringify(backend)};\n`;
  await writeFile(path.join(outDir, "config.js"), configJs, "utf8");

  await writeFile(
    path.join(outDir, "index.html"),
    `<!DOCTYPE html><html lang="id"><head><meta charset="utf-8">` +
      `<meta http-equiv="refresh" content="0; url=taskpane.html">` +
      `<title>Independent Researcher</title></head>` +
      `<body style="font-family:system-ui;padding:24px">` +
      `Membuka panel… <a href="taskpane.html">taskpane.html</a></body></html>\n`,
    "utf8"
  );

  await writeFile(
    path.join(outDir, "_headers"),
    [
      "/*",
      "  X-Content-Type-Options: nosniff",
      "  Referrer-Policy: strict-origin-when-cross-origin",
      "",
      "/taskpane.html",
      "  Cache-Control: no-cache",
      "",
      "/config.js",
      "  Cache-Control: no-cache",
      "",
      "/manifest.xml",
      "  Cache-Control: no-cache",
      "",
    ].join("\n"),
    "utf8"
  );

  console.log("Build Pages selesai.");
  console.log(`  Sumber  : ${SRC}`);
  console.log(`  Output  : ${outDir}`);
  console.log(`  Backend : ${backend || "(satu origin / kosong)"}`);
  if (!backend) {
    console.log("  [info] Tanpa backend URL, taskpane akan memanggil API di origin yang sama.");
  }
}

main().catch((e) => {
  console.error("GAGAL:", e.message);
  process.exitCode = 1;
});
