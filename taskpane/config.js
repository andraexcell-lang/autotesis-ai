// ============================================================================
// config.js - Konfigurasi runtime taskpane (di-generate saat build, tapi boleh
// diisi manual untuk deployment lokal/uji).
//
// AUTOTESIS_API_BASE:
//   "" (default)  -> frontend & backend satu origin (mis. server.py lokal atau
//                    backend Render yang sekaligus menyajikan taskpane).
//   "https://xxx.onrender.com" -> frontend statis (Cloudflare Pages / GitHub
//                    Pages) memanggil API backend di host lain.
//
// Build script deploy/build_pages.mjs menimpa file ini di folder dist/ memakai
// env AUTOTESIS_BACKEND_URL, jadi file di repo ini TIDAK perlu diubah.
// ============================================================================
window.AUTOTESIS_API_BASE = window.AUTOTESIS_API_BASE || "";
