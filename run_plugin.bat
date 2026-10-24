@echo off
title Independent Researcher - Microsoft Word Add-in & Engine Launcher
color 0b

echo =====================================================================
echo       INDEPENDENT RESEARCHER - WORD ADD-IN & THESIS GENERATOR
echo         Standar Tesis Magister Manajemen By SkripsiPlg & SmartPLS 4
echo =====================================================================
echo.

cd /d "%~dp0"

echo [1/3] Menyiapkan Sertifikat SSL Localhost...
python backend\generate_cert.py

echo.
echo [2/3] Menyalakan Independent Researcher Dual Server (HTTP: 5050 ^| HTTPS: 5051)...
start "Independent Researcher Server" /b python backend\server.py

timeout /t 3 >nul

echo.
echo [3/3] Membuka Antarmuka Dashboard Independent Researcher...
start http://localhost:5050/taskpane.html

echo.
echo =====================================================================
echo SERVER AKTIF & SIAP DIGUNAKAN!
echo - Web Dashboard (Browser) : http://localhost:5050/taskpane.html
echo - Secure Taskpane (HTTPS) : https://localhost:5051/taskpane.html
echo - File Manifest Word     : C:\Addins\autotesis_manifest.xml
echo.
echo FITUR UNGGULAN:
echo 1. Tulis Bab ke Word: Pilih 'Native COM' untuk mengetik langsung ke lembar Word aktif!
echo 2. 1 Unduh Lengkap  : Merangkai Cover, Bab 1-5, Tabel SmartPLS, & Dapus ke file .docx.
echo 3. Simulator SmartPLS : Menghasilkan dataset Excel 100% lolos Hair et al.
echo 4. Generator GForm    : Menghasilkan kode Apps Script kuesioner otomatis.
echo =====================================================================
echo Tekan sembarang tombol untuk menutup jendela launcher (server tetap berjalan).
pause
