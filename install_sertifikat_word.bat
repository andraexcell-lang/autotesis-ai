@echo off
title AutoTesis AI - Pemasangan Sertifikat Keamanan Microsoft Word Add-in
color 0a

echo =====================================================================
echo    PEMASANGAN SERTIFIKAT KEAMANAN RESMI OFFICE ADD-IN UNTUK WORD
echo =====================================================================
echo.
echo Menambahkan sertifikat pengembang ke Windows Certificate Store...
echo.
echo PENTING: Jika muncul jendela peringatan Windows "Security Warning",
echo          silakan klik "YA" (YES) untuk mengizinkan sertifikat.
echo.

certutil -user -addstore -f "Root" "%~dp0backend\ca.cer"
certutil -user -addstore -f "TrustedPeople" "%~dp0backend\localhost.cer"

echo.
echo =====================================================================
echo SERTIFIKAT SELESAI DIPASANG!
echo.
echo Langkah selanjutnya:
echo 1. Kembali ke Microsoft Word Anda.
echo 2. Pada panel kesalahan berwarna kuning, klik tombol [Mulai ulang]
echo    (atau klik kembali tombol "Buka Asisten Tesis" di pita Ribbon).
echo 3. Panel AutoTesis Magister AI akan langsung terbuka sempurna!
echo =====================================================================
echo.
pause
