@echo off
title AUTOTESIS AI - Server Lokal + Cloudflare Named Tunnel
color 0b

echo =====================================================================
echo   AUTOTESIS AI - BACKEND PUBLIK (Named Tunnel)
echo   https://autotesis.skripsiplg.my.id  ->  http://localhost:5050
echo =====================================================================
echo.

cd /d "%~dp0"

echo [1/2] Memastikan server Python aktif di port 5050...
powershell -NoProfile -Command "if (-not (Get-NetTCPConnection -LocalPort 5050 -State Listen -ErrorAction SilentlyContinue)) { Start-Process python -ArgumentList '-X','utf8','backend\server.py' -WorkingDirectory (Get-Location) -WindowStyle Hidden; Start-Sleep -Seconds 3 }"
echo      OK.

echo.
echo [2/2] Menjalankan Cloudflare Tunnel (autotesis-api)...
echo      Jendela ini harus tetap terbuka. Tutup = backend publik mati.
echo      Frontend tetap online di: https://autotesis-ai.pages.dev
echo =====================================================================
.\cloudflared.exe tunnel --config cloudflared_autotesis.yml run autotesis-api

pause
