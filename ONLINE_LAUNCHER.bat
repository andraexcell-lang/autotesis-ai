@echo off
title INDEPENDENT RESEARCHER - CLOUDFLARE ONLINE LAUNCHER
color 0b

echo =====================================================================
echo         INDEPENDENT RESEARCHER - CLOUDFLARE ONLINE TUNNEL
echo       Standar Tesis Magister Manajemen By SkripsiPlg & SmartPLS 4
echo =====================================================================
echo.

cd /d "%~dp0"

echo [1/2] Memastikan Server Python Lokal Aktif di Port 5050...
powershell -Command "$c = Get-NetTCPConnection -LocalPort 5050 -ErrorAction SilentlyContinue; if (!$c) { Start-Process python -ArgumentList 'backend\server.py' -WindowStyle Hidden; Start-Sleep -Seconds 2 }"

echo.
echo [2/2] Mengaktifkan Cloudflare HTTPS Tunnel ke Seluruh Dunia...
echo Jendela ini menjaga tunnel tetap ONLINE. Jangan ditutup selama ingin digunakan.
echo =====================================================================
python run_tunnel.py

pause
