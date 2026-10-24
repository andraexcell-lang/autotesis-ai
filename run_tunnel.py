import subprocess
import re
import sys
import os
import time

# Ensure utf-8 output on Windows
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CLOUDFLARED_EXE = os.path.join(BASE_DIR, "cloudflared.exe")
URL_FILE = os.path.join(BASE_DIR, "current_tunnel_url.txt")
MANIFEST_TEMPLATE = os.path.join(BASE_DIR, "manifest.xml")
MANIFEST_ONLINE = os.path.join(BASE_DIR, "manifest_online.xml")

def update_manifest_online(online_url):
    try:
        with open(MANIFEST_TEMPLATE, "r", encoding="utf-8") as f:
            content = f.read()
        
        # Replace localhost URLs with online_url
        content_online = content.replace("https://localhost:5051", online_url)
        content_online = content_online.replace("http://localhost:5050", online_url)
        
        with open(MANIFEST_ONLINE, "w", encoding="utf-8") as f:
            f.write(content_online)
        print(f"[MANIFEST] manifest_online.xml diperbarui -> {online_url}")
    except Exception as e:
        print(f"[MANIFEST ERROR] {e}")

if not os.path.exists(CLOUDFLARED_EXE):
    print(f"[ERROR] cloudflared.exe tidak ditemukan di {CLOUDFLARED_EXE}")
    sys.exit(1)

cmd = [CLOUDFLARED_EXE, "tunnel", "--url", "http://127.0.0.1:5050"]
print(f"[TUNNEL] Menjalankan: {' '.join(cmd)}")

proc = subprocess.Popen(
    cmd,
    stdout=subprocess.PIPE,
    stderr=subprocess.STDOUT,
    text=True,
    bufsize=1,
    encoding="utf-8",
    errors="replace"
)

tunnel_url = None
pattern = re.compile(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com")

for line in iter(proc.stdout.readline, ""):
    sys.stdout.write(line)
    sys.stdout.flush()
    if not tunnel_url:
        m = pattern.search(line)
        if m:
            tunnel_url = m.group(0)
            print("=" * 60)
            print(f"[SUCCESS] CLOUDFLARE TUNNEL BERHASIL AKTIF!")
            print(f"URL PUBLIK   : {tunnel_url}")
            print(f"Taskpane UI  : {tunnel_url}/taskpane.html")
            print("=" * 60)
            with open(URL_FILE, "w", encoding="utf-8") as f:
                f.write(tunnel_url)
            update_manifest_online(tunnel_url)
            
            # Tunggu DNS Cloudflare menyebar (biasanya 5-10 detik) sebelum membuka browser
            def open_when_ready(target_url):
                import urllib.request
                import webbrowser
                print("[DNS] Menunggu penyebaran DNS Cloudflare (beberapa detik)...")
                for _ in range(30):
                    time.sleep(2)
                    try:
                        req = urllib.request.Request(f"{target_url}/api/healthz", headers={"User-Agent": "Mozilla/5.0"})
                        with urllib.request.urlopen(req, timeout=5) as res:
                            if res.status == 200:
                                print(f"[READY] DNS & Server Terhubung! Membuka browser: {target_url}/taskpane.html")
                                webbrowser.open(f"{target_url}/taskpane.html")
                                return
                    except Exception:
                        pass
                # Fallback jika timeout
                try:
                    webbrowser.open(f"{target_url}/taskpane.html")
                except Exception:
                    pass

            import threading
            threading.Thread(target=open_when_ready, args=(tunnel_url,), daemon=True).start()

proc.wait()
