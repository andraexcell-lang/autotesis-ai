# -*- coding: utf-8 -*-
"""make_manifest.py - Membuat manifest add-in Word yang menunjuk ke URL publik.

Manifest Word WAJIB memakai https:// dan URL absolut. Skrip ini menyalin
manifest_online.xml lalu mengganti SEMUA origin lama (trycloudflare / pages.dev /
sebelumnya) dengan origin baru, sehingga ikon, SourceLocation, AppDomain, dan
Taskpane.Url konsisten.

Contoh:
    python deploy/make_manifest.py --base-url https://autotesis.pages.dev
    python deploy/make_manifest.py --base-url https://autotesis.pages.dev --out manifest_online.xml

Tanpa --out, hasil ditulis ke manifest_<host>.xml agar manifest lama tidak
tertimpa tanpa sengaja.
"""
import argparse
import re
import sys
import urllib.parse
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "manifest_online.xml"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base-url", required=True,
                    help="Origin publik, mis. https://autotesis.pages.dev")
    ap.add_argument("--template", default=str(TEMPLATE))
    ap.add_argument("--out", default="")
    args = ap.parse_args()

    raw = args.base_url.strip().rstrip("/")
    parsed = urllib.parse.urlparse(raw)
    if parsed.scheme != "https":
        sys.exit("GAGAL: manifest add-in Word mewajibkan URL https:// (juga untuk uji sideload).")
    if not parsed.netloc:
        sys.exit(f"GAGAL: URL tidak valid: {raw}")
    base = f"https://{parsed.netloc}"

    tpl_path = Path(args.template)
    if not tpl_path.exists():
        sys.exit(f"GAGAL: template tidak ditemukan: {tpl_path}")
    xml = tpl_path.read_text(encoding="utf-8")

    m = re.search(r'<SourceLocation DefaultValue="(https://[^/"]+)', xml)
    if not m:
        sys.exit("GAGAL: tidak menemukan <SourceLocation> https:// di template.")
    old_origin = m.group(1)

    if old_origin == base:
        print(f"Manifest sudah menunjuk ke {base} (tidak ada perubahan).")
    else:
        xml = xml.replace(old_origin, base)
        print(f"Origin diganti: {old_origin} -> {base}")

    out = Path(args.out) if args.out else ROOT / f"manifest_{parsed.netloc.replace('.', '_')}.xml"
    out.write_text(xml, encoding="utf-8")

    sisa = sorted(set(re.findall(r"https://[a-z0-9.-]+", xml)))
    print(f"Ditulis: {out}")
    print("Origin di manifest sekarang:", ", ".join(sisa))
    print("Langkah berikutnya: sisipkan manifest via Word > File > Options > "
          "Trust Center > Trusted Add-in Catalogs (atau upload ke M365 admin center).")


if __name__ == "__main__":
    main()
