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


def tidy(xml: str, base: str) -> str:
    """Rapikan manifest: AppDomain unik + Requirements WordApi bila belum ada."""
    # 1) Buang <AppDomain> duplikat (template lama mendaftarkan origin 2x).
    seen = set()

    def keep(match):
        url = match.group(1).strip()
        if url in seen:
            return ""
        seen.add(url)
        return match.group(0)

    xml = re.sub(r"[ \t]*<AppDomain>(.*?)</AppDomain>\s*\n", keep, xml)

    # 2) Tambahkan Requirements WordApi 1.1 setelah </Hosts> tingkat atas.
    if "<Requirements>" not in xml:
        req = ('  <Requirements>\n'
               '    <Sets DefaultMinVersion="1.1">\n'
               '      <Set Name="WordApi" MinVersion="1.1" />\n'
               '    </Sets>\n'
               '  </Requirements>\n\n')
        xml = xml.replace("</Hosts>\n", "</Hosts>\n\n" + req, 1)

    return xml


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

    xml = tidy(xml, base)

    out = Path(args.out) if args.out else ROOT / f"manifest_{parsed.netloc.replace('.', '_')}.xml"
    out.write_text(xml, encoding="utf-8")

    sisa = sorted(set(re.findall(r"https://[a-z0-9.-]+", xml)))
    print(f"Ditulis: {out}")
    print("Origin di manifest sekarang:", ", ".join(sisa))
    print("Langkah berikutnya: sisipkan manifest via Word > File > Options > "
          "Trust Center > Trusted Add-in Catalogs (atau upload ke M365 admin center).")


if __name__ == "__main__":
    main()
