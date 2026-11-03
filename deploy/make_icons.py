# -*- coding: utf-8 -*-
"""make_icons.py - Membuat ikon add-in (16/32/64/80 px) tanpa dependency.

Output: taskpane/assets/icon-{16,32,64,80}.png
Desain: kotak membulat biru (gradien) + dokumen putih dengan garis teks,
dirender 4x lalu diperkecil (supersampling) agar tepi halus.

Jalankan: python deploy/make_icons.py
"""
import os
import struct
import zlib

OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "taskpane", "assets")
SS = 4  # supersampling factor


def mix(c1, c2, t):
    return tuple(int(round(a + (b - a) * t)) for a, b in zip(c1, c2))


def rounded_alpha(px, py, w, h, r):
    """1 di dalam kotak membulat, 0 di luar (koordinat piksel)."""
    cx = min(max(px, r), w - r)
    cy = min(max(py, r), h - r)
    d = ((px - cx) ** 2 + (py - cy) ** 2) ** 0.5
    return 1.0 if d <= r else 0.0


def render(size):
    w = size * SS
    bg_top, bg_bottom = (37, 99, 235), (29, 78, 216)   # #2563eb -> #1d4ed8
    white = (255, 255, 255)

    pixels = [[(0, 0, 0, 0)] * w for _ in range(w)]
    radius = w * 0.22

    # Geometri dokumen (relatif terhadap kanvas)
    doc_x0, doc_x1 = w * 0.30, w * 0.70
    doc_y0, doc_y1 = w * 0.22, w * 0.78
    doc_r = w * 0.035
    line_h = max(1.0, w * 0.045)
    lines = [0.36, 0.48, 0.60]

    for y in range(w):
        bg = mix(bg_top, bg_bottom, y / max(1, w - 1))
        for x in range(w):
            a = rounded_alpha(x + 0.5, y + 0.5, w, w, radius)
            if a <= 0:
                continue
            color = bg

            # Dokumen putih
            inside = 0.0
            if doc_x0 <= x < doc_x1 and doc_y0 <= y < doc_y1:
                inside = rounded_alpha(x + 0.5 - doc_x0, y + 0.5 - doc_y0,
                                       doc_x1 - doc_x0, doc_y1 - doc_y0, doc_r)
                if inside:
                    color = white

            # Garis teks (biru) hanya di dalam area dokumen
            if inside and inside_doc(x, y, doc_x0, doc_x1, doc_y0, doc_y1, lines, line_h, w):
                color = bg_bottom

            pixels[y][x] = (color[0], color[1], color[2], 255 if a else 0)

    # Downsample
    out = []
    for y in range(size):
        row = []
        for x in range(size):
            rs = gs = bs = as_ = 0
            for dy in range(SS):
                for dx in range(SS):
                    r, g, b, a = pixels[y * SS + dy][x * SS + dx]
                    rs += r * a
                    gs += g * a
                    bs += b * a
                    as_ += a
            n = SS * SS
            if as_ == 0:
                row.append((0, 0, 0, 0))
            else:
                row.append((rs // as_, gs // as_, bs // as_, as_ // n))
        out.append(row)
    return out


def inside_doc(x, y, x0, x1, y0, y1, lines, line_h, w):
    margin = (x1 - x0) * 0.16
    if x < x0 + margin or x > x1 - margin:
        return False
    for ly in lines:
        cy = y0 + (y1 - y0) * ly
        if cy <= y <= cy + line_h:
            # baris terakhir lebih pendek
            if ly == lines[-1] and x > x1 - margin - (x1 - x0) * 0.22:
                continue
            return True
    return False


def write_png(path, pixels):
    size = len(pixels)
    raw = b""
    for row in pixels:
        raw += b"\x00" + b"".join(bytes(px) for px in row)

    def chunk(kind, data):
        c = struct.pack(">I", len(data)) + kind + data
        return c + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(raw, 9))
    png += chunk(b"IEND", b"")
    with open(path, "wb") as f:
        f.write(png)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    for size in (16, 32, 64, 80):
        path = os.path.join(OUT_DIR, f"icon-{size}.png")
        write_png(path, render(size))
        print(f"{path} ({os.path.getsize(path)} byte)")


if __name__ == "__main__":
    main()
