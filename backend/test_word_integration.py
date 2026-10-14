import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from engine_word_native import get_or_create_word
from engine_hierarchical_pipeline import write_subchapter_native

print("Testing Word connection...")
word = get_or_create_word(visible=True)
doc = word.ActiveDocument
print(f"Word active document: {doc.Name}")

cfg = {
    "judul": "PENGARUH KECERDASAN EMOSIONAL DAN MOTIVASI KERJA TERHADAP KINERJA MELALUI ETOS KERJA SEBAGAI VARIABEL INTERVENING DI BPKAD KABUPATEN PALI",
    "penulis": "MUHAMAD YUSI",
    "npm": "255341028",
    "instansi": "BPKAD KABUPATEN PALI",
    "x1": "Kecerdasan Emosional",
    "x2": "Motivasi Kerja",
    "m": "Etos Kerja",
    "y": "Kinerja Pegawai"
}

print("Testing write_subchapter_native for sub_1_1...")
ok = write_subchapter_native("sub_1_1", cfg)
print(f"Success! Inserted sub_1_1 into Word active document: {ok}")
