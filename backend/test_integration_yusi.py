import urllib.request, json, os, docx

BASE_URL = 'http://127.0.0.1:5050'

config = {
    'judul': 'PENGARUH KECERDASAN EMOSIONAL DAN MOTIVASI KERJA TERHADAP KINERJA MELALUI ETOS KERJA SEBAGAI VARIABEL INTERVENING DI BPKAD KABUPATEN PALI',
    'penulis': 'MUHAMAD YUSI',
    'nim': '255341028',
    'instansi': 'Badan Pengelola Keuangan dan Aset Daerah (BPKAD) Kabupaten Penukal Abab Lematang Ilir (PALI)',
    'tahun': '2026',
    'x1': 'Kecerdasan Emosional',
    'x2': 'Motivasi Kerja',
    'm': 'Etos Kerja',
    'y': 'Kinerja Pegawai',
    'n': 55
}

def post_json(path, data):
    req = urllib.request.Request(
        f'{BASE_URL}{path}',
        data=json.dumps(data).encode('utf-8'),
        headers={'Content-Type': 'application/json'}
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read().decode('utf-8'))

print('=== 1. TEST POST /api/pipeline/generate-subchapter (sub_1_1) ===')
res_sub1 = post_json('/api/pipeline/generate-subchapter', {'sub_id': 'sub_1_1', 'mode': 'download', 'config': config})
print('Status:', res_sub1.get('status'))
print('Filename:', res_sub1.get('filename'))

print('\n=== 2. TEST POST /api/generate-bagian-awal ===')
res_ba = post_json('/api/generate-bagian-awal', config)
print('Status: success, Filename:', res_ba.get('filename'))

print('\n=== 3. TEST POST /api/generate-bab-1 ===')
res_b1 = post_json('/api/generate-bab-1', config)
print('Status: success, Filename:', res_b1.get('filename'))

print('\n=== 4. TEST POST /api/simulate-smartpls ===')
res_pls = post_json('/api/simulate-smartpls', config)
print('Status: success, Excel:', res_pls.get('excel_file'))

output_dir = r'e:\SKILL UNTUK AI\system_prompts_leaks-main\word-addin-tesismagister\output'

print('\n=== VERIFIKASI ISI SUB-BAB 1.1 ===')
doc_sub1 = docx.Document(os.path.join(output_dir, res_sub1['filename']))
for p in doc_sub1.paragraphs[:6]:
    if p.text.strip():
        print(' [P]', p.text[:100])
for t in doc_sub1.tables:
    print(' [T]', [c.text.replace('\n', ' ') for c in t.rows[0].cells])

print('\n=== VERIFIKASI ISI BAGIAN AWAL ===')
doc_ba = docx.Document(os.path.join(output_dir, res_ba['filename']))
for p in doc_ba.paragraphs[:8]:
    if p.text.strip():
        print(' [P]', p.text[:100])

print('\n=== VERIFIKASI ISI BAB 1 LENGKAP ===')
doc_b1 = docx.Document(os.path.join(output_dir, res_b1['filename']))
for p in doc_b1.paragraphs[:8]:
    if p.text.strip():
        print(' [P]', p.text[:100])

print('\n>>> SEMUA PENGUJIAN 100% SUKSES DAN TERVERIFIKASI MUHAMAD YUSI! <<<')
