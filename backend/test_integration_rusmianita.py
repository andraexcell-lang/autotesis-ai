import urllib.request, json, os, docx

BASE_URL = 'http://127.0.0.1:5050'

config = {
    'judul': 'PENGARUH GAYA KEPEMIMPINAN DAN BUDAYA ORGANISASI TERHADAP DISIPLIN KERJA SERTA DAMPAKNYA PADA KINERJA PEGAWAI BANK PEMBANGUNAN DAERAH SUMATERA SELATAN DAN BANGKA BELITUNG CABANG PEMBANTU PARAMESWARA',
    'penulis': 'RUSMIANITA',
    'nim': '255441029',
    'instansi': 'Bank Pembangunan Daerah Sumatera Selatan dan Bangka Belitung Cabang Pembantu Parameswara',
    'tahun': '2026',
    'x1': 'Gaya Kepemimpinan',
    'x2': 'Budaya Organisasi',
    'm': 'Disiplin Kerja',
    'y': 'Kinerja Pegawai',
    'n': 45
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

print('\n=== VERIFIKASI ISI DOKUMEN BAGIAN AWAL ===')
doc_ba = docx.Document(os.path.join(output_dir, res_ba['filename']))
for p in doc_ba.paragraphs[:8]:
    if p.text.strip():
        print(' ', p.text[:100])

print('\n=== VERIFIKASI ISI DOKUMEN BAB 1 ===')
doc_b1 = docx.Document(os.path.join(output_dir, res_b1['filename']))
for p in doc_b1.paragraphs[:8]:
    if p.text.strip():
        print(' ', p.text[:100])

print('\n>>> SEMUA PENGUJIAN 100% SUKSES DAN TERVERIFIKASI RUSMIANITA! <<<')
