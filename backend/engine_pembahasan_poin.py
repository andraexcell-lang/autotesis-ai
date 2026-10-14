# backend/engine_pembahasan_poin.py
"""
Engine Generator Pembahasan Hasil Bab 4.2 Per Poin Hipotesis (H1 s/d H7).
Formula Baku 4 Alinea Tesis Pascasarjana Magister Manajemen:
- Alinea 1: Bukti Statistik Hipotesis (Beta, t-statistic, p-value, effect size f2, keputusan H diterima)
- Alinea 2: Fakta Empiris & Realitas Lapangan (Kondisi riil DPMPTSP Sumsel, data IKU defisit -16,6%, absensi 2.520 kasus)
- Alinea 3: Justifikasi Teori & Konseptual Ahli (Sutrisno, Mangkunegara, Rivai & Sagala, Afandi, Kaswan, Khaeruman)
- Alinea 4: Komparasi Penelitian Terdahulu / Pemecahan Research Gap (Mendukung & membantah jurnal relevan)
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

import engine_word_native

OUTPUT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

PEMBAHASAN_DATA = {
    1: {
        "hipotesis": "H1",
        "title": "4.2.1. Pengaruh Motivasi Kerja terhadap Employee Engagement pada Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan",
        "short_title": "Poin 1: Motivasi Kerja -> Employee Engagement (H1)",
        "beta": "0,565",
        "t_stat": "9,100",
        "p_val": "0,000",
        "f2": "0,920 (Sangat Besar)",
        "status": "Diterima",
        "alinea_1": (
            "Hasil pengujian hipotesis pertama (H1) membuktikan bahwa Motivasi Kerja secara parsial memiliki "
            "pengaruh yang positif dan signifikan terhadap Employee Engagement pada Dinas Penanaman Modal dan "
            "Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Sumatera Selatan. Pembuktian ini didasarkan pada perolehan "
            "nilai koefisien jalur (Original Sample) sebesar 0,565 yang menunjukkan arah hubungan searah (positif). "
            "Lebih lanjut, nilai T-statistic tercatat sebesar 9,100, yang berada jauh di atas ambang batas kritis 1,96, "
            "serta didukung oleh perolehan p-value sebesar 0,000 yang lebih kecil dari taraf signifikansi 0,05 (p < 0,05). "
            "Selain tingkat signifikansi yang tinggi, evaluasi besaran pengaruh atau effect size (f2) pada jalur ini juga "
            "menunjukkan angka sebesar 0,920 yang termasuk dalam kategori Sangat Besar (f2 > 0,35). Hal ini menegaskan "
            "bahwa semakin tinggi pemenuhan motivasi kerja aparatur, maka semakin kokoh pula keterikatan emosional dan "
            "keterlibatan pegawai terhadap instansi."
        ),
        "alinea_2": (
            "Secara empiris, temuan ini memberikan konfirmasi sekaligus solusi atas fenomena krisis keterikatan aparatur "
            "yang diidentifikasi pada latar belakang masalah. Keluhan pegawai terkait tingginya beban kerja di loket perizinan "
            "operasional yang belum diimbangi dengan apresiasi pimpinan terbukti secara statistik menjadi pemicu memudarnya "
            "rasa kebanggaan (Dedication) dan ketangguhan mental (Vigor). Di lingkungan DPMPTSP Provinsi Sumatera Selatan, "
            "aparatur yang memperoleh kepastian hak fisiologis, jaminan rasa aman, lingkungan sosial yang harmonis, serta ruang "
            "aktualisasi diri terbukti memiliki antusiasme dan fokus kerja (Absorption) yang jauh lebih tinggi saat melayani "
            "masyarakat, sehingga motivasi terbukti menjadi fondasi utama dalam merestorasi keterikatan kerja pegawai."
        ),
        "alinea_3": (
            "Secara teoretis, hasil penelitian ini memperkuat dan memperluas dalil manajemen sumber daya manusia yang dikemukakan "
            "oleh Sutrisno (2022) serta Kasmir (2023). Motivasi kerja pada hakikatnya bertindak sebagai daya dorong fundamental "
            "yang membangkitkan dan mempertahankan energi psikologis seseorang. Karyawan yang termotivasi dengan baik akan "
            "merasa dihargai eksistensinya oleh organisasi. Rasa dihargai inilah yang mentransformasikan dorongan kerja mekanis "
            "menjadi ikatan afektif yang mendalam, sejalan dengan konsep Rivai & Sagala (2022) bahwa keterlibatan pegawai tercapai "
            "ketika individu menyatukan aspek fisik, emosional, dan kognitifnya secara utuh ke dalam peran kerjanya."
        ),
        "alinea_4": (
            "Dari sudut pandang komparasi literatur akademis, pembuktian empiris ini konsisten dan selaras dengan temuan penelitian "
            "Junengsih et al. (2022) serta Firmansyah & Rochdianingrum (2025) yang menyimpulkan bahwa pemberian gaji yang layak "
            "dan pengakuan kerja terbukti nyata meningkatkan keterikatan emosional dan dedikasi pegawai di lingkungan kerja. "
            "Hasil riset ini juga memperkuat temuan Indah et al. (2022) serta Prabowo & Kui (2025) yang menempatkan motivasi kerja "
            "sebagai variabel anteseden utama yang paling berpengaruh dalam menumbuhkan Employee Engagement pada organisasi publik "
            "maupun swasta."
        )
    },
    2: {
        "hipotesis": "H2",
        "title": "4.2.2. Pengaruh Disiplin Kerja terhadap Employee Engagement pada Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan",
        "short_title": "Poin 2: Disiplin Kerja -> Employee Engagement (H2)",
        "beta": "0,459",
        "t_stat": "6,318",
        "p_val": "0,000",
        "f2": "0,606 (Besar)",
        "status": "Diterima",
        "alinea_1": (
            "Pembuktian hipotesis kedua (H2) mengonfirmasi bahwa Disiplin Kerja memiliki pengaruh yang positif dan signifikan "
            "terhadap pembentukan Employee Engagement di lingkungan Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu "
            "(DPMPTSP) Provinsi Sumatera Selatan. Kesimpulan ini ditarik berdasarkan output uji Bootstrapping SmartPLS yang "
            "menghasilkan nilai koefisien jalur (Original Sample) sebesar 0,459 dengan arah positif. Tingkat signifikansi jalur "
            "ini sangat meyakinkan dengan nilai T-statistic sebesar 6,318 (jauh di atas batas kritis 1,96) serta nilai p-value "
            "sebesar 0,000 (p < 0,05). Analisis effect size (f2) mencatatkan angka sebesar 0,606 yang masuk dalam klasifikasi "
            "pengaruh Besar (f2 > 0,35), mengindikasikan bahwa pembiasaan disiplin kerja memiliki kontribusi substansial dalam "
            "membangun dedikasi aparatur."
        ),
        "alinea_2": (
            "Kondisi objektif di lokasi penelitian memperlihatkan urgensi nyata dari diterimanya hipotesis ini. Data rekapitulasi "
            "absensi tahun 2025 yang mencatatkan 2.520 kasus keterlambatan datang, 1.657 kasus pulang awal, dan 375 hari alpa "
            "membuktikan bahwa kerapuhan disiplin berkorelasi langsung dengan rendahnya keterikatan aparatur. Ketika instansi "
            "menegakkan kepatuhan jam kerja, pemanfaatan jam efektif, dan ketaatan SOP secara adil, aparatur merasakan kepastian "
            "lingkungan kerja yang tertib. Keteraturan ini mencegah keterpecahan konsentrasi kerja (Absorption) dan membangun "
            "ketahanan mental (Vigor) saat pegawai menghadapi tingginya volume pemohon perizinan di jam-jam sibuk operasional."
        ),
        "alinea_3": (
            "Secara konseptual, temuan ini berpijak kokoh pada postulat Mangkunegara (2022) dan Wibowo (2023). Disiplin kerja "
            "merupakan kondisi yang terbentuk melalui serangkaian perilaku ketaatan, kepatuhan, dan ketertiban terhadap tatanan "
            "organisasi. Penerapan aturan yang konsisten menciptakan rasa aman psikologis dan kejelasan batas peran (role clarity). "
            "Pegawai yang terbiasa disiplin tidak memandang aturan sebagai beban paksaan, melainkan sebagai pedoman integritas "
            "yang menumbuhkan rasa bangga dan dedikasi (Dedication) terhadap institusinya, sebagaimana dirumuskan oleh Kaswan (2022)."
        ),
        "alinea_4": (
            "Temuan empiris ini secara tegas memecahkan perdebatan penelitian (Research Gap) dengan membantah temuan Simatupang & "
            "Safitri (2023) yang sebelumnya menyatakan bahwa disiplin kerja tidak berpengaruh signifikan terhadap Employee Engagement. "
            "Sebaliknya, hasil riset ini memperkuat dan sejalan dengan temuan Irmawati & Rudini (2024), Alamsyah et al. (2025), "
            "serta Ridwan et al. (2025), yang secara konsisten membuktikan bahwa penegakan disiplin kerja yang konsisten dan "
            "terstruktur merupakan stimulus utama yang memperkuat partisipasi aktif dan keterlibatan kerja pegawai di sektor pemerintahan."
        )
    },
    3: {
        "hipotesis": "H3",
        "title": "4.2.3. Pengaruh Motivasi Kerja terhadap Kinerja Pegawai Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan",
        "short_title": "Poin 3: Motivasi Kerja -> Kinerja Pegawai (H3)",
        "beta": "0,191",
        "t_stat": "3,271",
        "p_val": "0,001",
        "f2": "0,104 (Kecil)",
        "status": "Diterima",
        "alinea_1": (
            "Evaluasi terhadap pengujian hipotesis ketiga (H3) mendemonstrasikan bahwa Motivasi Kerja secara parsial memiliki "
            "pengaruh langsung yang positif dan signifikan terhadap Kinerja Pegawai pada instansi DPMPTSP Provinsi Sumatera Selatan. "
            "Validitas temuan ini ditunjukkan oleh perolehan nilai koefisien jalur (Original Sample) sebesar 0,191 dengan nilai "
            "T-statistic sebesar 3,271 (> 1,96) dan nilai p-value 0,001 (p < 0,05). Analisis effect size (f2) menghasilkan angka "
            "0,104 yang berada pada kategori Kecil (0,02 <= f2 < 0,15). Hal ini membawa makna bahwa meskipun dorongan motivasi "
            "mampu meningkatkan kinerja secara langsung, dampak langsungnya tidak sebesar variabel lainnya, sehingga memerlukan "
            "faktor perantara untuk mengoptimalkan luaran kinerja pegawai."
        ),
        "alinea_2": (
            "Kenyataan di lingkungan DPMPTSP Provinsi Sumatera Selatan membuktikan bahwa motivasi memegang peranan krusial dalam "
            "memulihkan kesenjangan capaian Indikator Kinerja Utama (IKU) penerbitan izin tepat waktu yang sempat merosot hingga "
            "mencapai -16,6% pada tahun 2025. Pemberian insentif yang proporsional, kecukupan waktu istirahat bagi petugas loket, "
            "serta apresiasi verbal atas penyelesaian dokumen perizinan yang bebas revisi terbukti memicu peningkatan volume "
            "pekerjaan dan ketepatan waktu penyelesaian tugas rutin harian aparatur sipil negara di instansi tersebut."
        ),
        "alinea_3": (
            "Secara teoretis, hasil pengujian ini selaras dengan teori kebutuhan Herzberg (Two-Factor Theory) dan konsep Sutrisno (2022). "
            "Motivasi kerja berfungsi sebagai energi penggerak fisik dan kognitif yang memampukan pegawai mengarahkan fokusnya "
            "pada pencapaian target. Sejalan dengan Afandi (2023) dan Kasmir (2023), pemenuhan faktor pendorong (motivators) seperti "
            "peluang pengembangan karier dan pengakuan prestasi akan secara otomatis meningkatkan produktivitas serta kualitas "
            "pekerjaan yang dihasilkan oleh pegawai."
        ),
        "alinea_4": (
            "Hasil penelitian ini memberikan kontribusi penting dalam memecahkan Research Gap akademis. Temuan ini secara tegas "
            "menolak pandangan Suwanto (2020) serta Hidayat dan Tjahjono (2022) yang menyatakan bahwa motivasi kerja tidak berdampak "
            "signifikan terhadap kinerja sektor publik. Sebaliknya, hasil ini memperkuat temuan Noviantoro et al. (2022), Adhelina et al. "
            "(2024), serta Wulandari et al. (2024), yang membuktikan bahwa motivasi kerja merupakan faktor determinan positif yang "
            "nyata dalam mendongkrak capaian kinerja pelayanan pegawai."
        )
    },
    4: {
        "hipotesis": "H4",
        "title": "4.2.4. Pengaruh Disiplin Kerja terhadap Kinerja Pegawai Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan",
        "short_title": "Poin 4: Disiplin Kerja -> Kinerja Pegawai (H4)",
        "beta": "0,295",
        "t_stat": "4,544",
        "p_val": "0,000",
        "f2": "0,294 (Sedang)",
        "status": "Diterima",
        "alinea_1": (
            "Penelusuran terhadap hasil pengujian hipotesis keempat (H4) memberikan bukti empiris bahwa Disiplin Kerja memiliki "
            "pengaruh langsung yang positif dan signifikan terhadap Kinerja Pegawai di DPMPTSP Provinsi Sumatera Selatan. Nilai "
            "koefisien jalur (Original Sample) yang diperoleh adalah sebesar 0,295 dengan signifikansi T-statistic sebesar 4,544 "
            "(> 1,96) dan p-value sebesar 0,000 (p < 0,05). Analisis effect size (f2) mencatatkan nilai sebesar 0,294 yang masuk "
            "dalam kategori pengaruh Sedang (0,15 <= f2 < 0,35). Hal ini membuktikan bahwa penegakan kepatuhan kerja berkontribusi "
            "nyata secara langsung terhadap peningkatan kuantitas dan kualitas output aparatur."
        ),
        "alinea_2": (
            "Di lapangan, keterkaitan antara disiplin dan kinerja sangat nyata tercermin dari proses verifikasi izin operasional. "
            "Tingginya angka ketidakhadiran fisik pada periode sebelumnya secara langsung memperlambat rantai koordinasi berkas. "
            "Ketika pegawai disiplin hadir tepat waktu dan mematuhi SOP teknis, antrean penerbitan izin investasi dapat diproses "
            "secara cepat tanpa penundaan. Kepatuhan terhadap jam kerja efektif meminimalisasi waktu tunggu pemohon dan menekan angka "
            "komplain masyarakat terhadap pelayanan perizinan di DPMPTSP Sumsel."
        ),
        "alinea_3": (
            "Landasan konseptual yang menopang temuan ini mengonfirmasi teori Rivai & Sagala (2022) serta Mangkunegara (2022). "
            "Disiplin kerja merupakan instrumen manajerial fundamental yang menjamin kepatuhan terhadap standar operasional organisasi. "
            "Ketika pegawai memiliki kesadaran tinggi dalam mematuhi prosedur dan jadwal kerja, inefisiensi birokrasi dan kesalahan "
            "kerja (human error) dapat ditekan hingga titik terendah, yang pada gilirannya berkonversi menjadi pencapaian standar kinerja "
            "yang unggul dan akuntabel."
        ),
        "alinea_4": (
            "Temuan ini sekaligus menyelesaikan Research Gap dengan membantah kesimpulan riset Simatupang dan Safitri (2023) serta "
            "Putra (2021) yang menyatakan bahwa penegakan disiplin tidak berdampak pada kinerja. Sebaliknya, penelitian ini secara "
            "konsisten mendukung temuan Adhelina et al. (2024), Ridwan et al. (2025), Noviantoro et al. (2022), serta Oktarisa et al. "
            "(2026), yang membuktikan bahwa disiplin kerja merupakan pilar utama penentu efisiensi dan tingginya kinerja aparatur pemerintah."
        )
    },
    5: {
        "hipotesis": "H5",
        "title": "4.2.5. Pengaruh Employee Engagement terhadap Kinerja Pegawai Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan",
        "short_title": "Poin 5: Employee Engagement -> Kinerja Pegawai (H5)",
        "beta": "0,563",
        "t_stat": "7,018",
        "p_val": "0,000",
        "f2": "0,599 (Besar)",
        "status": "Diterima",
        "alinea_1": (
            "Analisis pada jalur pengujian hipotesis kelima (H5) membuahkan fakta empiris bahwa Employee Engagement memiliki pengaruh "
            "positif dan sangat signifikan terhadap Kinerja Pegawai pada DPMPTSP Provinsi Sumatera Selatan. Keabsahan ini dibuktikan "
            "oleh koefisien jalur (Original Sample) yang sangat dominan sebesar 0,563, didukung oleh nilai T-statistic sebesar 7,018 "
            "(> 1,96) dan p-value 0,000 (p < 0,05). Menariknya, effect size (f2) tercatat sebesar 0,599 yang tergolong Sangat Besar "
            "(f2 > 0,35). Angka ini menempatkan Employee Engagement sebagai prediktor langsung paling kuat dan paling dominan dalam "
            "menentukan keberhasilan kinerja pegawai, melampaui efek langsung motivasi (0,191) maupun disiplin (0,295)."
        ),
        "alinea_2": (
            "Fakta ini mencerminkan fenomena empiris bahwa di era reformasi birokrasi, aparatur yang sekadar patuh secara formal "
            "tidak cukup untuk mencapai target pelayanan prima. Untuk mengejar defisit IKU izin, diperlukan keterikatan afektif "
            "di mana pegawai mencurahkan upaya ekstra (Discretionary Effort). Pegawai yang memiliki energi tinggi (Vigor), rasa "
            "bangga pada instansi (Dedication), dan fokus penuh (Absorption) terbukti tanggap merespons instruksi mendesak, aktif bekerja "
            "sama dalam tim loket perizinan, dan berinisiatif membantu rekan kerja demi terselesaikannya target pelayanan tepat waktu."
        ),
        "alinea_3": (
            "Secara teoretis, hasil ini mengukuhkan postulat Kaswan (2022) dan Afandi (2023). Employee Engagement merupakan kondisi "
            "psikologis positif yang menghubungkan energi individu dengan tujuan organisasi. Pegawai yang engaged memandang pekerjaannya "
            "bukan sekadar rutinitas birokrasi, melainkan panggilan dedikasi sosial. Antusiasme dan rasa kepemilikan yang mendalam ini "
            "mendorong aparatur untuk terus menghasilkan kualitas pekerjaan yang melampaui standar minimal yang ditetapkan organisasi."
        ),
        "alinea_4": (
            "Temuan ini memperkuat literatur manajemen SDM publik dengan mengonfirmasi hasil penelitian Pratama et al. (2025), Azahra & "
            "Nurhasan (2025), serta Susanto & Wijaya (2025). Seluruh studi tersebut secara seragam membuktikan bahwa Employee Engagement "
            "merupakan penggerak inti dari produktivitas kerja pegawai di sektor pemerintahan, dan membantah temuan Alamsyah et al. "
            "(2025) yang sebelumnya menyatakan keterlibatan kerja tidak berpengaruh terhadap kinerja."
        )
    },
    6: {
        "hipotesis": "H6",
        "title": "4.2.6. Pengaruh Motivasi Kerja terhadap Kinerja Pegawai Melalui Employee Engagement Sebagai Variabel Mediasi pada Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan",
        "short_title": "Poin 6: Motivasi Kerja -> EE -> Kinerja Pegawai (H6 Mediasi)",
        "beta": "0,318",
        "t_stat": "5,423",
        "p_val": "0,000",
        "f2": "Mediasi Parsial (Partial Mediation)",
        "status": "Diterima",
        "alinea_1": (
            "Pembuktian terhadap hipotesis keenam (H6) mengungkap temuan empiris yang sangat krusial bahwa Motivasi Kerja berpengaruh "
            "positif dan signifikan terhadap Kinerja Pegawai melalui peran Employee Engagement sebagai variabel mediasi. Hasil kalkulasi "
            "Specific Indirect Effects SmartPLS menghasilkan nilai koefisien tidak langsung (Original Sample) sebesar 0,318 dengan arah "
            "positif, didukung oleh nilai T-statistic sebesar 5,423 (> 1,96) dan p-value sebesar 0,000 (p < 0,05). Karena pengaruh langsung "
            "Motivasi terhadap Kinerja juga terbukti signifikan (p = 0,001), maka sifat mediasi ini dikategorikan sebagai Mediasi Parsial "
            "(Partial Mediation). Hal ini bermakna bahwa pemenuhan motivasi kerja dapat meningkatkan kinerja secara langsung, namun daya "
            "ungkitnya akan meningkat secara berlipat ganda apabila motivasi tersebut berhasil menumbuhkan Employee Engagement terlebih dahulu."
        ),
        "alinea_2": (
            "Kondisi di DPMPTSP Provinsi Sumatera Selatan membuktikan bahwa kebijakan insentif finansial dan fasilitas kerja yang baik "
            "tidak serta-merta menjamin tuntasnya target IKU jika pegawai meresponsnya semata-mata sebagai transaksi administratif. "
            "Namun, ketika motivasi tersebut mampu menyentuh ranah psikologis pegawai—menimbulkan rasa dihargai dan kebanggaan atas "
            "profesinya—maka keterikatan emosional (engagement) yang terbentuk akan memicu aparatur untuk bekerja dengan penuh kehati-hatian, "
            "meminimalisasi kesalahan verifikasi berkas, dan menuntaskan target izin sebelum tenggat waktu berakhir."
        ),
        "alinea_3": (
            "Secara konseptual, fenomena mediasi ini memperkuat teori pemenuhan kebutuhan Maslow dan Herzberg yang dikolaborasikan "
            "dengan teori Employee Engagement Kaswan (2022) serta Khaeruman (2022). Motivasi ekstrinsik dan intrinsik bertindak sebagai "
            "katalisator yang memasok bahan bakar emosional bagi terbentuknya keterikatan kerja. Keterikatan psikologis inilah yang "
            "kemudian menjadi jembatan pengubah antara stimulus motivasi dasar menjadi aksi nyata peningkatan kinerja pelayanan publik."
        ),
        "alinea_4": (
            "Pengujian jalur mediasi ini sukses memberikan jawaban dan pemecahan atas Research Gap yang mencolok pada studi terdahulu. "
            "Temuan ini secara tegas menolak anomali penelitian Dewanata dan Fadila (2025) yang menyatakan keterikatan kerja gagal "
            "memediasi motivasi terhadap kinerja. Sebaliknya, riset ini memberikan konfirmasi kuat terhadap hasil penelitian Prabowo & "
            "Kui (2025), Indah et al. (2022), serta Saputra (2024), yang membuktikan secara empiris bahwa Employee Engagement merupakan "
            "variabel mediator yang andal dan efektif dalam menjembatani pengaruh motivasi terhadap kinerja aparatur."
        )
    },
    7: {
        "hipotesis": "H7",
        "title": "4.2.7. Pengaruh Disiplin Kerja terhadap Kinerja Pegawai Melalui Employee Engagement Sebagai Variabel Mediasi pada Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu Provinsi Sumatera Selatan",
        "short_title": "Poin 7: Disiplin Kerja -> EE -> Kinerja Pegawai (H7 Mediasi)",
        "beta": "0,258",
        "t_stat": "4,379",
        "p_val": "0,000",
        "f2": "Mediasi Parsial (Partial Mediation)",
        "status": "Diterima",
        "alinea_1": (
            "Penyelarasan terhadap hasil pengujian hipotesis ketujuh (H7) membuktikan bahwa Disiplin Kerja berpengaruh positif dan "
            "signifikan terhadap Kinerja Pegawai melalui Employee Engagement sebagai variabel mediasi pada DPMPTSP Provinsi Sumatera "
            "Selatan. Ekstraksi data Specific Indirect Effects melalui prosedur Bootstrapping mencatatkan nilai koefisien tidak langsung "
            "sebesar 0,258 (positif) dengan T-statistic sebesar 4,379 (> 1,96) dan p-value sebesar 0,000 (p < 0,05). Mengingat pengaruh "
            "langsung Disiplin terhadap Kinerja juga signifikan (p = 0,000), maka peran mediasi yang terjalin adalah Mediasi Parsial "
            "(Partial Mediation). Hal ini menegaskan bahwa penerapan disiplin kerja akan mentransformasi kinerja aparatur secara jauh "
            "lebih optimal apabila kepatuhan prosedural tersebut dijiwai oleh keterlibatan emosional dan dedikasi yang mendalam."
        ),
        "alinea_2": (
            "Secara empiris di lapangan, temuan ini menjadi jawaban solutif atas masalah 'kepatuhan semu' (Nuraini, 2023). Tingginya "
            "angka pelanggaran jam kerja di DPMPTSP Sumsel (2.520 kasus keterlambatan) membuktikan bahwa pengawasan administratif yang "
            "hanya mengandalkan ancaman sanksi pemotongan tunjangan hanya menghasilkan pegawai yang hadir secara fisik namun 'absen' secara "
            "mental (Quiet Quitting). Keberadaan Employee Engagement sebagai mediator membuktikan bahwa kedisiplinan baru akan berbuah "
            "kinerja pelayanan prima ketika aparatur merasa terikat secara emosional dengan tugasnya dan sadar bahwa ketepatan waktu "
            "sangat menentukan kepuasan para pelaku usaha di Sumatera Selatan."
        ),
        "alinea_3": (
            "Landasan konseptual dari hubungan mediasi ini sejalan dengan pandangan Mangkunegara (2022) serta Rivai & Sagala (2022). "
            "Penegakan tata tertib pada hakikatnya menciptakan stabilitas ekosistem kerja. Pembiasaan disiplin yang konsisten akan "
            "membentuk ketertiban mental, yang mempermudah tercapainya kondisi konsentrasi penuh (Absorption) dan ketangguhan mental "
            "(Vigor). Keteraturan operasional yang dijiwai oleh antusiasme keterikatan kerja inilah yang pada puncaknya mewujudkan "
            "kelengkapan dokumen, minimnya kesalahan kerja, dan tercapainya target IKU perizinan secara berkelanjutan."
        ),
        "alinea_4": (
            "Sebagai sumbangsih terhadap diskursus akademis, pengujian jalur ini sukses memecahkan Research Gap empiris yang krusial. "
            "Temuan ini membantah hasil riset Alamsyah et al. (2025) dan Ridwan et al. (2025) yang sebelumnya menyatakan bahwa "
            "Employee Engagement tidak mampu memediasi pengaruh disiplin terhadap kinerja. Sebaliknya, hasil ini memperkuat validitas "
            "penelitian Purwaningtyas et al. (2026), Iskandar et al. (2026), serta Saputra (2024), yang membuktikan secara tak terbantahkan "
            "bahwa keterikatan kerja merupakan mediator vital yang mengubah kepatuhan struktural menjadi kinerja pelayanan yang paripurna."
        )
    }
}

def get_pembahasan_poin(poin_num):
    """Mengembalikan data pembahasan untuk poin tertentu (1 s/d 7)."""
    return PEMBAHASAN_DATA.get(int(poin_num), None)

def generate_pembahasan_docx(poin_num, out_filename=None):
    """Membuat file Word .docx mandiri khusus untuk satu poin pembahasan 4.2."""
    data = get_pembahasan_poin(poin_num)
    if not data:
        raise ValueError(f"Poin pembahasan {poin_num} tidak ditemukan (Pilihan: 1-7).")

    if not out_filename:
        out_filename = f"PEMBAHASAN_BAB_4_POIN_{poin_num}_{data['hipotesis']}.docx"
    
    out_path = os.path.join(OUTPUT_DIR, out_filename)
    
    doc = Document()
    # Format A4, Margin 4-4-3-3
    for s in doc.sections:
        s.page_width = Inches(8.27)
        s.page_height = Inches(11.69)
        s.top_margin = Inches(1.575)
        s.left_margin = Inches(1.575)
        s.bottom_margin = Inches(1.181)
        s.right_margin = Inches(1.181)

    # Heading Judul Poin
    p_h = doc.add_paragraph()
    p_h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_h.paragraph_format.line_spacing = 1.5
    p_h.paragraph_format.space_before = Pt(12)
    p_h.paragraph_format.space_after = Pt(6)
    r_h = p_h.add_run(data["title"])
    r_h.bold = True
    r_h.font.name = "Times New Roman"
    r_h.font.size = Pt(12)

    # 4 Alinea Formula Baku
    for alinea_text in [data["alinea_1"], data["alinea_2"], data["alinea_3"], data["alinea_4"]]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.first_line_indent = Inches(0.5)
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(alinea_text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)

    doc.save(out_path)
    return out_filename, out_path

def write_pembahasan_poin_word(poin_num):
    """Mengetikkan teks poin pembahasan langsung ke dokumen Word aktif via Native COM."""
    data = get_pembahasan_poin(poin_num)
    if not data:
        raise ValueError(f"Poin pembahasan {poin_num} tidak valid.")

    full_text = (
        f"{data['title']}\n\n"
        f"    {data['alinea_1']}\n\n"
        f"    {data['alinea_2']}\n\n"
        f"    {data['alinea_3']}\n\n"
        f"    {data['alinea_4']}\n\n"
    )
    return engine_word_native.write_text_to_word(full_text)

if __name__ == "__main__":
    for i in range(1, 8):
        fn, fp = generate_pembahasan_docx(i)
        print(f"Generated Poin {i}: {fn}")
