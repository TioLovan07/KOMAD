# KOMAD — Konversi Matriks Menjadi Draft

Versi 1.3.0. Aplikasi untuk menyusun draft Word dari matriks hasil harmonisasi, memeriksa catatan beserta lokasi teksnya, lalu mengunduh hasil `.docx`.

**Mulai dari [PANDUAN_GITHUB.md](PANDUAN_GITHUB.md)** untuk menayangkan aplikasi. Seluruh situs siap unggah berada di folder `docs/`; tidak perlu memasang dependensi atau menjalankan proses build untuk unggahan pertama paket ini.

## Pemakaian oleh perancang

1. Buka alamat KOMAD yang dibagikan pengelola.
2. Tekan **Pilih File Matriks Word** dan pilih matriks `.docx`.
3. Periksa seluruh catatan. Tekan **Lihat Bagian** untuk menuju teks. Arahkan penunjuk, klik, ketuk, atau gunakan keyboard pada teks berwarna untuk membaca keterangannya.
4. Untuk rekomendasi salah ketik, pilih **Terapkan Rekomendasi** atau **Biarkan Sesuai Matriks**. Catatan struktur atau isi yang meragukan perlu diperbaiki pada matriks, kemudian unggah melalui **Pilih Matriks Revisi**.
5. Setelah semua catatan ditangani dan pratinjau selesai, lanjutkan ke **Unduh Word**. Lengkapi nomor, tanggal, nama, dan isian lain yang belum tersedia.

**Panduan**, **Lihat Contoh**, dan **Tentang penggunaan offline** tersedia di aplikasi. Matriks contoh berisi data ilustrasi; gunakan tombol **Gunakan Matriks Contoh** untuk mencoba seluruh alur tanpa dokumen kantor.

Ukuran tulisan antarmuka dan zoom draft dapat diatur sendiri-sendiri dari 25% sampai 200%. Keduanya tidak mengubah format Word.

## Aturan penyusunan

- Seluruh isi kolom hasil digunakan, termasuk teks hitam dan merah. Warna merah bukan satu-satunya isi yang diambil.
- Kolom hasil kosong atau hanya bertuliskan `Tetap` menggunakan rumusan awal. `Tetap` yang disisipkan pada butir dicocokkan dengan penanda bagian yang sesuai; pencocokan yang gagal menjadi catatan penghalang unduhan.
- Coretan, penghapusan terlacak, dan teks tersembunyi dikeluarkan dari rumusan yang disusun. Teks sumber dan berkas matriks asli tidak ditulis ulang.
- Penomoran Word, penggabungan kolom, dan beberapa variasi judul kolom dikenali. Variasi yang tidak dapat dikenali dihentikan dengan pesan pemeriksaan.
- Pemindahan pasal dari keterangan ditangani apabila pola yang dikenali cocok secara jelas; ini bukan penafsir umum semua instruksi dalam kolom keterangan.
- Lampiran dengan kolom hasil hanya berisi judul memerlukan keputusan pengguna sebelum isi awalnya disertakan.
- Koreksi kata yang terdaftar diterapkan hanya setelah pengguna memilihnya. Ejaan dalam isi matriks tidak diubah hanya karena antarmuka menggunakan istilah “draft”.

Format keluaran mengikuti profil contoh final yang diintegrasikan: Bookman Old Style 12, F4, margin 2 cm, tanpa bold. Profil dapat ditinjau di `docs/js/format-profile.js`. Tabel lampiran mempertahankan struktur tabel sumber. Pratinjau dirender dari **Blob Word yang sama** dengan unduhan; pemenggalan halaman dapat berbeda di Microsoft Word karena font dan mesin tata letak. Ini tidak menjamin semua bentuk matriks dapat diproses tanpa pemeriksaan.

## Offline dan data dokumen

Aplikasi merupakan situs statis dengan PWA. Pembacaan matriks dan penyusunan Word berlangsung di browser perangkat; tidak memanggil layanan pemrosesan dokumen atau membutuhkan akun GPT. Library JavaScript, ikon, dan contoh matriks sudah disertakan, tanpa CDN.

Buka melalui HTTPS saat pertama kali menggunakan. Tunggu status **siap offline**, kemudian pasang melalui kontrol aplikasi/browser yang tersedia. Uji sekali dengan memutus internet dan membuka kembali aplikasi. Pemasangan saja belum berarti semua aset selesai disimpan. Dukungan dan menu pemasangan bergantung pada browser/OS.

Cache hanya berisi berkas aplikasi dan contoh publik, bukan matriks kantor atau hasil draft. Dokumen kerja berada dalam memori tab; unduh hasil sebelum menutup atau memuat ulang. Menghapus data situs/browser dapat menghapus kesiapan offline. Pembaruan memerlukan internet; pembaruan ditunda ketika dokumen masih dibuka agar pekerjaan tidak terhapus oleh pemuatan ulang.

## Struktur kode

| Lokasi | Tanggung jawab |
| --- | --- |
| `docs/index.html`, `style.css` | Struktur dan tampilan antarmuka |
| `docs/js/app.js` | Alur tiga tahap, pemilihan berkas, dialog bantuan |
| `docs/js/xml.js` | Utilitas XML/DOM dan pembacaan teks Word |
| `docs/js/matrix-reader.js` | Kolom logis, numbering Word, pencocokan `Tetap` |
| `docs/js/matrix-parser.js` | Model matriks, aturan pemilihan, pemeriksaan struktur |
| `docs/js/state.js`, `draft-model.js` | Model aktif, keputusan pengguna, rumusan efektif |
| `docs/js/issue-locations.js` | Lokasi, paragraf, kalimat, dan kutipan catatan |
| `docs/js/review-panel.js`, `issue-highlights.js` | Kartu catatan dan pengaitan sorotan |
| `docs/js/word-builder.js`, `format-profile.js` | Penyusunan OOXML dan profil format final |
| `docs/js/preview.js`, `docs/viewer.*` | Pratinjau dari berkas Word hasil konversi |
| `docs/js/downloads.js` | Unduhan dari revisi yang sudah tampil di pratinjau |
| `docs/pwa.js`, `pwa-status.js`, `sw.js` | Pemasangan, status offline, cache dan pembaruan |
| `scripts/build-pwa.py` | Membuat versi cache berdasarkan isi aset |
| `tests/` | Uji aturan pembacaan, lokasi catatan, dan service worker |
| `docs/vendor/` | Dependensi lokal beserta lisensi |

Tidak ada framework atau bundler. JavaScript aplikasi menggunakan ES modules. Vendor tetap dipisahkan dari kode aplikasi; perubahan pada renderer pihak ketiga dicatat di `THIRD_PARTY.md`.

## Pengembangan dan debugging

Untuk menjalankan secara lokal, gunakan Python 3:

```sh
python -m http.server 8080 --directory docs
```

Buka `http://localhost:8080`. Jangan membuka `index.html` langsung dari ZIP atau melalui `file://`, karena modul dan service worker membutuhkan server.

Untuk uji otomatis, gunakan Node.js 18 atau lebih baru:

```sh
npm test
```

Perintah tersebut tidak memerlukan `npm install`. Untuk setiap perubahan aset aplikasi sebelum mengunggah pembaruan, jalankan:

```sh
python scripts/build-pwa.py
npm test
```

Unggah aset yang berubah **beserta `docs/sw.js` yang baru**. Jangan mengedit `sw.js` langsung: perbaikilah `scripts/service-worker.template.js` atau generatornya. Catatan nama/versi aplikasi tampak di `index.html`, `pwa.js`, `downloads.js`, `package.json`, dan generator cache; perbarui bersama saat merilis versi baru.

Saat menemukan bug:

1. Coba matriks contoh untuk membedakan masalah umum dan bentuk matriks tertentu.
2. Catat versi, tahap, baris matriks, catatan, dan langkah reproduksi. Gunakan **Unduh Catatan Pemeriksaan** jika tersedia.
3. Periksa Console untuk error JavaScript dan Network untuk berkas gagal dimuat. Periksa Application → Service Workers/Cache Storage untuk masalah offline. Salin diagnosis dari panel pemasangan bila tersedia.
4. Untuk salah pilih rumusan, mulai dari `matrix-reader.js`/`matrix-parser.js`; untuk rumusan benar tetapi format salah, mulai dari `word-builder.js`/`format-profile.js`.
5. Tambahkan contoh minimal tanpa data kantor ke tes sebelum memperbaiki logika. Jangan menaruh matriks kantor, hasil draft kantor, token, atau kredensial di repositori publik.

Saat development, cache lama dapat membuat perubahan belum terlihat. Tutup dokumen terlebih dahulu, lalu gunakan pembaruan aplikasi. Khusus pengembangan, service worker dapat di-unregister melalui DevTools dan halaman dimuat ulang.

## Validasi paket

Uji otomatis mencakup variasi header/urutan kolom, gridSpan, `Tetap` utuh/campuran, numbering tersembunyi, lokasi penanda berulang, isian berulang, catatan yang hanya ada di sumber, serta cache offline pada subfolder GitHub Pages. Cache parsial tidak boleh dilaporkan siap; dokumen pengguna tidak boleh masuk cache; dokumen aktif menahan pembaruan.

Uji integrasi browser dilakukan pada matriks PAUD (14 pasal), tarif pelelangan ikan (3 pasal), dan contoh bawaan (4 pasal). Berkas Word hasilnya diperiksa struktur XML; hasil PAUD dan tarif juga dirender untuk pemeriksaan tata letak. Pengujian offline penuh pada alamat GitHub HTTPS perlu dilakukan pengelola setelah publikasi, menggunakan daftar periksa di panduan.

Dokumen kantor untuk pengujian tidak disertakan dalam paket/repository. Lihat lisensi komponen pihak ketiga di `THIRD_PARTY.md`.


## Pemeriksaan ejaan 1.1.0
Kamus Hunspell Indonesia (46.905 entri, LGPL-3.0-only) dan nspell 2.1.5 (MIT) disertakan lokal. Lihat docs/assets/dictionary/NOTICE.txt. Istilah tidak dikenali adalah saran pemeriksaan, bukan bukti typo. Saran dibatasi kandidat satu edit, maksimal lima. Nama/singkatan/istilah tetap perlu ditinjau. Tidak menilai substansi hukum.

Versi 1.3.0 menyaring peringatan: kata belum dikenali masuk daftar opsional; hanya pola typo konservatif menjadi catatan utama. Lihat PEMBARUAN-1.1.1.md.


## Rilis gabungan 1.2.0
Lihat PEMBARUAN-1.2.0.md untuk format margin 8 cm seluruh halaman, pemeriksaan pasal lintas baris, validasi dan batas pengujian.


## Rilis 1.3.0
Lihat PEMBARUAN-1.3.0.md untuk aturan halaman pertama/lanjutan, penomoran perubahan, penanganan rumusan belum terpetakan, dan batas lampiran/.doc yang ditunda.
