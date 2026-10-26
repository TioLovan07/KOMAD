# Menayangkan KOMAD melalui GitHub Pages

Panduan untuk pengelola. Perancang cukup menerima alamat aplikasi, membukanya, lalu memasangnya jika diperlukan. Mereka tidak perlu mengunduh ZIP kode.

## 1. Siapkan isi paket

Ekstrak `KOMAD-v1.0.0-GitHub.zip` di komputer Anda. Di dalamnya terdapat `README.md`, panduan ini, `docs/`, `scripts/`, `tests/`, dan `package.json`.

Unggah **isi hasil ekstraksi**, bukan berkas ZIP. Pastikan folder `docs/` langsung berada di tingkat pertama repositori; jangan menambah folder pembungkus seperti `komad-release/docs/`.

## 2. Buat repositori

Masuk ke GitHub dan buat repositori, misalnya bernama `komad`. Untuk GitHub Pages dengan akun Free, gunakan repositori **Public**. Kode, logo, serta contoh yang diunggah akan dapat dilihat publik. Jangan sertakan matriks kantor, hasil draft kantor, data pribadi, atau kata sandi.

Aplikasi publik berarti alamatnya dapat dibuka siapa saja; ini bukan pembatasan akses internal. Pemrosesan dokumen berlangsung di browser pengguna. Jika kebijakan kantor mengharuskan akses aplikasi dibatasi, pengelola perlu memilih hosting internal atau autentikasi terpisah.

## 3. Unggah melalui halaman GitHub

1. Pada repositori, pilih **Add file → Upload files** (atau tautan upload pada repositori kosong).
2. Seret folder `docs`, `scripts`, `tests`, dan berkas lain dari hasil ekstraksi ke area unggah. Pastikan nama berkas ditampilkan dengan jalur seperti `docs/index.html`.
3. Isi pesan commit, misalnya `Rilis awal KOMAD`, lalu tekan **Commit changes** ke branch `main`.
4. Pastikan `docs/index.html`, `docs/sw.js`, `docs/js/app.js`, serta `docs/assets/logo-komad.png` dapat ditemukan dalam repositori.
5. Berkas `docs/.nojekyll` sebaiknya ikut ada. Jika sistem operasi menyembunyikannya sehingga tidak terunggah, pilih **Add file → Create new file**, beri nama `docs/.nojekyll`, isi satu baris kosong, lalu commit.

Paket ini kurang dari 100 berkas dan setiap berkas di bawah 25 MiB, sehingga sesuai batas unggahan melalui browser GitHub.

## 4. Aktifkan Pages

1. Buka **Settings → Pages** pada repositori.
2. Pada **Build and deployment → Source**, pilih **Deploy from a branch**.
3. Pilih branch **main**, folder **/docs**, kemudian **Save**.
4. Tunggu proses publikasi selesai. Buka alamat situs yang ditampilkan GitHub pada halaman Pages. Gunakan alamat itu untuk dibagikan, bukan alamat repositori kode.

Jangan memindahkan isi `docs/` ke root setelah konfigurasi ini. Semua jalur aplikasi relatif sehingga dapat berjalan pada subfolder repositori GitHub Pages.

Referensi resmi:

- https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site
- https://docs.github.com/en/repositories/working-with-files/managing-files/adding-a-file-to-a-repository

## 5. Periksa sebelum membagikan

- Alamat menggunakan HTTPS; logo dan seluruh menu terbuka.
- Pilih **Lihat Contoh → Gunakan Matriks Contoh**. Pastikan ada 4 pasal; Pasal 1 menggunakan 3 bulan; Pasal 2 dan 3 tetap terbaca.
- Coba lokasi catatan, sorotan teks, rekomendasi, dan zoom draft. Selesaikan catatan, unduh Word, lalu buka di Microsoft Word.
- Coba matriks kantor yang sudah dikenal dan bandingkan hasilnya dengan sumber sebelum penggunaan rutin.
- Tunggu status siap offline, lalu pasang aplikasi jika menu tersedia. Jangan hanya menilai dari adanya ikon aplikasi.
- Tutup dokumen kerja, putus internet, buka kembali aplikasi dari ikon/alamat yang sama. Coba matriks contoh dan konversi Word. Uji pada jenis perangkat yang akan digunakan perancang.

Jika status gagal, buka panel pemasangan/pembaruan dan catat kode diagnosis serta komponen gagal. Bila `index.html` gagal, pastikan Pages menayangkan folder `/docs` dan halaman utama benar-benar merupakan KOMAD, bukan halaman repositori atau halaman masuk.

## Memasang pada perangkat perancang

Setelah seluruh aset siap offline, gunakan tombol pemasangan yang tersedia. Jika browser tidak menampilkan permintaan otomatis, ikuti petunjuk pada **Tentang penggunaan offline** atau menu pemasangan browser. Nama menu dapat berbeda antarperangkat. Perangkat yang tidak mendukung pemasangan tetap dapat membuka alamat web; dukungan offline perlu diuji pada browser tersebut.

KOMAD yang dahulu dipasang dari alamat hosting lama tidak otomatis berpindah ke GitHub. Buka alamat GitHub baru, siapkan offline, dan pasang dari alamat baru. Setelah dipastikan bekerja, pengguna dapat melepas ikon lama agar tidak tertukar.

## Memperbarui aplikasi kemudian

1. Simpan cadangan versi terakhir atau gunakan riwayat commit GitHub.
2. Ubah kode, jalankan tes, lalu jalankan `python scripts/build-pwa.py` untuk membentuk versi cache baru. Paket rilis pertama ini sudah dibentuk; langkah ini diperlukan jika ada perubahan sesudahnya.
3. Unggah berkas yang berubah ke jalur yang sama, termasuk `docs/sw.js` hasil generator. Commit. Tunggu Pages selesai.
4. Pengguna membuka aplikasi saat terhubung internet, menyimpan draft yang sedang dikerjakan, lalu mengikuti pemberitahuan pembaruan. Memuat ulang halaman membantu pemeriksaan pembaruan, tetapi tab yang masih berisi dokumen aktif menahan aktivasi versi baru.
5. Setelah versi baru siap, uji offline lagi pada satu perangkat sebelum mengumumkannya ke seluruh perancang.

Kirim alamat aplikasi yang sama untuk penggunaan harian. ZIP hanya untuk pengelola menyimpan atau memperbarui kode, bukan dibagikan berulang kepada setiap perancang.
