# KOMAD 1.3.2

Paket lengkap, termasuk perbaikan 1.3.1. Tidak perlu memasang versi perantara.

## Revisi margin dalam paket 1.3.2 yang sama

Paket ini sudah menggantikan paket 1.3.2 sebelumnya. Nomor versi tetap 1.3.2; cache offline dibangun ulang agar perubahan dikenali.

- Paragraf teks pertama di bawah Garuda dimulai pada posisi 8 cm dari tepi atas kertas. Garuda tetap pada posisi atas 3,5 cm dengan tinggi 3,5 cm.
- Halaman berikutnya tetap mempunyai margin atas/kiri/kanan/bawah 2,5 cm.
- Ruang khusus halaman pertama kini memakai satu blok ruang tetap 5,5 cm di awal isi, bersama margin dasar 2,5 cm. Header kosong lama dikeluarkan karena dapat diabaikan saat render. Tidak menambahkan teks atau mengubah rumusan.
- Page Setup menampilkan margin section 2,5 cm. Posisi awal teks halaman pertama tetap 8 cm karena ruang judul hanya muncul sekali; ini bukan jarak 8 cm sesudah Garuda.
- Diuji langsung pada Draft (6): hasil render menempatkan tepi huruf pertama sekitar 8,08 cm (posisi paragraf 8 cm), sedangkan halaman 2–9 sekitar 2,58 cm (margin 2,5 cm). Selisih kecil berasal dari bentuk huruf di dalam baris.
- Ekspor dari matriks Peninjoan dan Risiko serta kasus uji lampiran juga diperiksa. Pemeriksaan posisi teks pertama ditambahkan ke tes browser, selain pemeriksaan tinggi halaman, gambar, dan kesamaan seluruh teks.
- Deteksi typo, pemilihan rumusan, desain tombol, penomoran, dan format lampiran tetap seperti 1.3.2 sebelumnya. Pemenggalan halaman dapat bergeser karena ruang halaman pertama kini benar.

Unduh ulang paket ini, salin seluruh isinya ke repo, push, lalu tutup tab aplikasi lama dan buka online. Unggah ulang matriks dan unduh draft baru; file Word yang sudah tersimpan tidak berubah otomatis.

## Catatan butir perubahan nomor 3 pada matriks Risiko

Baris matriks 11 mempunyai perbedaan nyata: kolom awal mengubah Pasal 14 tentang kriteria Risiko, sementara kolom hasil mengubah Pasal 15 tentang identifikasi Risiko dan menambah ayat (3). Draft mengikuti kolom hasil. Pemeriksaan sebelumnya menampilkan Periksa kelengkapan Pasal 14, yang dapat disalahpahami sebagai kehilangan isi draft.

Untuk perubahan sasaran pasal yang dinyatakan dalam pengantar perubahan, catatan kini berbunyi Konfirmasi perubahan sasaran pasal, menyebut nomor awal dan hasil, dan menjelaskan bahwa draft mengikuti hasil. Tetap membutuhkan konfirmasi perancang karena aplikasi tidak dapat memastikan apakah perubahan substansi itu disengaja. Tidak mengembalikan Pasal 14 secara otomatis dan tidak menghapus catatan perbedaan sumber secara diam-diam. Pemeriksaan pasal yang sungguh tidak tersedia tetap berjalan.

## Alur file

- Tombol pengganti di bagian pemeriksaan bernama Ganti File Matriks. Ini mengganti file yang sedang dikerjakan; tidak harus dipakai jika matriks sudah benar.
- Tombol Mulai Dokumen Baru yang berdampingan dengannya dihapus. Pengguna cukup memakai Ganti File Matriks untuk membuka file lain. Mulai Dokumen Baru tetap tersedia setelah unduhan untuk kembali ke langkah awal.
- Kartu yang benar-benar memerlukan perbaikan sumber mengarahkan pengguna ke tombol pengganti di atas, tanpa menggandakan tombol unggah pada setiap kartu.
- File pengganti yang gagal dibaca tidak menghapus dokumen lama; unggahan yang berhasil memulai pemeriksaan baru.

## Koreksi ejaan

Kartu koreksi manual dan hasil kamus offline kini mempunyai judul, pesan, kotak rekomendasi, serta tombol Terapkan Rekomendasi dan Biarkan Sesuai Matriks yang sama. Kandidat tunggal tampil sebagai teks, bukan tombol pilihan yang tidak diperlukan. Beberapa kandidat tetap dapat dipilih jika tersedia.

Pilihan menganggap kata benar pada perangkat disimpan dalam bagian Pilihan kamus. Ini untuk nama/istilah yang memang benar, bukan untuk menyembunyikan salah ketik.

Koreksi eksplisit: penelahaan menjadi penelaahan; pertangungjawaban menjadi pertanggungjawaban; oragnisasi menjadi organisasi. Bentuk salah tidak dimasukkan sebagai kata yang diterima. Kosakata terkait yang benar ditambahkan, antara lain menelaah, penelaah, mempertanggungjawabkan, dipertanggungjawabkan, pengorganisasian, keorganisasian, mengorganisasikan, terorganisasi, akuntabilitas, pengawasan, pengendalian, dan pelaporan.

## Unduhan

Tautan Simpan Draft.docx sebelumnya adalah cadangan untuk unduhan. Sekarang hanya tombol unduh yang ditampilkan; tautan sementara di belakang tombol tidak terlihat dan dihapus sesudah diklik. Tombol boleh ditekan kembali jika berkas belum tersimpan. Berlaku untuk Word dan catatan pemeriksaan.

## Validasi

Tes regresi lulus, termasuk pembeda pasal benar-benar hilang versus perubahan sasaran pada matriks perubahan. Browser diuji dengan matriks Risiko: catatan baru, tampilan ejaan seragam, satu tombol pengganti file, penerapan koreksi, penyelesaian catatan, unduhan Word dua kali, dan unduhan catatan tanpa tautan tambahan atau galat JavaScript. Tidak menetapkan bahwa substansi kolom hasil sudah benar secara hukum.

## Memasang

Salin isi paket ke C:\Users\ASUS\KOMAD dengan mempertahankan .git. Jangan unggah matriks kantor ke repositori publik.

    git add -A
    git diff --cached --stat
    git commit -m "Rapikan pemeriksaan dan alur KOMAD v1.3.2"
    git push origin main

Setelah publikasi, buka aplikasi saat online dan pastikan Versi 1.3.2 serta kesiapan offline. Unggah ulang matriks. Paket ini tidak melakukan push atau mengubah hosting otomatis.
