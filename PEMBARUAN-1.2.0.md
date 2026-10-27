# KOMAD 1.2.0 — Paket gabungan persiapan penggunaan kantor

## Format yang disepakati
Bookman Old Style 12 pt, 1 spasi. F4 potret dipertahankan pada ukuran profil 215 x 330 mm.
Margin atas 8 cm pada SEMUA halaman, termasuk lampiran; kiri/kanan/bawah 2,5 cm.
Halaman bertambah karena ruang isi berkurang. Pastikan pilihan kertas printer cocok (jangan otomatis Letter/A4).

## Perbaikan gabungan
- Semua perbaikan 1.1.0–1.1.2: penggabungan parsial, penyaringan ejaan, kamus offline, tampilan rekomendasi, penomoran pratinjau.
- Pemeriksaan pasal mencakup semua baris. Pasal 9 Peninjoan pada baris 19 tidak lagi dilaporkan hilang dari baris 18.
- Ketidakpastian nomor pasal menjadi catatan yang dapat ditinjau pengguna, bukan tuduhan matriks salah. Pemrosesan yang ambigu/struktur rusak tertentu tetap ditahan.
- Kata kapital yang belum dikenali tidak dipromosikan otomatis sebagai typo karena dapat merupakan nama orang/tempat. Koreksi eksplisit tetap tersedia.
- Dokumen tetap harus diperiksa; tidak menjamin bebas typo atau semua variasi matriks akurat.
- Tautan pratinjau tidak aktif; audit keamanan terdahulu belum seluruhnya diselesaikan.

## Validasi yang dilakukan
- Tes otomatis pembaca, penggabungan parsial/ambigu, lokasi, PWA, kamus, penyaringan, dan pasal lintas baris.
- Browser lokal: impor, pratinjau dan pembentukan Word pada Corpu dan Peninjoan.
- Corpu: 22 pasal; BAB I, Pasal 1, dan 21 definisi dipertahankan; angka 18–19 memakai penyempurnaan.
- Peninjoan: 9 pasal; tidak ada penghambat palsu Pasal 9 hilang.
- Kedua dokumen berhasil dimuat/diperiksa dengan jaringan offline setelah persiapan cache.
- XML Word: seluruh section memiliki margin atas 4535 twip (~8 cm) dan sisi lain 1417 twip (~2,5 cm); ukuran F4; spasi baris 240 (single).
- Font pengganti dan pagination browser dapat berbeda dari Word. Pengujian tidak mencakup seluruh matriks kantor.

## Pembaruan
Ekstrak ZIP dan salin ISI paket ke folder repositori C:\Users\ASUS\KOMAD. Pertahankan .git.
Sertakan docs/assets/dictionary, seluruh js/vendor, serta sw.js.
Jangan menyertakan dokumen kantor.
    git add -A
    git diff --cached --stat
    git commit -m "Perbarui KOMAD ke v1.2.0"
    git push origin main
Tunggu GitHub Pages selesai. Tutup tab aplikasi lama, buka online, pastikan Versi 1.2.0 dan kesiapan offline selesai. Unggah ulang matriks.
Tidak perlu mengunggah paket versi sebelumnya satu per satu.

## Pemeriksaan sebelum penggunaan resmi
1. Perancang membandingkan rumusan akhir dengan matriks yang telah disepakati.
2. Buka Word pada komputer kantor: cek margin, font, penomoran, judul, lampiran, dan cetak percobaan.
3. Uji perangkat pengguna sasaran, koreksi/biarkan/sorotan/unduhan, serta offline.
4. Catat versi yang diuji dan masalah tersisa; jangan mengganti versi diam-diam di tengah validasi.
5. Perbaikan fitur baru ditunda selama validasi; fokus pada kehilangan isi, keluaran salah, atau aplikasi gagal.
