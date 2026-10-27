# KOMAD 1.3.0 — Perbaikan gabungan hasil pengujian matriks
Paket lengkap; tidak perlu memasang versi perantara.

## Format
- F4 potret 215 x 330 mm, Bookman Old Style 12 pt, spasi tunggal.
- Halaman pertama: area teks pertama di bawah Garuda dimulai sekitar 8 cm dari atas.
- Halaman selanjutnya: margin atas 2,5 cm; kiri/kanan/bawah 2,5 cm.
- Di Word, ruang halaman pertama dibuat dengan header khusus halaman pertama; margin section 2,5 cm. Karena itu kotak Page Setup tidak menampilkan 8 cm sebagai margin semua halaman.
- Garuda yang tersedia pada sumber ditempatkan tengah, tinggi 3,5 cm, tepi atas 3,5 cm, tepi bawah 7 cm. Tidak membuat gambar baru jika sumber tidak memilikinya.
- Jarak sebelum MEMUTUSKAN dikurangi dari 72 pt ke 12 pt. Isi langsung setelah judul pasal memakai jarak sebelum 0 pt.
- Lebar tabel format mengikuti ruang di dalam margin.

## Kelengkapan dan penomoran
- Tetap pada satu pasal memakai judul/nomor hasil serta isi awal, termasuk ketika judul awal dicoret atau nomor berubah.
- Metadata judul yang dicoret hanya dipakai untuk pemetaan, tidak dimunculkan lagi sebagai judul lama.
- Jika hasil berhenti pada judul pasal atau awal rumusan cocok tetapi lanjutan lebih pendek, aplikasi menunjukkan kandidat awal dan meminta keputusan. Tidak menganggap bagian tersebut dihapus.
- Gunakan Lanjutan Awal menambahkan kandidat lalu memeriksa ejaan kembali; Hasil Sudah Sesuai, Tanpa Tambahan mempertahankan hasil.
- Ini heuristik, bukan penilaian substansi hukum. Pencocokan kompleks tetap harus dibandingkan oleh perancang.
- Butir perubahan pada Pasal I memakai urutan tersendiri, tidak ter-reset oleh daftar definisi/ayat di dalamnya. BAB IIIA dan Pasal I/II dikenali sebagai judul.

## Pratinjau
- Sorotan dibuat per potongan teks, tanpa membongkar kolom penanda/isi paragraf.
- Uji sorotan Corpu berulang tidak mengubah teks maupun tinggi paragraf.
- Renderer dipatch agar ruang judul halaman pertama tidak disalin ke semua halaman.
- Pemenggalan halaman dan font pengganti browser tetap dapat berbeda dari Word.

## Ejaan
Kosakata tambahan: nonalam, suburusan, pascabencana, kedaruratan, terencana, kesiapsiagaan, kesiapan, berakhlak, berkepribadian, luhur, monitoring, kependudukan, kependidikan, berintegritas, fisik-motorik.
Koreksi eksplisit ditambahkan untuk koordinai, singkronisasi, efcktivitas, Penyelengaraan, Konflilk, Iuhur, bekerjasama, diatas, dilapangan, dibelahan, diluar, diantaranya, diantara, mensupervisi, dengan variasi kapitalisasi.
Kata perundang tidak diterima sebagai kata mandiri; periksa perundang-undangan secara utuh. Monitoring dan fisik-motorik tidak diubah otomatis. Koreksi tetap memerlukan keputusan pengguna.
Kamus tidak menjamin semua typo ditemukan atau konteks kalimat benar.

## Lampiran dan tipe file — batas scope
- Kepala lampiran sumber digunakan satu kali. Kepala otomatis hanya jika tidak tersedia.
- Belum mengekstrak materi lampiran di luar matriks utama (termasuk 16 tabel tambahan file Buleleng).
- Dukungan Word 97–2003 .doc ditunda. Masukan tetap .docx.
- Tidak ada perubahan hosting dari paket ini.

## Validasi
Kasus lokal: Corpu, Peninjoan, perubahan Pengelolaan Risiko, Juknis Konflik Kepentingan Buleleng.
Browser: impor, pembentukan pratinjau, sorotan berulang, pilihan tambahan sumber dan pembentukan Word.
Word: profil halaman pertama/lanjutan diverifikasi lewat XML dan hasil render.
Tes otomatis: pembaca matriks, revisi parsial, Tetap pasal berubah nomor/judul dicoret, lokasi, ejaan, integritas nomor lintas baris, cache offline.
Buleleng belum final: banyak kandidat sumber yang membutuhkan keputusan perancang. Jangan memilih seluruh tambahan secara massal tanpa membaca.
Kasus uji tidak mencakup semua bentuk matriks kantor. Audit keamanan terdahulu belum seluruhnya dituntaskan.

## Cara memperbarui
Salin ISI paket ke C:\Users\ASUS\KOMAD dan pertahankan .git.
Jangan ikut menyalin dokumen kantor ke repositori publik.
    git add -A
    git diff --cached --stat
    git commit -m "Perbarui KOMAD ke v1.3.0"
    git push origin main
Tunggu publikasi. Tutup tab aplikasi lama, buka online, pastikan Versi 1.3.0 dan kesiapan offline. Unggah ulang matriks.
Lakukan uji penerimaan perancang sebelum penggunaan resmi.
