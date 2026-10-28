# KOMAD 1.3.1

Paket lengkap, melanjutkan 1.3.0. Salin isi paket ke folder repo, pertahankan folder .git. Tidak menyertakan matriks kantor dan tidak mengubah hosting otomatis.

## Perbaikan

- Garuda pada Word tetap menggunakan gambar sumber. Pratinjau kini menangani gambar mengambang relatif terhadap halaman dengan lebar/tinggi yang nyata, sehingga tidak tersembunyi oleh kotak selebar nol.
- Pembagian halaman pratinjau dapat memecah baris tabel panjang, termasuk tabel tanpa elemen tbody. Ini mengurangi ruang kosong ketika Menimbang/Mengingat melewati batas halaman. Pemisahan teks mempertahankan seluruh potongan run dan formatnya.
- Indentasi daftar huruf di bawah ayat mengikuti tingkat semantik ayat/huruf; angka level bawaan penomoran matriks tidak lagi membuat butir tertentu menjorok berlebihan.
- Seluruh baris judul BAB yang dilanjutkan dengan paragraf huruf kapital memakai format judul yang sama, termasuk DAN TIPE. Teks judul tidak dipotong/dihilangkan.
- Kepala LAMPIRAN I, II, dan seterusnya memulai ulang status kepala dan tanda tangan. Kepala sumber tetap dipakai sekali, dengan format konsisten. Lampiran berikutnya mempunyai page break sebelum kepala, baik pada Word maupun pratinjau.
- Butir pengantar perubahan di Pasal I diberi jarak sebelum 12 pt (satu baris nominal), tanpa memasukkan paragraf kosong ke teks. Contoh UU 13 Tahun 2022 menjadi pembanding susunan, bukan dasar klaim jarak wajib 12 pt.
- Catatan lanjutan rumusan awal kini membandingkan isi yang dinormalisasi, bukan sekadar menganggap paragraf lebih sedikit berarti ada isi hilang. Teks sama yang digabung menjadi paragraf/tabel tidak lagi menjadi kandidat tambahan. Angka koordinat tetap dibedakan. Perubahan redaksi yang disengaja masih memerlukan keputusan manusia.

## Ejaan

Koreksi yang ditawarkan (tetap menunggu pilihan pengguna):

| Bentuk sumber | Saran |
| --- | --- |
| terakomodir | terakomodasi |
| komperensif | komprehensif |
| Restra | Renstra |
| priorotas | prioritas |
| anlisis | analisis |
| sistema | sistem |

Pada matriks Risiko, sistema berada dalam frasa kegagalan sistema, sehingga saran sistem sesuai konteks tersebut. Jangan memakai rekomendasi itu untuk nama/istilah asing tanpa membaca konteks.

Ketidakcukupan, mengestimasi, reviu, dan menatausahakan diterima. Stakeholders dikenali sebagai istilah asing, bukan dinyatakan kata baku Indonesia; pemangku kepentingan dapat dipilih secara editorial sesuai konteks. Bentuk hasil koreksi yang benar juga masuk kosakata tambahan.

## Validasi dan batasnya

- Impor dan pembentukan Word diuji pada Peninjoan serta matriks perubahan Pengelolaan Risiko.
- Kasus buatan menguji tabel satu baris yang panjang, judul BAB dua paragraf, serta Lampiran I dan II dalam satu baris matriks.
- Seluruh teks Word dan pratinjau dibandingkan pada tiga kasus tersebut: sama setelah whitespace diabaikan. Semua halaman pratinjau mempunyai tinggi kertas yang benar, gambar terlihat, dan dua lampiran berada pada halaman berbeda.
- Keluaran Word dirender untuk pemeriksaan format. Pratinjau tetap dapat berbeda pemenggalan halaman dari Microsoft Word karena mesin/font berbeda.
- Matriks Pulasari dan matriks persis pada foto Garuda/Lampiran I–II tidak tersedia pada saat revisi ini. Perbaikan umum sudah diuji, tetapi verifikasi persis file tersebut masih memerlukan sumbernya. Jangan menganggap foto keenam telah diverifikasi terhadap matriks Pulasari.
- Pengolahan materi lampiran di luar tabel utama dan dukungan .doc tetap ditunda.

## Memasang

Ekstrak, salin seluruh isi ke C:\Users\ASUS\KOMAD, lalu dari folder repo:

    git add -A
    git diff --cached --stat
    git commit -m "Perbaiki pratinjau dan format KOMAD v1.3.1"
    git push origin main

Setelah publikasi selesai, tutup tab aplikasi lama, buka kembali saat online, pastikan versi 1.3.1 dan status offline siap. Unggah ulang matriks. Lakukan pemeriksaan perancang sebelum dokumen digunakan resmi.
