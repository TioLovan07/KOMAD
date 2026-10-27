# KOMAD 1.1.0

## Perubahan
- Revisi parsial bernomor tanpa judul Pasal digabungkan hanya jika penanda unik cocok. Judul, pengantar, dan butir lain dipertahankan. Pengguna harus memeriksa hasil gabungan.
- Pencocokan ambigu ditahan; judul Pasal yang hilang tanpa petunjuk penghapusan/pemindahan menjadi penghambat unduhan.
- Pratinjau mempertahankan font serif, perataan judul dan baris terakhir justify. Tidak mengubah profil format Word.
- Kamus offline Hunspell Indonesia dengan aturan imbuhan, saran satu edit, kata ulang dan klitik terbatas. Semua keputusan tetap pada pengguna.
- Tambahkan istilah ke kamus perangkat; penyimpanan lokal bisa terhapus saat data situs dibersihkan dan tidak tersinkron antarperangkat.
- Tautan pada pratinjau dinonaktifkan sebagai pengamanan tambahan. Ini belum merupakan perbaikan menyeluruh audit keamanan sebelumnya.
- Nama keluaran Draft.docx konsisten pada ringkasan dan unduhan.

## Pengujian
Corpu: 22 judul Pasal (1–22) dan 21 definisi Pasal 1; angka 18–19 memakai penyempurnaan. Word dibentuk dan diverifikasi kembali.
Uji otomatis: pembaca matriks, revisi parsial/ambigu, lokasi catatan, cache PWA, kamus dan saran.
Uji browser lokal: unggah Corpu, render, pembuatan Word; mode jaringan offline setelah cache siap tetap menghasilkan 22 pasal dan pemeriksaan kamus berjalan tanpa kegagalan. Tidak dilakukan perubahan situs publik.

## Batas
- Bukan pemeriksa KBBI resmi atau pemeriksa substansi hukum.
- Kata tidak dikenali dapat benar. Singkatan kapital dilewati.
- Saran satu edit tidak menemukan semua typo; kata benar yang salah konteks tidak terdeteksi.
- Maksimal 200 kartu ejaan. Jika terlampaui, peringatan menyatakan cakupan yang belum ditampilkan.
- Waktu pemeriksaan maksimal 60 detik. Jika gagal, catatan menyatakan pemeriksaan kamus belum selesai.
- Pratinjau HTML tidak identik dengan Microsoft Word; font yang tersedia di perangkat memengaruhi layout.
- Penggabungan tidak menebak pemindahan/penomoran ambigu. Semua jenis variasi matriks belum teruji.

## Memperbarui GitHub
Salin seluruh isi paket ke folder clone repositori (jangan hapus .git). Sertakan folder docs/assets/dictionary, vendor/nspell.js, seluruh docs/js dan docs/sw.js.
Jangan hanya menyalin file JS karena kamus juga diperlukan untuk offline.
Kemudian:
    git add -A
    git diff --cached --stat
    git commit -m "Perbarui KOMAD ke v1.1.0"
    git push origin main
Pastikan hanya kode/aplikasi yang masuk; jangan tambahkan matriks kantor.
Tunggu publikasi selesai. Buka aplikasi online, pastikan versi 1.1.0 dan kesiapan offline selesai; tutup seluruh tab aplikasi lama lalu buka lagi bila pembaruan tertahan.
Uji offline setelah kesiapan selesai. Unggah ulang matriks untuk pemeriksaan baru.
