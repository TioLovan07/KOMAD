<p align="center">
  <img src="docs/assets/logo-komad.png" alt="Logo KOMAD" width="280">
</p>

# KOMAD

**Konversi Matriks Menjadi Draft**

Aplikasi web untuk membantu perancang peraturan mengolah matriks harmonisasi menjadi draft Word dengan format yang seragam. KOMAD membantu memindahkan rumusan, menelusuri bagian yang perlu diperiksa, dan mengurangi pekerjaan salin-tempel berulang.

**[Buka aplikasi](https://tiolovan07.github.io/KOMAD/)** · [Panduan pemasangan dan pembaruan](PANDUAN_GITHUB.md) · [Catatan rilis](PEMBARUAN-1.3.4.md)

## Fitur utama

- **Pembacaan matriks Word** — menerima DOCX serta DOC 97–2003 pada struktur yang didukung.
- **Penyusunan rumusan** — mengolah hasil penyempurnaan, petunjuk “Tetap”, dan revisi sebagian yang sasarannya dapat dicocokkan dengan jelas.
- **Pemeriksaan ejaan** — menampilkan dugaan typo dan rekomendasi dari kamus offline untuk dipilih pengguna.
- **Catatan dengan lokasi** — menghubungkan bagian yang perlu diperiksa dengan sorotan pada pratinjau, disertai penjelasan dan tindakan.
- **Draft dengan format seragam** — menata judul, Pasal, ayat, daftar, margin, serta halaman judul lampiran.
- **Ekspor Word** — menghasilkan `Draft.docx` untuk pemeriksaan dan penyelesaian akhir.
- **Penggunaan offline** — tersedia setelah seluruh aset aplikasi tersimpan dan status kesiapan offline terkonfirmasi.

## Cara menggunakan

1. Buka aplikasi dan pilih file matriks Word.
2. Periksa catatan, rekomendasi ejaan, dan bagian yang disorot pada pratinjau.
3. Cocokkan rumusan draft dengan matriks sumber dan selesaikan bagian yang memerlukan kepastian.
4. Unduh `Draft.docx`, lalu periksa dan lengkapi dokumen melalui Microsoft Word.

## Pengolahan dokumen

Pembacaan matriks dan pembentukan draft dilakukan di browser pengguna. KOMAD menggunakan aturan pemilihan dan penggabungan rumusan untuk membantu penyusunan; keputusan substansi tetap berada pada perancang.

Kata yang belum dikenali kamus belum tentu salah. Rekomendasi ejaan perlu diperiksa berdasarkan konteks sebelum diterapkan.

## Batas dukungan

- Variasi susunan tabel dan pengisian matriks dapat memengaruhi pembacaan. Struktur atau sasaran revisi yang belum jelas memerlukan perbaikan pada sumber.
- DOC 97–2003 memiliki keterbatasan pada beberapa jenis gambar, bentuk, dan tabel kompleks. Gunakan DOCX bila tersedia.
- Lampiran disiapkan sebagai halaman portrait dengan judul otomatis. Isi lampiran dilengkapi atau digabungkan di Word.
- Ketersediaan font dan aplikasi pembuka dokumen dapat memengaruhi pemenggalan halaman. Periksa isi dan tata letak akhir sebelum digunakan.

## Teknologi

| Komponen | Teknologi |
| --- | --- |
| Antarmuka dan pemrosesan | HTML, CSS, JavaScript |
| Pengolahan paket DOCX | JSZip |
| Pratinjau Word | docx-preview |
| Pemeriksaan ejaan | nspell dan kamus offline |
| Kesiapan offline dan pembaruan | Progressive Web App dan Service Worker |

Informasi komponen pihak ketiga tersedia pada [THIRD_PARTY.md](THIRD_PARTY.md).

## Menjalankan secara lokal

Dengan Python 3, jalankan dari folder proyek:

```bash
python3 buka_komad.py
```

Pada Windows, gunakan `BUKA-KOMAD.cmd` bila Python sudah tersedia.

Untuk pengembangan, jalankan tes dengan Node.js dan bentuk ulang cache setelah mengubah aset aplikasi:

```bash
npm test
python3 scripts/build-pwa.py
```

Petunjuk publikasi melalui GitHub Pages tersedia dalam [PANDUAN_GITHUB.md](PANDUAN_GITHUB.md).
