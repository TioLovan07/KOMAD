# Komponen pihak ketiga

## JSZip 3.10.1

Dipakai untuk membaca dan membentuk arsip `.docx`. Berkas lokal: `docs/vendor/jszip.min.js`. Lisensi disertakan pada `docs/vendor/JSZip-LICENSE.txt` dan header distribusi. JSZip juga menyertakan pako; pemberitahuannya terdapat pada berkas distribusi/lisensi tersebut.

## docx-preview

Renderer lokal dari fork `@apollo-design/docx-preview`, berdasarkan docxjs oleh Volodymyr Baydalka, lisensi Apache 2.0. Lihat `docs/vendor/docx-preview-LICENSE.txt` dan `docs/vendor/docx-preview-NOTICE.txt`.

`docs/vendor/docx-preview.js` adalah distribusi yang digunakan aplikasi. Berkas ini diwarisi dari mesin konversi sebelumnya dan memiliki penyesuaian posisi tab agar tidak menghitung indentasi paragraf dua kali. Jangan menggantinya dengan CDN atau distribusi lain tanpa pengujian format, tabel, penomoran, dan sorotan bookmark. `docx-preview.min.js` disertakan sebagai berkas pendamping sumber sebelumnya dan tidak dimuat aplikasi.

Logo KOMAD berasal dari aset yang diberikan pengguna. Dokumen contoh bawaan berisi ilustrasi untuk latihan. Tidak ada font komersial yang dibundel; perangkat menggunakan font yang tersedia.
