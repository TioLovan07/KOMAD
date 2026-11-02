# Komponen pihak ketiga

## JSZip 3.10.1

Dipakai untuk membaca dan membentuk arsip `.docx`. Berkas lokal: `docs/vendor/jszip.min.js`. Lisensi disertakan pada `docs/vendor/JSZip-LICENSE.txt` dan header distribusi. JSZip juga menyertakan pako; pemberitahuannya terdapat pada berkas distribusi/lisensi tersebut.

## docx-preview

Renderer lokal dari fork `@apollo-design/docx-preview`, berdasarkan docxjs oleh Volodymyr Baydalka, lisensi Apache 2.0. Lihat `docs/vendor/docx-preview-LICENSE.txt` dan `docs/vendor/docx-preview-NOTICE.txt`.

`docs/vendor/docx-preview.js` adalah distribusi yang digunakan aplikasi. Berkas ini diwarisi dari mesin konversi sebelumnya dan memiliki penyesuaian posisi tab agar tidak menghitung indentasi paragraf dua kali. Jangan menggantinya dengan CDN atau distribusi lain tanpa pengujian format, tabel, penomoran, dan sorotan bookmark. `docx-preview.min.js` disertakan sebagai berkas pendamping sumber sebelumnya dan tidak dimuat aplikasi.

Logo KOMAD berasal dari aset yang diberikan pengguna. Dokumen contoh bawaan berisi ilustrasi untuk latihan. Tidak ada font komersial yang dibundel; perangkat menggunakan font yang tersedia.

## Local preview pagination patch (KOMAD 1.3.0)
The bundled @apollo-design/docx-preview renderer is modified in paginateSection:
the KOMAD first-page spacer header is not cloned to continuation pages, and
available page height accounts for occupied header space. Upstream license
notice is retained. This patch does not change DOCX content.

KOMAD 1.3.1 additional local docx-preview changes: page-relative floating-image positioning, direct paragraph pageBreakBefore, a pre-pagination layout hook, DOM Range splitting with intact inline descendants, and splitting long table rows with or without tbody. Original upstream license remains included.

## Legacy DOC reader

Alpaq92/JSDoc, src/docToText.js, 0BSD. Pinned commit 821695a884e0c0bb8592a635d9524bb3e116cd67, https://github.com/Alpaq92/JSDoc. License: docs/vendor/legacy-doc-LICENSE.txt. Local change: preserve paragraph/list properties at cell endings; suppress marker-only paragraphs when all source text is struck/hidden in the OOXML adapter. No writer or sample document is bundled. The reader is browser-local and never sends document contents to a service. WMF/EMF and some shapes are not supported.

Penyesuaian adapter DOC: pertahankan indent paragraf, rekonstruksi gridSpan dari batas koordinat tabel, dan pemulihan terbatas gambar raster mengambang pada sel identitas. Pembaca vendor tetap tidak menjamin seluruh shape/textbox DOC.
