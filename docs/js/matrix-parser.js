import { missingSourceBlocks } from './source-coverage.js';
import { missingArticles } from './article-integrity.js';
import { replacementMatches } from './draft-model.js';
import { checkSpelling } from './spell-check.js';
import {
    SPELLING_CORRECTIONS,
    isEditorialInstruction
} from './editorial-checks.js';
/**
 * Membaca ZIP Word dan memilih rumusan sumber. Perubahan isi hanya melalui keputusan eksplisit pengguna.
 */
import {
    $,
    nodes,
    direct,
    norm,
    xml,
    parse,
    textOf,
    plain,
    articleSections,
    sectionBody,
    addIssue
} from './xml.js';
import {
    matrixCells,
    matrixHeader,
    matrixRowLabel,
    matrixNumberingCatalog,
    matrixParagraphs,
    matrixCloneBlock,
    matrixSelectRumusan
} from './matrix-reader.js';
import {
    attachIssueLocations
} from './issue-locations.js';
async function readMatrix(file) {
    if (!/\.docx$/i.test(file.name)) throw Error('Pilih file Word berformat .docx.');
    if (file.size > 20 * 1024 * 1024) throw Error('Ukuran file melebihi 20 MB.');
    const z = await JSZip.loadAsync(file);
    let total = 0;
    for (const f of Object.values(z.files)) {
        total += f._data?.uncompressedSize || 0;
    }
    if (total > 100 * 1024 * 1024) throw Error('Isi dokumen terlalu besar untuk diproses pada perangkat ini.');
    const f = z.file('word/document.xml');
    if (!f) throw Error('Dokumen ini bukan file Word yang didukung.');
    const d = parse(await f.async('string'));
    const numberingFile = z.file('word/numbering.xml'),
        stylesFile = z.file('word/styles.xml');
    const catalog = matrixNumberingCatalog(numberingFile ? parse(await numberingFile.async('string')) : null, stylesFile ? parse(await stylesFile.async('string')) : null);
    const body = nodes(d, 'body')[0],
        tables = direct(body, 'tbl');
    let table, header, cols;
    for (const candidate of tables) {
        const match = matrixHeader(candidate);
        if (match) {
            table = candidate;
            ({
                header,
                cols
            } = match);
            break;
        }
    }
    if (!table) throw Error('Kolom rancangan awal/rancangan peraturan dan hasil penyempurnaan belum dapat dikenali. Periksa judul kolom matriks.');
    const m = {
        zip: z,
        doc: d,
        name: file.name,
        rows: [],
        issues: [],
        decisions: [],
        filename: 'Draft.docx'
    };
    if (tables.length > 1) addIssue(m, {
        title: 'Ada lebih dari satu tabel',
        message: 'Tabel lain di luar matriks utama perlu diperiksa untuk memastikan tidak ada lanjutan naskah.',
        kind: 'review'
    });
    if (nodes(d, 'ins').length || nodes(d, 'del').length) addIssue(m, {
        title: 'Ada Track Changes',
        message: 'Sisipan digunakan dan penghapusan yang terlacak dikeluarkan. Pastikan perubahan sesuai keputusan.',
        kind: 'review'
    });
    const leftState = new Map(),
        rightState = new Map(),
        rs = direct(table, 'tr');
    m.decisions.push('Kolom dikenali: awal ' + (cols.awal + 1) + ', hasil ' + (cols.hasil + 1) + (cols.ket >= 0 ? ', keterangan ' + (cols.ket + 1) : '; tanpa kolom keterangan') + '.');
    for (let i = header + 1; i < rs.length; i++) {
        const cs = matrixCells(rs[i]),
            lc = cs[cols.awal],
            rc = cs[cols.hasil];
        if (!lc || !rc || lc === rc) throw Error('Baris ' + (i + 1) + ' menggabungkan kolom rumusan sehingga sumber awal dan hasil belum dapat dipisahkan.');
        if (/^\d+$/.test(norm(textOf(lc))) && /^\d+$/.test(norm(textOf(rc)))) continue;
        if (matrixHeader({
                children: [rs[i]]
            })) continue;
        const left = matrixParagraphs(lc, catalog, leftState),
            right = matrixParagraphs(rc, catalog, rightState),
            note = cols.ket >= 0 && cs[cols.ket] ? norm(textOf(cs[cols.ket])) : '',
            labels = matrixRowLabel(cs, Math.min(cols.awal, cols.hasil));
        const selected = matrixSelectRumusan(left, right, {
            m,
            rowIndex: i + 1,
            prefix: labels.prefix
        });
        const row = {
            index: i + 1,
            left,
            right,
            note,
            useRight: selected.useRight,
            ps: selected.ps,
            label: labels.label || ('Baris ' + (i + 1)),
            prefix: labels.prefix
        };
        if (row.ps.some(p => /^LAMPIRAN\b/i.test(norm(p.text)))) row.label = 'Lampiran';
        m.rows.push(row);
        m.decisions.push('Baris ' + row.index + ' · ' + selected.rule + (nodes(selected.useRight ? rc : lc, 'strike').length ? ' Teks dicoret dikeluarkan.' : ''));
        if (row.label === 'Lampiran' && right.length && !right.some(p => p.node.localName === 'tbl') && left.some(p => p.node.localName === 'tbl')) {
            const start = left.findIndex((p, j) => j > 1 && p.text.trim());
            addIssue(m, {
                kind: 'annex',
                row,
                title: 'Rincian lampiran belum ada di kolom hasil',
                message: 'Kolom hasil memuat judul lampiran, sementara rincian tabel tersedia di kolom awal. Tentukan apakah isi setelah judul dari kolom awal tetap digunakan.',
                annexBody: left.slice(start < 0 ? 2 : start).map(matrixCloneBlock)
            });
        }
    }
    let amendment = false, changeNumber = 0;
    for (const row of m.rows) for (const p of row.ps) {
        if (/^Pasal I$/i.test(norm(p.text))) amendment = true;
        if (/^Pasal II$/i.test(norm(p.text))) amendment = false;
        if (amendment && /^(?:\d+[.)]\s*)?(?:Ketentuan\s+(?:Pasal|huruf|Lampiran)|Di\s*antara\s+(?:BAB|Pasal))/i.test(norm(p.text))) {
            const manual = p.list?.prefixLength || 0;
            p.amendmentInstruction = true;
            p.list = {...p.list, marker: (++changeNumber) + '.', kind:'number', level:0, prefixLength:manual};
        }
    }
    // Identify incomplete replacements without silently filling legal substance.
    const articleRE = /^Pasal\s+\d+[A-Za-z]?$/i;
    const comparable = p => norm(p.text).replace(/^\(?\d+[.)]\s*/, '').replace(/\d+/g,'#').trim().toLowerCase();
    for (const row of m.rows) {
        if (!row.useRight || !row.left.length || !row.right.length) continue;
        const oldTitles = row.left.map((p,i)=>articleRE.test(norm(p.text))?i:-1).filter(i=>i>=0);
        const sourceArticles=[...new Set(row.left.map(p=>p.sourceArticle).filter(Boolean))];
        if (oldTitles.length !== 1 && !(oldTitles.length===0 && sourceArticles.length===1)) continue;
        const rightTitles = row.ps.map((p,i)=>articleRE.test(norm(p.text))?i:-1).filter(i=>i>=0);
        if (!rightTitles.length) continue;
        const oldBody = oldTitles.length ? row.left.slice(oldTitles[0]+1) : row.left.filter(p=>p.sourceArticle===sourceArticles[0]);
        const last = rightTitles.at(-1);
        if (last === row.ps.length-1 && oldBody.length) {
            addIssue(m, {kind:'source-gap', row, insertAt:last+1, candidates:oldBody.map(matrixCloneBlock),
                title:'Isi setelah '+norm(row.ps[last].text)+' belum tersedia',
                message:'Kolom hasil berakhir pada judul pasal. Rumusan awal tersedia pada baris yang sama. Bandingkan lalu pilih apakah isi awal digunakan; aplikasi tidak menganggapnya dihapus.'});
        } else if (rightTitles.length === 1 && oldBody.length>1) {
            const newBody=row.ps.slice(last+1);
            const firstOld=oldBody[0], firstNew=newBody[0];
            if (firstNew && comparable(firstOld).slice(0,65)===comparable(firstNew).slice(0,65) && newBody.length<oldBody.length) {
                const candidates=missingSourceBlocks(oldBody, newBody).map(matrixCloneBlock);
                if (!candidates.length) continue;
                addIssue(m, {kind:'source-gap', row, insertAt:row.ps.length, candidates,
                    title:'Ada lanjutan rumusan awal yang belum terpetakan',
                    message:'Awal rumusan cocok, tetapi beberapa blok teks awal belum ditemukan pada hasil setelah perbedaan paragraf, spasi, dan tanda baca diabaikan. Bandingkan teksnya; perubahan redaksi yang disengaja juga dapat memunculkan catatan ini.'});
            }
        }
    }
    // A move is applied only when an explicit note and matching destination text agree.
    const moveRE = /(?:reposisi|pemindahan|pindah\w*|materi)\s+(?:ketentuan\s+)?Pasal\s+(\d+[A-Za-z]?)/ig;
    const moves = new Map();
    for (const row of m.rows) {
        for (const x of row.note.matchAll(moveRE)) {
            if (/pindah|reposisi|masuk\s+ke/i.test(row.note)) moves.set(x[1], row.note);
        }
    }
    for (const [num, note] of moves) {
        const matches = [];
        for (const row of m.rows)
            for (const sec of articleSections(row.ps))
                if (sec.num === num) matches.push({
                    row,
                    sec
                });
        if (matches.length === 2) {
            const from = matches.find(x => !x.row.useRight),
                to = matches.find(x => x.row.useRight);
            if (from && to && sectionBody(from.row.ps, from.sec) === sectionBody(to.row.ps, to.sec) && sectionBody(to.row.ps, to.sec)) {
                from.row.ps.splice(from.sec.start, from.sec.end - from.sec.start);
                // Struck chapter titles have already been removed; a remaining chapter label needs review.
                m.decisions.push('Pasal ' + num + ' dipindahkan: bagian asal baris ' + from.row.index + ' dikeluarkan; rumusan tujuan baris ' + to.row.index + ' digunakan sekali. Keterangan: ' + note);
                to.row.moveHandled = true;
                from.row.moveHandled = true;
                for (const rr of m.rows)
                    if (new RegExp('Pasal\\s+' + num + '(?![0-9A-Za-z])', 'i').test(rr.note) && /pindah|reposisi|masuk\s+ke/i.test(rr.note)) rr.moveHandled = true;
            } else addIssue(m, {
                kind: 'review',
                title: 'Pemindahan Pasal ' + num + ' perlu diperiksa',
                message: 'Ada petunjuk pemindahan, tetapi rumusan asal dan tujuan tidak sama atau tidak lengkap. Periksa matriks lalu unggah ulang.',
                hard: true
            });
        } else if (matches.length !== 1) addIssue(m, {
            kind: 'review',
            title: 'Tujuan pemindahan Pasal ' + num + ' belum pasti',
            message: note,
            hard: true
        });
    }
    for (const {row, missing, resultTargets} of missingArticles(m.rows, articleSections)) {
        addIssue(m, {
            kind: 'review', row,
            title: resultTargets ? 'Konfirmasi perubahan sasaran pasal' : 'Periksa kelengkapan Pasal ' + missing.join(', '),
            message: resultTargets ? 'Pada baris ini, kolom awal membahas Pasal ' + missing.join(', ') + ', sedangkan kolom hasil membahas Pasal ' + resultTargets.join(', ') + '. Draft sudah mengikuti kolom hasil. Pastikan perubahan sasaran dan rumusannya memang disepakati, lalu pilih Sudah Saya Periksa. Catatan ini tidak menyatakan isi draft hilang.' : 'Nomor Pasal ini ada pada rancangan awal tetapi belum ditemukan di seluruh hasil terpilih. Periksa kemungkinan perubahan nomor, penghapusan, atau penyusunan ulang. Jika sudah sesuai, pilih Sudah Saya Periksa; jika belum, unggah matriks revisi. Aplikasi tidak menambah rumusan otomatis.'
        });
    }
    const seen = new Map();
    for (const row of m.rows) {
        for (const sec of articleSections(row.ps)) {
            if (seen.has(sec.num)) addIssue(m, {
                kind: 'review',
                title: 'Pasal ' + sec.num + ' muncul lebih dari sekali',
                message: 'Ditemukan pada baris ' + seen.get(sec.num) + ' dan ' + row.index + '. Petunjuk yang tersedia belum cukup untuk menentukan bagian yang dikeluarkan. Periksa matriks lalu unggah ulang.',
                hard: true
            });
            else seen.set(sec.num, row.index);
            const slice = row.ps.slice(sec.start + 1, sec.end),
                tail = slice.at(-1);
            if (tail && /sebagai berikut\s*:\s*$/i.test(tail.text)) {
                const original = articleSections(row.left).find(s => s.num === sec.num);
                const old = original ? row.left.slice(original.start, original.end) : [];
                addIssue(m, {
                    kind: 'incomplete',
                    hard: true,
                    title: 'Pasal ' + sec.num + ' belum memuat rincian',
                    message: 'Rumusan terpilih berakhir dengan “sebagai berikut:”. Rincian tidak ditambahkan otomatis dari kolom awal.',
                    row,
                    number: sec.num,
                    original: old
                });
            }
        }
        if (/hapus|hilang|pindah|reposisi|masuk\s+ke/i.test(row.note) && !row.moveHandled) addIssue(m, {
            kind: 'review',
            row,
            title: 'Petunjuk pada baris ' + row.index + ' perlu diperiksa',
            message: row.note
        });
        if (row.label.toLowerCase() === 'lampiran' && !row.ps.length) addIssue(m, {
            kind: 'review',
            row,
            title: 'Lampiran belum tersedia',
            message: 'Baris lampiran pada matriks kosong. Peta atau lampiran dari dokumen lain tidak disisipkan otomatis.'
        });
    }
    for (const row of m.rows) {
        let section = row.label,
            seenLetters = new Set();
        for (const p of row.ps) {
            if (/^Pasal\s+\d+/i.test(norm(p.text))) {
                section = norm(p.text);
                seenLetters.clear();
            }
            if (p.list?.kind === 'ayat') seenLetters.clear();
            if (p.list?.kind === 'letter') {
                const marker = p.list.marker;
                if (seenLetters.has(marker)) addIssue(m, {
                    kind: 'review',
                    row,
                    marker,
                    section,
                    title: 'Penanda ' + marker + ' berulang pada ' + section,
                    message: 'Baris ' + row.index + ' memuat penanda butir yang sama lebih dari sekali. Penanda dipertahankan sesuai sumber; periksa penomoran pada matriks.'
                });
                seenLetters.add(marker);
            }
        }
    }
    for (const row of m.rows)
        if ([...row.left, ...row.right].some(p => /^B\s{3,}BUPATI\b/.test(p.text))) addIssue(m, {
            kind: 'review',
            row,
            title: 'Karakter sebelum jabatan penandatangan perlu diperiksa',
            message: 'Baris ' + row.index + ' memuat huruf B terpisah sebelum BUPATI. Karakter dari sumber dipertahankan; periksa matriks jika huruf tersebut tidak dimaksudkan.'
        });
    for (const row of m.rows) {
        row.ps.forEach((block, blockIndex) => {
            if (isEditorialInstruction(block.text)) addIssue(m, {
                kind: 'editorial',
                row,
                blockIndex,
                title: 'Petunjuk penyunting berada dalam rumusan',
                message: 'Kalimat ini tampak sebagai instruksi penyunting, bukan rumusan peraturan. Pilih apakah catatan dikeluarkan dari draft atau dipertahankan sesuai matriks. Program tidak menyusun penjelasan atau materi yang diminta secara otomatis.'
            });
        });
    }
    const full = m.rows.map(r => plain(r.ps)).join('\n');
    for (const [word, replacement] of new Map(SPELLING_CORRECTIONS))
        if (replacementMatches(full, word).length) {
            const locations = [];
            for (const row of m.rows) {
                let article = '';
                for (let pi = 0; pi < row.ps.length; pi++) {
                    const p = row.ps[pi];
                    if (/^Pasal\s+\d+[A-Za-z]?$/i.test(norm(p.text))) article = norm(p.text);
                    if (replacementMatches(p.text, word).length) locations.push({
                        rowIndex: row.index,
                        blockIndex: pi,
                        label: 'Baris ' + row.index + (article ? ' · ' + article : ''),
                        context: p.text.slice(Math.max(0, p.text.indexOf(word) - 65), p.text.indexOf(word) + word.length + 85)
                    });
                }
            }
            addIssue(m, {
                kind: 'correction',
                title: '“' + word + '” → “' + replacement + '”',
                message: 'Pilih apakah rekomendasi ini diterapkan pada semua kemunculan yang ditampilkan.',
                word,
                replacement,
                locations
            });
        }
    if (/\.{3,}|…/.test(full)) addIssue(m, {
        kind: 'advisory',
        title: 'Ada tempat isian yang belum diisi',
        message: 'Nomor, tahun, tanggal, atau nama masih menggunakan titik/ellipsis. Nilainya perlu ditentukan perancang; tidak ada nilai pengganti yang dapat disarankan otomatis.'
    });

    await checkSpelling(m);
    attachIssueLocations(m);
    return m;
}

export {
    readMatrix
};
