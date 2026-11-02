import {keepInstruction} from './control-text.js';
import {isAnnexInput,annexTitles} from './annex-titles.js';
import {documentMetadata} from './document-layout.js';
import {selectEnactment} from './enactment.js';
import {selectStructurePlan} from './structure-plan.js';
import {annotateLists} from './list-layout.js';
import {prepareWordFile} from './legacy-import.js';
import {experimentalSelect} from './merge-rumusan.js';
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
    we,
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
    const originalName=file.name;
    const imported=await prepareWordFile(file);
    file=imported.file;
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
        name: originalName,
        rows: [],
        issues: [],
        decisions: [],
        filename: 'Draft.docx'
    };
    if(imported.legacy) {
        m.legacyImport={recoveredRows:imported.recoveredRows,images:imported.images};
        if(imported.skippedAnnex)m.decisions.push('Pembacaan DOC berhenti sesudah baris judul lampiran yang berada setelah penutup matriks. Isi lampiran tidak disalin; halaman portrait berjudul disiapkan sesuai ruang lingkup aplikasi.');
        m.decisions.push('DOC 97–2003 dibaca secara lokal; struktur tabel dan run teks diteruskan ke pembaca yang sama.');
        addIssue(m,{kind:'review',title:'Periksa hasil pembacaan DOC 97–2003',message:'Teks, daftar, dan tabel DOC dibaca di perangkat. Gambar WMF/EMF, bentuk, dan layout lampiran tertentu belum didukung. Periksa juga penomoran otomatis dan bandingkan sumber; gunakan .docx jika unsur tersebut diperlukan.'});
    }
    if (tables.length > 1) addIssue(m, {
        title: 'Materi di luar matriks utama belum digabungkan',
        message: (tables.length-1)+' tabel di luar matriks utama belum dimasukkan otomatis. Tabel tersebut dapat berisi lampiran atau lanjutan naskah. Bandingkan sumber dan gabungkan di Word jika diperlukan; jangan menganggap draft sudah mencakup seluruh lampiran.',
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
        const selectionOptions = {m,rowIndex:i+1,prefix:labels.prefix,note};
        const annexInput=isAnnexInput(labels.label,left,right);
        if(!annexInput && right.some(p=>keepInstruction(p)?.kind==='confirm-keep')) addIssue(m,{
            kind:'review',rowIndex:i+1,title:'Konfirmasi petunjuk Tetsp pada baris '+(i+1),
            message:'Kolom hasil memuat Tetsp, yang diduga berarti Tetap. Petunjuk tidak dicetak. Isi awal digunakan pada cakupan yang dapat dipastikan; pastikan maksud petunjuk sebelum menyetujui.'
        });
        const keepOnly=!annexInput && right.length>0 && right.every(p=>keepInstruction(p));
        if(keepOnly&&!left.length) addIssue(m,{kind:'review',hard:true,rowIndex:i+1,
            title:'Petunjuk Tetap tidak memiliki isi awal pada baris '+(i+1),
            message:'Hasil hanya memuat petunjuk Tetap/Tetsp, tetapi rumusan awal kosong. Sediakan rumusan yang harus dipertahankan; aplikasi tidak mengarang isi.'});

        const annexAmendment=left.some(p=>/^Pasal I$/i.test(norm(p.text))) && right.length===2 && /^Pasal I$/i.test(norm(right[0].text)) && /^Lampiran\s+Peraturan\b/i.test(norm(right[1].text)) && /diubah.*sebagaimana.*Lampiran/i.test(norm(right[1].text));
        if(annexAmendment)addIssue(m,{kind:'review',rowIndex:i+1,title:'Konfirmasi penggantian rumusan Pasal I',message:'Kolom hasil mengganti rumusan Pasal I menjadi perubahan lampiran. Rincian Pasal 14 dari kolom awal tidak dicetak karena tidak ditulis lagi pada rumusan pengganti. Pastikan keputusan perubahan lampiran ini sesuai hasil harmonisasi.',action:'Bandingkan seluruh rumusan Pasal I pada kolom hasil dengan draft. Konfirmasikan penggantian cakupannya; jangan mengembalikan Pasal 14 lama jika memang diganti oleh rumusan hasil.'});
        const selected = annexAmendment ? {ps:right.map(matrixCloneBlock),useRight:true,rule:'Rumusan Pasal I hasil yang khusus mengubah lampiran digunakan seluruhnya.'} : annexInput ? {ps:annexTitles(left,right,documentMetadata(m.rows),labels.label).flatMap(lines=>lines.map(text=>{const node=we(d,'p'),r=we(d,'r'),t=we(d,'t');t.textContent=text;r.appendChild(t);node.appendChild(r);return {node,text,list:null};})),useRight:false,rule:'Lampiran hanya berupa halaman portrait dengan judul; isi sumber tidak disalin.'} : keepOnly ? {ps:left.map(matrixCloneBlock),useRight:false,rule:'Hasil hanya memuat petunjuk Tetap/Tetsp; seluruh rumusan awal pada baris dipertahankan.'} : selectEnactment(left,right,selectionOptions) || selectStructurePlan(left,right,selectionOptions,experimentalSelect) || experimentalSelect(left,right,selectionOptions) || matrixSelectRumusan(left, right, {
            m,
            rowIndex: i + 1,
            prefix: labels.prefix
        });
        // Preserve a picture-only emblem in the opening identity row when the
        // result repeats the identity text but omits its illustration. No body images
        // or text are copied, and explicit removal instructions are respected.
        const hasPicture = p => nodes(p.node,'drawing').length || nodes(p.node,'pict').length;
        if (!m.rows.length && selected.useRight &&
            selected.ps.some(p=>/^(?:BUPATI|WALI\s*KOTA|GUBERNUR)\b/i.test(norm(p.text))) &&
            !selected.ps.some(hasPicture) && !(/\b(?:hapus|dihapus|hilangkan|tanpa)\b/i.test(note) && /\b(?:gambar|lambang|garuda)\b/i.test(note))) {
            const openingPictures=left.slice(0,left.findIndex(p=>p.text.trim())).filter(p=>!p.text.trim()&&hasPicture(p));
            if(openingPictures.length===1) {
                selected.ps.unshift(matrixCloneBlock(openingPictures[0]));
                m.decisions.push('Gambar pembuka dari kolom awal dipertahankan; teks identitas mengikuti hasil penyempurnaan.');
            }
        }
        // Some templates place the decision heading in a label cell of its own.
        if(/MEMUTUSKAN\s*:/i.test(labels.label) &&
            ![...m.rows.flatMap(r=>r.ps),...selected.ps].some(p=>/^MEMUTUSKAN\s*:/i.test(norm(p.text)))) {
            const node=we(d,'p'),run=we(d,'r'),text=we(d,'t');text.textContent='MEMUTUSKAN:';run.appendChild(text);node.appendChild(run);
            selected.ps.unshift({node,text:'MEMUTUSKAN:',list:null});
        }
        const row = {
            index: i + 1,
            structurePlan: !!selected.structurePlan,
            left,
            right,
            note,
            useRight: selected.useRight,
            ps: selected.ps,
            label: labels.label || ('Baris ' + (i + 1)),
            prefix: labels.prefix
        };
        if (row.ps.some(p=>/^PENJELASAN$/i.test(norm(p.text)))) row.sectionKind='explanation';
        if (annexInput) {row.label='Lampiran';row.note='';row.annexTitleOnly=true;}
        m.rows.push(row);
        m.decisions.push('Baris ' + row.index + ' · ' + selected.rule + (nodes(selected.useRight ? rc : lc, 'strike').length ? ' Teks dicoret dikeluarkan.' : ''));
        if (!row.annexTitleOnly && row.label === 'Lampiran' && right.length && !right.some(p => p.node.localName === 'tbl') && left.some(p => p.node.localName === 'tbl')) {
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
        if (amendment && /^(?:\d+[.)]\s*)?(?:Ketentuan\s+(?:Pasal|huruf|Lampiran|(?:ayat|angka)\b.*\bdiubah)|Di\s*antara\s+(?:BAB|Pasal))/i.test(norm(p.text))) {
            const manual = p.list?.prefixLength || 0;
            p.amendmentInstruction = true;
            const original=p.list?.marker || norm(p.text).match(/^(\d+)[.)]\s/)?.[1];
            changeNumber++;
            p.list = {...p.list, marker: original ? original.replace(/[.)]$/, '') + '.' : changeNumber + '.', kind:'number', level:0, prefixLength:manual};
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
    for (const {row, missing, resultTargets} of missingArticles(m.rows.filter(row=>row.sectionKind!=='explanation'), articleSections)) {
        addIssue(m, {
            kind: 'review', row,
            title: resultTargets ? 'Konfirmasi perubahan sasaran pasal' : 'Periksa kelengkapan Pasal ' + missing.join(', '),
            message: resultTargets ? 'Pada baris ini, kolom awal membahas Pasal ' + missing.join(', ') + ', sedangkan kolom hasil membahas Pasal ' + resultTargets.join(', ') + '. Draft sudah mengikuti kolom hasil. Pastikan perubahan sasaran dan rumusannya memang disepakati, lalu pilih Sudah Saya Periksa. Catatan ini tidak menyatakan isi draft hilang.' : 'Nomor Pasal ini ada pada rancangan awal tetapi belum ditemukan di seluruh hasil terpilih. Periksa kemungkinan perubahan nomor, penghapusan, atau penyusunan ulang. Jika sudah sesuai, pilih Sudah Saya Periksa; jika belum, unggah matriks revisi. Aplikasi tidak menambah rumusan otomatis.'
        });
    }
    for(const row of m.rows) for(const blocks of [row.left,row.ps]) {
        let outline=false;
        for(const p of blocks) {
            if(/^Pasal\s+\d+/i.test(norm(p.text)))outline=false;
            if(/sistematika\s+sebagai berikut\s*:/i.test(norm(p.text))) {outline=true;continue;}
            if(outline&&p.list?.kind==='ayat')outline=false;
            if(outline)p.inlineStructure=true;
        }
    }
    const seen = new Map();
    for (const row of m.rows) {
        if(row.sectionKind==='explanation')continue;
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

    for(const row of m.rows) row.ps.forEach((p,blockIndex)=>{
        if(keepInstruction(p))addIssue(m,{kind:'review',hard:true,row,rowIndex:row.index,blockIndex,
            title:'Petunjuk belum terselesaikan pada baris '+row.index,
            message:'Paragraf ini masih hanya berisi Tetap/Tetsp. Isi awal belum dipulihkan secara pasti. Unduhan ditahan; jangan menghapus petunjuk tanpa memulihkan rumusan yang dimaksud.'});
    });
    annotateLists(m);
    await checkSpelling(m);
    attachIssueLocations(m);
    return m;
}

export {
    readMatrix
};
