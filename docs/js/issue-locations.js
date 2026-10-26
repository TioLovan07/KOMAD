/**
 * Mengaitkan setiap catatan dengan baris, blok, dan kutipan sumber.
 * Lokasi sumber yang tidak masuk draft ditandai sourceOnly: tidak dibuat sorotan palsu.
 */
import {
    norm
} from './xml.js';

function makeLocation(row, blockIndex, {
    word = '',
    related = false
} = {}) {
    const block = row.ps[blockIndex];
    const text = block?.text || '';
    let article = '';
    let documentPart = /^Baris \d+$/.test(row.label) ? '' : row.label;
    for (const item of row.ps.slice(0, blockIndex + 1)) {
        if (/^Pasal\s+\d+[A-Za-z]?$/i.test(norm(item.text))) article = norm(item.text);
        if (/^Ditetapkan di/i.test(norm(item.text))) documentPart = 'Penetapan';
        if (/^Diundangkan di/i.test(norm(item.text))) documentPart = 'Pengundangan';
        if (/^BERITA (DAERAH|NEGARA)/.test(norm(item.text))) documentPart = 'Berita daerah';
    }
    const position = word ? text.indexOf(word) : -1;
    const sentence = position >= 0 ? (text.slice(0, position).match(/[.;!?]\s+/g) || []).length + 1 : null;
    return {
        rowIndex: row.index,
        blockIndex,
        label: ['Baris matriks ' + row.index, article || documentPart || (/^NOMOR\b/.test(norm(text)) ? 'Judul peraturan' : 'Rumusan'), 'Paragraf ' + (blockIndex + 1), sentence && 'Kalimat ' + sentence].filter(Boolean).join(' · '),
        context: text.slice(Math.max(0, position - 70), position >= 0 ? position + word.length + 110 : 200),
        word,
        related,
        sourceOnly: !block || !text.trim(),
    };
}

function rowLocations(row, predicate = () => true, options = {}) {
    return row.ps.flatMap((block, index) => predicate(block, index) ? [makeLocation(row, index, options)] : []);
}

export function attachIssueLocations(documentModel) {
    for (const issue of documentModel.issues) {
        let locations = [];
        if (issue.kind === 'editorial') {
            locations = [makeLocation(issue.row, issue.blockIndex)];
        } else if (issue.kind === 'correction') {
            for (const row of documentModel.rows) {
                locations.push(...rowLocations(row, block => block.text.includes(issue.word), {
                    word: issue.word
                }));
            }
        } else if (issue.kind === 'advisory') {
            for (const row of documentModel.rows) {
                row.ps.forEach((block, index) => {
                    for (const match of block.text.matchAll(/\.{3,}|…/g)) {
                        const location = makeLocation(row, index, {
                            word: match[0]
                        });
                        location.sourceStart = match.index;
                        location.label += ' · Isian pada karakter ' + (match.index + 1);
                        location.context = block.text.slice(Math.max(0, match.index - 55), match.index + match[0].length + 75);
                        locations.push(location);
                    }
                });
            }
        } else if (issue.marker && issue.row) {
            let currentSection = issue.row.label;
            issue.row.ps.forEach((block, index) => {
                if (/^Pasal\s+\d+/i.test(norm(block.text))) currentSection = norm(block.text);
                if (block.list?.marker === issue.marker && currentSection === issue.section) {
                    locations.push(makeLocation(issue.row, index));
                }
            });
        } else if (/Karakter sebelum jabatan/.test(issue.title) && issue.row) {
            locations = rowLocations(issue.row, block => /^B\s+BUPATI\b/.test(block.text), {
                word: 'B'
            });
            if (!locations.length) {
                const source = [...issue.row.left, ...issue.row.right].find(block => /^B\s+BUPATI\b/.test(block.text));
                locations.push({
                    rowIndex: issue.row.index,
                    label: 'Baris matriks ' + issue.row.index + ' · Lampiran · Kolom awal',
                    context: source?.text || issue.message,
                    sourceOnly: true
                });
            }
        } else if (issue.row) {
            const articleNumber = issue.number;
            locations = rowLocations(issue.row, block => articleNumber ? norm(block.text) === 'Pasal ' + articleNumber : !!block.text.trim(), {
                related: true
            });
            if (!locations.length) locations.push({
                rowIndex: issue.row.index,
                label: 'Baris matriks ' + issue.row.index + ' · ' + issue.row.label,
                context: 'Bagian ini belum tersedia dalam rumusan terpilih.',
                sourceOnly: true
            });
            if (issue.kind === 'annex') locations = locations.slice(0, 2);
        } else if (issue.rowIndex) {
            const row = documentModel.rows.find(row => row.index === issue.rowIndex);
            locations.push({
                rowIndex: issue.rowIndex,
                label: 'Baris matriks ' + issue.rowIndex + ' · Kolom hasil',
                context: row?.right.map(block => block.text).join('\n') || issue.message,
                sourceOnly: true
            });
        } else {
            const article = issue.title.match(/Pasal\s+(\d+[A-Za-z]?)/i);
            for (const row of documentModel.rows) {
                if (article) {
                    locations.push(...rowLocations(row, block => norm(block.text) === 'Pasal ' + article[1], {
                        related: true
                    }));
                } else if (/Track Changes/.test(issue.title)) {
                    locations.push(...rowLocations(row, block => /<w:(ins|del)\b/.test(block.source || ''), {
                        related: true
                    }));
                }
            }
            if (!locations.length) locations.push({
                label: 'Matriks · Pemeriksaan tingkat dokumen',
                context: issue.message,
                sourceOnly: true
            });
        }
        for (const location of locations) {
            if (documentModel.issues.some(note => note.kind === 'editorial' && note.choice === 'remove' &&
                    note.row.index === location.rowIndex && note.blockIndex === location.blockIndex)) {
                location.sourceOnly = true;
            }
        }
        issue.locations = locations;
    }
}
