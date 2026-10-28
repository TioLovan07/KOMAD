import { rememberWord, refreshSpelling } from './spell-check.js';
/** Kartu catatan, pilihan koreksi, dan navigasi lokasi. Tidak membentuk Word. */
import {
    $,
    el
} from './xml.js';
import {
    model
} from './state.js';
import {
    attachIssueLocations
} from './issue-locations.js';
import {
    highlightIssue
} from './issue-highlights.js';

function action(label, onClick, className = 'small') {
    const button = el('button', label, className);
    button.type = 'button';
    button.onclick = onClick;
    return button;
}

function showLocation(issue, index) {
    model.openIssueId = issue.id;
    document.querySelectorAll('.note').forEach(card => card.classList.remove('selected'));
    $('issue-' + issue.id)?.classList.add('selected');
    const location = issue.locations[index];
    const count = highlightIssue(issue, true, index);
    $('viewer-status').textContent = location.label + ' · ' + (location.sourceOnly ? 'Tersedia pada matriks sumber; belum ada teks yang dapat disorot di draft.' : count ? (location.related ? 'Bagian terkait disorot.' : 'Teks yang perlu diperiksa disorot.') : 'Lokasi sumber tersedia. Teks belum dapat dicocokkan secara tepat pada pratinjau.');
}

export function renderReview(onChange) {
    attachIssueLocations(model);
    $('issues').replaceChildren();
    for (const issue of model.issues) {
        const card = el('article', undefined, 'panel note' + (issue.resolved ? ' done' : ''));
        card.id = 'issue-' + issue.id;
        const status = issue.resolved ? (issue.choice === 'remove' ? 'CATATAN DIKELUARKAN' : issue.choice === 'apply' ? 'REKOMENDASI DITERAPKAN' : issue.choice === 'keep' ? 'SESUAI MATRIKS' : 'SUDAH DIPERIKSA') : issue.hard ? 'PERBAIKI MATRIKS' : 'PERLU DIPERIKSA';
        card.append(el('div', status, 'status'), el('h3', issue.kind === 'correction' ? 'Periksa ejaan: “' + issue.word + '”' : issue.title), el('p', issue.kind === 'correction' ? 'Bandingkan rekomendasi dengan konteks kalimat, lalu tentukan apakah perbaikan diterapkan.' : issue.message));
        const list = el('div', undefined, 'location-list');
        issue.locations.forEach((location, index) => {
            const item = el('div', undefined, 'location-item');
            item.append(el('strong', location.label), el('p', location.context || 'Bagian ini belum berisi teks.', 'quote'));
            if (location.sourceOnly) item.append(el('p', 'Bagian belum tersedia di draft atau hanya tercantum pada matriks sumber.', 'source-only'));
            item.append(action(location.sourceOnly ? 'Lihat Keterangan Lokasi' : issue.locations.length > 1 ? 'Lihat Bagian ' + (index + 1) : 'Lihat Bagian', () => showLocation(issue, index)));
            list.append(item);
        });
        card.append(list);
        const resolve = choice => {
            issue.resolved = true;
            if (choice) issue.choice = choice;
            onChange();
        };
        if (issue.kind === 'correction') {
            const suggestions = issue.suggestions?.length ? issue.suggestions : [issue.replacement].filter(Boolean);
            const group = el('div', undefined, 'suggestion-group');
            group.append(el('span', 'Rekomendasi perbaikan', 'suggestion-label'));
            if (suggestions.length === 1) group.append(el('div', suggestions[0], 'recommendation-value'));
            else for (const suggestion of suggestions) {
                const button = action(suggestion, () => {
                    issue.replacement = suggestion;
                    issue.resolved = false;
                    issue.choice = undefined;
                    onChange();
                }, 'suggestion-option' + (suggestion === issue.replacement ? ' selected' : ''));
                button.setAttribute('aria-pressed', String(suggestion === issue.replacement));
                group.append(button);
            }
            card.append(group);
            const apply = action('Terapkan Rekomendasi', () => resolve('apply'));
            apply.disabled = !issue.replacement;
            const keep = action('Biarkan Sesuai Matriks', () => resolve('keep'));
            apply.setAttribute('aria-pressed', String(issue.choice === 'apply'));
            keep.setAttribute('aria-pressed', String(issue.choice === 'keep'));
            card.append(apply, keep);
            if (issue.spelling) {
                const advanced = el('details', undefined, 'dictionary-options');
                advanced.append(el('summary', 'Pilihan kamus'), el('p', 'Gunakan hanya jika kata pada matriks memang benar, misalnya nama atau istilah khusus.'));
                advanced.append(action('Anggap Kata Ini Benar di Perangkat', () => {
                if (rememberWord(issue.word)) {
                    for (const other of model.issues.filter(i => i.spelling && i.word.toLowerCase() === issue.word.toLowerCase())) {
                        other.resolved = true; other.choice = 'keep';
                    }
                    resolve('keep');
                } else {
                    alert('Kamus perangkat tidak dapat disimpan. Gunakan Biarkan Sesuai Matriks untuk dokumen ini.');
                }
                }));
                card.append(advanced);
            }
        } else if (issue.kind === 'source-gap' && !issue.resolved) {
            const details=el('details',undefined,'source-comparison');
            details.append(el('summary','Bandingkan rumusan awal'));
            for(const block of issue.candidates) details.append(el('p',block.text,'quote'));
            card.append(details);
            card.append(action('Gunakan Lanjutan Awal',async()=>{
                if (issue.resolved) return;
                const selected=model;
                selected.previewReady=null; selected.outputBlob=null;
                document.dispatchEvent(new CustomEvent('komad:preview-state'));
                if(issue.choice!=='apply') issue.row.ps.splice(issue.insertAt,0,...issue.candidates.map(block=>({...block,node:block.node.cloneNode(true),list:block.list?{...block.list}:null})));
                model.decisions.push('Baris '+issue.row.index+' · Pengguna memilih menambahkan rumusan awal yang belum terpetakan.');
                issue.resolved=true;issue.choice='apply';
                await refreshSpelling(selected);
                if(model===selected) onChange();
            }),action('Hasil Sudah Sesuai, Tanpa Tambahan',()=>resolve('keep'),'text small'));
        } else if (issue.kind === 'editorial') {
            card.append(action('Keluarkan Catatan dari Draft', () => resolve('remove')),
                action('Biarkan Sesuai Matriks', () => resolve('keep'), 'text small'));
        } else if (issue.hard) {
            card.append(el('p', 'Perbaiki file sumber, lalu gunakan tombol Ganti File Matriks di bagian atas.', 'source-only'));
        } else if (!issue.resolved && issue.kind === 'annex') {
            card.append(action('Gunakan Isi Lampiran Awal', () => {
                issue.row.ps.push(...issue.annexBody.map(block => ({
                    ...block,
                    list: block.list ? {
                        ...block.list
                    } : null,
                    node: block.node.cloneNode(true)
                })));
                model.decisions.push('Baris ' + issue.row.index + ' · Judul lampiran hasil digunakan bersama isi setelah judul dari kolom awal, sesuai pilihan pengguna.');
                resolve('apply');
            }), action('Lengkapi Nanti', () => resolve('keep'), 'text small'));
        } else if (!issue.resolved) {
            card.append(action(issue.kind === 'advisory' ? 'Lengkapi Nanti' : 'Sudah Saya Periksa', () => resolve(issue.kind === 'advisory' ? 'keep' : undefined), 'text small'));
        }
        $('issues').append(card);
    }
    if (model.unconfirmedWords?.length) {
        const details = el('details', undefined, 'panel spelling-secondary');
        details.append(el('summary', model.unconfirmedWords.length + ' kata belum dikenali — opsional'));
        details.append(el('p', 'Bukan dugaan typo yang cukup kuat. Daftar ini tidak menghambat unduhan dan kata tidak diubah.'));
        details.append(el('p', model.unconfirmedWords.join(', '), 'quote'));
        $('issues').append(details);
    }
    if (!model.issues.length) $('issues').append(el('p', 'Tidak ada pengecualian yang terdeteksi. Tetap periksa seluruh draft sebelum digunakan.', 'panel issue-empty'));
    $('decisions').replaceChildren(...[
        ...model.decisions,
        ...model.issues.filter(issue => issue.kind === 'editorial' && issue.choice).map(issue =>
            (issue.choice === 'remove' ? 'Catatan dikeluarkan dari draft: ' : 'Catatan dipertahankan: ') +
            issue.row.ps[issue.blockIndex].text),
        ...model.issues.filter(issue => issue.kind === 'correction' && issue.choice).map(issue => issue.choice === 'apply' ? 'Diterapkan: “' + issue.word + '” menjadi “' + issue.replacement + '”.' : 'Dipertahankan: “' + issue.word + '” sesuai matriks.'),
    ].map(text => el('p', text)));
}
