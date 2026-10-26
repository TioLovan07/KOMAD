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
        card.append(el('div', status, 'status'), el('h3', issue.title), el('p', issue.message));
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
            const apply = action('Terapkan Rekomendasi', () => resolve('apply'));
            const keep = action('Biarkan Sesuai Matriks', () => resolve('keep'), 'text small');
            apply.setAttribute('aria-pressed', String(issue.choice === 'apply'));
            keep.setAttribute('aria-pressed', String(issue.choice === 'keep'));
            card.append(apply, keep);
        } else if (issue.kind === 'editorial') {
            card.append(action('Keluarkan Catatan dari Draft', () => resolve('remove')),
                action('Biarkan Sesuai Matriks', () => resolve('keep'), 'text small'));
        } else if (issue.hard) {
            card.append(action('Pilih Matriks Revisi', () => $('file').click()));
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
    if (!model.issues.length) $('issues').append(el('p', 'Tidak ada pengecualian yang terdeteksi. Tetap periksa seluruh draft sebelum digunakan.', 'panel issue-empty'));
    $('decisions').replaceChildren(...[
        ...model.decisions,
        ...model.issues.filter(issue => issue.kind === 'editorial' && issue.choice).map(issue =>
            (issue.choice === 'remove' ? 'Catatan dikeluarkan dari draft: ' : 'Catatan dipertahankan: ') +
            issue.row.ps[issue.blockIndex].text),
        ...model.issues.filter(issue => issue.kind === 'correction' && issue.choice).map(issue => issue.choice === 'apply' ? 'Diterapkan: “' + issue.word + '” menjadi “' + issue.replacement + '”.' : 'Dipertahankan: “' + issue.word + '” sesuai matriks.'),
    ].map(text => el('p', text)));
}
