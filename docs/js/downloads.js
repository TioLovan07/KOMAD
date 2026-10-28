/** Ekspor hanya memakai Blob revisi yang sudah selesai ditampilkan pada pratinjau. */
import {
    $,
    el
} from './xml.js';
import {
    model
} from './state.js';
import {
    pending
} from './draft-model.js';

const objectUrls = new Set();
export function clearDownloadLinks() {
    for (const url of objectUrls) URL.revokeObjectURL(url);
    objectUrls.clear();
    $('download-links').replaceChildren();
}

function saveBlob(blob, filename) {
    clearDownloadLinks();
    const link = el('a');
    link.hidden = true;
    link.href = URL.createObjectURL(blob);
    objectUrls.add(link.href);
    link.download = filename;
    $('download-links').append(link);
    link.click();
    link.remove();
}

export function isDraftReady() {
    return !!model && !pending().length && model.previewReady === model.revision && !!model.outputBlob;
}

export function downloadWord() {
    if (!isDraftReady()) return;
    saveBlob(model.outputBlob, 'Draft.docx');
    $('status').textContent = 'Unduhan Word dimulai. Jika belum tersimpan, tekan Unduh Draft Word kembali.';
}

export function downloadReport() {
    if (!model) return;
    const lines = ['CATATAN PENYUSUNAN DRAFT', model.name, 'KOMAD 1.3.2', '', 'KEPUTUSAN PENYUSUNAN', ...model.decisions, '', 'PEMERIKSAAN'];
    for (const issue of model.issues) {
        lines.push((!issue.resolved ? 'BELUM DIPERIKSA' : issue.choice === 'remove' ? 'CATATAN DIKELUARKAN' : issue.choice === 'apply' ? 'REKOMENDASI DITERAPKAN' : issue.choice === 'keep' ? 'SESUAI MATRIKS' : 'SUDAH DIPERIKSA') + ' · ' + issue.title, issue.message);
        for (const location of issue.locations || []) lines.push(location.label + ' · ' + location.context);
        lines.push('');
    }
    saveBlob(new Blob([lines.join('\n')], {
        type: 'text/plain;charset=utf-8'
    }), 'Catatan_Penyusunan_Draft.txt');
}
