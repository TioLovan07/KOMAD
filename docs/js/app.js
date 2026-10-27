/** Titik masuk antarmuka KOMAD. Alur: pilih matriks → periksa → unduh Word. */
import {
    $,
    el,
    norm
} from './xml.js';
import {
    model,
    setModel
} from './state.js';
import {
    readMatrix
} from './matrix-parser.js';
import {
    pending,
    effectiveRows
} from './draft-model.js';
import {
    renderReview
} from './review-panel.js';
import {
    initializePreview,
    schedulePreview,
    cancelPreview
} from './preview.js';
import {
    clearDownloadLinks,
    isDraftReady,
    downloadWord,
    downloadReport
} from './downloads.js';

let currentStage = 1;
let importSerial = 0;

function setStatus(message, error = false) {
    $('status').textContent = message;
    $('status').className = 'status-message' + (error ? ' error' : '');
}

function showStage(number) {
    if (number > 1 && !model) return;
    if (number === 3 && !isDraftReady()) {
        setStatus('Selesaikan seluruh catatan dan tunggu pratinjau selesai sebelum melanjutkan ke unduhan.');
        return;
    }
    currentStage = number;
    $('stage-1').hidden = number !== 1;
    $('work').hidden = number === 1;
    $('stage-2').hidden = number !== 2;
    $('stage-3').hidden = number !== 3;
    document.querySelectorAll('.step').forEach(button => {
        if (Number(button.dataset.stage) === number) button.setAttribute('aria-current', 'step');
        else button.removeAttribute('aria-current');
    });
    if (number === 3) renderDownloadSummary();
    window.scrollTo({
        top: 0,
        behavior: 'smooth'
    });
}

function updateReadiness() {
    const count = model ? pending().length : 0;
    const ready = isDraftReady();
    $('to-download').disabled = !ready;
    $('download').disabled = !ready;
    $('count').textContent = String(count);
    $('ready').textContent = !model ? '' : count ? count + ' catatan perlu ditangani sebelum unduhan.' : ready ? 'Seluruh catatan telah ditangani. Periksa draft, lalu lanjutkan ke unduhan.' : 'Menunggu seluruh halaman selesai ditampilkan…';
    document.querySelector('[data-stage="2"]').disabled = !model;
    document.querySelector('[data-stage="3"]').disabled = !ready;
}

function renderDownloadSummary() {
    $('download-filename').textContent = model.filename;
    const remainingData = model.issues.filter(issue => issue.kind === 'advisory').length;
    const correctionCount = model.issues.filter(issue => issue.kind === 'correction' && issue.choice === 'apply').length;
    const entries = [
        ['Status pemeriksaan', pending().length ? pending().length + ' catatan belum ditangani' : 'Seluruh catatan ditangani'],
        ['Koreksi ejaan diterapkan', correctionCount + ' rekomendasi'],
        ['Tempat isian', remainingData ? 'Masih perlu dilengkapi sesuai sumber' : 'Tidak ada penanda kosong terdeteksi'],
        ['Format dokumen', 'Bookman Old Style 12 pt · F4 potret · halaman pertama: teks 8 cm dari atas; halaman lanjutan dan sisi lain 2,5 cm'],
        ['Lampiran', 'Halaman potret di akhir; periksa judul dan lengkapi isi yang belum tersedia'],
    ];
    $('download-summary').replaceChildren(...entries.map(([label, value]) => {
        const row = el('div', undefined, 'summary-row');
        row.append(el('span', label), el('strong', value));
        return row;
    }));
}

function renderDocument() {
    $('filename').textContent = model.name;
    const articleCount = effectiveRows().flatMap(row => row.ps).filter(block => /^Pasal\s+\d+[A-Za-z]?$/i.test(norm(block.text))).length;
    $('stats').replaceChildren(...[model.rows.length + ' baris matriks', articleCount + ' pasal'].map(text => el('span', text)));
    renderReview(() => {
        clearDownloadLinks();
        renderDocument();
        setStatus('Pilihan dicatat. Pratinjau dan draft Word diperbarui.');
    });
    schedulePreview();
    updateReadiness();
}

async function importMatrix(file) {
    if (!file) return;
    const serial = ++importSerial;
    $('pick').disabled = true;
    setStatus('Membaca matriks dan memeriksa rumusan…');
    try {
        const nextModel = await readMatrix(file);
        if (serial !== importSerial) return;
        cancelPreview();
        clearDownloadLinks();
        setModel(nextModel);
        renderDocument();
        showStage(2);
        setStatus('Matriks terbaca. Periksa catatan dan seluruh halaman draft.');
    } catch (error) {
        if (serial !== importSerial) return;
        // Draft yang sedang dikerjakan tetap tersedia jika unggahan revisi gagal.
        setStatus('Matriks belum dapat diproses. ' + error.message, true);
    } finally {
        if (serial === importSerial) $('pick').disabled = false;
    }
}

function resetDocument() {
    importSerial++;
    cancelPreview();
    clearDownloadLinks();
    setModel(null);
    $('file').value = '';
    $('pick').disabled = false;
    setStatus('');
    updateReadiness();
    showStage(1);
}

function initializeHelp() {
    $('help').onclick = () => $('guide').showModal();
    $('example').onclick = () => $('example-dialog').showModal();
    $('guide-example').onclick = () => {
        $('guide').close();
        $('example-dialog').showModal();
    };
    $('install-info').onclick = () => $('offline').showModal();
    document.querySelectorAll('.close').forEach(button => button.onclick = () => button.closest('dialog').close());
    $('try-example').onclick = async () => {
        $('example-dialog').close();
        setStatus('Membuka matriks contoh…');
        try {
            const response = await fetch('./assets/Contoh_Matriks_KOMAD.docx');
            if (!response.ok) throw Error('Contoh belum tersedia.');
            await importMatrix(new File([await response.blob()], 'Contoh_Matriks_KOMAD.docx', {
                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            }));
        } catch (error) {
            setStatus('Contoh belum dapat dibuka. ' + error.message, true);
        }
    };
}

function initializeInterface() {
    $('pick').onclick = () => $('file').click();
    $('file').onchange = event => importMatrix(event.target.files[0]);
    $('revise').onclick = () => {
        $('file').value = '';
        $('file').click();
    };
    $('reset').onclick = resetDocument;
    $('new-doc').onclick = resetDocument;
    $('to-download').onclick = () => showStage(3);
    $('back-review').onclick = () => showStage(2);
    $('download').onclick = downloadWord;
    $('report').onclick = downloadReport;
    document.querySelectorAll('.step').forEach(button => button.onclick = () => showStage(Number(button.dataset.stage)));
    $('text-size').onchange = () => {
        const percent = Number($('text-size').value);
        document.documentElement.style.setProperty('--text-scale', percent / 100);
        $('small-text-hint').hidden = percent >= 75;
    };
    document.addEventListener('komad:preview-state', updateReadiness);
    window.addEventListener('beforeunload', event => {
        if (model) {
            event.preventDefault();
            event.returnValue = '';
        }
    });
    initializeHelp();
    initializePreview();
    updateReadiness();
}

initializeInterface();
