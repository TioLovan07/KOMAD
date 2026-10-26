/** Serialisasi render mencegah pilihan lama menggantikan pilihan terbaru. */
import {
    $
} from './xml.js';
import {
    model
} from './state.js';
import {
    buildWord
} from './word-builder.js';
import {
    highlightIssue
} from './issue-highlights.js';

let renderChain = Promise.resolve();
let renderSerial = 0;

function notifyState() {
    document.dispatchEvent(new CustomEvent('komad:preview-state'));
}

function awaitViewer() {
    const frame = $('draft-frame');
    if (frame.contentWindow?.renderDraft) return Promise.resolve(frame.contentWindow);
    return new Promise((resolve, reject) => {
        const timeout = setTimeout(() => reject(Error('Pembaca halaman belum siap. Muat ulang halaman.')), 12000);
        frame.addEventListener('load', () => {
            clearTimeout(timeout);
            if (frame.contentWindow?.renderDraft) resolve(frame.contentWindow);
            else reject(Error('Pembaca halaman gagal dimuat.'));
        }, {
            once: true
        });
    });
}

export function cancelPreview() {
    renderSerial++;
}

export function schedulePreview() {
    const selected = model;
    const revision = (selected.revision || 0) + 1;
    selected.revision = revision;
    selected.previewReady = null;
    notifyState();
    $('preview-info').textContent = 'Menyiapkan seluruh halaman…';
    const serial = ++renderSerial;
    renderChain = renderChain.catch(() => {}).then(async () => {
        if (serial !== renderSerial || model !== selected) return;
        try {
            const viewer = await awaitViewer();
            const blob = await buildWord({
                allowPending: true
            });
            if (serial !== renderSerial || model !== selected) return;
            const info = await viewer.renderDraft(await blob.arrayBuffer());
            if (serial !== renderSerial || model !== selected) return;
            selected.outputBlob = blob;
            selected.previewReady = revision;
            $('preview-info').textContent = info.pages + ' halaman · ' + info.images + ' gambar';
            notifyState();
            highlightIssue(model.issues.find(issue => issue.id === model.openIssueId));
        } catch (error) {
            if (serial !== renderSerial) return;
            $('preview-info').textContent = 'Pratinjau gagal: ' + error.message;
            $('status').className = 'status-message error';
            $('status').textContent = 'Draft belum berhasil ditampilkan. ' + error.message;
            notifyState();
        }
    });
}

export function initializePreview() {
    $('zoom').onchange = () => $('draft-frame').contentWindow?.setDraftScale?.($('zoom').value);
    window.addEventListener('message', event => {
        const frame = $('draft-frame');
        if (event.origin !== location.origin || event.source !== frame.contentWindow) return;
        const data = event.data;
        if (data?.type === 'draft-viewer-size' && Number.isFinite(data.height)) {
            const height = Math.max(420, Math.min(500000, data.height));
            if (Math.abs(frame.getBoundingClientRect().height - height) > 2) frame.style.height = height + 'px';
        }
        if (data?.type === 'draft-viewer-location' && Number.isFinite(data.top)) {
            $('document').scrollTo({
                top: Math.max(0, data.top - 110),
                behavior: 'smooth'
            });
            if ($('document').getBoundingClientRect().top > innerHeight - 180) $('document').scrollIntoView({
                block: 'start',
                behavior: 'smooth'
            });
        }
    });
}
