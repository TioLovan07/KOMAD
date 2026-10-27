'use strict';
let scaleMode = '100',
    pageWidths = [];
const pages = document.getElementById('pages');

function resize() {
    const height = Math.ceil(pages.getBoundingClientRect().height + 8);
    parent.postMessage({
        type: 'draft-viewer-size',
        height
    }, parent.location.origin === 'null' ? '*' : parent.location.origin);
}

function setScale(value) {
    scaleMode = String(value);
    const wrapper = pages.querySelector('.worddoc-wrapper');
    if (!wrapper) return;
    const width = Math.max(1, ...pageWidths);
    const zoom = value === 'fit' ? Math.min(1, (document.documentElement.clientWidth - 44) / width) : Number(value) / 100;
    wrapper.style.zoom = String(Math.max(.25, zoom));
    requestAnimationFrame(() => requestAnimationFrame(resize));
    return Math.round(zoom * 100);
}
window.setDraftScale = setScale;
window.renderDraft = async function(buffer) {
    document.getElementById('empty').hidden = true;
    await docx.renderAsync(new Uint8Array(buffer), pages, document.getElementById('styles'), {
        className: 'worddoc',
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        ignoreFonts: false,
        breakPages: true,
        paginate: true,
        experimental: true,
        renderHeaders: true,
        renderFooters: true,
        renderFootnotes: true,
        renderEndnotes: true,
        renderChanges: false,
        renderComments: false,
        renderAltChunks: false,
        ignoreLastRenderedPageBreak: true,
        useBase64URL: false
    });
    await Promise.all(Array.from(pages.querySelectorAll('img')).map(img => img.complete ? Promise.resolve() : new Promise(resolve => {
        img.onload = resolve;
        img.onerror = resolve;
    })));
    // Links are preview text, never executable/navigation targets.
    for (const link of pages.querySelectorAll('a[href]')) link.removeAttribute('href');
    for (const p of pages.querySelectorAll('article p')) {
        if (getComputedStyle(p).textAlign === 'justify') p.style.textAlignLast = 'start';
    }
    await document.fonts.ready;
    // docx renderer leaves explicit tab widths in inline spans. Keep those intact.
    const sections = Array.from(pages.querySelectorAll('section.worddoc'));
    pageWidths = sections.map(s => s.offsetWidth);
    sections.forEach((s, i) => {
        s.dataset.pageLabel = 'Halaman ' + (i + 1) + ' / ' + sections.length;
    });
    const zoom = setScale(scaleMode);
    resize();
    return {
        pages: sections.length,
        images: pages.querySelectorAll('img').length,
        zoom
    };
};
window.clearDraftHighlights = function() {
    hideTooltip();
    for (const m of pages.querySelectorAll('.draft-highlight')) m.replaceWith(document.createTextNode(m.textContent));
    pages.normalize();
};

function normalizedWithMap(text) {
    let s = '',
        map = [];
    for (let i = 0; i < text.length; i++) {
        if (/\s/.test(text[i])) {
            if (s && !s.endsWith(' ')) {
                s += ' ';
                map.push(i);
            }
        } else {
            s += text[i];
            map.push(i);
        }
    }
    return {
        text: s,
        map
    };
}

function paintRange(p, start, end) {
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    const parts = [];
    let n, offset = 0;
    while ((n = walker.nextNode())) {
        parts.push({
            node: n,
            start: offset,
            end: offset + n.textContent.length
        });
        offset += n.textContent.length;
    }
    const a = parts.find(q => q.start <= start && q.end > start),
        b = parts.find(q => q.start < end && q.end >= end);
    if (!a || !b) return null;
    const r = document.createRange();
    r.setStart(a.node, start - a.start);
    r.setEnd(b.node, end - b.start);
    const mark = document.createElement('span');
    mark.className = 'draft-highlight';
    mark.appendChild(r.extractContents());
    r.insertNode(mark);
    return mark;
}
window.highlightDraft = function(targets, scroll) {
    window.clearDraftHighlights();
    let first = null;
    for (const target of targets) {
        const anchors = Array.from(pages.querySelectorAll('[id="' + target.anchor + '"],[id^="' + target.anchor + '__"]'));
        if (!anchors.length) continue;
        const all = Array.from(pages.querySelectorAll('article p'));
        const candidates = [];
        for (const anchor of anchors) {
            const firstP = anchor.closest('p');
            if (!firstP) continue;
            const startAt = all.indexOf(firstP);
            for (let i = startAt; i < all.length; i++) {
                if (i > startAt && all[i].querySelector('[id^="p3h_"]')) break;
                if (!candidates.includes(all[i])) candidates.push(all[i]);
            }
        }

        const source = normalizedWithMap(target.text);
        for (const range of target.ranges) {
            const sourceStart = normalizedWithMap(target.text.slice(0, range.start)).text.length;
            const word = normalizedWithMap(target.text.slice(range.start, range.end)).text;
            const before = source.text.slice(Math.max(0, sourceStart - 28), sourceStart).trim(),
                after = source.text.slice(sourceStart + word.length, sourceStart + word.length + 28).trim();
            let best = null;
            for (const p of candidates) {
                const actual = normalizedWithMap(p.textContent);
                let pos = -1;
                while ((pos = actual.text.indexOf(word, pos + 1)) >= 0) {
                    const left = actual.text.slice(Math.max(0, pos - 28), pos).trim(),
                        right = actual.text.slice(pos + word.length, pos + word.length + 28).trim();
                    let score = 0;
                    for (let k = 1; k <= Math.min(before.length, left.length); k++)
                        if (before.slice(-k) === left.slice(-k)) score = k;
                    for (let k = 1; k <= Math.min(after.length, right.length); k++)
                        if (after.slice(0, k) === right.slice(0, k)) score += 1;
                    if (!best || score > best.score) best = {
                        p,
                        score,
                        start: actual.map[pos],
                        end: actual.map[pos + word.length - 1] + 1
                    };
                }
            }
            if (best) {
                const mark = paintRange(best.p, best.start, best.end);
                if (mark) {
                    mark.tabIndex = 0;
                    mark.dataset.note = target.tooltip || 'Bagian ini perlu diperiksa.';
                    mark.setAttribute('role', 'button');
                    mark.setAttribute('aria-label', mark.textContent + '. ' + mark.dataset.note);
                }
                first = first || mark;
            }
        }
    }
    if (scroll && first) {
        const rect = first.getBoundingClientRect();
        parent.postMessage({
            type: 'draft-viewer-location',
            top: rect.top + window.scrollY
        }, parent.location.origin === 'null' ? '*' : parent.location.origin);
    }
    return pages.querySelectorAll('.draft-highlight').length;
};
window.addEventListener('resize', () => {
    if (scaleMode === 'fit') setScale(scaleMode);
    resize();
});
// Tooltip berada di luar paragraf Word agar tidak mengubah pemenggalan halaman.
const notePopup = document.createElement('div');
notePopup.id = 'note-popup';
notePopup.setAttribute('role', 'tooltip');
notePopup.hidden = true;
document.body.appendChild(notePopup);
let pinnedMark = null;

function hideTooltip() {
    notePopup.hidden = true;
    pinnedMark = null;
}

function showTooltip(mark, pinned = false) {
    notePopup.textContent = mark.dataset.note;
    notePopup.hidden = false;
    if (pinned) pinnedMark = mark;
    const rect = mark.getBoundingClientRect();
    const width = Math.min(340, innerWidth - 24);
    notePopup.style.width = width + 'px';
    notePopup.style.left = Math.max(12, Math.min(rect.left, innerWidth - width - 12)) + 'px';
    const height = notePopup.getBoundingClientRect().height;
    notePopup.style.top = Math.max(12, rect.top > height + 18 ? rect.top - height - 10 : Math.min(innerHeight - height - 12, rect.bottom + 10)) + 'px';
}
pages.addEventListener('pointerover', event => {
    const mark = event.target.closest('.draft-highlight');
    if (mark) showTooltip(mark);
});
pages.addEventListener('pointerout', event => {
    if (!pinnedMark && !notePopup.matches(':hover')) notePopup.hidden = true;
});
pages.addEventListener('focusin', event => {
    const mark = event.target.closest('.draft-highlight');
    if (mark) showTooltip(mark);
});
pages.addEventListener('focusout', () => {
    if (!pinnedMark) notePopup.hidden = true;
});
pages.addEventListener('click', event => {
    const mark = event.target.closest('.draft-highlight');
    if (mark) showTooltip(mark, true);
    else hideTooltip();
});
pages.addEventListener('keydown', event => {
    const mark = event.target.closest('.draft-highlight');
    if (mark && ['Enter', ' '].includes(event.key)) {
        event.preventDefault();
        showTooltip(mark, true);
    }
    if (event.key === 'Escape') hideTooltip();
});
window.addEventListener('scroll', hideTooltip);
