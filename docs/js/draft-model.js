/**
 * Rumusan efektif dan koreksi lintas run Word. ZIP matriks asli tidak dimutasi.
 */
import {
    nodes,
    xml,
    we,
    blockText,
    paragraphs
} from './xml.js';
import {
    model
} from './state.js';

function pending() {
    return model.issues.filter(x => !x.resolved);
}

function replacementMatches(s, word) {
    const matches = [];
    let off = 0;
    while ((off = s.indexOf(word, off)) >= 0) {
        const before = s[off - 1] || '',
            after = s[off + word.length] || '';
        if (!/[\p{L}\p{N}_]/u.test(before) && !/[\p{L}\p{N}_]/u.test(after)) matches.push(off);
        off += word.length;
    }
    return matches;
}

function replaceRuns(node, word, replacement) {
    const paragraphs = node.localName === 'p' ? [node] : nodes(node, 'p');
    for (const p of paragraphs) {
        const ts = nodes(p, 't');
        let full = '',
            ranges = [];
        for (const t of ts) {
            ranges.push({
                t,
                start: full.length,
                end: full.length + t.textContent.length
            });
            full += t.textContent;
        }
        for (const start of replacementMatches(full, word).reverse()) {
            const end = start + word.length;
            let first = true;
            for (const q of ranges) {
                if (q.end <= start || q.start >= end) continue;
                const a = Math.max(0, start - q.start),
                    b = Math.min(q.end - q.start, end - q.start);
                q.t.textContent = q.t.textContent.slice(0, a) + (first ? replacement : '') + q.t.textContent.slice(b);
                q.t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
                first = false;
            }
        }
    }
}

function effectiveRows() {
    return model.rows.map(row => ({
        ...row,
        ps: row.ps.map(p => {
            const node = p.node.cloneNode(true);
            for (const issue of model.issues)
                if (issue.kind === 'correction' && issue.choice === 'apply') replaceRuns(node, issue.word, issue.replacement);
            return {
                ...p,
                list: p.list ? {
                    ...p.list
                } : null,
                node,
                text: blockText(node)
            };
        })
    }));
}

function makePara(t) {
    const d = model.doc;
    const p = we(d, 'p'),
        r = we(d, 'r'),
        tx = we(d, 't');
    tx.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    tx.textContent = t;
    r.appendChild(tx);
    p.appendChild(r);
    return {
        node: p,
        text: t,
        source: xml(p)
    };
}

export {
    pending,
    replacementMatches,
    replaceRuns,
    effectiveRows,
    makePara
};
