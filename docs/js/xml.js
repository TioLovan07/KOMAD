/**
 * Utilitas DOM dan OOXML bersama. Tidak memuat keadaan dokumen atau antarmuka.
 */
'use strict';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const $ = id => document.getElementById(id);
const nodes = (e, n) => Array.from(e.getElementsByTagNameNS(W, n));
const direct = (e, n) => Array.from(e.children || []).filter(x => x.namespaceURI === W && x.localName === n);
const val = (e, n = 'val') => e?.getAttributeNS(W, n) || '';
const norm = s => s.replace(/\s+/g, ' ').trim();
const xml = e => new XMLSerializer().serializeToString(e);
const parse = s => {
    const d = new DOMParser().parseFromString(s, 'application/xml');
    if (d.getElementsByTagName('parsererror').length) throw Error('Struktur Word tidak dapat dibaca. Simpan ulang sebagai .docx melalui Word.');
    return d;
};

function el(n, t, c) {
    const e = document.createElement(n);
    if (t !== undefined) e.textContent = t;
    if (c) e.className = c;
    return e;
}

function we(doc, name, attrs = {}) {
    const n = doc.createElementNS(W, 'w:' + name);
    for (const [k, v] of Object.entries(attrs)) n.setAttributeNS(W, 'w:' + k, String(v));
    return n;
}

function property(parent, name, attrs) {
    direct(parent, name).forEach(n => n.remove());
    const n = we(parent.ownerDocument, name, attrs);
    parent.appendChild(n);
    return n;
}

function textOf(e) {
    let s = '';
    for (const n of Array.from(e.childNodes)) {
        if (n.nodeType !== 1) continue;
        if (n.namespaceURI === W) {
            if (['pPr', 'rPr', 'sectPr'].includes(n.localName)) continue;
            if (n.localName === 't') s += n.textContent;
            else if (n.localName === 'tab') s += '\t';
            else if (['br', 'cr'].includes(n.localName)) s += '\n';
            else s += textOf(n);
        } else s += textOf(n);
    }
    return s;
}

function visiblePara(p) {
    const c = p.cloneNode(true);
    for (const n of nodes(c, 'del')) n.remove();
    for (const r of nodes(c, 'r')) {
        const pr = direct(r, 'rPr')[0];
        if (pr && ['strike', 'dstrike', 'vanish'].some(k => direct(pr, k).some(x => !['0', 'false', 'off'].includes(val(x))))) r.remove();
    }
    return {
        node: c,
        text: textOf(c),
        source: xml(p)
    };
}

function blockText(n) {
    return n.localName === 'tbl' ? direct(n, 'tr').map(r => direct(r, 'tc').map(c => direct(c, 'p').map(textOf).join('\n')).join('\t')).join('\n') : textOf(n);
}

function paragraphs(cell) {
    return Array.from(cell.children).filter(n => ['p', 'tbl'].includes(n.localName)).map(n => {
        const p = visiblePara(n);
        p.text = blockText(p.node);
        return p;
    }).filter(p => p.text.trim() || nodes(p.node, 'drawing').length || nodes(p.node, 'pict').length);
}

function plain(ps) {
    return ps.map(p => p.text).join('\n');
}

function hasContent(cell) {
    return nodes(cell, 't').some(n => n.textContent.trim()) || nodes(cell, 'drawing').length || nodes(cell, 'pict').length;
}

function articleSections(ps) {
    const out = [];
    let current = null;
    for (let i = 0; i < ps.length; i++) {
        const t = norm(ps[i].text);
        const m = t.match(/^Pasal\s+(\d+[A-Za-z]?)$/i);
        if (m) {
            if (current) out.push(current);
            current = {
                num: m[1],
                start: i,
                end: ps.length
            };
        } else if (/^BAB\s+[IVXLCDM]+$/i.test(t) && current) {
            current.end = i;
            out.push(current);
            current = null;
        }
        if (current) current.end = i + 1;
    }
    if (current) out.push(current);
    return out;
}

function sectionBody(ps, s) {
    return norm(plain(ps.slice(s.start + 1, s.end)));
}

function addIssue(m, obj) {
    m.issues.push({
        id: m.issues.length,
        resolved: false,
        ...obj
    });
}

export {
    W,
    R,
    $,
    nodes,
    direct,
    val,
    norm,
    xml,
    parse,
    el,
    we,
    property,
    textOf,
    visiblePara,
    blockText,
    paragraphs,
    plain,
    hasContent,
    articleSections,
    sectionBody,
    addIssue
};
