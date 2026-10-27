/**
 * Pengenalan kolom, daftar Word, Tetap per butir, dan normalisasi gambar inline.
 */
import {
    W,
    $,
    nodes,
    direct,
    val,
    norm,
    textOf,
    visiblePara,
    blockText,
    addIssue
} from './xml.js';

function matrixCells(row) {
    const cells = [];
    let col = Number(val(direct(row, 'trPr')[0]?.getElementsByTagNameNS(W, 'gridBefore')[0])) || 0;
    for (const c of direct(row, 'tc')) {
        const span = Number(val(nodes(direct(c, 'tcPr')[0] || c, 'gridSpan')[0])) || 1;
        for (let n = 0; n < span; n++) cells[col++] = c;
    }
    return cells;
}

function matrixHeader(table) {
    const rows = direct(table, 'tr');
    for (let i = 0; i < Math.min(rows.length, 8); i++) {
        const cells = matrixCells(rows[i]),
            texts = cells.map(c => c ? norm(textOf(c)).toLowerCase().replace(/[.:]/g, '') : '');
        const hasil = texts.findIndex(t => /hasil.*(penyempurnaan|penyesuaian|harmonisasi)|rumusan (akhir|hasil)|hasil perbaikan/.test(t));
        if (hasil < 0) continue;
        const choices = texts.map((t, j) => ({
            t,
            j
        })).filter(x => x.j !== hasil && /\bawal\b|^rancangan\s+(perbup|ranperbup|perda|ranperda|perkada|peraturan)\b|^naskah rancangan\b/.test(x.t));
        if (choices.length !== 1) continue;
        const awal = choices[0].j,
            ket = texts.findIndex(t => /^keterangan\b|^catatan\b/.test(t));
        return {
            header: i,
            cols: {
                awal,
                hasil,
                ket
            }
        };
    }
    return null;
}

function matrixRowLabel(cells, awal) {
    const candidates = cells.slice(0, awal).filter(Boolean).map(c => norm(textOf(c))).filter(t => t && !/^\d+$/.test(t) && t !== ':');
    const prefix = candidates.map(t => t.match(/\b(Menimbang|Mengingat|Menetapkan)\b/i)).find(Boolean);
    return {
        label: prefix ? prefix[1][0].toUpperCase() + prefix[1].slice(1).toLowerCase() : candidates.join(' '),
        prefix: prefix ? prefix[1][0].toUpperCase() + prefix[1].slice(1).toLowerCase() : null
    };
}

function matrixNumberingCatalog(numbering, styles) {
    const nums = new Map(),
        abstracts = new Map(),
        styleMap = new Map();
    if (numbering) {
        for (const a of nodes(numbering, 'abstractNum')) abstracts.set(val(a, 'abstractNumId'), a);
        for (const n of direct(numbering.documentElement, 'num')) nums.set(val(n, 'numId'), n);
    }
    if (styles)
        for (const s of nodes(styles, 'style')) styleMap.set(val(s, 'styleId'), s);
    return {
        nums,
        abstracts,
        styleMap
    };
}

function matrixNumberProperties(p, catalog) {
    let pr = direct(p, 'pPr')[0],
        np = pr && direct(pr, 'numPr')[0];
    if (np) return np;
    let style = pr && val(direct(pr, 'pStyle')[0]),
        seen = new Set();
    while (style && !seen.has(style)) {
        seen.add(style);
        const s = catalog.styleMap.get(style);
        if (!s) break;
        const sp = direct(s, 'pPr')[0];
        np = sp && direct(sp, 'numPr')[0];
        if (np) return np;
        style = val(direct(s, 'basedOn')[0]);
    }
    return null;
}

function matrixOrdinal(n, fmt) {
    if (fmt === 'lowerLetter' || fmt === 'upperLetter') {
        let s = '';
        do {
            n--;
            s = String.fromCharCode(97 + n % 26) + s;
            n = Math.floor(n / 26);
        } while (n > 0);
        return fmt === 'upperLetter' ? s.toUpperCase() : s;
    }
    if (fmt === 'lowerRoman' || fmt === 'upperRoman') {
        let s = '';
        for (const [v, c] of [
                [1000, 'M'],
                [900, 'CM'],
                [500, 'D'],
                [400, 'CD'],
                [100, 'C'],
                [90, 'XC'],
                [50, 'L'],
                [40, 'XL'],
                [10, 'X'],
                [9, 'IX'],
                [5, 'V'],
                [4, 'IV'],
                [1, 'I']
            ])
            while (n >= v) {
                s += c;
                n -= v;
            }
        return fmt === 'lowerRoman' ? s.toLowerCase() : s;
    }
    return String(n);
}

function matrixNativeMarker(p, catalog, state) {
    const np = matrixNumberProperties(p, catalog);
    if (!np) return null;
    const numId = val(direct(np, 'numId')[0]);
    if (!numId || numId === '0') return null;
    const level = Number(val(direct(np, 'ilvl')[0])) || 0,
        num = catalog.nums.get(numId),
        abs = catalog.abstracts.get(num && val(direct(num, 'abstractNumId')[0]));
    if (!abs) return null;
    const override = direct(num, 'lvlOverride').find(x => Number(val(x, 'ilvl')) === level),
        lvl = override && direct(override, 'lvl')[0] || direct(abs, 'lvl').find(x => Number(val(x, 'ilvl')) === level);
    if (!lvl) return null;
    const fmt = val(direct(lvl, 'numFmt')[0]) || 'decimal';
    if (['bullet', 'none'].includes(fmt)) return {
        marker: val(direct(lvl, 'lvlText')[0]) || '•',
        level,
        kind: 'bullet'
    };
    const start = Number(val(override && direct(override, 'startOverride')[0]) || val(direct(lvl, 'start')[0])) || 1;
    const key = numId + ':' + level,
        n = state.has(key) ? state.get(key) + 1 : start;
    state.set(key, n);
    for (const k of [...state.keys()])
        if (k.startsWith(numId + ':') && Number(k.split(':')[1]) > level) state.delete(k);
    const template = val(direct(lvl, 'lvlText')[0]) || '%' + (level + 1) + '.';
    const marker = template.replace(/%(\d+)/g, (_, raw) => {
        const l = Number(raw) - 1,
            def = direct(abs, 'lvl').find(x => Number(val(x, 'ilvl')) === l);
        return matrixOrdinal(state.get(numId + ':' + l) || Number(val(def && direct(def, 'start')[0])) || 1, val(def && direct(def, 'numFmt')[0]) || fmt);
    });
    return {
        marker,
        level,
        kind: /^\(/.test(marker) ? 'ayat' : /Letter/.test(fmt) ? 'letter' : 'number'
    };
}

function matrixManualMarker(text) {
    const m = text.match(/^\s*(\(\d+\)|[a-zA-Z][.)]|\d+[.)])\s*/);
    if (!m) return null;
    return {
        marker: m[1],
        prefixLength: m[0].length,
        level: 0,
        kind: m[1][0] === '(' ? 'ayat' : /^[a-zA-Z]/.test(m[1]) ? 'letter' : 'number'
    };
}

function matrixParagraphs(cell, catalog, state) {
    const result = [];
    let sourceArticle = '';
    for (const node of Array.from(cell.children).filter(n => ['p', 'tbl'].includes(n.localName))) {
        if (/^\s*(Pasal\s+\d+|BAB\s+[IVXLCDM]+[A-Z]?)\s*$/i.test(textOf(node))) state.clear();
        const rawTitle = norm(textOf(node)).match(/^Pasa[l]?\s+(\d+[A-Za-z]?)$/i);
        if (rawTitle) sourceArticle = rawTitle[1];
        const native = node.localName === 'p' ? matrixNativeMarker(node, catalog, state) : null;
        const p = visiblePara(node);
        p.text = blockText(p.node);
        p.list = native || matrixManualMarker(p.text);
        p.sourceArticle = sourceArticle;
        if (p.text.trim() || nodes(p.node, 'drawing').length || nodes(p.node, 'pict').length) result.push(p);
    }
    return result;
}

function matrixCloneBlock(p) {
    return {
        ...p,
        list: p.list ? {
            ...p.list
        } : null,
        node: p.node.cloneNode(true)
    };
}

function matrixKeepText(p) {
    return /^(?:\s*(?:\d+[.)]|[a-z][.)]|\(\d+\))\s*)?tetap\s*[.;]?\s*$/i.test(p.text);
}

function matrixItemKey(p) {
    const marker = p.list?.marker || matrixManualMarker(p.text)?.marker;
    return marker ? marker.toLowerCase().replace(/[().\s]/g, '') : null;
}

function matrixSelectRumusan(left, right, {
    m,
    rowIndex,
    prefix
}) {
    if (!right.length) return {
        ps: left.map(matrixCloneBlock),
        useRight: false,
        rule: 'Kolom hasil kosong; kolom awal digunakan.'
    };
    if (right.length === 1 && matrixKeepText(right[0]) && !right[0].list) return {
        ps: left.map(matrixCloneBlock),
        useRight: false,
        rule: 'Kolom hasil bertuliskan Tetap; rumusan awal digunakan.'
    };
    const titleRE = /^Pasal\s+(?:\d+[A-Za-z]?|[IVXLCDM]+)$/i;
    const oldTitle = left.findIndex(p => /^Pasa[l]?\s+(?:\d+[A-Za-z]?|[IVXLCDM]+)$/i.test(norm(p.text)));
    const newTitles = right.map((p,i) => titleRE.test(norm(p.text)) ? i : -1).filter(i=>i>=0);
    const keeps = right.filter(matrixKeepText);
    const sourceArticles=[...new Set(left.map(p=>p.sourceArticle).filter(Boolean))];
    const preservedBody=oldTitle>=0 ? left.slice(oldTitle+1) :
        sourceArticles.length===1 ? left.filter(p=>p.sourceArticle===sourceArticles[0]) : [];
    if (preservedBody.length && newTitles.length === 1 && keeps.length === 1 &&
        right.slice(newTitles[0]+1).every(matrixKeepText)) {
        return {
            ps: [...right.slice(0,newTitles[0]+1).map(matrixCloneBlock), ...preservedBody.map(matrixCloneBlock)],
            useRight: true,
            rule: 'Tetap untuk isi pasal: judul/nomor hasil digunakan; seluruh isi pasal awal dipertahankan.'
        };
    }
    // A numbered fragment is not a complete replacement of a chapter/article.
    const articleHead = p => /^Pasal\s+\d+[A-Za-z]?$/i.test(norm(p.text));
    const lostArticle = left.some(articleHead) && !right.some(articleHead);
    if (lostArticle && !right.some(matrixKeepText)) {
        const keys = new Map();
        left.forEach((p, i) => {
            const key = matrixItemKey(p);
            if (key) keys.set(key, keys.has(key) ? null : i);
        });
        const updates = new Map();
        const safe = right.every(p => {
            const key = matrixItemKey(p), index = key && keys.get(key);
            if (!key || index == null || updates.has(index)) return false;
            updates.set(index, p);
            return true;
        });
        if (safe && updates.size) {
            addIssue(m, {
                kind: 'review', rowIndex,
                title: 'Revisi parsial digabungkan pada baris ' + rowIndex,
                message: 'Judul BAB/Pasal, pengantar, dan butir lain dipertahankan dari awal. Hanya penanda ' +
                    right.map(p => matrixItemKey(p)).join(', ') +
                    ' diganti dari hasil penyempurnaan. Periksa kecocokan nomor dan substansi sebelum menyetujui.'
            });
            return {
                ps: left.map((p, i) => matrixCloneBlock(updates.get(i) || p)),
                useRight: true,
                rule: 'Revisi parsial: butir unik diganti berdasarkan penanda; bagian lainnya dipertahankan.'
            };
        }
        addIssue(m, {
            kind: 'review', hard: true, rowIndex,
            title: 'Rumusan parsial belum dapat digabungkan pada baris ' + rowIndex,
            message: 'Hasil tidak memuat judul Pasal asal dan butir tidak dapat dicocokkan secara unik. Draft menampilkan sumber awal sementara; perbaiki matriks sebelum unduhan.'
        });
        return {ps: left.map(matrixCloneBlock), useRight: false, rule: 'Sumber awal dipertahankan sementara; penggabungan ambigu memblokir unduhan.'};
    }
    if (!right.some(matrixKeepText)) return {
        ps: right.map(matrixCloneBlock),
        useRight: true,
        rule: 'Kolom hasil digunakan.'
    };
    const leftKeys = new Map();
    for (const p of left) {
        const key = matrixItemKey(p);
        if (key) {
            if (leftKeys.has(key)) leftKeys.set(key, null);
            else leftKeys.set(key, p);
        }
    }
    const result = [];
    let ordinal = 0;
    for (const p of right) {
        if (!matrixKeepText(p)) {
            result.push(matrixCloneBlock(p));
            if (p.list || prefix) ordinal++;
            continue;
        }
        const key = matrixItemKey(p);
        let original = key ? leftKeys.get(key) : null;
        if (!key && left.length === right.length) original = left[ordinal];
        if (!key && prefix && ordinal < left.length) original = left[ordinal];
        if (!original) {
            addIssue(m, {
                kind: 'review',
                hard: true,
                rowIndex,
                title: 'Tetap pada baris ' + rowIndex + ' belum dapat dicocokkan',
                message: 'Penanda butir tidak cocok atau muncul lebih dari sekali pada rumusan awal. Perbaiki penanda pada matriks lalu unggah ulang; isi tidak ditebak.'
            });
        } else {
            result.push(matrixCloneBlock(original));
            m.decisions.push('Baris ' + rowIndex + ' · Butir ' + (key || ordinal + 1) + ' bertuliskan Tetap; rumusan awal dipertahankan.');
        }
        ordinal++;
    }
    return {
        ps: result,
        useRight: true,
        rule: 'Rumusan hasil dan butir Tetap dicocokkan berdasarkan penanda.'
    };
}

function matrixStripLeading(node, count) {
    let remaining = count;

    function walk(parent) {
        for (const child of Array.from(parent.childNodes)) {
            if (!remaining) break;
            if (child.nodeType !== 1) continue;
            if (child.namespaceURI === W && child.localName === 't') {
                const cut = Math.min(remaining, child.textContent.length);
                child.textContent = child.textContent.slice(cut);
                remaining -= cut;
            } else if (child.namespaceURI === W && ['tab', 'br', 'cr'].includes(child.localName)) {
                child.remove();
                remaining--;
            } else if (!['pPr', 'rPr'].includes(child.localName)) walk(child);
        }
    }
    walk(node);
    if (remaining) throw Error('Penanda butir tidak dapat dipisahkan secara utuh.');
}

function matrixInlineDrawings(node) {
    const ns = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
    for (const anchor of Array.from(node.getElementsByTagNameNS(ns, 'anchor'))) {
        const inline = node.ownerDocument.createElementNS(ns, 'wp:inline');
        for (const name of ['extent', 'effectExtent', 'docPr', 'cNvGraphicFramePr', 'graphic']) {
            const child = Array.from(anchor.children).find(c => c.localName === name);
            if (child) inline.appendChild(child.cloneNode(true));
        }
        for (const name of ['distT', 'distB', 'distL', 'distR']) inline.setAttribute(name, '0');
        anchor.replaceWith(inline);
    }
}

export {
    matrixCells,
    matrixHeader,
    matrixRowLabel,
    matrixNumberingCatalog,
    matrixNumberProperties,
    matrixOrdinal,
    matrixNativeMarker,
    matrixManualMarker,
    matrixParagraphs,
    matrixCloneBlock,
    matrixKeepText,
    matrixItemKey,
    matrixSelectRumusan,
    matrixStripLeading,
    matrixInlineDrawings
};
