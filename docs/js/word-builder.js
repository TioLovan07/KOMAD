import {configureFirstPage, positionGaruda} from './page-layout.js';
/**
 * Pembentukan Word memakai profil contoh final. Pemeriksaan teks menghentikan ekspor jika isi tidak cocok.
 */
import {
    W,
    $,
    nodes,
    direct,
    val,
    norm,
    xml,
    parse,
    we,
    property,
    textOf,
    blockText
} from './xml.js';
import {
    model
} from './state.js';
import {
    pending,
    effectiveRows,
    makePara
} from './draft-model.js';
import {
    matrixStripLeading,
    matrixInlineDrawings
} from './matrix-reader.js';
import {
    DRAFT_FORMAT
} from './format-profile.js';
import {
    startsJointApproval,
    isDecision,
    documentMetadata,
    annexHeadingLines,
    isSourceAnnexHeading
} from './document-layout.js';

function fromProfile(doc, obj) {
    const e = we(doc, obj.tag, obj.attrs);
    for (const c of obj.children || []) e.appendChild(fromProfile(doc, c));
    return e;
}

function applyRunFormat(p) {
    for (const t of nodes(p, 't')) {
        t.removeAttribute('xml:space');
        t.removeAttributeNS('http://www.w3.org/XML/1998/namespace', 'space');
        t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    }
    for (const r of nodes(p, 'r')) {
        let rp = direct(r, 'rPr')[0];
        if (!rp) {
            rp = we(p.ownerDocument, 'rPr');
            r.insertBefore(rp, r.firstChild);
        }
        for (const n of ['rStyle', 'highlight', 'shd', 'color', 'strike', 'dstrike', 'vanish', 'sz', 'szCs', 'rFonts', 'b', 'bCs', 'u', 'spacing', 'w', 'fitText']) direct(rp, n).forEach(x => x.remove());
        property(rp, 'rFonts', {
            ascii: DRAFT_FORMAT.font,
            hAnsi: DRAFT_FORMAT.font,
            eastAsia: DRAFT_FORMAT.font,
            cs: DRAFT_FORMAT.font
        });
        property(rp, 'sz', {
            val: DRAFT_FORMAT.size
        });
        property(rp, 'szCs', {
            val: DRAFT_FORMAT.size
        });
        property(rp, 'color', {
            val: '000000'
        });
        property(rp, 'b', {
            val: 0
        });
        property(rp, 'bCs', {
            val: 0
        });
    }
}

function prefixText(p, text) {
    if (!text) return;
    const d = p.ownerDocument,
        r = we(d, 'r');
    for (const part of text.split(/(\t|\n)/)) {
        if (part === '\t') r.appendChild(we(d, 'tab'));
        else if (part === '\n') r.appendChild(we(d, 'br'));
        else if (part) {
            const t = we(d, 't');
            t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
            t.textContent = part;
            r.appendChild(t);
        }
    }
    const pr = direct(p, 'pPr')[0];
    p.insertBefore(r, pr ? pr.nextSibling : p.firstChild);
}

function forceNoBold(doc) {
    for (const b of [...nodes(doc, 'b'), ...nodes(doc, 'bCs')]) b.setAttributeNS(W, 'w:val', '0');
    for (const level of nodes(doc, 'lvl'))
        if (!direct(level, 'rPr').length) level.appendChild(we(doc, 'rPr'));
    for (const pr of nodes(doc, 'rPr')) {
        property(pr, 'b', {
            val: 0
        });
        property(pr, 'bCs', {
            val: 0
        });
    }
}

function orderPPr(pr) {
    const order = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
    const rank = n => {
        const i = order.indexOf(n.localName);
        return i < 0 ? 100 : i;
    };
    for (const n of Array.from(pr.children).sort((a, b) => rank(a) - rank(b))) pr.appendChild(n);
}
async function buildWord({
    allowPending = false
} = {}) {
    if (!allowPending && pending().length) throw Error('Tentukan pilihan untuk semua catatan terlebih dahulu.');
    const selectedRows = effectiveRows();
    const isAnnex = row => row.label.toLowerCase() === 'lampiran';
    // Lampiran selalu berada sesudah pengundangan, walau baris matriks berbeda urutan.
    const annexRows = selectedRows.filter(isAnnex);
    const rows = [...selectedRows.filter(row => !isAnnex(row)), ...annexRows];
    if (!annexRows.length) rows.push({
        index: 'annex',
        label: 'Lampiran',
        ps: []
    });
    const metadata = documentMetadata(selectedRows);
    const
        d = parse(xml(model.doc)),
        body = nodes(d, 'body')[0];
    body.replaceChildren();
    // New archive each time: choices can be changed without altering the source or a prior export.
    const out = await JSZip.loadAsync(await model.zip.generateAsync({
        type: 'uint8array'
    }));
    const numberingFile = out.file('word/numbering.xml');
    const numbering = numberingFile ? parse(await numberingFile.async('string')) : parse('<w:numbering xmlns:w="' + W + '"/>');
    let maxNum = Math.max(0, ...nodes(numbering, 'num').map(n => Number(val(n, 'numId')))),
        maxAbs = Math.max(0, ...nodes(numbering, 'abstractNum').map(n => Number(val(n, 'abstractNumId'))));

    function newNumbering(fmt, start) {
        const ai = ++maxAbs,
            ni = ++maxNum,
            a = we(numbering, 'abstractNum', {
                abstractNumId: ai
            }),
            lvl = we(numbering, 'lvl', {
                ilvl: 0
            });
        lvl.append(we(numbering, 'start', {
            val: start
        }), we(numbering, 'numFmt', {
            val: fmt
        }), we(numbering, 'lvlText', {
            val: '%1.'
        }), we(numbering, 'lvlJc', {
            val: 'left'
        }));
        const rp = we(numbering, 'rPr');
        rp.append(we(numbering, 'rFonts', {
            ascii: DRAFT_FORMAT.font,
            hAnsi: DRAFT_FORMAT.font
        }), we(numbering, 'color', {
            val: '000000'
        }), we(numbering, 'sz', {
            val: 24
        }));
        lvl.appendChild(rp);
        a.appendChild(lvl);
        const num = we(numbering, 'num', {
            numId: ni
        });
        num.appendChild(we(numbering, 'abstractNumId', {
            val: ai
        }));
        const root = numbering.documentElement;
        root.insertBefore(a, direct(root, 'num')[0] || null);
        root.appendChild(num);
        return ni;
    }
    const expected = [];
    let anchorContext = '',
        bookmarkId = 100000,
        anchorCount = 0;
    let inTop = true,
        previous = '',
        chapterTitleNext = false,
        definitions = false,
        signatureStage = '',
        annexStarted = false,
        inAyat = false,
        sectionTitleNext = false,
        jointApproval = false,
        annexHeading = true;
    let annexCount = 0;

    // Potret F4 mengikuti halaman utama; lebar tabel dihitung dari margin aktual.
    const portraitSection = () => fromProfile(d, DRAFT_FORMAT.sections[0]);
    const annexPage = portraitSection();
    const pageWidth = Number(val(direct(annexPage, 'pgSz')[0], 'w'));
    const margins = direct(annexPage, 'pgMar')[0];
    const annexWidth = pageWidth - Number(val(margins, 'left')) - Number(val(margins, 'right'));

    function startAnnex(sourceHeading = false) {
        annexStarted = true;
        inTop = false;
        signatureStage = '';
        jointApproval = false;
        annexHeading = true;
        const section = portraitSection();
        property(section, 'type', {
            val: 'nextPage'
        });
        const boundary = we(d, 'p'),
            pr = we(d, 'pPr');
        pr.appendChild(section);
        boundary.appendChild(pr);
        body.appendChild(boundary);
        const oldAnchor = anchorContext;
        anchorContext = '';
        for (const line of sourceHeading ? [] : annexHeadingLines(metadata)) {
            const p = addParagraph(makePara(line).node, 'annexHeader', {
                before: 0,
                removeNum: true
            });
            property(direct(p, 'pPr')[0], 'keepLines', {
                val: 1
            });
        }
        anchorContext = oldAnchor;
    }

    function layoutTable(widths, indent = 0) {
        const excess = widths.reduce((a,b)=>a+b,0) + indent - annexWidth;
        if (excess > 0) widths[widths.length-1] = Math.max(240,widths.at(-1)-excess);
        const table = we(d, 'tbl'),
            pr = we(d, 'tblPr'),
            width = widths.reduce((a, b) => a + b, 0);
        pr.append(we(d, 'tblW', {
            w: width,
            type: 'dxa'
        }), we(d, 'tblInd', {
            w: indent,
            type: 'dxa'
        }), we(d, 'tblLayout', {
            type: 'fixed'
        }));
        const borders = we(d, 'tblBorders');
        for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) borders.appendChild(we(d, side, {
            val: 'nil'
        }));
        pr.appendChild(borders);
        const margins = we(d, 'tblCellMar');
        for (const side of ['top', 'left', 'bottom', 'right']) margins.appendChild(we(d, side, {
            w: 0,
            type: 'dxa'
        }));
        pr.appendChild(margins);
        table.appendChild(pr);
        const grid = we(d, 'tblGrid');
        widths.forEach(w => grid.appendChild(we(d, 'gridCol', {
            w
        })));
        table.appendChild(grid);
        return table;
    }

    function addParagraph(node, role, extra = {}) {
        const p = d.importNode(node, true);
        const leading = blockText(p).match(/^[ \t]+/);
        if (leading) matrixStripLeading(p, leading[0].length);
        for (const t of nodes(p, 't')) t.textContent = t.textContent.replace(/ {2,}/g, ' ');
        for (const n of [...nodes(p, 'sectPr'), ...nodes(p, 'commentRangeStart'), ...nodes(p, 'commentRangeEnd'), ...nodes(p, 'commentReference'), ...nodes(p, 'bookmarkStart'), ...nodes(p, 'bookmarkEnd')]) n.remove();
        const oldPr = direct(p, 'pPr')[0],
            oldNum = oldPr && direct(oldPr, 'numPr')[0]?.cloneNode(true);
        oldPr?.remove();
        const pr = fromProfile(d, DRAFT_FORMAT.roles[role] || DRAFT_FORMAT.roles.body);
        p.insertBefore(pr, p.firstChild);
        if (!direct(pr, 'jc').length) property(pr, 'jc', {
            val: 'left'
        });
        if (!direct(pr, 'ind').length) property(pr, 'ind', {
            left: 0,
            right: 0,
            firstLine: 0
        });
        const before = extra.before ?? DRAFT_FORMAT.before[role] ?? 0;
        property(pr, 'spacing', {
            before,
            after: 0,
            line: 240,
            lineRule: 'auto'
        });
        if (extra.numbering) {
            const np = we(d, 'numPr');
            np.append(we(d, 'ilvl', {
                val: 0
            }), we(d, 'numId', {
                val: extra.numbering
            }));
            pr.insertBefore(np, pr.firstChild);
        } else if (oldNum && !extra.removeNum) pr.insertBefore(d.importNode(oldNum, true), pr.firstChild);
        if (['chapter', 'chapterTitle', 'article', 'definitionLead', 'decide', 'set', 'signatureRole', 'annexHeader'].includes(role)) property(pr, 'keepNext', {
            val: 1
        });
        if (extra.flat) {
            property(pr, 'ind', {
                left: 0,
                right: 0,
                firstLine: 0
            });
            direct(pr, 'tabs').forEach(n => n.remove());
            property(pr, 'jc', {
                val: extra.align || 'left'
            });
        }
        if (extra.listLevel) {
            const left = 3119 + 567 * extra.listLevel;
            property(pr, 'ind', {
                left,
                hanging: 567
            });
            property(pr, 'tabs', {}).appendChild(we(d, 'tab', {
                val: 'left',
                pos: left
            }));
        }
        matrixInlineDrawings(p);
        property(pr, 'suppressAutoHyphens', {
            val: 1
        });
        property(pr, 'widowControl', {
            val: 1
        });
        orderPPr(pr);
        prefixText(p, extra.prefix || '');
        applyRunFormat(p);
        if (anchorContext) {
            const name = anchorContext + (anchorCount ? '__' + anchorCount : '');
            anchorCount++;
            const id = bookmarkId++;
            const bm = we(d, 'bookmarkStart', {
                id,
                name
            });
            p.insertBefore(bm, pr.nextSibling);
            p.appendChild(we(d, 'bookmarkEnd', {
                id
            }));
        }(extra.target || body).appendChild(p);
        if (!extra.target) expected.push((extra.prefix || '') + blockText(node));
        return p;
    }

    function copyTable(node) {
        const table = d.importNode(node, true);
        let pr = direct(table, 'tblPr')[0];
        if (!pr) {
            pr = we(d, 'tblPr');
            table.insertBefore(pr, table.firstChild);
        }
        property(pr, 'tblInd', {
            w: 0,
            type: 'dxa'
        });
        property(pr, 'tblLayout', {
            type: 'fixed'
        });
        const grid = direct(table, 'tblGrid')[0],
            columns = grid ? direct(grid, 'gridCol') : [],
            widths = columns.map(c => Number(val(c, 'w')) || 1),
            total = widths.reduce((a, b) => a + b, 0),
            available = annexStarted ? annexWidth : 7371,
            scale = total ? (annexStarted ? available / total : Math.min(1, available / total)) : 1;
        if (total) {
            columns.forEach((c, i) => c.setAttributeNS(W, 'w:w', Math.round(widths[i] * scale)));
            property(pr, 'tblW', {
                w: Math.round(total * scale),
                type: 'dxa'
            });
            for (const tr of direct(table, 'tr')) {
                let col = 0;
                for (const tc of direct(tr, 'tc')) {
                    let cp = direct(tc, 'tcPr')[0];
                    if (!cp) {
                        cp = we(d, 'tcPr');
                        tc.insertBefore(cp, tc.firstChild);
                    }
                    const span = Number(val(direct(cp, 'gridSpan')[0])) || 1;
                    property(cp, 'tcW', {
                        w: Math.round(widths.slice(col, col + span).reduce((a, b) => a + b, 0) * scale),
                        type: 'dxa'
                    });
                    col += span;
                }
            }
        }
        for (const p of nodes(table, 'p')) {
            const old = direct(p, 'pPr')[0];
            old?.remove();
            const pp = we(d, 'pPr');
            pp.append(we(d, 'spacing', {
                before: 0,
                after: 0,
                line: 240,
                lineRule: 'auto'
            }), we(d, 'ind', {
                left: 0,
                right: 0,
                firstLine: 0
            }), we(d, 'jc', {
                val: 'left'
            }));
            p.insertBefore(pp, p.firstChild);
            applyRunFormat(p);
            matrixInlineDrawings(p);
        }
        for (const h of nodes(table, 'trHeight')) h.remove();
        body.appendChild(table);
        expected.push(blockText(node));
    }
    for (const row of rows) {
        let group = null,
            groupIndex = 0,
            groupClosed = false;
        const flushGroup = () => {
            if (group) expected.push(blockText(group));
            group = null;
            groupClosed = true;
        };
        if (isAnnex(row) && !annexStarted) {
            row.sourceAnnexHeading = row.ps.some(p=>/^LAMPIRAN\b/i.test(norm(p.text)));
            const topicIndex=row.ps.findIndex(p=>/^TENTANG\b/i.test(norm(p.text)));
            row.sourceAnnexHeaderEnd=topicIndex>=0 ? topicIndex + (norm(row.ps[topicIndex].text)==='TENTANG' ? 1 : 0) : Math.min(4,row.ps.length-1);
            startAnnex(row.sourceAnnexHeading);
        }
        const groupWidths = row.prefix === 'Menetapkan' ? [2268, 284, 7371] : [2268, 284, 567, 6804];
        if (row.prefix) inTop = false;
        for (let blockIndex = 0; blockIndex < row.ps.length; blockIndex++) {
            const item = row.ps[blockIndex];
            anchorContext = 'p3h_' + row.index + '_' + (item.sourceBlockIndex ?? blockIndex);
            anchorCount = 0;
            const t = norm(item.text),
                isDrawing = nodes(item.node, 'drawing').length > 0;
            if (annexStarted && /^LAMPIRAN(?:\s+[IVXLCDM]+|\s+\d+)?$/i.test(t)) {
                row.sourceAnnexHeading = true;
                const remaining = row.ps.slice(blockIndex);
                const topic = remaining.findIndex(p => /^TENTANG\b/i.test(norm(p.text)));
                row.sourceAnnexHeaderEnd = topic >= 0 ? blockIndex + topic + (norm(remaining[topic].text).toUpperCase() === 'TENTANG' ? 1 : 0) : blockIndex + 4;
                annexHeading = true;
                signatureStage = '';
                row.annexStartIndex = blockIndex;
            }
            if(annexStarted && row.sourceAnnexHeading && blockIndex>row.sourceAnnexHeaderEnd) annexHeading=false;
            let role = 'body',
                extra = {
                    removeNum: true
                };
            if (annexStarted && !row.sourceAnnexHeading && annexHeading && item.node.localName !== 'tbl' &&
                isSourceAnnexHeading(t, metadata)) continue;
            if (startsJointApproval(t)) {
                flushGroup();
                jointApproval = true;
                inTop = false;
            }
            if (isDecision(t)) {
                if (group) flushGroup();
                jointApproval = false;
                if (row.prefix === 'Menetapkan') groupClosed = false;
            }
            if (item.node.localName === 'tbl') {
                const trs = direct(item.node, 'tr');
                const isBoundaryTable = !annexStarted && trs.length && trs.every(tr => {
                    const cs = direct(tr, 'tc');
                    return cs.length === 4 && /^[a-z]\.$/i.test(norm(textOf(cs[0]))) && norm(textOf(cs[2])) === ':';
                });
                if (isBoundaryTable) {
                    const widths = [567, 3118, 426, 3260],
                        table = layoutTable(widths, 2552);
                    body.appendChild(table);
                    for (const tr of trs) {
                        const outRow = we(d, 'tr');
                        table.appendChild(outRow);
                        direct(tr, 'tc').forEach((cell, index) => {
                            const tc = we(d, 'tc'),
                                pr = we(d, 'tcPr');
                            pr.append(we(d, 'tcW', {
                                w: widths[index],
                                type: 'dxa'
                            }), we(d, 'vAlign', {
                                val: 'top'
                            }));
                            tc.appendChild(pr);
                            outRow.appendChild(tc);
                            for (const para of direct(cell, 'p')) addParagraph(para, 'body', {
                                target: tc,
                                flat: true,
                                before: 0,
                                removeNum: true
                            });
                        });
                    }
                    expected.push(blockText(table));
                } else copyTable(item.node);
                annexHeading = false;
                previous = t;
                continue;
            }
            if (row.prefix && !groupClosed && !jointApproval && !isDecision(t)) {
                if (!group) {
                    group = layoutTable(groupWidths);
                    body.appendChild(group);
                }
                const first = groupIndex === 0,
                    tr = we(d, 'tr');
                group.appendChild(tr);
                const isSet = row.prefix === 'Menetapkan',
                    marker = isSet ? '' : item.list?.marker || (row.prefix === 'Menimbang' ? String.fromCharCode(97 + groupIndex) + '.' : String(groupIndex + 1) + '.');
                const source = item.node.cloneNode(true);
                if (item.list?.prefixLength) matrixStripLeading(source, item.list.prefixLength);
                const contents = isSet ? [first ? row.prefix : '', first ? ':' : '', source] : [first ? row.prefix : '', first ? ':' : '', marker, source];
                contents.forEach((content, i) => {
                    const tc = we(d, 'tc'),
                        pr = we(d, 'tcPr');
                    pr.append(we(d, 'tcW', {
                        w: groupWidths[i],
                        type: 'dxa'
                    }), we(d, 'vAlign', {
                        val: 'top'
                    }));
                    tc.appendChild(pr);
                    tr.appendChild(tc);
                    const node = typeof content === 'string' ? makePara(content).node : content;
                    addParagraph(node, 'body', {
                        target: tc,
                        flat: true,
                        align: i === contents.length - 1 ? 'both' : 'left',
                        before: first ? 240 : 0,
                        removeNum: true
                    });
                });
                groupIndex++;
                previous = t;
                continue;
            }
            if (annexStarted && row.sourceAnnexHeading && blockIndex<=row.sourceAnnexHeaderEnd) {
                role = 'annexHeader';
                extra.before = 0;
            } else if (jointApproval) {
                role = 'approval';
                extra.before = startsJointApproval(t) ? 240 : 0;
            } else if (/^MEMUTUSKAN\s*:$/i.test(t)) {
                inTop = false;
                role = 'decide';
            } else if (/^BAB\s+[IVXLCDM]+[A-Z]?$/i.test(t)) {
                inTop = false;
                chapterTitleNext = 1;
                definitions = false;
                inAyat = false;
                role = 'chapter';
            } else if (chapterTitleNext && !item.list && !/^(?:Pasal|Bagian|Lampiran)\b/i.test(t) && (chapterTitleNext === 1 || (/[A-ZÀ-Ž]/.test(t) && t === t.toUpperCase()))) {
                role = 'chapterTitle';
                chapterTitleNext = 2;
                extra.before = 0;
            } else if (/^Bagian\s+\S+$/i.test(t)) {
                role = 'article';
                sectionTitleNext = true;
            } else if (sectionTitleNext) {
                role = 'chapterTitle';
                sectionTitleNext = false;
            } else if (/^Pasal\s+(?:\d+[A-Za-z]?|[IVXLCDM]+)$/i.test(t)) {
                inTop = false;
                definitions = false;
                inAyat = false;
                role = 'article';
            } else if (/yang dimaksud dengan\s*:\s*$/i.test(t)) {
                definitions = true;
                role = 'definitionLead';
            } else if (annexStarted && !annexHeading && /^(?:B\s+)?BUPATI\b/.test(t)) {
                signatureStage = 'annex';
                role = 'annex';
                extra.before = 240;
            } else if (signatureStage === 'annex') {
                role = 'annex';
                extra.before = 960;
            } else if (/^Ditetapkan di\b/i.test(t)) {
                signatureStage = 'set';
                role = 'signature';
            } else if (/^Diundangkan di\b/i.test(t)) {
                signatureStage = 'promulgate';
                role = 'promulgation';
            } else if (/^BERITA (DAERAH|NEGARA)\b/.test(t)) {
                signatureStage = '';
                role = 'gazette';
            } else if (signatureStage === 'set') {
                role = /^pada tanggal/i.test(t) ? 'signature' : /^BUPATI\b/.test(t) ? 'signatureRole' : 'signatureName';
            } else if (signatureStage === 'promulgate') {
                role = /^[.\s…]+$/.test(t) ? 'promulgationName' : 'promulgation';
                extra.before = /^pada tanggal/i.test(t) ? 0 : 240;
            } else if (annexStarted && /^Struktur dan Besaran/i.test(t)) {
                annexHeading = false;
                role = 'top';
                extra.before = 240;
            } else if (annexStarted && annexHeading) {
                annexHeading = false;
                role = 'top';
                extra.before = 240;
            } else if (inTop) {
                role = 'top';
                if (/^RANCANGAN$/.test(t)) extra.before = 240;
                else if (/^PERATURAN\b/.test(t)) extra.before = previous === 'RANCANGAN' ? 0 : 240;
                else if (/^(TENTANG|DENGAN RAHMAT)\b/.test(t) || /^BUPATI\s.*,\s*$/.test(t) || previous === 'TENTANG') extra.before = 240;
            } else if (item.list) {
                role = definitions ? 'definition' : 'list';
                extra.before = 0;
            }
            if (!['chapter','chapterTitle'].includes(role)) chapterTitleNext = false;
            if (role === 'signatureName') {
                role = 'signature';
                extra.before = 960;
            } else if (role === 'signature' && /^pada tanggal/i.test(t)) extra.before = 0;
            // Perataan tengah eksplisit pada rumusan umum tidak dibuang bersama pPr matriks.
            const sourcePr = direct(item.node, 'pPr')[0];
            if (role === 'body' && sourcePr && val(direct(sourcePr, 'jc')[0]) === 'center') {
                extra.flat = true;
                extra.align = 'center';
            }
            let source = item.node;
            if (item.list && !inTop && !annexStarted && !signatureStage && !['article', 'chapter', 'chapterTitle'].includes(role)) {
                source = item.node.cloneNode(true);
                if (item.list.prefixLength) matrixStripLeading(source, item.list.prefixLength);
                extra.prefix = item.list.marker + '\t';
                extra.removeNum = true;
                if (item.list.kind === 'ayat') inAyat = true;
                extra.listLevel = item.list.kind === 'letter' && inAyat ? 1 : 0;
            }
            if (item.amendmentInstruction) { extra.before = 240; inAyat = false; }
            if (role === 'body' && /^Pasal\s+(?:\d+[A-Za-z]?|[IVXLCDM]+)$/i.test(previous)) extra.before = 0;
            const p = addParagraph(source, role, extra);
            if (annexStarted && /^LAMPIRAN(?:\s+[IVXLCDM]+|\s+\d+)?$/i.test(t)) {
                if (annexCount++) {
                    property(direct(p,'pPr')[0], 'pageBreakBefore', {val:1});
                    orderPPr(direct(p,'pPr')[0]);
                }
            }
            if (isDrawing && inTop) {
                for (const extent of p.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'extent')) {
                    const scale = DRAFT_FORMAT.logo.cy / Number(extent.getAttribute('cy'));
                    extent.setAttribute('cx', String(Math.round(Number(extent.getAttribute('cx')) * scale)));
                    extent.setAttribute('cy', String(DRAFT_FORMAT.logo.cy));
                }
                for (const extent of p.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'ext'))
                    if (extent.hasAttribute('cx') && extent.hasAttribute('cy')) {
                        const scale = DRAFT_FORMAT.logo.cy / Number(extent.getAttribute('cy'));
                        extent.setAttribute('cx', String(Math.round(Number(extent.getAttribute('cx')) * scale)));
                        extent.setAttribute('cy', String(DRAFT_FORMAT.logo.cy));
                    }
            }
            if (isDrawing && inTop) {
                positionGaruda(p);
                if (!blockText(p).trim()) property(direct(p,'pPr')[0],'spacing',{before:0,after:0,line:1,lineRule:'exact'});
            }
            previous = t;
        }
        if (group) expected.push(blockText(group));
    }
    body.appendChild(portraitSection());
    const actual = Array.from(body.children).filter(n => ['p', 'tbl'].includes(n.localName) && !(n.localName === 'p' && !blockText(n) && !nodes(n, 'drawing').length)).map(blockText).join('\n');
    const contentKey = text => text.replace(/[ \t]+/g, ' ').replace(/(^|\n) +/g, '$1');
    if (contentKey(expected.join('\n')) !== contentKey(actual)) throw Error('Pemeriksaan isi tidak cocok. Unduhan dihentikan.');
    const settingsFile = out.file('word/settings.xml');
    if (settingsFile) {
        const settings = parse(await settingsFile.async('string'));
        property(settings.documentElement, 'autoHyphenation', {
            val: 0
        });
        out.file('word/settings.xml', xml(settings));
    }
    await configureFirstPage(d, out);
    forceNoBold(d);
    const documentXml = xml(d);
    parse(documentXml);
    out.file('word/document.xml', documentXml);
    out.remove('word/comments.xml');
    for (const c of nodes(numbering, 'color')) {
        c.setAttributeNS(W, 'w:val', '000000');
        for (const a of ['themeColor', 'themeTint', 'themeShade']) c.removeAttributeNS(W, a);
    }
    forceNoBold(numbering);
    out.file('word/numbering.xml', xml(numbering));
    const stylesFile = out.file('word/styles.xml');
    if (stylesFile) {
        const st = parse(await stylesFile.async('string'));
        const defs = nodes(st, 'docDefaults')[0];
        if (defs) {
            let rd = direct(defs, 'rPrDefault')[0];
            if (!rd) {
                rd = we(st, 'rPrDefault');
                defs.insertBefore(rd, defs.firstChild);
            }
            let rp = direct(rd, 'rPr')[0];
            if (!rp) {
                rp = we(st, 'rPr');
                rd.appendChild(rp);
            }
            property(rp, 'rFonts', {
                ascii: DRAFT_FORMAT.font,
                hAnsi: DRAFT_FORMAT.font,
                eastAsia: DRAFT_FORMAT.font,
                cs: DRAFT_FORMAT.font
            });
            property(rp, 'sz', {
                val: 24
            });
            property(rp, 'szCs', {
                val: 24
            });
            let pd = direct(defs, 'pPrDefault')[0];
            if (!pd) {
                pd = we(st, 'pPrDefault');
                defs.appendChild(pd);
            }
            pd.replaceChildren();
            const pp = we(st, 'pPr');
            pp.appendChild(we(st, 'spacing', {
                before: 0,
                after: 0,
                line: 240,
                lineRule: 'auto'
            }));
            pd.appendChild(pp);
        }
        for (const c of nodes(st, 'color')) {
            c.setAttributeNS(W, 'w:val', '000000');
            for (const a of ['themeColor', 'themeTint', 'themeShade']) c.removeAttributeNS(W, a);
        }
        forceNoBold(st);
        out.file('word/styles.xml', xml(st));
    }
    return await out.generateAsync({
        type: 'blob',
        compression: 'DEFLATE',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
}

export {
    buildWord
};
