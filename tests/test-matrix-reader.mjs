import vm from 'node:vm';
import fs from 'node:fs';
import assert from 'node:assert/strict';
const W = 'w';
class Element {
    constructor(name, text = '', children = [], attrs = {}) {
        this.localName = name;
        this.namespaceURI = W;
        this.text = text;
        this.children = children;
        this.attrs = attrs;
    }
    get textContent() {
        return this.text + this.children.map(x => x.textContent).join('');
    }
    getElementsByTagNameNS(ns, name) {
        return this.children.flatMap(c => [...(c.localName === name ? [c] : []), ...c.getElementsByTagNameNS(ns, name)]);
    }
    cloneNode() {
        return new Element(this.localName, this.text, this.children.map(c => c.cloneNode()), {
            ...this.attrs
        });
    }
}
const e = (n, t = '', cs = [], a = {}) => new Element(n, t, cs, a),
    p = (t, list) => ({
        text: t,
        node: e('p', t),
        list
    }),
    direct = (n, k) => (n?.children || []).filter(c => c.localName === k),
    nodes = (n, k) => n.getElementsByTagNameNS(W, k),
    val = (n, k = 'val') => n?.attrs[k],
    norm = s => s.replace(/\s+/g, ' ').trim();
const ctx = {
    W,
    direct,
    nodes,
    val,
    norm,
    textOf: n => n.textContent,
    blockText: n => n.textContent,
    visiblePara: n => ({
        node: n.cloneNode()
    }),
    addIssue: (m, o) => m.issues.push(o)
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(new URL('../docs/js/matrix-reader.js', import.meta.url), 'utf8').replace(/^import[\s\S]*?from '\.\/xml\.js';/m, '').replace(/export \{[\s\S]*?\};/g, ''), ctx);
const table = headers => e('tbl', '', [e('tr', '', headers.map(t => e('tc', t)))]);
for (const title of ['RANCANGAN AWAL', 'RANCANGAN PERBUP', 'RUMUSAN AWAL']) {
    const h = ctx.matrixHeader(table(['No', 'Materi', title, 'HASIL PENYESUAIAN', 'Catatan']));
    assert.equal(h.cols.awal, 2);
    assert.equal(h.cols.hasil, 3);
}
const swapped = ctx.matrixHeader(table(['Materi', 'HASIL PENYEMPURNAAN', 'RANCANGAN PERDA']));
assert.equal(swapped.cols.hasil, 1);
assert.equal(swapped.cols.awal, 2);
const span = e('tc', '', [e('tcPr', '', [e('gridSpan', '', [], {
    val: '2'
})])]);
const cells = ctx.matrixCells(e('tr', '', [span, e('tc', 'X')]));
assert.equal(cells[0], cells[1]);
assert.equal(cells[2].textContent, 'X');
assert.equal(ctx.matrixRowLabel([e('tc', 'Menimbang'), e('tc', ':'), e('tc', 'awal')], 2).prefix, 'Menimbang');
const left = ['a', 'b', 'c', 'd'].map(x => p('Isi awal ' + x, {
    marker: x + '.',
    kind: 'letter'
}));
const m = {
    issues: [],
    decisions: []
};
let selected = ctx.matrixSelectRumusan(left, [p('Tetap', {
    marker: 'a.'
}), p('Rumusan baru b', {
    marker: 'b.'
}), p('c. Tetap', {
    marker: 'c.',
    prefixLength: 3
}), p('d. Tetap', {
    marker: 'd.',
    prefixLength: 3
})], {
    m,
    rowIndex: 7,
    prefix: 'Menimbang'
});
assert.deepEqual(Array.from(selected.ps, x => x.text), ['Isi awal a', 'Rumusan baru b', 'Isi awal c', 'Isi awal d']);
assert.equal(m.issues.length, 0);
selected = ctx.matrixSelectRumusan(left, [p('Tetap')], {
    m,
    rowIndex: 1
});
assert.equal(selected.ps.length, 4);
assert.equal(selected.useRight, false);
assert.equal(ctx.matrixSelectRumusan(left, [], {
    m,
    rowIndex: 1
}).ps.length, 4);
ctx.matrixSelectRumusan(left, [p('z. Tetap', {
    marker: 'z.'
})], {
    m,
    rowIndex: 1
});
assert.equal(m.issues.at(-1).hard, true);
// A deleted list paragraph still advances Word's original numbering.
const lvl = e('lvl', '', [e('start', '', [], {
        val: '1'
    }), e('numFmt', '', [], {
        val: 'lowerLetter'
    }), e('lvlText', '', [], {
        val: '%1.'
    })], {
        ilvl: '0'
    }),
    abs = e('abstractNum', '', [lvl], {
        abstractNumId: '1'
    }),
    num = e('num', '', [e('abstractNumId', '', [], {
        val: '1'
    })], {
        numId: '9'
    });
const catalog = {
    nums: new Map([
        ['9', num]
    ]),
    abstracts: new Map([
        ['1', abs]
    ]),
    styleMap: new Map()
};
const native = t => e('p', '', [e('pPr', '', [e('numPr', '', [e('ilvl', '', [], {
    val: '0'
}), e('numId', '', [], {
    val: '9'
})])]), e('t', t)]);
const cell = e('tc', '', [native('first'), native(''), native('third')]);
const result = ctx.matrixParagraphs(cell, catalog, new Map());
assert.equal(result.length, 2);
assert.equal(result[1].list.marker, 'c.');
console.log('PASS: header aliases/reordered columns, merged grids, split labels, empty/Tetap/mixed items, unmatched-item blocking, hidden-item numbering.');
const fullArticle = [p('BAB I'), p('KETENTUAN UMUM'), p('Pasal 1'), p('Dalam Peraturan ini:'), ...Array.from({length:21},(_,i)=>p('Definisi '+(i+1),{marker:(i+1)+'.',kind:'number'}))];
const partialModel = {issues:[],decisions:[]};
const merged = ctx.matrixSelectRumusan(fullArticle,[p('LMS direvisi',{marker:'18.',kind:'number'}),p('Teknologi direvisi',{marker:'19.',kind:'number'})],{m:partialModel,rowIndex:10,prefix:''});
assert.equal(merged.ps.length, fullArticle.length);
assert.equal(merged.ps[2].text,'Pasal 1');
assert.equal(merged.ps[21].text,'LMS direvisi');
assert.equal(merged.ps[22].text,'Teknologi direvisi');
assert.equal(merged.ps.at(-1).text,'Definisi 21');
assert.equal(partialModel.issues.length,1);
const ambiguous={issues:[],decisions:[]};
ctx.matrixSelectRumusan(fullArticle,[p('Tidak ada penanda')],{m:ambiguous,rowIndex:10,prefix:''});
assert.equal(ambiguous.issues[0].hard,true);
console.log('PASS: partial revision preserves chapter, article, introduction and all 21 definitions; ambiguous fragment blocks export.');
const renamed=ctx.matrixSelectRumusan(
 [p('BAB V'),p('KETENTUAN PENUTUP'),p('Pasal 26'),p('Peraturan lama dicabut.')],
 [p('BAB V'),p('KETENTUAN PENUTUP'),p('Pasal 29'),p('Tetap')],
 {m:{issues:[],decisions:[]},rowIndex:35,prefix:''}
);
assert.deepEqual(Array.from(renamed.ps,p=>p.text),['BAB V','KETENTUAN PENUTUP','Pasal 29','Peraturan lama dicabut.']);
const struckBody=p('Rumusan lama tetap digunakan.');
struckBody.sourceArticle='26';
const struck=ctx.matrixSelectRumusan([p('BAB V'),struckBody],[p('Pasal 29'),p('Tetap')],{m:{issues:[],decisions:[]},rowIndex:35,prefix:''});
assert.equal(struck.ps.at(-1).text,'Rumusan lama tetap digunakan.');
console.log('PASS: whole-article Tetap preserves body with a renamed or struck source title.');
