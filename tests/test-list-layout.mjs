import assert from 'node:assert/strict';
import {annotateLists} from '../docs/js/list-layout.js';
const p=(text,marker,kind,sourceIndent)=>({text,node:{localName:'p'},sourceIndent,...(marker?{list:{marker,kind,level:0}}:{})});
const row={index:1,label:'Batang Tubuh',ps:[p('Pasal 12'),p('(3) Ketentuan berikut:','(3)','ayat',628),p('a. keberangkatan','a.','letter',1053),p('1. dari kantor','1.','number',1478),p('2. dari terminal','2.','number',1478),p('b. kepulangan','b.','letter',1053),p('1. dari tujuan','1.','number',1478),p('(6) Ketentuan berikut','(6)','ayat',628),p('g. biaya transport','g.','letter',961),p('(7) Biaya lain','(7)','ayat',628),p('(3) Biaya berikut','(3)','ayat',628),p('Pasal 15'),p('(1) Biaya jenazah meliputi:','(1)','ayat',583),p('a. biaya pertama','a.','letter',1053),p('b. biaya kedua','b.','letter',1053),p('Pejabat Negara dan ASN.',null,null,551),p('Pasal 1'),p('Dalam peraturan ini yang dimaksud dengan:'),p('17. Risiko Operasional disebabkan oleh:','17.','number',601),p('a. ketidakcukupan','a.','letter',1026),p('b. kejadian lain','b.','letter',1026)]};
const m={rows:[row],issues:[]};annotateLists(m);
assert.equal(row.ps[3].layoutLevel,2);assert.equal(row.ps[4].layoutLevel,2);assert.equal(row.ps[6].layoutLevel,2);assert.equal(row.ps[15].continuationLevel,0);assert.equal(row.ps[19].layoutLevel,1);assert.equal(row.ps[20].layoutLevel,1);
assert.ok(m.issues.some(i=>i.blockIndex===8));assert.ok(m.issues.some(i=>i.blockIndex===10));assert.ok(!m.issues.some(i=>i.blockIndex===6));
assert.equal(row.ps[10].list.marker,'(3)');
console.log('PASS: nested numbers, definition children, continuation ownership, suspicious markers retained and located.');

const amend={rows:[{index:2,label:'Perubahan',ps:[{...p('Ketentuan Pasal 1 diubah','1.','number',0),amendmentInstruction:true},p('Pasal 1'),p('Dalam peraturan ini yang dimaksud dengan:'),p('1. Definisi','1.','number',600),{...p('Ketentuan Pasal 8 diubah','2.','number',0),amendmentInstruction:true}]}],issues:[]};annotateLists(amend);assert.equal(amend.issues.length,0);

const flat={rows:[{index:3,label:'Isi',ps:[p('Pasal 1'),p('(1) Rincian:','(1)','ayat',600),p('a. rincian','a.','letter',600)]}],issues:[]};annotateLists(flat);assert.ok(flat.issues.some(i=>i.title.startsWith('Periksa indent')));

const dup={rows:[{index:7,label:'Perubahan',ps:[p('Pasal I'),{...p('1. Ketentuan Pasal 14 diubah','1.','number',0),amendmentInstruction:true},p('Pasal 14'),p('(1) Isi lama','(1)','ayat',600),{...p('1. Ketentuan Lampiran diubah','1.','number',0),amendmentInstruction:true}]}],issues:[]};
annotateLists(dup);assert.ok(dup.issues.some(i=>i.numbering?.change==='berulang'));assert.equal(dup.rows[0].ps[4].list.marker,'1.');
console.log('PASS: repeated amendment numbers survive intervening Pasal headings and are reported without renumbering.');
