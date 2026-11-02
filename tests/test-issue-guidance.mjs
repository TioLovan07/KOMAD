import assert from 'node:assert/strict';
import {issueGuidance} from '../docs/js/issue-guidance.js';
import {annotateLists} from '../docs/js/list-layout.js';
const p=(text,marker,kind)=>({text,node:{localName:'p'},list:marker?{marker,kind,level:0}:null});
const m={rows:[{index:10,label:'Isi',ps:[p('Pasal 1'),p('Dalam Peraturan ini yang dimaksud dengan:'),p('Definisi tiga puluh dua','32.','number'),p('Definisi tiga puluh empat','34.','number'),p('Pasal 12'),p('Isi ayat tujuh','(7)','ayat'),p('Isi ayat tiga','(3)','ayat'),p('Rincian','g.','letter')]}],issues:[]};
annotateLists(m);
const jump=m.issues.find(i=>i.numbering?.previous==='32.'&&i.numbering.current==='34.');assert.ok(jump);assert.match(jump.message,/33\./);assert.match(jump.action,/dicoret\/dihapus/);assert.equal(jump.blockIndex,3);
const back=m.issues.find(i=>i.numbering?.previous==='(7)'&&i.numbering.current==='(3)');assert.match(back.message,/mundur/);
const letter=m.issues.find(i=>i.numbering?.current==='g.');assert.match(letter.message,/a\./);assert.match(letter.action,/anak daftar/);
for(const i of m.issues){const g=issueGuidance(i);assert.ok(g.problem.length>20);assert.ok(g.action.length>20);}
for(const kind of ['correction','source-gap','editorial','advisory','annex','review']){const g=issueGuidance({kind,title:'Periksa hasil pembacaan DOC 97–2003',message:'Pesan pemeriksaan',word:'salah',replacement:'benar'});assert.ok(g.action.length>20);}
assert.deepEqual(m.rows[0].ps[3].list.marker,'34.');
console.log('PASS: exact 32→34/33 missing, 7→3, g/a explanations and actionable guidance without renumbering.');
