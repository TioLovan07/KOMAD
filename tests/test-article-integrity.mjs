import assert from 'node:assert/strict';
import {missingArticles} from '../docs/js/article-integrity.js';
const sections=ps=>ps.filter(p=>/^Pasal \d+$/.test(p.text)).map(p=>({num:p.text.split(' ')[1]}));
const ps=(...items)=>items.map(text=>({text}));
const rows=[
 {index:18,useRight:true,note:'',left:ps('BAB V','Pasal 9'),ps:ps('BAB III','Pasal 8')},
 {index:19,useRight:true,note:'',left:[],ps:ps('Pasal 9','Peraturan ini mulai berlaku.')}
];
assert.deepEqual(missingArticles(rows,sections),[]);
assert.deepEqual(missingArticles(rows.slice(0,1),sections)[0].missing,['9']);
console.log('PASS: Peninjoan article 9 on next row is retained; truly absent article is flagged.');
const amendmentRows = [{index:11,useRight:true,note:'',
 left:ps('Ketentuan pada ayat (2) Pasal 14 diubah, sehingga Pasal 14 berbunyi sebagai berikut:','Pasal 14','Isi awal.'),
 ps:[{text:'Ketentuan huruf a ayat (2) Pasal 15 diubah sehingga Pasal 15 berbunyi sebagai berikut:',amendmentInstruction:true},...ps('Pasal 15','Isi hasil.')]}];
const changed = missingArticles(amendmentRows,sections);
assert.deepEqual(changed[0].missing,['14']);
assert.deepEqual(changed[0].resultTargets,['15']);
assert.equal(missingArticles(rows.slice(0,1),sections)[0].resultTargets,undefined);
console.log('PASS: explicit amendment target differences are distinguished from actually absent articles.');
