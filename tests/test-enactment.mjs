import assert from 'node:assert/strict';
import {selectEnactment} from '../docs/js/enactment.js';
const p=text=>({text,node:{cloneNode(){return this}}});
const old=[p('MEMUTUSKAN :'),p('PERATURAN BUPATI TENTANG CONTOH.')];
for(const result of [[],[p('Tetap')],[p('MEMUTUSKAN :')],[p('MEMUTUSKAN :'),p('Tetap')]])assert.deepEqual(selectEnactment(old,result,{prefix:'Menetapkan'}).ps.map(p=>p.text),old.map(p=>p.text));
assert.equal(selectEnactment(old,[p('MEMUTUSKAN :'),p('PERATURAN BUPATI TENTANG BARU.')],{prefix:'Menetapkan'}),null);
assert.deepEqual(selectEnactment([old[1]],[],{prefix:'Menetapkan'}).ps.map(p=>p.text),[old[1].text]);
console.log('PASS: blank, Tetap, heading-only, separated decision and complete replacement semantics.');
