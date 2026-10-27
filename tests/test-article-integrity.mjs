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
