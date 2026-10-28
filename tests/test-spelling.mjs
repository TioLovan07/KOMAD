import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
const nativeFetch = globalThis.fetch;
globalThis.fetch = async url => {
    if (url.protocol === 'file:') return {ok:true, text:()=>fs.readFile(url,'utf8')};
    return nativeFetch(url);
};
let response;
globalThis.self = {postMessage: result => {response=result;}};
await import('../docs/js/spell-worker.js');
await self.onmessage({data:{words:['mengambarkan','dilengakpi','pemerintah','kompetensinya','isu-isu','Corpu','IstilahKantorUji'],custom:['istilahkantoruji']}});
assert.equal(response.error,undefined);
assert.deepEqual(response.unknown.map(x=>x.word),['mengambarkan','dilengakpi']);
assert.ok(response.unknown[0].suggestions.includes('menggambarkan'));
assert.ok(response.unknown[1].suggestions.includes('dilengkapi'));
console.log('PASS: offline dictionary, affixes, adjacent swaps, missing letters, clitics, reduplication, office terms and custom terms.');
const {triageSpelling}=await import('../docs/js/spelling-policy.js');
const triage=triageSpelling([
 {word:'mengambarkan',suggestions:['mengabarkan','mengamarkan','mengembarkan','menggambarkan']},
 {word:'terintegrasi',suggestions:['berintegrasi']},
 {word:'Corporate',suggestions:[]},
 {word:'Undisan',suggestions:['Undissan']},
 {word:'dimana',suggestions:['dimana-mana']}
]);
assert.deepEqual(triage.flagged.map(x=>x.word),['mengambarkan']);
assert.equal(triage.unconfirmed.length,4);
console.log('PASS: unknown terms do not become required notes; prefix substitution rejected, actual missing-letter candidate retained.');

await self.onmessage({data:{words:['ketidakcukupan','mengestimasi','Reviu','menatausahakan','stakeholders','terakomodasi','komprehensif','Renstra'],custom:[]}});
assert.equal(response.error,undefined);
assert.deepEqual(response.unknown,[]);
const {SPELLING_CORRECTIONS}=await import('../docs/js/editorial-checks.js');
for (const [a,b] of [['terakomodir','terakomodasi'],['komperensif','komprehensif'],['Restra','Renstra'],['priorotas','prioritas'],['anlisis','analisis'],['sistema','sistem']]) assert.ok(SPELLING_CORRECTIONS.some(pair=>pair[0]===a && pair[1]===b));
console.log('PASS: requested accepted terms and explicit spelling corrections.');
