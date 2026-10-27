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
