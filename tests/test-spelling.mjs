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
await self.onmessage({data:{words:['penelaahan','menelaah','pertanggungjawaban','mempertanggungjawabkan','organisasi','pengorganisasian','keorganisasian'],custom:[]}});
assert.deepEqual(response.unknown,[]);
for (const [wrong,right] of [['penelahaan','penelaahan'],['pertangungjawaban','pertanggungjawaban'],['oragnisasi','organisasi']]) assert.ok(SPELLING_CORRECTIONS.some(([a,b])=>a===wrong && b===right));
console.log('PASS: screenshot typos have explicit corrections; related correct terms are accepted.');

await self.onmessage({data:{words:['Kotamadya','Paralegal','Nonlitigasi','terakreditasi','direkrut','drafting','mengadvokasi','dipidana','secondary','employment','moonlighting','revolving','door','cooling','period','beneficial','ownership','selambat-lambatnya','off','difasilitasi','dugaan','terlapor','online','offline'],custom:[]}});
assert.deepEqual(response.unknown,[]);
for (const [a,b] of [['dentifikasi','identifikasi'],['disetiap','di setiap'],['dimasing-masing','di masing-masing'],['dimana','di mana'],['tatacara','tata cara'],['dilingkungan','di lingkungan']]) assert.ok(SPELLING_CORRECTIONS.some(p=>p[0]===a&&p[1]===b));
console.log('PASS: requested legal and foreign terms plus six explicit corrections.');

await self.onmessage({data:{words:['ketidaktaatan','ketidakpatuhan','dipersyaratkan','mutatis','mutandis','ketidaksesuaian','diindikasikan','positioning','system','tracker','rincian','KKPR-nya','pengendalian','menaati','sanksi','administratif','perundang-undangan'],custom:[]}});
assert.deepEqual(response.unknown,[]);
for(const [a,b] of [['pengedalian','pengendalian'],['yarng','yang'],['Sanski','Sanksi'],['Administratrif','Administratif'],['mentaati','menaati'],['sebaimana','sebagaimana'],['administatif','administratif'],['administrastif','administratif'],['peraturan-perundang-undangan','peraturan perundang-undangan'],['perundangan-undangan','perundang-undangan']])assert.ok(SPELLING_CORRECTIONS.some(p=>p[0]===a&&p[1]===b));
console.log('PASS: legal terms and spelling corrections 30 October 2026.');

assert.ok(SPELLING_CORRECTIONS.some(p=>p[0]==='KKPRnya'&&p[1]==='KKPR-nya'));

await self.onmessage({data:{words:['penatausahaan','Penatausahaan','Undang-Undang','dasawisma','Parisadha','Bhayangkari','Persit','sub','dianggarkan','dirinci','Portable','Document','e-mail','direnovasi','pemeringkatan','Maret','Mei','April','Juni','SPP-LS','SPM-LS','Perundang','ditransfer','Januari','menganggarkan'],custom:[]}});
assert.deepEqual(response.unknown,[]);
for(const [a,b] of [['Bagiam','Bagian'],['bagiam','bagian'],['Plaporan','Pelaporan'],['Monitorng','Monitoring'],['Perundangundangan','Perundang-undangan'],['bermaterai','bermeterai'],['kerkenaan','berkenaan'],['disetempel','distempel'],['besertajasa','beserta jasa'],['Akuntasi','Akuntansi']])assert.ok(SPELLING_CORRECTIONS.some(p=>p[0]===a&&p[1]===b));
console.log('PASS: 1.3.4 accepted terms and explicit case-aware typo recommendations.');

await self.onmessage({data:{words:['terukur','Reduce','Reuse','Recycle','pemilahan','pendauran','e-BK','website','terverifikasi','feasibility','study','appraisal','permusyawaratan','mempedomani','output','pemindahbukuan','dipersamakan','lumpsum','lumsum','force','majeure','tertuang','e-Hibah','Agrikultur'],custom:[]}});
assert.deepEqual(response.unknown,[]);
for(const [a,b] of [['isingkat','disingkat'],['documentformat','document format'],['merubah','mengubah'],['dirnaksud','dimaksud'],['perkegiatan','per kegiatan'],['nimimal','minimal'],['dipertanggungjawaban','dipertanggungjawabkan'],['Undang-Undang-Undang','Undang-Undang'],['dibidang','di bidang']])assert.ok(SPELLING_CORRECTIONS.some(p=>p[0]===a&&p[1]===b));
console.log('PASS: added waste management, administration and foreign terms; nine new typo cases.');
