import {triageSpelling} from './spelling-policy.js';
/** Word checks are advisory. Unknown does not mean incorrect. Documents never leave this device. */
import {addIssue} from './xml.js';
const KEY = 'komad-custom-words-v1';
export function customWords() {
    try {
        const value = JSON.parse(localStorage.getItem(KEY) || '[]');
        return Array.isArray(value) ? value.filter(w => typeof w === 'string').slice(0, 10000) : [];
    } catch { return []; }
}
export function rememberWord(word) {
    const words = new Set(customWords());
    words.add(word.toLocaleLowerCase('id'));
    try { localStorage.setItem(KEY, JSON.stringify([...words])); return true; }
    catch { return false; }
}
export async function checkSpelling(model) {
    const existing = new Set(model.issues.filter(i => i.kind === 'correction').map(i => i.word));
    const words = [...new Set(model.rows.flatMap(row => row.ps.flatMap(p =>
        [...p.text.matchAll(/\p{L}+(?:-\p{L}+)*/gu)].map(m => m[0]))))].filter(w =>
            w.length > 2 && !existing.has(w) && !/^[A-Z]{2,}$/.test(w));
    let worker;
    try {
        worker = new Worker(new URL('./spell-worker.js', import.meta.url), {type: 'module'});
        const result = await new Promise((resolve, reject) => {
            const timer = setTimeout(() => reject(Error('Pemeriksaan ejaan melewati batas waktu 60 detik.')), 60000);
            worker.onmessage = event => {clearTimeout(timer); resolve(event.data);};
            worker.onerror = () => {clearTimeout(timer); reject(Error('Mesin ejaan tidak dapat dimuat.'));};
            worker.postMessage({words, custom: customWords()});
        });
        if (result.error) throw Error(result.error);
        const {flagged, unconfirmed} = triageSpelling(result.unknown);
        model.unconfirmedWords = unconfirmed;
        for (const {word, suggestions} of flagged.slice(0, 200)) addIssue(model, {
            kind: 'correction', spelling: true, word, suggestions,
            replacement: suggestions[0] || '',
            title: 'Dugaan typo: “' + word + '”',
            message: 'Ditemukan pola salah ketik dengan satu kandidat yang lolos penyaringan. Periksa konteks sebelum menerapkan.'
        });
        model.decisions.push('Kamus offline: ' + words.length + ' kata unik diperiksa; ' +
            flagged.length + ' dugaan typo; ' + unconfirmed.length + ' kata belum dikenali tidak dijadikan catatan wajib. Singkatan huruf kapital dilewati. Pemeriksaan tidak menilai konteks atau substansi hukum.');
        if (flagged.length > 200) addIssue(model, {
            kind: 'review', title: 'Sebagian catatan ejaan belum ditampilkan',
            message: 'Ditemukan ' + flagged.length + ' kata belum dikenali; 200 ditampilkan agar panel tetap dapat digunakan. Periksa ejaan seluruh dokumen di Word juga.'
        });
    } catch (error) {
        addIssue(model, {kind:'review', title:'Pemeriksaan kamus belum selesai',
            message:error.message + ' Deteksi typo terbatas tetap berjalan. Jangan menganggap draft bebas kesalahan.'});
    } finally { worker?.terminate(); }
}

/** Recheck newly accepted source text; preserve earlier word decisions by identity. */
export async function refreshSpelling(model) {
    const [{SPELLING_CORRECTIONS},{replacementMatches}] = await Promise.all([
        import('./editorial-checks.js'), import('./draft-model.js')
    ]);
    const prior=new Map(model.issues.filter(i=>i.kind==='correction').map(i=>[i.word,i]));
    model.issues=model.issues.filter(i=>i.kind!=='correction');
    const full=model.rows.flatMap(r=>r.ps.map(p=>p.text)).join('\n');
    for(const [word,replacement] of new Map(SPELLING_CORRECTIONS))
        if(replacementMatches(full,word).length) addIssue(model,{
            kind:'correction',word,replacement,
            title:'“'+word+'” → “'+replacement+'”',
            message:'Periksa saran ejaan sebelum menerapkan pada kemunculan yang ditampilkan.'
        });
    await checkSpelling(model);
    // IDs must be unique even after removing/rebuilding correction cards.
    model.issues.forEach((issue,id)=>{
        issue.id=id;
        if(issue.kind==='correction' && prior.has(issue.word)) {
            const old=prior.get(issue.word);
            issue.choice=old.choice; issue.resolved=old.resolved;
            if(!issue.suggestions || issue.suggestions.includes(old.replacement)) issue.replacement=old.replacement;
        }
    });
}
