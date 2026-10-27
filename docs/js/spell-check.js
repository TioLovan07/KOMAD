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
        for (const {word, suggestions} of result.unknown.slice(0, 200)) addIssue(model, {
            kind: 'correction', spelling: true, word, suggestions,
            replacement: suggestions[0] || '',
            title: 'Periksa ejaan: “' + word + '”',
            message: 'Kata belum dikenali kamus. Nama, istilah khusus, atau kata asing dapat benar. Saran berdasarkan kemiripan ejaan; pilih sendiri, bukan penggantian otomatis.'
        });
        model.decisions.push('Kamus offline: ' + words.length + ' kata unik diperiksa; ' +
            result.unknown.length + ' kata belum dikenali. Singkatan huruf kapital dilewati. Pemeriksaan tidak menilai konteks atau substansi hukum.');
        if (result.unknown.length > 200) addIssue(model, {
            kind: 'review', title: 'Sebagian catatan ejaan belum ditampilkan',
            message: 'Ditemukan ' + result.unknown.length + ' kata belum dikenali; 200 ditampilkan agar panel tetap dapat digunakan. Periksa ejaan seluruh dokumen di Word juga.'
        });
    } catch (error) {
        addIssue(model, {kind:'review', title:'Pemeriksaan kamus belum selesai',
            message:error.message + ' Deteksi typo terbatas tetap berjalan. Jangan menganggap draft bebas kesalahan.'});
    } finally { worker?.terminate(); }
}
