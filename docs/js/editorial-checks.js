/** Kamus koreksi terbatas, bukan pemeriksa ejaan bahasa Indonesia menyeluruh. */
export const SPELLING_CORRECTIONS = [
    ['Paraturan', 'Peraturan'],
    ['Selanjuntnya', 'Selanjutnya'],
    ['TK TK', 'TK'],
    ['sebalah', 'sebelah'],
    ['dilengakpi', 'dilengkapi'],
    ['Dilengakpi', 'Dilengkapi'],
    ['DILENGAKPI', 'DILENGKAPI'],
];

// Explicit corrections: wrong forms are never added as accepted dictionary entries.
const additions = [
 ['terakomodir','terakomodasi'], ['komperensif','komprehensif'],
 ['restra','renstra'], ['priorotas','prioritas'], ['anlisis','analisis'], ['sistema','sistem'],
 ['koordinai','koordinasi'], ['singkronisasi','sinkronisasi'],
 ['efcktivitas','efektivitas'], ['penyelengaraan','penyelenggaraan'],
 ['konflilk','konflik'], ['Iuhur','luhur'], ['bekerjasama','bekerja sama'],
 ['diatas','di atas'], ['dilapangan','di lapangan'], ['dibelahan','di belahan'],
 ['diluar','di luar'], ['diantaranya','di antaranya'], ['diantara','di antara'], ['mensupervisi','menyupervisi']
];
for (const [word,replacement] of additions) {
    SPELLING_CORRECTIONS.push([word,replacement]);
    SPELLING_CORRECTIONS.push([word[0].toUpperCase()+word.slice(1),replacement[0].toUpperCase()+replacement.slice(1)]);
    SPELLING_CORRECTIONS.push([word.toUpperCase(),replacement.toUpperCase()]);
}
export function isEditorialInstruction(text) {
    return /^\s*(?:Mohon|Harap)\s+(?:untuk\s+)?(?:dilengakpi|dilengkapi|ditambahkan|diperbaiki)\b/i.test(text);
}
