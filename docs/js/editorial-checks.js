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
 ['isingkat','disingkat'], ['documentformat','document format'], ['merubah','mengubah'], ['dirnaksud','dimaksud'], ['perkegiatan','per kegiatan'], ['nimimal','minimal'], ['dipertanggungjawaban','dipertanggungjawabkan'], ['dibidang','di bidang'], ['undang-undang-undang','undang-undang'], ['Undang-Undang-Undang','Undang-Undang'],
 ['adalan','adalah'], ['disngkat','disingkat'], ['disaingkat','disingkat'], ['meujudkan','mewujudkan'], ['dearah','daerah'], ['akuntasi','akuntansi'], ['diperuntukan','diperuntukkan'], ['sebagaiman','sebagaimana'], ['standarisasi','standardisasi'], ['meminimalisir','meminimalkan'], ['mengakomodir','mengakomodasi'],
 ['bagiam','bagian'], ['plaporan','pelaporan'], ['monitorng','monitoring'], ['perundangundangan','perundang-undangan'], ['bermaterai','bermeterai'], ['kerkenaan','berkenaan'], ['disetempel','distempel'], ['besertajasa','beserta jasa'],
 ['KKPRnya','KKPR-nya'], ['kkprnya','KKPR-nya'],
 ['pengedalian','pengendalian'], ['yarng','yang'], ['sanski','sanksi'], ['administratrif','administratif'], ['mentaati','menaati'], ['sebaimana','sebagaimana'], ['administatif','administratif'], ['administrastif','administratif'], ['peraturan-perundang-undangan','peraturan perundang-undangan'], ['perundangan-undangan','perundang-undangan'],
 ['dentifikasi','identifikasi'], ['disetiap','di setiap'], ['dimasing-masing','di masing-masing'], ['dimana','di mana'], ['tatacara','tata cara'], ['dilingkungan','di lingkungan'],
 ['penelahaan','penelaahan'], ['pertangungjawaban','pertanggungjawaban'], ['oragnisasi','organisasi'],
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
