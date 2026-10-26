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

export function isEditorialInstruction(text) {
    return /^\s*(?:Mohon|Harap)\s+(?:untuk\s+)?(?:dilengakpi|dilengkapi|ditambahkan|diperbaiki)\b/i.test(text);
}
