/** Offline Hunspell-compatible engine, isolated from the UI thread. */
import nspell from '../vendor/nspell.js';
let engine;
const terms = 'nonalam suburusan pascabencana kedaruratan terencana kesiapsiagaan kesiapan berakhlak berkepribadian luhur monitoring kependudukan kependidikan berintegritas fisik-motorik bali komad ranperbup ranperda perkada perda perbup pemda pemkab pemprov kanwil kemenkum p3h corpu lms asn anev jdih buleleng gianyar singaraja denpasar badung bangli karangasem tabanan jembrana klungkung posbankum'.split(' ');
async function initialize() {
    if (engine) return;
    const base = new URL('../assets/dictionary/', import.meta.url);
    const responses = await Promise.all(['id_ID.aff', 'id_ID.dic'].map(name => fetch(new URL(name, base))));
    if (responses.some(r => !r.ok)) throw Error('Berkas kamus offline belum tersedia.');
    const [aff, dic] = await Promise.all(responses.map(r => r.text()));
    engine = nspell(aff, dic);
    terms.forEach(word => engine.add(word));
}
function recognized(word) {
    if (engine.correct(word)) return true;
    // Clitics/reduplication are checked against known forms, never blindly accepted.
    const unclitic = word.replace(/(?:nya|lah|kah|pun)$/, '');
    if (unclitic !== word && engine.correct(unclitic)) return true;
    const parts = word.split('-');
    return parts.length === 2 && parts[0] === parts[1] && engine.correct(parts[0]);
}
function candidates(word) {
    // One-edit candidates including adjacent swaps; never invent replacement text.
    const result = new Set(), alphabet = 'abcdefghijklmnopqrstuvwxyz';
    const accept = s => { if (s !== word && engine.correct(s)) result.add(s); };
    for (let i = 0; i < word.length; i++) {
        accept(word.slice(0, i) + word.slice(i + 1));
        if (i + 1 < word.length) accept(word.slice(0, i) + word[i + 1] + word[i] + word.slice(i + 2));
        for (const c of alphabet) accept(word.slice(0, i) + c + word.slice(i + 1));
    }
    for (let i = 0; i <= word.length; i++)
        for (const c of alphabet) accept(word.slice(0, i) + c + word.slice(i));
    return [...result].sort().slice(0, 5);
}
self.onmessage = async event => {
    try {
        await initialize();
        const {words, custom} = event.data;
        const known = new Set(custom);
        const unknown = [];
        for (const word of words) {
            const low = word.toLocaleLowerCase('id');
            if (known.has(low) || recognized(low)) continue;
            const proposed = word.length <= 32 ? candidates(low) : [];
            const caseLike = s => word === word.toUpperCase() ? s.toUpperCase() : /^[A-Z]/.test(word) ? s[0].toUpperCase() + s.slice(1) : s;
            unknown.push({word, suggestions: proposed.map(caseLike)});
        }
        self.postMessage({unknown});
    } catch (error) { self.postMessage({error: error.message}); }
};
