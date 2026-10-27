/** Conservative triage: a dictionary miss is not itself a typo. */
export function strongCandidate(word, suggestion) {
    const a = word.toLocaleLowerCase('id'), b = suggestion.toLocaleLowerCase('id');
    if (a.length < 6 || b.length < 6) return false;
    if (a.length === b.length) {
        const positions = [...a].map((c,i)=>c !== b[i] ? i : -1).filter(i=>i>=0);
        return positions.length === 2 && positions[1] === positions[0]+1 &&
            a[positions[0]] === b[positions[1]] && a[positions[1]] === b[positions[0]];
    }
    if (Math.abs(a.length-b.length) !== 1 || a.slice(0,3)!==b.slice(0,3) || a.slice(-3)!==b.slice(-3)) return false;
    const longer=a.length>b.length?a:b, shorter=a.length>b.length?b:a;
    return [...longer].some((_,i)=>i>=3 && i<longer.length-3 &&
        (longer[i] === longer[i-1] || longer[i] === longer[i+1]) &&
        longer.slice(0,i)+longer.slice(i+1)===shorter);
}
export function triageSpelling(unknown) {
    const flagged=[], unconfirmed=[];
    for (const entry of unknown) {
        // Capitalized unknowns may be names/places; do not promote them automatically.
        if (/^\p{Lu}/u.test(entry.word)) { unconfirmed.push(entry.word); continue; }
        const plausible=entry.suggestions.filter(s=>strongCandidate(entry.word,s));
        if (plausible.length===1) flagged.push({...entry,suggestions:plausible});
        else unconfirmed.push(entry.word);
    }
    return {flagged,unconfirmed};
}
