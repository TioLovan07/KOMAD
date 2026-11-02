/** Compare all selected rows before declaring an article absent.
 * A matrix row is a source coordinate, not an article boundary.
 */
export function missingArticles(rows, sections) {
    const selected = new Set(rows.flatMap(row => sections(row.ps).map(s => s.num)));
    return rows.flatMap(row => {
        if (!row.useRight || /hapus|pindah|reposisi/i.test(row.note)) return [];
        const missing = sections(row.left).map(s => s.num).filter(num => !selected.has(num));
        if (!missing.length) return [];
        // Amendment matrices may deliberately change the target article in the result.
        // This is a source/result difference to confirm, not proof of dropped draft text.
        const targetChange = row.ps.some(p => p.amendmentInstruction) &&
            row.left.some(p => /^\s*(?:\d+[.)]\s*)?Ketentuan\b.*\bPasal\s+\d+/i.test(p.text));
        const oldSections=sections(row.left),newSections=sections(row.ps);
        const normText=ps=>ps.map(p=>p.text).join(' ').replace(/\s+/g,' ').trim().toLowerCase();
        const oldBody=oldSections.length===1?normText(row.left.slice(oldSections[0].start+1,oldSections[0].end)):'';
        const newBody=newSections.length===1?normText(row.ps.slice(newSections[0].start+1,newSections[0].end)):'';
        const sameLead=oldBody.length>60&&newBody.length>60&&oldBody.slice(0,60)===newBody.slice(0,60);
        const resultTargets = targetChange||sameLead ? [...new Set(sections(row.ps).map(s=>s.num))] : [];
        return [{row, missing:[...new Set(missing)], ...(resultTargets.length ? {resultTargets} : {})}];
    });
}
