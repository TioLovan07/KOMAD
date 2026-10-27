/** Compare all selected rows before declaring an article absent.
 * A matrix row is a source coordinate, not an article boundary.
 */
export function missingArticles(rows, sections) {
    const selected = new Set(rows.flatMap(row => sections(row.ps).map(s => s.num)));
    return rows.flatMap(row => {
        if (!row.useRight || /hapus|pindah|reposisi/i.test(row.note)) return [];
        const missing = sections(row.left).map(s => s.num).filter(num => !selected.has(num));
        return missing.length ? [{row, missing:[...new Set(missing)]}] : [];
    });
}
