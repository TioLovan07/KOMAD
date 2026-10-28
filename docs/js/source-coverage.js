/** Compare content independently of Word paragraph/table boundaries. */
export function coverageKey(block) {
    let text = block.text || '';
    if (block.list?.prefixLength) text = text.slice(block.list.prefixLength);
    return text.normalize('NFKC').toLowerCase().replace(/[^\p{L}\p{N}]/gu, '');
}
export function missingSourceBlocks(source, result) {
    const joined = result.map(coverageKey).join('');
    return source.filter(block => {
        const key = coverageKey(block);
        return key && !joined.includes(key);
    });
}
