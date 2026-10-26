/** Sorotan hanya hidup di iframe pratinjau; tidak dimasukkan ke file Word. */
import {
    $
} from './xml.js';
import {
    model
} from './state.js';
import {
    effectiveRows,
    replacementMatches
} from './draft-model.js';

export function highlightIssue(issue, scroll = false, locationIndex = null) {
    const viewer = $('draft-frame')?.contentWindow;
    if (!viewer?.highlightDraft) return 0;
    viewer.clearDraftHighlights();
    if (!issue || !model) return 0;
    const rows = effectiveRows();
    const targets = [];
    const locations = locationIndex === null ? issue.locations : [issue.locations[locationIndex]];
    for (const location of locations || []) {
        if (location.sourceOnly) continue;
        const source = rows.find(row => row.index === location.rowIndex)?.ps[location.blockIndex]?.text;
        if (!source) continue;
        let needle = location.word;
        if (issue.kind === 'correction' && issue.choice === 'apply') needle = issue.replacement;
        let ranges = [];
        if (needle) {
            let positions = issue.kind === 'correction' ? replacementMatches(source, needle) : [...source.matchAll(new RegExp(needle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'))].map(match => match.index);
            if (location.sourceStart !== undefined) positions = positions.filter(position => position === location.sourceStart);
            ranges = positions.map(start => ({
                start,
                end: start + needle.length
            }));
        } else {
            const trimmed = source.trim();
            if (trimmed) ranges = [{
                start: source.indexOf(trimmed),
                end: source.indexOf(trimmed) + Math.min(trimmed.length, 180)
            }];
        }
        targets.push({
            anchor: 'p3h_' + location.rowIndex + '_' + location.blockIndex,
            text: source,
            ranges,
            tooltip: location.label + ' — ' + issue.title + '. ' + issue.message
        });
    }
    return viewer.highlightDraft(targets, scroll);
}
