import assert from 'node:assert/strict';
import {
    attachIssueLocations
} from '../docs/js/issue-locations.js';

// Dua penanda identik harus tetap menunjuk paragraf yang berbeda.
const row = {
    index: 8,
    label: 'Pasal 4',
    ps: [{
        text: 'Pasal 4'
    }, {
        text: 'Ruang lingkup Paraturan ini meliputi:'
    }, {
        text: 'd. partisipasi;',
        list: {
            marker: 'd.'
        }
    }, {
        text: 'd. pembiayaan.',
        list: {
            marker: 'd.'
        }
    }, {
        text: 'Ditetapkan di … pada tanggal …'
    }, ],
    left: [],
    right: []
};
const issues = [{
    kind: 'review',
    marker: 'd.',
    section: 'Pasal 4',
    row
}, {
    kind: 'correction',
    word: 'Paraturan'
}, {
    kind: 'advisory'
}, {
    kind: 'review',
    rowIndex: 9,
    message: 'Tetap tidak dapat dicocokkan.'
}, {
    kind: 'review',
    title: 'Catatan tingkat dokumen',
    message: 'Periksa matriks.'
}];
attachIssueLocations({
    rows: [row],
    issues
});
assert.deepEqual(issues[0].locations.map(x => x.blockIndex), [2, 3]);
assert.equal(issues[1].locations[0].blockIndex, 1);
assert.equal(issues[1].locations[0].word, 'Paraturan');
assert.equal(issues[2].locations.length, 2);
assert.notEqual(issues[2].locations[0].sourceStart, issues[2].locations[1].sourceStart);
assert(issues[3].locations[0].sourceOnly);
assert(issues[4].locations[0].sourceOnly);
assert(issues.every(x => x.locations.length));
console.log('PASS: duplicate markers, correction location, separate placeholders, source-only and document notes.');
