import assert from 'node:assert/strict';
import {missingSourceBlocks} from '../docs/js/source-coverage.js';
const old = [
 {text:'Batas wilayah sebagai berikut:'},
 {text:'b. batas sebelah utara : Desa Yangapi;',list:{prefixLength:3}},
 {text:'Batas sebelah selatan : Desa Bangbang dan Desa Peninjoan;'},
 {text:'Batas sebelah barat : Desa Peninjoan dan Desa Yangapi;'}
];
const merged = [{text:'Batas wilayah sebagai berikut: a. Batas sebelah timur: Desa X; b. Batas sebelah utara: Desa Yangapi; c. Batas sebelah selatan: Desa Bangbang dan Desa Peninjoan; d. Batas sebelah barat: Desa Peninjoan dan Desa Yangapi;'}];
assert.deepEqual(missingSourceBlocks(old,merged),[]);
assert.deepEqual(missingSourceBlocks(old,[{text:'Batas wilayah sebagai berikut:'}]),old.slice(1));
assert.deepEqual(missingSourceBlocks([{text:'Koordinat 8 24 50.288 LS'}],[{text:'Koordinat 8 24 50.289 LS'}]),[{text:'Koordinat 8 24 50.288 LS'}]);
assert.deepEqual(missingSourceBlocks([{text:'Batas Desa Peninjoan'}],[{text:'Batas Desa Peni'},{text:'njoan'}]),[]);
console.log('PASS: paragraph/table reflow does not create missing-content notes; genuinely absent text and changed coordinates remain detectable.');
