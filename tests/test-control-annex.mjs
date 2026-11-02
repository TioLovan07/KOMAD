import {documentMetadata} from '../docs/js/document-layout.js';
import assert from 'node:assert/strict';
import {keepInstruction} from '../docs/js/control-text.js';
import {annexTitles,isAnnexInput} from '../docs/js/annex-titles.js';
for(const s of [' Tetap','\t(2) Tetap','\u00a0Tetap  ','1. Tetap','Tetsp'])assert.ok(keepInstruction(s),s);
for(const s of ['peraturan ini tetap berlaku','Tetapkan','Tetap berlaku'])assert.equal(keepInstruction(s),null,s);
const p=text=>({text});const meta={regulation:'PERATURAN BUPATI X',number:'…',year:'…',title:'PERATURAN CONTOH'};
const source=[p('LAMPIRAN I'),p('PERATURAN BUPATI X'),p('NOMOR … TAHUN …'),p('TENTANG'),p('JUDUL SUMBER'),p('ISI RAHASIA TABEL'),p('LAMPIRAN II'),p('TENTANG'),p('JUDUL KEDUA'),p('ISI KEDUA')];
const titles=annexTitles(source,[],meta,'Lampiran');assert.equal(titles.length,2);assert.equal(titles[0][4],'PERATURAN CONTOH');assert.equal(titles[1][0],'LAMPIRAN II');assert.ok(!JSON.stringify(titles).includes('ISI RAHASIA'));assert.ok(isAnnexInput('Lampiran',source,[]));
console.log('PASS: standalone Tetap variants, legal sentences, two title-only annexes.');

assert.equal(annexTitles([p('LAMPIRAN'),p('TENTANG'),p('PERUBAHAN KEDUA ATAS PERATURAN')],[],{...meta,title:'PERUBAHAN KEDUA ATAS PERATURAN BUPATI NOMOR 37 TAHUN 2023 TENTANG KEBIJAKAN AKUNTANSI'},'Lampiran')[0][4],'PERUBAHAN KEDUA ATAS PERATURAN BUPATI NOMOR 37 TAHUN 2023 TENTANG KEBIJAKAN AKUNTANSI');

const completeTitle='PERUBAHAN KEDUA ATAS PERATURAN BUPATI NOMOR 37 TAHUN 2023 TENTANG KEBIJAKAN AKUNTANSI PEMERINTAH DAERAH';
assert.equal(documentMetadata([{label:'Judul',ps:['RANCANGAN','PERATURAN BUPATI BANGLI','NOMOR ... TAHUN ..','TENTANG',completeTitle,'DENGAN RAHMAT TUHAN YANG MAHA ESA','BUPATI BANGLI,'].map(text=>({text}))}]).title,completeTitle);
const annualTitle='STANDAR HARGA SATUAN PEMERINTAH DAERAH TAHUN ANGGARAN 2028';
const annualMeta=documentMetadata([{label:'Judul',ps:['PERATURAN BUPATI CONTOH','NOMOR ... TAHUN ...','TENTANG','STANDAR HARGA SATUAN PEMERINTAH DAERAH','TAHUN ANGGARAN 2028','DENGAN RAHMAT TUHAN YANG MAHA ESA'].map(p)}]);
assert.equal(annualMeta.title,annualTitle);
assert.equal(annexTitles([p('LAMPIRAN I'),p('TENTANG'),p('STANDAR HARGA SATUAN PEMERINTAH DAERAH'),p('TAHUN ANGGARAN 2028')],[],annualMeta)[0][4],annualTitle);
