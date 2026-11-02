/** A bounded plan for added chapters with consequent article renumbering.
 * Only used when unique chapter titles establish at least two shifted matches.
 * Every plan is reviewable. Ambiguous targets remain hard findings.
 */
import {norm,addIssue} from './xml.js';
import {matrixCloneBlock as clone,matrixKeepText as keep,matrixSelectRumusan} from './matrix-reader.js';
const article=p=>norm(p.text).match(/^Pasal\s+(\d+[A-Za-z]?)$/i)?.[1];
const chapter=p=>/^BAB\s+(?:[IVXLCDM]+|\d+)$/i.test(norm(p.text));
const heading=p=>article(p)||chapter(p)||/^Bagian\s+\S+$/i.test(norm(p.text));
const clean=p=>norm(p.text.slice(p.list?.prefixLength||0)).toLowerCase().replace(/\bpasal\s+\d+[a-z]?(?:\s+ayat\s*\(?\d+\)?)?(?:\s+huruf\s+[a-z])?/g,'').replace(/\bayat\s*\(?\d+\)?/g,'').replace(/[^\p{L}\p{N}]+/gu,' ').replace(/\bsebagaimana dimaksud (?:dalam|pada)\b/g,'sebagaimana dimaksud').trim();
function parts(ps) {
 const out=[];let title='';
 for(let i=0;i<ps.length;i++) {
  if(chapter(ps[i]))title=clean(ps[i+1]||{text:''});
  if(article(ps[i])) {let end=i+1;while(end<ps.length&&!heading(ps[end]))end++;out.push({num:article(ps[i]),title,start:i,end,body:ps.slice(i+1,end),head:ps[i]});}
 }
 return out;
}
function leadScore(a,b) {
 const x=clean(a),y=clean(b);if(!x||!y)return 0;if(x===y)return 1;
 const xx=new Set(x.split(' ')),yy=new Set(y.split(' '));let n=0;for(const t of xx)if(yy.has(t))n++;
 return 2*n/(xx.size+yy.size);
}
export function selectStructurePlan(left,right,options,selectPartial) {
 const old=parts(left),now=parts(right);
 if(old.length<4||now.length<old.length||!right.some(chapter)||new Set(now.map(s=>s.num)).size!==now.length)return null;
 if(now.some((s,i)=>i&&Number(s.num)<=Number(now[i-1].num)))return null;
 const newChapters=[...new Set(now.map(s=>s.title))].filter(t=>t&&!old.some(s=>s.title===t));
 if(!newChapters.length)return null;
 const matches=new Map(),used=new Set();
 // First anchor unchanged chapter titles with equal article counts.
 for(const title of new Set(now.map(s=>s.title))) {
  const aa=old.filter(s=>s.title===title),bb=now.filter(s=>s.title===title);
  if(aa.length&&aa.length===bb.length)bb.forEach((b,i)=>{matches.set(b,aa[i]);used.add(aa[i]);});
 }
 if([...matches].filter(([b,a])=>b.num!==a.num).length<2)return null;
 // Within changed chapters, use long, distinctive lead text before numeric equality.
 for(const b of now.filter(s=>!matches.has(s))) {
  const lead=b.body.find(p=>!keep(p)&&!/^tetsp[.;]?$/i.test(norm(p.text)));
  const ranked=old.filter(a=>!used.has(a)&&a.title===b.title).map(a=>({a,score:lead&&a.body[0]?leadScore(a.body[0],lead):0})).sort((x,y)=>y.score-x.score);
  if(clean(lead||{text:''}).length>=35&&ranked[0]?.score>=.88&&(!ranked[1]||ranked[0].score-ranked[1].score>=.15)) {matches.set(b,ranked[0].a);used.add(ranked[0].a);continue;}
  const same=old.filter(a=>!used.has(a)&&a.title===b.title&&a.num===b.num);
  if(same.length===1){matches.set(b,same[0]);used.add(same[0]);}
 }
 const selections=new Map(),findings=[],mapping=[];
 for(const b of now) {
  if(b.body.some(p=>/^tetsp[.;]?$/i.test(norm(p.text))))addIssue(options.m,{kind:'review',rowIndex:options.rowIndex,title:'Konfirmasi petunjuk Tetsp pada Pasal '+b.num,message:'Rumusan awal dipertahankan sebagai kandidat Tetap. Pastikan petunjuk ini memang berarti Tetap.',excerpt:'Tetsp'});
  const a=matches.get(b),meaningful=b.body.filter(p=>!keep(p)&&!/^tetsp[.;]?$/i.test(norm(p.text)));
  if(!a) {
   if(meaningful.length&&meaningful.length===b.body.length) {selections.set(b,right.slice(b.start,b.end));mapping.push('Pasal '+b.num+' ditambahkan');}
   else findings.push({number:b.num,title:'Sasaran Pasal '+b.num+' belum pasti',message:'Pasal hasil belum mempunyai pasangan awal atau rumusan tambahan yang lengkap.',hard:true});
   continue;
  }
  mapping.push('Pasal '+a.num+' awal → Pasal '+b.num+' hasil');
  if(!meaningful.length) {if(!a.body.length)findings.push({number:b.num,title:'Isi Pasal '+b.num+' belum tersedia',message:'Judul hasil tidak mempunyai isi awal yang dapat dipertahankan.',hard:true});selections.set(b,[b.head,...a.body]);continue;}
  if(meaningful.length===1&&!meaningful[0].list&&a.body[0]&&clean(meaningful[0])===clean(a.body[0])) {selections.set(b,[b.head,meaningful[0],...a.body.slice(1)]);continue;}
  // Do not copy units that have explicitly moved into another complete result article.
  const source=[];
  for(let i=0;i<a.body.length;) {
   const p=a.body[i];let end=i+1;
   if(p.list?.kind==='ayat')while(end<a.body.length&&a.body[end].list?.kind!=='ayat')end++;
   const moved=p.list?.kind==='ayat'&&now.some(other=>other!==b&&other.body.some(q=>q.list?.kind==='ayat'&&clean(q).length>=35&&clean(q)===clean(p)));
   if(!moved)source.push(...a.body.slice(i,end));
   i=end;
  }
  const sourceBlocks=[b.head,...source],resultBlocks=[b.head,...b.body];
  const selected=selectPartial(sourceBlocks,resultBlocks,options)||matrixSelectRumusan(sourceBlocks,resultBlocks,options);
  selections.set(b,selected.ps);
 }
 for(const a of old.filter(a=>!used.has(a)))findings.push({number:a.num,title:'Rumusan awal Pasal '+a.num+' belum terpetakan',message:'Ada isi awal yang tidak mempunyai tujuan tunggal dalam rencana struktur. Perjelas penghapusan atau pemindahannya; unduhan tetap ditahan.',hard:true});
 if(findings.length)for(const f of findings)addIssue(options.m,{kind:'review',rowIndex:options.rowIndex,...f});
 const starts=new Map(now.map(s=>[s.start,s])),out=[];
 for(let i=0;i<right.length;) {const s=starts.get(i);if(s){out.push(...(selections.get(s)||[s.head]));i=s.end;}else{if(!keep(right[i]))out.push(right[i]);i++;}}
 addIssue(options.m,{kind:'review',rowIndex:options.rowIndex,title:'Periksa tambahan dan perubahan nomor pasal',message:mapping.join('; ')+'. Pemetaan memakai judul BAB, urutan, dan rumusan. Periksa khusus bagian yang hanya berisi judul atau Tetap sebelum menyetujui.',excerpt:'Tambahan dan pergeseran pasal pada kolom hasil'});
 return {ps:out.map(clone),useRight:true,rule:'Rencana struktur: tambahan BAB dan perubahan nomor dipetakan; konfirmasi perancang diperlukan.',structurePlan:true};
}
