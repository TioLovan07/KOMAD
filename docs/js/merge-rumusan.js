/** Experimental, conservative overlay. The stable selector remains the fallback.
 * No instruction is emitted as draft text. Every unmatched replacement blocks export.
 */
import {norm,addIssue} from './xml.js';
import {matrixCloneBlock as clone,matrixKeepText as isKeep,matrixItemKey} from './matrix-reader.js';
const article=p=>norm(p.text).match(/^Pasal\s+(\d+[A-Za-z]?)$/i)?.[1];
const chapter=p=>/^BAB\s+(?:[IVXLCDM]+[A-Z]?|\d+)$/i.test(norm(p.text));
const structural=p=>!!article(p)||chapter(p)||/^Bagian\s+\S+$/i.test(norm(p.text));
const keep=p=>isKeep(p)||/^tetsp[.;]?$/i.test(norm(p.text));
const key=p=>p.list ? (p.list.kind||'')+':'+matrixItemKey(p) : null;
function clean(p) { return norm(p.text.slice(p.list?.prefixLength||0)).toLowerCase().replace(/[\p{P}\p{S}]+/gu,' ').replace(/\s+/g,' ').trim(); }
function sameLeadApartFromReferences(a,b) {
    const key=p=>clean(p).replace(/\bpasal\s+\d+[a-z]?(?:\s+ayat\s+\d+)?(?:\s+huruf\s+[a-z])?/g,'').replace(/\s+/g,' ').trim();
    return key(a)===key(b);
}
function similarity(a,b) {
    a=clean(a);b=clean(b);if(a===b)return 1;
    if(Math.min(a.length,b.length)<35)return 0;
    const aa=a.split(' '),bb=b.split(' '),counts=new Map();
    aa.forEach(t=>counts.set(t,(counts.get(t)||0)+1));let shared=0;
    for(const t of bb)if(counts.get(t)>0){shared++;counts.set(t,counts.get(t)-1);}
    const overlap=2*shared/(aa.length+bb.length);
    return a.slice(0,40)===b.slice(0,40)?overlap:overlap*.8;
}
function units(ps,start=0,end=ps.length) {
    const out=[];let current=null,ayat=false;
    for(let i=start;i<end;i++) {
        const p=ps[i];if(structural(p)){current=null;ayat=false;continue;}
        const child=p.list?.kind==='letter'&&ayat;
        if(!current || (p.list&&!child)) { current={start:i,end:i+1,p,key:key(p)};out.push(current); }
        else current.end=i+1;
        if(p.list?.kind==='ayat')ayat=true;
    }
    return out;
}
export function experimentalSelect(left,right,{m,rowIndex,prefix,note=''}) {
    if(!left.length||!right.length)return null;
    const titles=left.map((p,i)=>article(p)?i:-1).filter(i=>i>=0);
    const mixed=right.some(keep);
    const rightTitles=right.map((p,i)=>article(p)?i:-1).filter(i=>i>=0);
    const rawArticles=[...new Set(left.map(p=>p.sourceArticle).filter(Boolean))];
    const oldBody=titles.length===1?left.slice(titles[0]+1):titles.length===0&&rawArticles.length===1?left.filter(p=>p.sourceArticle===rawArticles[0]):[];
    const resultBody=rightTitles.length===1?right.slice(rightTitles[0]+1):[];
    // A struck source title is metadata, not permission to discard its surviving body.
    if(!titles.length && rawArticles.length===1 && rightTitles.length===1 && resultBody.some(p=>p.list) && !mixed) {
        const firstBody=left.findIndex(p=>p.sourceArticle===rawArticles[0]);
        const prefixBlocks=firstBody>=0?left.slice(0,firstBody):[];
        return experimentalSelect([...prefixBlocks,right[rightTitles[0]],...oldBody],right,{m,rowIndex,prefix,note});
    }
    if(rightTitles.length>=1 && rightTitles.at(-1)===right.length-1 && oldBody.length) {
        const deletion=/\b(?:dihapus|hapus|dipindahkan|reposisi)\b/i.test(note);
        addIssue(m,{kind:'review',rowIndex,hard:deletion,title:'Isi awal dipertahankan pada baris '+rowIndex,
            message:deletion?'Hasil hanya memuat judul, tetapi keterangan menyebut penghapusan/pemindahan. Perjelas sumber sebelum unduhan.':'Hasil hanya memuat judul Pasal. Isi awal dipertahankan; periksa bahwa tidak ada penghapusan yang dimaksud.'});
        return {ps:[...right.map(clone),...oldBody.map(clone)],useRight:true,rule:'Judul hasil digunakan; isi awal dipertahankan karena hasil tidak menulis isi pengganti.'};
    }
    const oldKeys=units(oldBody).map(u=>u.key).filter(Boolean),newKeys=units(resultBody).map(u=>u.key).filter(Boolean);
    const partialSingle=![...oldBody,...resultBody].some(p=>p.node.localName==='tbl')&&titles.length===1&&rightTitles.length===1&&!mixed&&newKeys.length>0&&newKeys.length<oldKeys.length&&newKeys.every(k=>oldKeys.includes(k));
    const startsAtBeginning=['1','a'].includes(matrixItemKey(right[0]));
    const explicitPartial=/(?:angka|huruf|butir)\s+[0-9a-z]+\s+(?:diubah|diperbaiki|disempurnakan)/i.test(note);
    const prefixPartial=/^(Mengingat|Menimbang)$/.test(prefix||'') && (!startsAtBeginning||explicitPartial) && right.length<left.length && right.every(p=>p.list&&!keep(p));
    if(!prefixPartial && !partialSingle && !(mixed&&(titles.length>1 || (titles.length===1 && right.some(p=>p.list&&!keep(p))))))return null;
    // Whole-article Tetap/renumbering remains owned by the stable rule.
    if(titles.length===1 && right.some(article) && right.filter(p=>!keep(p)&&!structural(p)).length===0)return null;
    const edits=new Map(),issues=[],used=new Set(),insertions=new Map();
    const issue=(title,message,hard=false)=>issues.push({kind:'review',rowIndex,title,message,hard});
    const boundaries=[];
    left.forEach((p,i)=>{if(article(p)){let end=i+1;while(end<left.length&&!structural(left[end]))end++;boundaries.push({num:article(p),start:i,end});}});
    let scope=null;
    const oldUnits=units(left);
    function put(start,end,ps) {
        for(const [at,e] of edits)if(start<e.end&&end>at)return false;
        edits.set(start,{end,ps});return true;
    }
    const ordinary=[];
    for(let i=0;i<right.length;i++) {
        const p=right[i];
        if(keep(p)) {
            if(/^tetsp/i.test(norm(p.text)))issue('Konfirmasi petunjuk “Tetsp” pada baris '+rowIndex,'Petunjuk mirip Tetap ditemukan. Rumusan awal dipertahankan sebagai kandidat; pastikan maksudnya sebelum menyetujui.');
            continue;
        }
        if(article(p)) {
            const matches=partialSingle?boundaries:boundaries.filter(s=>s.num===article(p));
            if(matches.length!==1){
                let end=i+1;while(end<right.length&&!structural(right[end]))end++;
                const body=right.slice(i+1,end);
                if(matches.length===0 && body.length && !body.some(keep)) {
                    const next=boundaries.find(s=>Number(s.num)>Number(article(p)));
                    const at=next?.start??left.length;
                    if(!insertions.has(at))insertions.set(at,[]);
                    insertions.get(at).push(...right.slice(i,end));
                    issue('Pasal baru dari hasil pada baris '+rowIndex,'Pasal '+article(p)+' beserta rumusannya hanya tersedia pada hasil dan ditambahkan sebelum pasal berikutnya. Periksa penempatannya.');
                    i=end-1;scope=null;continue;
                }
                issue('Sasaran Pasal '+article(p)+' belum pasti','Judul hasil tidak cocok dengan satu judul awal atau hanya berupa judul kosong. Rumusan awal dipertahankan sementara.',true);scope=null;continue;
            }
            scope=matches[0];put(scope.start,scope.start+1,[p]);continue;
        }
        if(chapter(p)||/^Bagian\s+\S+$/i.test(norm(p.text))) {
            const next=right[i+1];
            let candidates=left.map((q,j)=>({q,j})).filter(({q,j})=>(chapter(p)?chapter(q):/^Bagian\s+\S+$/i.test(norm(q.text))) && next && left[j+1] && clean(left[j+1])===clean(next));
            if(!candidates.length) {
                const following=right.slice(i+1).find(article);
                candidates=left.map((q,j)=>({q,j})).filter(({q,j})=>norm(q.text).toLowerCase()===norm(p.text).toLowerCase() && following && article(left.slice(j+1).find(article)||{text:''})===article(following));
            }
            if(candidates.length===1){put(candidates[0].j,candidates[0].j+2,[p,next]);i++;}
            else issue('Judul struktur belum dapat dicocokkan pada baris '+rowIndex,'Judul BAB/Bagian hasil tidak mempunyai pasangan tunggal pada awal. Periksa matriks.',true);
            scope=null;continue;
        }
        let end=i+1;
        while(end<right.length&&!structural(right[end])&&!keep(right[end])&&!right[end].list)end++;
        // Collect child letters of an ayat in the same replacement unit.
        if(p.list?.kind==='ayat')while(end<right.length && right[end].list?.kind==='letter' && !keep(right[end])) {
            end++;while(end<right.length&&!structural(right[end])&&!keep(right[end])&&!right[end].list)end++;
        }
        ordinary.push({p,ps:right.slice(i,end),scope});i=end-1;
    }
    for(const replacement of ordinary) {
        const {p,ps,scope}=replacement;
        const candidates=oldUnits.filter(u=>!used.has(u.start) && (!scope || (u.start>scope.start&&u.start<scope.end)) && (!p.list || u.key===key(p)));
        let match;
        if(candidates.length===1 && p.list)match=candidates[0];
        else {
            const ranked=candidates.map(u=>({u,score:similarity(u.p,p)})).sort((a,b)=>b.score-a.score);
            if(ranked[0]?.score>=.88 && (!ranked[1]||ranked[0].score-ranked[1].score>=.15))match=ranked[0].u;
        }
        let replacementBlocks=ps;
        if(match && p.list?.kind==='ayat' && (similarity(match.p,p)>=.88||sameLeadApartFromReferences(match.p,p))) {
            const oldChildren=units(left,match.start+1,match.end),newChildren=units(ps,1,ps.length);
            if(oldChildren.length && newChildren.length<oldChildren.length && newChildren.every(c=>c.key&&oldChildren.filter(o=>o.key===c.key).length===1)) {
                const byKey=new Map(newChildren.map(c=>[c.key,ps.slice(c.start,c.end)]));
                replacementBlocks=[p,...oldChildren.flatMap(c=>byKey.get(c.key)||left.slice(c.start,c.end))];
            }
        }
        if(!match || !put(match.start,match.end,replacementBlocks)) {
            issue('Revisi belum dapat dipetakan pada baris '+rowIndex,'Penggalan “'+norm(p.text).slice(0,100)+'” tidak mempunyai pasangan tunggal. Rumusan awal ditampilkan sementara; perbaiki sumber untuk bagian ini.',true);continue;
        }
        used.add(match.start);
    }
    if(!edits.size&&!issues.length)return {ps:left.map(clone),useRight:false,rule:'Semua petunjuk Tetap; isi awal yang telah dibersihkan dari coretan digunakan.'};
    const result=[];
    for(let i=0;i<left.length;) {
        if(insertions.has(i))result.push(...insertions.get(i).map(clone));
        const edit=edits.get(i);if(edit){result.push(...edit.ps.map(clone));i=edit.end;}else result.push(clone(left[i++]));
    }
    if(insertions.has(left.length))result.push(...insertions.get(left.length).map(clone));
    for(const i of issues)addIssue(m,i);
    if(edits.size) {
        const targets=[...edits.values()].flatMap(e=>e.ps).filter(p=>!keep(p)).map(p=>norm(p.text)).filter(Boolean);
        const preview=targets.slice(0,3).map(t=>'“'+t.slice(0,110)+'”').join('; ');
        issue('Penggabungan bagian perlu diperiksa pada baris '+rowIndex,edits.size+' bagian dipetakan dari kolom hasil. Contoh bagian yang digunakan: '+preview+'. Isi awal di luar bagian itu tetap dipertahankan. Pastikan bagian yang tidak ditulis ulang masih berlaku.');
        issues.at(-1).action='Bandingkan bagian hasil yang dikutip dengan draft pada baris '+rowIndex+'. Periksa juga paragraf/butir awal yang masih dipertahankan, terutama sesudah petunjuk Tetap. Konfirmasikan hanya jika penggantian dan bagian yang dipertahankan sesuai keputusan harmonisasi.';
    }
    // Add the informational merge note only once, after hard findings.
    if(edits.size)addIssue(m,issues.at(-1));
    return {ps:result,useRight:true,rule:'Eksperimen: '+edits.size+' penggantian terpetakan; isi lain dipertahankan dari awal.'};
}
