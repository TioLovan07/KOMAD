/** Title-only annexes: never copy source annex body, tables, pictures or landscape sections. */
import {annexHeadingLines} from './document-layout.js';
const clean=s=>String(s||'').replace(/\s+/g,' ').trim();
const marker=s=>clean(s).match(/^LAMPIRAN(?:\s+([IVXLCDM]+|\d+))?(?=\s|$)/i);
export function isAnnexInput(label,left,right) {
 if (/^(?:\d+[.)]?\s*)?lampiran\b/i.test(clean(label))) return true;
 const first=[left,right].map(ps=>ps.find(p=>clean(p.text))).filter(Boolean);
 return first.length>0 && first.every(p=>marker(p.text));
}
export function annexTitles(left,right,metadata,label='') {
 const source=right.some(p=>marker(p.text))?right:left;
 const starts=source.map((p,i)=>marker(p.text)?i:-1).filter(i=>i>=0);
 if(!starts.length) {
  const lines=annexHeadingLines(metadata),m=clean(label).match(/lampiran\s+([IVXLCDM]+|\d+)\b/i);
  if(m)lines[0]='LAMPIRAN '+m[1].toUpperCase();
  return [lines];
 }
 const seen=new Set(),out=[];
 for(const [n,at] of starts.entries()) {
  const part=source.slice(at,starts[n+1]??source.length),m=marker(part[0].text),key=m[1]?.toUpperCase()||'';
  if(seen.has(key))continue;seen.add(key);
  const lines=annexHeadingLines(metadata);lines[0]='LAMPIRAN'+(key?' '+key:'');
  // The regulation title is reconstructed from the complete opening, never a clipped annex line.
  // Preserve a source-specific title only when the opening has no usable title.
  if(!metadata.title || metadata.title==='…') {
   const about=part.slice(0,8).findIndex(p=>/\bTENTANG\b/i.test(clean(p.text)));
   if(about>=0) {
    const title=[];const tail=clean(part[about].text).split(/\bTENTANG\b/i).slice(1).join(' ').trim();
    if(tail)title.push(tail);
    for(const p of part.slice(about+1)) {
     const t=clean(p.text);if(!t)continue;
     if(p.node?.localName==='tbl'||/^(?:BUPATI|WALI\s*KOTA|GUBERNUR)\b/.test(t)||/^\.{2,}$|^…$/.test(t)||/^\(?\d+[.)]\s/.test(t))break;
     title.push(t);
    }
    if(title.length)lines[4]=title.join(' ').toUpperCase();
   }
  }
  out.push(lines);
 }
 return out.length?out:[annexHeadingLines(metadata)];
}
