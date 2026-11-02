/** Infer hierarchy without renumbering or changing source words. */
import {norm,addIssue} from './xml.js';
const structural=t=>/^(?:Pasal\s+\d+[A-Za-z]?|BAB\s+(?:[IVXLCDM]+|\d+)|Bagian\s+\S+|Paragraf\s+\d+)$/i.test(t);
const ordinal=p=>{const s=p.list.marker.replace(/[().\s]/g,'');return p.list.kind==='letter'?s.toLowerCase().charCodeAt(0)-96:Number(s);};
export function annotateLists(m) {
 let stack=[],definition=false,section='',previous='',sequences=new Map(),instructionLast=null;
 for(const row of m.rows) for(let i=0;i<row.ps.length;i++) {
  const p=row.ps[i],t=norm(p.text);
  if(p.inlineStructure)continue;
  if(structural(t)){stack=[];definition=false;sequences=new Map();section=t;previous=t;continue;}
  if(/yang dimaksud dengan\s*:\s*$/i.test(t)){definition=true;stack=[];sequences=new Map();previous=t;continue;}
  if(row.prefix){stack=[];definition=false;continue;}
  if(p.amendmentInstruction){stack=[];definition=false;}
  if(p.list) {
   const kind=p.list.kind;
   let level=0;
   if(kind==='letter') {
    const parent=[...stack].reverse().find(x=>x.kind==='ayat'||(definition&&x.kind==='number'&&x.level===0));
    level=parent?parent.level+1:0;
   } else if(kind==='number') {
    const letter=[...stack].reverse().find(x=>x.kind==='letter');
    if(letter && ((p.sourceIndent??0)>(letter.indent??Infinity)+60 || p.list.level>letter.nativeLevel || /:\s*$/.test(previous))) level=letter.level+1;
    else if(stack.at(-1)?.kind==='number'&&stack.at(-1).level>0)level=stack.at(-1).level;
   }
   // Legal ayat always starts a fresh root; no automatic repairs to its marker.
   if(kind==='ayat')level=0;
   const parent=stack.filter(x=>x.level<level).map(x=>x.marker).join('/');
   const group=section+'|'+parent+'|'+level+'|'+kind;
   const n=ordinal(p),record=p.amendmentInstruction?instructionLast:sequences.get(group),last=record?.value;
   const display=value=>kind==='letter'?String.fromCharCode(96+value)+'.':kind==='ayat'?'('+value+')':value+'.';
   if(Number.isFinite(n) && ((last!=null&&n!==last+1)||(last==null&&n!==1))) {
    const expected=display(last==null?1:last+1);
    const finding=last==null?'Penanda pertama yang terlihat adalah '+p.list.marker.replace(/\.$/,'')+'. Penanda '+expected+' belum terlihat sebelumnya dalam kelompok ini.':n>last+1?'Setelah '+record.marker+' langsung muncul '+p.list.marker.replace(/\.$/,'')+'. Penanda '+(n-last===2?expected:display(last+1)+' sampai '+display(n-1))+' tidak terlihat di antaranya pada rumusan terpilih.':n===last?'Penanda '+p.list.marker+' muncul lagi pada kelompok yang sama.':'Urutan mundur dari '+record.marker+' ke '+p.list.marker+' pada kelompok yang sama.';
    const change=last==null?'awal kelompok':n>last+1?'melompat':n===last?'berulang':'mundur';
    addIssue(m,{kind:'review',checkType:'sequence',row,blockIndex:i,title:'Penomoran '+change+': '+(record?record.marker+' → ':'')+p.list.marker+' pada '+(section||row.label),message:finding,
      action:last==null?'Periksa apakah butir ini lanjutan kelompok sebelumnya, anak daftar yang salah tingkat, atau seharusnya dimulai dari '+expected.replace(/\.$/,'')+'. Jika sumber perlu diperbaiki, ubah penanda/tingkatnya dan unggah ulang. Jika memang disengaja, pilih Sudah Saya Periksa.':'Bandingkan penanda sebelum dan sesudahnya pada matriks. Pastikan apakah ada butir dicoret/dihapus, awal kelompok baru, atau nomor yang salah. Jika harus berurutan, perbaiki sumber dan unggah ulang; jika urutan ini memang disepakati, pilih Sudah Saya Periksa. Nomor tidak diubah otomatis.',
      numbering:{previous:record?.marker,current:p.list.marker,expected,change}});
   }
   const current={value:n,marker:p.list.marker,rowIndex:row.index,blockIndex:i};
   if(p.amendmentInstruction)instructionLast=current;else sequences.set(group,current);
   const indentParent=stack.filter(x=>x.level<level).at(-1);
   if(indentParent && Number.isFinite(p.sourceIndent) && Number.isFinite(indentParent.indent) && p.sourceIndent<=indentParent.indent+60) {
    addIssue(m,{kind:'review',checkType:'indent',row,blockIndex:i,title:'Periksa indent butir '+p.list.marker+' pada '+(section||row.label),message:'Penanda dibaca sebagai anak daftar, tetapi indent sumber tidak lebih masuk daripada induknya. Draft memakai tingkat anak daftar; pastikan hubungan induk dan rinciannya sebelum menyetujui.',action:'Periksa apakah butir ini rincian dari butir induk atau butir sejajar. Jika rincian, pastikan indent lebih masuk; jika sejajar, perbaiki tingkat daftar pada sumber lalu unggah ulang. Jika hasil sudah benar, pilih Sudah Saya Periksa.'});
   }
   // Starting a new parent invalidates child sequences, even when a child uses the same marker.
   stack=stack.filter(x=>x.level<level);
   stack.push({kind,level,marker:p.list.marker,indent:p.sourceIndent,nativeLevel:p.list.level||0});
   p.layoutLevel=level;
  } else if(stack.length && t && !/^(?:Ditetapkan|Diundangkan|BERITA|LEMBARAN|LAMPIRAN)\b/i.test(t) && p.node.localName==='p') {
   const matches=stack.filter(x=>Number.isFinite(x.indent)&&Number.isFinite(p.sourceIndent));
   if(matches.length) {
    const ranked=matches.map(x=>({x,d:Math.abs(x.indent-p.sourceIndent)})).sort((a,b)=>a.d-b.d);
    // Only use a nearby source text indent, never infer membership solely from proximity.
    if(ranked[0].d<=160)p.continuationLevel=ranked[0].x.level;
    else if(p.sourceIndent>stack.at(-1).indent)p.continuationLevel=stack.at(-1).level;
   }
  }
  previous=t;
 }
}
