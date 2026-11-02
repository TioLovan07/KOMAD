/** Exact standalone matrix instructions; legal sentences containing tetap are content. */
export function keepInstruction(value) {
 const raw=typeof value==='string'?value:value?.text;
 const text=String(raw||'').normalize('NFKC').replace(/[\u200B\uFEFF]/g,'').replace(/\s+/g,' ').trim();
 const m=text.match(/^(?:(?:\(\d+\)|\d+[.)]|[a-z][.)])\s*)?(tetap|tetsp)[.;]?$/i);
 return m?{kind:m[1].toLowerCase()==='tetap'?'keep':'confirm-keep'}:null;
}
