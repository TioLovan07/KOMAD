/** Resolve the decision heading separately from the enactment text. */
import {norm} from './xml.js';
import {matrixKeepText,matrixCloneBlock as clone} from './matrix-reader.js';
const decision=p=>/^MEMUTUSKAN\s*:$/i.test(norm(p.text));
export function selectEnactment(left,right,{prefix}) {
 if(prefix!=='Menetapkan'&&!left.some(decision))return null;
 const original=left.filter(p=>!decision(p));
 if(!original.some(p=>/^PERATURAN\b/i.test(norm(p.text))))return null;
 if(right.some(p=>!decision(p)&&!matrixKeepText(p)))return null;
 return {ps:[...(right.filter(decision).length?right.filter(decision):left.filter(decision)),...original].map(clone),useRight:false,rule:'Isi penetapan dari awal dipertahankan; judul MEMUTUSKAN dipisahkan dari petunjuk Tetap.'};
}
