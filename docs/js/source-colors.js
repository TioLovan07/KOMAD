/** Materialize source colour inheritance before matrix blocks are detached. */
import {W,nodes,direct,val,we} from './xml.js';
const pick=(p,n)=>p&&direct(p,n)[0];
export function materializeSourceColors(doc,styles,theme) {
 const byId=new Map((styles?nodes(styles,'style'):[]).map(s=>[s.getAttributeNS(W,'styleId'),s]));
 const chain=(id,seen=new Set())=>{if(!id||seen.has(id)||!byId.has(id))return [];seen.add(id);const s=byId.get(id);return [...chain(val(pick(s,'basedOn')),seen),s];};
 const defaults=styles&&nodes(styles,'docDefaults')[0];
 const defaultRun=defaults&&nodes(defaults,'rPr')[0];
 const themeColors=new Map();
 if(theme)for(const scheme of [...theme.getElementsByTagName('*')].filter(n=>n.localName==='clrScheme'))for(const e of scheme.children){const c=e.firstElementChild;if(c)themeColors.set(e.localName,c.getAttribute('val')||c.getAttribute('lastClr'));}
 const resolve=(n)=>{
  const c=n.cloneNode(true),aliases={dark1:'dk1',light1:'lt1',dark2:'dk2',light2:'lt2',text1:'dk1',background1:'lt1',text2:'dk2',background2:'lt2'};
  for(const [key,tint,shade,target] of [['themeColor','themeTint','themeShade',n.localName==='shd'?'color':'val'],['themeFill','themeFillTint','themeFillShade','fill']]) {
   const name=n.getAttributeNS(W,key);let rgb=themeColors.get(aliases[name]||name);
   if(!/^[\da-f]{6}$/i.test(rgb||''))continue;
   const factor=n.getAttributeNS(W,tint),dark=n.getAttributeNS(W,shade);
   rgb=rgb.match(/../g).map(x=>{let v=parseInt(x,16);if(dark)v*=parseInt(dark,16)/255;if(factor)v+=(255-v)*parseInt(factor,16)/255;return Math.round(v).toString(16).padStart(2,'0');}).join('').toUpperCase();
   c.setAttributeNS(W,'w:'+target,rgb);for(const a of [key,tint,shade])c.removeAttributeNS(W,a);
  }
  return c;
 };
 for(const p of nodes(doc,'p')) {
  const pp=pick(p,'pPr'),styleId=val(pick(pp,'pStyle'))||[...byId.entries()].find(([,s])=>s.getAttributeNS(W,'type')==='paragraph'&&s.getAttributeNS(W,'default')==='1')?.[0];
  const paragraphStyles=chain(styleId);
  let cell=p.parentElement;while(cell&&cell.localName!=='tc')cell=cell.parentElement;
  const shades=[pick(pick(cell,'tcPr'),'shd'),...paragraphStyles.map(s=>pick(pick(s,'pPr'),'shd')),pick(pp,'shd')].filter(Boolean);
  if(shades.length){const pr=pp||we(doc,'pPr');if(!pp)p.insertBefore(pr,p.firstChild);direct(pr,'shd').forEach(n=>n.remove());pr.appendChild(resolve(shades.at(-1)));}
  for(const r of nodes(p,'r')) {
   const rp=pick(r,'rPr'),props=[defaultRun,...paragraphStyles.map(s=>pick(s,'rPr')),pick(pp,'rPr'),...chain(val(pick(rp,'rStyle'))).map(s=>pick(s,'rPr')),rp].filter(Boolean);
   const chosen=new Map();for(const pr of props)for(const n of direct(pr,'color').concat(direct(pr,'highlight'),direct(pr,'shd')))chosen.set(n.localName,n);
   if(chosen.size){const pr=rp||we(doc,'rPr');if(!rp)r.insertBefore(pr,r.firstChild);for(const [name,n] of chosen){const c=resolve(n);direct(pr,name).forEach(x=>x.remove());pr.appendChild(c);}}
  }
 }
}
