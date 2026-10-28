/** Reserve 5.5 cm in the FIRST body block, above the first text paragraph.
 * With 2.5 cm section margins, text starts at 8 cm; later pages stay at 2.5 cm.
 * A body spacer is used because empty first-page headers can be suppressed.
 * The floating Garuda keeps its independent page-relative 3.5 cm position.
 */
export async function configureFirstPage(doc, archive) {
    const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const parser=new DOMParser(), serializer=new XMLSerializer();
    const body=doc.getElementsByTagNameNS(W,'body')[0];
    const make=(name,attrs={})=>{
        const node=doc.createElementNS(W,'w:'+name);
        for(const [key,value] of Object.entries(attrs)) node.setAttributeNS(W,'w:'+key,String(value));
        return node;
    };
    const direct=(node,name)=>[...node.children].filter(c=>c.namespaceURI===W && c.localName===name);
    const set=(node,name,attrs)=>{
        direct(node,name).forEach(c=>c.remove());
        const child=make(name,attrs);node.append(child);return child;
    };
    const orderPr=pr=>{
        const names=['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'];
        const rank=node=>names.includes(node.localName)?names.indexOf(node.localName):100;
        for(const child of [...pr.children].sort((a,b)=>rank(a)-rank(b))) pr.append(child);
    };
    const text=node=>[...node.getElementsByTagNameNS(W,'t')].map(t=>t.textContent).join('').trim();
    // Remove only the obsolete KOMAD-generated header and its references.
    const relFile=archive.file('word/_rels/document.xml.rels');
    const oldIds=new Set();
    if(relFile) {
        const rels=parser.parseFromString(await relFile.async('string'),'application/xml');
        for(const rel of [...rels.documentElement.children]) {
            if(rel.getAttribute('Target')==='komad-first-page.xml') {
                oldIds.add(rel.getAttribute('Id'));rel.remove();
            }
        }
        archive.file('word/_rels/document.xml.rels',serializer.serializeToString(rels));
    }
    for(const section of [...doc.getElementsByTagNameNS(W,'sectPr')]) {
        const obsolete=direct(section,'headerReference').filter(ref=>oldIds.has(ref.getAttributeNS(R,'id')));
        if(obsolete.length) {
            obsolete.forEach(ref=>ref.remove());
            direct(section,'titlePg').forEach(node=>node.remove());
        }
    }
    archive.remove('word/komad-first-page.xml');
    const types=parser.parseFromString(await archive.file('[Content_Types].xml').async('string'),'application/xml');
    for(const part of [...types.documentElement.children])
        if(part.getAttribute('PartName')==='/word/komad-first-page.xml') part.remove();
    archive.file('[Content_Types].xml',serializer.serializeToString(types));

    const firstBlock=[...body.children].find(node=>['p','tbl'].includes(node.localName));
    if(!firstBlock) return;
    // Reuse the existing picture-only paragraph, without adding another line.
    const spacer=firstBlock.localName==='p' && !text(firstBlock) ? firstBlock : make('p');
    if(spacer!==firstBlock) body.insertBefore(spacer,firstBlock);
    let pr=direct(spacer,'pPr')[0];
    if(!pr) {pr=make('pPr');spacer.insertBefore(pr,spacer.firstChild);}
    for(const name of ['keepNext','keepLines','pageBreakBefore','sectPr','numPr']) direct(pr,name).forEach(node=>node.remove());
    set(pr,'spacing',{before:0,after:0,line:3118,lineRule:'exact'});
    set(pr,'ind',{left:0,right:0,firstLine:0});
    orderPr(pr);
    // Stable marker lets the regression checks identify the spacer paragraph.
    if(![...spacer.getElementsByTagNameNS(W,'bookmarkStart')].some(b=>b.getAttributeNS(W,'name')==='komad_first_page_spacer')) {
        const bookmark=make('bookmarkStart',{id:2147000000,name:'komad_first_page_spacer'});
        spacer.insertBefore(bookmark,pr.nextSibling);
        spacer.append(make('bookmarkEnd',{id:2147000000}));
    }
    if(!spacer.getElementsByTagNameNS(W,'r').length) {
        const run=make('r'),blank=make('t');
        blank.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');
        blank.textContent=' ';run.append(blank);spacer.append(run);
    }
    const firstText=[...body.getElementsByTagNameNS(W,'p')].find(p=>p!==spacer && text(p));
    if(firstText) {
        let firstPr=direct(firstText,'pPr')[0];
        if(!firstPr) {firstPr=make('pPr');firstText.insertBefore(firstPr,firstText.firstChild);}
        const spacing=direct(firstPr,'spacing')[0] || set(firstPr,'spacing',{});
        spacing.setAttributeNS(W,'w:before','0');
    }
}
export function positionGaruda(paragraph) {
    const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
    const doc=paragraph.ownerDocument;
    const make=(name,attrs={})=>{
        const node=doc.createElementNS(WP,'wp:'+name);
        for(const [key,value] of Object.entries(attrs)) node.setAttribute(key,String(value));
        return node;
    };
    for(const inline of [...paragraph.getElementsByTagNameNS(WP,'inline')]) {
        const extent=inline.getElementsByTagNameNS(WP,'extent')[0];
        if(!extent) continue;
        const oldHeight=Number(extent.getAttribute('cy'));
        const height=1260000,width=Math.round(Number(extent.getAttribute('cx'))*height/oldHeight);
        extent.setAttribute('cx',width);extent.setAttribute('cy',height);
        const anchor=make('anchor',{distT:0,distB:0,distL:0,distR:0,simplePos:0,relativeHeight:0,behindDoc:0,locked:0,layoutInCell:1,allowOverlap:1});
        anchor.append(make('simplePos',{x:0,y:0}));
        const horizontal=make('positionH',{relativeFrom:'page'}),align=make('align');align.textContent='center';horizontal.append(align);
        const vertical=make('positionV',{relativeFrom:'page'}),offset=make('posOffset');offset.textContent='1260000';vertical.append(offset);
        anchor.append(horizontal,vertical,extent,make('wrapNone'));
        for(const child of [...inline.children]) if(['effectExtent','docPr','cNvGraphicFramePr','graphic'].includes(child.localName)) anchor.append(child);
        for(const x of anchor.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main','ext'))
            if(x.hasAttribute('cx')) {x.setAttribute('cx',width);x.setAttribute('cy',height);}
        inline.replaceWith(anchor);
    }
}
