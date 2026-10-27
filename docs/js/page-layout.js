/** First-page header reserves title space without forcing later pages to 8 cm.
 * Section margins stay 2.5 cm. A first-page-only 7 cm header at y=1 cm
 * pushes the first body line to 8 cm in Word/LibreOffice.
 */
export async function configureFirstPage(doc, archive) {
    const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const P='http://schemas.openxmlformats.org/package/2006/relationships';
    const file=archive.file('word/_rels/document.xml.rels');
    const parser=new DOMParser(), serializer=new XMLSerializer();
    const rels=parser.parseFromString(file?await file.async('string'):'<Relationships xmlns="'+P+'"/>','application/xml');
    let id='rIdKomadFirstPage',n=0;
    while([...rels.documentElement.children].some(r=>r.getAttribute('Id')===id)) id='rIdKomadFirstPage'+(++n);
    const rel=rels.createElementNS(P,'Relationship');
    rel.setAttribute('Id',id);rel.setAttribute('Type',R+'/header');rel.setAttribute('Target','komad-first-page.xml');
    rels.documentElement.append(rel);
    archive.file('word/_rels/document.xml.rels',serializer.serializeToString(rels));
    archive.file('word/komad-first-page.xml','<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="'+W+'"><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="3969" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr><w:bookmarkStart w:id="2147000000" w:name="komad_first_page_spacer"/><w:bookmarkEnd w:id="2147000000"/><w:r><w:rPr><w:sz w:val="2"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r></w:p></w:hdr>');
    const types=parser.parseFromString(await archive.file('[Content_Types].xml').async('string'),'application/xml');
    const part='/word/komad-first-page.xml';
    if (![...types.documentElement.children].some(x=>x.getAttribute('PartName')===part)) {
        const override=types.createElementNS(types.documentElement.namespaceURI,'Override');
        override.setAttribute('PartName',part);
        override.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
        types.documentElement.append(override);
    }
    archive.file('[Content_Types].xml',serializer.serializeToString(types));
    const first=doc.getElementsByTagNameNS(W,'sectPr')[0];
    const ref=doc.createElementNS(W,'w:headerReference');ref.setAttributeNS(W,'w:type','first');ref.setAttributeNS(R,'r:id',id);
    first.insertBefore(ref,first.firstChild);
    const title=doc.createElementNS(W,'w:titlePg');first.append(title);
    const margin=first.getElementsByTagNameNS(W,'pgMar')[0];
    margin.setAttributeNS(W,'w:header','567');
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
