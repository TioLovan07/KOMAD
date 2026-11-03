/** Structured DOC 97-2003 → OOXML adapter, entirely local. No network conversion.
 * Unsupported row topology is rejected, never flattened into guessed columns.
 */
import '../vendor/legacy-doc.js';
const rgb=n=>[n&255,(n>>>8)&255,(n>>>16)&255].map(v=>v.toString(16).padStart(2,'0')).join('').toUpperCase();
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const escape=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
export async function prepareWordFile(file) {
    if(/\.docx$/i.test(file.name))return {file,legacy:false};
    if(!/\.doc$/i.test(file.name))throw Error('Pilih file Word .docx atau .doc 97–2003.');
    if(file.size>20*1024*1024)throw Error('Ukuran file melebihi 20 MB.');
    const bytes=new Uint8Array(await file.arrayBuffer());
    if(bytes[0]===0x50&&bytes[1]===0x4b)return {file:new File([bytes],file.name.replace(/\.doc$/i,'.docx')),legacy:false,renamed:true};
    const model=globalThis.docToText.model(bytes);
    if(!model?.body?.length)throw Error('DOC ini tidak dapat dibaca (rusak, terenkripsi, atau bukan Word 97–2003). Simpan sebagai .docx melalui Word.');
    if(model.body.length>100000)throw Error('DOC terlalu besar untuk dibaca pada perangkat ini.');
    // Floating raster emblems may use a shape anchor rather than character 0x01.
    // Only recover when one unique raster exists, the shape is in an identity
    // cell, and there is no textbox story to confuse it with ordinary content.
    const rasters=globalThis.docToText.images(bytes)||[];
    if(rasters.length===1 && !(model.textboxes||[]).length) for(let i=0;i<model.body.length;i++) {
        const p=model.body[i];
        for(const r of p.runs||[]) if(r.tbxRef!=null && model.shapes?.[r.tbxRef]) {
            let end=i;while(end<model.body.length&&!['cell','rowEnd'].includes(model.body[end].kind))end++;
            const following=model.body.slice(i,end+1).flatMap(p=>p.runs||[]).map(r=>r.text||'').join(' ');
            const shape=model.shapes[r.tbxRef];
            if(/\b(?:BUPATI|WALI\s*KOTA|GUBERNUR)\b/i.test(following)) r.image={...rasters[0],dxa:shape.xR-shape.xL,dya:shape.yB-shape.yT};
        }
    }
    const zip=new JSZip();let imageNo=0,recoveredRows=0;
    const rels=[];
    function para(p) {
        const marker=p.list?.marker;
        const props=[];
        if(p.pp) { const attrs={};if(Number.isFinite(p.pp.indL))attrs.left=p.pp.indL;if(Number.isFinite(p.pp.indR))attrs.right=p.pp.indR;if(Number.isFinite(p.pp.ind1))attrs[p.pp.ind1<0?'hanging':'firstLine']=Math.abs(p.pp.ind1);if(Object.keys(attrs).length)props.push('<w:ind '+Object.entries(attrs).map(([k,v])=>'w:'+k+'="'+v+'"').join(' ')+'/>'); }
        if(p.align)props.push('<w:jc w:val="'+(['left','center','right','both'][p.align]||'left')+'"/>');
        if(p.pp?.shd!=null)props.push('<w:shd w:val="clear" w:fill="'+rgb(p.pp.shd)+'"/>');
        const runs=[];
        if(marker&&(p.runs||[]).some(r=>r.text?.trim()&&!r.strike&&!r.dstrike&&!r.hidden))runs.push('<w:r><w:t xml:space="preserve">'+escape(marker)+' </w:t></w:r>');
        for(const r of p.runs||[]) {
            if(r.image) {
                const img=r.image,id='rIdLegacyImage'+(++imageNo),ext=img.mime==='image/png'?'png':'jpg';
                zip.file('word/media/legacy-'+imageNo+'.'+ext,new Uint8Array(img.bytes));
                rels.push('<Relationship Id="'+id+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/legacy-'+imageNo+'.'+ext+'"/>');
                const cx=Math.max(1,img.dxa||1800)*635,cy=Math.max(1,img.dya||1800)*635;
                runs.push('<w:r><w:drawing><wp:inline><wp:extent cx="'+cx+'" cy="'+cy+'"/><wp:docPr id="'+imageNo+'" name="Image '+imageNo+'"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="'+imageNo+'" name="Image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="'+id+'"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'+cx+'" cy="'+cy+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>');continue;
            }
            if(!r.text)continue;
            const rp=[];
            if(r.color!=null)rp.push('<w:color w:val="'+rgb(r.color)+'"/>');
            if(r.highlight!=null)rp.push('<w:shd w:val="clear" w:fill="'+rgb(r.highlight)+'"/>');
            for(const [k,tag] of [['b','b'],['i','i'],['strike','strike'],['dstrike','dstrike'],['hidden','vanish']])if(r[k])rp.push('<w:'+tag+'/>');
            if(r.u)rp.push('<w:u w:val="single"/>');
            if(r.size)rp.push('<w:sz w:val="'+Math.round(r.size*2)+'"/>');
            if(r.font)rp.push('<w:rFonts w:ascii="'+escape(r.font)+'" w:hAnsi="'+escape(r.font)+'"/>');
            const parts=r.text.split(/(\t|\n)/).map(t=>t==='\t'?'<w:tab/>':t==='\n'?'<w:br/>':'<w:t xml:space="preserve">'+escape(t)+'</w:t>').join('');
            runs.push('<w:r><w:rPr>'+rp.join('')+'</w:rPr>'+parts+'</w:r>');
        }
        return '<w:p><w:pPr>'+props.join('')+'</w:pPr>'+runs.join('')+'</w:p>';
    }
    const rows=[];let cells=[],pending=[],skippedAnnex=false;
    const cellText=c=>c.map(p=>(p.runs||[]).map(r=>r.text||'').join('')).join(' ').replace(/\s+/g,' ').trim();
    for(const p of model.body) {
        pending.push(p);
        if(p.kind==='cell'||p.kind==='rowEnd'){cells.push(pending);pending=[];}
        if(p.kind==='rowEnd') {
            const expected=(p.tblw?.length||6)-1;
            if(cells.length===expected)rows.push({cells,widths:p.tblw,shades:p.tblShd});
            else if(expected===5&&cells.length>10&&cells.slice(5,-5).every(c=>c.every(p=>(p.runs||[]).every(r=>!r.text?.trim()&&!r.image)))) {
                rows.push({cells:cells.slice(0,5),widths:p.tblw},{cells:cells.slice(-5),widths:p.tblw});recoveredRows++;
            } else throw Error('Struktur tabel DOC tidak pasti ('+cells.length+' sel, diharapkan '+expected+'). Simpan sebagai .docx melalui Word; aplikasi tidak menebak kolom.');
            const closing=rows.slice(-4,-1).some(r=>r.cells.some(c=>/^(?:BERITA|LEMBARAN) DAERAH\b/i.test(cellText(c))));
            const annexRow=cells.length>=4 && /^Lampiran(?:\s+[IVXLCDM]+)?$/i.test(cellText(cells[1])) && cells.slice(2).some(c=>/^LAMPIRAN\b/i.test(cellText(c)));
            cells=[];
            // Validated matrix closing and an explicit annex row: annex contents are outside current scope.
            if(closing&&annexRow){skippedAnnex=true;pending=[];break;}
        }
    }
    if(cells.length)throw Error('Baris tabel DOC tidak lengkap. Simpan ulang sebagai .docx.');
    const grid=rows.reduce((best,r)=>(r.widths?.length||0)>best.length?r.widths:best,[]);
    function tableRow(row) {
        let spans=row.cells.map(()=>1);
        if(row.widths && row.cells.length<grid.length-1) {
            const indices=row.widths.map(x=>grid.findIndex(y=>Math.abs(x-y)<=20));
            if(indices.every((n,i)=>n>=0&&(!i||n>indices[i-1])))spans=indices.slice(1).map((n,i)=>n-indices[i]);
            else throw Error('Batas gabungan sel DOC belum pasti. Simpan sebagai .docx agar kolom tidak bergeser.');
        }
        return '<w:tr>'+row.cells.map((c,i)=>'<w:tc><w:tcPr>'+(row.shades?.[i]!=null?'<w:shd w:val="clear" w:fill="'+rgb(row.shades[i])+'"/>':'')+(spans[i]>1?'<w:gridSpan w:val="'+spans[i]+'"/>':'')+'</w:tcPr>'+c.map(para).join('')+'</w:tc>').join('')+'</w:tr>';
    }
    const body=rows.length?'<w:tbl>'+rows.map(tableRow).join('')+'</w:tbl>'+pending.map(para).join(''):pending.map(para).join('');
    const main='<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="'+W+'" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>'+body+'<w:sectPr><w:pgSz w:w="12191" w:h="18711"/><w:pgMar w:top="1417" w:left="1417" w:right="1417" w:bottom="1417"/></w:sectPr></w:body></w:document>';
    zip.file('word/document.xml',main);
    zip.file('word/_rels/document.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+rels.join('')+'</Relationships>');
    zip.file('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    return {file:new File([await zip.generateAsync({type:'uint8array'})],file.name.replace(/\.doc$/i,'.docx')),legacy:true,recoveredRows,images:imageNo,skippedAnnex,model};
}
