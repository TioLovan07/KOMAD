/* Optional regression: supply a matrix with an opening image in the initial
 * identity row and Paragraf headings. No office source file is bundled.
 * Start docs server at KOMAD_TEST_URL, then run with the input .doc or .docx.
 */
const fs=require('node:fs'),assert=require('node:assert/strict');
const {chromium}=require(process.env.KOMAD_PLAYWRIGHT_MODULE||'playwright');
(async()=>{
 const file=process.argv[2];assert.ok(file,'Supply a matrix source');
 const browser=await chromium.launch({executablePath:process.env.KOMAD_BROWSER_EXECUTABLE,args:['--no-sandbox','--disable-gpu','--disable-dev-shm-usage','--no-zygote','--single-process']});
 try {
  const page=await browser.newPage();await page.goto(process.env.KOMAD_TEST_URL||'http://127.0.0.1:8877/');
  const result=await page.evaluate(async({bytes,name})=>{
   const {readMatrix}=await import('./js/matrix-parser.js'),{setModel}=await import('./js/state.js'),{buildWord}=await import('./js/word-builder.js');
   const m=await readMatrix(new File([new Uint8Array(bytes)],name));setModel(m);
   const blob=await buildWord({allowPending:true}),z=await JSZip.loadAsync(await blob.arrayBuffer());
   const d=new DOMParser().parseFromString(await z.file('word/document.xml').async('string'),'application/xml'),W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
   const ps=[...d.getElementsByTagNameNS(W,'body')[0].children].filter(p=>p.localName==='p');
   const text=p=>[...p.getElementsByTagNameNS(W,'t')].map(t=>t.textContent).join('');
   const layout=p=>{const pr=p.getElementsByTagNameNS(W,'pPr')[0],ind=pr?.getElementsByTagNameNS(W,'ind')[0];return {align:pr?.getElementsByTagNameNS(W,'jc')[0]?.getAttributeNS(W,'val'),left:ind?.getAttributeNS(W,'left'),right:ind?.getAttributeNS(W,'right')};};
   const article=ps.find(p=>/^Pasal \d+$/.test(text(p)));
   const labels=ps.flatMap((p,i)=>/^Paragraf \d+$/.test(text(p))?[layout(p),layout(ps[i+1])]:[]);
   return {article:layout(article),labels,image:d.getElementsByTagNameNS(W,'drawing').length,firstBlockImage:ps[0].getElementsByTagNameNS(W,'drawing').length};
  },{bytes:[...fs.readFileSync(file)],name:require('node:path').basename(file)});
  assert.ok(result.labels.length,'Fixture needs Paragraf labels');
  for(const layout of result.labels)assert.deepEqual(layout,result.article,'Label and title must use the same centered field as Pasal');
  assert.equal(result.image,1,'Opening emblem missing or duplicated');assert.equal(result.firstBlockImage,1,'Emblem must occupy the opening spacer');
  console.log('PASS: opening emblem retained once; Paragraf labels and titles match Pasal centering.');
 } finally {await browser.close();}
})().catch(e=>{console.error(e);process.exit(1)});
