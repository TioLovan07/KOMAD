/* Optional integration check for exported drafts. Start a local docs server first.
 * KOMAD_PLAYWRIGHT_MODULE=/path/to/playwright KOMAD_BROWSER_EXECUTABLE=/path/to/chromium
 * node tests/test-preview-browser.cjs /path/to/Draft.docx [...]
 * No office documents are bundled with this test.
 */
const fs = require('node:fs');
const assert = require('node:assert/strict');
const {chromium} = require(process.env.KOMAD_PLAYWRIGHT_MODULE || 'playwright');
(async () => {
    assert.ok(process.argv.length > 2, 'Supply exported DOCX paths');
    for (const file of process.argv.slice(2)) {
        const browser = await chromium.launch({
            executablePath: process.env.KOMAD_BROWSER_EXECUTABLE,
            args: ['--no-sandbox','--disable-gpu','--disable-dev-shm-usage','--no-zygote','--single-process']
        });
        try {
            const page = await browser.newPage();
            await page.goto((process.env.KOMAD_TEST_URL || 'http://localhost:8765') + '/viewer.html');
            const result = await page.evaluate(async bytes => {
                const archive = await JSZip.loadAsync(new Uint8Array(bytes));
                const xml = new DOMParser().parseFromString(await archive.file('word/document.xml').async('string'), 'application/xml');
                const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                const body = xml.getElementsByTagNameNS(W,'body')[0];
                const key = text => text.replace(/\s/g,'');
                const expected = key([...body.getElementsByTagNameNS(W,'t')].map(t=>t.textContent).join(''));
                await renderDraft(new Uint8Array(bytes).buffer);
                const actual = key([...document.querySelectorAll('article')].map(a=>a.textContent).join(''));
                const sections = [...document.querySelectorAll('section.worddoc')];
                const annex = [...document.querySelectorAll('article p')].filter(p=>/^LAMPIRAN [IVX]+$/.test(p.textContent.trim()));
                return {
                    same: expected === actual,
                    firstTextTop: (()=>{ const p=[...sections[0].querySelectorAll('article p')].find(p=>p.textContent.trim());return (p.getBoundingClientRect().top-sections[0].getBoundingClientRect().top)*2.54/96; })(),
                    sized: sections.every(s=>Math.abs(s.getBoundingClientRect().height-parseFloat(getComputedStyle(s).minHeight)) < 2),
                    images: [...document.images].every(i=>i.complete && i.naturalWidth>0 && i.getBoundingClientRect().width>0),
                    annexPages: annex.map(p=>sections.indexOf(p.closest('section'))),
                    pages: sections.length
                };
            }, [...fs.readFileSync(file)]);
            assert.ok(Math.abs(result.firstTextTop-8)<0.15, 'First text does not start at 8 cm');
            assert.ok(result.same, 'Preview lost or reordered document text');
            assert.ok(result.sized, 'Preview page exceeds the specified paper height');
            assert.ok(result.images, 'Preview contains an invisible/broken image');
            assert.equal(new Set(result.annexPages).size,result.annexPages.length,'Appendix titles share a page');
            console.log('PASS:',file,result.pages,'pages; complete text, images, page size, appendix breaks.');
        } finally { await browser.close(); }
    }
})().catch(error=>{ console.error(error); process.exit(1); });
