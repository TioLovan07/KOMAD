import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
const root=path.resolve(new URL('../docs',import.meta.url).pathname);
const sw=fs.readFileSync(path.join(root,'sw.js'),'utf8');
const assets=new Set(JSON.parse(sw.match(/ASSETS = (\[[^;]+\]);/)[1]));
for(const file of fs.readdirSync(path.join(root,'js')).filter(f=>f.endsWith('.js'))) {
 const name='js/'+file,source=fs.readFileSync(path.join(root,name),'utf8');
 for(const m of source.matchAll(/(?:from\s*|import\s*)['"]([^'"]+\.js)['"]/g)) {
  const target=path.posix.normalize(path.posix.join('js',m[1]));
  assert.ok(assets.has(target),`${name} imports ${target}, absent from offline cache`);
 }
}
console.log('PASS: every statically imported module, including the legacy DOC reader, is in the offline shell.');
