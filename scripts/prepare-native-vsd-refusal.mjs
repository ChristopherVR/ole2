import {readFileSync,writeFileSync,mkdirSync} from 'node:fs';
import {createHash} from 'node:crypto';
import {resolve,dirname,join} from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import assert from 'node:assert/strict';
// Accept only explicitly supplied trusted built ole2 modules. No Office or
// document code is executed by this driver; COM is a separate native gate.
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
if(!process.argv[2]) throw new Error('Historical 0.9.1 containment gate: supply an explicit trusted built 0.9.1 dist/index.js. Current bounded text writers admit the tested World edit.');
const modulePath=resolve(process.argv[2]);
const output=resolve(process.argv[3]??join(root,'.native-validation/native-visio'));
const {parseVsd}=await import(pathToFileURL(modulePath).href);
const sourcePath=join(root,'test/fixtures/vsd/native-visio16-v11.vsd');
const source=readFileSync(sourcePath),sha=b=>createHash('sha256').update(b).digest('hex');
mkdirSync(output,{recursive:true});
const noOp=parseVsd(source);noOp.pages[0].shapes.find(s=>s.id===1).text='Hello\n\n';
assert.equal(noOp.revision,0);assert.equal(noOp.dirty,false);assert.deepEqual(Buffer.from(noOp.serialize()),source);
const refused=parseVsd(source),shape=refused.pages[0].shapes.find(s=>s.id===1);
assert.equal(shape.text,'Hello\n\n');
assert.throws(()=>{shape.text='World\n\n'},e=>e.name==='UnsupportedOle2EditError'&&e.reason==='unsafe-block-relocation');
assert.equal(refused.revision,0);assert.equal(refused.dirty,false);assert.equal(shape.text,'Hello\n\n');
assert.deepEqual(Buffer.from(refused.serialize()),source);
const files=[];
for(const [name,doc] of [['native-noop.vsd',noOp],['native-refused.vsd',refused]]){const bytes=Buffer.from(doc.serialize()),path=join(output,name);writeFileSync(path,bytes);files.push({path,sha256:sha(bytes),bytes:bytes.length})}
const modules=['index.js','vsd-reader.js','vsd-writer.js','vsd-document.js','vsd-compression.js'].map(name=>{const path=join(dirname(modulePath),name);return{path,sha256:sha(readFileSync(path))}});
const manifest={sourcePath,sourceSha256:sha(source),modulePath,modules,files,refusal:'unsafe-block-relocation',assertions:['same-text setter stays clean','unsupported World\\n\\n setter refuses','revision stays zero','source text remains unchanged','both serializations preserve every source byte'],nativeConsumerGate:'Pending independent readonly macros-disabled Visio snapshots; no native fidelity inferred from these byte assertions.'};
writeFileSync(join(output,'refusal-manifest.json'),JSON.stringify(manifest,null,2));console.log(JSON.stringify(manifest,null,2));
