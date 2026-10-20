import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtempSync,cpSync,mkdirSync,writeFileSync,readFileSync,readdirSync,existsSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {spawnSync} from 'node:child_process';
import {fileURLToPath} from 'node:url';
test('a collector crash preserves the last result, reports every source, and never reads a missing brief',()=>{
 const dir=mkdtempSync(join(tmpdir(),'metanav-failure-'));
 try {
  const source=fileURLToPath(new URL('..',import.meta.url)), out=join(dir,'out');
  for(const file of readdirSync(source).filter(f=>f.endsWith('.mjs')||f==='template.html'))cpSync(join(source,file),join(dir,file));
  mkdirSync(join(out,'runs'),{recursive:true});
  const previous={headline:'Existing marketing actions',queue:[],waiting:[],calendar:[],highlights:[],done:[],generated_at:'2026-10-19T00:00:00Z'};
  writeFileSync(join(out,'runs','20200101-0000.json'),JSON.stringify(previous));
  writeFileSync(join(out,'.manual'),'');
  writeFileSync(join(dir,'config.json'),JSON.stringify({output_dir:out,ado:{base_url:'https://dev.azure.com/ssw/SSW.Marketing'}}));
  writeFileSync(join(dir,'collect.mjs'),"console.error('fixture collection failed');process.exit(1)");
  // Avoid sending a test toast on Windows; the production path is verified by the real service run.
  writeFileSync(join(dir,'notifications.mjs'),'export const notify=()=>true;');
  const result=spawnSync(process.execPath,[join(dir,'run.mjs')],{encoding:'utf8',timeout:10000});
  assert.equal(result.status,0,result.stderr);
  const log=JSON.parse(readFileSync(join(out,'logs',readdirSync(join(out,'logs'))[0]),'utf8'));
  assert.equal(log.is_error,true);assert.equal(log.collection_complete,false);
  assert.deepEqual(log.sources.map(s=>[s.name,s.ok]),[['Outlook',false],['Teams',false],['GitHub',false],['Azure DevOps',false]]);
  assert.equal('brief_error' in log,false);assert.equal(log.notification_sent,true);
  assert.equal(existsSync(join(out,'.inputs')),false);assert.equal(existsSync(join(out,'.running')),false);
  assert.deepEqual(JSON.parse(readFileSync(join(out,'runs','20200101-0000.json'),'utf8')),previous);
 } finally {rmSync(dir,{recursive:true,force:true});}
});
