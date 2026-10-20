import test from 'node:test';
import assert from 'node:assert/strict';
import {EventEmitter} from 'node:events';
import {sdkRequest,runSDK,SDK_CLIENT_ID} from '../graph-sdk.mjs';
import {deviceLogin,graphInputs} from '../graph.mjs';
test('SDK payload contains only tenant and collection facts, never the Reader secret or app ID',()=>{
 const request=sdkRequest('read',{tenant:'tenant.example',secret:'do-not-pass',id:'reader-app',from:'a',to:'b',timezone:'Australia/Brisbane'});
 assert.equal('secret' in request,false);assert.equal('id' in request,false);
 assert.throws(()=>sdkRequest('delete',{tenant:'tenant'}),/Unsupported/);
 assert.throws(()=>sdkRequest('login',{tenant:'bad/tenant'}),/invalid/);
 assert.throws(()=>sdkRequest('read',{tenant:'tenant',tracked_conversations:["x' or true"]}),/Invalid tracked/);
 assert.match(SDK_CLIENT_ID,/^[a-f0-9-]+$/);
});
test('SDK runs an executable directly, restricts actions, and does not expose authentication stderr',async()=>{
 let sent='',args;
 const spawnImpl=(exe,argv)=>{
  args=argv;const child=new EventEmitter();child.stdout=new EventEmitter();child.stderr=new EventEmitter();child.kill=()=>{};
  child.stdin=new EventEmitter();child.stdin.end=input=>{sent=input;queueMicrotask(()=>{child.stderr.emit('data','SECRET DETAILS');child.stdout.emit('data',JSON.stringify({ok:false,code:'signin_required',note:'Graph sign-in is required'}));child.emit('close',1)})};return child;
 };
 const result=await runSDK('read',{tenant:'ssw.example',secret:'never'}, {spawnImpl,powershell:'pwsh',timeout:1000});
 assert.equal(result.code,'signin_required');assert.equal(sent.includes('never'),false);assert.ok(args.includes('read'));assert.equal(JSON.stringify(result).includes('SECRET'),false);
});
test('the confidential Reader app is rejected before issuing another device code',async()=>{
 let calls=0;
 await assert.rejects(deviceLogin({tenant:'t',id:'reader',secret:'secret'},{},{fetchImpl:async()=>{calls++}}),/public-client/);
 assert.equal(calls,0);
});
test('SDK calendar input uses the existing calendar schema and reports the mail access check',async()=>{
 const data=await graphInputs({provider:'microsoft-sdk',app:{tenant:'t',id:'reader',secret:'secret'},now:new Date(2026,9,20),sdk:async(action,payload)=>{
  assert.equal(action,'read');assert.equal(payload.secret,undefined);
  return {ok:true,mail_read_ok:true,calendar:[{id:'x',subject:'Planning',start:{dateTime:'2026-10-21T10:00:00.0000000'},end:{dateTime:'2026-10-21T11:00:00.0000000'}}]};
 }});
 assert.equal(data.ok,true);assert.equal(data.calendar[0].start,'2026-10-21T10:00:00');assert.equal(data.mail_read_ok,true);
 const none=await graphInputs({provider:'microsoft-sdk',app:{tenant:'t'},sdk:async()=>({ok:false,skipped:true,note:'Graph is not signed in'})});
 assert.equal(none.skipped,true);
});

// Exercise the actual Windows boundary without network access, credentials or a browser.
test('Windows SDK script parses and a fresh profile returns not signed in without prompting', {skip:process.platform!=='win32'}, async()=>{
 const {spawnSync}=await import('node:child_process');
 const {mkdtempSync,rmSync}=await import('node:fs');
 const {tmpdir}=await import('node:os');
 const {join}=await import('node:path');
 const {fileURLToPath}=await import('node:url');
 const isolated=mkdtempSync(join(tmpdir(),'metanav-sdk-'));
 try {
  const script=fileURLToPath(new URL('../windows/graph-sdk.ps1',import.meta.url));
  const exe=join(process.env.ProgramFiles,'PowerShell','7','pwsh.exe');
  const run=spawnSync(exe,['-NoProfile','-File',script,'-Action','status'],{input:JSON.stringify({tenant:'example.onmicrosoft.com'}),encoding:'utf8',timeout:10000,windowsHide:true,env:{...process.env,LOCALAPPDATA:isolated}});
  assert.equal(run.status,0,run.stderr);
  assert.deepEqual(JSON.parse(run.stdout),{ok:false,skipped:true,code:'not_signed_in',note:'Graph is not signed in'});
 } finally {rmSync(isolated,{recursive:true,force:true})}
});

test('an enabled calendar failure is visible and its sign-in alert points to the dedicated desktop shortcut',async()=>{
 const {captureSources,captureComplete}=await import('../capture-status.mjs');
 const {failureMessage,missingSources}=await import('../health.mjs');
 const inputs={mail:{folders:{},calendar:[]},teams:{conversations:[]},github:{ok:true}};
 assert.equal(captureComplete(inputs),true,'unconfigured legacy calendar remains optional');
 inputs.graph={provider:'microsoft-sdk',ok:false,code:'signin_required'};
 const calendar=captureSources(inputs).find(s=>s.name==='Calendar');
 assert.equal(calendar.ok,false);assert.match(calendar.note,/sign-in is required/);
 assert.equal(captureComplete(inputs),false);
 assert.match(failureMessage({sources:[calendar]}),/Meta-Nav Calendar Sign-in on your desktop/);
 inputs.graph={provider:'microsoft-sdk',ok:true,calendar:[]};
 assert.equal(captureComplete(inputs),true,'an empty calendar is still a successful read');
 assert.equal(missingSources(true,true).at(-1).name,'Calendar');
});

test('a stalled SDK is terminated inside its collection budget instead of keeping the collector alive',async()=>{
 let killed=false;
 const spawnImpl=()=>{
  const child=new EventEmitter();child.stdout=new EventEmitter();child.stderr=new EventEmitter();child.stdin=new EventEmitter();
  child.stdin.end=()=>{};child.kill=()=>{killed=true};return child;
 };
 await assert.rejects(runSDK('read',{tenant:'example.org'},{spawnImpl,timeout:10}),/timed out/);
 assert.equal(killed,true);
});
