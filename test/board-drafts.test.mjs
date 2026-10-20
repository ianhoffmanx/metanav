import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtempSync,readFileSync,readdirSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {attachBoardContext} from '../board-context.mjs';
import {draftRequest,replyDrafts} from '../reply-drafts.mjs';
import {readerReply,readerLogin} from '../reader-auth.mjs';
const run=()=>({queue:[{id:'conversation',source:'mail',title:'Review content',reply_target:{message_id:'message',sender:'editor@example.test'}}],waiting:[]});
const request=()=>({key:'conversation',text:'Draft test only',request_id:'00000000-0000-4000-8000-000000000001'});
test('only collected incoming mail may supply a draft target; model targets are overwritten',()=>{
 const b=run();b.queue.push({id:'forged',source:'mail',reply_target:{message_id:'forged'}});
 attachBoardContext(b,{mail:{folders:{Inbox:[{convid:'conversation',graph_id:'latest',at:'2026-10-21',from:{address:'editor@example.test'},subject:'Blog'}],Sent:[{convid:'conversation',graph_id:'sent',at:'2026-10-22',from:{address:'ian@example.test'}}]}},graph:{ok:false}});
 assert.equal(b.queue[0].reply_target.message_id,'latest');assert.equal(b.queue[1].reply_target,undefined);
 assert.equal(draftRequest({...request(),key:'forged'},b),null);
 assert.equal(draftRequest({...request(),text:''},b),null);
});
test('calendar facts come from Graph; cancelled and declined events are excluded',()=>{
 const b=attachBoardContext({},{collected_at:'2026-10-21T00:00:00Z',graph:{ok:true,timezone:'Australia/Sydney',calendar:[{subject:'Content planning',start:'2026-10-21T10:00:00',end:'2026-10-21T10:30:00'},{cancelled:true},{response:'declined'}]}});
 assert.equal(b.agenda.events.length,1);assert.equal(b.agenda.timezone,'Australia/Sydney');assert.equal(b.agenda.generated_at,'2026-10-21T00:00:00Z');
});
test('duplicate requests, process restarts and uncertain writes cannot create duplicate drafts',async()=>{
 const directory=mkdtempSync(join(tmpdir(),'draft-receipts-'));let calls=0;
 try{
  const options={directory,loadRun:run,create:async()=>{calls++;return {status:'created',url:'https://outlook.office.com/mail/draft/1'}}};
  const create=replyDrafts(options);
  assert.equal((await create(request())).status,'created');
  assert.equal((await replyDrafts(options)(request())).status,'created');assert.equal(calls,1);
  assert.equal((await create({...request(),request_id:'00000000-0000-4000-8000-000000000003'})).status,'created');assert.equal(calls,1);
  assert.equal((await create({...request(),text:'changed'})).status,'request_conflict');
  const serialized=readFileSync(join(directory,readdirSync(directory)[0]),'utf8');assert.ok(!serialized.includes('Draft test only'));
  const uncertain={...options,create:async()=>{calls++;throw Error('private error')}};
  const r={...request(),text:'Another draft',request_id:'00000000-0000-4000-8000-000000000002'};
  assert.equal((await replyDrafts(uncertain)(r)).status,'check_drafts');
  assert.equal((await replyDrafts(uncertain)(r)).status,'check_drafts');assert.equal(calls,2);
 }finally{rmSync(directory,{recursive:true,force:true})}
});
test('a definite sign-in failure can be retried after connecting without changing the request',async()=>{
 const directory=mkdtempSync(join(tmpdir(),'draft-auth-'));let ready=false;
 try{
  const create=replyDrafts({directory,loadRun:run,create:async()=>({status:ready?'created':'signin_required'})});
  assert.equal((await create(request())).status,'signin_required');ready=true;
  assert.equal((await create(request())).status,'created');
 }finally{rmSync(directory,{recursive:true,force:true})}
});
test('Graph uses a delegated token and createReply only; no sends or read-state mutations',async()=>{
 const calls=[];let saved;
 const fetchImpl=async(url,o)=>{
  calls.push([url,o]);return {ok:true,json:async()=>url.includes('/token')?{access_token:'test-access',refresh_token:'rotated',scope:'Mail.ReadWrite Calendars.Read'}:{webLink:'https://outlook.office.com/mail/draft/test'}};
 };
 const r=await readerReply({tenant:'tenant',id:'reader',secret:'test-secret'},{load:()=>({refresh_token:'saved'}),save:v=>saved=v},{message_id:'a/b',text:'Draft only'},{fetchImpl});
 assert.equal(r.status,'created');assert.equal(saved.refresh_token,'rotated');assert.match(calls[1][0],/a%2Fb\/createReply$/);
 assert.equal(JSON.parse(calls[1][1].body).message.body.contentType,'Text');
 assert.ok(!calls.some(([u])=>/send|isRead/.test(u)));
 assert.equal(JSON.stringify(r).includes('test-access'),false);
});
test('missing write scope stops before a mailbox POST',async()=>{
 let calls=0;
 const result=await readerReply({tenant:'t',id:'r',secret:'s'},{load:()=>({refresh_token:'saved'})},{message_id:'m',text:'text'},
 {fetchImpl:async()=>{calls++;return {ok:true,json:async()=>({access_token:'read-only',scope:'Mail.Read'})}}});
 assert.equal(result.status,'permission_required');assert.equal(calls,1);
});
test('Reader sign-in checks callback state, PKCE, nonce and the configured account',async()=>{
 let authUrl,saved,callback,tokenCalls=0;
 const app={tenant:'test-tenant',id:'reader',secret:'secret'};
 const login=readerLogin(app,{save:v=>saved=v},'ian@example.test',{port:0,timeout:5000,
 open:u=>{authUrl=new URL(u);callback=(async()=>{
  const target=authUrl.searchParams.get('redirect_uri');
  const bad=await fetch(`${target}?state=wrong`);assert.equal(bad.status,400);
  const good=await fetch(`${target}?state=${authUrl.searchParams.get('state')}&code=test-code`);assert.equal(good.status,200);
 })()},fetchImpl:async(url,o)=>{
  tokenCalls++;const params=new URLSearchParams(o.body);assert.ok(params.get('code_verifier'));assert.equal(params.get('grant_type'),'authorization_code');
  const c={tid:app.tenant,nonce:authUrl.searchParams.get('nonce'),preferred_username:'ian@example.test'};
  return {ok:true,json:async()=>({access_token:'access',refresh_token:'refresh',scope:'Mail.ReadWrite Calendars.Read',id_token:'x.'+Buffer.from(JSON.stringify(c)).toString('base64url')+'.x'})};
 }});
 assert.equal(await login,true);await callback;assert.equal(tokenCalls,1);assert.equal(saved.account,'ian@example.test');assert.equal(saved.access_token,undefined);
});
