import test from 'node:test';
import assert from 'node:assert/strict';
import { cdpCall, recoverTargets } from '../cdp.mjs';
test('a stalled target times out and releases its socket', async () => {
  let closed = false;
  class Socket { constructor() { queueMicrotask(() => this.onopen?.()); } send() {} close() { closed = true; } }
  await assert.rejects(cdpCall('ws://test', 'Page.getFrameTree', {}, { timeout: 10, Socket }), /timed out/);
  assert.equal(closed, true);
});
test('recovery closes only the frozen page, leaves healthy and non-page targets alone', async () => {
  const closed = [], calls = [];
  const targets = [{ id:'healthy', type:'page', webSocketDebuggerUrl:'healthy' }, { id:'frozen', type:'page', webSocketDebuggerUrl:'frozen' }, { id:'worker', type:'worker' }];
  const result = await recoverTargets('http://localhost', {
    request: async url => url.endsWith('/json/list') ? {json:async()=>targets} : (closed.push(url),{ok:true}),
    call: async (url, method) => { calls.push([url,method]); if(url==='frozen') throw Error('timeout'); },
  });
  assert.deepEqual(result, ['frozen']);
  assert.deepEqual(closed, ['http://localhost/json/close/frozen']);
  assert.equal(calls.filter(c=>c[0]==='healthy').length, 2);
});
test('duplicate application tabs are closed so repeated sign-in attempts do not accumulate renderers', async()=>{
 const closed=[];
 const targets=[{id:'one',type:'page',url:'outlook',webSocketDebuggerUrl:'one'},{id:'two',type:'page',url:'outlook',webSocketDebuggerUrl:'two'}];
 await recoverTargets('http://localhost',{application:u=>u,call:async()=>({}),request:async url=>url.endsWith('/json/list')?{json:async()=>targets}:(closed.push(url),{ok:true})});
 assert.deepEqual(closed,['http://localhost/json/close/two']);
});
