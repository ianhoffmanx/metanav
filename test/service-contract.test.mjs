import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtempSync,cpSync,writeFileSync,readFileSync,readdirSync,mkdirSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {spawn} from 'node:child_process';
test('service keeps dashboard contracts and accepts Calendar sign-in without opening auth during a busy run',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'metanav-service-'));let child;
 try {
  const source=fileURLToPath(new URL('..',import.meta.url)),out=join(dir,'out');
  for(const file of readdirSync(source).filter(f=>f.endsWith('.mjs')))cpSync(join(source,file),join(dir,file));
  mkdirSync(join(out,'.running'),{recursive:true});
  writeFileSync(join(dir,'config.json'),JSON.stringify({output_dir:out,graph:{provider:'microsoft-sdk'}}));
  const service=join(dir,'state.mjs');
  writeFileSync(service,readFileSync(service,'utf8').replace('const PORT = 47615;','const PORT = 0;')+"\nserver.on('listening',()=>process.send({port:server.address().port}));\n");
  child=spawn(process.execPath,[service],{stdio:['ignore','ignore','pipe','ipc']});
  let errors='';child.stderr.on('data',x=>errors+=x);
  const port=await new Promise((resolve,reject)=>{
   const timer=setTimeout(()=>reject(Error('Service did not start: '+errors)),8000);
   child.once('message',m=>{clearTimeout(timer);resolve(m.port)});
   child.once('error',e=>{clearTimeout(timer);reject(e)});
  });
  const {request}=await import('node:http');
  const post=(path,body={})=>new Promise((resolve,reject)=>{
   const req=request({host:'127.0.0.1',port,path,method:'POST',headers:{Host:'127.0.0.1:0','Content-Type':'application/json'}},response=>{
    let text='';response.on('data',x=>text+=x);response.on('end',()=>resolve({status:response.statusCode,body:response.statusCode===200?JSON.parse(text):null}));
   });req.on('error',reject);req.end(JSON.stringify(body));
  });
  const status=await post('/status');assert.equal(status.status,200);assert.deepEqual(Object.keys(status.body).sort(),['running','signing','signin'].sort());assert.equal(status.body.running,true);assert.deepEqual(status.body.signing,[]);
  assert.deepEqual((await post('/refresh')).body,{status:'busy'});
  assert.deepEqual(await post('/signin',{source:'Calendar'}),{status:200,body:{status:'busy'}});
  assert.equal((await post('/signin',{source:'Unknown'})).status,400);
 } finally {
  if(child && child.exitCode===null){const closed=new Promise(resolve=>child.once('close',resolve));child.kill();await closed}
  rmSync(dir,{recursive:true,force:true});
 }
});
