// The Microsoft SDK handles public-client browser sign-in, encrypted caching and silent refresh.
// The Reader app's secret is never passed to this provider. Only an explicit login may open a window.
import {spawn} from 'node:child_process';
import {join,dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
const HERE=dirname(fileURLToPath(import.meta.url));
export const SDK_CLIENT_ID='14d82eec-204b-4c2f-b7e8-296a70dab67e';
export function sdkRequest(action, {tenant, login_hint, from, to, timezone, mail_since, mail_folders, first,tracked_conversations}={}) {
  if(!['login','read','status'].includes(action))throw Error('Unsupported Graph SDK action');
  if(!/^[a-zA-Z0-9.-]+$/.test(tenant||''))throw Error('Graph tenant is missing or invalid');
  if(mail_since && !Number.isFinite(Date.parse(mail_since)))throw Error('Invalid Graph mail window');
  if(mail_folders && (!Array.isArray(mail_folders)||mail_folders.some(f=>!f.name||!f.id||typeof f.id!=='string')))throw Error('Invalid Graph mail folders');
  if(tracked_conversations&&(!Array.isArray(tracked_conversations)||tracked_conversations.length>12||tracked_conversations.some(s=>typeof s!=='string'||!/^[A-Za-z0-9+/_=-]{20,2048}$/.test(s))))throw Error('Invalid tracked mail conversations');
  return {tenant,login_hint,from,to,timezone,mail_since,first:!!first,mail_folders:mail_folders?.map(({name,id,sent})=>({name,id,sent:!!sent})),tracked_conversations};
}
export function runSDK(action, options={}, {spawnImpl=spawn,powershell=process.platform==='win32'?join(process.env.ProgramFiles||'C:\\Program Files','PowerShell','7','pwsh.exe'):'pwsh',timeout=action==='login'?330000:action==='read'?45000:60000}={}) {
  const payload=sdkRequest(action,options);
  return new Promise((resolve,reject)=>{
    const child=spawnImpl(powershell,['-NoLogo','-NoProfile','-File',join(HERE,'windows','graph-sdk.ps1'),'-Action',action],{windowsHide:action!=='login',stdio:['pipe','pipe','pipe']});
    let output='',finished=false;
    const finish=(error,value)=>{if(finished)return;finished=true;clearTimeout(timer);error?reject(error):resolve(value)};
    const timer=setTimeout(()=>{child.kill();finish(Error('Graph SDK timed out; the saved session was retained.'))},timeout);
    child.once('error',()=>finish(Error('Graph SDK could not start. Install PowerShell 7 and Microsoft.Graph.Authentication.')));
    child.stdout.on('data',chunk=>{output+=chunk;if(output.length>32<<20){child.kill();finish(Error('Graph SDK output exceeded its collection limit.'))}});
    // SDK stderr can contain authentication details: never forward it into run logs.
    child.stderr.on('data',()=>{});
    child.once('close',()=>{
      if(finished)return;
      try {const result=JSON.parse(output.replace(/^\uFEFF/,'').trim());if(!result||typeof result.ok!=='boolean')throw Error();finish(null,result)}
      catch {finish(Error('Graph SDK did not return a usable result.'))}
    });
    child.stdin.on('error',()=>{});
    child.stdin.end(JSON.stringify(payload));
  });
}
