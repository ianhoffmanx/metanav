// The tenant-approved Reader app uses its registered localhost web callback.
// Tokens stay in this Windows user's DPAPI store; only explicit login opens a browser.
import {randomBytes,createHash} from 'node:crypto';
import {createServer} from 'node:http';
import {spawn} from 'node:child_process';
export const CALLBACK='http://localhost:47616/auth/callback';
export const READER_SCOPES='openid profile offline_access https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Calendars.Read';
const claims=t=>{try{return JSON.parse(Buffer.from(t.split('.')[1],'base64url').toString())}catch{return {}}};
export async function exchange(app,params,fetchImpl=fetch){
  const r=await fetchImpl(`https://login.microsoftonline.com/${encodeURIComponent(app.tenant)}/oauth2/v2.0/token`,{
    method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
    body:new URLSearchParams({client_id:app.id,client_secret:app.secret,...params}),signal:AbortSignal.timeout(20000)});
  const j=await r.json();
  if(!r.ok||!j.access_token)return {ok:false,code:j.error==='invalid_grant'?'signin_required':'permission_required'};
  if(!String(j.scope||'').split(' ').some(s=>s==='Mail.ReadWrite'||s==='https://graph.microsoft.com/Mail.ReadWrite'))return {ok:false,code:'permission_required'};
  return {ok:true,...j};
}
export async function readerAccess(app,store,{fetchImpl=fetch}={}){
  const saved=store.load();
  if(!saved?.refresh_token)return {status:'signin_required'};
  const t=await exchange(app,{grant_type:'refresh_token',refresh_token:saved.refresh_token,scope:READER_SCOPES},fetchImpl);
  if(!t.ok)return {status:t.code};
  if(t.refresh_token)store.save({...saved,refresh_token:t.refresh_token,scope:t.scope});
  return {status:'ready',token:t.access_token};
}
export async function readerReply(app,store,plan,{fetchImpl=fetch}={}){
  const auth=await readerAccess(app,store,{fetchImpl});
  if(auth.status!=='ready')return {status:auth.status};
  if(!plan)return {status:'ready'};
  // A single POST creates a draft to the message's reply recipients. No send call.
  try{
    const r=await fetchImpl(`https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(plan.message_id)}/createReply`,{
      method:'POST',headers:{Authorization:'Bearer '+auth.token,'Content-Type':'application/json'},
      body:JSON.stringify({message:{body:{contentType:'Text',content:plan.text}}}),signal:AbortSignal.timeout(20000)});
    if(!r.ok)return {status:r.status===403?'permission_required':'check_drafts'};
    const draft=await r.json();return {status:'created',url:draft.webLink};
  }catch{return {status:'check_drafts'}}
}
export async function readerLogin(app,store,account,{fetchImpl=fetch,open=url=>{
  const child=spawn('rundll32.exe',['url.dll,FileProtocolHandler',url],{stdio:'ignore',windowsHide:true});
  child.on('error',()=>{});child.unref();
},port=47616,timeout=300000}={}){
  if(!app.secret||!app.id||!app.tenant||!account)throw Error('Reader application credentials are not configured.');
  const state=randomBytes(32).toString('base64url'),nonce=randomBytes(32).toString('base64url'),verifier=randomBytes(48).toString('base64url');
  const challenge=createHash('sha256').update(verifier).digest('base64url');
  let callback;
  let finish,timer,handling=false;
  const result=new Promise(resolve=>finish=resolve);
  const server=createServer(async(req,res)=>{
    res.setHeader('Cache-Control','no-store');res.setHeader('Content-Type','text/plain; charset=utf-8');
    res.setHeader('Content-Security-Policy',"default-src 'none'; frame-ancestors 'none'");
    const url=new URL(req.url,'http://localhost');
    if(req.method!=='GET'||req.headers.host!==`localhost:${port}`||url.pathname!=='/auth/callback'||url.searchParams.get('state')!==state){res.writeHead(400);res.end('Invalid sign-in callback.');return}
    if(handling){res.writeHead(409);res.end('Sign-in is already being checked.');return}handling=true;
    try{
      const code=url.searchParams.get('code');if(!code)throw Error();
      const t=await exchange(app,{grant_type:'authorization_code',code,redirect_uri:callback,code_verifier:verifier},fetchImpl);
      const id=claims(t.id_token),user=id.preferred_username||id.upn||id.email;
      if(!t.ok||!t.refresh_token||id.nonce!==nonce||user?.toLowerCase()!==account.toLowerCase()||id.tid!==app.tenant)throw Error();
      store.save({refresh_token:t.refresh_token,scope:t.scope,account:user,signed_in_at:new Date().toISOString()});
      res.end('Meta-Nav reply drafts are connected. You can close this tab and return to your board.',()=>finish(true));
    }catch{res.writeHead(400);res.end('Sign-in could not be completed for the configured SSW account. Return to Meta-Nav to try again.',()=>finish(false))}
  });
  try{
    await new Promise((resolve,reject)=>{server.once('error',reject);server.listen(port,'127.0.0.1',resolve)});
    port=server.address().port;
    callback=port===47616?CALLBACK:`http://localhost:${port}/auth/callback`;
    timer=setTimeout(()=>finish(false),timeout);
    const query=new URLSearchParams({client_id:app.id,response_type:'code',redirect_uri:callback,response_mode:'query',scope:READER_SCOPES,state,nonce,
      code_challenge:challenge,code_challenge_method:'S256',login_hint:account});
    open(`https://login.microsoftonline.com/${encodeURIComponent(app.tenant)}/oauth2/v2.0/authorize?${query}`);
    return await result;
  }finally{clearTimeout(timer);server.closeAllConnections();await new Promise(resolve=>server.close(resolve))}
}
