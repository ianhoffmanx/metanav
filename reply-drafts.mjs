// Only creates reply drafts. No send, mark-read, delete or arbitrary Graph endpoint.
import {createHash} from 'node:crypto';
import {existsSync,mkdirSync,readFileSync,readdirSync,renameSync,rmSync,writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {itemKey} from './board-context.mjs';
const hash=s=>createHash('sha256').update(s).digest('hex');
export function draftRequest(body,run){
  if(!body||typeof body.key!=='string'||typeof body.text!=='string'||!body.text.trim()||body.text.length>10000||
    typeof body.request_id!=='string'||!/^[a-f0-9-]{36}$/i.test(body.request_id))return null;
  const item=[...(run.queue||[]),...(run.waiting||[])].find(i=>itemKey(i)===body.key&&i.source==='mail'&&i.reply_target?.message_id);
  if(!item)return null;
  return {request_id:body.request_id,message_id:item.reply_target.message_id,text:body.text,
    fingerprint:hash(JSON.stringify([body.key,item.reply_target.message_id,body.text]))};
}
export function replyDrafts({directory,loadRun,create}){
  const active=new Set();
  return async body=>{
    const plan=draftRequest(body,loadRun());
    if(!plan)return {status:'unknown_item'};
    if(active.size)return {status:'busy'};
    active.add(plan.request_id);
    const file=join(directory,plan.request_id+'.json');
    const save=value=>{writeFileSync(file+'.tmp',JSON.stringify(value),{mode:0o600});renameSync(file+'.tmp',file)};
    try{
      mkdirSync(directory,{recursive:true,mode:0o700});
      if(existsSync(file)){
        const saved=JSON.parse(readFileSync(file,'utf8'));
        if(saved.fingerprint!==plan.fingerprint)return {status:'request_conflict'};
        return saved.result||{status:'check_drafts'}; // Interrupted or ambiguous POST: never repeat it.
      }
      // Reloading the page creates a new browser request ID. Identical text to
      // the same collected message still resolves to the existing receipt.
      for(const name of readdirSync(directory).filter(f=>/^[a-f0-9-]{36}\.json$/i.test(f))){
        const saved=JSON.parse(readFileSync(join(directory,name),'utf8'));
        if(saved.fingerprint===plan.fingerprint)return saved.result||{status:'check_drafts'};
      }
      save({fingerprint:plan.fingerprint,at:new Date().toISOString()});
      let result;
      try{result=await create(plan)}catch{result={status:'check_drafts'}}
      // Keep receipts, not the email text, and do not expose raw SDK errors.
      const safe={status:['created','permission_required','signin_required','not_configured','check_drafts'].includes(result?.status)?result.status:'check_drafts'};
      if(safe.status==='created'&&typeof result.url==='string'){
        try{const u=new URL(result.url);if(u.protocol==='https:'&&/^(outlook\.office(?:365)?\.com|outlook\.cloud\.microsoft)$/.test(u.hostname))safe.url=u.href}catch{}
      }
      // These failures establish that no draft was created. Permit the same
      // request after the user repairs consent/sign-in; ambiguous writes stay locked.
      if(['permission_required','signin_required','not_configured'].includes(safe.status))rmSync(file,{force:true});
      else save({fingerprint:plan.fingerprint,at:new Date().toISOString(),result:safe});
      return safe;
    }finally{active.delete(plan.request_id)}
  };
}
