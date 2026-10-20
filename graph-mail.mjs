// Keep the collector's mail shape while replacing Outlook page reads with read-only Graph data.
export function mailFolders(mail={}) {
  const received=(mail.received_folders||[]).map(folder=>{
    const id=folder.graph_id || (folder.name?.toLowerCase()==='inbox'?'inbox':null);
    if(!id)throw Error(`Graph mail needs an explicit graph_id for the configured folder ${folder.name}.`);
    return {name:folder.name,id,sent:false};
  });
  return [...received,...(mail.sent?[{name:'Sent',id:'sentitems',sent:true}]:[])];
}
const clean=(text,n=800)=>String(text||'').replace(/\s+/g,' ').trim().slice(0,n);
const person=p=>p?.emailAddress?.name||p?.emailAddress?.address||'';
export function mailRow(message,{sent=false,timezone='Australia/Brisbane'}={}) {
  if(!message.conversationId||message.isDraft)return null;
  const at=sent?message.sentDateTime:message.receivedDateTime;
  if(!Number.isFinite(Date.parse(at)))return null;
  const who=sent?(message.toRecipients||[]).map(person).filter(Boolean).join('; '):person(message.from);
  const flags=[message.flag?.flagStatus==='flagged'?'Flagged':'',message.hasAttachments?'Has attachments':'',message.importance==='high'?'High importance':''].filter(Boolean).join(' ');
  const time=new Intl.DateTimeFormat('en-AU',{timeZone:timezone,day:'numeric',month:'numeric',year:'numeric',hour:'numeric',minute:'2-digit',hour12:true}).format(new Date(at));
  let url=null;try{const u=new URL(message.webLink);if(u.protocol==='https:'&&/^(outlook\.office(?:365)?\.com|outlook\.cloud\.microsoft)$/.test(u.hostname))url=u.href;}catch{}
  return {convid:message.conversationId,graph_id:message.id,label:clean([flags,who,message.subject,message.bodyPreview].filter(Boolean).join(' · ')),
    unread:!message.isRead,time,at,url,subject:clean(message.subject,300),from:message.from?.emailAddress||null,to:(message.toRecipients||[]).map(x=>x.emailAddress),cc:(message.ccRecipients||[]).map(x=>x.emailAddress),preview:clean(message.bodyPreview,500)};
}
export function mailInput(raw,{timezone}={}) {
  const folders={};
  for(const [name,data] of Object.entries(raw?.folders||{})){
    const seen=new Set();folders[name]=(data.messages||[]).filter(m=>{if(seen.has(m.id))return false;seen.add(m.id);return true})
      .map(m=>mailRow(m,{sent:!!data.sent,timezone})).filter(Boolean).sort((a,b)=>b.at.localeCompare(a.at));
  }
  return {provider:'microsoft-sdk',signin_needed:false,folders,calendar:[],deletion_evidence:false,
    ...(raw?.ok?{}:{error:raw?.error||'Graph mail collection failed.'})};
}
export function graphMailFailure(graph={}) {
 const signin=!!graph.skipped||graph.code==='signin_required';
 return {provider:'microsoft-sdk',signin_needed:signin,folders:{},calendar:[],deletion_evidence:false,
   error:signin?'Graph sign-in is required. Open Meta-Nav Calendar Sign-in on your desktop.':graph.error||graph.note||'Graph mail collection failed.'};
}
export const graphSigninSource=(source,config={})=>config.graph?.provider==='microsoft-sdk' && (source==='Calendar'||source==='Outlook'&&config.graph.mail!==false);
