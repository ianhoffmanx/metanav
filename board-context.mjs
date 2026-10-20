// Deterministic Graph facts for the board. Luna still chooses and ranks the work.
const clean=(s,n)=>typeof s==='string'?s.trim().slice(0,n):'';
const outlook=u=>{try{const x=new URL(u);return x.protocol==='https:'&&/^(outlook\.office(?:365)?\.com|outlook\.cloud\.microsoft)$/.test(x.hostname)?x.href:null}catch{return null}};
export const itemKey=i=>String(i.id||i.url||i.title||i.what||'');
export function attachBoardContext(brief,inputs){
  // Never trust a target proposed by the judge or carried over from an older run.
  const incoming=Object.entries(inputs.mail?.folders||{}).filter(([name])=>name!=='Sent')
    .flatMap(([,rows])=>Array.isArray(rows)?rows:[]).filter(m=>m.graph_id&&m.from?.address);
  brief.reply_coverage={received:incoming.length,mail_items:0,matched:0,id_matches:0,url_matches:0};
  for(const section of ['queue','waiting'])for(const item of brief[section]||[]){
    delete item.reply_target;
    if(item.source!=='mail')continue;
    brief.reply_coverage.mail_items++;
    if(incoming.some(m=>item.id===m.convid||item.id===m.graph_id))brief.reply_coverage.id_matches++;
    if(incoming.some(m=>outlook(item.reply_url||item.url)&&outlook(item.reply_url||item.url)===outlook(m.url)))brief.reply_coverage.url_matches++;
    const candidates=incoming.filter(m=>item.id===m.convid||item.id===m.graph_id||
      (outlook(item.reply_url||item.url)&&outlook(item.reply_url||item.url)===outlook(m.url)))
      .sort((a,b)=>String(b.at).localeCompare(String(a.at)));
    const message=candidates[0];
    if(message){brief.reply_coverage.matched++;item.reply_target={message_id:clean(message.graph_id,2048),sender:clean(message.from.address,254),
      name:clean(message.from.name,100),subject:clean(message.subject,300)};}
  }
  const g=inputs.graph;
  brief.agenda={ok:g?.ok===true,generated_at:inputs.collected_at||brief.generated_at,timezone:clean(g?.timezone,80),
    events:g?.ok===true?(g.calendar||[]).filter(e=>!e.cancelled&&e.response!=='declined').slice(0,200).map(e=>({
      title:clean(e.subject,140),start:clean(e.start,40),end:clean(e.end,40),all_day:e.all_day===true,
      location:clean(e.location,120),url:outlook(e.url),online_url:clean(e.online_url,1000)||null})):[]};
  return brief;
}
