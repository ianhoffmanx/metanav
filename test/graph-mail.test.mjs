import test from 'node:test';
import assert from 'node:assert/strict';
import {mailFolders,mailRow,mailInput,graphMailFailure} from '../graph-mail.mjs';
import {graphInputs} from '../graph.mjs';
import {captureSources} from '../capture-status.mjs';
import {signinEvidence} from '../health.mjs';
const message={id:'m1',conversationId:'same-thread',subject:'Please approve the blog',from:{emailAddress:{name:'Tina team',address:'tina@example.test'}},toRecipients:[{emailAddress:{name:'Ian'}}],receivedDateTime:'2026-10-19T23:00:00Z',sentDateTime:'2026-10-19T22:00:00Z',bodyPreview:'Can you approve it by Thursday?',isRead:true,flag:{flagStatus:'flagged'},webLink:'https://outlook.office.com/mail/id/m1'};
test('only configured mail folders are read; custom folders need an explicit Graph ID',()=>{
 assert.deepEqual(mailFolders({received_folders:[{name:'Inbox'}],sent:'url',skipped_folders:['GitHub']}),[{name:'Inbox',id:'inbox',sent:false},{name:'Sent',id:'sentitems',sent:true}]);
 assert.throws(()=>mailFolders({received_folders:[{name:'Unknown'}]}),/explicit graph_id/);
 assert.equal(mailFolders({received_folders:[{name:'Campaigns',graph_id:'folder-id'}]})[0].id,'folder-id');
});
test('Graph mail preserves people, flags, read state and the official link; received and sent use their own dates',()=>{
 const row=mailRow(message,{timezone:'Australia/Brisbane'});
 assert.equal(row.convid,'same-thread');assert.equal(row.unread,false);assert.equal(row.url,message.webLink);assert.match(row.label,/Flagged.*Tina team.*Thursday/);assert.equal(row.at,message.receivedDateTime);
 const sent=mailRow(message,{sent:true});assert.match(sent.label,/Ian/);assert.equal(sent.at,message.sentDateTime);
 assert.equal(mailRow({...message,isDraft:true}),null);
 assert.equal(mailRow({...message,receivedDateTime:'bad'}),null);
 assert.equal(mailRow({...message,webLink:'javascript:alert(1)'}).url,null);
});
test('duplicate messages are removed but multiple replies in one thread remain; absence is not deletion evidence',()=>{
 const input=mailInput({ok:true,folders:{Inbox:{messages:[message,message,{...message,id:'m2',bodyPreview:'Approval received'}]}}});
 assert.equal(input.folders.Inbox.length,2);assert.equal(input.deletion_evidence,false);assert.equal(input.signin_needed,false);assert.equal(input.provider,'microsoft-sdk');
});
test('Graph collection supplies mail independently of a calendar failure and never forwards app credentials',async()=>{
 const result=await graphInputs({provider:'microsoft-sdk',app:{tenant:'tenant',secret:'secret'},mailSince:'2026-09-20T00:00:00Z',first:true,mailConfig:{received_folders:[{name:'Inbox'}],sent:'url'},sdk:async(action,payload)=>{
  assert.equal(payload.secret,undefined);assert.equal(payload.first,true);assert.equal(payload.mail_folders.length,2);assert.match(payload.mail_since,/2026-09/);
  return {ok:false,error:'Graph calendar collection failed.',mail:{ok:true,folders:{Inbox:{messages:[message]}}}};
 }});
 assert.equal(result.ok,false);assert.equal(result.mail.folders.Inbox.length,1);assert.equal(result.mail.error,undefined);assert.match(result.error,/calendar/);
});
test('a Teams sign-in failure cannot turn working Graph mail into an Outlook failure',()=>{
 const mail=mailInput({ok:true,folders:{Inbox:{messages:[message]}}});
 const inputs={mail,teams:{signin_needed:true},github:{ok:true},graph:{provider:'microsoft-sdk',ok:true,calendar:[]}};
 const sources=captureSources(inputs);assert.equal(sources.find(s=>s.name==='Outlook').ok,true);assert.equal(sources.find(s=>s.name==='Calendar').ok,true);assert.equal(sources.find(s=>s.name==='Teams').ok,false);
 assert.match(signinEvidence(inputs).message,/Teams browser sign-in/);
 const failed=graphMailFailure({code:'signin_required'});assert.equal(failed.signin_needed,true);
 assert.match(captureSources({...inputs,mail:failed}).find(s=>s.name==='Outlook').note,/Graph sign-in is required/);
});

test('panel sign-in buttons route mail and calendar to Graph, and Teams to its browser',async()=>{
 const {graphSigninSource}=await import('../graph-mail.mjs');
 const cfg={graph:{provider:'microsoft-sdk'}};
 assert.equal(graphSigninSource('Outlook',cfg),true);
 assert.equal(graphSigninSource('Calendar',cfg),true);
 assert.equal(graphSigninSource('Teams',cfg),false);
 assert.equal(graphSigninSource('Outlook',{graph:{provider:'microsoft-sdk',mail:false}}),false);
 assert.equal(graphSigninSource('Outlook',{}),false);
});
