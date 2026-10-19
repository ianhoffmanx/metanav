import test from 'node:test';
import assert from 'node:assert/strict';
import { freshSignin, signinEvidence, missingSources, failureMessage } from '../health.mjs';
import { contentRelevant, marketingGithub, marketingPrevious } from '../marketing.mjs';
const options={orgs:['SSWConsulting','tinacms'],content_repos:['SSWConsulting/SSW.Products','tinacms/tina.io'],exclude_repos:['koreanpatch','polysemylabs','VictorCazanave/svg-maps']};
test('old ready status becomes stale and expiry is explicit',()=>{
 assert.equal(freshSignin({status:'ready',at:'2026-10-09T00:00:00Z'},Date.parse('2026-10-19T00:00:00Z')).status,'stale');
 assert.equal(signinEvidence({mail:{signin_needed:true},teams:{}}).status,'required');
 assert.equal(missingSources(true).length,4);
 assert.match(failureMessage({sources:[{ok:false,note:'Microsoft sign-in is required'}]}),/SIGN IN/);
});
test('content approval survives, technical PR and developer metrics do not',()=>{
 const content={url:'https://SSWConsulting/invalid',title:'Blog: YakShaver release',size:{files:9},ci:'FAILURE',unresolved_threads:30};
 content.url='https://github.com/SSWConsulting/SSW.Products/pull/1';
 const technical={url:'https://github.com/SSWConsulting/SSW.Products/pull/2',title:'Fix sitemap static rendering',files:[{path:'src/server.ts'}]};
 assert.equal(contentRelevant(technical,options),false);
 const got=marketingGithub({ok:true,work:[],review_requests:[content,technical]},options).review_requests;
 assert.equal(got.length,1);assert.equal(got[0].publication_blocked,true);assert.equal('size' in got[0],false);assert.equal('ci' in got[0],false);
});
test('excluded projects are removed from previous context and current input',()=>{
 const previous={queue:[{url:'https://github.com/VictorCazanave/svg-maps/issues/1'},{url:'https://github.com/tinacms/tina.io/pull/2'}]};
 assert.equal(marketingPrevious(previous,options).queue.length,1);
 assert.equal(contentRelevant({url:'https://github.com/tinacms/koreanpatch/pull/1',title:'Blog content'},options),false);
});
test('application code under app/blog and a CSS footer PR are not editorial content',()=>{
 for(const item of [
  {title:'Fix sitemap + static rendering',files:[{path:'app/[product]/blog/[slug]/page.tsx'}]},
  {title:'Use a Tailwind class for footer background',files:[{path:'content/footer/EagleEye/footer.json'}]},
  {title:'feat(ses): add content-neutral email template',files:[{path:'src/template.ts'}]},
 ])assert.equal(contentRelevant({...item,url:'https://github.com/SSWConsulting/SSW.Products/pull/1'},options),false);
 const old={queue:[{source:'github',url:'https://github.com/SSWConsulting/SSW.Products/pull/1'},{source:'teams',url:'msteams:/request'}]};
 assert.equal(marketingPrevious(old,options,{ok:true,excluded_prs:['https://github.com/SSWConsulting/SSW.Products/pull/1']}).queue.length,1);
});
