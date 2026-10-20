import test from 'node:test';
import assert from 'node:assert/strict';
import {browserArgs} from '../browser-args.mjs';
test('headless Chrome launches one initial target, then Teams opens through the context',()=>{
 const args=browserArgs('C:\\profile',{headless:true});
 assert.equal(args.includes('--headless=new'),true);assert.equal(args.filter(x=>x.startsWith('https://')).length,1);
 assert.equal(browserArgs('C:\\profile').filter(x=>x.startsWith('https://')).length,2);
});

test('Graph mail leaves only Teams in the browser startup URLs',()=>{
 const args=browserArgs('profile',{apps:['teams']});
 assert.equal(args.some(x=>x.includes('outlook.office.com')),false);
 assert.ok(args.includes('https://teams.microsoft.com/v2/'));
});
