import test from 'node:test';
import assert from 'node:assert/strict';
import {browserArgs} from '../browser-args.mjs';
test('headless Chrome launches one initial target, then Teams opens through the context',()=>{
 const args=browserArgs('C:\\profile',{headless:true});
 assert.equal(args.includes('--headless=new'),true);assert.equal(args.filter(x=>x.startsWith('https://')).length,1);
 assert.equal(browserArgs('C:\\profile').filter(x=>x.startsWith('https://')).length,2);
});
