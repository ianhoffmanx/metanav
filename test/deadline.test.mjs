import test from 'node:test';
import assert from 'node:assert/strict';
import { withDeadline } from '../deadline.mjs';
test('an unresponsive browser evaluation cannot keep a collection stage open forever',async()=>{
 await assert.rejects(withDeadline(new Promise(()=>{}),15,'Teams cache'),/Teams cache timed out/);
 assert.equal(await withDeadline(Promise.resolve('captured'),15,'Outlook'),'captured');
});
