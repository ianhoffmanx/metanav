import test from 'node:test';
import assert from 'node:assert/strict';
import {runSummary} from '../run-summary.mjs';
test('the run summary reports current collected source facts, independently of old AI failure text',()=>{
 const brief={headline:'Prepare the blog\nfor approval.',queue:[{}],waiting:[],highlights:[],done:[]};
 const summary=runSummary(brief,[{name:'Outlook',ok:true},{name:'Teams',ok:false}]);
 assert.match(summary,/Outlook OK; Teams FAILED/);assert.equal(summary.includes('\n'),false);
});
