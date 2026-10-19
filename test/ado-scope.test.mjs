import test from 'node:test';
import assert from 'node:assert/strict';
import { projectScope, iterationScope } from '../ado-scope.mjs';
test('marketing project and sprint constrain assigned and mentioned items',()=>{
 assert.equal(projectScope('SSW.Marketing')," AND [System.TeamProject] = 'SSW.Marketing'");
 assert.match(iterationScope([{path:'SSW.Marketing\\Sprint 348'}]),/SSW.Marketing\\Sprint 348/);
 assert.equal(projectScope("Client's project")," AND [System.TeamProject] = 'Client''s project'");
});
