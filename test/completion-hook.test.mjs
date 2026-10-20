import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { test } from 'node:test';
import { completionHook } from '../completion-hook.mjs';

test('no exporter configured leaves the service unchanged', () => {
  completionHook(null, { launch() { assert.fail('must not launch'); } })();
});

test('completion and clicks coalesce; changed state is sent after the active send', () => {
  const children = [], calls = [];
  const trigger = completionHook('C:\\a path\\export.mjs', { launch(...args) {
    calls.push(args); const c = new EventEmitter(); c.unref = () => {}; children.push(c); return c;
  }});
  trigger(); trigger(); trigger();
  assert.equal(children.length, 1);
  assert.deepEqual(calls[0][1], ['C:\\a path\\export.mjs']);
  assert.equal(calls[0][2].windowsHide, true);
  children[0].emit('exit', 0);
  assert.equal(children.length, 2);
  children[1].emit('exit', 0);
  assert.equal(children.length, 2);
});

test('a failed exporter can be retried by the next completion', () => {
  const children = [];
  const trigger = completionHook('export.mjs', { launch() {
    const c = new EventEmitter(); children.push(c); return c;
  }});
  trigger(); children[0].emit('error', Error('missing'));
  children[0].emit('exit', 1); // error + exit only completes once
  trigger(); assert.equal(children.length, 2); children[1].emit('exit', 0);
});
