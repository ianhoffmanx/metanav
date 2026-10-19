import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { normaliseSocial, SOCIAL_KINDS, SOCIAL_WORDS } from '../social.mjs';
import { parseJudgeResult } from '../judge-output.mjs';

// "now" is a local midday; dates are built the same way the normaliser reads them (the user's local calendar)
const now = new Date(2026, 9, 20, 12, 0);
const day = n => { const d = new Date(2026, 9, 20 + n); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
const item = (id, over = {}) => ({ id, kind: 'birthday', title: "Penny's birthday", who: ['Penny'], when: day(1), where: 'Brisbane office', detail: 'Morning tea', url: 'https://outlook.office.com/mail/inbox/id/x', ...over });
const ids = xs => xs.map(x => x.id);

test('a good item comes out with exactly the shipped fields, and new is true the first time', () => {
  const [x] = normaliseSocial([item('a')], { now });
  assert.deepEqual(Object.keys(x), ['id', 'kind', 'title', 'who', 'when', 'where', 'detail', 'url', 'new']);
  assert.equal(x.new, true);
  assert.deepEqual(x.who, ['Penny']);
});

test('new is false for ids the previous run already had', () => {
  const out = normaliseSocial([item('a'), item('b')], { now, previous: [{ id: 'a' }] });
  assert.deepEqual(out.map(x => [x.id, x.new]), [['a', false], ['b', true]]);
});

test('the window is the last 2 days to the next 7 days, today included', () => {
  const out = normaliseSocial([-3, -2, 0, 7, 8].map(n => item('d' + n, { when: day(n) })), { now });
  assert.deepEqual(ids(out), ['d-2', 'd0', 'd7']);
});

test('a datetime and a null date are both kept; an unreadable date becomes null', () => {
  const out = normaliseSocial([
    item('t', { when: day(2) + 'T10:30:00+10:00' }), item('n', { when: null }), item('bad', { when: 'next Friday' })], { now });
  assert.deepEqual(out.map(x => [x.id, x.when === null]), [['t', false], ['n', true], ['bad', true]]);
});

test('malformed items are dropped, never an error: bad kind, no id, no title, not an object', () => {
  const out = normaliseSocial([item('ok'), item('k', { kind: 'party' }), item('', {}), item('t', { title: '  ' }), null, 'x', []], { now });
  assert.deepEqual(ids(out), ['ok']);
});

test('only https:// and msteams: links survive; other fields get safe defaults', () => {
  const out = normaliseSocial([
    item('a', { url: 'javascript:alert(1)' }), item('b', { url: 'msteams:/l/message/x/1' }), item('c', { url: 'http://insecure' }),
    item('d', { who: 'Sam', where: undefined, detail: '' })], { now });
  assert.equal(out.find(x => x.id === 'a').url, null);
  assert.equal(out.find(x => x.id === 'b').url, 'msteams:/l/message/x/1');
  assert.equal(out.find(x => x.id === 'c').url, null);
  const d = out.find(x => x.id === 'd');
  assert.deepEqual([d.who, d.where, d.detail], [['Sam'], null, null]);
});

test('duplicate ids keep the first, results are soonest first with undated last, and the list is capped', () => {
  const out = normaliseSocial([item('late', { when: day(5) }), item('none', { when: null }), item('soon', { when: day(0) }), item('soon', { when: day(1) })], { now });
  assert.deepEqual(ids(out), ['soon', 'late', 'none']);
  const many = normaliseSocial(Array.from({ length: 40 }, (_, i) => item('i' + i)), { now });
  assert.equal(many.length, 25);
});

test('nothing in, or something that is not a list, gives an empty array', () => {
  assert.deepEqual(normaliseSocial([], { now }), []);
  assert.deepEqual(normaliseSocial(undefined, { now }), []);
  assert.deepEqual(normaliseSocial({ id: 'a' }, { now }), []);
});

test('the judge result may omit social, but not return a non-list', () => {
  const brief = { headline: 'h', queue: [], waiting: [], calendar: [], highlights: [], done: [] };
  assert.equal(parseJudgeResult(JSON.stringify({ summary: 's', brief })).brief.social, undefined);
  assert.deepEqual(parseJudgeResult(JSON.stringify({ summary: 's', brief: { ...brief, social: [] } })).brief.social, []);
  assert.throws(() => parseJudgeResult(JSON.stringify({ summary: 's', brief: { ...brief, social: 'none' } })), /social/);
});

test('the schema asks the judge for social, with the same kinds and fields the runner ships', () => {
  const schema = JSON.parse(readFileSync(new URL('../judge-output.schema.json', import.meta.url), 'utf8'));
  const brief = schema.properties.brief;
  assert.ok(brief.required.includes('social'));
  const s = brief.properties.social;
  assert.equal(s.type, 'array');
  assert.deepEqual(s.items.properties.kind.enum, SOCIAL_KINDS);
  assert.deepEqual(s.items.required, ['id', 'kind', 'title', 'who', 'when', 'where', 'detail', 'url']);
  assert.deepEqual(Object.keys(s.items.properties), s.items.required);
  assert.equal(s.items.additionalProperties, false);
  assert.ok(!('new' in s.items.properties), 'the runner sets new, not the judge');
  for (const f of ['when', 'where', 'detail', 'url']) assert.deepEqual(s.items.properties[f].type, ['string', 'null']);
  // an item the schema accepts is one the normaliser keeps
  assert.equal(normaliseSocial([item('a')], { now }).length, 1);
});

test('the Teams channel filter and the collector share one social pattern', () => {
  const extract = readFileSync(new URL('../teams-extract.js', import.meta.url), 'utf8');
  assert.match(extract, /new RegExp\(__SOCIAL__, 'i'\)/);
  const rx = new RegExp(SOCIAL_WORDS, 'i');
  for (const yes of ["Happy Birthday Penny! 🎂", 'Welcome to the team, Deborrah', 'Drinks at 5 on Friday', 'Farewell lunch for Sam', 'Congrats on the baby!']) assert.ok(rx.test(yes), yes);
  for (const no of ['Can you review the blog PR by Thursday', 'Release 3.2 is out']) assert.ok(!rx.test(no), no);
  assert.doesNotThrow(() => new Function(`return (${extract.replace('__SOCIAL__', JSON.stringify(SOCIAL_WORDS)).replace(/__\w+__/g, '2026-01-01')})`), 'extractor still parses');
});
