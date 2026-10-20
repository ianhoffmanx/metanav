import test from 'node:test';
import { mkdtempSync, writeFileSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';
import { graphApp, parseEnv, eventOf, calendarWindow, deviceLogin, accessToken, calendar, graphInputs, fileStore } from '../graph.mjs';

const app = { tenant: 't-1', id: 'c-1', secret: 'shh-secret' };
const memory = (initial = null) => { let v = initial; return { load: () => v, save: x => { v = x; }, get: () => v }; };
const reply = (status, body) => ({ ok: status < 400, status, json: async () => body });
const idToken = claims => 'x.' + Buffer.from(JSON.stringify(claims)).toString('base64url') + '.y';

test('the .env is read by name; quotes, comments and blanks are handled and nothing else leaks in', () => {
  const env = parseEnv('# graph\nMSGRAPH_TENANT_ID=abc\n\nMSGRAPH_CLIENT_SECRET="s e c"\n  MSGRAPH_CLIENT_ID = id-1 \nbad line\n');
  assert.deepEqual(env, { MSGRAPH_TENANT_ID: 'abc', MSGRAPH_CLIENT_SECRET: 's e c', MSGRAPH_CLIENT_ID: 'id-1' });
});

test('a calendar entry becomes wall-clock times, all-day is a date, and body text is short and one line', () => {
  const e = eventOf({ id: 'a', subject: 'Penny birthday morning tea', isAllDay: false, isCancelled: false, start: { dateTime: '2026-10-21T10:30:00.0000000' },
    end: { dateTime: '2026-10-21T11:00:00.0000000' }, location: { displayName: 'Brisbane kitchen' }, organizer: { emailAddress: { name: 'Penny' } },
    bodyPreview: 'Cake\r\n  and tea ' + 'x'.repeat(300), responseStatus: { response: 'accepted' }, webLink: 'https://outlook.office.com/calendar/item/a' });
  assert.equal(e.start, '2026-10-21T10:30:00');
  assert.equal(e.location, 'Brisbane kitchen');
  assert.ok(e.preview.startsWith('Cake and tea') && e.preview.length === 200);
  assert.equal(eventOf({ isAllDay: true, start: { dateTime: '2026-10-22T00:00:00.0000000' }, end: { dateTime: '2026-10-23T00:00:00.0000000' } }).start, '2026-10-22');
  assert.equal(eventOf({}).subject, '(no subject)');
});

test('the window is the last 2 days to the end of the 7th day ahead', () => {
  const { from, to } = calendarWindow(new Date(2026, 9, 20, 15, 0));
  assert.deepEqual([from.getDate(), from.getHours(), to.getDate(), to.getHours()], [18, 0, 28, 0]);
});

test('device login shows the code, waits through pending, and keeps only the refresh token and who signed in', async () => {
  const store = memory(), shown = [], calls = [];
  const responses = [
    reply(200, { device_code: 'dc', user_code: 'ABC123', verification_uri: 'https://login.microsoft.com/device', expires_in: 900, interval: 1 }),
    reply(400, { error: 'authorization_pending' }),
    reply(400, { error: 'slow_down' }),
    reply(200, { access_token: 'AT', refresh_token: 'RT1', scope: 'Calendars.Read', id_token: idToken({ preferred_username: 'ian@ssw.com.au' }) }),
  ];
  const done = await deviceLogin({ ...app, secret: undefined }, store, { onCode: c => shown.push(c), sleep: async () => {}, fetchImpl: async (u, o) => { calls.push([u, String(o.body)]); return responses.shift(); } });
  assert.equal(shown[0].user_code, 'ABC123');
  assert.deepEqual(done, { account: 'ian@ssw.com.au', scope: 'Calendars.Read' });
  assert.equal(store.get().refresh_token, 'RT1');
  assert.ok(!('access_token' in store.get()), 'only the refresh token is kept');
  assert.ok(calls.every(([, body]) => !body.includes('shh-secret') && !body.includes('client_secret')), 'a public app is sent no secret');
});

test('a refusal ends the sign-in with Microsoft\'s reason, and an expired code says so', async () => {
  const refused = [reply(200, { device_code: 'dc', user_code: 'A', verification_uri: 'u', expires_in: 900, interval: 1 }),
    reply(400, { error: 'access_denied', error_description: 'AADSTS65004: The user declined.\r\nTrace ID: x' })];
  await assert.rejects(deviceLogin({...app,secret:undefined}, memory(), { sleep: async () => {}, fetchImpl: async () => refused.shift() }), /declined/);
  const none = [reply(400, { error: 'invalid_request', error_description: 'AADSTS70002: nope' })];
  await assert.rejects(deviceLogin({...app,secret:undefined}, memory(), { fetchImpl: async () => none.shift() }), /would not start/);
});

test('an app with a secret presents it from the first request; a public app is retried without it; a rotated refresh token is saved', async () => {
  const store = memory({ refresh_token: 'RT1' }), bodies = [];
  const responses = [reply(400, { error: 'invalid_client', error_description: 'AADSTS700025: Client is public so neither client_assertion nor client_secret should be presented.' }), reply(200, { access_token: 'AT2', refresh_token: 'RT2' })];
  const token = await accessToken(app, store, { fetchImpl: async (u, o) => { bodies.push(String(o.body)); return responses.shift(); } });
  assert.equal(token, 'AT2');
  assert.equal(store.get().refresh_token, 'RT2');
  assert.ok(bodies[0].includes('client_secret') && !bodies[1].includes('client_secret'));
});

test('a confidential app that still fails reports which attempts were made, and never the secret', async () => {
  const sent = [];
  await assert.rejects(accessToken(app, memory({ refresh_token: 'RT' }), { fetchImpl: async (u, o) => { sent.push(String(o.body)); return reply(400, { error: 'invalid_client', error_description: 'AADSTS7000215: Invalid client secret provided.' }); } }),
    e => /with secret\]/.test(e.message) && !/shh-secret/.test(e.message));
  assert.equal(sent.length, 1, 'a wrong-secret error is not retried without the secret');
});

test('no sign-in or an expired one is a plain instruction, never a crash', async () => {
  await assert.rejects(accessToken(app, memory(), {}), /node graph\.mjs login/);
  await assert.rejects(accessToken(app, memory({ refresh_token: 'old' }), { fetchImpl: async () => reply(400, { error: 'invalid_grant', error_description: 'AADSTS70000: expired' }) }), /expired or was refused/);
});

test('the calendar read pages, asks for the user\'s zone, and reports a denial without the body', async () => {
  const store = memory({ refresh_token: 'RT' });
  const seen = [];
  const pages = [reply(200, { access_token: 'AT' }),
    reply(200, { value: [{ id: '1', subject: 'A', start: { dateTime: '2026-10-21T09:00:00.0' }, end: { dateTime: '2026-10-21T09:30:00.0' } }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next' }),
    reply(200, { value: [{ id: '2', subject: 'B', isAllDay: true, start: { dateTime: '2026-10-22T00:00:00.0' }, end: { dateTime: '2026-10-23T00:00:00.0' } }] })];
  const { from, to } = calendarWindow(new Date(2026, 9, 20));
  const events = await calendar(app, store, { from, to, timezone: 'Australia/Brisbane', fetchImpl: async (u, o) => { seen.push([u, o.headers?.Prefer]); return pages.shift(); } });
  assert.deepEqual(events.map(e => [e.subject, e.start]), [['A', '2026-10-21T09:00:00'], ['B', '2026-10-22']]);
  assert.match(seen[1][1], /Australia\/Brisbane/);
  assert.match(seen[1][0], /calendarView/);
  const denied = [reply(200, { access_token: 'AT' }), reply(403, { error: { code: 'ErrorAccessDenied', message: 'secret details' } })];
  await assert.rejects(calendar(app, store, { from, to, fetchImpl: async () => denied.shift() }), e => /HTTP 403 ErrorAccessDenied/.test(e.message) && !/secret details/.test(e.message));
});

test('the collector input says "not signed in" quietly and turns any failure into data', async () => {
  assert.deepEqual(await graphInputs({ provider: 'legacy', app, store: memory() }), { ok: false, skipped: true, note: 'Graph is not signed in' });
  const bad = await graphInputs({ provider: 'legacy', app, store: memory({ refresh_token: 'RT' }), fetchImpl: async () => { throw new Error('network down'); } });
  assert.equal(bad.ok, false); assert.match(bad.error, /network down/);
  const good = await graphInputs({ provider: 'legacy', app, store: memory({ refresh_token: 'RT' }), now: new Date(2026, 9, 20),
    fetchImpl: (r => async () => r.shift())([reply(200, { access_token: 'AT' }), reply(200, { value: [] })]) });
  assert.equal(good.ok, true); assert.deepEqual(good.calendar, []); assert.ok(good.timezone);
});

test('the token file store keeps the token encrypted on disk and reads it back', () => {
  const dir = mkdtempSync(join(tmpdir(), 'metanav-graph-')), file = join(dir, 'nested', 'graph-token.dat');
  try {
    const store = fileStore(file, { protect: t => 'ENC:' + Buffer.from(t).toString('base64'), unprotect: t => Buffer.from(t.slice(4), 'base64').toString() });
    assert.equal(store.load(), null);
    store.save({ refresh_token: 'RT-secret-value', account: 'ian@ssw.com.au' });
    assert.ok(!readFileSync(file, 'utf8').includes('RT-secret-value'), 'the token is not readable in the file');
    assert.deepEqual(store.load(), { refresh_token: 'RT-secret-value', account: 'ian@ssw.com.au' });
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test('the Mail Reader app is used; the SharePoint app never is; older .env names still work', () => {
  const dir = mkdtempSync(join(tmpdir(), 'metanav-env-')), file = join(dir, '.env');
  try {
    writeFileSync(file, 'SHAREPOINT_MSGRAPH_TENANT_ID=st\nSHAREPOINT_MSGRAPH_CLIENT_ID=sp\nSHAREPOINT_MSGRAPH_CLIENT_SECRET=sp-secret\n'
      + 'READER_MSGRAPH_TENANT_ID=t\nREADER_MSGRAPH_APP_ID=reader\nREADER_MSGRAPH_CLIENT_SECRET=reader-secret\n');
    assert.deepEqual(graphApp(file), { tenant: 't', id: 'reader', secret: 'reader-secret' });
    writeFileSync(file, 'SHAREPOINT_MSGRAPH_TENANT_ID=st\nSHAREPOINT_MSGRAPH_CLIENT_ID=sp\nSHAREPOINT_MSGRAPH_CLIENT_SECRET=sp-secret\n');
    assert.throws(() => graphApp(file), /not configured/);   // only the SharePoint app: nothing usable, and it is not picked up
    writeFileSync(file, 'MSGRAPH_TENANT_ID=t\nMSGRAPH_CLIENT_ID=old\nMSGRAPH_CLIENT_SECRET=old-secret\nMSGRAPH_APP_ID=m365\n');
    assert.deepEqual(graphApp(file), { tenant: 't', id: 'm365', secret: undefined });
    writeFileSync(file, 'MSGRAPH_TENANT_ID=t\nMSGRAPH_CLIENT_ID=old\nMSGRAPH_CLIENT_SECRET=old-secret\n');
    assert.deepEqual(graphApp(file), { tenant: 't', id: 'old', secret: 'old-secret' });
  } finally { rmSync(dir, { recursive: true, force: true }); }
});
