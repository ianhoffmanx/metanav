#!/usr/bin/env node
// Direct Microsoft Graph access as Ian, with read-only delegated permissions.
//   node graph.mjs login     explicit browser sign-in; never started by background collection
//   node graph.mjs status    silently check the saved sign-in
//   node graph.mjs calendar  read the collector calendar window, reporting counts only for the SDK
// graph.provider=microsoft-sdk uses Microsoft's approved public client and encrypted SDK cache.
// Only the tenant is read from the existing .env; the Reader app secret is never sent to the SDK.
// The legacy provider below remains available for correctly configured public-client registrations.
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { PROFILE, SKILL, WIN, config } from './common.mjs';
import { runSDK } from './graph-sdk.mjs';
import { mailFolders, mailInput } from './graph-mail.mjs';

export const SCOPES = 'Calendars.Read offline_access openid profile';
const CACHE = dirname(PROFILE);
export const TOKEN_FILE = join(CACHE, 'graph-token.dat');
export const LOGIN_FILE = join(CACHE, 'graph-login.json');
const GRAPH = 'https://graph.microsoft.com/v1.0';

// KEY=value lines, optional quotes; comments and blanks ignored
export function parseEnv(text) {
  const out = {};
  for (const line of String(text).split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Za-z_]\w*)\s*=\s*(.*?)\s*$/);
    if (m && !line.trim().startsWith('#')) out[m[1]] = m[2].replace(/^(["'])(.*)\1$/, '$2');
  }
  return out;
}

export function envFile() {
  const candidates = [config.graph?.env_file, join(SKILL, '.env'), join(homedir(), 'projects', 'metanav', '.env')].filter(Boolean);
  return candidates.find(existsSync) || null;
}

export function graphApp(file = envFile()) {
  const env = file ? parseEnv(readFileSync(file, 'utf8')) : {};
  // The Mail Reader app (READER_MSGRAPH_*) is the one for Microsoft 365 data. The SharePoint app (SHAREPOINT_MSGRAPH_*) is never used.
  // Older .env files named it MSGRAPH_APP_ID (with MSGRAPH_APP_SECRET), or used MSGRAPH_CLIENT_ID / _SECRET; those still work.
  const app = env.READER_MSGRAPH_APP_ID
    ? { tenant: env.READER_MSGRAPH_TENANT_ID, id: env.READER_MSGRAPH_APP_ID, secret: env.READER_MSGRAPH_CLIENT_SECRET }
    : env.MSGRAPH_APP_ID && env.MSGRAPH_APP_ID !== env.MSGRAPH_CLIENT_ID
      ? { tenant: env.MSGRAPH_TENANT_ID, id: env.MSGRAPH_APP_ID, secret: env.MSGRAPH_APP_SECRET }
      : { tenant: env.MSGRAPH_TENANT_ID, id: env.MSGRAPH_CLIENT_ID, secret: env.MSGRAPH_CLIENT_SECRET };
  if (!app.tenant || !app.id) throw new Error('Graph is not configured: READER_MSGRAPH_APP_ID and READER_MSGRAPH_TENANT_ID are missing from the .env');
  return app;
}

// Windows DPAPI, for the current user: only this account on this laptop can read the token back
function dpapi(mode, text) {
  if (!WIN) throw new Error('Token storage needs Windows (DPAPI)');
  const body = mode === 'protect'
    ? '[Convert]::ToBase64String([Security.Cryptography.ProtectedData]::Protect([Text.Encoding]::UTF8.GetBytes($t),$null,"CurrentUser"))'
    : '[Text.Encoding]::UTF8.GetString([Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String($t.Trim()),$null,"CurrentUser"))';
  const r = spawnSync('powershell', ['-NoProfile', '-NonInteractive', '-Command',
    `$ProgressPreference='SilentlyContinue';Add-Type -AssemblyName System.Security;$t=[Console]::In.ReadToEnd();${body}`],
  { input: text, encoding: 'utf8', windowsHide: true, timeout: 20000 });
  if (r.status !== 0) throw new Error('Could not ' + mode + ' the Graph token');
  return r.stdout;
}
export const fileStore = (file = TOKEN_FILE, { protect = t => dpapi('protect', t), unprotect = t => dpapi('unprotect', t) } = {}) => ({
  load() { try { return JSON.parse(unprotect(readFileSync(file, 'utf8'))); } catch { return null; } },
  save(value) { mkdirSync(dirname(file), { recursive: true }); writeFileSync(file, protect(JSON.stringify(value))); },
});

const form = o => new URLSearchParams(o);
async function post(url, body, fetchImpl) {
  const r = await fetchImpl(url, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: form(body), signal: AbortSignal.timeout(20000) });
  return { status: r.status, json: await r.json().catch(() => ({})) };
}
const failure = j => `${j.error || 'error'}: ${String(j.error_description || '').split(/\r?\n/)[0].replace(/\s*Trace ID.*$/, '')}`;
const accountOf = idToken => { try { const c = JSON.parse(Buffer.from(idToken.split('.')[1], 'base64url').toString()); return c.preferred_username || c.upn || c.email || null; } catch { return null; } };

// A token request for this app. A confidential app (one with a secret) must present it; a public one must not (AADSTS700025).
// So: with the secret first when there is one, and without it only if Microsoft says the app is public - never a probe that
// could use up a one-shot device code. The returned error names which attempts were made, never the secret.
async function tokenRequest(app, params, fetchImpl) {
  const url = `https://login.microsoftonline.com/${app.tenant}/oauth2/v2.0/token`;
  const attempts = app.secret ? [true, false] : [false];
  let r, tried = [];
  for (const withSecret of attempts) {
    tried.push(withSecret ? 'with secret' : 'without secret');
    r = await post(url, { client_id: app.id, ...(withSecret && { client_secret: app.secret }), ...params }, fetchImpl);
    const text = `${r.json.error} ${r.json.error_description}`;
    const retry = withSecret ? /AADSTS700025/.test(text) : false;   // a public app: try again without the secret
    if (r.json.access_token || !retry) break;
  }
  if (r.json.error) r.json.error_description = `${String(r.json.error_description || '').split(/\r?\n/)[0]} [${tried.join(', ')}]`;
  return r;
}

// Device-code sign-in: onCode gets the code to show the person, then this waits (up to the code's lifetime) for them to finish
export async function deviceLogin(app, store, { onCode = () => {}, fetchImpl = fetch, sleep = ms => new Promise(r => setTimeout(r, ms)), publicClient = false } = {}) {
  if (app.secret && !publicClient) throw new Error('Device-code sign-in requires a public-client app. Use the configured Microsoft SDK sign-in; the Reader secret cannot enable this flow.');
  const dc = await post(`https://login.microsoftonline.com/${app.tenant}/oauth2/v2.0/devicecode`, { client_id: app.id, scope: SCOPES }, fetchImpl);
  if (!dc.json.device_code) throw new Error('Microsoft would not start a sign-in: ' + failure(dc.json));
  onCode({ user_code: dc.json.user_code, verification_uri: dc.json.verification_uri, expires_at: new Date(Date.now() + dc.json.expires_in * 1000).toISOString() });
  let wait = (dc.json.interval || 5) * 1000;
  for (const end = Date.now() + dc.json.expires_in * 1000; Date.now() < end;) {
    await sleep(wait);
    const t = await post(`https://login.microsoftonline.com/${app.tenant}/oauth2/v2.0/token`, { client_id: app.id, grant_type: 'urn:ietf:params:oauth:grant-type:device_code', device_code: dc.json.device_code }, fetchImpl);
    if (t.json.access_token) {
      const saved = { refresh_token: t.json.refresh_token, scope: t.json.scope, account: accountOf(t.json.id_token), signed_in_at: new Date().toISOString() };
      if (!saved.refresh_token) throw new Error('Microsoft returned no refresh token (offline_access was not granted)');
      store.save(saved);
      return { account: saved.account, scope: saved.scope };
    }
    const e = t.json.error;
    if (e === 'authorization_pending') continue;
    if (e === 'slow_down') { wait += 5000; continue; }
    throw new Error('Sign-in was not completed: ' + failure(t.json));
  }
  throw new Error('The sign-in code expired before it was used');
}

// A fresh access token from the stored refresh token (which Microsoft rotates: the new one is saved)
export async function accessToken(app, store, { fetchImpl = fetch } = {}) {
  const saved = store.load();
  if (!saved?.refresh_token) throw new Error('Graph is not signed in; run: node graph.mjs login');
  const t = await tokenRequest(app, { grant_type: 'refresh_token', refresh_token: saved.refresh_token, scope: SCOPES }, fetchImpl);
  if (!t.json.access_token) throw new Error('Graph sign-in has expired or was refused (' + failure(t.json) + '); run: node graph.mjs login');
  if (t.json.refresh_token && t.json.refresh_token !== saved.refresh_token) store.save({ ...saved, refresh_token: t.json.refresh_token });
  return t.json.access_token;
}

// One calendar entry, in the fields the judge and the dashboard care about. Times are the user's wall-clock time in the
// zone the read asked for ("2026-10-21T10:00:00", no offset; an all-day event is just its date).
export function eventOf(e) {
  const at = x => (x?.dateTime ? (e.isAllDay ? x.dateTime.slice(0, 10) : x.dateTime.slice(0, 19)) : null);
  return {
    id: e.id, subject: e.subject || '(no subject)', start: at(e.start), end: at(e.end), all_day: !!e.isAllDay,
    cancelled: !!e.isCancelled, location: e.location?.displayName || null, organizer: e.organizer?.emailAddress?.name || null,
    show_as: e.showAs || null, response: e.responseStatus?.response || null, online_url: e.onlineMeeting?.joinUrl || null,
    url: e.webLink || null, preview: (e.bodyPreview || '').replace(/\s+/g, ' ').trim().slice(0, 200) || null,
  };
}

export async function calendar(app, store, { from, to, fetchImpl = fetch, timezone = Intl.DateTimeFormat().resolvedOptions().timeZone } = {}) {
  const token = await accessToken(app, store, { fetchImpl });
  const q = new URLSearchParams({ startDateTime: from.toISOString(), endDateTime: to.toISOString(), $top: '100', $orderby: 'start/dateTime',
    $select: 'id,subject,start,end,isAllDay,isCancelled,location,organizer,showAs,responseStatus,onlineMeeting,webLink,bodyPreview' });
  let url = `${GRAPH}/me/calendarView?${q}`;
  const events = [];
  for (let page = 0; url && page < 3; page++) {
    const r = await fetchImpl(url, { headers: { Authorization: 'Bearer ' + token, Prefer: `outlook.timezone="${timezone}", outlook.body-content-type="text"` }, signal: AbortSignal.timeout(20000) });
    if (!r.ok) {
      const e = await r.json().catch(() => ({}));
      throw new Error(`Graph calendar read failed (HTTP ${r.status}${e.error?.code ? ' ' + e.error.code : ''})`);
    }
    const j = await r.json();
    events.push(...j.value.map(eventOf));
    url = j['@odata.nextLink'] || null;
  }
  return events;
}

// What the collector puts in the inputs: never throws, and says plainly when Graph is not signed in or not reachable
export async function graphInputs({ now = new Date(), app, store = fileStore(), fetchImpl = fetch, provider = config.graph?.provider, sdk = runSDK, mailSince, first = false, mailConfig = config.mail, trackedConversations=[] } = {}) {
  try {
    if (provider === 'microsoft-sdk') {
      const w=calendarWindow(now), timezone=Intl.DateTimeFormat().resolvedOptions().timeZone;
      const result=await sdk('read',{tenant:(app||graphApp()).tenant,login_hint:config.microsoft_account,from:w.from.toISOString(),to:w.to.toISOString(),timezone,...(mailSince?{mail_since:mailSince,mail_folders:mailFolders(mailConfig),first,tracked_conversations:trackedConversations}:{})});
      if(result.mail)result.mail=mailInput(result.mail,{timezone});
      if(!result.ok)return {...result,provider,error:result.note||result.error};
      return {...result,timezone,from:w.from.toISOString(),to:w.to.toISOString(),calendar:result.calendar.map(eventOf)};
    }
    if (!store.load()?.refresh_token) return { ok: false, skipped: true, note: 'Graph is not signed in' };
    const w = calendarWindow(now);
    const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    const events = await calendar(app || graphApp(), store, { ...w, fetchImpl, timezone });
    return { ok: true, timezone, from: w.from.toISOString(), to: w.to.toISOString(), calendar: events };
  } catch (e) { return { ok: false, provider, error: String(e.message || e).slice(0, 200) }; }
}

// The window the collector asks for: the last 2 days to the next 7, the same as `social`
export function calendarWindow(now = new Date(), back = 2, ahead = 7) {
  const day = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return { from: new Date(day.getFullYear(), day.getMonth(), day.getDate() - back), to: new Date(day.getFullYear(), day.getMonth(), day.getDate() + ahead + 1) };
}

async function cli(cmd, arg) {
  if (!['login','status','calendar'].includes(cmd)) throw new Error('usage: node graph.mjs login | status | calendar [days]');
  const app = graphApp(), store = fileStore();
  if (config.graph?.provider === 'microsoft-sdk') {
    const w=calendarWindow();
    const result=await runSDK(cmd==='login'?'login':cmd==='status'?'status':'read',{tenant:app.tenant,login_hint:config.microsoft_account,from:w.from.toISOString(),to:w.to.toISOString(),timezone:Intl.DateTimeFormat().resolvedOptions().timeZone});
    if(!result.ok)throw new Error([result.note||'Graph SDK sign-in did not complete',result.code,result.aad_code,result.error_type].filter(Boolean).join(' | '));
    console.log(JSON.stringify({ok:true,provider:'microsoft-sdk',account:result.account,calendar_events:result.calendar?.length,mail_read_ok:result.mail_read_ok}));
    return;
  }
  if (cmd === 'login') {
    const done = await deviceLogin(app, store, { onCode: c => {
      mkdirSync(CACHE, { recursive: true });
      writeFileSync(LOGIN_FILE, JSON.stringify(c));
      console.log(`Go to ${c.verification_uri} and enter the code ${c.user_code} (valid until ${c.expires_at}).`);
    } });
    console.log(`Signed in${done.account ? ' as ' + done.account : ''}. Permissions: ${done.scope}`);
  } else if (cmd === 'status') {
    const saved = store.load();
    console.log(saved ? `Signed in${saved.account ? ' as ' + saved.account : ''} since ${saved.signed_in_at}` : 'Not signed in');
    if (saved) { const w = calendarWindow(); console.log(`Calendar read: ${(await calendar(app, store, w)).length} events in the window`); }
  } else if (cmd === 'calendar') {
    const w = calendarWindow(new Date(), 0, Number(arg) || 7);
    const events = await calendar(app, store, w);
    console.log(`${events.length} events`);
    for (const e of events) console.log(`${e.start} → ${e.end}${e.all_day ? ' (all day)' : ''}${e.cancelled ? ' [cancelled]' : ''} ${e.subject}${e.location ? ' @ ' + e.location : ''}`);
  } else console.log('usage: node graph.mjs login | status | calendar [days]');
}
if (fileURLToPath(import.meta.url) === resolve(process.argv[1] || '')) {
  cli(process.argv[2], process.argv[3]).catch(e => { console.error(String(e.message || e)); process.exit(1); });
}
