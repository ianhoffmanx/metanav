// Meta-Nav's background service, started at login by launchd (macOS) or Task Scheduler (Windows):
// - the hourly run: it starts run.mjs on the hour (run.mjs itself keeps to weekdays within config.hours);
// - the click store: the panel's "done", "resolved" and "Got it" clicks, kept in one file that every browser and every
//   run share. A page opened from disk can read files beside it but can't write any, so this service writes for it.
// It listens on 127.0.0.1 only.
//   POST /state   {"key": "tower:dismissed" | "tower:read", "value": {...}}   -> the whole state, as JSON
//   POST /signin  {"source": "Azure DevOps" | "Outlook" | "Teams"}          -> {"status": "started" | "busy"}
//                 The banner's SIGN IN button: a window to sign in, then a refresh.
//   POST /refresh {}                                                         -> {"status": "started" | "busy"}
//                 The panel's SYNC NOW button: a run now, whatever the hour (busy if one is running, or a sign-in window is open).
// Writes <output dir>/state.json (read by each run, JUDGE.md step 5) and state.js (the same, for the page's <script> tag).
import { spawn } from 'node:child_process';
import { completionHook } from './completion-hook.mjs';
import { existsSync, mkdirSync, readFileSync, readdirSync, renameSync, rmSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { appOf, devtoolsEndpoint, recognizedApps } from './microsoft.mjs';
import { dirname, join } from 'node:path';
import { OUT, PROFILE, SKILL, config, chromePath, tool } from './common.mjs';
import { recoverTargets } from './cdp.mjs';
import { freshSignin } from './health.mjs';
import { browserArgs } from './browser-args.mjs';
import { graphSigninSource } from './graph-mail.mjs';
import { runSDK } from './graph-sdk.mjs';
import { fileStore, graphApp } from './graph.mjs';
import { replyDrafts } from './reply-drafts.mjs';
import { readerLogin, readerReply } from './reader-auth.mjs';

const PORT = 47615;
const KEYS = ['tower:dismissed', 'tower:read'];
const STATE = join(OUT, 'state.json');
const signing = new Set();   // what's being signed in to right now, so a second click doesn't open a second window
const sleep = ms => new Promise(r => setTimeout(r, ms));
const exportCompleted = completionHook(config.exporter?.script);
async function draftAPI(plan){
  return readerReply(graphApp(),fileStore(join(dirname(PROFILE),'graph-drafts-token.dat')),plan);
}
const createReply=replyDrafts({directory:join(OUT,'draft-receipts'),create:draftAPI,loadRun:()=>{
  const files=readdirSync(join(OUT,'runs')).filter(f=>/^\d{8}-\d{4}\.json$/.test(f)).sort();
  return files.length?JSON.parse(readFileSync(join(OUT,'runs',files.at(-1)),'utf8')):{};
}});

function load() {
  try { return JSON.parse(readFileSync(STATE, 'utf8')); } catch { return { rev: 0 }; }
}
// write-then-rename, so a run or a page never reads half a file
function write(file, text) {
  writeFileSync(`${file}.tmp`, text);
  renameSync(`${file}.tmp`, file);
}
function save(state) {
  write(STATE, JSON.stringify(state, null, 1));
  write(join(OUT, 'state.js'), `window.towerState = ${JSON.stringify(state).replace(/<\//g, '<\\/')};\n`);
}

// A run. `manual`: one the user asked for (SYNC NOW, `metanav`, a sign-in that worked) - run.mjs skips the hours for it.
let activeRun = null;
function start(manual) {
  if (activeRun || existsSync(join(OUT, '.running'))) return;
  if (manual) writeFileSync(join(OUT, '.manual'), '');
  activeRun = spawn(process.execPath, [join(SKILL, 'run.mjs')], { detached: true, stdio: 'ignore', windowsHide: true });
  activeRun.once('error', () => { activeRun = null; });
  activeRun.once('exit', () => { activeRun = null; exportCompleted(); });
  activeRun.unref();
}
const refresh = () => start(true);

// On the hour, every hour. A timer looks twice a minute, so after the computer wakes up, the hour it slept through runs
// straight away. Not when the service starts: logging in isn't the hour turning.
const hourOf = d => `${d.toDateString()} ${d.getHours()}`;
let lastHour = hourOf(new Date());
setInterval(() => {
  const h = hourOf(new Date());
  if (h === lastHour) return;
  lastHour = h;
  // one run at a time (run.mjs locks too), and not while a sign-in window holds the browser profile
  if (!existsSync(join(OUT, '.running')) && !signing.has('microsoft')) start(false);
}, 30e3);


// A visible window on the runs' own profile, at Outlook and Teams: a new profile may have to pick the account for
// Teams once. Keep regular Chrome open for collection to reuse its authenticated tabs and proxy origins.
async function signInMicrosoft() {
  const browserApps = config.graph?.provider === 'microsoft-sdk' && config.graph?.mail !== false ? ['teams'] : ['outlook','teams'];
  const chrome = chromePath();
  if (!chrome) throw new Error('Google Chrome was not found');
  mkdirSync(PROFILE, { recursive: true });
  let endpoint = devtoolsEndpoint(PROFILE);
  let pages = [];
  if (endpoint) {
    try { pages = await (await fetch(endpoint + '/json', { signal: AbortSignal.timeout(3000) })).json(); }
    catch { endpoint = ''; }
  }
  const report = (status, apps = {}) => write(join(OUT, 'signin-status.json'), JSON.stringify({ status, apps, at: new Date().toISOString() }));
  report('waiting', recognizedApps(pages));
  if (endpoint) {
    await recoverTargets(endpoint, { application: appOf });
    // Reuse the live profile. Deleting its port file or trusting a second Chrome
    // launch's exit loses the browser that actually owns the signed-in session.
    for (const [app, url] of [['outlook', 'https://outlook.office.com/mail/inbox'], ['teams', 'https://teams.microsoft.com/v2/']].filter(([app])=>browserApps.includes(app))) {
      if (!pages.some(p => p.type === 'page' && appOf(p.url) === app)) {
        await fetch(endpoint + '/json/new?' + encodeURI(url), { method: 'PUT', signal: AbortSignal.timeout(5000) });
      }
    }
  } else {
    const child = spawn(chrome, browserArgs(PROFILE,{apps:browserApps}), { detached: true, stdio: 'ignore' });
    child.on('error', e => report('error', { message: e.message }));
    child.unref();
  }
  for (const deadline = Date.now() + 600e3; Date.now() < deadline;) {
    await sleep(2000);
    endpoint = devtoolsEndpoint(PROFILE);
    if (!endpoint) continue;
    try {
      pages = await (await fetch(endpoint + '/json', { signal: AbortSignal.timeout(3000) })).json();
      const apps = recognizedApps(pages);
      report('waiting', apps);
      if (browserApps.every(app=>apps[app])) { report('checking', apps); return true; }
    } catch {}
  }
  report('timed_out', recognizedApps(pages));
  // Authentication is interactive. Leave the window open instead of destroying
  // the user's session; another SIGN IN resumes watching the same browser.
  return false;
}

// az login opens the browser itself. BROWSER=open on macOS: its default way is AppleScript, which macOS stops when it
// comes from a background job - `open` always works.
function signInAzure() {
  const env = { ...process.env, ...(process.platform === 'darwin' && { BROWSER: 'open' }) };
  const tenant = (tool('az', ['account', 'show', '--query', 'tenantId', '-o', 'tsv'], { env }).stdout || '').trim();
  return tool('az', ['login', '--output', 'none', ...(tenant ? ['--tenant', tenant] : [])], { env, timeout: 600e3 }).status === 0;
}

async function signInGraph() {
  const result=await runSDK('login',{tenant:graphApp().tenant,login_hint:config.microsoft_account});
  if(!result.ok)throw Error(result.note||'Microsoft Graph sign-in did not complete.');
  return true;
}
function signIn(source) {
  const what = source === 'Mail drafts' ? 'drafts' : source === 'Azure DevOps' ? 'azure' : 'microsoft';
  // a run holds the browser profile while it reads Outlook and Teams
  if (signing.has(what) || (what === 'microsoft' && existsSync(join(OUT, '.running')))) return 'busy';
  signing.add(what);
  (async () => {
    try { if (await (source==='Mail drafts' ? readerLogin(graphApp(),fileStore(join(dirname(PROFILE),'graph-drafts-token.dat')),config.microsoft_account) : what === 'azure' ? signInAzure() : graphSigninSource(source,config) ? signInGraph() : signInMicrosoft())) refresh(); }
    catch (e) { write(join(OUT, source==='Mail drafts'?'draft-signin-status.json':'signin-status.json'), JSON.stringify({ status: 'error', message: source==='Mail drafts'?'Reply draft sign-in did not complete.':String(e.message || e), at: new Date().toISOString() })); }
    finally { signing.delete(what); }
  })();
  return 'started';
}

function reply(res, status, obj) {
  const data = obj === undefined ? '' : JSON.stringify(obj);
  res.writeHead(status, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': 'null', 'Content-Length': Buffer.byteLength(data) });
  res.end(data);
}

const server = createServer((req, res) => {
  // Only the panel itself may call in: a page on disk sends Origin "null"; a website open in the same browser sends its
  // own origin. The Host check keeps out DNS-rebinding tricks.
  if (req.method !== 'POST' || !['/state', '/signin', '/refresh', '/status','/mail/reply-draft','/mail/draft-status'].includes(req.url) || (req.headers.origin ?? 'null') !== 'null'
    || ![`127.0.0.1:${PORT}`, `localhost:${PORT}`].includes(req.headers.host)) return reply(res, 403);
  let raw = '';
  req.on('data', c => { raw += c; if (raw.length > 1 << 22) req.destroy(); });
  req.on('end', async () => {
    let body;
    try {
      body = JSON.parse(raw || '{}');
      if (req.url === '/signin' && !['Azure DevOps', 'Outlook', 'Teams', 'Calendar','Mail drafts'].includes(body.source)) throw 0;
      if (req.url === '/state' && (!KEYS.includes(body.key) || typeof body.value !== 'object' || !body.value || Array.isArray(body.value))) throw 0;
    } catch { return reply(res, 400); }
    if(req.url==='/mail/reply-draft'||req.url==='/mail/draft-status'){
      try{return reply(res,200,await (req.url==='/mail/reply-draft'?createReply(body):draftAPI(null)))}
      catch{return reply(res,200,{status:'not_configured'})}
    }
    if (req.url === '/signin') return reply(res, 200, { status: signIn(body.source) });
    if (req.url === '/status') {
      let signin = null;
      try { signin = JSON.parse(readFileSync(join(OUT, 'signin-status.json'), 'utf8')); } catch {}
      return reply(res, 200, { running: !!activeRun || existsSync(join(OUT, '.running')), signing: [...signing], signin: freshSignin(signin) });
    }
    if (req.url === '/refresh') {
      // one run at a time, and not while a sign-in window holds the browser profile
      const busy = !!activeRun || existsSync(join(OUT, '.running')) || signing.has('microsoft');
      if (!busy) refresh();
      return reply(res, 200, { status: busy ? 'busy' : 'started' });
    }
    const state = load();
    state[body.key] = body.value;
    state.rev = (state.rev || 0) + 1;   // lets an open page tell that another browser changed something
    save(state);
    exportCompleted();
    reply(res, 200, state);
  });
});

mkdirSync(OUT, { recursive: true });
save(load());   // state.js always exists once the service is up, so the page knows to use it
server.listen(PORT, '127.0.0.1');
