// Microsoft sessions may live behind a tenant's Conditional Access proxy.
// Keep each application on its authenticated origin and in its own browser tab.
import { readFileSync, mkdirSync } from 'node:fs';
import { spawn } from 'node:child_process';
import { join } from 'node:path';

const HOSTS = {
  outlook: ['outlook.office.com', 'outlook.office365.com', 'outlook.cloud.microsoft', 'outlook.cloud.microsoft.mcas.ms'],
  teams: ['teams.microsoft.com', 'teams.cloud.microsoft', 'teams.cloud.microsoft.mcas.ms'],
};
const requested = new WeakMap();
let savedAccount = '';
export async function continueSavedAccount(page, account = savedAccount) {
  if (!/^[\w.+-]+@[\w.-]+$/.test(account)) return false;
  try {
    if (new URL(page.url()).hostname !== 'login.microsoftonline.com') return false;
    const tile = page.locator(`[role="button"][data-test-id="${account}"]`);
    if (await tile.count() !== 1) return false;
    await tile.click({ timeout: 2000 });
    return true;
  } catch { return false; }
}
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
export function appOf(value) {
  try {
    const url = new URL(value);
    if (url.protocol !== 'https:') return null;
    return Object.keys(HOSTS).find(app => HOSTS[app].includes(url.hostname)) || null;
  } catch { return null; }
}
export function sessionUrl(current, target) {
  const app = appOf(target);
  if (app && appOf(current) === app) {
    const url = new URL(target);
    url.host = new URL(current).host;
    return url.href;
  }
  return target;
}
export function devtoolsEndpoint(profile) {
  try {
    const port = readFileSync(join(profile, 'DevToolsActivePort'), 'utf8').split(/\r?\n/)[0].trim();
    return /^\d+$/.test(port) ? `http://127.0.0.1:${port}` : '';
  } catch { return ''; }
}
export async function connectMicrosoftBrowser(chromium, profile, executable, account = '') {
  savedAccount = account;
  let endpoint = devtoolsEndpoint(profile);
  if (endpoint) {
    try {
      const browser = await chromium.connectOverCDP(endpoint, { timeout: 3000 });
      return { browser, context: browser.contexts()[0] };
    } catch {}
  }
  if (!executable) throw new Error('Google Chrome was not found');
  mkdirSync(profile, { recursive: true });
  // Use regular Chrome for Conditional Access, even when no sign-in window is
  // already open. Automation's headless launch can invalidate an otherwise valid session.
  let launchError;
  const child = spawn(executable, [`--user-data-dir=${profile}`, '--remote-debugging-port=0', '--no-first-run', '--no-default-browser-check',
    'https://outlook.office.com/mail/inbox', 'https://teams.microsoft.com/v2/'], { detached: true, stdio: 'ignore' });
  child.on('error', error => { launchError = error; });
  child.unref();
  const deadline = Date.now() + 30000;
  do {
    if (launchError) throw launchError;
    endpoint = devtoolsEndpoint(profile);
    if (endpoint) {
      try {
        const browser = await chromium.connectOverCDP(endpoint, { timeout: 1000 });
        const context = browser.contexts()[0];
        // Allow the initial cookie/Conditional Access redirects to settle.
        const readyBy = Date.now() + 20000;
        while (Date.now() < readyBy) {
          for (const page of context.pages()) await continueSavedAccount(page);
          const pages = await Promise.all(context.pages().map(async p => ({ url: p.url(), title: await p.title().catch(() => '') })));
          const apps = recognizedApps(pages);
          if (apps.outlook && apps.teams) break;
          await sleep(500);
        }
        return { browser, context };
      } catch {}
    }
    await sleep(500);
  } while (Date.now() < deadline);
  throw new Error('Could not attach to Meta-Nav Chrome. Close only its dedicated window, then use SIGN IN to reopen it.');
}
export async function pageForApp(context, app) {
  const page = context.pages().find(page => !page.isClosed() && appOf(page.url()) === app) || await context.newPage();
  requested.set(page, app);
  return page;
}
export async function signedIn(page, timeout = 20000) {
  const app = requested.get(page) || appOf(page.url());
  if (!app) return false;
  const deadline = Date.now() + timeout;
  do {
    if (page.isClosed()) return false;
    if (appOf(page.url()) === app) return true;
    await continueSavedAccount(page);
    await sleep(250);
  } while (Date.now() < deadline);
  return false;
}
export async function openApp(page, target) {
  const app = appOf(target);
  if (!app) throw new Error('Unsupported Microsoft application URL');
  requested.set(page, app);
  const url = sessionUrl(page.url(), target);
  // Teams' proxy lands on "/" even though its public entry point is "/v2/".
  // Its authenticated cache is already available there; reloading adds an auth round trip.
  if (app === 'teams' && appOf(page.url()) === 'teams') return signedIn(page);
  // An already loaded application tab has the live proxy session. Avoid a fresh auth round trip.
  if (page.url().replace(/\/$/, '') !== url.replace(/\/$/, '')) {
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 45000 });
  }
  return signedIn(page);
}
export function recognizedApps(pages) {
  return {
    outlook: pages.some(p => appOf(p.url) === 'outlook' && /Outlook/i.test(p.title || '') && !/sign in/i.test(p.title || '')),
    teams: pages.some(p => appOf(p.url) === 'teams' && /\| Microsoft Teams$/.test(p.title || '')),
  };
}
export async function evaluateStable(page, expression, arg) {
  for (let attempt = 0; ; attempt++) {
    try { return arg === undefined ? await page.evaluate(expression) : await page.evaluate(expression, arg); }
    catch (e) {
      if (attempt >= 2 || !/execution context was destroyed|frame was detached|navigation/i.test(String(e.message || e))) throw e;
      await page.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => {});
      if (!await signedIn(page)) throw new Error('Microsoft session expired during collection');
    }
  }
}

