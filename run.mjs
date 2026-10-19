#!/usr/bin/env node
// Refresh Meta-Nav: hours guard, lock, collect, judge, render, clean up.
// launchd (macOS) or Task Scheduler (Windows) fires it on the hour; `metanav` and the panel's SYNC NOW drop a .manual
// flag first, which skips the hours.
// Everything here is deterministic. The only judgement - what's yours, what others owe you, in what order - is one
// call to the agent in config.json ("claude" or "codex"), which reads files and writes one file, brief.json.
//   node run.mjs            a scheduled run (weekdays within config.hours, unless .manual is there)
//   node run.mjs --rescan   a full rescan: rebuild everything over the whole lookback
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { CONFIG, OUT, SKILL, config, tool } from './common.mjs';
import { renderBrief, repaint } from './render.mjs';
import { captureSources, captureComplete } from './capture-status.mjs';
import { judgePrompt, parseJudgeResult } from './judge-output.mjs';
import { runSummary } from './run-summary.mjs';
import { signinEvidence, missingSources, failureMessage } from './health.mjs';
import { marketingPrevious } from './marketing.mjs';
import { notify } from './notifications.mjs';

const RESCAN = process.argv.includes('--rescan');
const now = new Date();
const stamp = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}`;
function pad(n) { return String(n).padStart(2, '0'); }

// A run the user asked for ignores the hours
const manual = join(OUT, '.manual');
if (existsSync(manual)) rmSync(manual);
else {
  const day = now.getDay(), hour = now.getHours();
  const { start = 9, end = 18 } = config.hours || {};
  if (day < 1 || day > 5 || hour < start || hour > end) process.exit(0);
}

mkdirSync(join(OUT, 'logs'), { recursive: true });

// One run at a time - a slow run shouldn't overlap the next tick. A lock older than 20 minutes is left over
// from a run that died (lid closed, power off): clear it, or every later run would skip silently.
const lock = join(OUT, '.running');
if (existsSync(lock) && Date.now() - statSync(lock).mtimeMs > 20 * 60e3) rmSync(lock, { recursive: true });
try { mkdirSync(lock); } catch { process.exit(0); }
process.on('exit', () => rmSync(lock, { recursive: true, force: true }));

let log;
try {
repaint(true);   // "Syncing…" on the open panel until this run rewrites it

// The window: from the last run (never further back than the lookback), or the whole lookback on a first run
const read = f => { try { return readFileSync(f, 'utf8').trim(); } catch { return ''; } };
const runs = existsSync(join(OUT, 'runs')) ? readdirSync(join(OUT, 'runs')).filter(f => f.endsWith('.json')).sort() : [];
const PREV = runs.length && !RESCAN ? join(OUT, 'runs', runs.at(-1)) : '';
const LOOKBACK = new Date(Date.now() - (config.lookback_days ?? 30) * 864e5).toISOString().slice(0, 19) + 'Z';
let SINCE = (!RESCAN && read(join(OUT, '.last-run'))) || LOOKBACK;
if (SINCE < LOOKBACK) SINCE = LOOKBACK;   // back from a long break: loops don't reach further than this
const FIRST = PREV ? 'no' : 'yes';
const TODAY_PROJECT = (config.schedule || {})[now.toLocaleDateString('en-US', { weekday: 'short' })] || 'none';

// Collect: Outlook and Teams in a headless Chrome, GitHub and Azure DevOps alongside, into one file. About 25 s.
const DIR = join(OUT, '.inputs');
mkdirSync(DIR, { recursive: true });
const INPUTS = join(DIR, 'inputs.json'), BRIEF = join(DIR, 'brief.json');
// Never let an interrupted run's files masquerade as this run's results.
for (const file of [INPUTS, BRIEF, join(DIR, 'last-message.txt')]) rmSync(file, { force: true });
process.on('exit', () => rmSync(DIR, { recursive: true, force: true }));
const collect = spawnSync(process.execPath, [join(SKILL, 'collect.mjs'), SINCE, LOOKBACK, FIRST, PREV, INPUTS], { encoding: 'utf8', maxBuffer: 64 << 20, timeout: 5 * 60e3, windowsHide: true });
const collected = [`${collect.stdout || ''}${collect.stderr || ''}`.trim(),
  collect.error && `${collect.error.code || 'spawn error'}: ${collect.error.message}`,
  collect.signal && `signal: ${collect.signal}`, collect.status === null && !collect.error && 'collector exited without a status'].filter(Boolean).join('\n');

let inputs = null;
try { inputs = JSON.parse(read(INPUTS)); } catch {}
const sources = inputs ? captureSources(inputs) : missingSources(!!config.ado);
const collection_complete = !!inputs && collect.status === 0 && captureComplete(inputs);
writeFileSync(join(OUT, 'signin-status.json'), JSON.stringify(signinEvidence(inputs || {})));
let earlyAlert = false;
if (!collection_complete) {
  repaint(true, sources);
  earlyAlert = notify(failureMessage({sources}));
}

// Judge. The agent reads the inputs, the previous result, the user's clicks and the config, follows JUDGE.md,
// and writes BRIEF - nothing else: no network, no shell of its own, so it runs with the least access either agent has.
// The prompt goes in on stdin: it's long, and on Windows the agents start through the shell.
const prompt = `You are Meta-Nav's judge, running unattended: nobody will answer a question, so decide, finish and write the file.
Read ${join(SKILL, 'JUDGE.md')} and follow it, with these facts for this run:
SINCE=${SINCE}
LOOKBACK=${LOOKBACK}
FIRST=${FIRST}
TODAY_PROJECT=${TODAY_PROJECT}
INPUTS=${INPUTS}
PREV=${PREV || 'none'}
STATE=${join(OUT, 'state.json')}
CONFIG=${CONFIG}
BRIEF=${BRIEF}
Write your result to BRIEF and change no other file. Then reply with the summary JUDGE.md asks for.`;

const agent = config.agent || 'codex';
const model = config.judge_model || 'gpt-6-luna';
const t0 = Date.now();
if (collect.status !== 0 || !existsSync(INPUTS)) {
  log = { agent, model: model || 'default', is_error: true, result: 'Collection failed; the previous panel has been preserved.', collector_error: collected };
} else if (agent === 'codex') {
  // Run unattended without approval prompts. The user's Codex config (MCP servers, rules) and hooks aren't loaded;
  // hooks.json isn't covered by --ignore-user-config, and a hook waiting on a person would hang the run.
  const last = join(DIR, 'last-message.txt');
  const supplied = judgePrompt({
    instructions: read(join(SKILL, 'JUDGE.md')),
    inputs, previous: marketingPrevious(PREV ? JSON.parse(read(PREV)) : null, config.marketing?.github, inputs.github),
    state: JSON.parse(read(join(OUT, 'state.json')) || '{}'), config,
    facts: { SINCE, LOOKBACK, FIRST, TODAY_PROJECT, now: now.toISOString(), timezone: Intl.DateTimeFormat().resolvedOptions().timeZone },
  });
  const r = tool('codex', ['exec', '--json', '--ignore-user-config', '--disable', 'hooks', '--skip-git-repo-check', '--ephemeral',
    '--sandbox', 'read-only', '--cd', DIR, '-c', 'model_reasoning_effort=high',
    '--output-schema', join(SKILL, 'judge-output.schema.json'),
    ...(model && !/^(opus|sonnet|haiku|claude)/.test(model) ? ['--model', model] : []),
    '--output-last-message', last, '-'], { cwd: DIR, input: supplied, timeout: 10 * 60e3 });
  // token use, from each turn's usage event, so a run's cost can be worked out (Claude Code reports its own)
  const usage = {};
  for (const line of (r.stdout || '').split('\n')) {
    let e; try { e = JSON.parse(line); } catch { continue; }
    if (e.type === 'turn.completed') for (const [k, v] of Object.entries(e.usage || {})) usage[k] = (usage[k] || 0) + v;
  }
  log = { agent, model: model || 'default', is_error: r.status !== 0, result: read(last) || (r.stderr || '').slice(-4000), usage };
  if (!log.is_error) {
    try {
      const result = parseJudgeResult(read(last));
      writeFileSync(BRIEF, JSON.stringify(result.brief));
      log.judge_summary = result.summary;
      log.result = runSummary(result.brief, sources);
    } catch (e) { log.is_error = true; log.result = 'Judge output rejected: ' + String(e.message || e); }
  }
} else {
  // Only Read and Write: no MCP servers (--strict-mcp-config with none given), no shell, no other tool
  const r = tool('claude', ['-p', '--model', model || 'opus', '--strict-mcp-config',
    '--add-dir', SKILL, '--allowedTools', 'Read', 'Write', '--output-format', 'json'],
    { cwd: OUT, input: prompt, timeout: 20 * 60e3 });
  try { log = { agent, ...JSON.parse(r.stdout) }; } catch { log = { agent, is_error: true, result: `${r.stdout || ''}${r.stderr || ''}`.slice(-4000) }; }
}
log.duration_ms ??= Date.now() - t0;
log.collect = collected;

// Render what the judge wrote. A run that wrote nothing usable leaves the last good result up - the panel marks it
// stale once it's over 90 minutes old.
let brief = null;
if (!log.is_error) {
  try { brief = JSON.parse(readFileSync(BRIEF, 'utf8')); } catch (e) { log.is_error = true; log.brief_error = 'Judge did not produce a usable brief: ' + String(e.message || e); }
}
if (brief && !log.is_error) {
  try { renderBrief(BRIEF, INPUTS); } catch (e) { log.is_error = true; log.render_error = String(e.stack || e).slice(0, 2000); }
}
if (/"syncing":\s*true/.test(read(join(OUT, 'index.html')))) repaint(false);
log.sources = sources;
log.collection_complete = collection_complete;
if (!collection_complete) log.is_error = true;
if (log.is_error) {
  repaint(false, [...sources, ...(collection_complete ? [{name:'Action list',ok:false,note:'Luna could not update the panel; previous information is retained.'}] : [])]);
  log.notification_sent = earlyAlert || notify(failureMessage(log));
}
log.result = String(log.result || 'Meta-Nav run did not complete.').replace(/\s+/g, ' ').trim();
writeFileSync(join(OUT, 'logs', `${stamp}${RESCAN ? '-rescan' : ''}.json`), JSON.stringify(log, null, 1));

// The inputs hold private chats and mail: deleted every time, even after a failed run
rmSync(DIR, { recursive: true, force: true });
for (const f of readdirSync(join(OUT, 'logs'))) {
  const p = join(OUT, 'logs', f);
  if (f.endsWith('.json') && Date.now() - statSync(p).mtimeMs > 7 * 864e5) rmSync(p);
}

} catch (error) {
  log = { ...(log || {}), is_error:true, collection_complete:false,
    result:'Meta-Nav could not complete this refresh; previous information is retained.',
    runner_error:String(error.message || error).slice(0,1000), sources:missingSources(!!config.ado) };
  try { repaint(false, log.sources); } catch {}
  try { log.notification_sent=notify(failureMessage(log)); } catch {}
  writeFileSync(join(OUT,'logs',`${stamp}.json`),JSON.stringify(log,null,1));
} finally {
  rmSync(join(OUT,'.inputs'), {recursive:true,force:true});
}
