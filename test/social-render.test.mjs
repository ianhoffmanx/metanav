import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, cpSync, mkdirSync, writeFileSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test('a rendered run carries a social array, with new set against the previous run and the other fields untouched', () => {
  const dir = mkdtempSync(join(tmpdir(), 'metanav-social-'));
  try {
    const source = fileURLToPath(new URL('..', import.meta.url)), out = join(dir, 'out');
    for (const f of readdirSync(source).filter(f => f.endsWith('.mjs') || f === 'template.html')) cpSync(join(source, f), join(dir, f));
    mkdirSync(join(out, 'runs'), { recursive: true });
    writeFileSync(join(dir, 'config.json'), JSON.stringify({ output_dir: out }));
    const d = new Date(), today = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    const brief = social => ({ date: today, generated_at: '2026-10-20T01:00:00Z', interval_min: 60, since: '2026-10-20T00:00:00Z', headline: 'H', tradeoffs: [],
      queue: [{ id: 'q1', title: 'Approve blog', urgency: 'normal' }], waiting: [], calendar: [], highlights: [], done: [], ...(social && { social }) });
    const run = (name, b) => {
      writeFileSync(join(dir, name), JSON.stringify(b));
      const r = spawnSync(process.execPath, [join(dir, 'render.mjs'), join(dir, name)], { encoding: 'utf8', timeout: 20000, cwd: dir });
      assert.equal(r.status, 0, r.stderr);
      const files = readdirSync(join(out, 'runs')).sort();
      return JSON.parse(readFileSync(join(out, 'runs', files.at(-1)), 'utf8'));
    };
    const party = { id: 'c1', kind: 'event', title: 'Team lunch', who: [], when: today, where: null, detail: null, url: 'https://outlook.office.com/mail/inbox/id/c1' };
    const old = { ...party, id: 'c2', title: 'Old drinks', when: '2020-01-01' };

    const first = run('b1.json', brief([party, old]));
    assert.deepEqual(first.social.map(x => [x.id, x.new]), [['c1', true]]);   // the 2020 item is outside the window
    assert.deepEqual(first.queue, [{ id: 'q1', title: 'Approve blog', urgency: 'normal' }]);   // first run: nothing to diff against, as before
    assert.equal(first.headline, 'H');

    // the run files are named to the minute: make the second one a later minute
    const runs = join(out, 'runs'), f = readdirSync(runs)[0];
    cpSync(join(runs, f), join(runs, '20200101-0000.json')); rmSync(join(runs, f));
    const fresh = { ...party, id: 'c3', title: 'Welcome Sam', kind: 'welcome', who: ['Sam'] };
    const second = run('b2.json', brief([party, fresh]));
    assert.deepEqual(second.social.map(x => [x.id, x.new]), [['c1', false], ['c3', true]]);   // c1 was in the previous run
    assert.equal(second.queue[0].new, false);   // the existing new/updated logic is unchanged

    const third = run('b3.json', brief());   // judge gave none: still an array
    assert.ok(Array.isArray(third.social));
  } finally { rmSync(dir, { recursive: true, force: true }); }
});
