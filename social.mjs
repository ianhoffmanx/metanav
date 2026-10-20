// Workplace social (birthdays, welcomes, farewells, drinks, news): the judge proposes, this file makes sure what lands
// in runs/*.json always has the same shape - the dashboard reads it without checking.
export const SOCIAL_KINDS = ['birthday', 'anniversary', 'welcome', 'farewell', 'event', 'news'];
export const SOCIAL_DAYS_BACK = 2;
export const SOCIAL_DAYS_AHEAD = 7;
const MAX_ITEMS = 25;

// Teams channels aren't read for asks, but a channel is where a birthday or a farewell is posted: only messages that
// look social are let through (teams-extract.js gets this exact pattern from collect.mjs).
export const SOCIAL_WORDS = 'birthday|b-?day|anniversar|work ?iversary|welcome|new starter|farewell|last day|goodbye|congrat|'
  + 'baby|wedding|engaged|engagement|drinks|lunch|dinner|party|celebrat|morning tea|happy hour|cake|trivia|karaoke|'
  + 'christmas|retire|🎂|🎉|🥳|🍰|🍾';

const DATE = /^\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:\d{2})?)?$/;
const text = (v, max) => (typeof v === 'string' ? v.replace(/\s+/g, ' ').trim().slice(0, max) : '');
const startOfDay = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
// A date alone ("2026-10-21") is that day in the user's time zone; Date.parse would call it UTC midnight
const instant = s => (/^\d{4}-\d{2}-\d{2}$/.test(s) ? new Date(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)).getTime() : Date.parse(s));

function clean(item) {
  if (!item || typeof item !== 'object' || Array.isArray(item)) return null;
  const id = text(item.id, 300), title = text(item.title, 120);
  if (!id || !title || !SOCIAL_KINDS.includes(item.kind)) return null;
  const when = typeof item.when === 'string' && DATE.test(item.when.trim()) && !Number.isNaN(instant(item.when.trim())) ? item.when.trim() : null;
  const url = typeof item.url === 'string' && /^(?:https:\/\/|msteams:)/i.test(item.url.trim()) ? item.url.trim() : null;
  const who = (Array.isArray(item.who) ? item.who : typeof item.who === 'string' ? [item.who] : [])
    .map(w => text(w, 80)).filter(Boolean).slice(0, 8);
  return { id, kind: item.kind, title, who, when, where: text(item.where, 120) || null, detail: text(item.detail, 200) || null, url };
}

// What the judge returned -> the `social` array of a run: only well-formed items, the last 2 days to the next 7, no duplicate
// ids, soonest first (no date last), and `new` against the previous run's social ids.
export function normaliseSocial(items, { now = new Date(), previous = [] } = {}) {
  const from = addDays(startOfDay(now), -SOCIAL_DAYS_BACK).getTime();
  const to = addDays(startOfDay(now), SOCIAL_DAYS_AHEAD + 1).getTime();   // end of the 7th day ahead
  const seen = new Set(Array.isArray(previous) ? previous.map(p => p && p.id) : []);
  const ids = new Set(), out = [];
  for (const raw of Array.isArray(items) ? items : []) {
    const item = clean(raw);
    if (!item || ids.has(item.id)) continue;
    if (item.when) { const t = instant(item.when); if (t < from || t >= to) continue; }
    ids.add(item.id);
    out.push({ ...item, new: !seen.has(item.id) });
  }
  const order = i => (i.when ? instant(i.when) : Number.MAX_SAFE_INTEGER);
  out.sort((a, b) => order(a) - order(b));
  return out.slice(0, MAX_ITEMS);
}
