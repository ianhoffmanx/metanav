export function runSummary(brief, sources) {
  const oneLine = s => String(s || '').replace(/\s+/g,' ').trim();
  return `${oneLine(brief.headline)} Queue ${brief.queue.length}; waiting ${brief.waiting.length}; highlights ${brief.highlights.length}; done ${brief.done.length}. Sources: ${sources.map(s=>`${s.name} ${s.ok?'OK':'FAILED'}`).join('; ')}.`;
}
