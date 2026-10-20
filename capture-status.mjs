// Source health is collection evidence, not an AI judgement.
export function captureSources(inputs) {
  const browser = (name, data, ready) => {
    const ok = !!data && !data.error && !data.signin_needed && ready(data);
    return { name, ok, note: ok ? 'Captured successfully.' : data?.signin_needed
      ? (data?.provider === 'microsoft-sdk' ? 'Graph sign-in is required. Open Meta-Nav Calendar Sign-in on your desktop.' : 'Microsoft sign-in is required in the Meta-Nav browser.')
      : (data?.error || 'Collection did not complete.') };
  };
  const sources = [
    browser('Outlook', inputs.mail, data => !!data.folders && Array.isArray(data.calendar)),
    browser('Teams', inputs.teams, data => Array.isArray(data.conversations)),
    { name: 'GitHub', ok: !!inputs.github?.ok, note: inputs.github?.error || (inputs.github?.ok ? 'Captured successfully.' : 'Collection did not complete.') },
  ];
  if (inputs.ado && !inputs.ado.skipped) sources.push({ name: 'Azure DevOps', ok: !!inputs.ado.ok, note: inputs.ado.error || (inputs.ado.ok ? 'Captured successfully.' : 'Collection did not complete.') });
  if (inputs.graph?.provider === 'microsoft-sdk') {
    const graph=inputs.graph, ok=!!graph.ok && Array.isArray(graph.calendar);
    const needsSignin=graph.skipped || graph.code === 'signin_required';
    sources.push({name:'Calendar',ok,note:ok ? 'Captured successfully.' : needsSignin
      ? 'Graph sign-in is required. Open Meta-Nav Calendar Sign-in on your desktop.'
      : (graph.error || graph.note || 'Calendar collection did not complete.')});
  }
  return sources;
}
export function captureComplete(inputs) { return captureSources(inputs).every(source => source.ok); }

