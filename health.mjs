// Collection facts drive status and alerts; the judge never decides source health.
export function signinEvidence(inputs, now = new Date()) {
  const check = data => !data ? 'unknown' : data.signin_needed ? 'required' : data.error ? 'error' : 'ready';
  const outlook = check(inputs.mail), teams = check(inputs.teams);
  const status = [outlook, teams].includes('required') ? 'required' : [outlook, teams].every(x => x === 'ready') ? 'ready' : 'error';
  return { status, apps: { outlook: outlook === 'ready', teams: teams === 'ready' }, at: now.toISOString(),
    message: status === 'required' ? 'Click SIGN IN in Meta-Nav, then complete Microsoft sign-in in its Chrome window.'
      : status === 'error' ? 'Microsoft collection could not be verified. Click SYNC NOW to check again.' : 'Outlook and Teams were checked during collection.' };
}
export function freshSignin(signin, now = Date.now()) {
  if (!signin) return {status:'unknown',apps:{outlook:false,teams:false},at:null,message:'No Microsoft collection check has completed.'};
  if (!Number.isFinite(Date.parse(signin.at)) || now - Date.parse(signin.at) > 90 * 60000) return { ...signin, status:'stale', message:'Microsoft sign-in has not been checked in the last 90 minutes. Click SYNC NOW.' };
  return signin;
}
export function missingSources(adoEnabled = false, graphEnabled = false) {
  return ['Outlook','Teams','GitHub',...(adoEnabled?['Azure DevOps']:[]),...(graphEnabled?['Calendar']:[])].map(name=>({name,ok:false,note:'Collection did not finish; the previous information is retained.'}));
}
export function failureMessage(log) {
  const failed = (log.sources || []).filter(s=>!s.ok);
  if (failed.some(s=>s.name === 'Calendar' && /sign.in.*required/i.test(s.note))) return 'Calendar sign-in is required. Open Meta-Nav Calendar Sign-in on your desktop and choose your SSW account. Other source failures are shown in Meta-Nav.';
  if (failed.some(s=>/sign.in.*required/i.test(s.note))) return 'Microsoft sign-in expired. Open Meta-Nav, click SIGN IN, and complete sign-in in its Chrome window.';
  return failed.length ? `${failed.map(s=>s.name).join(' and ')} could not be checked. Open Meta-Nav and click SYNC NOW. Previous information may be out of date.`
    : 'Meta-Nav could not update your action list. Open Meta-Nav and click SYNC NOW. Previous information may be out of date.';
}
