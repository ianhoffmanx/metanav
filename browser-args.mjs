export function browserArgs(profile, {headless=false,apps=['outlook','teams']}={}) {
  const urls={outlook:'https://outlook.office.com/mail/inbox',teams:'https://teams.microsoft.com/v2/'};
  const selected=apps.filter(app=>urls[app]);
  if(!selected.length)throw Error('No supported browser source was selected');
  return [...(headless?['--headless=new']:[]),`--user-data-dir=${profile}`,'--remote-debugging-port=0','--no-first-run','--no-default-browser-check',
    ...(headless?selected.slice(0,1):selected).map(app=>urls[app])];
}
