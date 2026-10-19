export function browserArgs(profile, {headless=false}={}) {
  return [...(headless?['--headless=new']:[]),`--user-data-dir=${profile}`,'--remote-debugging-port=0','--no-first-run','--no-default-browser-check',
    'https://outlook.office.com/mail/inbox',...(!headless?['https://teams.microsoft.com/v2/']:[])];
}
