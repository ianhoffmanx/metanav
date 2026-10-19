import test from 'node:test';
import assert from 'node:assert/strict';
import { recognizedApps } from '../microsoft.mjs';

const page = (url, title) => ({ url, title });

test('a signed-in mailbox and Teams chat are recognised', () => {
  const apps = recognizedApps([
    page('https://outlook.office.com/mail/inbox', 'Inbox | Ian Ho | Outlook'),
    page('https://teams.microsoft.com/v2/', 'Chat | Ian Ho | Microsoft Teams'),
  ]);
  assert.deepEqual(apps, { outlook: true, teams: true });
});

test('the Outlook 401 error page and the Teams public landing page are not signed in', () => {
  const apps = recognizedApps([
    page('https://outlook.office.com/mail/', 'Outlook'),
    page('https://teams.microsoft.com/v2/', 'Microsoft Teams'),
  ]);
  assert.deepEqual(apps, { outlook: false, teams: false });
});
