# Microsoft calendar sign-in on Ian's laptop

## Why the Reader sign-in failed

The `AI Mail reader` registration has delegated permissions, no web/native/SPA redirect
addresses, and no enabled public-client fallback. Device-code authentication is a public-client
flow; adding its client secret does not turn the registration into one. A confidential browser
flow would need a registered redirect address. Ian cannot edit either setting, so Meta-Nav
must not send him through another device-code attempt using that registration.

SSW already has the Microsoft Graph PowerShell client (`14d82eec-204b-4c2f-b7e8-296a70dab67e`)
with tenant-wide delegated consent for `Calendars.Read` and `Mail.Read`. The new provider uses
the public Azure.Identity API shipped by Microsoft.Graph.Authentication 2.41.2, requesting
exactly those two read scopes. No Reader or SharePoint secret is passed to that provider and
no tenant/app-registration setting is changed. This is user access, not an unattended app
permission grant. It does not grant Teams chat access.

## Configuration and operation

Set `graph.provider` to `microsoft-sdk` in the ignored `config.json`. The existing Reader
`.env` supplies the tenant; `microsoft_account` supplies the expected SSW account. Leave that
`.env` on the laptop. PowerShell 7 and the CurrentUser `Microsoft.Graph.Authentication` module
are installed there. The adapter loads its public Azure.Identity assemblies without relying
on the cmdlets' isolated load context.

From the signed-in Windows desktop, `node graph.mjs login` performs explicit browser sign-in.
`windows/install-graph-login.ps1` installs the manual `MetaNav-GraphLogin` scheduled task and
the **Meta-Nav Calendar Sign-in** desktop shortcut. The task runs as Ian in his desktop session,
without elevation or saved credentials. It is never triggered by an hourly refresh. The panel's existing SIGN IN button now routes
Outlook (when using Graph mail) and Calendar to this same explicit SDK sign-in. A Teams SIGN IN
opens only Teams in the dedicated browser. Neither requires Ian to edit an app registration.

`node graph.mjs status` checks the saved sign-in silently. `node graph.mjs calendar` reads the
collector calendar window and reports only the count and successful inbox read check. These
commands must run in Ian's desktop session; SSH has a different credential environment.
Real collection uses the resident service's `POST /refresh`, never `run.mjs` over SSH.

Tokens are managed by Azure.Identity in its encrypted Windows cache named
`metanav.graph.readonly`; unencrypted fallback is disabled. The SDK authentication record
(account metadata, not a bearer token) is in `%LOCALAPPDATA%\metanav\graph-sdk-authrecord.json`.
No tokens or profiles are copied to Linux. Background token requests disable automatic
interactive authentication. When silent renewal cannot proceed, collection reports the
failure instead of opening repeated sign-in windows. A successful login is not proof that
SSW policy will allow refresh indefinitely; future policy/MFA changes can still require Ian.
The adapter's read process has an 45-second limit, inside the collector's 50-second budget,
so a timed-out SDK cannot keep the collector alive after its summary.

## What reaches the panel

The existing `inputs.graph.calendar` schema is retained (local start/end, meeting subject,
organizer, location, online link and short body preview). Luna remains `gpt-6-luna` and uses
that evidence for meeting preparation and workplace social dates. No dashboard calendar
`inputs/` export is created. The SDK also replaces the Outlook browser mail collector: it reads only configured received
folders and Sent, using read-only GETs with message previews, people, read/flag state, dates,
conversation IDs and Microsoft's own links. Inbox and Sent are the only enabled folders on
Ian's laptop; GitHub mail remains excluded. The full 30-day lookback is read, plus older flags
on a first run. Pagination is bounded and fails visibly if incomplete. Custom received folders
need an explicit `graph_id`; they are never silently omitted. No GET marks mail read.

`graph.mail: false` retains browser mail as an explicit fallback; the SDK default is Graph mail.
Calendar and mail failures are independent. The existing `mail.folders` row shape is preserved,
with extra Graph metadata. `deletion_evidence: false` prevents a missing or migrated browser ID
from being treated as completed work. Luna prefers the supplied official deep links and combines
replies by conversation/deliverable. Teams still needs its browser session because the approved
Graph scopes do not include chat read access. With SDK mail enabled, browser startup selects
only Teams; Outlook is no longer opened for collection.

When the SDK is configured, the existing `sources[]` gains `name: "Calendar"`; its shape is
unchanged. A failed calendar read makes `collection_complete` false and raises the existing
Windows alert. Existing browser sign-in notes keep their wording. The already-present attach-timeout branch
now retains the dashboard's classification phrase: `Could not attach to Meta-Nav Chrome: it is open but not responding. Close only its dedicated window, then use SIGN IN to reopen it.` Exact added status notes are:

- Success: `Captured successfully.`
- Missing/expired SDK sign-in: `Graph sign-in is required. Open Meta-Nav Calendar Sign-in on your desktop.`
- SDK failures use the sanitized adapter reason, including `Microsoft Graph could not be read.`,
  `SSW approval is required for the requested Graph permissions.`,
  `Microsoft Graph authentication or collection timed out.`,
  `The selected Microsoft account does not match the configured SSW account.`,
  `Graph SDK timed out; the saved session was retained.`,
  `Graph SDK could not start. Install PowerShell 7 and Microsoft.Graph.Authentication.`,
  `Graph SDK did not return a usable result.`, or `Graph SDK output exceeded its collection limit.`
- Mail-specific failure: `Graph mail collection failed.`; calendar-specific failure: `Graph calendar collection failed.`
- Combined mail/calendar outer deadline: `Graph mail and calendar timed out; collection is incomplete`
- Outer collection deadline: `Graph calendar timed out; collection is incomplete` (if that independent guard wins).
- Fallback: `Calendar collection did not complete.`

Collector crashes retain `Collection did not finish; the previous information is retained.`
for Calendar as well as the original sources. `POST /status`, `POST /refresh`, the log fields,
and `.last-run` keep their existing contracts.

## Validation

The explicit desktop login succeeded for Ian's configured SSW account. Independent fresh
processes then successfully checked the cached sign-in and read **38 calendar events** plus
the inbox read check without another prompt. A further live SDK read returned **141 Inbox messages, 53 Sent messages and 38 calendar events**,
without another login. **53 Windows tests passed**, including panel authentication routing and the HTTP status/refresh contracts. The service was restarted through its scheduled task; the dedicated Chrome PID remained 39336 before and after, so activation did not close that browser or reset its profile. Regression coverage includes the real Windows
script's clean no-cache response, secret-free SDK inputs, the legacy confidential-device-flow
guard, normalized calendar data, visible calendar failures and the dedicated sign-in alert.

Administrator fallback if the Reader registration is approved later: Web redirect URI
`http://localhost:47616/auth/callback`, using authorization code with PKCE and the existing
confidential client secret. Implicit grants stay disabled. This endpoint is reserved for that
future implementation; the working SDK provider uses its own already-registered loopback redirect.
