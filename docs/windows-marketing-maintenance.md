# Ian's Windows marketing maintenance — 19 October 2026

The running checkout is `C:\Users\ianho\.claude\skills\metanav`. Real refreshes are started only
with `POST /refresh` on the resident service, so GitHub credentials and Chrome belong to the
signed-in Windows desktop user. The collector remains read-only.

## Findings and fixes

- **Frozen tabs:** the 23:22 run (finished 23:23:43 AEST) and the following diagnostic run failed
  to attach to Chrome. Their `outlook_signin_needed` and `teams_signin_needed` flags are false.
  Raw CDP probes found three unresponsive application tabs while another Teams tab responded.
  Playwright's browser attachment waits on all targets, so one frozen renderer defeats it.
  Recovery now probes each target with bounded frame/runtime checks, closes frozen targets,
  removes duplicate app tabs and then attaches once. It never deletes a profile or kills the
  user's normal browser. The laptop had only about 150–350 MB RAM available out of 16 GB;
  that is an observed contributing risk, not a proven sole cause of the renderer stalls.
- **Sign-in evidence:** the main Chrome PID 28732 remained alive from about 22:20 throughout the
  23:23 failure. The first recovery was at about 23:26, after that failure. Collection at 23:35
  succeeded for Outlook and Teams without Ian signing in again. The failure does not establish
  a session lifetime of under an hour. No password, MFA, cookie or profile was copied or reset.
- **Stale status:** `signin-status.json` was last updated by the explicit sign-in workflow on
  9 October. Each collection now writes fresh evidence, `/status` labels checks over 90 minutes
  old as stale, and a URL/title alone is labelled checking rather than ready.
- **Missing brief:** the runner read `.inputs/brief.json` even when it deliberately skipped the
  judge because collection had failed. It now reads a brief only after a successful judge and
  always logs `is_error`, `collection_complete`, a one-line `result` and source statuses.
- **Silent failures:** collection and judge failures now repaint the old panel with current
  failure status and raise a Windows toast. A planned collector crash through `/refresh` left
  `.last-run` unchanged, kept the old result, recorded all four failed sources, produced no
  misleading missing-brief error and returned `notification_sent: true`. Windows notification
  history independently contained one matching Meta-Nav alert.
- **Misleading summary:** Luna repeated a previous Outlook failure in its prose even when all
  current sources were healthy. Current source facts are now explicit in the prompt; the runner
  composes the log's headline, counts and source-health clause deterministically. The original
  judge summary remains separately available for auditing.
- **Historical post-summary timeout:** the 17:00 log contains a collector summary at 129.2 s and
  a subsequent five-minute `spawnSync ETIMEDOUT`. The diagnostic attach-failure run exited
  naturally, and the recovered 223 s run also exited naturally. A leaked connection has not
  been established. Browser stages, credential commands and network calls now have bounded
  waits; `collector-stage.json` records the current stage, and a process remaining alive after
  its summary records active resource types. Do not claim the historical exit hang's root cause
  is proven until it reproduces with that evidence.

## Marketing scope

Luna remains the judge (`agent: codex`, `judge_model: gpt-6-luna`). `JUDGE.md` now prioritises
asks, promises, approvals, content deadlines, meeting preparation/follow-up and marketing
collection failures. It names the asker, uses evidenced dates or null, deduplicates board/chat
asks and treats a product release with unknown coverage as a blog-coverage check.

Azure DevOps was disabled. The configuration now reads only assigned items in the current
`SSW.Marketing / SSW.Marketing Team` iteration, without collecting ADO PRs. The first real run
found nine items in Sprint 348 (9–23 October). GitHub is restricted to SSWConsulting/tinacms,
content evidence and explicit marketing asks. Side projects are excluded. PR metrics are removed
from marketing inputs and panel data; publication blockers remain in plain language.

The first filter wrongly admitted application code under `app/blog` and CSS footer JSON. The
real files for SSW.Products #827/#839 established these were technical changes. Regression
checks now cover those cases, and old automatic review rows are removed when current filtering
identifies them as out of scope. Independent content asks in mail/chat remain eligible.

## Branch reconciliation

The active base is flat-layout WIP commit `044b3e6`. Preserve its running Node service and the
Microsoft rework. The older `windows-support` branch uses obsolete Bash/Python runners:

- Windows scheduling and toast support are already in `94d7f1e` / the Node service.
- Multiple GitHub organisations are already in `a3e218d`.
- Saved-account selection and preserving the signed-in Chrome profile are in `microsoft.mjs`.
- The legacy second Outlook-list wait is retained in the repaired collector.
- The old Python/Bash task installer and runner must not replace `state.mjs` / `run.mjs`.

The project checkout's `b206e66` is an adapted copy of the same Microsoft rework on a newer,
reorganised tree. Preserve that WIP branch for recovery; the operational checkouts are to use
the same tested maintenance branch rather than two independently edited implementations.

`origin` was still AttackOnMorty's repository despite the handover. It is now Ian's fork;
`upstream` keeps the original fetch URL with pushes disabled. Nothing has been pushed.

## Hermes and credentials

The encrypted exporter was already delivering successfully: the 22:37 run reached the Linux
bridge. The separate Git fetch error came from trying to use Windows Credential Manager in an
SSH session. A desktop-session scheduled maintenance task fetched both repositories successfully
with `gh auth git-credential`, without reading or printing tokens. `windows/desktop-fetch.ps1`
provides the reusable command; it must run in the interactive Windows session. Export still uses
the existing age recipient and restricted receiver key in place.

## Contracts

`POST /status` still returns `{running, signing, signin}` and `POST /refresh` still returns
`{status: started|busy}` on loopback port 47615. The log and `.last-run` contracts are retained.
No dashboard calendar export, tenant, Entra, Microsoft policy or Tailscale setting is changed.

## Independent Luna review

A GPT-6-Luna child reviewed a real panel from Ian's point of view. Accepted fixes: combine the
SSWTV tagline board/chat duplicates, distinguish ASF planning from drafting, label sprint-end
estimates explicitly, identify requesters/approvers rather than implying reassignment, and make
content-PR next steps clear. Clarified rather than implemented: Waiting On already uses `what`
as its title; the existing renderer already shows a source-failure banner; new release coverage,
automation failures or meeting follow-up cannot be invented without source evidence.
