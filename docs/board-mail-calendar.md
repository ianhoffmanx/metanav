# Calendar and reply drafts on Ian's board

The collector still reads mail/calendar with the approved Microsoft SDK client.
Teams uses the existing signed-in browser. Luna remains the judge and decides
which asks matter. `board-context.mjs` attaches collected facts after judging:

- `agenda`: the full Graph calendar window, original collection time and timezone.
  Cancelled/declined meetings are omitted. The board converts Sydney/Brisbane time
  correctly and displays Today/Tomorrow without writing `inputs/calendar.json`.
- `reply_target`: a message ID and display metadata, only when a mail task matches
  a collected incoming conversation/message/link. No subject-only guessing. The
  board cache omits the message ID; the laptop resolves it again from its own run.

Reply drafts use the tenant-approved **AI Mail reader** registration. Its delegated
`Mail.ReadWrite` grant is separate from the Microsoft SDK client's read consent.
The web callback already registered by SSW is exactly:
`http://localhost:47616/auth/callback`.

`POST /signin {"source":"Mail drafts"}` is an explicit user action. It opens the
default browser, uses authorization code + PKCE + state/nonce, checks the configured
SSW account and tenant, then saves only its refresh token using Windows DPAPI in
`%LOCALAPPDATA%/metanav/graph-drafts-token.dat`. Nothing is copied to Linux. The
existing read-only SDK session and Chrome profile are not reset or changed.

Additive loopback-only service routes:

- `POST /mail/draft-status {}`: silent readiness check, never prompts or writes mail.
- `POST /mail/reply-draft {key,text,request_id}`: looks up the mail item in the latest
  run, validates its collected message ID, then calls `createReply` once, with plain
  text. No send/mark-read/delete route exists. Draft text is never logged or exported.

Receipts contain hashes and the resulting draft link, not reply text. Repeated
requests (including after a page reload or service restart) return the original
receipt. Interrupted/uncertain writes ask Ian to check Outlook Drafts, without
repeating the Graph POST. Definite pre-write sign-in/consent failures are retryable.

Existing `/status`, `/refresh`, run-log fields and source-note strings are unchanged.
The server Docker package carries the UI and SSH control only; this Windows token
and the app secret are not present in the image or bundle.
