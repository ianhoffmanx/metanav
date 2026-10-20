# Shared server feed (Windows connector)

The companion `metanav-hermes` exporter can be called by Meta-Nav's resident service
on run exit and after a Done/Undo/Got-it change. It is optional and changes no
HTTP endpoints, log fields, source-note text, hours or judge configuration.

Add this to the machine's ignored `config.json` (use the actual installed path):

```json
"exporter": {
  "script": "C:\\Users\\ianho\\.claude\\skills\\metanav-hermes\\export.mjs"
}
```

The service invokes Node directly with an argument array, without a shell, secrets
or UI. Concurrent triggers coalesce and send again once the active export exits.
A three-minute timeout bounds a stuck exporter. Keep the companion's existing
15-minute Task Scheduler job as the retry fallback. Export failure never blocks
collection or a click. Restart only the resident Node service after changing config;
leave its Chrome process/profile alone and restart the Meta-Nav task afterward.

The companion's schema 1.1 adds optional social/context, panel marks, source health
and a content-derived delivery identifier. A click changes delivery identity but
preserves Luna's original run identifier and collection time. The server validates
and imports the result, then publishes one sanitized dashboard snapshot with the
same identifier as Hermes. The existing SSH status/control route remains available.

Server deployment is packaged in work-dashboard's `deploy/docker/compose.yaml`.
It runs the dashboard, status collectors and encrypted importer. Windows collection
and the `gpt-6-luna` Codex judge remain on the laptop. No credentials, browser
profiles or Windows SDK token caches are bundled. No dashboard calendar input is
created. A future Linux collection move requires separate authentication testing.
