# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Slack next-gen platform (Deno SDK) app whose only real purpose is to expose
one custom function, `create_calendar_event`, as a step usable in **Slack
Workflow Builder** (the no-code workflow tool, configured entirely in the Slack
UI — not in this repo). The actual business workflow (a trigger on a Slack List
field update, a form, etc.) lives in Workflow Builder and is out of scope for
this codebase; this repo just supplies the one step that workflow calls out to.

That step calls the Microsoft Graph API (client-credentials / app-only OAuth2)
to create a calendar event on a shared scheduling mailbox on behalf of a
request submitted through the Slack workflow.

Full narrative context (architecture diagram, IT credential process, known
limitations) is in `README.md` — read that first for the "why."

## Commands

```sh
# Run locally against the emtechllc workspace (installs "<app name> (local)")
slack run --team emtechllc --hide-triggers

# Health-check the toolchain (CLI, auth, deno, SDK deps)
slack doctor

# Run tests
deno test

# Format / lint (also run by `deno task test`)
deno fmt --check
deno lint

# Deploy to Slack-hosted infra (prompts for a trigger on first deploy)
slack deploy

# Set an env var for the deployed app (local dev uses .env instead)
slack env add MS_TENANT_ID <value>
```

**Windows/PATH gotcha:** Deno is installed via winget but this shell session
was live before the PATH update landed, so `deno`/`slack` may not resolve in
a fresh shell within this session. If a `slack` command hangs with literally
zero output, check `deno --version` first — if it's not found, prefix the
command with:
```sh
export PATH="/c/Users/cragandb/AppData/Local/Microsoft/WinGet/Packages/DenoLand.Deno_Microsoft.Winget.Source_8wekyb3d8bbwe:$PATH"
```
A genuinely new terminal/VS Code window won't need this.

**Git Bash/pty gotcha:** independent of the PATH issue, `slack.exe` hangs with
zero output under Git Bash's pty for some commands (`doctor`, `run` without
`--team`) — they're waiting on an interactive prompt that Git Bash can't
satisfy, and the process produces no output until killed. This is not a real
error; the same commands complete normally within a couple seconds from native
PowerShell/cmd. `slack run --team emtechllc` explicitly passes the team to skip
the workspace-selection prompt that triggers this.

## Architecture

- **`manifest.ts`** — registers `CreateCalendarEventDefinition` under
  `functions:` (not `workflows:`). That's the specific field that makes a
  function selectable as a step inside Slack's no-code Workflow Builder;
  functions only referenced via a code-defined `DefineWorkflow` wouldn't show
  up there. `outgoingDomains` must list `login.microsoftonline.com` and
  `graph.microsoft.com` — the Deno SDK sandboxes `fetch` to declared domains.

- **`functions/create_calendar_event.ts`** — the one file that matters.
  - `input_parameters`: submitter alias/email, request type, `category`,
    start/end ISO 8601 datetimes, description, location, and an array of
    additional attendee emails (mapped in Workflow Builder from the
    triggering Slack List row / form fields).
  - `category` is deliberately a separate input from `request_type`, not
    reused — `request_type` only affects the human-readable title text,
    while `category` must exactly match one of the categories IT
    pre-configured on the shared mailbox (`OOTO`, `4x10 OOTO`, `WFH`,
    `On-Site`, currently — see README's "Outlook categories" section for the
    live list) and gets sent as Graph's `categories` field for calendar
    color-coding. It's a plain string, not a code-enforced enum — the valid
    set grows over time as IT adds per-site categories, and enum-locking it
    here would mean a redeploy every time IT adds one.
  - `getGraphAccessToken()`: does the OAuth2 client-credentials POST to
    `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token` to get an
    app-only bearer token — no user session involved, since the workflow runs
    unattended.
  - Handler builds the title as `<alias> - <request type> - <date>`, builds
    the attendee list (submitter + additional attendees, all as Graph
    `attendees` so they get real Outlook invites), and POSTs to
    `/v1.0/users/{MS_SHARED_MAILBOX}/events` with `showAs: "free"` and
    `categories: [category]`.
  - Env vars (`env` destructured from the handler's context, not
    `Deno.env`): `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`,
    `MS_SHARED_MAILBOX`, `MS_EVENT_TIMEZONE`. All required-but-missing cases
    return `{ error: "..." }` rather than throwing — Workflow Builder surfaces
    that as a clean failed-step message.

- **Do not request `Calendars.ReadWrite` as a tenant-wide app permission.**
  IT deliberately did *not* grant that as an application permission with
  admin consent — instead they scoped this app's mailbox access at the
  Exchange level (an Application Access Policy restricting it to just
  `teamavailability@emtech.us`). Per IT's explicit instruction, adding that
  permission or requesting admin consent would override their scoping and
  open access to every mailbox in the tenant. If a real API call ever fails
  with a 401/403 permission error, that's a signal to go back to IT with the
  exact error, not to go add Graph API permissions in Azure Portal.

- **Known, deliberate limitation**: `showAs: "free"` only affects the shared
  mailbox's own calendar copy of the event. Each attendee gets an independent
  copy when they accept, and Outlook defaults that copy to Busy — there is no
  single Graph call that makes an event show as Free on every attendee's
  calendar simultaneously. This was discussed and accepted, not an oversight —
  don't "fix" it without checking with the user first, since the alternative
  (patching each attendee's own event copy) requires Graph permissions scoped
  to every attendee's mailbox, a much bigger ask of IT.

- **Leftover scaffold, not part of the real app**: `workflows/sample_workflow.ts`,
  `functions/sample_function.ts` (+ `_test.ts`), `datastores/sample_datastore.ts`,
  `triggers/sample_trigger.ts`. These are the default `slack create` template
  and are unrelated to the scheduling feature. Don't extend them by habit when
  asked to add functionality — check whether the ask actually belongs in
  `create_calendar_event.ts` instead.

## Current blocking dependency

IT has responded with the finalized setup (mailbox, endpoint, permission
scoping, categories — see README's "Getting credentials from IT" section),
but the actual Tenant ID/Client ID/Client Secret values are not yet
necessarily in `.env` — verify `.env` is present and actually populated
before claiming an end-to-end test exercised the real Graph API call. Also
note: give it a few minutes after `.env` is first populated before the very
first real API call, since IT's permission scoping takes a short time to
propagate.

## Operational considerations once deployed (see README for detail)

`slack run` is local-dev-only (needs this machine/terminal running) — real
usage requires `slack deploy` to Slack-hosted infra, plus `slack env add` for
each production env var (separate from local `.env`). Client secret expires
2027-07-07 (IT has their own rotation reminder) — an autonomous deployment
will fail silently post-expiration until `MS_CLIENT_SECRET` is rotated.
