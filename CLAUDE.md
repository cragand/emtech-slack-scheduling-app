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

# Set an env var for the deployed app (local dev uses .env instead) —
# omit the value so the CLI prompts for it, and always pass --app explicitly
# (see gotcha below); find the deployed app's ID via `slack deploy`'s output
# or .slack/apps.json
slack env add MS_TENANT_ID --app <deployed app ID>
```

**Two app identities, one project — `--app` matters for `env add`/`env list`:**
`.slack/apps.dev.json` is the dev/local app (`IsDev: true`); `.slack/apps.json`
is the deployed production app. `slack env add`/`slack env list` without an
explicit `--app <app ID>` targets the **dev** app, not the deployed one — this
already caused a real outage-in-waiting: all 5 Graph env vars were set
successfully but landed on the dev app, so the freshly-deployed production step
failed on every live run with a missing-env-var error until redone with
`--app <deployed app ID>`. Always pass `--app` explicitly for the deployed app
when asked to set or verify production env vars; `slack env list --app <id>`
is safe to run freely (names only, values stay masked).

**Pasted env var values can silently double.** A separate real incident on the
same deploy: `MS_TENANT_ID` got saved as the real GUID repeated twice
back-to-back after a paste into the `slack env add` prompt. Graph surfaced this
clearly (`AADSTS900023: ... is neither a valid DNS name, nor a valid external
domain`, with the doubled value visible in the error text) — if a similar
"tenant/client identifier not valid" error comes back after credentials were
just set, suspect a doubled paste before suspecting a permissions problem.

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

- **`manifest.ts`** — registers both `CreateCalendarEventDefinition` and
  `PostDailyScheduleDigestDefinition` under `functions:` (not `workflows:`).
  That's the specific field that makes a function selectable as a step inside
  Slack's no-code Workflow Builder; functions only referenced via a
  code-defined `DefineWorkflow` wouldn't show up there. `outgoingDomains` must
  list `login.microsoftonline.com` and `graph.microsoft.com` — the Deno SDK
  sandboxes `fetch` to declared domains.

- **`functions/lib/graph_auth.ts`** — `getGraphAccessToken()`, shared by both
  functions below. Extracted here specifically to avoid two copies of the
  same OAuth logic drifting apart once a second function needed it — don't
  duplicate it back into either function file.

- **`functions/post_daily_schedule_digest.ts`** — a second, independent
  function, not a step used by any of the four request workflows. Meant to
  be wired to its own Workflow Builder workflow via a native **"On a
  schedule"** trigger (daily), not a form — the function does everything
  itself (reads Graph, posts to Slack), so nothing else needs building on the
  Workflow Builder side beyond that one trigger + this one step.
  - Uses Graph's `calendarView` endpoint, not a plain `/events` GET —
    `calendarView` is what correctly expands a recurring 4x10 series into
    that week's specific occurrence (with its own `webLink`), rather than
    returning just the series master. Don't switch this to `/events` without
    checking recurring events still show up correctly.
  - Computes "today" via `Intl.DateTimeFormat` in `MS_EVENT_TIMEZONE`, not
    server-local time or a raw UTC day boundary — same category of care as
    `create_calendar_event.ts`'s date handling, just computing "now" instead
    of parsing a submitted value.
  - `channel_id` is a required input (`Schema.slack.types.channel_id`) —
    typically set as a fixed value in the step config (the scheduling
    channel), not something a human answers, same pattern as the fixed
    `request_type`/`is_recurring` values on the 4x10 workflow.
  - **Chosen over a per-request delayed message deliberately** — Workflow
    Builder's native Delay step caps at 7 days (too short for a vacation
    submitted months out), and a dynamically-created scheduled trigger per
    request has no cap but would need its own storage + cancellation logic
    if a request is later edited/canceled (a real feature not yet built).
    This design has nothing per-request to track or clean up — it just
    re-reads the calendar fresh every day. Don't "simplify" this back to a
    per-request delay without re-solving that cancellation problem.
  - If nothing's scheduled, it posts `"Nothing on the schedule today."`
    rather than staying silent, so the daily run itself being alive is
    visible, not just inferred from an absence of messages.

- **`functions/create_calendar_event.ts`** — the one file that matters most.
  - `input_parameters`: `amazon_alias`, `submitter_name`, submitter email,
    request type, start/end ISO 8601 datetimes, description, location,
    `additional_attendees`, and `external_attendees` (mapped in Workflow
    Builder from the triggering Slack List row / form fields).
    `amazon_alias` is sourced from the company roster lookup (Workflow
    Builder step 3), not the initial form — distinct from `submitter_name`,
    which is the person's actual name. This used to be a single field
    (`submitter_alias`) doing double duty as both the title's first and only
    identifying segment; it was split into two once a real Amazon Alias
    became available as separate data.
  - **`additional_attendees` is `Schema.slack.types.user_id[]`, not plain
    email strings — this was deliberate, confirmed by testing, don't revert
    it.** A plain-string version was tried first (and does work when a
    literal email is typed in), but the real data source is a List's Person
    column ("Internal OOTO Recipients"), which gives a Slack user ID, not an
    email — Microsoft Graph has no concept of Slack users, so each ID is
    resolved via `client.users.info` before being added as a Graph attendee.
    This needs the `users:read`/`users:read.email` bot scopes (in
    `manifest.ts`). Resolution is best-effort — an attendee whose email can't
    be found is skipped (event still created with whoever did resolve)
    rather than failing the whole step, and the submitter gets a best-effort
    DM afterward (via `users.lookupByEmail` on `submitter_email`, then
    `chat.postMessage`) naming who was skipped. Confirmed working end-to-end
    with a real live run.
  - **`external_attendees` is a single `Schema.types.string`, not an
    array** — for non-org contacts with no Slack account to resolve from. It
    maps to "External OOTO Recipients," a plain **Text** List column, parsed
    by `parseEmailList()` (splits on any mix of whitespace/commas/semicolons,
    validates each entry looks like an email, returns `{ error }` naming the
    first bad one otherwise). This is NOT the original design — an array of
    strings mapped to an Email-type List column was tried first (matching
    Slack's own API reference, which documents that field type as natively
    multi-value: `["a@b.com", "c@d.com"]`), but saving more than one value to
    an Email-type field fails in the actual product with a generic "Failed to
    save changes!" error, reproduced repeatedly — a real gap between the
    documented data model and what the UI/backend actually supports as of
    this writing. Don't revert to the array/Email-column version without
    first confirming Slack has fixed that. Combined with the resolved
    internal emails into one Graph `attendees` list. First live test
    surfaced a real bug (see next bullet); not yet re-confirmed live since
    the fix. See README's "Additional attendees" / "External attendees"
    sections for the full reasoning.
  - **`parseEmailList()` strips a leading `mailto:` (case-insensitive) from
    each entry — don't remove this.** A real bug shipped without it: Slack's
    List Text column auto-linkifies email-looking text, and whatever
    serialization turns that rich-text value into a plain string variable
    carries the `mailto:` scheme along with it. `EMAIL_PATTERN` doesn't
    exclude colons, so `mailto:name@x.com` passed validation and Graph
    accepted the event with no error — but Exchange never delivered an
    invite to the real address underneath. Confirmed live (before the fix):
    Outlook showed "N emails are invalid" and the external attendees
    received nothing.
  - There is **no separate `category` input** — a first attempt added one,
    but it caused a real "failed to start due to an invalid parameter"
    failure in production because the Workflow Builder step wasn't
    re-configured to map the new required field after it was added. It also
    turned out unnecessary: the request-type-to-category mapping is fully
    deterministic (`Sick`→`OOTO`, `Vacation`→`OOTO`, `OOTO`→`OOTO`,
    `4x10 OOTO`→`4x10 OOTO`, `WFH`→`WFH`, `Location Assignment`→`On-Site`),
    so `getCategoryForRequestType()` / `REQUEST_TYPE_TO_CATEGORY` derives it
    from `request_type` in code instead. If the "Request Type" dropdown in
    Workflow Builder ever gets a new option, this lookup table needs a
    matching entry (see README's "Outlook categories" section) — an unknown
    `request_type` returns a clean `{ error: "No Outlook category mapping
    configured..." }` rather than guessing or silently omitting the category.
    **The lookup is case-insensitive (and trims whitespace) — don't switch
    it back to an exact-case object-key lookup.** A real request failed
    because the dropdown's actual value (`4X10 OOTO`) didn't match this
    table's key (`4x10 OOTO`) under exact-case comparison.
  - `getGraphAccessToken()`: does the OAuth2 client-credentials POST to
    `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token` to get an
    app-only bearer token — no user session involved, since the workflow runs
    unattended.
  - Handler builds the title as
    `@<amazon_alias> - <request_type> - <date> - <submitter_name>`, builds
    the attendee list (submitter + additional attendees, all as Graph
    `attendees` so they get real Outlook invites), and POSTs to
    `/v1.0/users/{MS_SHARED_MAILBOX}/events` with `showAs: "free"` and
    `categories: [category]` (category derived as above).
  - **`stripLeadingAt()` strips a leading `@` (and following whitespace) from
    both `amazon_alias` and `submitter_name` before building the title — the
    literal `@` in the title is always added in code, on `amazon_alias`
    specifically.** `submitter_name` is sometimes sourced from a Slack
    Person-type variable, which renders as `@Display Name` when interpolated
    into a plain string (same category of bug as the `mailto:` issue) — a
    real request came through with `submitter_name` carrying the `@` and
    `amazon_alias` not, producing the wrong title. Stripping from both sides
    makes the output correct regardless of which input happens to carry one.
  - **`start_date_time`/`end_date_time` arrive in two genuinely different
    raw formats depending on which Workflow Builder field type sourced
    them, and `resolveDateTimeParts()` has to detect which one it got —
    don't assume either format universally.**
    1. Plain **Date** field (Sick, Vacation, 4x10 OOTO, and OOTO's own
       all-day case): an ISO string like `"2026-07-08T00:00:00Z"` — a
       calendar date someone picked, not a real moment in time, even though
       it's dressed up as midnight UTC. **Never timezone-convert this one.**
       Two real bugs happened before landing on this, in this order — don't
       re-introduce either:
       1. First attempt: `formatTitleDate` used `new Date(isoDateTime)` +
          default (system-ambient) timezone formatting. Wrong day near a
          midnight boundary.
       2. "Fix": ran the value through `Intl.DateTimeFormat` with
          `MS_EVENT_TIMEZONE`. Made things *worse* — converting
          midnight-UTC to `America/Los_Angeles` (hours behind UTC) rolls it
          back to the *previous* day, so both the title and the actual
          Graph event landed exactly one day early for every request
          (confirmed live: submitted 7/8–7/9, got 7/7–7/8).
       3. Actual fix: stop timezone-converting ISO-string values at all —
          extract the literal `YYYY-MM-DD`/`HH:MM:SS` digits via regex,
          full stop. `timeZone` is still needed, but only paired with the
          constructed strings when talking to Graph, never used to decide
          *which day* something is.
    2. **"Date and time"** field (OOTO's `start_date_time`/`end_date_time`,
       used for partial-day support): a bare Unix timestamp, e.g.
       `"1784235600"` — genuinely a real, correctly-computed instant this
       time, confirmed against a live submission (a 2:00 PM–7:00 PM Pacific
       request came through as exactly `21:00:00Z`–next-day `02:00:00Z`).
       **This one needs the opposite handling of the ISO-string case: it
       must be converted into `MS_EVENT_TIMEZONE`** (via
       `Intl.DateTimeFormat.formatToParts`) to recover the wall-clock
       date/time actually picked. Real bug: this wasn't handled at all
       originally, and every OOTO submission (partial-day or not, since
       OOTO's fields are Date-and-time unconditionally) crashed with
       `Expected an ISO date string, got "1784235600"` until
       `resolveDateTimeParts()` learned to detect and handle both formats.
  - **`is_partial_day`/`is_recurring` are strings ("Yes"/"No"), not
    booleans — don't change these back to `Schema.types.boolean`.** A
    boolean-typed step input can't be mapped to a per-submission variable in
    Workflow Builder's step configuration; it only renders as a fixed
    checkbox, the same value for every run regardless of what the form
    collected. Discovered live when `is_partial_day` was first built as a
    boolean and the checkbox had no way to map a variable at all. Matched
    case-insensitively via `isPartialDayRequest()`/`isRecurringRequest()`,
    same pattern as `getCategoryForRequestType()`.
  - `is_partial_day` (OOTO only): when true, `getPartialDayEventRange()` is
    used instead of `getAllDayEventRange()` — real clock times,
    `isAllDay: false`. When false/absent, the all-day path never even reads
    the time-of-day portion of `start_date_time`/`end_date_time`, so it's
    safe regardless of what hours happen to be embedded in the raw value
    (confirmed: a full-day request with real shift hours like 9am–5:30pm
    still produces a correct all-day event when `is_partial_day` is "No").
  - `is_recurring`/`recurrence_end_date`/`recurrence_day_of_week` (4x10 OOTO
    only): when true, `buildWeeklyRecurrence()` attaches a Graph `recurrence`
    object (weekly pattern + bounded `endDate` range) to an otherwise-normal
    all-day event. **`recurrence_end_date` is always required when
    `is_recurring` is true — never build a "no end" / indefinite option** —
    an open-ended recurring day-off doesn't account for a schedule changing,
    so every recurring request needs a bounded horizon and resubmission to
    continue past it. `recurrence_day_of_week` is matched case-insensitively
    against the seven day names via `normalizeDayOfWeek()`, sourced from an
    existing single-select dropdown — it does **not** need to match
    `start_date_time`'s actual weekday; Graph's `range.startDate` just
    anchors "start generating from here forward" and finds the first
    matching weekday itself. `is_partial_day` and `is_recurring` are
    mutually exclusive — the handler rejects a request that sets both
    rather than picking one silently. No self-service update/cancel
    capability exists for a recurring series yet — see README's "Recurring
    requests (4x10 OOTO)" section.
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

## Current status

This app is fully deployed and running live (`slack deploy`, production app
`A0BGD5F67RQ` in Emtech, LLC) — not blocked on anything. Treat it as a
production app when making changes, not a project still being set up.
Verify against README's Status checklist before assuming otherwise.

## Operational considerations once deployed (see README for detail)

`slack run` is local-dev-only (needs this machine/terminal running), used for
local iteration before `slack deploy`ing a change to production — it is not
how the app runs day-to-day. Client secret expires 2027-07-07 (IT has their
own rotation reminder) — an autonomous deployment will fail silently
post-expiration until `MS_CLIENT_SECRET` is rotated.
