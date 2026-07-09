# Emtech Scheduling

A Slack app (Deno/next-gen platform) that adds a custom Workflow Builder step,
**"Create scheduling calendar event,"** which creates a calendar event on a
shared Microsoft 365 mailbox via the Microsoft Graph API.

## How it fits together

The actual workflow logic — the trigger, the form, the Slack List — lives in
**Slack Workflow Builder** (no-code), not in this repo. This repo provides one
building block that workflow uses as a step: a function that, given a submitter,
request details, and a list of attendees, calls Microsoft Graph to put an event
on the shared scheduling mailbox's calendar.

```
Slack List field updated
        │
        ▼
Workflow Builder workflow (built in Slack UI, not in this repo)
        │
        ▼
"Create scheduling calendar event" step  ──►  functions/create_calendar_event.ts
                                                      │
                                                      ▼
                                        Microsoft Graph API (POST /users/{mailbox}/events)
                                                      │
                                                      ▼
                                        Event appears on shared scheduling mailbox's calendar,
                                        attendees (submitter + additional) get an Outlook invite
```

## Status

- [x] Slack app scaffolded, authenticated to the `emtechllc` workspace
- [x] `create_calendar_event` function written and registered as a Workflow
      Builder step
- [x] Step added to the real workflow in Workflow Builder, inputs mapped
- [x] IT set up the Azure AD app registration, scoped to the shared mailbox via
      an Exchange Application Access Policy — see
      [Getting credentials from IT](#getting-credentials-from-it) for the final
      setup details
- [x] Real Tenant ID / Client ID / Client Secret values added to `.env`
- [x] End-to-end test against the real shared mailbox — title, category color,
      free/busy, all-day range, and multi-day date handling all confirmed
      correct
- [x] `additional_attendees` reworked to resolve Slack user references (a List
      Person column) to real email addresses — see
      [Additional attendees](#additional-attendees). Confirmed working with a
      real live run (Internal OOTO Recipients)
- [x] `external_attendees` added for non-Slack contacts, parsed from a delimited
      string (a plain **Text** List column, "External OOTO Recipients" — the
      Email-type column was tried first and turned out to not support saving
      more than one value) — see [External attendees](#external-attendees).
      Built and unit-tested; not yet confirmed with a real live run
- [ ] Deploy to Slack-hosted infra (`slack deploy`) for real/autonomous use —
      still only running via local `slack run` today

Real end-to-end testing is working locally via `slack run`. See
[Operational considerations](#operational-considerations-for-liveautonomous-use)
for what's still needed before this runs unattended in production.

## Prerequisites

- [Slack CLI](https://api.slack.com/automation/quickstart), already
  authenticated to the `emtechllc` workspace
- [Deno](https://deno.com) (installed via `winget install DenoLand.Deno` on this
  machine)

> **Windows/Git Bash note:** if `slack` commands hang with no output at all,
> check `deno --version` in that same shell first — the Slack CLI shells out to
> Deno for most commands and fails silently (or hangs) if it's not on `PATH`.
> Also, some `slack` commands (`doctor`, `run` without `--team`) don't play well
> with Git Bash's pty and will hang waiting on a prompt that can't be answered;
> prefer running them from a native PowerShell/cmd terminal.

## Running locally

```sh
cd emtech-scheduling-slack-app
slack run --team emtechllc --hide-triggers
```

You'll know it's running when you see `Connected, awaiting events`. While it's
running, a local version of the app (`Emtech Scheduling (local)`) is installed
in the workspace and its step is selectable in Workflow Builder.

`--team emtechllc` skips an interactive workspace-selection prompt that fails in
non-interactive shells; `--hide-triggers` skips prompts related to the unused
sample trigger (see [Leftover scaffold files](#leftover-scaffold-files)).

## Environment variables

Copy `.env.example` to `.env` and fill in real values once IT provides them:

| Variable            | Purpose                                                                                                   |
| ------------------- | --------------------------------------------------------------------------------------------------------- |
| `MS_TENANT_ID`      | Identifies Emtech's Azure AD directory                                                                    |
| `MS_CLIENT_ID`      | Identifies this app's registration within that directory                                                  |
| `MS_CLIENT_SECRET`  | The app's credential, used to authenticate as itself (no user login involved)                             |
| `MS_SHARED_MAILBOX` | The shared scheduling mailbox: `teamavailability@emtech.us` ("Team Availability")                         |
| `MS_EVENT_TIMEZONE` | IANA timezone used for the event's start/end: `America/Los_Angeles` (Emtech operations are Pacific-based) |

`slack run` automatically picks up `.env` locally. For the deployed app, set
these with `slack env add <NAME> <VALUE>` instead — `.env` is git-ignored and
never used once deployed.

### Getting credentials from IT

[`IT_REQUEST.md`](./IT_REQUEST.md) is the original ask sent to IT — kept as a
historical record. Here's what IT actually set up in response:

- **Mailbox**: `teamavailability@emtech.us` ("Team Availability"). No Microsoft
  365 license was needed for this — turned out to be unnecessary for a shared
  mailbox used this way, despite what the original request assumed. Andrew
  Crager and Simon Newkirk have full access (it auto-appears in their own
  Outlook); everyone else in the company has read-only view access to events on
  it.
- **Endpoint**:
  `POST https://graph.microsoft.com/v1.0/users/teamavailability@emtech.us/events`,
  authenticated via the standard OAuth2 client-credentials flow — matches what
  `getGraphAccessToken()` already does.
- **Access scoping**: IT did **not** grant `Calendars.ReadWrite` as a
  tenant-wide application permission with admin consent. Instead, they scoped
  the app's access at the Exchange level (an Application Access Policy) so it
  can only ever touch this one mailbox. **Do not add `Calendars.ReadWrite` as an
  API permission or request admin consent on the app registration** — per IT,
  that would override their scoping and open access to every mailbox in the
  tenant. (This repo has no mechanism to do that anyway — Graph permissions are
  configured entirely in Azure Portal, not in this codebase — but flagging
  clearly so no one goes looking for a way to add it.)
- **Propagation delay**: give it a few minutes after setup before the first real
  API call — the permission scoping takes a short time to propagate on
  Microsoft's side.
- **Client secret expiration**: July 7, 2027. IT has their own reminder to
  rotate it before then — flag them if it gets rotated sooner than that for any
  reason, so their reminder stays accurate.

## Outlook categories

The event's `categories` field drives calendar color-coding. There's no separate
`category` input in the step — it's derived automatically from `request_type`
(the workflow's "Request Type" dropdown) via a lookup table in
`create_calendar_event.ts`, since the mapping is fully deterministic and a
manually-mapped extra field turned out to be redundant (and was the actual cause
of an early "invalid parameter" failure, since the step wasn't re-configured
after that field was added).

Current mapping (`REQUEST_TYPE_TO_CATEGORY` in `create_calendar_event.ts`):

| Request Type        | Outlook category                                                                      |
| ------------------- | ------------------------------------------------------------------------------------- |
| Sick                | `OOTO`                                                                                |
| Vacation            | `OOTO`                                                                                |
| OOTO                | `OOTO`                                                                                |
| 4x10 OOTO           | `4x10 OOTO`                                                                           |
| WFH                 | `WFH`                                                                                 |
| Location Assignment | `On-Site`                                                                             |
| Scheduling          | _(removed from the dropdown — handled via direct message instead, no calendar event)_ |

If the "Request Type" dropdown ever gets a new option,
`REQUEST_TYPE_TO_CATEGORY` needs a matching entry — otherwise the step fails
cleanly with `No Outlook
category mapping configured for request type "..."`
rather than silently mis-categorizing (or, worse, silently succeeding with the
wrong color). To add a new per-site category (e.g. `On-Site - Seattle`) to the
mailbox itself, send IT the site name and they'll configure it with a color.

## All-day events

Every event this step creates is a full-day calendar entry (Graph's
`isAllDay: true`), regardless of request type — there's no clock-time component,
matching how absence/status requests are actually used. A couple of details
worth knowing:

- **`end_date_time` is the _last_ day of the request, inclusive** — e.g. a
  single-day request has the same date for `start_date_time` and
  `end_date_time`. Graph itself requires an _exclusive_ end (the day _after_ the
  last day), so `getAllDayEventRange()` adds one calendar day internally before
  sending it to Graph. If the workflow's date field ever changes to already
  provide an exclusive end date, this would double-count by a day — check
  `getAllDayEventRange()`'s tests if that's ever suspected.
- **Time-of-day and any offset/`Z` suffix in the inputs are ignored entirely.**
  `start_date_time`/`end_date_time` represent a calendar date someone picked,
  not a real moment in time — the literal `YYYY-MM-DD` written in the string is
  taken as-is, with no timezone conversion applied. This is deliberate, and the
  opposite of what an earlier version of this code did: it tried to convert
  these values into `MS_EVENT_TIMEZONE` to decide "which day is this," which
  seems reasonable but is actually wrong for a date-only value — Slack's date
  field sends a plain date selection as midnight UTC (e.g.
  `2026-07-08T00:00:00Z` for "July 8"), and converting midnight UTC into
  `America/Los_Angeles` (hours behind UTC) rolls it back to the _previous_ day.
  That bug was real and shipped briefly — a two-day request submitted as 7/8–7/9
  came out as 7/7–7/8 on the actual calendar. `MS_EVENT_TIMEZONE` is still used,
  just only to tell Graph what "midnight" means for the constructed date range,
  never to determine which date it is in the first place.

## Operational considerations for live/autonomous use

- **Client secret expiration**: see
  [Getting credentials from IT](#getting-credentials-from-it) above for the
  actual expiration date.
- **Deployed env vars are separate from `.env`**: see
  [Environment variables](#environment-variables) above — `slack env add` has to
  be run once against the deployed app; local `.env` values don't carry over.
- `slack run` is local-dev-only and requires this machine/terminal to stay on;
  it is not how the app runs for real usage. See [Deploying](#deploying).

## Additional attendees

`additional_attendees` takes Slack user references (e.g. a List's "Person"
column — currently "Internal OOTO Recipients" on the "Emtech Company Roster"
list), not plain email addresses. Microsoft Graph has no concept of Slack users,
so each one is resolved to a real email address via Slack's `users.info` API
before being added to the event — this requires the
`users:read`/`users:read.email` bot scopes.

Resolution is **best-effort**: if a given attendee's email can't be found
(hidden email, guest account, etc.), the event is still created with whichever
attendees _did_ resolve, and the submitter gets a DM afterward naming who was
skipped, rather than the whole request failing over one unresolvable person. The
DM itself is also best-effort — if the submitter's own Slack account can't be
resolved from `submitter_email` either, the notification is silently skipped
rather than failing the (already successful) event creation.

A plain-string (email) version of this field was tried first and does work when
someone types a literal email address in — the failure is specifically with
Person-type values, confirmed by that manual test succeeding while a
Person-column-sourced value failed with a generic "invalid parameter" step
failure (the same unhelpful message Slack gives for any pre-execution parameter
type mismatch — nothing more specific is available for this failure class,
checked via both `slack activity` and Workflow Builder's own activity log).
Switching the List column itself to an Email type was also considered for
`additional_attendees`, since that would need zero code changes — set aside
since internal recipients were asked to stay Person type. (Turned out to be moot
anyway for multi-value Email columns generally — see
[External attendees](#external-attendees) below.)

### External attendees

`external_attendees` is a separate input for non-org contacts (clients, other
POCs) who have no Slack account to resolve from. It's a **single string**, not
an array — email addresses separated by whitespace, commas, or semicolons (any
mix), parsed by `parseEmailList()` and validated to look like real emails before
being added to the event. An entry that doesn't look like a valid email returns
a clean error naming which one, rather than silently sending something malformed
to Graph.

This is mapped to a plain **Text** column ("External OOTO Recipients"), not an
Email-type column, despite Slack's own API reference confirming the Email field
type is _supposed_ to be natively multi-value (`["a@b.com", "c@d.com"]`). In
practice, saving more than one value to an Email-type List field fails with a
generic "Failed to save changes! Please reload slack." error — reproduced
repeatedly, and not something reloading or re-entering fixes. The field happily
accepts and displays multiple pasted values right up until the actual save,
which points to a backend limitation/bug in that field type rather than anything
on our end. The Text-column-plus-parsing approach here is the workaround.

These get combined with the resolved internal attendees into one Graph
`attendees` list — external contacts receive the same real Outlook invite as
everyone else.

## Known limitation: attendees' free/busy status

The event is created with `showAs: "free"`, but this only controls the free/busy
status on the **shared mailbox's own calendar copy**. Each invited attendee gets
their own separate copy of the event once they accept, and Outlook defaults that
copy to **Busy** regardless of the organizer's `showAs` setting — there's no
Graph API call that makes it "Free" on every attendee's calendar at once. This
was a deliberate tradeoff (real Outlook invites were chosen over a mailbox-only
entry) — see `functions/create_calendar_event.ts` for where this is implemented.

## Testing

```sh
deno test
```

`functions/create_calendar_event_test.ts` covers the title formatting, the
request-type-to-category mapping, `parseEmailList()`'s delimiter handling and
invalid-entry detection, the exact request sent to Microsoft Graph (subject,
attendees, `showAs`, `categories`), attendee resolution (internal + external
combined, and the best-effort skip-and-notify behavior when an internal attendee
can't be resolved), and the missing-env-var / unrecognized-request-type /
invalid-external-attendee / token-failure error paths — all with `fetch` mocked
via `@std/testing/mock` (including Slack's own Web API calls), so no live
credential or network call is needed to run it.
`functions/sample_function_test.ts` is unrelated, from the default scaffold (see
below).

## Deploying

Once verified locally, deploy to Slack-hosted infrastructure:

```sh
slack deploy
```

Set production environment variables first with `slack env add`, since `.env` is
only used for local `slack run`.

## Leftover scaffold files

`workflows/sample_workflow.ts`, `functions/sample_function.ts` (+ its test),
`datastores/sample_datastore.ts`, and `triggers/sample_trigger.ts` are unused
template files from `slack create`, unrelated to the scheduling app. They're
left in place but can be deleted if the project needs cleaning up.

## Resources

- [Slack automation overview](https://api.slack.com/automation)
- [Creating a custom step for Workflow Builder (Deno SDK)](https://docs.slack.dev/tools/deno-slack-sdk/tutorials/workflow-builder-custom-step/)
- [Microsoft Graph — create event](https://learn.microsoft.com/en-us/graph/api/user-post-events)
- [Microsoft identity platform — client credentials flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow)
