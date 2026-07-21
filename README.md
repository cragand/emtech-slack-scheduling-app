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
      First live test surfaced a real bug (Slack's rich-text auto-linkifying
      added a `mailto:` prefix that silently broke delivery); fixed and
      confirmed working live afterward
- [x] Fixed a case-sensitivity bug in the request-type-to-category lookup (the
      dropdown's actual `4X10 OOTO` value didn't match the table's `4x10 OOTO`
      key) — matching is now case-insensitive, see
      [Outlook categories](#outlook-categories)
- [x] Split `submitter_alias` into `amazon_alias` (from the roster lookup) and
      `submitter_name`, and reordered the title — see
      [Event title](#event-title). Confirmed working with a real live run
- [x] Deployed to Slack-hosted infra (`slack deploy`) — the app runs as
      **"Emtech Scheduling"** (no `(local)` suffix) in Emtech, LLC, with
      production env vars set via `slack env add --app <deployed app ID>` (see
      [Environment variables](#environment-variables)). The workflow's step now
      points at the deployed function, not the local dev one. Confirmed working
      with a real live run
- [x] Split the single catch-all workflow into purpose-specific workflows (Sick,
      Vacation, OOTO, 4x10 OOTO) so only some request types require approval —
      see [Purpose-specific workflows](#purpose-specific-workflows)
- [x] Added `is_partial_day` so OOTO requests can cover part of a day instead of
      always being full-day — see [Partial-day requests](#partial-day-requests).
      Confirmed working with a real live run
- [x] Added `is_recurring`/`recurrence_end_date`/`recurrence_day_of_week` so
      4x10 OOTO requests can recur weekly instead of requiring a fresh
      submission every week — see
      [Recurring requests (4x10 OOTO)](#recurring-requests-4x10-ooto). Confirmed
      working with a real live run, generating correct occurrences across
      multiple actual weeks, not just at initial creation
- [x] All four purpose-specific workflows (Sick, Vacation, OOTO, 4x10 OOTO) live
      and confirmed working end-to-end. Feature-complete; currently awaiting
      management review and approval
- [x] Added a second function, `post_daily_schedule_digest`, for company-wide
      visibility into the shared calendar — see
      [Daily schedule digest](#daily-schedule-digest). Built and unit-tested;
      not yet wired to a live scheduled workflow or confirmed with a real run

This app is fully deployed and running live — no local `slack run` process needs
to stay on for it to work. See
[Operational considerations](#operational-considerations-for-liveautonomous-use)
for ongoing maintenance notes (secret rotation, etc.).

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
non-interactive shells; `--hide-triggers` skips a trigger-creation prompt (this
app has no triggers of its own — the workflow is triggered from within Workflow
Builder, not a Slack CLI trigger).

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
these with `slack env add <NAME>` instead (omit the value so the CLI prompts for
it, keeping secrets out of shell history) — `.env` is git-ignored and never used
once deployed.

**This project has two separate app identities** — a dev/local one
(`.slack/apps.dev.json`) and the deployed production one (`.slack/apps.json`).
`slack env add` without an explicit `--app <app ID>` flag targets the dev app,
_not_ the deployed one — a real deploy surfaced this: all 5 vars were set
successfully but on the wrong app, so the live step failed with a missing-var
error until they were redone with `--app <deployed app ID>` (find the deployed
app's ID from `slack deploy`'s output, or `.slack/apps.json`). Use
`slack env list --app <app ID>` to check which variables (by name only, values
stay masked) exist on a given app — handy for confirming this without exposing
any secret.

A second real issue on the same deploy: pasting a value into the `slack env add`
prompt somehow doubled it (e.g. `MS_TENANT_ID` got saved as the real tenant ID
repeated twice back-to-back, which Microsoft's `AADSTS900023` error will call
out clearly if it happens again). If a Graph call fails with a "tenant/client
identifier is not valid" error, check whether the corresponding value looks
doubled before assuming it's a permissions problem — redo just that one
variable, watching the terminal to make sure only one clean paste lands before
hitting enter.

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

## Event title

Format: `@<amazon_alias> - <request_type> - <date> - <submitter_name>`, e.g.
`@jdoe - Sick - Aug 1 - John Doe`. The `@` always belongs on `amazon_alias`,
never on `submitter_name`. The date deliberately omits the year — it's always
obvious from the calendar page the event appears on.

`amazon_alias` and `submitter_name` are two separate inputs, not one —
`amazon_alias` comes from the company roster lookup (Workflow Builder step 3,
the same lookup that provides "Internal OOTO Recipients"), while
`submitter_name` is the person's actual name. These used to be a single field
(`submitter_alias`) serving double duty; it was split once a real Amazon Alias
became available as its own piece of roster data.

`submitter_name` is sometimes sourced from a Slack Person-type variable, which
renders as `@Display Name` (with the `@` baked into the string) when
interpolated into a plain-text field — the same category of formatting
bleed-through as the `mailto:` issue with external attendees. `stripLeadingAt()`
strips a leading `@` from both `amazon_alias` and `submitter_name` before
building the title, so the `@` always ends up in exactly the right place
regardless of which input happened to carry one in already.

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

The lookup is case-insensitive (and trims surrounding whitespace) — a real
request failed because the dropdown's actual value (`4X10 OOTO`) didn't match
this table's key (`4x10 OOTO`) under an exact-case comparison. The table's keys
can stay whatever casing is convenient to read; matching doesn't depend on it
lining up exactly with the dropdown.

## Purpose-specific workflows

Rather than one catch-all workflow handling every request type, each request
type now has its own Workflow Builder workflow (Sick, Vacation, OOTO, 4x10 OOTO)
— done so some request types can skip an approval step entirely (a sick employee
doesn't need approval to be sick), which a single shared workflow couldn't
express. All of them call the same `create_calendar_event` step; they just map
different subsets of its inputs. Sick and Vacation only use the original fields.
OOTO additionally supports partial-day requests. 4x10 OOTO additionally supports
weekly recurrence. Every new input added for these is optional, so adding one
doesn't require touching the other workflows' step configurations at all.

## All-day events

Every event this step creates is a full-day calendar entry (Graph's
`isAllDay: true`) by default, regardless of request type — there's no clock-time
component, matching how absence/status requests are actually used. The OOTO
workflow can override this per request via `is_partial_day` — see
[Partial-day requests](#partial-day-requests). A couple of details worth knowing
about the all-day path itself:

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
  never to determine which date it is in the first place. This only applies to
  values from a plain **Date** field — see
  [Partial-day requests](#partial-day-requests) for the other format this code
  now has to handle.

## Partial-day requests

The OOTO workflow supports `is_partial_day` ("Yes"/"No", from a dropdown — not a
boolean; see below) so a request can cover only part of a day instead of always
being a full-day event. When true, the event is created with `isAllDay: false`
and real clock times instead of midnight-to-midnight.

**`is_partial_day` is a string, not a boolean, on purpose.** Boolean-typed step
inputs can't be mapped to a per-submission variable in Workflow Builder's step
configuration — they only render as a fixed checkbox, the same value for every
run regardless of what the form collected. Switched to a string sourced from a
Yes/No dropdown instead, matched case-insensitively via `isPartialDayRequest()`,
the same pattern used elsewhere in this code for exact-case mismatches (see
[Outlook categories](#outlook-categories)).

**Because partial-day requests need real clock times, not just a date, the OOTO
workflow's `start_date_time`/`end_date_time` fields use Workflow Builder's "Date
and time" field type instead of the plain "Date" field Sick/Vacation/4x10 use —
and this sends a completely different raw format: a bare Unix timestamp (e.g.
`"1784235600"`) instead of an ISO string.** This surfaced as a real crash the
first time it was tested (`Expected an ISO date string, got "1784235600"`), and
broke _every_ OOTO submission, partial-day or not, since
`start_date_time`/`end_date_time` are Date-and-time fields unconditionally on
that workflow. Unlike the plain Date field's midnight-UTC quirk (never
timezone-convert it — see above), this Unix timestamp is a genuine,
correctly-computed real instant, confirmed against a live submission (a 2:00
PM–7:00 PM Pacific request came through as exactly `21:00:00Z`–`02:00:00Z` the
next day UTC) — so it needs the _opposite_ handling: converting it into
`MS_EVENT_TIMEZONE` to recover the wall-clock date/time that was actually
picked. `create_calendar_event.ts` now detects which of the two formats it
received and applies the correct handling for each.

**A "No" answer to `is_partial_day` is always safe, even if the underlying
Date-and-time field still carries specific hours** (e.g. someone's normal 9:00
AM–5:30 PM shift) — the all-day code path never reads the time-of-day portion at
all when `is_partial_day` is false, regardless of which raw format it came from.
Confirmed via test and live use.

## Recurring requests (4x10 OOTO)

The 4x10 OOTO workflow supports a weekly-recurring day off, for the common case
of someone consistently taking the same day off every week, so they don't have
to submit a fresh request every single week. Three optional inputs, all only
mapped by this workflow:

- `is_recurring` — "Yes"/"No" (a string, same reasoning as `is_partial_day`
  above), matched via `isRecurringRequest()`
- `recurrence_end_date` — the last date the series should generate through,
  inclusive. Deliberately **always required (never open-ended/"no end")** — an
  indefinitely recurring day-off doesn't account for someone's schedule
  changing, so every recurring request has a bounded horizon and needs
  resubmitting to continue past it
- `recurrence_day_of_week` — which weekday recurs, sourced from a single-select
  day-of-week dropdown already in the form, matched case-insensitively via
  `normalizeDayOfWeek()` against the seven day names

`start_date_time`/`end_date_time` are still required and both mapped to the same
single "day off" date field (the same trick a single-day Sick/Vacation request
already uses) — this workflow keeps the plain **Date** field type throughout,
never Date-and-time, since it never needs partial-day logic. `is_partial_day`
and `is_recurring` are mutually exclusive; the step rejects a request that sets
both.

**`start_date_time`'s weekday doesn't need to match `recurrence_day_of_week`.**
Graph's recurrence `range.startDate` just tells it "start generating occurrences
from here forward" — it finds the first actual matching weekday on or after that
date itself, so these are two independent pieces of information, not two values
that could disagree the way the partial-day time fields could have (see
`buildWeeklyRecurrence()`'s tests).

**No self-service way to change or cancel a recurring request yet** — this app
only ever creates events, with no update/cancel capability at all. For now,
changing a schedule or ending a series early means manually editing it directly
in Outlook (Andrew and Simon have full access to the shared mailbox). Worth
revisiting if this becomes a frequent need.

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

**A real bug already happened here, now fixed:** Slack's List Text column
auto-linkifies email-looking text, and whatever serialization turns that
rich-text value into a plain string variable carries the `mailto:` scheme along
with it (e.g. `mailto:name@x.com` instead of `name@x.com`). This still passed
our own validation (the pattern doesn't exclude colons) and Graph accepted the
event creation without error, but Exchange never actually delivered an invite to
the real address underneath — confirmed live: the event showed "N emails are
invalid" in Outlook and the external attendees never received anything.
`parseEmailList()` now strips a leading `mailto:` (case-insensitive) from each
entry before validating it.

## Known limitation: attendees' free/busy status

The event is created with `showAs: "free"`, but this only controls the free/busy
status on the **shared mailbox's own calendar copy**. Each invited attendee gets
their own separate copy of the event once they accept, and Outlook defaults that
copy to **Busy** regardless of the organizer's `showAs` setting — there's no
Graph API call that makes it "Free" on every attendee's calendar at once. This
was a deliberate tradeoff (real Outlook invites were chosen over a mailbox-only
entry) — see `functions/create_calendar_event.ts` for where this is implemented.

## Daily schedule digest

`functions/post_daily_schedule_digest.ts` is a second, independent function (not
a step used by any of the four request workflows) for company-wide visibility
into the shared calendar. It's meant to be wired to its own Workflow Builder
workflow using a native **"On a schedule"** trigger (daily), not a form — the
function itself does everything (reads the calendar, posts the message), so that
workflow is just the schedule trigger plus this one step.

Each run reads today's events from the shared mailbox via Graph's `calendarView`
endpoint (not a plain `/events` GET — `calendarView` is what correctly expands a
recurring 4x10 series into that week's specific occurrence, rather than
returning just the series master), computing "today" in `MS_EVENT_TIMEZONE`
rather than trusting server-local time. It posts a bulleted, clickable list —
each line is the event's existing `subject` (the same title
`create_calendar_event.ts` already built at creation time) wrapped as a link to
that event's own `webLink`, e.g.:

```
Today's schedule:
• <https://outlook.office.com/.../abc|@jsmith - Vacation - Jul 20 - Jane Smith>
• <https://outlook.office.com/.../xyz|@jdoe - Sick - Jul 20 - John Doe>
```

If nothing is scheduled, it posts "Nothing on the schedule today." rather than
staying silent — so the daily run being alive is itself visible, not just
inferred from an absence of messages.

Chosen over two alternatives, both discussed and set aside:

- A per-request delay (Workflow Builder's native Delay step is capped at 7 days,
  too short for a vacation submitted months out; a dynamically-created scheduled
  trigger per request has no cap, but would need its own storage and cleanup
  logic to cancel that trigger if the request is later edited or canceled — a
  real feature in itself, not yet built)
- Sick's existing "post immediately" behavior, which works because Sick is
  always for the day it's submitted — not true for Vacation/OOTO, which can be
  submitted well in advance

This design has nothing per-request to track, cancel, or clean up — it just
re-reads the calendar fresh every day, so an edited or canceled request is
simply reflected (or absent) in the next digest, automatically.

No new Microsoft/Azure credentials are needed — this function reuses the same
Graph access `create_calendar_event.ts` already has (`Calendars.ReadWrite`
inherently includes read, not just write). The shared OAuth token logic was
extracted to `functions/lib/graph_auth.ts` so both functions use the exact same
implementation rather than two copies silently drifting apart.

To stop the daily posts, remove (or deactivate, if reversibility matters) the
one Workflow Builder workflow using this function — nothing else references it,
so nothing else is affected either way.

## Testing

```sh
deno test
```

`functions/create_calendar_event_test.ts` covers the title formatting, the
request-type-to-category mapping, `parseEmailList()`'s delimiter handling and
invalid-entry detection, the exact request sent to Microsoft Graph (subject,
attendees, `showAs`, `categories`), attendee resolution (internal + external
combined, and the best-effort skip-and-notify behavior when an internal attendee
can't be resolved), the partial-day and recurring-event logic (both ISO-string
and Unix-timestamp date/time formats, the mutual-exclusion check between them,
and the missing/invalid-field error paths), and the missing-env-var /
unrecognized-request-type / invalid-external-attendee / token-failure error
paths — all with `fetch` mocked via `@std/testing/mock` (including Slack's own
Web API calls), so no live credential or network call is needed to run it.

`functions/post_daily_schedule_digest_test.ts` covers the digest message
formatting (bulleted links, the missing-webLink/subject fallbacks, the "nothing
scheduled" case), and the handler's Graph `calendarView` call and
`chat.postMessage` call, plus its missing-env-var and Graph-failure error paths
— same mocking approach, no live call needed.

## Deploying

Once verified locally, deploy to Slack-hosted infrastructure:

```sh
slack deploy
```

Set production environment variables first with `slack env add`, since `.env` is
only used for local `slack run`.

After the first deploy, swap the workflow's step from the `(local)` version of
"Create scheduling calendar event" to the deployed version in Workflow Builder,
re-mapping every input — Slack doesn't carry input mappings over between steps
automatically.

## Resources

- [Slack automation overview](https://api.slack.com/automation)
- [Creating a custom step for Workflow Builder (Deno SDK)](https://docs.slack.dev/tools/deno-slack-sdk/tutorials/workflow-builder-custom-step/)
- [Microsoft Graph — create event](https://learn.microsoft.com/en-us/graph/api/user-post-events)
- [Microsoft identity platform — client credentials flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow)
