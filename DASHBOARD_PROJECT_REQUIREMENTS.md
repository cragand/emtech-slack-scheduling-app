# Scheduling dashboard — requirements summary (draft)

Status: **exploratory** — this describes a possible future project, not yet
started, not yet approved. It's a separate system from this repo's Slack app,
though the two would integrate. This document exists to scope what it would
actually take, so it can be sized and prioritized.

## What it's for

A hosted web dashboard where scheduling requests, responses, and assignments
(the same data currently flowing through Slack Lists and the "Emtech
Scheduling" Slack app) can be viewed and managed outside of Slack. Discussed
so far as potentially covering:

- Viewing scheduling requests/events (submitted, pending, assigned)
- Reviewing what's landed on the shared "Team Availability" calendar
- A possible phase 2: automatically syncing individual employees' Slack status
  (OOTO/Sick/Vacation/WFH/etc.) based on events on that shared calendar

## Architecture: a separate app, not a step in the existing workflow

This is a fully separate application from the "Emtech Scheduling" Slack app in
this repo — its own codebase, its own AWS-hosted backend/database/login, not
a Deno function and not a Workflow Builder step. It is not possible (and not
desirable) to build this "inside" the `create_calendar_event` custom step —
that step is Slack-hosted, single-purpose, and runs as one function
invocation per request; a multi-page dashboard with its own storage and auth
can't live inside that.

The two systems only ever exchange data at two loosely-coupled points:

1. `create_calendar_event.ts` optionally makes one additional outbound `POST`
   (the webhook) after successfully creating an event — a small addition to
   the existing step, not a merge of the two codebases.
2. The dashboard independently reads Microsoft Graph and/or Slack Lists on its
   own — this requires zero changes to the existing step.

Everything else in this document (AWS hosting, the dashboard's own database,
its own login) is scoped entirely to the new, separate application.

## Integration: Slack

- **Receiving data from the existing Slack app.** The "Emtech Scheduling"
  Slack app would `POST` a webhook to this dashboard whenever it creates a
  calendar event, so the dashboard has a live feed of scheduling activity
  without needing to poll anything. Needs from the dashboard side:
  - A stable HTTPS endpoint URL
  - An authentication scheme for incoming requests (bearer token, HMAC
    signature, etc. — not yet decided)
  - An agreed payload shape (currently assumed to be similar to what's already
    submitted per request: submitter, request type, dates, category,
    attendees, plus the Graph `event_id`/`web_link` — not yet finalized)
- **Reading/writing Slack Lists directly?** Open question — does the
  dashboard read the "Emtech Company Roster" and scheduling request Lists
  directly via the Slack API (`lists:read`/`lists:write` scopes) to display
  and manage requests, or does it maintain its own database as the real
  source of truth and treat the Slack List only as the original intake form?
  This is a real architectural decision, not yet made.
- **Slack OAuth, if phase 2 (status sync) happens.** The dashboard would need
  to host:
  - A link/button that sends each employee through Slack's standard OAuth
    consent screen (requesting the `users.profile:write` user scope)
  - A callback route that exchanges the resulting code for a per-user Slack
    access token
  - Secure, persistent storage for those tokens, one per opted-in employee
  - The Slack app's OAuth client secret, needed to perform that code exchange

## Integration: Microsoft (Graph / the shared calendar)

- **Read access to the "Team Availability" shared mailbox's calendar.** The
  existing Slack app already has Graph access to this mailbox (scoped via an
  Exchange Application Access Policy, not a tenant-wide permission), and that
  access already includes read, not just write. Open question for IT: reuse
  the existing app registration/credentials, or issue a separate, read-only
  Azure AD app registration for the dashboard specifically (better
  separation/least-privilege, at the cost of one more thing for IT to set up
  and rotate)?
- **How the dashboard learns about new/changed events** — two options:
  - Poll the calendar on a schedule (simpler, has inherent lag)
  - Subscribe to Microsoft Graph change notifications (webhooks) on the
    mailbox's calendar (near real-time, but subscriptions expire — roughly
    every 3 days for calendar resources — and need active renewal)
- **Credential handling**: same discipline as the existing Slack app —
  Tenant ID/Client ID/Client Secret stored as environment/secrets config,
  never in source control or chat, secret expiration tracked.

## Integration: AWS (hosting)

Not yet decided in any detail — flagging what any hosting choice will need
to account for, rather than proposing a specific architecture:

- **A public HTTPS endpoint** for the Slack webhook receiver and (if phase 2
  happens) the Slack OAuth callback
- **Persistent storage** for scheduling data (if the dashboard becomes its
  own source of truth rather than just a Slack List viewer) and, if phase 2
  happens, per-user Slack tokens
- **Secrets management** for Microsoft/Slack credentials (e.g. AWS Secrets
  Manager or SSM Parameter Store, rather than plain environment variables on
  a box)
- **The dashboard's own access control** — who can log in and view/manage
  this data at all? Not yet decided: could be Slack-based sign-in (Slack
  supports "Sign in with Slack," Slack's own OpenID Connect flow, which
  would be a light lift given the Slack integration already exists),
  company SSO, or something simpler — this needs an answer before much else
  here is built
- **Logging/monitoring** for a system that will be handling scheduling data
  and, eventually, calendar/status data for real employees

## A note on data sensitivity

Some of the request types already in use (`Sick`, in particular) are
HR-adjacent in a way plain calendar events on a shared mailbox may not have
fully surfaced yet. Worth looping in whoever owns HR/compliance concerns
early on questions like who can see what in the dashboard, and how long
scheduling history is retained — before access control gets designed, not
after.

## Explicitly not decided yet

- Dashboard tech stack / framework
- Whether this is one project or two (core dashboard vs. status-sync phase 2)
- Whether the dashboard becomes the source of truth for scheduling data, or
  stays a read/management view over Slack Lists
- Timeline / whether this is being greenlit at all

## Related context

- `README.md` / `CLAUDE.md` in this repo — the existing Slack app this would
  integrate with
- `IT_REQUEST.md` — the original Azure AD ask for the existing app, useful as
  a template for whatever the dashboard's own Microsoft/IT ask ends up being
