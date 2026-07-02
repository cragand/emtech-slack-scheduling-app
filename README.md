# Emtech Scheduling

A Slack app (Deno/next-gen platform) that adds a custom Workflow Builder step,
**"Create scheduling calendar event,"** which creates a calendar event on a
shared Microsoft 365 mailbox via the Microsoft Graph API.

## How it fits together

The actual workflow logic — the trigger, the form, the Slack List — lives in
**Slack Workflow Builder** (no-code), not in this repo. This repo provides one
building block that workflow uses as a step: a function that, given a
submitter, request details, and a list of attendees, calls Microsoft Graph to
put an event on the shared scheduling mailbox's calendar.

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
- [ ] **Waiting on IT**: Azure AD app registration (Tenant ID, Client ID,
      Client Secret) — see [Getting credentials from IT](#getting-credentials-from-it)
- [ ] End-to-end test once credentials land

Until credentials are supplied, running the step will fail fast with a clear
error (`Missing MS_TENANT_ID, MS_CLIENT_ID, or MS_CLIENT_SECRET environment
variable`) rather than doing anything destructive — this is expected, not a bug.

## Prerequisites

- [Slack CLI](https://api.slack.com/automation/quickstart), already
  authenticated to the `emtechllc` workspace
- [Deno](https://deno.com) (installed via `winget install DenoLand.Deno` on
  this machine)

> **Windows/Git Bash note:** if `slack` commands hang with no output at all,
> check `deno --version` in that same shell first — the Slack CLI shells out to
> Deno for most commands and fails silently (or hangs) if it's not on `PATH`.
> Also, some `slack` commands (`doctor`, `run` without `--team`) don't play
> well with Git Bash's pty and will hang waiting on a prompt that can't be
> answered; prefer running them from a native PowerShell/cmd terminal.

## Running locally

```sh
cd emtech-scheduling-slack-app
slack run --team emtechllc --hide-triggers
```

You'll know it's running when you see `Connected, awaiting events`. While it's
running, a local version of the app (`Emtech Scheduling (local)`) is installed
in the workspace and its step is selectable in Workflow Builder.

`--team emtechllc` skips an interactive workspace-selection prompt that fails
in non-interactive shells; `--hide-triggers` skips prompts related to the
unused sample trigger (see [Leftover scaffold files](#leftover-scaffold-files)).

## Environment variables

Copy `.env.example` to `.env` and fill in real values once IT provides them:

| Variable | Purpose |
|---|---|
| `MS_TENANT_ID` | Identifies Emtech's Azure AD directory |
| `MS_CLIENT_ID` | Identifies this app's registration within that directory |
| `MS_CLIENT_SECRET` | The app's credential, used to authenticate as itself (no user login involved) |
| `MS_SHARED_MAILBOX` | The shared scheduling mailbox address the event gets created on (e.g. `emtech-scheduling@yourdomain.com`) |
| `MS_EVENT_TIMEZONE` | IANA timezone used for the event's start/end (e.g. `America/New_York`) |

`slack run` automatically picks up `.env` locally. For the deployed app, set
these with `slack env add <NAME> <VALUE>` instead — `.env` is git-ignored and
never used once deployed.

### Getting credentials from IT

This app authenticates to Microsoft Graph via the OAuth2 **client credentials**
flow — it runs unattended (triggered by a Slack List update), so it has to
authenticate as itself rather than as a signed-in person. That requires:

1. An **Azure AD (Entra ID) app registration** for this app, which produces the
   Tenant ID, Client ID, and Client Secret above.
2. The `Calendars.ReadWrite` **application permission** on that registration,
   with **admin consent** granted (application permissions always require a
   tenant admin's sign-off, since no individual user consents on the app's
   behalf).
3. The shared mailbox (`emtech-scheduling@...` or equivalent) assigned a
   Microsoft 365 license, needed for full calendar support via Graph.

Worth flagging to IT proactively: by default, `Calendars.ReadWrite` as an
*application* permission grants access to **every mailbox in the tenant**, not
just the shared one. An Exchange Online **Application Access Policy** can scope
this app down to only the shared scheduling mailbox — likely to speed up
approval from a security-conscious admin.

## Operational considerations for live/autonomous use

- **Client secret expiration**: Azure AD client secrets expire on whatever
  schedule IT sets (commonly 6–24 months). Once deployed and running
  unattended, an expired secret means the app fails silently until someone
  rotates `MS_CLIENT_SECRET` via `slack env add`. Worth asking IT for the
  expiration date up front and calendaring a reminder, or asking whether a
  longer-lived credential type is available.
- **Deployed env vars are separate from `.env`**: see
  [Environment variables](#environment-variables) above — `slack env add` has
  to be run once against the deployed app; local `.env` values don't carry
  over.
- `slack run` is local-dev-only and requires this machine/terminal to stay on;
  it is not how the app runs for real usage. See [Deploying](#deploying).

## Known limitation: attendees' free/busy status

The event is created with `showAs: "free"`, but this only controls the
free/busy status on the **shared mailbox's own calendar copy**. Each invited
attendee gets their own separate copy of the event once they accept, and
Outlook defaults that copy to **Busy** regardless of the organizer's `showAs`
setting — there's no Graph API call that makes it "Free" on every attendee's
calendar at once. This was a deliberate tradeoff (real Outlook invites were
chosen over a mailbox-only entry) — see `functions/create_calendar_event.ts`
for where this is implemented.

## Testing

```sh
deno test
```

`functions/sample_function_test.ts` is the only test file present, from the
default scaffold (see below). No test currently exists for
`create_calendar_event.ts`; the Graph API calls would need to be mocked since
there's no live credential to test against yet.

## Deploying

Once verified locally, deploy to Slack-hosted infrastructure:

```sh
slack deploy
```

Set production environment variables first with `slack env add`, since `.env`
is only used for local `slack run`.

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
