# Azure AD app registration — what we need

Hey! We're building a small Slack automation (Emtech Scheduling) that needs to create calendar events on a shared mailbox. It runs unattended, so it needs its own app identity in Azure AD rather than logging in as a person. Here's what we need from you:

## 1. App registration

A new Azure AD (Entra ID) app registration for this. Doesn't need to be anything fancy — just needs:

- **`Calendars.ReadWrite`** as an **application permission** (not delegated), with admin consent granted
- Once that's set up, send us the **Tenant ID**, **Client ID**, and a **Client Secret**

## 2. Shared mailbox

The shared mailbox we'll be creating events on (e.g. `emtech-scheduling@emtech.us`, or whatever you'd like to use) needs a Microsoft 365 license attached, so it has full calendar support via the Graph API.

## 3. One ask to make your life easier (optional but recommended)

By default, `Calendars.ReadWrite` as an app permission gives this app access to **every mailbox in the tenant**, not just the shared one. If you'd rather scope it down, an Exchange Online **Application Access Policy** can restrict it to just the shared scheduling mailbox — happy to help test that if you set one up.

## 4. One thing to plan for later

Client secrets expire (whatever timeline you set — 6/12/24 months, up to you). When you create it, could you let us know the expiration date? We'll put a reminder on the calendar so this doesn't quietly break down the road.

## Sending the credentials back

Whatever's easiest/most secure on your end works for us — just avoid pasting the Client Secret somewhere that sticks around in plaintext (email, Slack DM history, etc.) if you can help it.

Thanks for the help — happy to hop on a call if anything above is unclear!
