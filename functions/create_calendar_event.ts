import { DefineFunction, Schema, SlackFunction } from "deno-slack-sdk/mod.ts";

/**
 * Creates a calendar event on the shared scheduling mailbox via Microsoft
 * Graph, invited as a step from a Workflow Builder workflow.
 * https://learn.microsoft.com/en-us/graph/api/user-post-events
 */
export const CreateCalendarEventDefinition = DefineFunction({
  callback_id: "create_calendar_event",
  title: "Create scheduling calendar event",
  description:
    "Creates a calendar event on the shared scheduling mailbox via Microsoft Graph",
  source_file: "functions/create_calendar_event.ts",
  input_parameters: {
    properties: {
      submitter_alias: {
        type: Schema.types.string,
        description:
          "Alias/name of the request submitter, used in the event title",
      },
      submitter_email: {
        type: Schema.types.string,
        description: "Submitter's email, invited as the main attendee",
      },
      request_type: {
        type: Schema.types.string,
        description:
          "Type of request (from the Request Type dropdown), used in the event title and mapped internally to an Outlook category",
      },
      start_date_time: {
        type: Schema.types.string,
        description:
          "First day of the request, ISO 8601 (e.g. 2026-07-10T09:00:00). Always creates a full-day event — the time-of-day portion is ignored.",
      },
      end_date_time: {
        type: Schema.types.string,
        description:
          "Last day of the request, inclusive (e.g. a single-day request has the same date as start_date_time). ISO 8601; time-of-day is ignored.",
      },
      description: {
        type: Schema.types.string,
        description: "Event description, pulled from the request form",
      },
      location: {
        type: Schema.types.string,
        description: "Event location",
      },
      additional_attendees: {
        type: Schema.types.array,
        items: { type: Schema.slack.types.user_id },
        description:
          "Additional internal attendees as Slack user references (e.g. from a List's Person column). Resolved to real email addresses internally via the Slack API.",
      },
      external_attendees: {
        type: Schema.types.string,
        description:
          "Additional external attendees (e.g. non-org POCs/clients with no Slack account) as a single string of email addresses separated by whitespace, commas, or semicolons — e.g. from a List's plain Text column.",
      },
    },
    required: [
      "submitter_alias",
      "submitter_email",
      "request_type",
      "start_date_time",
      "end_date_time",
    ],
  },
  output_parameters: {
    properties: {
      event_id: {
        type: Schema.types.string,
        description: "The created Microsoft Graph event ID",
      },
      web_link: {
        type: Schema.types.string,
        description: "Outlook web link to the created event",
      },
    },
    required: ["event_id"],
  },
});

async function getGraphAccessToken(
  env: Record<string, string | undefined>,
): Promise<string> {
  const tenantId = env["MS_TENANT_ID"];
  const clientId = env["MS_CLIENT_ID"];
  const clientSecret = env["MS_CLIENT_SECRET"];
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      "Missing MS_TENANT_ID, MS_CLIENT_ID, or MS_CLIENT_SECRET environment variable",
    );
  }

  const response = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
      }),
    },
  );

  if (!response.ok) {
    throw new Error(
      `Failed to acquire Graph access token: ${response.status} ${await response
        .text()}`,
    );
  }

  const { access_token } = await response.json();
  return access_token as string;
}

type DateParts = { year: number; month: number; day: number };

const SHORT_MONTH_NAMES = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

// Extracts the literal "YYYY-MM-DD" prefix of an ISO date/datetime string.
//
// start_date_time / end_date_time represent a calendar date someone picked
// (via a date picker) — there's no real time-of-day or timezone meaning
// intended, even though the raw value may carry a time/offset suffix (e.g.
// "T00:00:00Z") as an artifact of the Slack List field's storage format.
// A real bug already happened from getting this wrong: treating that
// midnight-UTC suffix as a genuine moment in time and converting it to
// America/Los_Angeles (hours behind UTC) rolled every date back by one full
// day. The fix is to never timezone-convert these values at all — the
// digits the user picked are the answer, regardless of what suffix is
// attached.
function getCalendarDateParts(isoDateTime: string): DateParts {
  const match = isoDateTime.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) {
    throw new Error(`Expected an ISO date string, got "${isoDateTime}"`);
  }
  const [, year, month, day] = match;
  return { year: Number(year), month: Number(month), day: Number(day) };
}

// Title format: "<submitter alias> - <request type> - <start date>"
export function formatTitleDate(isoDateTime: string): string {
  const { year, month, day } = getCalendarDateParts(isoDateTime);
  return `${SHORT_MONTH_NAMES[month - 1]} ${day}, ${year}`;
}

// Adds `days` calendar days to a Y/M/D triple. Uses UTC as a neutral, DST-free
// computation frame — this is pure calendar arithmetic on already-extracted
// date parts, not a timezone conversion.
function addCalendarDays(parts: DateParts, days: number): DateParts {
  const utc = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
  utc.setUTCDate(utc.getUTCDate() + days);
  return {
    year: utc.getUTCFullYear(),
    month: utc.getUTCMonth() + 1,
    day: utc.getUTCDate(),
  };
}

function toMidnightIso({ year, month, day }: DateParts): string {
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${year}-${pad(month)}-${pad(day)}T00:00:00`;
}

// Computes the start/end dateTime strings for an all-day Graph event. Graph
// requires all-day start/end to both be midnight (paired with a timeZone by
// the caller), with an *exclusive* end — i.e., the day after the last day
// of the event, not the last day itself. endDateTime here is the workflow's
// "end date," which is the inclusive last day someone is out, so this adds
// one calendar day. This is pure calendar-date arithmetic — no timezone
// conversion — see getCalendarDateParts for why.
// https://learn.microsoft.com/en-us/graph/api/resources/event#properties
export function getAllDayEventRange(
  startDateTime: string,
  endDateTime: string,
): { start: string; end: string } {
  const startParts = getCalendarDateParts(startDateTime);
  const endParts = addCalendarDays(getCalendarDateParts(endDateTime), 1);
  return {
    start: toMidnightIso(startParts),
    end: toMidnightIso(endParts),
  };
}

// Maps the workflow's "Request Type" dropdown to the Outlook category
// pre-configured on the shared mailbox by IT. Keep in sync with both the
// dropdown's options in Workflow Builder and README's "Outlook categories"
// section whenever either changes.
const REQUEST_TYPE_TO_CATEGORY: Record<string, string> = {
  "Sick": "OOTO",
  "OOTO": "OOTO",
  "4x10 OOTO": "4x10 OOTO",
  "WFH": "WFH",
  "Vacation": "OOTO",
  "Location Assignment": "On-Site",
};

export function getCategoryForRequestType(
  requestType: string,
): string | undefined {
  return REQUEST_TYPE_TO_CATEGORY[requestType];
}

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

// Splits external_attendees (a single string) into individual email
// addresses. Slack's List "Email" column type is documented as multi-value,
// but saving more than one value to it fails in practice ("Failed to save
// changes!") — this works around that by using a plain Text column instead,
// with addresses separated by whitespace, commas, or semicolons (any mix),
// matching however someone naturally types a list of emails. Returns an
// error naming the first entry that doesn't look like a valid email, rather
// than silently sending something malformed to Graph.
//
// Strips a leading "mailto:" from each entry — Slack's List Text column
// auto-linkifies email-looking text, and whatever serialization turns that
// rich-text value into a plain string variable carries the "mailto:" scheme
// along with it. A real bug happened from missing this: Graph/Exchange
// silently accepted "mailto:name@x.com" as an attendee address (it still
// matches EMAIL_PATTERN — the pattern doesn't exclude colons) but never
// actually delivered an invite to the real underlying address.
export function parseEmailList(
  value: string | undefined,
): { emails: string[] } | { error: string } {
  const entries = (value ?? "")
    .split(/[\s,;]+/)
    .map((entry) => entry.trim().replace(/^mailto:/i, ""))
    .filter((entry) => entry.length > 0);

  for (const entry of entries) {
    if (!EMAIL_PATTERN.test(entry)) {
      return { error: `"${entry}" doesn't look like a valid email address` };
    }
  }

  return { emails: entries };
}

export default SlackFunction(
  CreateCalendarEventDefinition,
  async ({ inputs, env, client }) => {
    const {
      submitter_alias,
      submitter_email,
      request_type,
      start_date_time,
      end_date_time,
      description,
      location,
      additional_attendees,
      external_attendees,
    } = inputs;

    const mailbox = env["MS_SHARED_MAILBOX"];
    if (!mailbox) {
      return { error: "Missing MS_SHARED_MAILBOX environment variable" };
    }

    const category = getCategoryForRequestType(request_type);
    if (!category) {
      return {
        error:
          `No Outlook category mapping configured for request type "${request_type}"`,
      };
    }

    const externalAttendeesResult = parseEmailList(external_attendees);
    if ("error" in externalAttendeesResult) {
      return {
        error: `Invalid external_attendees: ${externalAttendeesResult.error}`,
      };
    }

    const timeZone = env["MS_EVENT_TIMEZONE"] || "UTC";

    let accessToken: string;
    try {
      accessToken = await getGraphAccessToken(env);
    } catch (err) {
      return { error: `${err}` };
    }

    const title = `${submitter_alias} - ${request_type} - ${
      formatTitleDate(start_date_time)
    }`;

    // additional_attendees are Slack user references (e.g. a List's Person
    // column) — Microsoft Graph has no concept of Slack users, so each one
    // needs to be resolved to a real email address first. Resolution is
    // best-effort: an attendee who can't be resolved (hidden email, guest
    // account, etc.) is skipped rather than failing the whole request, and
    // the submitter gets a DM afterward listing who was skipped so they can
    // follow up.
    const resolvedAttendeeEmails: string[] = [];
    const unresolvedAttendeeIds: string[] = [];
    for (const userId of additional_attendees ?? []) {
      const userInfo = await client.users.info({ user: userId });
      const email = userInfo.ok ? userInfo.user?.profile?.email : undefined;
      if (email) {
        resolvedAttendeeEmails.push(email);
      } else {
        unresolvedAttendeeIds.push(userId);
      }
    }

    const attendees = [
      { emailAddress: { address: submitter_email }, type: "required" },
      // External attendees are already plain email addresses — no Slack
      // account to resolve, so they're used as-is alongside the resolved
      // internal ones.
      ...[...resolvedAttendeeEmails, ...externalAttendeesResult.emails].map(
        (email) => ({
          emailAddress: { address: email },
          type: "required",
        }),
      ),
    ];

    const allDayRange = getAllDayEventRange(start_date_time, end_date_time);

    const eventBody = {
      subject: title,
      body: { contentType: "Text", content: description ?? "" },
      start: { dateTime: allDayRange.start, timeZone },
      end: { dateTime: allDayRange.end, timeZone },
      isAllDay: true,
      location: { displayName: location ?? "" },
      attendees,
      // Keeps the shared mailbox's own calendar entry shown as free; does not
      // change how attendees' own calendars display the event once accepted.
      showAs: "free",
      categories: [category],
    };

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/users/${
        encodeURIComponent(mailbox)
      }/events`,
      {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(eventBody),
      },
    );

    if (!response.ok) {
      return {
        error:
          `Failed to create calendar event: ${response.status} ${await response
            .text()}`,
      };
    }

    const event = await response.json();

    if (unresolvedAttendeeIds.length > 0) {
      // Best-effort: the event was already created successfully above, so a
      // failure here should never turn this into a failed step.
      try {
        const submitter = await client.users.lookupByEmail({
          email: submitter_email,
        });
        if (submitter.ok && submitter.user?.id) {
          const mentions = unresolvedAttendeeIds
            .map((id) => `<@${id}>`)
            .join(", ");
          await client.chat.postMessage({
            channel: submitter.user.id,
            text:
              `Your "${title}" calendar event was created, but ${mentions} couldn't be added as an attendee — their email address wasn't available. They may need to check their Slack profile's contact info.`,
          });
        }
      } catch {
        // Notification is a courtesy, not a requirement — swallow any error.
      }
    }

    return {
      outputs: {
        event_id: event.id,
        web_link: event.webLink ?? "",
      },
    };
  },
);
