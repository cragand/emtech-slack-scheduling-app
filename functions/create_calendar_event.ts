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
        items: { type: Schema.types.string },
        description: "Additional attendee emails, pulled from the Slack List",
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

// Title format: "<submitter alias> - <request type> - <start date>"
// timeZone must match the event's own timeZone (MS_EVENT_TIMEZONE) so the
// title's date can never disagree with which day the event actually lands
// on — without an explicit timeZone here, this would silently fall back to
// whatever zone the process happens to be running in, which can differ from
// MS_EVENT_TIMEZONE depending on where the app is running.
export function formatTitleDate(isoDateTime: string, timeZone: string): string {
  const date = new Date(isoDateTime);
  return date.toLocaleDateString("en-US", {
    year: "numeric",
    month: "short",
    day: "numeric",
    timeZone,
  });
}

// Calendar-date parts (year/month/day) as observed in the given IANA
// timeZone — used instead of raw string-slicing since the offset in
// isoDateTime (if any) can put it on a different calendar day than what's
// intended once viewed in timeZone.
function getCalendarDateParts(
  isoDateTime: string,
  timeZone: string,
): { year: number; month: number; day: number } {
  const date = new Date(isoDateTime);
  // en-CA formats as YYYY-MM-DD, convenient to split.
  const formatted = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date);
  const [year, month, day] = formatted.split("-").map(Number);
  return { year, month, day };
}

type DateParts = { year: number; month: number; day: number };

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
// requires all-day start/end to both be midnight in the same timeZone, with
// an *exclusive* end — i.e., the day after the last day of the event, not
// the last day itself. endDateTime here is the workflow's "end date," which
// is the inclusive last day someone is out, so this adds one calendar day.
// https://learn.microsoft.com/en-us/graph/api/resources/event#properties
export function getAllDayEventRange(
  startDateTime: string,
  endDateTime: string,
  timeZone: string,
): { start: string; end: string } {
  const startParts = getCalendarDateParts(startDateTime, timeZone);
  const endParts = addCalendarDays(
    getCalendarDateParts(endDateTime, timeZone),
    1,
  );
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

export default SlackFunction(
  CreateCalendarEventDefinition,
  async ({ inputs, env }) => {
    const {
      submitter_alias,
      submitter_email,
      request_type,
      start_date_time,
      end_date_time,
      description,
      location,
      additional_attendees,
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

    const timeZone = env["MS_EVENT_TIMEZONE"] || "UTC";

    let accessToken: string;
    try {
      accessToken = await getGraphAccessToken(env);
    } catch (err) {
      return { error: `${err}` };
    }

    const title = `${submitter_alias} - ${request_type} - ${
      formatTitleDate(start_date_time, timeZone)
    }`;

    const attendees = [
      { emailAddress: { address: submitter_email }, type: "required" },
      ...(additional_attendees ?? []).map((email) => ({
        emailAddress: { address: email },
        type: "required",
      })),
    ];

    const allDayRange = getAllDayEventRange(
      start_date_time,
      end_date_time,
      timeZone,
    );

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

    return {
      outputs: {
        event_id: event.id,
        web_link: event.webLink ?? "",
      },
    };
  },
);
