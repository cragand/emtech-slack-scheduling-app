import { DefineFunction, Schema, SlackFunction } from "deno-slack-sdk/mod.ts";
import { getGraphAccessToken } from "./lib/graph_auth.ts";

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
      amazon_alias: {
        type: Schema.types.string,
        description:
          "Submitter's Amazon Alias, from the company roster lookup (Workflow Builder step 3). First segment of the event title.",
      },
      submitter_name: {
        type: Schema.types.string,
        description:
          "Name of the request submitter. Last segment of the event title.",
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
          "First day of the request, ISO 8601 (e.g. 2026-07-10T09:00:00). Creates a full-day event by default — the time-of-day portion is ignored unless is_partial_day is true, in which case it's used as the event's real start time.",
      },
      end_date_time: {
        type: Schema.types.string,
        description:
          "Last day of the request, inclusive (e.g. a single-day request has the same date as start_date_time). ISO 8601; time-of-day is ignored unless is_partial_day is true, in which case it's used as the event's real end time.",
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
      is_partial_day: {
        type: Schema.types.string,
        description:
          "Whether this request covers only part of a day rather than the full day(s) — a string rather than a boolean because boolean-typed inputs can't be mapped to a per-submission variable in Workflow Builder's step configuration (they only render as a fixed checkbox). Sourced from a 'Yes'/'No' dropdown; matched case-insensitively against \"yes\" via isPartialDayRequest(). When true, the time-of-day portions of start_date_time/end_date_time are used to create a timed (non-all-day) event instead of being ignored. Optional — omitted entirely by workflows that don't support partial days (Sick, Vacation, 4x10 OOTO), which keeps them on the existing all-day behavior untouched.",
      },
      is_recurring: {
        type: Schema.types.string,
        description:
          "Whether this is a recurring weekly request (e.g. a 4x10 schedule's consistent day off) — a string 'Yes'/'No' for the same reason is_partial_day is, matched case-insensitively via isRecurringRequest(). When true, recurrence_end_date and recurrence_day_of_week are required, and the created event repeats weekly through that end date. Optional — omitted by workflows that don't support recurrence (Sick, Vacation, OOTO), which keeps them on today's single-event behavior untouched. Mutually exclusive with is_partial_day.",
      },
      recurrence_end_date: {
        type: Schema.types.string,
        description:
          "The last date the recurring series should generate occurrences through, inclusive. ISO 8601 (e.g. from a plain Date field) — only used when is_recurring is true.",
      },
      recurrence_day_of_week: {
        type: Schema.types.string,
        description:
          "Which day of the week the recurring event falls on (e.g. from a single-select day-of-week dropdown already in the form). Matched case-insensitively against the seven day names via normalizeDayOfWeek() — only used when is_recurring is true. Doesn't need to match the weekday of start_date_time; Graph starts generating from start_date_time forward and finds the first matching weekday on or after it.",
      },
    },
    required: [
      "amazon_alias",
      "submitter_name",
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

type TimeParts = { hour: number; minute: number; second: number };

const ISO_DATE_PATTERN = /^(\d{4})-(\d{2})-(\d{2})/;
const UNIX_TIMESTAMP_PATTERN = /^\d+$/;

// Slack sends date/time values in two genuinely different formats depending
// on which List/form field type they came from, and each needs the opposite
// handling:
//
// - A plain "Date" field (Sick, Vacation, 4x10 OOTO, and OOTO's own all-day
//   case) sends an ISO string like "2026-07-08T00:00:00Z" — but that's a
//   calendar date someone picked, not a real moment in time, even though it's
//   dressed up as midnight UTC. A real bug already shipped from getting this
//   wrong: converting that midnight-UTC suffix to America/Los_Angeles (hours
//   behind UTC) rolled every date back by one full day. The fix there is to
//   never timezone-convert it at all — take the literal digits as-is.
// - A "Date and time" field (used for OOTO's partial-day support) sends a
//   bare Unix timestamp instead, e.g. "1784235600" — genuinely a real moment
//   in time this time, confirmed against a live test (a 2:00 PM Pacific
//   submission came through as exactly 21:00 UTC). This one needs the
//   opposite handling: it must be converted into timeZone to recover the
//   wall-clock date/time that was actually picked.
function resolveDateTimeParts(
  value: string,
  timeZone: string,
): { date: DateParts; time: TimeParts } {
  const isoMatch = value.match(ISO_DATE_PATTERN);
  if (isoMatch) {
    const [, year, month, day] = isoMatch;
    const timeMatch = value.match(/T(\d{2}):(\d{2}):(\d{2})/);
    const [, hour, minute, second] = timeMatch ?? [, "0", "0", "0"];
    return {
      date: { year: Number(year), month: Number(month), day: Number(day) },
      time: {
        hour: Number(hour),
        minute: Number(minute),
        second: Number(second),
      },
    };
  }

  if (UNIX_TIMESTAMP_PATTERN.test(value)) {
    const instant = new Date(Number(value) * 1000);
    const formatter = new Intl.DateTimeFormat("en-US", {
      timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23",
    });
    const parts = formatter.formatToParts(instant);
    const part = (type: string) =>
      Number(parts.find((p) => p.type === type)?.value);
    return {
      date: { year: part("year"), month: part("month"), day: part("day") },
      time: {
        hour: part("hour") % 24,
        minute: part("minute"),
        second: part("second"),
      },
    };
  }

  throw new Error(
    `Expected an ISO date string or Unix timestamp, got "${value}"`,
  );
}

function getCalendarDateParts(value: string, timeZone: string): DateParts {
  return resolveDateTimeParts(value, timeZone).date;
}

function getTimeOfDayParts(value: string, timeZone: string): TimeParts {
  return resolveDateTimeParts(value, timeZone).time;
}

// Title format: "@<amazon alias> - <request type> - <start date> - <submitter name>"
// The year is deliberately omitted — it's always obvious from the calendar
// page the event appears on, and including it just added noise.
export function formatTitleDate(isoDateTime: string, timeZone: string): string {
  const { month, day } = getCalendarDateParts(isoDateTime, timeZone);
  return `${SHORT_MONTH_NAMES[month - 1]} ${day}`;
}

// Strips a leading "@" (and any whitespace right after it) from a name.
// submitter_name is sometimes sourced from a Slack Person-type variable,
// which renders as "@Display Name" when interpolated into a plain string —
// the "@" belongs on amazon_alias in the title, not here, regardless of
// whether the incoming value happens to carry one already.
export function stripLeadingAt(value: string): string {
  return value.replace(/^@\s*/, "");
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

function toLocalIso(parts: DateParts, time: TimeParts): string {
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${parts.year}-${pad(parts.month)}-${pad(parts.day)}T${
    pad(time.hour)
  }:${pad(time.minute)}:${pad(time.second)}`;
}

// Computes the start/end dateTime strings for a partial-day (timed, non-
// all-day) Graph event, using the literal date *and* time-of-day from
// startDateTime/endDateTime directly — unlike the all-day path, the time
// portion is meaningful here, not discarded.
export function getPartialDayEventRange(
  startDateTime: string,
  endDateTime: string,
  timeZone: string,
): { start: string; end: string } {
  return {
    start: toLocalIso(
      getCalendarDateParts(startDateTime, timeZone),
      getTimeOfDayParts(startDateTime, timeZone),
    ),
    end: toLocalIso(
      getCalendarDateParts(endDateTime, timeZone),
      getTimeOfDayParts(endDateTime, timeZone),
    ),
  };
}

// Matches case-insensitively (and trims whitespace) against "yes", same
// approach as getCategoryForRequestType — the dropdown's exact wording/case
// shouldn't matter.
export function isPartialDayRequest(value: string | undefined): boolean {
  return (value ?? "").trim().toLowerCase() === "yes";
}

export function isRecurringRequest(value: string | undefined): boolean {
  return (value ?? "").trim().toLowerCase() === "yes";
}

const VALID_DAYS_OF_WEEK = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
];

// Matches case-insensitively (and trims whitespace) against the seven day
// names, same reasoning as getCategoryForRequestType and
// isPartialDayRequest — the dropdown's exact wording/case shouldn't matter.
// Returns the canonically-cased name Graph expects, or undefined if the
// value isn't a recognized day.
export function normalizeDayOfWeek(
  value: string | undefined,
): string | undefined {
  const normalized = (value ?? "").trim().toLowerCase();
  return VALID_DAYS_OF_WEEK.find((day) => day.toLowerCase() === normalized);
}

// Builds a Graph weekly recurrence: repeats every week on dayOfWeek,
// starting from startDateTime's date and continuing through
// recurrenceEndDate, inclusive. startDateTime doesn't need to fall on
// dayOfWeek itself — Graph starts generating occurrences from that date
// forward and finds the first matching weekday on or after it.
// https://learn.microsoft.com/en-us/graph/api/resources/recurrencepattern
export function buildWeeklyRecurrence(
  startDateTime: string,
  recurrenceEndDate: string,
  dayOfWeek: string,
  timeZone: string,
) {
  const pad = (n: number) => String(n).padStart(2, "0");
  const toDateOnly = ({ year, month, day }: DateParts) =>
    `${year}-${pad(month)}-${pad(day)}`;

  return {
    pattern: { type: "weekly", interval: 1, daysOfWeek: [dayOfWeek] },
    range: {
      type: "endDate",
      startDate: toDateOnly(getCalendarDateParts(startDateTime, timeZone)),
      endDate: toDateOnly(getCalendarDateParts(recurrenceEndDate, timeZone)),
    },
  };
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

// Matches case-insensitively (and trims whitespace) against the dropdown's
// options — a real bug happened from an exact-case lookup: the dropdown's
// actual value ("4X10 OOTO") didn't match this table's key ("4x10 OOTO"),
// so a legitimate request type was rejected as unrecognized.
export function getCategoryForRequestType(
  requestType: string,
): string | undefined {
  const normalized = requestType.trim().toLowerCase();
  for (const [key, category] of Object.entries(REQUEST_TYPE_TO_CATEGORY)) {
    if (key.toLowerCase() === normalized) {
      return category;
    }
  }
  return undefined;
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
      amazon_alias,
      submitter_name,
      submitter_email,
      request_type,
      start_date_time,
      end_date_time,
      description,
      location,
      additional_attendees,
      external_attendees,
      is_partial_day,
      is_recurring,
      recurrence_end_date,
      recurrence_day_of_week,
    } = inputs;

    const mailbox = env["MS_SHARED_MAILBOX"];
    if (!mailbox) {
      return { error: "Missing MS_SHARED_MAILBOX environment variable" };
    }

    const partialDay = isPartialDayRequest(is_partial_day);
    const recurring = isRecurringRequest(is_recurring);

    if (partialDay && recurring) {
      return {
        error:
          "is_partial_day and is_recurring can't both be true — a request is never both a partial day and a recurring day off",
      };
    }

    let recurrenceDayOfWeek: string | undefined;
    let recurrenceEndDate: string | undefined;
    if (recurring) {
      if (!recurrence_end_date || !recurrence_day_of_week) {
        return {
          error:
            "is_recurring is true, but recurrence_end_date and/or recurrence_day_of_week weren't provided",
        };
      }
      recurrenceDayOfWeek = normalizeDayOfWeek(recurrence_day_of_week);
      if (!recurrenceDayOfWeek) {
        return {
          error:
            `"${recurrence_day_of_week}" isn't a recognized day of the week`,
        };
      }
      recurrenceEndDate = recurrence_end_date;
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

    const title = `@${stripLeadingAt(amazon_alias)} - ${request_type} - ${
      formatTitleDate(start_date_time, timeZone)
    } - ${stripLeadingAt(submitter_name)}`;

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

    const range = partialDay
      ? getPartialDayEventRange(start_date_time, end_date_time, timeZone)
      : getAllDayEventRange(start_date_time, end_date_time, timeZone);

    const eventBody = {
      subject: title,
      body: { contentType: "Text", content: description ?? "" },
      start: { dateTime: range.start, timeZone },
      end: { dateTime: range.end, timeZone },
      isAllDay: !partialDay,
      location: { displayName: location ?? "" },
      attendees,
      // Keeps the shared mailbox's own calendar entry shown as free; does not
      // change how attendees' own calendars display the event once accepted.
      showAs: "free",
      categories: [category],
      ...(recurring && recurrenceDayOfWeek && recurrenceEndDate
        ? {
          recurrence: buildWeeklyRecurrence(
            start_date_time,
            recurrenceEndDate,
            recurrenceDayOfWeek,
            timeZone,
          ),
        }
        : {}),
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
