import { DefineFunction, Schema, SlackFunction } from "deno-slack-sdk/mod.ts";
import { getGraphAccessToken } from "./lib/graph_auth.ts";

/**
 * Posts a daily digest of today's events on the shared scheduling mailbox
 * to a Slack channel, invoked from a Workflow Builder workflow on an
 * "On a schedule" trigger (not a form) — this function does everything
 * itself (reads Graph, posts to Slack), so the workflow around it is just
 * the schedule trigger plus this one step.
 */
export const PostDailyScheduleDigestDefinition = DefineFunction({
  callback_id: "post_daily_schedule_digest",
  title: "Post daily schedule digest",
  description:
    "Posts today's events from the shared scheduling mailbox to a Slack channel",
  source_file: "functions/post_daily_schedule_digest.ts",
  input_parameters: {
    properties: {
      channel_id: {
        type: Schema.slack.types.channel_id,
        description: "Channel to post the digest to",
      },
    },
    required: ["channel_id"],
  },
  output_parameters: {
    properties: {
      event_count: {
        type: Schema.types.number,
        description: "Number of events included in the digest",
      },
    },
    required: [],
  },
});

type DateParts = { year: number; month: number; day: number };

// Computes today's Y/M/D in timeZone, then the UTC instants for that day's
// midnight-to-midnight boundaries — the range Graph's calendarView needs to
// return everything happening "today" in the org's own timezone, not UTC's.
function getTodayRangeUtc(
  now: Date,
  timeZone: string,
): { startUtc: string; endUtc: string; todayLabel: DateParts } {
  const formatter = new Intl.DateTimeFormat("en-US", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
  const parts = formatter.formatToParts(now);
  const part = (type: string) =>
    Number(parts.find((p) => p.type === type)?.value);
  const today: DateParts = {
    year: part("year"),
    month: part("month"),
    day: part("day"),
  };

  // Midnight today and midnight tomorrow, expressed as literal local
  // wall-clock strings — paired with timeZone in the query itself (via the
  // Prefer header), Graph resolves these correctly rather than us needing
  // to compute a UTC offset by hand.
  const pad = (n: number) => String(n).padStart(2, "0");
  const startUtc = `${today.year}-${pad(today.month)}-${
    pad(today.day)
  }T00:00:00`;
  const tomorrow = new Date(Date.UTC(today.year, today.month - 1, today.day));
  tomorrow.setUTCDate(tomorrow.getUTCDate() + 1);
  const endUtc = `${tomorrow.getUTCFullYear()}-${
    pad(tomorrow.getUTCMonth() + 1)
  }-${pad(tomorrow.getUTCDate())}T00:00:00`;

  return { startUtc, endUtc, todayLabel: today };
}

// Formats one Graph event as a bulleted, clickable line — the subject is
// already the full title create_calendar_event.ts built at creation time
// (e.g. "@jdoe - Sick - Jul 20 - John Doe"), so this just needs to wrap it
// as a link using the event's own webLink, same property
// create_calendar_event.ts already captures when creating an event.
export function formatDigestLine(
  event: { subject?: string; webLink?: string },
): string {
  const subject = event.subject ?? "(untitled event)";
  return event.webLink ? `• <${event.webLink}|${subject}>` : `• ${subject}`;
}

export function buildDigestMessage(
  events: { subject?: string; webLink?: string }[],
): string {
  if (events.length === 0) {
    return "Nothing on the schedule today.";
  }
  return `Today's schedule:\n${events.map(formatDigestLine).join("\n")}`;
}

export default SlackFunction(
  PostDailyScheduleDigestDefinition,
  async ({ inputs, env, client }) => {
    const { channel_id } = inputs;

    const mailbox = env["MS_SHARED_MAILBOX"];
    if (!mailbox) {
      return { error: "Missing MS_SHARED_MAILBOX environment variable" };
    }
    const timeZone = env["MS_EVENT_TIMEZONE"] || "UTC";

    let accessToken: string;
    try {
      accessToken = await getGraphAccessToken(env);
    } catch (err) {
      return { error: `${err}` };
    }

    const { startUtc, endUtc } = getTodayRangeUtc(new Date(), timeZone);

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/users/${
        encodeURIComponent(mailbox)
      }/calendarView?startDateTime=${
        encodeURIComponent(startUtc)
      }&endDateTime=${encodeURIComponent(endUtc)}`,
      {
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Prefer": `outlook.timezone="${timeZone}"`,
        },
      },
    );

    if (!response.ok) {
      return {
        error:
          `Failed to read the shared mailbox's calendar: ${response.status} ${await response
            .text()}`,
      };
    }

    const { value: events } = await response.json();

    await client.chat.postMessage({
      channel: channel_id,
      text: buildDigestMessage(events),
    });

    return {
      outputs: { event_count: events.length },
    };
  },
);
