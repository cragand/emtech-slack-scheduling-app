import { SlackFunctionTester } from "deno-slack-sdk/mod.ts";
import { assertEquals, assertExists, assertStringIncludes } from "@std/assert";
import { stub } from "@std/testing/mock";
import CreateCalendarEvent, {
  buildWeeklyRecurrence,
  formatTitleDate,
  getAllDayEventRange,
  getCategoryForRequestType,
  getPartialDayEventRange,
  isPartialDayRequest,
  isRecurringRequest,
  normalizeDayOfWeek,
  parseEmailList,
  stripLeadingAt,
} from "./create_calendar_event.ts";

const { createContext } = SlackFunctionTester("create_calendar_event");

const FAKE_ENV = {
  MS_TENANT_ID: "fake-tenant-id",
  MS_CLIENT_ID: "fake-client-id",
  MS_CLIENT_SECRET: "fake-client-secret",
  MS_SHARED_MAILBOX: "emtech-scheduling@example.com",
  // Matches the real MS_EVENT_TIMEZONE — Emtech operations are Pacific-based.
  MS_EVENT_TIMEZONE: "America/Los_Angeles",
};

const BASE_INPUTS = {
  amazon_alias: "jdoe",
  // Simulates a Slack Person-type variable rendering as "@Display Name" when
  // interpolated into a plain string — the real-world case that motivated
  // stripLeadingAt().
  submitter_name: "@John Doe",
  submitter_email: "jdoe@example.com",
  request_type: "Sick",
  // Midnight UTC, matching what Slack's date field actually sends for a
  // plain calendar-date selection (no real time-of-day intended).
  start_date_time: "2026-08-01T00:00:00Z",
  end_date_time: "2026-08-01T00:00:00Z",
  description: "Out of office",
  location: "N/A",
  // Slack user references (e.g. from a List's Person column) — resolved to
  // real email addresses internally via client.users.info.
  additional_attendees: ["U0MANAGER1", "U0BACKUP2"],
};

const RESOLVABLE_ATTENDEE_EMAILS: Record<string, string> = {
  "U0MANAGER1": "manager@example.com",
  "U0BACKUP2": "backup@example.com",
};

// Stubs the Slack Web API calls our attendee-resolution logic makes,
// alongside whatever Microsoft Graph stub the test itself provides.
// emailsByUserId maps user IDs to emails; omit an ID to simulate an
// unresolvable attendee (no email in their profile).
function stubSlackApi(
  emailsByUserId: Record<string, string>,
  options?: {
    onLookupByEmail?: (email: string) => { id: string } | undefined;
    onPostMessage?: (args: { channel: string; text: string }) => void;
  },
) {
  return async (request: Request): Promise<Response | undefined> => {
    if (request.url.includes("slack.com/api/users.info")) {
      const params = await request.formData();
      const userId = params.get("user") as string;
      const email = emailsByUserId[userId];
      return new Response(
        JSON.stringify({ ok: true, user: { profile: { email } } }),
        { status: 200 },
      );
    }
    if (request.url.includes("slack.com/api/users.lookupByEmail")) {
      const params = await request.formData();
      const email = params.get("email") as string;
      const user = options?.onLookupByEmail?.(email);
      return new Response(JSON.stringify({ ok: Boolean(user), user }), {
        status: 200,
      });
    }
    if (request.url.includes("slack.com/api/chat.postMessage")) {
      const params = await request.formData();
      options?.onPostMessage?.({
        channel: params.get("channel") as string,
        text: params.get("text") as string,
      });
      return new Response(JSON.stringify({ ok: true }), { status: 200 });
    }
    return undefined;
  };
}

Deno.test("stripLeadingAt removes a leading @ and any following whitespace", () => {
  assertEquals(stripLeadingAt("@John Doe"), "John Doe");
  assertEquals(stripLeadingAt("@ John Doe"), "John Doe");
  assertEquals(stripLeadingAt("John Doe"), "John Doe");
  assertEquals(stripLeadingAt("cragandb"), "cragandb");
});

const PACIFIC = "America/Los_Angeles";

Deno.test("formatTitleDate uses the literal date, ignoring any time/offset suffix", () => {
  assertEquals(formatTitleDate("2026-08-01T00:00:00Z", PACIFIC), "Aug 1");
  assertEquals(formatTitleDate("2026-08-01T23:59:00Z", PACIFIC), "Aug 1");
  assertEquals(formatTitleDate("2026-08-01T09:00:00", PACIFIC), "Aug 1");
});

Deno.test("formatTitleDate omits the year", () => {
  assertEquals(formatTitleDate("2026-08-01T00:00:00Z", PACIFIC), "Aug 1");
  assertEquals(formatTitleDate("2027-01-15T00:00:00Z", PACIFIC), "Jan 15");
});

Deno.test("formatTitleDate does not roll a midnight-UTC date back a day (regression)", () => {
  // This is the exact bug that was reported: a date-only value serialized
  // as midnight UTC ("2026-07-08T00:00:00Z") was being timezone-converted
  // to America/Los_Angeles, landing on July 7 evening instead of July 8.
  assertEquals(formatTitleDate("2026-07-08T00:00:00Z", PACIFIC), "Jul 8");
});

Deno.test("formatTitleDate converts a Unix timestamp into the target timezone (regression)", () => {
  // Real bug: a "Date and time" field (used for OOTO's partial-day support)
  // sends a bare Unix timestamp, not an ISO string, and the code crashed
  // trying to parse it as one. This is the exact value from that report:
  // 1784235600 = 2026-07-16T21:00:00Z = 2026-07-16T14:00:00 Pacific.
  assertEquals(formatTitleDate("1784235600", PACIFIC), "Jul 16");
});

Deno.test("getAllDayEventRange covers a single day with an exclusive end", () => {
  const range = getAllDayEventRange(
    "2026-07-09T00:00:00Z",
    "2026-07-09T00:00:00Z",
    PACIFIC,
  );
  assertEquals(range.start, "2026-07-09T00:00:00");
  assertEquals(range.end, "2026-07-10T00:00:00");
});

Deno.test("getAllDayEventRange spans multiple days correctly (regression)", () => {
  // The reported bug, reproduced directly: submitting 7/8 to 7/9 must not
  // land on 7/7 to 7/8.
  const range = getAllDayEventRange(
    "2026-07-08T00:00:00Z",
    "2026-07-09T00:00:00Z",
    PACIFIC,
  );
  assertEquals(range.start, "2026-07-08T00:00:00");
  assertEquals(range.end, "2026-07-10T00:00:00");
});

Deno.test("getAllDayEventRange rolls over month and year boundaries", () => {
  const monthRollover = getAllDayEventRange(
    "2026-07-31T00:00:00Z",
    "2026-07-31T00:00:00Z",
    PACIFIC,
  );
  assertEquals(monthRollover.end, "2026-08-01T00:00:00");

  const yearRollover = getAllDayEventRange(
    "2026-12-31T00:00:00Z",
    "2026-12-31T00:00:00Z",
    PACIFIC,
  );
  assertEquals(yearRollover.end, "2027-01-01T00:00:00");
});

Deno.test("getAllDayEventRange ignores time-of-day and offset entirely", () => {
  // Late in the UTC day, with a full offset — both should still resolve to
  // the literal July 9 date, not shift because of the "Z"/offset suffix.
  const range = getAllDayEventRange(
    "2026-07-09T23:00:00Z",
    "2026-07-09T05:00:00+05:00",
    PACIFIC,
  );
  assertEquals(range.start, "2026-07-09T00:00:00");
  assertEquals(range.end, "2026-07-10T00:00:00");
});

Deno.test("getAllDayEventRange resolves a Unix timestamp via the target timezone (regression)", () => {
  // A full-day OOTO request also goes through a "Date and time" field now
  // (it's the same field type as the partial-day case, just unused time
  // portion), so this path must handle a Unix timestamp too, not just ISO.
  const range = getAllDayEventRange("1784235600", "1784235600", PACIFIC);
  assertEquals(range.start, "2026-07-16T00:00:00");
  assertEquals(range.end, "2026-07-17T00:00:00");
});

Deno.test("getPartialDayEventRange uses the literal date and time-of-day, unlike the all-day path", () => {
  const range = getPartialDayEventRange(
    "2026-07-20T09:00:00",
    "2026-07-20T13:30:00",
    PACIFIC,
  );
  assertEquals(range.start, "2026-07-20T09:00:00");
  assertEquals(range.end, "2026-07-20T13:30:00");
});

Deno.test("getPartialDayEventRange converts Unix timestamps via the target timezone (regression)", () => {
  // The exact values from the real bug report: a 2:00 PM-7:00 PM Pacific
  // partial-day request came through as these two raw epoch seconds.
  const range = getPartialDayEventRange("1784235600", "1784253600", PACIFIC);
  assertEquals(range.start, "2026-07-16T14:00:00");
  assertEquals(range.end, "2026-07-16T19:00:00");
});

Deno.test("isPartialDayRequest matches 'yes' regardless of case or surrounding whitespace", () => {
  assertEquals(isPartialDayRequest("Yes"), true);
  assertEquals(isPartialDayRequest("yes"), true);
  assertEquals(isPartialDayRequest("YES"), true);
  assertEquals(isPartialDayRequest("  Yes  "), true);
  assertEquals(isPartialDayRequest("No"), false);
  assertEquals(isPartialDayRequest(undefined), false);
  assertEquals(isPartialDayRequest(""), false);
});

Deno.test("isRecurringRequest matches 'yes' regardless of case or surrounding whitespace", () => {
  assertEquals(isRecurringRequest("Yes"), true);
  assertEquals(isRecurringRequest("yes"), true);
  assertEquals(isRecurringRequest("  YES  "), true);
  assertEquals(isRecurringRequest("No"), false);
  assertEquals(isRecurringRequest(undefined), false);
});

Deno.test("normalizeDayOfWeek matches regardless of case, returning the canonical name", () => {
  assertEquals(normalizeDayOfWeek("Friday"), "Friday");
  assertEquals(normalizeDayOfWeek("friday"), "Friday");
  assertEquals(normalizeDayOfWeek("FRIDAY"), "Friday");
  assertEquals(normalizeDayOfWeek("  Monday  "), "Monday");
  assertEquals(normalizeDayOfWeek("Someday"), undefined);
  assertEquals(normalizeDayOfWeek(undefined), undefined);
});

Deno.test("buildWeeklyRecurrence builds a weekly pattern anchored on the start date", () => {
  const recurrence = buildWeeklyRecurrence(
    "2026-07-24T00:00:00Z",
    "2026-12-31T00:00:00Z",
    "Friday",
    PACIFIC,
  );
  assertEquals(recurrence, {
    pattern: { type: "weekly", interval: 1, daysOfWeek: ["Friday"] },
    range: {
      type: "endDate",
      startDate: "2026-07-24",
      endDate: "2026-12-31",
    },
  });
});

Deno.test("buildWeeklyRecurrence doesn't require start_date_time's weekday to match dayOfWeek", () => {
  // 2026-07-24 is a Friday, but that's incidental — Graph anchors from
  // startDate forward and finds the first matching weekday itself, so an
  // unrelated day (e.g. a Monday) works exactly the same way.
  const recurrence = buildWeeklyRecurrence(
    "2026-07-20T00:00:00Z", // a Monday
    "2026-12-31T00:00:00Z",
    "Friday",
    PACIFIC,
  );
  assertEquals(recurrence.range.startDate, "2026-07-20");
  assertEquals(recurrence.pattern.daysOfWeek, ["Friday"]);
});

Deno.test("parseEmailList splits on whitespace, commas, and semicolons", () => {
  assertEquals(
    parseEmailList("a@x.com b@y.com"),
    { emails: ["a@x.com", "b@y.com"] },
  );
  assertEquals(
    parseEmailList("a@x.com, b@y.com;c@z.com"),
    { emails: ["a@x.com", "b@y.com", "c@z.com"] },
  );
  assertEquals(
    parseEmailList("  a@x.com   b@y.com  "),
    { emails: ["a@x.com", "b@y.com"] },
  );
});

Deno.test("parseEmailList handles empty/undefined input", () => {
  assertEquals(parseEmailList(undefined), { emails: [] });
  assertEquals(parseEmailList(""), { emails: [] });
  assertEquals(parseEmailList("   "), { emails: [] });
});

Deno.test("parseEmailList returns an error naming the first invalid entry", () => {
  const result = parseEmailList("a@x.com not-an-email b@y.com");
  assertEquals("error" in result, true);
  assertStringIncludes((result as { error: string }).error, "not-an-email");
});

Deno.test("parseEmailList strips a leading mailto: from each entry (regression)", () => {
  // Real bug: Slack's List Text column auto-linkifies emails, and the
  // "mailto:" scheme survived into the plain string variable. Graph/Exchange
  // silently accepted "mailto:name@x.com" as an attendee address but never
  // delivered an invite to the real address underneath.
  assertEquals(
    parseEmailList("mailto:a@x.com Mailto:b@y.com MAILTO:c@z.com"),
    { emails: ["a@x.com", "b@y.com", "c@z.com"] },
  );
});

Deno.test("getCategoryForRequestType maps every known Request Type value", () => {
  assertEquals(getCategoryForRequestType("Sick"), "OOTO");
  assertEquals(getCategoryForRequestType("Vacation"), "OOTO");
  assertEquals(getCategoryForRequestType("OOTO"), "OOTO");
  assertEquals(getCategoryForRequestType("4x10 OOTO"), "4x10 OOTO");
  assertEquals(getCategoryForRequestType("WFH"), "WFH");
  assertEquals(getCategoryForRequestType("Location Assignment"), "On-Site");
  assertEquals(getCategoryForRequestType("Scheduling"), undefined);
  assertEquals(getCategoryForRequestType("Something Unrecognized"), undefined);
});

Deno.test("getCategoryForRequestType matches regardless of case or surrounding whitespace (regression)", () => {
  // Real bug: the dropdown's actual value ("4X10 OOTO") didn't match this
  // table's key ("4x10 OOTO") under an exact-case lookup, so a legitimate
  // request failed with "No Outlook category mapping configured...".
  assertEquals(getCategoryForRequestType("4X10 OOTO"), "4x10 OOTO");
  assertEquals(getCategoryForRequestType("4x10 ooto"), "4x10 OOTO");
  assertEquals(getCategoryForRequestType("sick"), "OOTO");
  assertEquals(getCategoryForRequestType("  WFH  "), "WFH");
});

Deno.test("create_calendar_event happy path sends the expected Graph request", async () => {
  let capturedBody: Record<string, unknown> | undefined;
  const slackApiStub = stubSlackApi(RESOLVABLE_ATTENDEE_EMAILS);

  using _stubFetch = stub(
    globalThis,
    "fetch",
    async (url: string | URL | Request, options?: RequestInit) => {
      const request = url instanceof Request ? url : new Request(url, options);

      if (request.url.includes("login.microsoftonline.com")) {
        return new Response(
          JSON.stringify({ access_token: "fake-access-token" }),
          { status: 200 },
        );
      }

      const slackResponse = await slackApiStub(request);
      if (slackResponse) return slackResponse;

      // graph.microsoft.com/v1.0/users/{mailbox}/events
      capturedBody = await request.json();
      return new Response(
        JSON.stringify({
          id: "AAMkAGI1AAA=",
          webLink: "https://outlook.office.com/calendar/item/AAMkAGI1AAA%3D",
        }),
        { status: 201 },
      );
    },
  );

  const { outputs, error } = await CreateCalendarEvent(
    createContext({ inputs: BASE_INPUTS, env: FAKE_ENV }),
  );

  assertEquals(error, undefined);
  assertExists(capturedBody);
  assertEquals(capturedBody?.subject, "@jdoe - Sick - Aug 1 - John Doe");
  assertEquals(capturedBody?.showAs, "free");
  assertEquals(capturedBody?.attendees, [
    { emailAddress: { address: "jdoe@example.com" }, type: "required" },
    { emailAddress: { address: "manager@example.com" }, type: "required" },
    { emailAddress: { address: "backup@example.com" }, type: "required" },
  ]);
  assertEquals(capturedBody?.isAllDay, true);
  assertEquals(capturedBody?.start, {
    dateTime: "2026-08-01T00:00:00",
    timeZone: "America/Los_Angeles",
  });
  assertEquals(capturedBody?.end, {
    dateTime: "2026-08-02T00:00:00",
    timeZone: "America/Los_Angeles",
  });
  assertEquals(capturedBody?.categories, ["OOTO"]);
  assertEquals(outputs?.event_id, "AAMkAGI1AAA=");
  assertEquals(
    outputs?.web_link,
    "https://outlook.office.com/calendar/item/AAMkAGI1AAA%3D",
  );
});

Deno.test('create_calendar_event creates a timed (non-all-day) event when is_partial_day is "Yes"', async () => {
  let capturedBody: Record<string, unknown> | undefined;
  const slackApiStub = stubSlackApi(RESOLVABLE_ATTENDEE_EMAILS);

  using _stubFetch = stub(
    globalThis,
    "fetch",
    async (url: string | URL | Request, options?: RequestInit) => {
      const request = url instanceof Request ? url : new Request(url, options);

      if (request.url.includes("login.microsoftonline.com")) {
        return new Response(
          JSON.stringify({ access_token: "fake-access-token" }),
          { status: 200 },
        );
      }

      const slackResponse = await slackApiStub(request);
      if (slackResponse) return slackResponse;

      capturedBody = await request.json();
      return new Response(
        JSON.stringify({ id: "AAMkAGI1AAA=", webLink: "https://x" }),
        { status: 201 },
      );
    },
  );

  const { error } = await CreateCalendarEvent(
    createContext({
      inputs: {
        ...BASE_INPUTS,
        request_type: "OOTO",
        is_partial_day: "Yes",
        start_date_time: "2026-07-20T09:00:00",
        end_date_time: "2026-07-20T13:30:00",
      },
      env: FAKE_ENV,
    }),
  );

  assertEquals(error, undefined);
  assertEquals(capturedBody?.isAllDay, false);
  assertEquals(capturedBody?.start, {
    dateTime: "2026-07-20T09:00:00",
    timeZone: "America/Los_Angeles",
  });
  assertEquals(capturedBody?.end, {
    dateTime: "2026-07-20T13:30:00",
    timeZone: "America/Los_Angeles",
  });
});

Deno.test("create_calendar_event handles a partial-day request sourced from a Date-and-time field's Unix timestamp (regression)", async () => {
  // Reproduces the real reported bug end-to-end: a "Date and time" field
  // sends a bare Unix timestamp rather than an ISO string, and the function
  // crashed with "Expected an ISO date string, got \"1784235600\"".
  let capturedBody: Record<string, unknown> | undefined;
  const slackApiStub = stubSlackApi(RESOLVABLE_ATTENDEE_EMAILS);

  using _stubFetch = stub(
    globalThis,
    "fetch",
    async (url: string | URL | Request, options?: RequestInit) => {
      const request = url instanceof Request ? url : new Request(url, options);

      if (request.url.includes("login.microsoftonline.com")) {
        return new Response(
          JSON.stringify({ access_token: "fake-access-token" }),
          { status: 200 },
        );
      }

      const slackResponse = await slackApiStub(request);
      if (slackResponse) return slackResponse;

      capturedBody = await request.json();
      return new Response(
        JSON.stringify({ id: "AAMkAGI1AAA=", webLink: "https://x" }),
        { status: 201 },
      );
    },
  );

  const { error } = await CreateCalendarEvent(
    createContext({
      inputs: {
        ...BASE_INPUTS,
        request_type: "OOTO",
        is_partial_day: "Yes",
        start_date_time: "1784235600",
        end_date_time: "1784253600",
      },
      env: FAKE_ENV,
    }),
  );

  assertEquals(error, undefined);
  assertEquals(capturedBody?.isAllDay, false);
  assertEquals(capturedBody?.subject, "@jdoe - OOTO - Jul 16 - John Doe");
  assertEquals(capturedBody?.start, {
    dateTime: "2026-07-16T14:00:00",
    timeZone: "America/Los_Angeles",
  });
  assertEquals(capturedBody?.end, {
    dateTime: "2026-07-16T19:00:00",
    timeZone: "America/Los_Angeles",
  });
});

Deno.test("create_calendar_event handles a full-day OOTO request sourced from a Date-and-time field (regression)", async () => {
  // A full-day (non-partial) OOTO request goes through the same Date-and-
  // time field type, so it must also survive a Unix-timestamp value.
  let capturedBody: Record<string, unknown> | undefined;
  const slackApiStub = stubSlackApi(RESOLVABLE_ATTENDEE_EMAILS);

  using _stubFetch = stub(
    globalThis,
    "fetch",
    async (url: string | URL | Request, options?: RequestInit) => {
      const request = url instanceof Request ? url : new Request(url, options);

      if (request.url.includes("login.microsoftonline.com")) {
        return new Response(
          JSON.stringify({ access_token: "fake-access-token" }),
          { status: 200 },
        );
      }

      const slackResponse = await slackApiStub(request);
      if (slackResponse) return slackResponse;

      capturedBody = await request.json();
      return new Response(
        JSON.stringify({ id: "AAMkAGI1AAA=", webLink: "https://x" }),
        { status: 201 },
      );
    },
  );

  const { error } = await CreateCalendarEvent(
    createContext({
      inputs: {
        ...BASE_INPUTS,
        request_type: "OOTO",
        start_date_time: "1784235600",
        end_date_time: "1784235600",
      },
      env: FAKE_ENV,
    }),
  );

  assertEquals(error, undefined);
  assertEquals(capturedBody?.isAllDay, true);
  assertEquals(capturedBody?.start, {
    dateTime: "2026-07-16T00:00:00",
    timeZone: "America/Los_Angeles",
  });
  assertEquals(capturedBody?.end, {
    dateTime: "2026-07-17T00:00:00",
    timeZone: "America/Los_Angeles",
  });
});

Deno.test('create_calendar_event attaches a weekly recurrence when is_recurring is "Yes"', async () => {
  let capturedBody: Record<string, unknown> | undefined;
  const slackApiStub = stubSlackApi(RESOLVABLE_ATTENDEE_EMAILS);

  using _stubFetch = stub(
    globalThis,
    "fetch",
    async (url: string | URL | Request, options?: RequestInit) => {
      const request = url instanceof Request ? url : new Request(url, options);

      if (request.url.includes("login.microsoftonline.com")) {
        return new Response(
          JSON.stringify({ access_token: "fake-access-token" }),
          { status: 200 },
        );
      }

      const slackResponse = await slackApiStub(request);
      if (slackResponse) return slackResponse;

      capturedBody = await request.json();
      return new Response(
        JSON.stringify({ id: "AAMkAGI1AAA=", webLink: "https://x" }),
        { status: 201 },
      );
    },
  );

  const { error } = await CreateCalendarEvent(
    createContext({
      inputs: {
        ...BASE_INPUTS,
        request_type: "4x10 OOTO",
        start_date_time: "2026-07-24T00:00:00Z",
        end_date_time: "2026-07-24T00:00:00Z",
        is_recurring: "Yes",
        recurrence_end_date: "2026-12-31T00:00:00Z",
        recurrence_day_of_week: "friday",
      },
      env: FAKE_ENV,
    }),
  );

  assertEquals(error, undefined);
  // Still a normal all-day event otherwise — recurrence is additive.
  assertEquals(capturedBody?.isAllDay, true);
  assertEquals(capturedBody?.recurrence, {
    pattern: { type: "weekly", interval: 1, daysOfWeek: ["Friday"] },
    range: {
      type: "endDate",
      startDate: "2026-07-24",
      endDate: "2026-12-31",
    },
  });
});

Deno.test("create_calendar_event fails fast when is_recurring is true but recurrence fields are missing", async () => {
  using _stubFetch = stub(
    globalThis,
    "fetch",
    () => {
      throw new Error(
        "fetch should never be called when recurrence fields are missing",
      );
    },
  );

  const { outputs, error } = await CreateCalendarEvent(
    createContext({
      inputs: { ...BASE_INPUTS, is_recurring: "Yes" },
      env: FAKE_ENV,
    }),
  );

  assertExists(error);
  assertStringIncludes(error, "recurrence_end_date");
  assertStringIncludes(error, "recurrence_day_of_week");
  assertEquals(outputs, undefined);
});

Deno.test("create_calendar_event fails fast on an unrecognized recurrence_day_of_week", async () => {
  using _stubFetch = stub(
    globalThis,
    "fetch",
    () => {
      throw new Error(
        "fetch should never be called for an unrecognized day of week",
      );
    },
  );

  const { outputs, error } = await CreateCalendarEvent(
    createContext({
      inputs: {
        ...BASE_INPUTS,
        is_recurring: "Yes",
        recurrence_end_date: "2026-12-31T00:00:00Z",
        recurrence_day_of_week: "Someday",
      },
      env: FAKE_ENV,
    }),
  );

  assertExists(error);
  assertStringIncludes(error, "Someday");
  assertEquals(outputs, undefined);
});

Deno.test("create_calendar_event fails fast when is_partial_day and is_recurring are both true", async () => {
  using _stubFetch = stub(
    globalThis,
    "fetch",
    () => {
      throw new Error(
        "fetch should never be called when partial-day and recurring are both set",
      );
    },
  );

  const { outputs, error } = await CreateCalendarEvent(
    createContext({
      inputs: {
        ...BASE_INPUTS,
        is_partial_day: "Yes",
        is_recurring: "Yes",
        recurrence_end_date: "2026-12-31T00:00:00Z",
        recurrence_day_of_week: "Friday",
      },
      env: FAKE_ENV,
    }),
  );

  assertExists(error);
  assertStringIncludes(error, "is_partial_day");
  assertStringIncludes(error, "is_recurring");
  assertEquals(outputs, undefined);
});

Deno.test("create_calendar_event combines resolved internal attendees with plain external emails", async () => {
  let capturedBody: Record<string, unknown> | undefined;
  const slackApiStub = stubSlackApi(RESOLVABLE_ATTENDEE_EMAILS);

  using _stubFetch = stub(
    globalThis,
    "fetch",
    async (url: string | URL | Request, options?: RequestInit) => {
      const request = url instanceof Request ? url : new Request(url, options);

      if (request.url.includes("login.microsoftonline.com")) {
        return new Response(
          JSON.stringify({ access_token: "fake-access-token" }),
          { status: 200 },
        );
      }

      const slackResponse = await slackApiStub(request);
      if (slackResponse) return slackResponse;

      capturedBody = await request.json();
      return new Response(
        JSON.stringify({ id: "AAMkAGI1AAA=", webLink: "https://x" }),
        { status: 201 },
      );
    },
  );

  const { error } = await CreateCalendarEvent(
    createContext({
      inputs: {
        ...BASE_INPUTS,
        // Mixed delimiters, matching however someone naturally types a list.
        external_attendees:
          "client@external.example.com; second@external.example.com",
      },
      env: FAKE_ENV,
    }),
  );

  assertEquals(error, undefined);
  assertEquals(capturedBody?.attendees, [
    { emailAddress: { address: "jdoe@example.com" }, type: "required" },
    { emailAddress: { address: "manager@example.com" }, type: "required" },
    { emailAddress: { address: "backup@example.com" }, type: "required" },
    {
      emailAddress: { address: "client@external.example.com" },
      type: "required",
    },
    {
      emailAddress: { address: "second@external.example.com" },
      type: "required",
    },
  ]);
});

Deno.test("create_calendar_event fails fast on an invalid external_attendees entry", async () => {
  using _stubFetch = stub(
    globalThis,
    "fetch",
    () => {
      throw new Error(
        "fetch should never be called for an invalid external_attendees entry",
      );
    },
  );

  const { outputs, error } = await CreateCalendarEvent(
    createContext({
      inputs: { ...BASE_INPUTS, external_attendees: "not-an-email" },
      env: FAKE_ENV,
    }),
  );

  assertExists(error);
  assertStringIncludes(error, "not-an-email");
  assertEquals(outputs, undefined);
});

Deno.test("create_calendar_event still creates the event when an attendee can't be resolved, and DMs the submitter", async () => {
  let capturedBody: Record<string, unknown> | undefined;
  let postedMessage: { channel: string; text: string } | undefined;

  // U0BACKUP2 has no email in this stub, simulating an unresolvable attendee.
  const slackApiStub = stubSlackApi({ "U0MANAGER1": "manager@example.com" }, {
    onLookupByEmail: (email) =>
      email === "jdoe@example.com" ? { id: "U0SUBMITTER" } : undefined,
    onPostMessage: (args) => {
      postedMessage = args;
    },
  });

  using _stubFetch = stub(
    globalThis,
    "fetch",
    async (url: string | URL | Request, options?: RequestInit) => {
      const request = url instanceof Request ? url : new Request(url, options);

      if (request.url.includes("login.microsoftonline.com")) {
        return new Response(
          JSON.stringify({ access_token: "fake-access-token" }),
          { status: 200 },
        );
      }

      const slackResponse = await slackApiStub(request);
      if (slackResponse) return slackResponse;

      capturedBody = await request.json();
      return new Response(
        JSON.stringify({ id: "AAMkAGI1AAA=", webLink: "https://x" }),
        { status: 201 },
      );
    },
  );

  const { outputs, error } = await CreateCalendarEvent(
    createContext({ inputs: BASE_INPUTS, env: FAKE_ENV }),
  );

  assertEquals(error, undefined);
  assertExists(outputs);
  // Only the resolvable attendee made it onto the event.
  assertEquals(capturedBody?.attendees, [
    { emailAddress: { address: "jdoe@example.com" }, type: "required" },
    { emailAddress: { address: "manager@example.com" }, type: "required" },
  ]);
  // The submitter was notified about the one that couldn't be added.
  assertExists(postedMessage);
  assertEquals(postedMessage?.channel, "U0SUBMITTER");
  assertStringIncludes(postedMessage?.text ?? "", "<@U0BACKUP2>");
});

Deno.test("create_calendar_event fails fast when MS_SHARED_MAILBOX is missing", async () => {
  using _stubFetch = stub(
    globalThis,
    "fetch",
    () => {
      throw new Error("fetch should never be called without MS_SHARED_MAILBOX");
    },
  );

  const { MS_SHARED_MAILBOX: _omit, ...envWithoutMailbox } = FAKE_ENV;
  const { outputs, error } = await CreateCalendarEvent(
    createContext({ inputs: BASE_INPUTS, env: envWithoutMailbox }),
  );

  assertExists(error);
  assertStringIncludes(error, "MS_SHARED_MAILBOX");
  assertEquals(outputs, undefined);
});

Deno.test("create_calendar_event fails fast on an unrecognized request_type", async () => {
  using _stubFetch = stub(
    globalThis,
    "fetch",
    () => {
      throw new Error(
        "fetch should never be called for an unrecognized request_type",
      );
    },
  );

  const { outputs, error } = await CreateCalendarEvent(
    createContext({
      inputs: { ...BASE_INPUTS, request_type: "Scheduling" },
      env: FAKE_ENV,
    }),
  );

  assertExists(error);
  assertStringIncludes(error, "No Outlook category mapping");
  assertStringIncludes(error, "Scheduling");
  assertEquals(outputs, undefined);
});

Deno.test("create_calendar_event surfaces a clean error when the Graph token request fails", async () => {
  let eventsEndpointCalled = false;

  using _stubFetch = stub(
    globalThis,
    "fetch",
    (url: string | URL | Request, options?: RequestInit) => {
      const request = url instanceof Request ? url : new Request(url, options);

      if (request.url.includes("login.microsoftonline.com")) {
        return Promise.resolve(
          new Response("invalid_client", { status: 401 }),
        );
      }

      eventsEndpointCalled = true;
      return Promise.resolve(
        new Response("should not be reached", { status: 500 }),
      );
    },
  );

  const { outputs, error } = await CreateCalendarEvent(
    createContext({ inputs: BASE_INPUTS, env: FAKE_ENV }),
  );

  assertExists(error);
  assertStringIncludes(error, "Failed to acquire Graph access token");
  assertEquals(eventsEndpointCalled, false);
  assertEquals(outputs, undefined);
});
