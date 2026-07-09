import { SlackFunctionTester } from "deno-slack-sdk/mod.ts";
import { assertEquals, assertExists, assertStringIncludes } from "@std/assert";
import { stub } from "@std/testing/mock";
import CreateCalendarEvent, {
  formatTitleDate,
  getAllDayEventRange,
  getCategoryForRequestType,
  parseEmailList,
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
  submitter_alias: "jdoe",
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

Deno.test("formatTitleDate uses the literal date, ignoring any time/offset suffix", () => {
  assertEquals(formatTitleDate("2026-08-01T00:00:00Z"), "Aug 1, 2026");
  assertEquals(formatTitleDate("2026-08-01T23:59:00Z"), "Aug 1, 2026");
  assertEquals(formatTitleDate("2026-08-01T09:00:00"), "Aug 1, 2026");
});

Deno.test("formatTitleDate does not roll a midnight-UTC date back a day (regression)", () => {
  // This is the exact bug that was reported: a date-only value serialized
  // as midnight UTC ("2026-07-08T00:00:00Z") was being timezone-converted
  // to America/Los_Angeles, landing on July 7 evening instead of July 8.
  assertEquals(formatTitleDate("2026-07-08T00:00:00Z"), "Jul 8, 2026");
});

Deno.test("getAllDayEventRange covers a single day with an exclusive end", () => {
  const range = getAllDayEventRange(
    "2026-07-09T00:00:00Z",
    "2026-07-09T00:00:00Z",
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
  );
  assertEquals(range.start, "2026-07-08T00:00:00");
  assertEquals(range.end, "2026-07-10T00:00:00");
});

Deno.test("getAllDayEventRange rolls over month and year boundaries", () => {
  const monthRollover = getAllDayEventRange(
    "2026-07-31T00:00:00Z",
    "2026-07-31T00:00:00Z",
  );
  assertEquals(monthRollover.end, "2026-08-01T00:00:00");

  const yearRollover = getAllDayEventRange(
    "2026-12-31T00:00:00Z",
    "2026-12-31T00:00:00Z",
  );
  assertEquals(yearRollover.end, "2027-01-01T00:00:00");
});

Deno.test("getAllDayEventRange ignores time-of-day and offset entirely", () => {
  // Late in the UTC day, with a full offset — both should still resolve to
  // the literal July 9 date, not shift because of the "Z"/offset suffix.
  const range = getAllDayEventRange(
    "2026-07-09T23:00:00Z",
    "2026-07-09T05:00:00+05:00",
  );
  assertEquals(range.start, "2026-07-09T00:00:00");
  assertEquals(range.end, "2026-07-10T00:00:00");
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
  assertEquals(capturedBody?.subject, "jdoe - Sick - Aug 1, 2026");
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
