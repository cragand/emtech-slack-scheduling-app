import { SlackFunctionTester } from "deno-slack-sdk/mod.ts";
import { assertEquals, assertExists, assertStringIncludes } from "@std/assert";
import { stub } from "@std/testing/mock";
import CreateCalendarEvent, {
  formatTitleDate,
  getAllDayEventRange,
  getCategoryForRequestType,
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
  additional_attendees: ["manager@example.com", "backup@example.com"],
};

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

Deno.test("create_calendar_event happy path sends the expected Graph request", async () => {
  let capturedBody: Record<string, unknown> | undefined;

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
