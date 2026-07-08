import { SlackFunctionTester } from "deno-slack-sdk/mod.ts";
import { assertEquals, assertExists, assertStringIncludes } from "@std/assert";
import { stub } from "@std/testing/mock";
import CreateCalendarEvent, {
  formatTitleDate,
  getCategoryForRequestType,
} from "./create_calendar_event.ts";

const { createContext } = SlackFunctionTester("create_calendar_event");

const FAKE_ENV = {
  MS_TENANT_ID: "fake-tenant-id",
  MS_CLIENT_ID: "fake-client-id",
  MS_CLIENT_SECRET: "fake-client-secret",
  MS_SHARED_MAILBOX: "emtech-scheduling@example.com",
  MS_EVENT_TIMEZONE: "America/New_York",
};

const BASE_INPUTS = {
  submitter_alias: "jdoe",
  submitter_email: "jdoe@example.com",
  request_type: "Sick",
  // Explicit UTC ("Z") times rather than naive local strings, so this test
  // is deterministic regardless of the machine/CI running it — 13:00Z /
  // 21:00Z is 09:00 / 17:00 America/New_York (EDT, UTC-4 in August).
  start_date_time: "2026-08-01T13:00:00Z",
  end_date_time: "2026-08-01T21:00:00Z",
  description: "Out of office",
  location: "N/A",
  additional_attendees: ["manager@example.com", "backup@example.com"],
};

Deno.test("formatTitleDate formats using the given timeZone, not the system default", () => {
  assertEquals(
    formatTitleDate("2026-08-01T13:00:00Z", "America/New_York"),
    "Aug 1, 2026",
  );
});

Deno.test("formatTitleDate can land on a different calendar day depending on timeZone", () => {
  // 04:00 UTC is still July 9 in America/Los_Angeles (PDT, UTC-7) but
  // already July 10 further east — this is the exact class of bug that
  // caused the title's date to disagree with the actual event date when
  // formatTitleDate didn't take an explicit timeZone.
  assertEquals(
    formatTitleDate("2026-07-10T04:00:00Z", "America/Los_Angeles"),
    "Jul 9, 2026",
  );
  assertEquals(
    formatTitleDate("2026-07-10T04:00:00Z", "America/New_York"),
    "Jul 10, 2026",
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
  assertEquals(
    (capturedBody?.start as { dateTime: string; timeZone: string }).timeZone,
    "America/New_York",
  );
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
