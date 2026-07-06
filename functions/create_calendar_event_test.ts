import { SlackFunctionTester } from "deno-slack-sdk/mod.ts";
import { assertEquals, assertExists, assertStringIncludes } from "@std/assert";
import { stub } from "@std/testing/mock";
import CreateCalendarEvent, {
  formatTitleDate,
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
  request_type: "Time Off",
  start_date_time: "2026-08-01T09:00:00",
  end_date_time: "2026-08-01T17:00:00",
  description: "Out of office",
  location: "N/A",
  additional_attendees: ["manager@example.com", "backup@example.com"],
};

Deno.test("formatTitleDate formats an ISO datetime as a short date", () => {
  assertEquals(formatTitleDate("2026-08-01T09:00:00"), "Aug 1, 2026");
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
  assertEquals(capturedBody?.subject, "jdoe - Time Off - Aug 1, 2026");
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
