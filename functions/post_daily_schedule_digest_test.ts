import { SlackFunctionTester } from "deno-slack-sdk/mod.ts";
import { assertEquals, assertExists, assertStringIncludes } from "@std/assert";
import { stub } from "@std/testing/mock";
import PostDailyScheduleDigest, {
  buildDigestMessage,
  formatDigestLine,
} from "./post_daily_schedule_digest.ts";

const { createContext } = SlackFunctionTester("post_daily_schedule_digest");

const FAKE_ENV = {
  MS_TENANT_ID: "fake-tenant-id",
  MS_CLIENT_ID: "fake-client-id",
  MS_CLIENT_SECRET: "fake-client-secret",
  MS_SHARED_MAILBOX: "emtech-scheduling@example.com",
  MS_EVENT_TIMEZONE: "America/Los_Angeles",
};

Deno.test("formatDigestLine wraps the subject as a link to the event's webLink", () => {
  assertEquals(
    formatDigestLine({
      subject: "@jdoe - Sick - Jul 20 - John Doe",
      webLink: "https://outlook.office.com/calendar/item/abc",
    }),
    "• <https://outlook.office.com/calendar/item/abc|@jdoe - Sick - Jul 20 - John Doe>",
  );
});

Deno.test("formatDigestLine falls back to plain text if webLink is missing", () => {
  assertEquals(
    formatDigestLine({ subject: "@jdoe - Sick - Jul 20 - John Doe" }),
    "• @jdoe - Sick - Jul 20 - John Doe",
  );
});

Deno.test("formatDigestLine falls back to a placeholder if subject is missing", () => {
  assertStringIncludes(
    formatDigestLine({ webLink: "https://outlook.office.com/x" }),
    "(untitled event)",
  );
});

Deno.test("buildDigestMessage lists every event as a bulleted line", () => {
  const message = buildDigestMessage([
    {
      subject: "@jdoe - Sick - Jul 20 - John Doe",
      webLink: "https://outlook.office.com/a",
    },
    {
      subject: "@asmith - Vacation - Jul 20 - Alex Smith",
      webLink: "https://outlook.office.com/b",
    },
  ]);
  assertEquals(
    message,
    "Today's schedule:\n" +
      "• <https://outlook.office.com/a|@jdoe - Sick - Jul 20 - John Doe>\n" +
      "• <https://outlook.office.com/b|@asmith - Vacation - Jul 20 - Alex Smith>",
  );
});

Deno.test("buildDigestMessage says so when there are no events", () => {
  assertEquals(buildDigestMessage([]), "Nothing on the schedule today.");
});

Deno.test("post_daily_schedule_digest posts a digest of today's Graph events to the given channel", async () => {
  let postedMessage: { channel: string; text: string } | undefined;
  let calendarViewUrl: string | undefined;

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

      if (request.url.includes("/calendarView")) {
        calendarViewUrl = request.url;
        return new Response(
          JSON.stringify({
            value: [
              {
                subject: "@jdoe - Sick - Jul 20 - John Doe",
                webLink: "https://outlook.office.com/calendar/item/abc",
              },
            ],
          }),
          { status: 200 },
        );
      }

      if (request.url.includes("slack.com/api/chat.postMessage")) {
        const params = await request.formData();
        postedMessage = {
          channel: params.get("channel") as string,
          text: params.get("text") as string,
        };
        return new Response(JSON.stringify({ ok: true }), { status: 200 });
      }

      throw new Error(`Unexpected fetch: ${request.url}`);
    },
  );

  const { outputs, error } = await PostDailyScheduleDigest(
    createContext({ inputs: { channel_id: "C0SCHED1" }, env: FAKE_ENV }),
  );

  assertEquals(error, undefined);
  assertExists(calendarViewUrl);
  assertStringIncludes(calendarViewUrl!, "calendarView");
  assertExists(postedMessage);
  assertEquals(postedMessage?.channel, "C0SCHED1");
  assertStringIncludes(postedMessage?.text ?? "", "@jdoe - Sick - Jul 20");
  assertStringIncludes(
    postedMessage?.text ?? "",
    "https://outlook.office.com/calendar/item/abc",
  );
  assertEquals(outputs?.event_count, 1);
});

Deno.test("post_daily_schedule_digest posts a 'nothing scheduled' message when the calendar is empty", async () => {
  let postedMessage: { channel: string; text: string } | undefined;

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

      if (request.url.includes("/calendarView")) {
        return new Response(JSON.stringify({ value: [] }), { status: 200 });
      }

      if (request.url.includes("slack.com/api/chat.postMessage")) {
        const params = await request.formData();
        postedMessage = {
          channel: params.get("channel") as string,
          text: params.get("text") as string,
        };
        return new Response(JSON.stringify({ ok: true }), { status: 200 });
      }

      throw new Error(`Unexpected fetch: ${request.url}`);
    },
  );

  const { outputs, error } = await PostDailyScheduleDigest(
    createContext({ inputs: { channel_id: "C0SCHED1" }, env: FAKE_ENV }),
  );

  assertEquals(error, undefined);
  assertEquals(postedMessage?.text, "Nothing on the schedule today.");
  assertEquals(outputs?.event_count, 0);
});

Deno.test("post_daily_schedule_digest fails fast when MS_SHARED_MAILBOX is missing", async () => {
  using _stubFetch = stub(
    globalThis,
    "fetch",
    () => {
      throw new Error("fetch should never be called without MS_SHARED_MAILBOX");
    },
  );

  const { MS_SHARED_MAILBOX: _omit, ...envWithoutMailbox } = FAKE_ENV;
  const { outputs, error } = await PostDailyScheduleDigest(
    createContext({
      inputs: { channel_id: "C0SCHED1" },
      env: envWithoutMailbox,
    }),
  );

  assertExists(error);
  assertStringIncludes(error, "MS_SHARED_MAILBOX");
  assertEquals(outputs, undefined);
});

Deno.test("post_daily_schedule_digest surfaces a clean error when reading the calendar fails", async () => {
  using _stubFetch = stub(
    globalThis,
    "fetch",
    (url: string | URL | Request, options?: RequestInit) => {
      const request = url instanceof Request ? url : new Request(url, options);

      if (request.url.includes("login.microsoftonline.com")) {
        return Promise.resolve(
          new Response(
            JSON.stringify({ access_token: "fake-access-token" }),
            { status: 200 },
          ),
        );
      }

      return Promise.resolve(
        new Response("mailbox not found", { status: 404 }),
      );
    },
  );

  const { outputs, error } = await PostDailyScheduleDigest(
    createContext({ inputs: { channel_id: "C0SCHED1" }, env: FAKE_ENV }),
  );

  assertExists(error);
  assertStringIncludes(error, "Failed to read the shared mailbox's calendar");
  assertEquals(outputs, undefined);
});
