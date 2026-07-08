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
        description: "Event start, ISO 8601 (e.g. 2026-07-10T09:00:00)",
      },
      end_date_time: {
        type: Schema.types.string,
        description: "Event end, ISO 8601 (e.g. 2026-07-10T17:00:00)",
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
export function formatTitleDate(isoDateTime: string): string {
  const date = new Date(isoDateTime);
  return date.toLocaleDateString("en-US", {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
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
      formatTitleDate(start_date_time)
    }`;

    const attendees = [
      { emailAddress: { address: submitter_email }, type: "required" },
      ...(additional_attendees ?? []).map((email) => ({
        emailAddress: { address: email },
        type: "required",
      })),
    ];

    const eventBody = {
      subject: title,
      body: { contentType: "Text", content: description ?? "" },
      start: { dateTime: start_date_time, timeZone },
      end: { dateTime: end_date_time, timeZone },
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
