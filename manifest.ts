import { Manifest } from "deno-slack-sdk/mod.ts";
import { CreateCalendarEventDefinition } from "./functions/create_calendar_event.ts";
import { PostDailyScheduleDigestDefinition } from "./functions/post_daily_schedule_digest.ts";

/**
 * The app manifest contains the app's configuration. This
 * file defines attributes like app name and description.
 * https://api.slack.com/automation/manifest
 */
export default Manifest({
  name: "Emtech Scheduling",
  description: "A template for building Slack apps with Deno",
  icon: "assets/default_new_app_icon.png",
  functions: [CreateCalendarEventDefinition, PostDailyScheduleDigestDefinition],
  outgoingDomains: ["login.microsoftonline.com", "graph.microsoft.com"],
  botScopes: [
    "commands",
    "chat:write",
    "chat:write.public",
    "users:read",
    "users:read.email",
  ],
});
