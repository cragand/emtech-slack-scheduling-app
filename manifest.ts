import { Manifest } from "deno-slack-sdk/mod.ts";
import SampleWorkflow from "./workflows/sample_workflow.ts";
import SampleObjectDatastore from "./datastores/sample_datastore.ts";
import { CreateCalendarEventDefinition } from "./functions/create_calendar_event.ts";

/**
 * The app manifest contains the app's configuration. This
 * file defines attributes like app name and description.
 * https://api.slack.com/automation/manifest
 */
export default Manifest({
  name: "Emtech Scheduling",
  description: "A template for building Slack apps with Deno",
  icon: "assets/default_new_app_icon.png",
  functions: [CreateCalendarEventDefinition],
  workflows: [SampleWorkflow],
  outgoingDomains: ["login.microsoftonline.com", "graph.microsoft.com"],
  datastores: [SampleObjectDatastore],
  botScopes: [
    "commands",
    "chat:write",
    "chat:write.public",
    "datastore:read",
    "datastore:write",
    "users:read",
    "users:read.email",
  ],
});
