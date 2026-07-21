// Shared OAuth2 client-credentials token fetch for Microsoft Graph, used by
// any function in this app that talks to Graph — extracted here once a
// second function (the daily digest) needed the exact same token, to avoid
// two copies of the same auth logic silently drifting apart.
export async function getGraphAccessToken(
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
