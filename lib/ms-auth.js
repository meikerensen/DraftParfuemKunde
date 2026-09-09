// Dependency-free Microsoft Graph client-credentials auth (app-only, no
// user sign-in) - mirrors lib/google-auth.js's role for the Google Sheets
// path. Requires an Azure AD app registration with the Files.Read.All
// application permission (admin-consented) - see project docs for setup.
async function getAccessToken(tenantId, clientId, clientSecret) {
  const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "client_credentials",
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
    }),
  });
  if (!res.ok) {
    throw new Error(`Microsoft OAuth token request failed: ${res.status} ${res.statusText} - ${await res.text()}`);
  }
  const data = await res.json();
  return data.access_token;
}

module.exports = { getAccessToken };
