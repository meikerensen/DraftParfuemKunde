// One-off verification script - NOT part of the regular generate.js
// pipeline yet. Run this once the Azure AD app registration is done and
// MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET are set as env vars, to
// confirm the Graph reader actually works against the real file before
// wiring it into the dashboard.
//
// Usage (locally, with the 3 env vars set in your shell):
//   node scripts/test-msgraph.js
const { getAccessToken } = require("../lib/ms-auth");
const { listWorksheetNames, loadWorkbookFromShareLink } = require("../lib/msgraph-lite");

const SHARE_URL = "https://digitaktco-my.sharepoint.com/:x:/g/personal/zh_digitakt_co/IQBKluJ8RMXdRKmybCnHIALtAVOSdcbdI_Kq3nZ7But7-Rs";

async function main() {
  const tenantId = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    console.error("Missing MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET environment variables.");
    process.exit(1);
  }

  console.log("Requesting access token...");
  const token = await getAccessToken(tenantId, clientId, clientSecret);
  console.log("Got token (length " + token.length + ").");

  console.log("\nListing worksheets in the workbook...");
  const sheets = await listWorksheetNames(SHARE_URL, token);
  console.log("Sheets:", sheets.join(" | "));

  console.log("\nReading 'Sonstiges' tab as a test...");
  const wb = await loadWorkbookFromShareLink(SHARE_URL, token, ["Sonstiges"]);
  const rows = wb.getRows("Sonstiges");
  console.log("Total rows:", rows.length);
  for (let i = 0; i < Math.min(6, rows.length); i++) {
    console.log("R" + rows[i].rowNum + ":", JSON.stringify(rows[i]));
  }
}

main().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
