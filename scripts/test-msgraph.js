// One-off verification/inspection script - NOT part of the regular
// generate.js pipeline. Run via the "Test Microsoft Graph connection"
// GitHub Actions workflow (workflow_dispatch), which injects the 3 secrets.
//
// Currently configured to inspect the 4 SharePoint tabs not yet checked for
// data quality: "3. New Listings", "Lilial", "PD-New sumbission",
// "PS-New submissions", "PD-brand approvals" - looking specifically for the
// same corruption pattern (stray numeric junk in text fields) already found
// in this file's "2. Account Violations" and "1. Account Health Score" tabs.
const { getAccessToken } = require("../lib/ms-auth");
const { loadWorkbookFromShareLink } = require("../lib/msgraph-lite");

const SHARE_URL = "https://digitaktco-my.sharepoint.com/:x:/g/personal/zh_digitakt_co/IQBKluJ8RMXdRKmybCnHIALtAVOSdcbdI_Kq3nZ7But7-Rs";

function looksNumericJunk(v) {
  const s = String(v == null ? "" : v).trim();
  return s !== "" && /^\d+(\.\d+)?$/.test(s);
}

function inspectTab(wb, name, textFieldIndexes, sampleCount) {
  const rows = wb.getRows(name);
  console.log(`\n=== "${name}" ===`);
  console.log("Total rows:", rows.length);
  console.log("First", sampleCount, "rows:");
  for (let i = 0; i < Math.min(sampleCount, rows.length); i++) {
    console.log("  R" + rows[i].rowNum + ":", JSON.stringify(rows[i]));
  }
  // Corruption screen: count how many of the specified "should be text"
  // fields across all rows are actually bare numbers instead.
  let junkCount = 0;
  const junkExamples = [];
  for (const r of rows) {
    for (const idx of textFieldIndexes) {
      const v = r[idx];
      if (looksNumericJunk(v)) {
        junkCount++;
        if (junkExamples.length < 5) junkExamples.push({ rowNum: r.rowNum, col: idx, value: v });
      }
    }
  }
  console.log(`Numeric-junk cells found in text-expected columns [${textFieldIndexes.join(",")}]:`, junkCount);
  if (junkExamples.length) console.log("Examples:", JSON.stringify(junkExamples));
}

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

  const tabNames = ["3. New Listings", "Lilial", "PD-New sumbission", "PS-New submissions", "PD-brand approvals"];
  console.log("\nFetching tabs:", tabNames.join(" | "));
  const wb = await loadWorkbookFromShareLink(SHARE_URL, token, tabNames);

  // Column indexes chosen per tab's expected schema (0-indexed), based on
  // the header rows/sample content already seen for these sheet names.
  inspectTab(wb, "3. New Listings", [3, 4, 6], 8); // Country, Account, Product-ish columns
  inspectTab(wb, "Lilial", [1], 8); // Account column
  inspectTab(wb, "PD-New sumbission", [0], 8); // Asin column
  inspectTab(wb, "PS-New submissions", [0], 8); // Asin column
  inspectTab(wb, "PD-brand approvals", [0], 12); // Brand column
}

main().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
