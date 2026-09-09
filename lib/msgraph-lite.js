// Reads an Excel workbook stored in OneDrive/SharePoint via Microsoft
// Graph, starting from a normal sharing link - no need to know the file's
// internal drive/item ID up front, Graph resolves that from the link
// itself (and this keeps working even if the link's "?e=" nonce changes
// on re-share, since that resolves to the same underlying driveItem).
// Exposes the same { listSheets(), getSheet(name), getRows(name) }
// interface as lib/xlsx-lite.js, so generate.js's row-processing logic
// (rowNum filters, hasContent, excelDateToISO etc.) works unchanged
// regardless of which backend a given tab actually came from.
//
// Dates: Graph's worksheet range API returns cell values with dates as
// Excel serial numbers - confirmed via Microsoft's own docs/Q&A, not
// assumed - the exact same "days since 1899-12-30" epoch xlsx-lite already
// uses, so excelDateToISO() needs zero changes to work with this source.
//
// Encoding a sharing URL into a Graph share-id: base64-encode the URL,
// convert to unpadded base64url (strip "=", "/"->"_", "+"->"-"), prefix
// with "u!" - this exact recipe is Microsoft's documented sharing-token
// format (graph.microsoft.com/.../shares-get), not a guessed convention.
function encodeShareUrl(url) {
  const base64 = Buffer.from(url, "utf8").toString("base64");
  return "u!" + base64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
}

async function graphGet(path, accessToken) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    throw new Error(`Graph API GET ${path} failed: ${res.status} ${res.statusText} - ${await res.text()}`);
  }
  return res.json();
}

async function resolveShareLink(shareUrl, accessToken) {
  const encoded = encodeShareUrl(shareUrl);
  const shared = await graphGet(`/shares/${encoded}/driveItem`, accessToken);
  const driveId = shared.parentReference && shared.parentReference.driveId;
  const itemId = shared.id;
  if (!driveId || !itemId) {
    throw new Error(`Could not resolve drive/item id from share link - response: ${JSON.stringify(shared)}`);
  }
  return { driveId, itemId };
}

// Lists every worksheet name in the workbook - useful for a first
// verification pass against the real file before deciding which tabs to
// actually read into the dashboard.
async function listWorksheetNames(shareUrl, accessToken) {
  const { driveId, itemId } = await resolveShareLink(shareUrl, accessToken);
  const data = await graphGet(`/drives/${driveId}/items/${itemId}/workbook/worksheets`, accessToken);
  return (data.value || []).map((s) => s.name);
}

// tabNames: array of worksheet names to fetch. Fetches each tab's full
// usedRange eagerly, then returns a synchronous-lookup object, because
// getRows() itself has to stay synchronous to match xlsx-lite's interface.
async function loadWorkbookFromShareLink(shareUrl, accessToken, tabNames) {
  const { driveId, itemId } = await resolveShareLink(shareUrl, accessToken);

  const cache = {};
  for (const name of tabNames) {
    const range = await graphGet(
      `/drives/${driveId}/items/${itemId}/workbook/worksheets/${encodeURIComponent(name)}/usedRange`,
      accessToken
    );
    // rowIndex is the 0-based sheet row where the used range starts - not
    // assumed to be 0, in case a tab's real data doesn't begin at row 1.
    const startRow = range.rowIndex || 0;
    const values = range.values || [];
    cache[name] = values.map((row, idx) => {
      const arr = row.slice();
      arr.rowNum = startRow + idx + 1;
      return arr;
    });
  }

  return {
    listSheets: () => tabNames,
    getSheet: (name) => cache[name],
    getRows: (name) => {
      if (!(name in cache)) throw new Error(`Tab "${name}" was not fetched - add it to the tabNames list.`);
      return cache[name];
    },
  };
}

module.exports = { loadWorkbookFromShareLink, listWorksheetNames, encodeShareUrl };
