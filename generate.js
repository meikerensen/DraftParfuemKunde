// Parfum Dashboard - generator script
// Reads the raw Excel/CSV exports and emits a static dist/index.html.
// Re-run with `node generate.js` whenever new exports land in data/.
const path = require("path");
const fs = require("fs");
const { loadWorkbook } = require("./lib/xlsx-lite");

const DATA_DIR = path.join(__dirname, "data");
const OUT_DIR = path.join(__dirname, "dist");

// ---------------------------------------------------------------- helpers --
function cleanAsin(v) {
  const s = String(v == null ? "" : v).trim().toUpperCase();
  return /^B0[A-Z0-9]{8}$/.test(s) ? s : null;
}
function scanForAsin(rowArr) {
  for (const v of rowArr) {
    const a = cleanAsin(v);
    if (a) return a;
  }
  return null;
}
function normalizeAccount(v) {
  const s = String(v == null ? "" : v).trim().toLowerCase();
  if (!s) return null;
  if (s.includes("direct")) return "Parfum Direct";
  if (s.includes("store")) return "Parfum Store";
  if (s === "pd") return "Parfum Direct";
  if (s === "ps") return "Parfum Store";
  return null;
}
function scanForAccount(rowArr) {
  for (const v of rowArr) {
    const a = normalizeAccount(v);
    if (a) return a;
  }
  return null;
}
function safeNumber(v) {
  if (v === undefined || v === null || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
function excelDateToISO(v) {
  const n = Number(v);
  if (!Number.isFinite(n) || n < 20000 || n > 60000) return null; // sane 1954-2064 window
  const epoch = Date.UTC(1899, 11, 30);
  return new Date(epoch + n * 86400000).toISOString().slice(0, 10);
}
// Parses either an Excel serial-date number (xlsx source) or German-locale
// date text (CSV source). Found necessary because Google's xlsx export
// silently blanks cells containing short-format date text like "14.8" (day.
// month, no year) - confirmed by comparing an xlsx export against a CSV
// export of the same tab taken minutes apart: the CSV had the text, the
// xlsx cell was empty. Full "D.M.YYYY"/"D.M.YY" text is also handled for
// completeness, though those survive the xlsx export fine.
//
// Short dates have no year at all - most rows in the New Listings tab use
// this format, and the last row with an actual year anywhere is from
// January 2025. A naive "assume current year" is provably wrong: it
// produced entries like "01.10" -> 2026-10-01, a date in the future
// relative to a snapshot taken 2026-09-01, on a row already marked "done" -
// a completed task can't be dated in the future. Since these rows are
// entered in roughly chronological order and the tracker spans a year
// boundary, the fix is to prefer the current year but roll back to the
// previous year whenever the current-year reading would land after the
// snapshot date. This can't perfectly recover the true year for old
// entries buried deep in 2024/2025 history (there's no year marker left to
// check against for those), but it guarantees no date is ever shown in the
// future, and it's correct for anything near the actual year boundary,
// which is what matters for the "recent" figures.
function parseFlexibleDate(v, snapshotISO) {
  const serial = excelDateToISO(v);
  if (serial) return serial;
  const s = String(v == null ? "" : v).trim();
  if (!s) return null;
  let m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2})$/);
  if (m) return `20${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  m = s.match(/^(\d{1,2})\.(\d{1,2})\.?$/);
  if (m) {
    const day = m[1].padStart(2, "0"), month = m[2].padStart(2, "0");
    const thisYear = new Date().getUTCFullYear();
    let iso = `${thisYear}-${month}-${day}`;
    if (snapshotISO && iso > snapshotISO) iso = `${thisYear - 1}-${month}-${day}`;
    return iso;
  }
  return null;
}
// Minimal RFC4180-ish CSV parser (semicolon-delimited, Google Sheets CSV
// export style) - handles quoted fields with "" escaping and embedded
// semicolons/newlines. Returns the same row-array-with-.rowNum shape as
// xlsx-lite's getRows(), so downstream code doesn't care which format a
// given tab actually came from.
function parseCsvSemicolon(text) {
  const rows = [];
  let row = [], field = "", inQuotes = false, i = 0;
  const pushField = () => { row.push(field); field = ""; };
  const pushRow = () => { pushField(); rows.push(row); row = []; };
  while (i < text.length) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
        inQuotes = false; i++; continue;
      }
      field += c; i++; continue;
    }
    if (c === '"') { inQuotes = true; i++; continue; }
    if (c === ";") { pushField(); i++; continue; }
    if (c === "\r") { i++; continue; }
    if (c === "\n") { pushRow(); i++; continue; }
    field += c; i++;
  }
  if (field !== "" || row.length) pushRow();
  return rows.map((cells, idx) => {
    const arr = cells.slice();
    arr.rowNum = idx + 1;
    return arr;
  });
}
function hasContent(rowArr) {
  return rowArr.some((v) => String(v == null ? "" : v).trim() !== "");
}
// A "status" cell that is purely numeric (e.g. "23", "3673") is not a real
// submission status - it's the same merged-cell column drift seen elsewhere
// in these sheets, just bleeding a different neighbour value into this field.
function looksNumericJunk(v) {
  const s = String(v == null ? "" : v).trim();
  return s !== "" && /^\d+(\.\d+)?$/.test(s);
}
function fmtEUR(n) {
  return n.toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €";
}
function fmtInt(n) {
  return n.toLocaleString("de-DE");
}
function esc(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function countBy(arr, keyFn) {
  const m = {};
  for (const item of arr) {
    const k = keyFn(item);
    const key = k === null || k === undefined || k === "" ? "(leer)" : k;
    m[key] = (m[key] || 0) + 1;
  }
  return m;
}
function sortedEntries(obj) {
  return Object.entries(obj).sort((a, b) => b[1] - a[1]);
}
// Status labels get a semantic colour based on what stage of a workflow they
// represent - separate from the per-store brand accent used elsewhere, so a
// "blocked" chip always reads as a problem regardless of which account it's
// on. Free-text categories (reasons, flags, types) don't get this treatment,
// only genuine status/progress fields.
function statusToken(label) {
  const s = String(label || "").toLowerCase();
  if (/done|approved|erledigt/.test(s)) return "good";
  if (/progress|submitted|review/.test(s)) return "warn";
  if (/blocked|open|offen/.test(s)) return "critical";
  return "neutral";
}

// SNAPSHOT_NOW is frozen at generation time and embedded into the page, both
// for the server-rendered initial numbers AND the client-side day-window
// selector. If the browser recomputed "now" at view time instead, the same
// page would show different numbers depending on when you happened to open
// it - a "weekly snapshot" needs one fixed reference date, not a live clock.
// Declared early (before any date parsing happens) so parseFlexibleDate can
// use it to reject impossible future dates.
const SNAPSHOT_NOW = new Date();
const SNAPSHOT_ISO = SNAPSHOT_NOW.toISOString().slice(0, 10);

const notes = []; // data-quality notes collected while building, rendered at the bottom (internal only - not shown on the page)
function note(de, en) { notes.push({ de, en }); }
// Inline bilingual label: German on the main line, English rendered smaller
// underneath via CSS (.en). Used for every static UI string - headings,
// legends, table headers, tags - never for raw data values (product names,
// reasons, comments), which stay exactly as the source sheet has them.
function bi(de, en) {
  return `${esc(de)}<span class="en">${esc(en)}</span>`;
}

// ============================================================================
// 1) UEBERSICHTSTABELLE
// ============================================================================
const ueb = loadWorkbook(path.join(DATA_DIR, "uebersicht.xlsx"));

// --- 3. New Listings ---------------------------------------------------
// Prefer a CSV export of this one tab over the xlsx workbook's copy, if
// present: the xlsx export was found to silently blank cells with
// short-format date text (see parseFlexibleDate above), which the CSV
// export preserves correctly. Same column layout either way, so only the
// row source and date parser differ.
const nlCsvPath = path.join(DATA_DIR, "new-listings.csv");
const nlUsingCsv = fs.existsSync(nlCsvPath);
const nlSourceRows = nlUsingCsv
  ? parseCsvSemicolon(fs.readFileSync(nlCsvPath, "utf8"))
  : ueb.getRows("3. New Listings");
const nlRows = nlSourceRows.filter((r) => r.rowNum > 3 && hasContent(r));
const newListings = nlRows.map((r) => ({
  rowNum: r.rowNum,
  date: parseFlexibleDate(r[2], SNAPSHOT_ISO),
  country: r[3] || null,
  account: normalizeAccount(r[4]) || scanForAccount(r), // col4 is the labeled Account col; some rows have it shifted by merged cells
  ean: r[5] || null,
  asin: cleanAsin(r[6]) || cleanAsin(r[14]) || scanForAsin(r), // sheet has two ASIN columns (F + O) depending on the row
  product: r[7] || null,
  status: (r[8] || "").trim() || null,
}));
const nlByAccount = countBy(newListings, (x) => x.account || "Nicht zuordenbar (Quelle unvollständig)");
const nlByStatus = countBy(newListings, (x) => x.status || "(kein Status)");
const nlDates = newListings.map((x) => x.date).filter(Boolean).sort();

// The "Date" column does NOT track when a listing was newly created - it
// clusters into a handful of distinct values with massive spikes on single
// dates, and rows in the most recent batches rarely have a product name
// filled in. That's the signature of periodic batch reviews of the EXISTING
// backlog (comments like "still under processing, check later"), not of
// individual new listings being created day by day. So this is deliberately
// labelled "bearbeitete Listings" (processed/reviewed), never "New
// Listings" + a day count, which would imply recent creation that the data
// doesn't support.
const NL_RECENT_DAYS = 30;
const nlCutoffISO = new Date(SNAPSHOT_NOW.getTime() - NL_RECENT_DAYS * 86400000).toISOString().slice(0, 10);
const nlLatestDateISO = nlDates[nlDates.length - 1] || null;
const nlDaysSinceLatest = nlLatestDateISO ? Math.floor((SNAPSHOT_NOW - new Date(nlLatestDateISO)) / 86400000) : null;
const newListingsRecent = newListings.filter((x) => x.date && x.date >= nlCutoffISO);
const nlRecentByAccount = countBy(newListingsRecent, (x) => x.account || "Nicht zuordenbar (Quelle unvollständig)");
const nlByStatusRecent = countBy(newListingsRecent, (x) => x.status || "(kein Status)");
const totalNewListingsRecent = newListingsRecent.length;

if (nlUsingCsv) {
  note(
    `„Bearbeitete Listings" liest ab jetzt aus einer separaten CSV-Datei dieses einen Tabs statt aus der xlsx-Gesamtdatei — Grund: der xlsx-Export von Google Sheets leert Zellen mit kurzem Datumstext ohne Jahr (z. B. „14.8") kommentarlos, während der CSV-Export denselben Text korrekt enthält. Verifiziert durch Vergleich beider Exportformate für dieselbe Tabelle, Minuten auseinander erstellt. ` +
    `Kurze Datumsangaben ohne Jahr nehmen das aktuelle Jahr an, außer das würde ein Datum NACH dem Snapshot-Stand ergeben (z. B. wurde „01.10" fälschlich als Zukunftsdatum erkannt), dann wird ein Jahr zurückgerechnet. Für vereinzelte alte Einträge tief in der Historie kann das Jahr dadurch trotzdem ungenau sein — für die „letzte X Tage"-Werte oben ist das jedoch korrekt, da es genau den Jahreswechsel behandelt.`,
    `"Processed Listings" now reads from a separate CSV export of this one tab instead of the xlsx workbook — reason: Google Sheets' xlsx export silently blanks cells with short date text lacking a year (e.g. "14.8"), while the CSV export preserves that text correctly. Verified by comparing both export formats of the same tab, taken minutes apart. ` +
    `Short dates with no year are assumed to be the current year, unless that would land the date AFTER the snapshot date (e.g. "01.10" was initially misread as a future date), in which case it falls back one year. This can still be imprecise for scattered old entries deep in the history, but it's correct for the "last X days" figures above, since it specifically handles the year boundary.`
  );
}
note(
  `Die Kachel heißt „Bearbeitete Listings" statt „New Listings", weil die Datumsspalte im Quell-Tab kein Erstellungsdatum ist, sondern das Datum des letzten Bearbeitungs-/Prüf-Durchgangs — ` +
  `die Datumswerte häufen sich auf wenige Einzeltage (Batch-Reviews eines bestehenden Rückstands), und die zuletzt bearbeiteten Zeilen haben meist keinen Produktnamen. ` +
  `Die Kommentare („still under processing, check later") beschreiben Nachverfolgung, keine Neuanlage.`,
  `The tile is called "Processed Listings" instead of "New Listings" because the date column in the source sheet is not a creation date, but the date of the last processing/review pass — ` +
  `the date values cluster on a handful of single days (batch reviews of an existing backlog), and the most recently touched rows usually have no product name. ` +
  `The comments ("still under processing, check later") describe follow-up, not new creation.`
);
note(
  `„Bearbeitete Listings" zeigt die letzten ${NL_RECENT_DAYS} Tage (ab ${nlCutoffISO}): ${totalNewListingsRecent} Einträge. ` +
  `Lifetime-Gesamtzahl (alle Zeilen seit Beginn der Tabelle, unverändert verfügbar): ${newListings.length}, davon ${nlRows.length - nlDates.length} ohne verwertbares Datum (fließen nicht in den 30-Tage-Wert ein). ` +
  (nlDaysSinceLatest !== null
    ? `Der Tracker wurde zuletzt am ${nlLatestDateISO} befüllt (vor ${nlDaysSinceLatest} Tagen) — ein niedriger Wert kann auch schlicht heißen, dass die Tabelle seither nicht bearbeitet wurde.`
    : ""),
  `"Processed Listings" shows the last ${NL_RECENT_DAYS} days (from ${nlCutoffISO}): ${totalNewListingsRecent} entries. ` +
  `Lifetime total (all rows since the tracker began, still available): ${newListings.length}, of which ${nlRows.length - nlDates.length} have no usable date (excluded from the ${NL_RECENT_DAYS}-day figure). ` +
  (nlDaysSinceLatest !== null
    ? `The tracker was last updated on ${nlLatestDateISO} (${nlDaysSinceLatest} days ago) — a low figure can simply mean the table hasn't been touched since, not that nothing happened.`
    : "")
);
note(
  `${nlByAccount["Nicht zuordenbar (Quelle unvollständig)"] || 0} Zeilen in „New Listings" (Lifetime-Gesamtzahl) konnten keinem Account zugeordnet werden ` +
  `(verbundene/verschobene Zellen in der Quelltabelle liefern dort Zahlen-Müll statt „parfum-direct"/„Parfum Store" — Quelldatenproblem, kein Skript-Bug).`,
  `${nlByAccount["Nicht zuordenbar (Quelle unvollständig)"] || 0} rows in "New Listings" (lifetime total) could not be matched to an account ` +
  `(merged/shifted cells in the source table return number junk there instead of "parfum-direct"/"Parfum Store" — a source-data issue, not a script bug).`
);
note(
  `Alle Zahlen auf dieser Seite sind ein eingefrorener Wochenstand („Stand: ${SNAPSHOT_ISO}") — die Tages-Fenster (7/14/30/60/90 Tage) rechnen relativ zu diesem Datum, nicht zur Uhrzeit beim Betrachten. ` +
  `Erst der nächste Lauf von generate.js aktualisiert den Stand.`,
  `All figures on this page are a frozen weekly snapshot ("as of: ${SNAPSHOT_ISO}") — the day windows (7/14/30/60/90 days) are calculated relative to this date, not to the time you're viewing the page. ` +
  `Only the next run of generate.js will update the snapshot.`
);
// Compact date+account+status payload for the client-side day-window selector
// below - only what's needed to recompute counts in the browser.
const nlFilterData = newListings.filter((x) => x.date).map((x) => [x.date, x.account || "Nicht zuordenbar (Quelle unvollständig)", x.status || "(kein Status)"]);

// --- 4. Blocked ASINs ----------------------------------------------------
function normalizeReason(v) {
  const s = (v || "").trim();
  if (/^lillial$/i.test(s)) return "Lilial"; // verified typo of "Lilial" in the source sheet
  if (s.startsWith("Sie haben keine Berechtigung, Artikel dieser Marke zu verkaufen")) {
    return `„${s}"`; // quoted Amazon policy message, not a short reason keyword like the others
  }
  return s || null;
}
// Column positions in this sheet have drifted between exports before (an
// extra column got inserted, shifting everything one to the right - caught
// by a raw-value sanity check: "Reason" briefly held full product names
// after the shift). Rather than trust fixed indices again, the header row
// is read once and columns are located by their label text.
function headerIndex(headerRowArr, labels) {
  const map = {};
  headerRowArr.forEach((cell, i) => {
    const v = String(cell || "").trim();
    if (labels.includes(v) && !(v in map)) map[v] = i;
  });
  return map;
}
const blockedHeaderRow = ueb.getRows("4. Blocked ASINs").find((r) => r.rowNum === 3);
const blockedCol = headerIndex(blockedHeaderRow, ["Product", "Reason", "Stock", "Status"]);
const blockedRows = ueb.getRows("4. Blocked ASINs").filter((r) => r.rowNum > 3 && hasContent(r));
const blockedAsins = blockedRows.map((r) => ({
  rowNum: r.rowNum,
  account: scanForAccount(r),
  asin: scanForAsin(r),
  product: r[blockedCol["Product"]] || null,
  reason: normalizeReason(r[blockedCol["Reason"]]),
  stock: safeNumber(r[blockedCol["Stock"]]),
  status: (r[blockedCol["Status"]] || "").trim() || null,
}));
const blockedByStatus = countBy(blockedAsins, (x) => x.status || "(kein Status)");
const blockedByReason = countBy(blockedAsins, (x) => x.reason || "(kein Grund)");
note(
  `In „Blocked ASINs" wurden die Schreibvarianten „Lillial" und „Lilial" (${(blockedAsins.filter((x) => x.reason === "Lilial").length)} Zeilen zusammen) als ein Grund gezählt — offensichtlicher Tippfehler in der Quelle, keine zwei unterschiedlichen Gründe.`,
  `In "Blocked ASINs", the spelling variants "Lillial" and "Lilial" (${(blockedAsins.filter((x) => x.reason === "Lilial").length)} rows combined) were counted as one reason — an obvious typo in the source, not two different reasons.`
);
note(
  `In diesem Export war in „Blocked ASINs" eine zusätzliche Spalte eingefügt, wodurch Product/Reason/Stock/Status um eine Position verschoben waren (Reason zeigte kurzzeitig Produktnamen statt Gründen). ` +
  `Behoben durch Spalten-Erkennung über die Kopfzeile statt feste Positionen — bleibt auch bei künftigen Spaltenverschiebungen in dieser Tabelle korrekt.`,
  `In this export, an extra column had been inserted in "Blocked ASINs", shifting Product/Reason/Stock/Status one position over (Reason briefly showed product names instead of reasons). ` +
  `Fixed via header-based column detection instead of fixed positions — stays correct even if this table's columns shift again in future exports.`
);

// --- 2. Account Violations -------------------------------------------------
const violationRows = ueb.getRows("2. Account Violations").filter((r) => r.rowNum > 3 && hasContent(r));
const violations = violationRows.map((r) => ({
  rowNum: r.rowNum,
  date: excelDateToISO(r[2]),
  country: r[3] || null,
  account: normalizeAccount(r[4]) || scanForAccount(r),
  type: r[5] || null,
  asin: cleanAsin(r[6]) || scanForAsin(r),
  product: r[7] || null,
  status: (r[9] || "").trim() || null,
}));
const violationsByStatus = countBy(violations, (x) => x.status || "(kein Status)");
const violationsByType = countBy(violations, (x) => x.type || "(kein Typ)");

// --- GPSR complaints (canonical source; "Lilial" and "GPSR-Graph" tabs in ---
// --- this workbook are byte-identical duplicates of "GPSR" - verified ------
// --- so only "GPSR" is read to avoid tripling ------------------------------
const gpsrRawAll = ueb.getRows("GPSR");
const gpsrCases = gpsrRawAll
  .filter((r) => r.rowNum > 2 && (r[3] || "").toString().trim() !== "") // real complaint = has an Order Nr.
  .map((r) => ({
    account: normalizeAccount(r[1]) || scanForAccount(r),
    marketplace: r[2] || null,
    orderNr: r[3] || null,
    date: excelDateToISO(r[4]),
    reason: r[5] || null,
    automaticRemoval: (r[6] || "").trim() || null,
    amazonCase: r[7] || null,
    flag: (r[8] || "").trim() || null,
  }));
const gpsrTemplateRowCount = gpsrRawAll.filter((r) => r.rowNum > 2).length - gpsrCases.length;
note(
  `„GPSR"/„Lilial"/„GPSR-Graph" sind in der Übersichtstabelle drei Tabs mit identischem Inhalt — ` +
  `dieses Skript liest nur „GPSR" einmal, um Drei-fach-Zählung zu vermeiden.`,
  `"GPSR"/"Lilial"/"GPSR-Graph" are three tabs with identical content in the overview workbook — ` +
  `this script reads only "GPSR" once, to avoid triple-counting.`
);
note(
  `Von ${gpsrRawAll.filter((r) => r.rowNum > 2).length} physischen Zeilen im GPSR-Tab enthalten nur ${gpsrCases.length} tatsächlich einen Fall (Order Nr. gefüllt); ` +
  `${gpsrTemplateRowCount} sind leere Vorlagenzeilen und zählen hier bewusst NICHT als „offener Fall".`,
  `Of ${gpsrRawAll.filter((r) => r.rowNum > 2).length} physical rows in the GPSR tab, only ${gpsrCases.length} actually hold a case (Order Nr. filled in); ` +
  `${gpsrTemplateRowCount} are empty template rows and are deliberately NOT counted here as an "open case".`
);
const gpsrByFlag = countBy(gpsrCases, (x) => x.flag || "(kein Flag)");

// Market coverage for the GPSR case log: restricted to DE per client request
// (01.09.) - note this means the table below no longer sums to the "GPSR
// cases (real order rows)" KPI card total further up (which still counts
// IT/ES cases too, since that total isn't scoped to this table).
const GPSR_STANDARD_MARKETS = ["DE"];
const gpsrCasesByMarket = countBy(gpsrCases, (x) => x.marketplace || "(kein Markt)");
const gpsrMarketCoverage = GPSR_STANDARD_MARKETS.map((mkt) => ({
  market: mkt,
  cases: gpsrCasesByMarket[mkt] || 0,
  tracked: (gpsrCasesByMarket[mkt] || 0) > 0,
}));
note(
  `„GPSR-Fälle nach Markt" zeigt seit dem 01.09. auf Kundenwunsch nur noch DE — IT (4 Fälle) und ES (9 Fälle) haben in der Quelle reale erfasste Fälle, werden hier aber bewusst ausgeblendet. ` +
  `Die KPI-Kachel „GPSR-Fälle (echte Order-Zeilen)" oben zählt weiterhin alle Märkte zusammen (26) — die beiden Zahlen stimmen deshalb absichtlich nicht mehr überein.`,
  `"GPSR cases by market" has shown DE only since 01.09. per client request — IT (4 cases) and ES (9 cases) have real tracked cases in the source but are deliberately hidden here. ` +
  `The "GPSR cases (real order rows)" KPI card above still counts all markets combined (26) — the two numbers are intentionally inconsistent as a result.`
);

// ============================================================================
// 2) GPSR ISSUE WORKBOOK (Priority submissions + New submissions)
// ============================================================================
const gi = loadWorkbook(path.join(DATA_DIR, "gpsr-issue.xlsx"));

// Priority sheets share a schema; "PD - Priority 1" has one extra leading
// column vs the other five, shifting Revenue from col5 to col6 (verified).
const PRIORITY_SHEETS = [
  { name: "PD - Priority 1", account: "Parfum Direct", priority: 1, revCol: 6, colOffset: 1 },
  { name: "PD - Priority 2", account: "Parfum Direct", priority: 2, revCol: 5, colOffset: 0 },
  { name: "PD - Priority 3", account: "Parfum Direct", priority: 3, revCol: 5, colOffset: 0 },
  { name: "PS - Priority 1", account: "Parfum Store", priority: 1, revCol: 5, colOffset: 0 },
  { name: "PS - Priority 2", account: "Parfum Store", priority: 2, revCol: 5, colOffset: 0 },
  { name: "PS - Priority 3", account: "Parfum Store", priority: 3, revCol: 5, colOffset: 0 },
];
let gpsrPriority = [];
let priorityStatusJunkCount = 0;
for (const spec of PRIORITY_SHEETS) {
  const off = spec.colOffset;
  const rows = gi.getRows(spec.name).filter((r) => r.rowNum > 2 && hasContent(r));
  for (const r of rows) {
    let submissionStatus = (r[9 + off] || "").trim();
    if (submissionStatus === "Submission status") continue; // stray embedded header row
    if (looksNumericJunk(submissionStatus)) { priorityStatusJunkCount++; submissionStatus = ""; }
    gpsrPriority.push({
      source: spec.name,
      account: spec.account,
      priority: spec.priority,
      sku: r[0] || null,
      asin: cleanAsin(r[1]) || scanForAsin(r),
      productName: spec.name === "PD - Priority 1" ? null : r[2] || null, // P1's col2 holds junk numbers, not a name (verified)
      revenueRaw: r[spec.revCol],
      revenue12m: safeNumber(r[spec.revCol]),
      stock: safeNumber(r[4 + off]),
      submissionStatus: submissionStatus || null,
      productStatus: (r[11 + off] || "").trim() || null,
    });
  }
}
// revenue12m===null covers two different situations: a blank cell (no
// revenue reported, totally normal) and non-numeric junk like "#N/A" (an
// actual data problem). Keep those apart so the note below stays accurate.
function revenueStats(items) {
  let sum = 0, numericCount = 0, blankCount = 0, junkCount = 0;
  for (const x of items) {
    if (x.revenue12m !== null) { sum += x.revenue12m; numericCount++; continue; }
    const raw = x.revenueRaw;
    if (raw === undefined || raw === null || String(raw).trim() === "") blankCount++;
    else junkCount++;
  }
  return { sum, numericCount, blankCount, junkCount };
}
const revSummary = revenueStats(gpsrPriority);
note(
  `Revenue-Summe wird als echte Zahlensumme berechnet, nicht als Text verkettet. ` +
  `${revSummary.junkCount} Zeile(n) hatten einen nicht-numerischen Revenue-Wert (z. B. "#N/A") und wurden ausgeschlossen; ${revSummary.blankCount} Zeilen hatten schlicht keinen Revenue-Wert eingetragen (normal, kein Fehler).`,
  `The revenue total is calculated as a true numeric sum, not concatenated as text. ` +
  `${revSummary.junkCount} row(s) had a non-numeric revenue value (e.g. "#N/A") and were excluded; ${revSummary.blankCount} rows simply had no revenue value entered (normal, not an error).`
);
const revenueByAccount = {};
for (const acc of ["Parfum Direct", "Parfum Store"]) {
  revenueByAccount[acc] = revenueStats(gpsrPriority.filter((x) => x.account === acc));
}
const productsWithRevenueByAccount = countBy(
  gpsrPriority.filter((x) => x.revenue12m !== null && x.revenue12m > 0),
  (x) => x.account
);
const submissionStatusCounts = countBy(gpsrPriority, (x) => x.submissionStatus || "(kein Status)");
if (priorityStatusJunkCount > 0) {
  note(
    `${priorityStatusJunkCount} Zeilen in den GPSR-Priority-Sheets hatten eine reine Zahl statt eines echten Status in der Submission-Status-Spalte ` +
    `(dieselbe Art von verschobener Zelle wie bei „New Listings") — diese werden als „(kein Status)" statt als Fantasie-Statuswert gezählt.`,
    `${priorityStatusJunkCount} rows in the GPSR priority sheets had a plain number instead of a real status in the submission-status column ` +
    `(the same kind of shifted cell as in "New Listings") — these are counted as "(no status)" instead of a made-up status value.`
  );
}

// New submissions sheets (different schema: Asin / Detail / Submission date / status)
const NEW_SUB_SHEETS = [
  { name: "PD-New sumbission", account: "Parfum Direct" },
  { name: "PS-New submissions", account: "Parfum Store" },
];
let gpsrNewSubmissions = [];
for (const spec of NEW_SUB_SHEETS) {
  const rows = gi.getRows(spec.name).filter((r) => r.rowNum > 2 && hasContent(r));
  for (const r of rows) {
    let status = (r[4] || "").trim();
    if (status === "Submission status") continue; // stray embedded header row (found in PS-New submissions)
    if (looksNumericJunk(status)) status = "";
    gpsrNewSubmissions.push({
      source: spec.name,
      account: spec.account,
      asin: cleanAsin(r[0]) || scanForAsin(r),
      detail: r[1] || null,
      submissionDate: r[3] || null,
      submissionStatus: status || null,
    });
  }
}

// ============================================================================
// 3) GPSR GRAPH WORKBOOK — weekly trend (only the confidently-parseable block)
// ============================================================================
const gg = loadWorkbook(path.join(DATA_DIR, "gpsr-graph.xlsx"));
const weeklyRaw = gg.getRows("Numbers per week").filter((r) => r.rowNum > 2 && (r[0] || "").trim());
const weeklyTrend = weeklyRaw
  .map((r) => ({
    period: r[0].trim(),
    pdApproved: safeNumber(r[2]),
    pdSubmitted: safeNumber(r[3]),
    pdNew: safeNumber(r[4]),
    psBeginning: safeNumber(r[7]),
    psApproved: safeNumber(r[8]),
    psSubmitted: safeNumber(r[9]),
    psNew: safeNumber(r[10]),
  }))
  .filter((w) => w.pdApproved !== null || w.pdSubmitted !== null || w.psApproved !== null || w.psSubmitted !== null);
note(
  `„Numbers per week" enthält rechts von der Wochentabelle weitere, uneindeutig ausgerichtete Mini-Tabellen (u. a. tägliche „Anzahl offener Fälle") — ` +
  `diese wurden NICHT übernommen, da die Spaltenausrichtung ohne Rückfrage beim Original-Ersteller nicht zuverlässig rekonstruierbar war. ` +
  `Nur der eindeutige Wochenblock (Approved/Submitted/New GPSR SKUs je Account) ist im Trend enthalten (${weeklyTrend.length} Wochen insgesamt).`,
  `"Numbers per week" contains further, ambiguously-aligned mini-tables to the right of the weekly table (among others a daily "number of open cases") — ` +
  `these were NOT included, since the column alignment could not be reliably reconstructed without asking the original author. ` +
  `Only the unambiguous weekly block (Approved/Submitted/New GPSR SKUs per account) is included in the trend (${weeklyTrend.length} weeks total).`
);
// Chart shows roughly the last 2 months. The sheet has no per-period exact
// date, only "YYYY.MM.<sub-period>" labels with an uneven number of
// sub-periods per month (I-V seen) - a strict calendar-month filter left as
// few as 4 points right after a month boundary, too thin for a line chart.
// Taking the last 8 rows (this data's actual weekly-ish cadence) gives a
// consistently readable ~2-month window regardless of where in the month
// the snapshot falls.
const WEEKLY_TREND_RECENT_COUNT = 8;
const weeklyTrendRecent = weeklyTrend.slice(-WEEKLY_TREND_RECENT_COUNT);
note(
  `Der Trend-Graph zeigt die letzten ${weeklyTrendRecent.length} von ${weeklyTrend.length} Wochen (≈ 2 Monate bei der Wochentaktung dieser Tabelle) — die komplette Historie steht weiterhin in der einklappbaren Detailtabelle darunter.`,
  `The trend chart shows the last ${weeklyTrendRecent.length} of ${weeklyTrend.length} weeks (≈ 2 months at this table's weekly cadence) — the full history is still available in the collapsible detail table below.`
);

// ============================================================================
// 4) PROHIBITED INGREDIENTS (Lilial/Lyral) — now delivered as XLSX, not CSV
// ============================================================================
// Source switched from a CSV export to a raw xlsx workbook with two tabs.
// "Sheet1" matches the same column layout as the previous CSV (Account,
// Title-mislabelled-"ASIN", Comment, Status, Issue, 3 extra doc-flag
// columns, Label Image Available, Case ID). "Sheet3" is a broken duplicate -
// verified: only 1 of 431 IDs there isn't already in Sheet1, and its
// Account/Title columns are numbers/blank instead of real values - so only
// Sheet1 is read, same as the single-source CSV before.
const pi = loadWorkbook(path.join(DATA_DIR, "prohibited-ingredients.xlsx"));
const piRows = pi.getRows("Sheet1").filter((r) => r.rowNum > 1); // row 1 is the header
function normalizePiStatus(v) {
  const s = String(v || "").trim().toLowerCase();
  if (!s) return "Unbekannt";
  if (s === "active") return "Aktiv";
  if (s === "inactive") return "Inaktiv";
  if (s.includes("out of stock")) return "Ausverkauft";
  return v.trim();
}
let prohibited = piRows
  .filter((r) => r.some((v) => (v || "").trim()))
  .map((r) => {
    const rawId = (r[0] || "").trim();
    const asin = cleanAsin(rawId);
    return {
      rawId,
      asin,
      eanOnly: asin ? null : rawId || null,
      account: normalizeAccount(r[1]) || r[1] || null,
      productTitle: r[2] || null, // source column is literally mislabelled "ASIN" - it's actually the title
      comment: r[3] || null,
      status: normalizePiStatus(r[4]),
      issue: r[5] || null,
      missingDocs: [r[5], r[6], r[7], r[8]].map((x) => (x || "").trim()).filter(Boolean),
      labelImageAvailable: (r[9] || "").trim() || null,
      caseId: r[10] || null,
    };
  });
// de-dupe exact repeats (verified e.g. one EAN appears 3x identically)
const seen = new Set();
let piDupes = 0;
prohibited = prohibited.filter((p) => {
  const key = [p.rawId, p.account, p.productTitle, p.status, p.caseId].join("|");
  if (seen.has(key)) { piDupes++; return false; }
  seen.add(key);
  return true;
});
note(
  `Die Inhaltsstoff-Liste kommt jetzt als xlsx-Workbook (2 Tabs) statt CSV. „Sheet1" hat dieselbe Spaltenstruktur wie die bisherige CSV-Quelle und wird gelesen; ` +
  `„Sheet3" ist eine kaputte Zweitversion (Account-Spalte enthält Zahlen statt Namen, Titel-Spalte leer) — verifiziert: nur 1 von 431 IDs dort fehlt in Sheet1, wird ignoriert.`,
  `The prohibited-ingredients list now arrives as an xlsx workbook (2 tabs) instead of a CSV. "Sheet1" has the same column layout as the previous CSV source and is read; ` +
  `"Sheet3" is a broken duplicate (Account column holds numbers instead of names, Title column is empty) — verified: only 1 of 431 IDs there is missing from Sheet1, ignored.`
);
note(
  `Liste „Verbotene Inhaltsstoffe (Lilial/Lyral)": ${prohibited.length} Einträge nach Entfernen von ${piDupes} exakten Duplikat-Zeilen. ${prohibited.filter((p) => p.eanOnly).length} Einträge haben nur eine EAN, keine ASIN.`,
  `"Prohibited ingredients (Lilial/Lyral)" list: ${prohibited.length} entries after removing ${piDupes} exact duplicate rows. ${prohibited.filter((p) => p.eanOnly).length} entries have only an EAN, no ASIN.`
);
const piByAccount = countBy(prohibited, (x) => x.account || "(kein Account)");
const piByStatus = countBy(prohibited, (x) => x.status);
const piAsins = new Set(prohibited.map((p) => p.asin).filter(Boolean));

// cross-check against other ASIN sources already loaded
const blockedAsinSet = new Set(blockedAsins.map((x) => x.asin).filter(Boolean));
const priorityAsinSet = new Set(gpsrPriority.map((x) => x.asin).filter(Boolean));
let piAlsoBlocked = 0, piAlsoInPriority = 0;
for (const a of piAsins) {
  if (blockedAsinSet.has(a)) piAlsoBlocked++;
  if (priorityAsinSet.has(a)) piAlsoInPriority++;
}
note(
  `Von ${piAsins.size} eindeutigen ASINs in der Inhaltsstoff-Liste stehen ${piAlsoBlocked} zusätzlich auch in „Blocked ASINs" und ${piAlsoInPriority} in den GPSR-Priority-Listen — ` +
  `plausibel, da Lilial/Lyral-Fälle oft auch Deaktivierungen/GPSR-Fälle auslösen.`,
  `Of ${piAsins.size} unique ASINs in the prohibited-ingredients list, ${piAlsoBlocked} also appear in "Blocked ASINs" and ${piAlsoInPriority} in the GPSR priority lists — ` +
  `plausible, since Lilial/Lyral cases often also trigger deactivations/GPSR cases.`
);

// ============================================================================
// 5) BRAND APPROVALS — Parfum Direct only (confirmed by the client), no
// date column, so it's shown as its own section rather than folded into
// an existing one, but styled/labelled as Parfum-Direct-scoped like every
// other account-specific number in this dashboard.
// ============================================================================
// Cells are free text, not a fixed enum ("Approved", "Approved, request
// approval-sell yours", "most are active, for the others not accepting
// applications", "Cant submit application", "didnt find products with this
// brandname or EANs" all appear) - categorizeBrandStatus() buckets by
// keyword, most-specific phrase first, and falls back to showing the raw
// text verbatim rather than guessing a bucket for anything unrecognized.
function categorizeBrandStatus(raw) {
  const v = (raw || "").toString().trim();
  if (!v) return { token: "neutral", label: "Keine Angabe", labelEn: "Not specified" };
  if (looksNumericJunk(v)) return { token: "neutral", label: "Dateneintrag unklar", labelEn: "Unclear data entry" };
  const s = v.toLowerCase();
  if (s.includes("cant submit") || s.includes("can't submit") || s.includes("cannot submit")) {
    return { token: "critical", label: "Nicht möglich", labelEn: "Not possible" };
  }
  if (s.includes("didnt find") || s.includes("didn't find") || s.includes("did not find")) {
    return { token: "neutral", label: "Keine Produkte gefunden", labelEn: "No products found" };
  }
  if (s.includes("not accepting applications") || (s.includes("approved") && s.includes("others"))) {
    return { token: "warn", label: "Teilweise aktiv", labelEn: "Partially active" };
  }
  if (s.includes("submitted")) {
    return { token: "warn", label: "Eingereicht", labelEn: "Submitted" };
  }
  if (s.includes("approved")) {
    return { token: "good", label: "Freigegeben", labelEn: "Approved" };
  }
  return { token: "neutral", label: v, labelEn: v };
}
const ba = loadWorkbook(path.join(DATA_DIR, "brand-approvals.xlsx"));
const baHeaderRow = ba.getRows("Sheet1").find((r) => r.rowNum === 1) || [];
// Market columns read from the header row itself (not a fixed list) - same
// header-based approach used for "4. Blocked ASINs" after that sheet's
// column-drift incident, so a future reordering here won't silently
// misalign brand names with the wrong marketplace.
const BRAND_MARKETS = baHeaderRow.slice(1).map((h) => (h || "").trim()).filter(Boolean);
const baDataRows = ba.getRows("Sheet1").filter((r) => r.rowNum > 1 && hasContent(r));
const brandApprovals = baDataRows
  .filter((r) => (r[0] || "").trim())
  .map((r) => ({
    brand: r[0].trim(),
    cells: BRAND_MARKETS.map((mkt, i) => ({ market: mkt, raw: (r[i + 1] || "").toString(), ...categorizeBrandStatus(r[i + 1]) })),
  }));
const allBaCells = brandApprovals.flatMap((b) => b.cells).filter((c) => c.raw.trim() !== "");
const baCategoryMap = new Map();
for (const c of allBaCells) {
  if (!baCategoryMap.has(c.label)) baCategoryMap.set(c.label, { label: c.label, labelEn: c.labelEn, token: c.token, count: 0 });
  baCategoryMap.get(c.label).count++;
}
const baCategoryEntries = [...baCategoryMap.values()].sort((a, b) => b.count - a.count);
const baJunkCount = allBaCells.filter((c) => looksNumericJunk(c.raw)).length;
note(
  `„Status of brand approvals" ist eine neue Quelle (Marke × Marktplatz-Freigabestatus) und gilt laut Kunde ausschließlich für Parfum Direct — ohne Datumsspalte wird sie deshalb als eigener, aber Parfum-Direct-gekennzeichneter Überblick dargestellt statt in eine bestehende Sektion eingerechnet.`,
  `"Status of brand approvals" is a new source (brand × marketplace approval status) that, per the client, applies exclusively to Parfum Direct — with no date column, it's shown as its own section but labelled/styled as Parfum-Direct-scoped rather than folded into an existing section.`
);
if (baJunkCount > 0) {
  note(
    `${baJunkCount} Zelle(n) in „Status of brand approvals" enthielten eine reine Zahl statt eines Freigabe-Status (z. B. „12", „14") — vermutlich ein Dateneingabefehler in der Quelle, wie schon bei früheren Status-Spalten in diesem Dashboard. Als „Dateneintrag unklar" statt als Fantasie-Status gezählt.`,
    `${baJunkCount} cell(s) in "Status of brand approvals" contained a plain number instead of an approval status (e.g. "12", "14") — likely a data-entry error in the source, the same pattern seen in earlier status columns in this dashboard. Counted as "Unclear data entry" instead of a made-up status.`
  );
}

// ============================================================================
// RENDER
// ============================================================================
// headers: array of [german, english] pairs
function accountTable(headers, rowsHtml) {
  return `<table><thead><tr>${headers.map(([de, en]) => `<th>${bi(de, en)}</th>`).join("")}</tr></thead><tbody>${rowsHtml}</tbody></table>`;
}
// Plain single-colour bar list - for free-text categorical breakdowns
// (reasons, flags, types) where there's no inherent "progress" to encode.
// Category labels (k) come straight from the source data, so they stay
// untranslated - only the surrounding UI chrome is bilingual.
function barHtml(entries, total) {
  return entries
    .map(([k, v]) => {
      const pct = total ? ((v / total) * 100).toFixed(1) : "0.0";
      return `<div class="bar-row"><span class="bar-label">${esc(k)}</span><div class="bar-track"><div class="bar-fill" style="width:${pct}%"></div></div><span class="bar-value">${fmtInt(v)} (${pct}%)</span></div>`;
    })
    .join("");
}
// Segmented progress bar - for genuine status/workflow fields (done/blocked/
// in progress, approved/submitted/in review). One stacked bar, colour per
// status token (good/warn/critical/neutral), plus a chip legend underneath.
function segBarHtml(entries, total) {
  if (!entries.length || !total) {
    return `<div class="seg-bar"></div><p style="color:var(--ink-faint); font-size:.8rem; margin:4px 0 0;">${bi("Keine Einträge im gewählten Zeitraum.", "No entries in the selected period.")}</p>`;
  }
  const segs = entries
    .map(([k, v]) => {
      const pct = ((v / total) * 100).toFixed(2);
      return `<div style="width:${pct}%; background:var(--${statusToken(k)});" title="${esc(k)}"></div>`;
    })
    .join("");
  const legend = entries
    .map(([k, v]) => {
      const pct = ((v / total) * 100).toFixed(1);
      return `<span><i style="background:var(--${statusToken(k)});"></i>${esc(k)} · ${fmtInt(v)} (${pct}%)</span>`;
    })
    .join("");
  return `<div class="seg-bar">${segs}</div><div class="seg-legend">${legend}</div>`;
}
// Like segBarHtml, but for entries whose colour token is already known
// (categorizeBrandStatus decides it up front from free text) rather than
// guessed from the label text - segBarHtml's own statusToken() only
// recognises English workflow words and would mis-colour German labels
// like "Freigegeben"/"Nicht möglich".
function brandStatusBarHtml(entries, total) {
  if (!entries.length || !total) {
    return `<div class="seg-bar"></div><p style="color:var(--ink-faint); font-size:.8rem; margin:4px 0 0;">${bi("Keine Einträge.", "No entries.")}</p>`;
  }
  const segs = entries
    .map((e) => {
      const pct = ((e.count / total) * 100).toFixed(2);
      return `<div style="width:${pct}%; background:var(--${e.token});" title="${esc(e.label)}"></div>`;
    })
    .join("");
  const legend = entries
    .map((e) => {
      const pct = ((e.count / total) * 100).toFixed(1);
      return `<span><i style="background:var(--${e.token});"></i>${bi(e.label, e.labelEn)} · ${fmtInt(e.count)} (${pct}%)</span>`;
    })
    .join("");
  return `<div class="seg-bar">${segs}</div><div class="seg-legend">${legend}</div>`;
}
function accentClass(account) {
  return account === "Parfum Direct" ? "acc-pd" : account === "Parfum Store" ? "acc-ps" : "";
}
// Short period label for the chart x-axis - "2026.07.III" -> "07.III"
// (year dropped, both months are visible from context, keeps labels tight).
function shortPeriod(period) {
  return period.replace(/^\d{4}\./, "");
}
// Minimal dependency-free SVG line chart - two series (e.g. Approved vs
// Submitted), sharing one y-scale, with a dot + native <title> tooltip per
// data point so exact values are available on hover without cluttering the
// chart itself with numbers. Axis titles (bilingual, matching the rest of
// the dashboard) make explicit what the two axes represent, since the
// gridline numbers alone don't say "SKUs" or "week" - a client glancing at
// the chart shouldn't have to guess.
function lineChartSvg(periods, series, axisLabels) {
  const yLabel = (axisLabels && axisLabels.y) || ["Anzahl SKUs", "Number of SKUs"];
  const xLabel = (axisLabels && axisLabels.x) || ["Woche", "Week"];
  const W = 560, H = 195, padL = 46, padR = 10, padT = 12, padB = 40;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const n = periods.length;
  const allVals = series.flatMap((s) => s.values).filter((v) => v !== null && v !== undefined);
  const maxV = Math.max(1, ...allVals);
  const x = (i) => padL + (n <= 1 ? plotW / 2 : (i / (n - 1)) * plotW);
  const y = (v) => padT + plotH - (v / maxV) * plotH;
  const gridY = [0, 0.5, 1].map((f) => {
    const yy = padT + plotH - f * plotH;
    return `<line x1="${padL}" y1="${yy.toFixed(1)}" x2="${W - padR}" y2="${yy.toFixed(1)}" stroke="var(--line)" stroke-width="1"/>` +
      `<text x="${padL - 6}" y="${(yy + 3).toFixed(1)}" font-size="8" text-anchor="end" fill="var(--ink-faint)">${fmtInt(Math.round(f * maxV))}</text>`;
  }).join("");
  const xLabelsY = padT + plotH + 10;
  const xLabels = periods
    .map((p, i) => {
      if (n > 10 && i % 2 === 1) return ""; // thin out labels when there are many weeks
      return `<text x="${x(i).toFixed(1)}" y="${xLabelsY.toFixed(1)}" font-size="8" text-anchor="middle" fill="var(--ink-faint)">${esc(shortPeriod(p))}</text>`;
    })
    .join("");
  // Y-axis title, rotated -90deg and vertically centred on the plot area.
  // Both languages sit on one rotated baseline (continuing along the same
  // line rather than stacked underneath), since stacking rotated text needs
  // a second independently-rotated <text> anyway and this is simpler to get
  // right without visual round-tripping.
  const cy = (padT + plotH / 2).toFixed(1);
  const yAxisTitle = `<text x="11" y="${cy}" font-size="8" text-anchor="middle" fill="var(--ink-faint)" transform="rotate(-90 11 ${cy})"><tspan>${esc(yLabel[0])}</tspan><tspan dx="6" font-style="italic">${esc(yLabel[1])}</tspan></text>`;
  // X-axis title, two stacked lines below the period tick labels.
  const xAxisTitleY1 = padT + plotH + 23, xAxisTitleY2 = padT + plotH + 33;
  const xAxisTitle = `<text x="${(padL + plotW / 2).toFixed(1)}" y="${xAxisTitleY1}" font-size="8" text-anchor="middle" fill="var(--ink-faint)">${esc(xLabel[0])}</text>` +
    `<text x="${(padL + plotW / 2).toFixed(1)}" y="${xAxisTitleY2}" font-size="8" font-style="italic" text-anchor="middle" fill="var(--ink-faint)">${esc(xLabel[1])}</text>`;
  const seriesSvg = series
    .map((s) => {
      const pts = s.values.map((v, i) => (v === null || v === undefined ? null : [x(i), y(v)]));
      const pathD = pts
        .map((p, i) => (p ? `${pts[i - 1] ? "L" : "M"}${p[0].toFixed(1)},${p[1].toFixed(1)}` : ""))
        .join("");
      const dots = pts
        .map((p, i) =>
          p ? `<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="2.6" fill="${s.color}"><title>${esc(periods[i])} — ${esc(s.name)}: ${fmtInt(s.values[i])}</title></circle>` : ""
        )
        .join("");
      return `<path d="${pathD}" fill="none" stroke="${s.color}" stroke-width="2"/>${dots}`;
    })
    .join("");
  return `<svg viewBox="0 0 ${W} ${H}" style="width:100%; height:auto; max-width:${W}px; display:block;">${gridY}${seriesSvg}${xLabels}${yAxisTitle}${xAxisTitle}</svg>`;
}
function chartLegend(series) {
  return `<div class="seg-legend" style="margin-top:4px;">${series
    .map((s) => `<span><i style="background:${s.color};"></i>${esc(s.name)}</span>`)
    .join("")}</div>`;
}

const totalRevenue = revSummary.sum;
const totalNewListings = newListings.length;
const totalGpsrCases = gpsrCases.length;
const totalPriorityEntries = gpsrPriority.length;
const totalProhibited = prohibited.length;

const html = `<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8" />
<title>Parfum Compliance Dashboard</title>
<style>
:root{
  --paper:#F4F2EC; --panel:#FFFFFF; --panel-2:#FBFAF6;
  --line:#DEDACF; --line-strong:#C9C3B4;
  --ink:#1E1C1A; --ink-soft:#5B564C; --ink-faint:#8A8477;
  --accent:#7A3B58; --accent-soft:#F1E4EA;
  --accent2:#1F6F6B; --accent2-soft:#E1EEEC;
  --good:#2E7A4F; --good-soft:#E3F0E7;
  --warn:#8C6208; --warn-soft:#F6ECD6;
  --critical:#A5322F; --critical-soft:#F7E3E1;
  --neutral:#8A8477; --neutral-soft:#EDEAE3;
  --shadow: 0 1px 2px rgba(30,28,26,.06), 0 6px 20px rgba(30,28,26,.05);
  --font-display: Georgia, "Iowan Old Style", "Palatino Linotype", "Book Antiqua", serif;
  --font-body: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  --font-mono: ui-monospace, "SF Mono", "Cascadia Code", Consolas, "Liberation Mono", monospace;
}
@media (prefers-color-scheme: dark){
  :root{
    --paper:#17181A; --panel:#202225; --panel-2:#1B1D1F;
    --line:#34363A; --line-strong:#3F4145;
    --ink:#ECE9E4; --ink-soft:#ACA79E; --ink-faint:#726E67;
    --accent:#D79BB6; --accent-soft:#2E2229;
    --accent2:#7FCFC9; --accent2-soft:#1C2C2A;
    --good:#6FBF8B; --good-soft:#1D2C22;
    --warn:#D9A83B; --warn-soft:#2E2716;
    --critical:#E2827F; --critical-soft:#332120;
    --neutral:#8E8A82; --neutral-soft:#26241F;
    --shadow: 0 1px 2px rgba(0,0,0,.3), 0 8px 24px rgba(0,0,0,.35);
  }
}
*{ box-sizing:border-box; }
body{ margin:0; background:var(--paper); color:var(--ink); font-family:var(--font-body); line-height:1.5; font-variant-numeric:tabular-nums; }
.wrap{ max-width:1120px; margin:0 auto; padding:32px 24px 64px; }
header.top{ display:flex; justify-content:space-between; align-items:baseline; flex-wrap:wrap; gap:10px; margin-bottom:6px; }
h1{ font-family:var(--font-display); font-weight:400; font-size:1.9rem; margin:0; }
.stand-chip{ font-family:var(--font-mono); font-size:.72rem; color:var(--ink-faint); background:var(--panel-2); border:1px solid var(--line); border-radius:3px; padding:4px 10px; white-space:nowrap; }
.sub{ color:var(--ink-soft); font-size:.9rem; margin:6px 0 30px; max-width:75ch; }
section{ margin-bottom:38px; }
h2{ font-family:var(--font-display); font-weight:400; font-size:1.2rem; margin:0 0 4px; color:var(--ink); }
h3{ font-family:var(--font-body); font-weight:650; font-size:.85rem; margin:20px 0 8px; color:var(--ink); }
.legend{ font-size:.8rem; color:var(--ink-faint); margin:0 0 14px; max-width:75ch; }
.grid{ display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:14px; }
.card{ background:var(--panel); border:1px solid var(--line); border-radius:6px; padding:16px; box-shadow:var(--shadow); }
.card strong{ display:block; font-family:var(--font-mono); font-size:1.55rem; font-weight:600; margin-bottom:6px; }
.card .lbl{ font-size:.82rem; color:var(--ink-soft); }
.card small{ display:block; color:var(--ink-faint); font-size:.74rem; margin-top:8px; line-height:1.4; }
table{ width:100%; border-collapse:collapse; margin-top:4px; font-size:.86rem; }
th,td{ text-align:left; padding:8px 10px; border-bottom:1px solid var(--line); }
th{ font-family:var(--font-mono); font-size:.66rem; text-transform:uppercase; letter-spacing:.03em; color:var(--ink-faint); font-weight:600; }
td:first-child{ border-left:3px solid transparent; }
tr.acc-pd td:first-child{ border-left-color:var(--accent); }
tr.acc-ps td:first-child{ border-left-color:var(--accent2); }
code{ background:var(--panel-2); padding:2px 5px; border-radius:3px; font-family:var(--font-mono); font-size:.82rem; }
.bar-row{ display:flex; align-items:center; gap:8px; margin:4px 0; font-size:.8rem; }
.bar-label{ width:220px; flex-shrink:0; color:var(--ink-soft); }
.bar-track{ flex:1; background:var(--line); border-radius:4px; height:12px; overflow:hidden; }
.bar-fill{ background:var(--ink-faint); height:100%; }
.bar-value{ width:120px; flex-shrink:0; text-align:right; color:var(--ink-faint); font-family:var(--font-mono); font-size:.76rem; }
.seg-bar{ display:flex; height:20px; border-radius:4px; overflow:hidden; margin:10px 0 8px; background:var(--line); }
.seg-legend{ display:flex; gap:16px; flex-wrap:wrap; font-size:.78rem; color:var(--ink-soft); }
.seg-legend i{ display:inline-block; width:8px; height:8px; border-radius:2px; margin-right:5px; vertical-align:middle; }
.notes{ background:var(--panel-2); border:1px dashed var(--line-strong); border-radius:6px; padding:16px 18px; font-size:.8rem; color:var(--ink-soft); }
.notes li{ margin-bottom:8px; }
.badge{ display:inline-flex; align-items:center; padding:2px 10px; border-radius:20px; font-size:.72rem; font-weight:600; }
.badge.pd{ background:var(--accent-soft); color:var(--accent); }
.badge.ps{ background:var(--accent2-soft); color:var(--accent2); }
.tag{ display:inline-block; padding:2px 9px; border-radius:20px; font-size:.7rem; font-weight:600; }
.tag.tracked{ background:var(--good-soft); color:var(--good); }
.tag.gap{ background:var(--line); color:var(--ink-soft); }
.tag.t-good{ background:var(--good-soft); color:var(--good); }
.tag.t-warn{ background:var(--warn-soft); color:var(--warn); }
.tag.t-critical{ background:var(--critical-soft); color:var(--critical); }
.tag.t-neutral{ background:var(--neutral-soft); color:var(--neutral); }
.brand-table td{ white-space:nowrap; }
.coverage-note{ font-size:.8rem; color:var(--ink-soft); margin-top:10px; }
select#nlDaysSelect{ font:inherit; font-family:var(--font-mono); background:var(--panel); border:1px solid var(--line-strong); border-radius:3px; padding:1px 6px; color:var(--ink); }
footer{ color:var(--ink-faint); font-size:.75rem; margin-top:48px; }
.legend b, .coverage-note b, .notes b{ color:var(--ink); font-weight:650; }
.acc-card{ border-left-width:4px; border-left-style:solid; }
.acc-card.acc-pd{ border-left-color:var(--accent); }
.acc-card.acc-ps{ border-left-color:var(--accent2); }
.acc-card-head{ margin-bottom:14px; }
.acc-stat-grid{ display:grid; grid-template-columns:1fr 1fr; gap:14px 18px; }
.acc-stat-grid strong{ display:block; font-family:var(--font-mono); font-size:1.3rem; font-weight:600; color:var(--ink); }
.acc-stat-grid .lbl{ font-size:.74rem; color:var(--ink-soft); margin-top:2px; }
.acc-card small{ display:block; color:var(--ink-faint); font-size:.74rem; margin-top:14px; }
.jumpnav{ display:flex; flex-wrap:wrap; gap:8px; margin:0 0 34px; }
.jumpnav a{ font-family:var(--font-mono); font-size:.72rem; color:var(--ink-soft); text-decoration:none; background:var(--panel); border:1px solid var(--line); border-radius:20px; padding:4px 12px; }
.jumpnav a:hover{ border-color:var(--line-strong); color:var(--ink); }
.group{ margin-bottom:12px; }
.group-title{ font-family:var(--font-mono); font-size:.72rem; letter-spacing:.06em; text-transform:uppercase; color:var(--accent); font-weight:700; margin:0 0 6px; padding-top:8px; border-top:1px solid var(--line); }
.group:first-of-type .group-title{ border-top:none; padding-top:0; }
details.pi-table summary{ cursor:pointer; font-family:var(--font-body); font-weight:650; font-size:.85rem; color:var(--ink); margin:20px 0 8px; list-style:none; display:flex; align-items:center; gap:6px; }
details.pi-table summary::before{ content:"▸"; color:var(--ink-faint); font-size:.7rem; }
details.pi-table[open] summary::before{ content:"▾"; }
details.pi-table summary::-webkit-details-marker{ display:none; }
@media (max-width:640px){ .bar-label{ width:140px; } }
.en{ display:block; font-weight:400; font-style:italic; color:var(--ink-faint); text-transform:none; letter-spacing:normal; }
h1 .en{ font-size:.5rem; margin-top:2px; }
h2 .en{ font-size:.62rem; margin-top:2px; }
h3 .en{ font-size:.72rem; margin-top:1px; }
.card .lbl .en, .acc-stat-grid .lbl .en{ font-size:.92em; margin-top:1px; }
.legend .en, .coverage-note .en{ font-size:.95em; margin-top:2px; }
th .en{ font-size:.9em; margin-top:2px; text-transform:none; letter-spacing:normal; }
.badge .en, .tag .en{ display:block; font-size:.85em; margin-top:0; }
.jumpnav a .en{ font-size:.85em; margin-top:1px; }
.group-title .en{ display:inline; font-size:.85em; margin-top:0; margin-left:6px; font-weight:500; }
.stand-chip .en{ font-size:.85em; margin-top:1px; }
.bar-label .en{ font-size:.85em; margin-top:1px; }
.notes li .en{ font-size:.92em; margin-top:2px; }
footer .en{ font-size:.9em; margin-top:2px; }
summary .en{ font-size:.85em; margin-top:1px; font-weight:400; }
</style>
</head>
<body>
<div class="wrap">
<header class="top">
  <h1>Parfum Compliance Dashboard<span class="en">Parfum Compliance Dashboard</span></h1>
  <span class="stand-chip" id="standChip">${bi(`Stand: ${SNAPSHOT_ISO}`, `As of: ${SNAPSHOT_ISO}`)}</span>
</header>
<p class="sub">${bi(
  `Generiert aus Übersichtstabelle, GPSR-Issue-Workbook, GPSR-Graph-Workbook, der Prohibited-Ingredients-Liste und der Markenfreigaben-Liste. Alle Zeitfenster (7–90 Tage) rechnen relativ zum eingefrorenen Stand oben, nicht zur Uhrzeit beim Betrachten.`,
  `Generated from the overview workbook, the GPSR issue workbook, the GPSR graph workbook, the prohibited-ingredients list, and the brand approvals list. All time windows (7–90 days) are calculated relative to the frozen snapshot date above, not to the time you're viewing the page.`
)}</p>
<div class="jumpnav">
  <a href="#sec-overview">${bi("Überblick", "Overview")}</a>
  <a href="#sec-gpsr">${bi("GPSR-Compliance", "GPSR Compliance")}</a>
  <a href="#sec-violations">${bi("Blocked ASINs & Violations", "Blocked ASINs & Violations")}</a>
  <a href="#sec-ingredients">${bi("Verbotene Inhaltsstoffe", "Prohibited Ingredients")}</a>
  <a href="#sec-brands">${bi("Markenfreigaben (PD)", "Brand Approvals (PD)")}</a>
</div>

<div class="group" id="sec-overview">
<p class="group-title">${bi("Überblick", "Overview")}</p>
<section>
<h2>${bi("Snapshot", "Snapshot")}</h2>
<p class="legend">${bi("Die vier wichtigsten Zahlen auf einen Blick — Details und Herkunft in den Sektionen darunter.", "The four key numbers at a glance — details and source in the sections below.")}</p>
<div class="grid">
  <div class="card">
    <strong id="nlCount">${fmtInt(totalNewListingsRecent)}</strong>
    <div class="lbl">${bi("Bearbeitete Listings", "Processed Listings")} (${bi("letzte", "last")} <select id="nlDaysSelect">
      <option value="7">7</option><option value="14">14</option><option value="30" selected>30</option><option value="60">60</option><option value="90">90</option>
    </select> ${bi("Tage", "days")})</div>
    <small id="nlBreakdown">${sortedEntries(nlRecentByAccount).map(([k, v]) => `${esc(k)}: ${fmtInt(v)}`).join(" · ") || "keine im Zeitraum / none in this period"} · ab / from ${nlCutoffISO}</small>
    <small>${bi(`Lifetime gesamt: ${fmtInt(totalNewListings)}${nlDaysSinceLatest !== null ? ` · Tracker zuletzt aktualisiert vor ${nlDaysSinceLatest} Tagen (${nlLatestDateISO})` : ""}`, `Lifetime total: ${fmtInt(totalNewListings)}${nlDaysSinceLatest !== null ? ` · tracker last updated ${nlDaysSinceLatest} days ago (${nlLatestDateISO})` : ""}`)}</small>
  </div>
  <div class="card"><strong>${fmtInt(totalGpsrCases)}</strong><div class="lbl">${bi("GPSR-Fälle (echte Order-Zeilen)", "GPSR cases (real order rows)")}</div><small>${bi(`vs. ${fmtInt(gpsrTemplateRowCount)} leere Vorlagenzeilen ausgeschlossen · nur DE/ES/IT erfasst`, `vs. ${fmtInt(gpsrTemplateRowCount)} empty template rows excluded · only DE/ES/IT tracked`)}</small></div>
  <div class="card"><strong>${fmtEUR(totalRevenue)}</strong><div class="lbl">${bi("Revenue (letzte 12 Monate, PD+PS)", "Revenue (last 12 months, PD+PS)")}</div><small>Parfum Direct: ${fmtEUR(revenueByAccount["Parfum Direct"].sum)} · Parfum Store: ${fmtEUR(revenueByAccount["Parfum Store"].sum)}</small></div>
  <div class="card"><strong>${fmtInt(totalProhibited)}</strong><div class="lbl">${bi("Verbotene Inhaltsstoffe (Lilial/Lyral)", "Prohibited ingredients (Lilial/Lyral)")}</div><small>${sortedEntries(piByAccount).map(([k, v]) => `${esc(k)}: ${fmtInt(v)}`).join(" · ")}</small></div>
</div>
</section>

<section>
<h2>${bi("Account-Übersicht", "Account Overview")}</h2>
<p class="legend">${bi("Die vier wichtigsten Kennzahlen je Account — GPSR-Fälle, verbotene Inhaltsstoffe (Lilial), bearbeitete Listings und Umsatz.", "The four key metrics per account — GPSR cases, prohibited ingredients (Lilial), processed listings, and revenue.")}</p>
<div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(280px,1fr));">
${["Parfum Direct", "Parfum Store"]
  .map((acc) => {
    const gpsrN = gpsrCases.filter((c) => c.account === acc).length;
    return `<div class="card acc-card ${accentClass(acc)}">
      <div class="acc-card-head"><span class="badge ${acc === "Parfum Direct" ? "pd" : "ps"}">${esc(acc)}</span></div>
      <div class="acc-stat-grid">
        <div><strong>${fmtInt(gpsrN)}</strong><div class="lbl">${bi("GPSR-Fälle", "GPSR cases")}</div></div>
        <div><strong>${fmtInt(piByAccount[acc] || 0)}</strong><div class="lbl">${bi("Verbotene Inhaltsstoffe (Lilial)", "Prohibited ingredients (Lilial)")}</div></div>
        <div><strong data-nl-account="${esc(acc)}">${fmtInt(nlRecentByAccount[acc] || 0)}</strong><div class="lbl">${bi(`Bearb. Listings (${NL_RECENT_DAYS}d)`, `Processed listings (${NL_RECENT_DAYS}d)`)}</div></div>
        <div><strong>${fmtEUR(revenueByAccount[acc].sum)}</strong><div class="lbl">${bi("Revenue (12M)", "Revenue (12M)")}</div></div>
      </div>
      <small>${bi(`Produkte mit Umsatz: ${fmtInt(productsWithRevenueByAccount[acc] || 0)}`, `Products with revenue: ${fmtInt(productsWithRevenueByAccount[acc] || 0)}`)}</small>
    </div>`;
  })
  .join("")}
</div>
</section>
</div>

<div class="group" id="sec-gpsr">
<p class="group-title">${bi("GPSR-Compliance", "GPSR Compliance")}</p>
<section>
<h2 id="nlStatusHeading">${bi(`Bearbeitete Listings — Status-Verteilung (letzte ${NL_RECENT_DAYS} Tage)`, `Processed Listings — status breakdown (last ${NL_RECENT_DAYS} days)`)}</h2>
<p class="legend">${bi("Bearbeitungsstand der im gewählten Zeitraum zuletzt geprüften Listing-Zeilen — kein Neuanlage-Datum.", "Processing status of the listing rows most recently reviewed in the selected period — not a creation date.")}</p>
<div id="nlStatusBars">${segBarHtml(sortedEntries(nlByStatusRecent), totalNewListingsRecent)}</div>
</section>

<section>
<h2>${bi("GPSR-Fälle nach Markt — Datenabdeckung", "GPSR cases by market — data coverage")}</h2>
<p class="legend">${bi(`Zeigt, für welche Märkte überhaupt Fälle in dieser Quelle erfasst werden — „–" bedeutet Datenlücke, nicht „keine Beschwerden".`, `Shows which markets have cases tracked in this source at all — "–" means a data gap, not "no complaints".`)}</p>
${accountTable(
  [["Markt", "Market"], ["Fälle", "Cases"], ["Status", "Status"]],
  gpsrMarketCoverage
    .map(
      (m) =>
        `<tr><td>${esc(m.market)}</td><td>${m.tracked ? fmtInt(m.cases) : "–"}</td><td>${m.tracked ? `<span class="tag tracked">${bi("erfasst", "tracked")}</span>` : `<span class="tag gap">${bi("keine Daten in dieser Quelle", "no data in this source")}</span>`}</td></tr>`
    )
    .join("")
)}
<p class="coverage-note">${bi("Für Märkte ohne Eintrag oben liegen in der GPSR-Fall-Tabelle keine Zeilen vor — das ist eine Lücke in der Datenerhebung, kein Beleg für weniger Beschwerden.", "For markets with no entry above, the GPSR case table has no rows — this is a gap in data collection, not evidence of fewer complaints.")}</p>
</section>

<section>
<h2>${bi("GPSR-Fälle nach Ursache (Flag)", "GPSR cases by cause (flag)")}</h2>
<p class="legend">${bi("Beschwerdegrund laut Amazon-Seller-Central-Case (siehe Datenabdeckung oben).", "Complaint reason per the Amazon Seller Central case (see data coverage above).")}</p>
${barHtml(sortedEntries(gpsrByFlag).slice(0, 12), totalGpsrCases)}
</section>

<section>
<h2>${bi("GPSR-Priority-Einreichungen — Status", "GPSR priority submissions — status")}</h2>
<p class="legend">${bi(
  `Bearbeitungsstand der bei Amazon eingereichten GPSR-Fälle (Priority-Listen), beide Accounts zusammen. „Approved" = Details eingereicht und von Amazon bereits genehmigt. „Submitted" = Details eingereicht, aber von Amazon noch nicht genehmigt oder abgelehnt.`,
  `Processing status of GPSR cases submitted to Amazon (priority lists), both accounts combined. "Approved" = details submitted and already approved by Amazon. "Submitted" = details submitted, but not yet approved or declined by Amazon.`
)}</p>
${segBarHtml(sortedEntries(submissionStatusCounts), totalPriorityEntries)}
</section>

<section>
<h2>${bi("Wöchentlicher GPSR-Trend", "Weekly GPSR trend")}</h2>
<p class="legend">${bi(
  `Approved- vs. Submitted-Verlauf der letzten 2 Monate, getrennt nach Account — Punkte zeigen den genauen Wert beim Überfahren mit der Maus. „Approved" = Details eingereicht und von Amazon bereits genehmigt. „Submitted" = Details eingereicht, aber von Amazon noch nicht genehmigt oder abgelehnt.`,
  `Approved vs. Submitted trend for the last 2 months, split by account — hover over a point to see its exact value. "Approved" = details submitted and already approved by Amazon. "Submitted" = details submitted, but not yet approved or declined by Amazon.`
)}</p>
<h3>Parfum Direct</h3>
${lineChartSvg(
  weeklyTrendRecent.map((w) => w.period),
  [
    { name: "Approved", color: "var(--good)", values: weeklyTrendRecent.map((w) => w.pdApproved) },
    { name: "Submitted", color: "var(--warn)", values: weeklyTrendRecent.map((w) => w.pdSubmitted) },
  ]
)}
${chartLegend([
  { name: "Approved", color: "var(--good)" },
  { name: "Submitted", color: "var(--warn)" },
])}
<h3>Parfum Store</h3>
${lineChartSvg(
  weeklyTrendRecent.map((w) => w.period),
  [
    { name: "Approved", color: "var(--good)", values: weeklyTrendRecent.map((w) => w.psApproved) },
    { name: "Submitted", color: "var(--warn)", values: weeklyTrendRecent.map((w) => w.psSubmitted) },
  ]
)}
${chartLegend([
  { name: "Approved", color: "var(--good)" },
  { name: "Submitted", color: "var(--warn)" },
])}
<details class="pi-table">
<summary>${bi(`Alle ${weeklyTrend.length} Wochen als Tabelle anzeigen`, `Show all ${weeklyTrend.length} weeks as a table`)}</summary>
${accountTable(
  [["Woche", "Week"], ["PD: Approved SKUs", "PD: Approved SKUs"], ["PD: Submitted SKUs", "PD: Submitted SKUs"], ["PD: Neue GPSR", "PD: New GPSR"], ["PS: Bestand Start", "PS: Starting stock"], ["PS: Approved SKUs", "PS: Approved SKUs"], ["PS: Submitted SKUs", "PS: Submitted SKUs"], ["PS: Neue GPSR", "PS: New GPSR"]],
  weeklyTrend
    .map(
      (w) =>
        `<tr><td>${esc(w.period)}</td><td>${w.pdApproved ?? "–"}</td><td>${w.pdSubmitted ?? "–"}</td><td>${w.pdNew ?? "–"}</td><td>${w.psBeginning ?? "–"}</td><td>${w.psApproved ?? "–"}</td><td>${w.psSubmitted ?? "–"}</td><td>${w.psNew ?? "–"}</td></tr>`
    )
    .join("")
)}
</details>
</section>
</div>

<div class="group" id="sec-violations">
<p class="group-title">${bi("Blocked ASINs & Violations", "Blocked ASINs & Violations")}</p>
<section>
<h2>${bi("Blocked ASINs — Status", "Blocked ASINs — status")}</h2>
<p class="legend">${bi("Aktueller Bearbeitungsstand blockierter Listings, beide Accounts zusammen.", "Current processing status of blocked listings, both accounts combined.")}</p>
${segBarHtml(sortedEntries(blockedByStatus), blockedAsins.length)}
<h3>${bi("Top-Gründe", "Top reasons")}</h3>
${barHtml(sortedEntries(blockedByReason).slice(0, 10), blockedAsins.length)}
</section>

<section>
<h2>${bi("Account Violations", "Account Violations")}</h2>
<p class="legend">${bi("Verstöße gegen Amazon-Richtlinien nach Bearbeitungsstand und Typ.", "Violations of Amazon policy by processing status and type.")}</p>
${segBarHtml(sortedEntries(violationsByStatus), violations.length)}
<h3>${bi("Nach Typ", "By type")}</h3>
${barHtml(sortedEntries(violationsByType).slice(0, 10), violations.length)}
</section>
</div>

<div class="group" id="sec-ingredients">
<p class="group-title">${bi("Verbotene Inhaltsstoffe", "Prohibited Ingredients")}</p>
<section>
<h2>${bi("Verbotene Inhaltsstoffe (Lilial/Lyral)", "Prohibited ingredients (Lilial/Lyral)")}</h2>
<p class="legend">${bi("ASINs mit gemeldetem Lilial-/Lyral-Gehalt — eigenständige Compliance-Liste, nicht Teil der GPSR-Fall-Tabelle oben.", "ASINs with reported Lilial/Lyral content — a standalone compliance list, not part of the GPSR case table above.")}</p>
<div class="grid">
  <div class="card"><strong>${fmtInt(totalProhibited)}</strong><div class="lbl">${bi("Einträge gesamt", "Total entries")}</div></div>
  <div class="card"><strong>${fmtInt(piAsins.size)}</strong><div class="lbl">${bi("Eindeutige ASINs", "Unique ASINs")}</div></div>
  <div class="card"><strong>${fmtInt(prohibited.filter((p) => p.eanOnly).length)}</strong><div class="lbl">${bi("Nur EAN, keine ASIN", "EAN only, no ASIN")}</div></div>
</div>
<h3>${bi("Status-Verteilung", "Status breakdown")}</h3>
${barHtml(sortedEntries(piByStatus), totalProhibited)}
<details class="pi-table">
<summary>${bi("Erste 30 Einträge anzeigen", "Show first 30 entries")}</summary>
${accountTable(
  [["#", "#"], ["Account", "Account"], ["ASIN/EAN", "ASIN/EAN"], ["Produkt", "Product"], ["Status", "Status"], ["Fehlende Unterlagen", "Missing documents"], ["Case ID", "Case ID"]],
  prohibited
    .slice(0, 30)
    .map(
      (p, i) =>
        `<tr class="${accentClass(p.account)}"><td>${i + 1}</td><td>${esc(p.account)}</td><td>${esc(p.asin || p.eanOnly || "–")}</td><td>${esc((p.productTitle || "").slice(0, 60))}</td><td>${esc(p.status)}</td><td>${esc(p.missingDocs.join("; ") || "–")}</td><td>${esc(p.caseId || "–")}</td></tr>`
    )
    .join("")
)}
</details>
</section>
</div>

<div class="group" id="sec-brands">
<p class="group-title">${bi("Markenfreigaben", "Brand Approvals")}</p>
<section class="acc-card acc-pd" style="padding-left:14px;">
<h2>${bi("Markenfreigaben nach Marktplatz", "Brand approvals by marketplace")} <span class="badge pd">Parfum Direct</span></h2>
<p class="legend">${bi("Gilt laut Kunde ausschließlich für Parfum Direct — keine Datumsspalte in der Quelle, deshalb als eigener Überblick statt in eine bestehende Sektion eingerechnet. Zellen zeigen den Originaltext beim Überfahren mit der Maus.", "Per the client, this applies exclusively to Parfum Direct — no date column in the source, so it's shown as its own overview rather than folded into an existing section. Hover a cell to see the original wording.")}</p>
<div class="grid">
  <div class="card"><strong>${fmtInt(brandApprovals.length)}</strong><div class="lbl">${bi("Marken erfasst", "Brands tracked")}</div></div>
  <div class="card"><strong>${fmtInt(BRAND_MARKETS.length)}</strong><div class="lbl">${bi("Marktplätze", "Marketplaces")}</div></div>
  <div class="card"><strong>${fmtInt(allBaCells.length)}</strong><div class="lbl">${bi("Ausgefüllte Zellen gesamt", "Filled cells total")}</div></div>
</div>
<h3>${bi("Verteilung über alle Zellen", "Breakdown across all cells")}</h3>
${brandStatusBarHtml(baCategoryEntries, allBaCells.length)}
<h3>${bi("Übersicht je Marke", "Overview by brand")}</h3>
<div style="overflow-x:auto;">
${accountTable(
  [["Marke", "Brand"], ...BRAND_MARKETS.map((m) => [m, m])],
  brandApprovals
    .map(
      (b) =>
        `<tr><td>${esc(b.brand)}</td>${b.cells
          .map((c) => `<td><span class="tag t-${c.token}" title="${esc(c.raw || "(leer)")}">${esc(c.label)}</span></td>`)
          .join("")}</tr>`
    )
    .join("")
).replace("<table>", '<table class="brand-table">')}
</div>
</section>
</div>

<footer>${bi(
  `Generiert von generate.js — Stand ${SNAPSHOT_ISO} — Quellen: Übersichtstabelle.xlsx, Parfum Direct products with GPSR issue.xlsx, GPSR Graph.xlsx, List of ASINs containing prohibited ingredients.xlsx, Status of brand approvals.xlsx`,
  `Generated by generate.js — as of ${SNAPSHOT_ISO} — sources: Übersichtstabelle.xlsx, Parfum Direct products with GPSR issue.xlsx, GPSR Graph.xlsx, List of ASINs containing prohibited ingredients.xlsx, Status of brand approvals.xlsx`
)}</footer>
</div>
<script>
(function () {
  var nlData = ${JSON.stringify(nlFilterData)}; // [[dateISO, account, status], ...] - only rows with a valid date
  var SNAPSHOT_TS = ${SNAPSHOT_NOW.getTime()}; // frozen generation time - the day-window filter anchors here, not to the viewer's clock
  var STATUS_TOKEN = { good: "var(--good)", warn: "var(--warn)", critical: "var(--critical)", neutral: "var(--neutral)" };
  function statusToken(label) {
    var s = String(label || "").toLowerCase();
    if (/done|approved|erledigt/.test(s)) return "good";
    if (/progress|submitted|review/.test(s)) return "warn";
    if (/blocked|open|offen/.test(s)) return "critical";
    return "neutral";
  }
  var select = document.getElementById("nlDaysSelect");
  function fmtIntJs(n) { return n.toLocaleString("de-DE"); }
  function escJs(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
  function renderSegBar(counts, total) {
    var entries = Object.keys(counts).map(function (k) { return [k, counts[k]]; });
    entries.sort(function (a, b) { return b[1] - a[1]; });
    if (!entries.length || !total) {
      return '<div class="seg-bar"></div><p style="color:var(--ink-faint); font-size:.8rem; margin:4px 0 0;">Keine Einträge im gewählten Zeitraum.<span class="en">No entries in the selected period.</span></p>';
    }
    var segs = entries.map(function (e) {
      var pct = ((e[1] / total) * 100).toFixed(2);
      return '<div style="width:' + pct + '%; background:' + STATUS_TOKEN[statusToken(e[0])] + ';" title="' + escJs(e[0]) + '"></div>';
    }).join("");
    var legend = entries.map(function (e) {
      var pct = ((e[1] / total) * 100).toFixed(1);
      return '<span><i style="background:' + STATUS_TOKEN[statusToken(e[0])] + ';"></i>' + escJs(e[0]) + ' · ' + fmtIntJs(e[1]) + ' (' + pct + '%)</span>';
    }).join("");
    return '<div class="seg-bar">' + segs + '</div><div class="seg-legend">' + legend + '</div>';
  }
  function update() {
    var days = parseInt(select.value, 10);
    var cutoff = new Date(SNAPSHOT_TS - days * 86400000).toISOString().slice(0, 10);
    var accCounts = {}, statusCounts = {}, total = 0;
    for (var i = 0; i < nlData.length; i++) {
      if (nlData[i][0] >= cutoff) {
        total++;
        var acc = nlData[i][1];
        var status = nlData[i][2] || "(kein Status)";
        accCounts[acc] = (accCounts[acc] || 0) + 1;
        statusCounts[status] = (statusCounts[status] || 0) + 1;
      }
    }
    var pd = accCounts["Parfum Direct"] || 0;
    var ps = accCounts["Parfum Store"] || 0;
    var other = total - pd - ps;
    var parts = [];
    if (pd) parts.push("Parfum Direct: " + fmtIntJs(pd));
    if (ps) parts.push("Parfum Store: " + fmtIntJs(ps));
    if (other) parts.push("Nicht zuordenbar: " + fmtIntJs(other));
    document.getElementById("nlCount").textContent = fmtIntJs(total);
    document.getElementById("nlBreakdown").textContent = (parts.join(" · ") || "keine im Zeitraum / none in this period") + " · ab / from " + cutoff;
    document.getElementById("nlStatusHeading").innerHTML = "Bearbeitete Listings — Status-Verteilung (letzte " + days + " Tage)" +
      '<span class="en">Processed Listings — status breakdown (last ' + days + ' days)</span>';
    document.getElementById("nlStatusBars").innerHTML = renderSegBar(statusCounts, total);
    var pdCell = document.querySelector('[data-nl-account="Parfum Direct"]');
    var psCell = document.querySelector('[data-nl-account="Parfum Store"]');
    if (pdCell) pdCell.textContent = fmtIntJs(pd);
    if (psCell) psCell.textContent = fmtIntJs(ps);
  }
  select.addEventListener("change", update);
  update(); // sync display to the frozen snapshot date, not the live clock
})();
</script>
</body>
</html>
`;

fs.mkdirSync(OUT_DIR, { recursive: true });
fs.writeFileSync(path.join(OUT_DIR, "index.html"), html, "utf8");
console.log("Written:", path.join(OUT_DIR, "index.html"));
console.log("\n--- Key numbers ---");
console.log("Snapshot date:", SNAPSHOT_ISO);
console.log("New Listings total:", totalNewListings, nlByAccount);
console.log("Bearbeitete Listings (recent):", totalNewListingsRecent, nlRecentByAccount);
console.log("GPSR real cases:", totalGpsrCases, "(template rows excluded:", gpsrTemplateRowCount, ")");
console.log("Revenue total (fixed):", fmtEUR(totalRevenue), revenueByAccount);
console.log("Prohibited ingredients:", totalProhibited, piByAccount);
console.log("\n--- Notes ---");
notes.forEach((n, i) => console.log(`${i + 1}. ${n.de}`));
