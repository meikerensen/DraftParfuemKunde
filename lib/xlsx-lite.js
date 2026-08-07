// Minimal dependency-free XLSX reader (zip + shared-strings + sheet XML).
// Good enough for reading tabular data out of workbooks without needing npm install.
const fs = require("fs");
const zlib = require("zlib");

function readZip(path) {
  const buf = fs.readFileSync(path);
  const entries = {};
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd === -1) throw new Error("Not a zip file: " + path);
  const cdOff = buf.readUInt32LE(eocd + 16);
  const cnt = buf.readUInt16LE(eocd + 10);
  let p = cdOff;
  for (let n = 0; n < cnt; n++) {
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commLen = buf.readUInt16LE(p + 32);
    const lho = buf.readUInt32LE(p + 42);
    const name = buf.toString("utf8", p + 46, p + 46 + nameLen);
    const lnameLen = buf.readUInt16LE(lho + 26);
    const lextraLen = buf.readUInt16LE(lho + 28);
    const dataStart = lho + 30 + lnameLen + lextraLen;
    const method = buf.readUInt16LE(lho + 8);
    const compSize = buf.readUInt32LE(p + 20);
    const raw = buf.slice(dataStart, dataStart + compSize);
    entries[name] = method === 0 ? raw : zlib.inflateRawSync(raw);
    p += 46 + nameLen + extraLen + commLen;
  }
  return entries;
}

function decodeXmlEntities(s) {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function sharedStrings(z) {
  const x = z["xl/sharedStrings.xml"];
  if (!x) return [];
  const xml = x.toString("utf8");
  const out = [];
  const parts = xml.split("<si>").slice(1);
  for (const pt of parts) {
    const si = pt.split("</si>")[0];
    const texts = [...si.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decodeXmlEntities(m[1]));
    out.push(texts.join(""));
  }
  return out;
}

function colToNum(c) {
  let n = 0;
  for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function parseSheetXml(xml, ss) {
  const rows = {};
  const rowMatches = [...xml.matchAll(/<row[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)];
  for (const rm of rowMatches) {
    const rn = parseInt(rm[1]);
    const cells = {};
    const cellM = [...rm[2].matchAll(/<c[^>]*r="([A-Z]+)\d+"([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/g)];
    for (const cm of cellM) {
      const col = colToNum(cm[1]);
      const attrs = cm[2];
      const inner = cm[3] || "";
      const t = (attrs.match(/t="([^"]*)"/) || [])[1];
      let v = "";
      if (t === "s") {
        const vv = (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
        v = vv !== undefined ? ss[parseInt(vv)] : "";
      } else if (t === "inlineStr") {
        v = decodeXmlEntities((inner.match(/<t[^>]*>([\s\S]*?)<\/t>/) || ["", ""])[1]);
      } else {
        const vv = (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
        v = vv !== undefined ? vv : "";
      }
      cells[col] = v;
    }
    rows[rn] = cells;
  }
  return rows;
}

function sheetOrder(z) {
  const wbxml = z["xl/workbook.xml"].toString("utf8");
  const rels = z["xl/_rels/workbook.xml.rels"].toString("utf8");
  const relMap = {};
  [...rels.matchAll(/<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"/g)].forEach(
    (m) => (relMap[m[1]] = m[2])
  );
  return [...wbxml.matchAll(/<sheet[^>]*name="([^"]*)"[^>]*r:id="([^"]*)"/g)].map((m) => ({
    name: decodeXmlEntities(m[1]),
    file: "xl/" + relMap[m[2]].replace(/^\//, "").replace(/^xl\//, ""),
  }));
}

// Loads a workbook once; returns { listSheets(), getSheet(name) -> {row:{col:value}} , getRows(name) -> array of row-arrays (0-indexed, dense) }
function loadWorkbook(path) {
  const z = readZip(path);
  const ss = sharedStrings(z);
  const sheets = sheetOrder(z);
  const cache = {};
  function getSheetRaw(name) {
    if (cache[name]) return cache[name];
    const sh = sheets.find((s) => s.name === name);
    if (!sh) throw new Error(`Sheet "${name}" not found in ${path}. Available: ${sheets.map((s) => s.name).join(", ")}`);
    const rows = parseSheetXml(z[sh.file].toString("utf8"), ss);
    cache[name] = rows;
    return rows;
  }
  return {
    listSheets: () => sheets.map((s) => s.name),
    getSheet: getSheetRaw,
    // Returns dense array-of-arrays, 1 row = 1 array indexed from col 0, in physical row order (1-indexed row numbers preserved as .rowNum)
    getRows(name) {
      const raw = getSheetRaw(name);
      const rowNums = Object.keys(raw).map(Number).sort((a, b) => a - b);
      return rowNums.map((rn) => {
        const cells = raw[rn];
        const maxCol = Math.max(-1, ...Object.keys(cells).map(Number));
        const arr = [];
        for (let c = 0; c <= maxCol; c++) arr[c] = cells[c] !== undefined ? cells[c] : "";
        arr.rowNum = rn;
        return arr;
      });
    },
  };
}

module.exports = { loadWorkbook };
