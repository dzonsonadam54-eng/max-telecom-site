// Minimal .xlsx reader without dependencies (Node zlib only).
// Reads the first worksheet into an array of rows (arrays of strings).
// Enough for 1C "Save as Excel 2007" report exports.
const zlib = require("zlib");

function unzip(buf) {
  // Locate the End Of Central Directory record.
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("not a zip/xlsx file");
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const files = {};
  for (let n = 0; n < count; n++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error("bad zip directory");
    const method = buf.readUInt16LE(p + 10);
    const csize = buf.readUInt32LE(p + 20);
    const nlen = buf.readUInt16LE(p + 28), xlen = buf.readUInt16LE(p + 30), clen = buf.readUInt16LE(p + 32);
    const local = buf.readUInt32LE(p + 42);
    const name = buf.slice(p + 46, p + 46 + nlen).toString("utf8");
    const lnlen = buf.readUInt16LE(local + 26), lxlen = buf.readUInt16LE(local + 28);
    const start = local + 30 + lnlen + lxlen;
    const data = buf.slice(start, start + csize);
    files[name] = () => (method === 0 ? data : zlib.inflateRawSync(data));
    p += 46 + nlen + xlen + clen;
  }
  return files;
}

const decode = (s) => s
  .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d))
  .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
  .replace(/&amp;/g, "&");

function colIndex(ref) {
  const letters = ref.replace(/\d+/g, "");
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function readXlsx(buf) {
  const files = unzip(buf);
  const get = (name) => (files[name] ? files[name]().toString("utf8") : null);

  const shared = [];
  const sst = get("xl/sharedStrings.xml");
  if (sst) {
    for (const si of sst.match(/<si>[\s\S]*?<\/si>/g) || []) {
      const parts = [...si.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decode(m[1]));
      shared.push(parts.join(""));
    }
  }

  // First sheet: follow workbook relationships when present, else sheet1.xml.
  let sheetPath = "xl/worksheets/sheet1.xml";
  const wb = get("xl/workbook.xml"), rels = get("xl/_rels/workbook.xml.rels");
  if (wb && rels) {
    const rid = (wb.match(/<sheet\b[^>]*r:id="([^"]+)"/) || [])[1];
    const target = rid && (rels.match(new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`)) || rels.match(new RegExp(`Target="([^"]+)"[^>]*Id="${rid}"`)) || [])[1];
    if (target) sheetPath = "xl/" + target.replace(/^\/?xl\//, "").replace(/^\//, "");
  }
  const sheet = get(sheetPath);
  if (!sheet) throw new Error("worksheet not found");

  const rows = [];
  for (const rm of sheet.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
    const row = [];
    for (const cm of rm[1].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const attrs = cm[1], inner = cm[2] || "";
      const ref = (attrs.match(/\br="([A-Z]+\d+)"/) || [])[1];
      const type = (attrs.match(/\bt="([^"]+)"/) || [])[1];
      let v = "";
      if (type === "inlineStr") v = [...inner.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decode(m[1])).join("");
      else {
        const raw = (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
        if (raw != null) v = type === "s" ? shared[+raw] ?? "" : decode(raw);
      }
      const idx = ref ? colIndex(ref) : row.length;
      row[idx] = String(v).trim();
    }
    rows.push(Array.from(row, (x) => x ?? ""));
  }
  return rows;
}

module.exports = { readXlsx };
