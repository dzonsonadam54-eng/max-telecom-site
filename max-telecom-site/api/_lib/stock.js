// Shared logic for the 1C stock integration.
// No external dependencies: runs on the plain Vercel Node.js runtime.
//
// Data flow (two supported modes, see docs/1c-integration.md):
//   PULL  — 1C is reachable from the internet (cloud 1C): /api/stock reads OData directly,
//           Vercel's CDN caches the answer for 15 minutes.
//   PUSH  — 1C sits on an office server: a 1C scheduled job POSTs raw data to
//           /api/stock-upload; we store the already-sanitised result in Upstash Redis.
//
// Raw item (what 1C provides):  { sku, name, group, unit, qty, price }
// Public item (what the site sees): { sku, name, cat, group, unit, status, price }
// Exact quantities never leave the server.

const config = require("./config.json");

// ---------- normalisation ----------

const KEYWORDS = config.keywords; // [[substring, categoryId], ...] checked in order

function categoryFor(name, group) {
  if (group && config.groups[group]) return config.groups[group];
  const hay = `${group || ""} ${name || ""}`.toLowerCase();
  for (const [needle, cat] of KEYWORDS) if (hay.includes(needle)) return cat;
  return null;
}

function lowThreshold(unit) {
  const u = String(unit || "").trim().toLowerCase();
  return config.lowThreshold[u] ?? config.lowThreshold.default;
}

function statusFor(qty, unit) {
  const q = Number(qty) || 0;
  if (q <= 0) return "order";
  if (q < lowThreshold(unit)) return "low";
  return "in";
}

const EXCLUDE = new Set((config.excludeGroups || []).map((g) => g.trim().toLowerCase()));
const ACCOUNTS = new Set((config.accounts || []).map(String));
const EXCLUDE_NAMES = (config.excludeNames || []).map((x) => x.toLowerCase());
const AUTO = config.autoGroups || [];

function autoGroup(name) {
  const n = String(name || "").toLowerCase();
  for (const [needle, id] of AUTO) if (n.includes(needle)) return id;
  return "other";
}

function toPublic(raw) {
  const items = [];
  for (const r of raw.items || []) {
    if (!r || !r.name) continue;
    if (r.group && EXCLUDE.has(String(r.group).trim().toLowerCase())) continue;
    if (ACCOUNTS.size && r.account && !ACCOUNTS.has(String(r.account).trim())) continue;
    const lname = String(r.name).toLowerCase();
    if (EXCLUDE_NAMES.some((x) => lname.includes(x))) continue;
    const price = Number(r.price);
    // Visibility rule the shop owner controls from 1C: an item is on the site only if it has a retail price.
    if (config.requirePrice && !(price > 0)) continue;
    items.push({
      sku: String(r.sku || "").slice(0, 40),
      name: String(r.name).slice(0, 200),
      cat: categoryFor(r.name, r.group),
      group: String(r.group || "").slice(0, 80),
      g: r.group ? undefined : autoGroup(r.name),
      unit: String(r.unit || "").slice(0, 12),
      status: statusFor(r.qty, r.unit),
      price: Number.isFinite(price) && price > 0 ? Math.round(price * 100) / 100 : null,
    });
  }
  items.sort((a, b) => a.name.localeCompare(b.name, "ru"));
  return { configured: true, updated: raw.updated || new Date().toISOString(), groups: config.groupNames || {}, items };
}

// ---------- 1C report export (Excel) -> raw items ----------
// Works with a "Universal report" saved to .xlsx: finds the header row by column titles,
// reads rows until "Итого". Column titles may vary slightly between 1C versions.

const H = {
  name: /^(товар|номенклатура|наименование)$/i,
  qty: /(количеств|остаток)/i,
  group: /(групп|родитель)/i,
  sku: /^(артикул|код)$/i,
  unit: /(ед\.?\s*изм|единиц)/i,
  account: /^сч[её]т/i,
  // Retail price only. Cost columns ("Учетная цена", "Себестоимость", purchase) are never read.
  price: /^(цена|цена продажи|розничная цена)$/i,
};
const FINAL = /конечн\w*\s+остат/i;

function toNumber(v) {
  const n = Number(String(v ?? "").replace(/[\s ]/g, "").replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

// Finds the header in a 1C report saved to Excel. Supports:
//  - one-row headers (Universal report: "Товар | Группа | Количество остаток")
//  - two-row headers (Material statement: "Конечный остаток" over "Учетная цена | Количество | Сумма"),
//    where the quantity is the "Количество" under "Конечный остаток".
function findHeader(rows) {
  const str = (r) => (r || []).map((c) => String(c ?? "").trim());
  for (let i = 0; i < Math.min(rows.length, 60); i++) {
    const r = str(rows[i]);
    const name = r.findIndex((c) => H.name.test(c));
    if (name < 0) continue;
    let qty = -1, start = i + 1;
    const fin = r.findIndex((c) => FINAL.test(c));
    if (fin >= 0) {
      const sub = str(rows[i + 1]);
      const nextHead = r.findIndex((c, j) => j > fin && c);
      const stop = nextHead > 0 ? nextHead : sub.length;
      for (let j = fin; j < stop; j++) if (/^количеств/i.test(sub[j])) { qty = j; break; }
      if (qty >= 0) start = i + 2;
    }
    if (qty < 0) qty = r.findIndex((c, j) => j !== name && H.qty.test(c) && !/начальн/i.test(c));
    if (qty < 0) continue;
    const col = { name, qty };
    for (const k of ["group", "sku", "unit", "account", "price"]) {
      const j = r.findIndex((c, x) => x !== name && x !== qty && H[k].test(c));
      if (j >= 0) col[k] = j;
    }
    return { col, start };
  }
  return null;
}

function fromReportRows(rows) {
  const h = findHeader(rows);
  if (!h) throw new Error("В файле не найдены колонки «Номенклатура» и «Количество»");
  const { col, start } = h;
  const cell = (r, k) => (col[k] != null ? String(r[col[k]] ?? "").trim() : "");
  const items = [];
  for (let i = start; i < rows.length; i++) {
    const r = rows[i] || [];
    const name = String(r[col.name] ?? "").replace(/\s+/g, " ").trim();
    if (/^итого/i.test(String(r[0] ?? "").trim()) || /^итого/i.test(name)) break;
    if (!name) continue;
    items.push({
      name,
      qty: toNumber(r[col.qty]),
      group: cell(r, "group"),
      sku: cell(r, "sku"),
      unit: cell(r, "unit"),
      account: cell(r, "account"),
      price: col.price != null ? toNumber(r[col.price]) : null,
    });
  }
  // The same product can appear on several rows (batches, several 1C codes): sum quantities.
  const merged = new Map();
  for (const it of items) {
    const key = `${it.account}|${it.name.toLowerCase()}|${it.unit.toLowerCase()}`;
    const prev = merged.get(key);
    if (prev) prev.qty += it.qty; else merged.set(key, { ...it });
  }
  return [...merged.values()];
}

// ---------- PUSH storage: Upstash Redis REST (Vercel Marketplace) ----------

const KEY = "max:stock:v1";
const PREV_KEY = "max:stock:prev";

function kv() {
  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
  return url && token ? { url: url.replace(/\/$/, ""), token } : null;
}

async function kvGetKey(key) {
  const c = kv();
  if (!c) return null;
  const r = await fetch(`${c.url}/get/${encodeURIComponent(key)}`, { headers: { Authorization: `Bearer ${c.token}` } });
  if (!r.ok) throw new Error(`storage read ${r.status}`);
  const { result } = await r.json();
  return result ? JSON.parse(result) : null;
}

async function kvSetKey(key, value) {
  const c = kv();
  if (!c) throw new Error("storage is not configured");
  const r = await fetch(`${c.url}/set/${encodeURIComponent(key)}`, {
    method: "POST",
    headers: { Authorization: `Bearer ${c.token}` },
    body: JSON.stringify(value),
  });
  if (!r.ok) throw new Error(`storage write ${r.status}`);
}

// Publish a new catalog, keeping the previous one for one-click rollback.
async function publish(value) {
  const cur = await kvGetKey(KEY).catch(() => null);
  if (cur) await kvSetKey(PREV_KEY, cur);
  await kvSetKey(KEY, value);
}

async function rollback() {
  const prev = await kvGetKey(PREV_KEY);
  if (!prev) return null;
  const cur = await kvGetKey(KEY).catch(() => null);
  await kvSetKey(KEY, prev);
  if (cur) await kvSetKey(PREV_KEY, cur);
  return prev;
}

async function kvGet() {
  const c = kv();
  if (!c) return null;
  const r = await fetch(`${c.url}/get/${encodeURIComponent(KEY)}`, { headers: { Authorization: `Bearer ${c.token}` } });
  if (!r.ok) throw new Error(`storage read ${r.status}`);
  const { result } = await r.json();
  return result ? JSON.parse(result) : null;
}

async function kvSet(value) {
  const c = kv();
  if (!c) throw new Error("storage is not configured");
  const r = await fetch(`${c.url}/set/${encodeURIComponent(KEY)}`, {
    method: "POST",
    headers: { Authorization: `Bearer ${c.token}` },
    body: JSON.stringify(value),
  });
  if (!r.ok) throw new Error(`storage write ${r.status}`);
}

// ---------- PULL: 1C OData ----------

function odataConfigured() {
  return Boolean(process.env.ODATA_URL && process.env.ODATA_USER);
}

async function odataGet(path) {
  const base = process.env.ODATA_URL.replace(/\/$/, "");
  const auth = Buffer.from(`${process.env.ODATA_USER}:${process.env.ODATA_PASS || ""}`).toString("base64");
  const sep = path.includes("?") ? "&" : "?";
  const r = await fetch(`${base}/${encodeURI(path)}${sep}$format=json`, {
    headers: { Authorization: `Basic ${auth}`, Accept: "application/json" },
    signal: AbortSignal.timeout(20000),
  });
  if (!r.ok) {
    const e = new Error(`1C OData ${r.status} on ${path.split("?")[0]}`);
    e.status = r.status;
    throw e;
  }
  return (await r.json()).value || [];
}

async function pullFromOData() {
  const env = process.env;
  const unitsCatalog = env.ODATA_UNITS_CATALOG || "Catalog_УпаковкиЕдиницыИзмерения";
  const stockReg = env.ODATA_STOCK_REGISTER || "AccumulationRegister_ТоварыНаСкладах";
  const whCond = env.ODATA_WAREHOUSE_KEY ? `Condition='Склад_Key eq guid'${env.ODATA_WAREHOUSE_KEY}''` : "";

  const [nomenclature, units, balance] = await Promise.all([
    odataGet("Catalog_Номенклатура?$select=Ref_Key,Description,Артикул,Parent_Key,IsFolder,DeletionMark,ЕдиницаИзмерения_Key,ТипНоменклатуры")
      .catch(() => odataGet("Catalog_Номенклатура?$select=Ref_Key,Description,Артикул,Parent_Key,IsFolder,DeletionMark,ЕдиницаИзмерения_Key")),
    odataGet(`${unitsCatalog}?$select=Ref_Key,Description`).catch(() => []),
    odataGet(`${stockReg}/Balance(${whCond})`),
  ]);

  // Price register is ЦеныНоменклатуры25 in УТ 11.5+, ЦеныНоменклатуры in 11.4 — try both.
  let prices = [];
  if (env.PRICE_TYPE_KEY) {
    const cond = `SliceLast(Condition='ВидЦены_Key eq guid'${env.PRICE_TYPE_KEY}'')`;
    const regs = env.ODATA_PRICE_REGISTER
      ? [env.ODATA_PRICE_REGISTER]
      : ["InformationRegister_ЦеныНоменклатуры25", "InformationRegister_ЦеныНоменклатуры"];
    for (const reg of regs) {
      try { prices = await odataGet(`${reg}/${cond}`); break; } catch (e) { if (e.status !== 404 && e.status !== 400) throw e; }
    }
  }

  const groupName = new Map(nomenclature.filter((n) => n.IsFolder).map((n) => [n.Ref_Key, n.Description]));
  const unitName = new Map(units.map((u) => [u.Ref_Key, u.Description]));
  const qty = new Map();
  for (const b of balance) {
    const free = (Number(b.ВНаличииBalance) || 0) - (Number(b.ВРезервеСоСкладаBalance) || 0);
    qty.set(b.Номенклатура_Key, (qty.get(b.Номенклатура_Key) || 0) + free);
  }
  const price = new Map();
  for (const p of prices) {
    const v = Number(p.Цена);
    if (!(v > 0)) continue;
    const prev = price.get(p.Номенклатура_Key);
    price.set(p.Номенклатура_Key, prev ? Math.min(prev, v) : v); // several characteristics → "from" price
  }

  const items = nomenclature
    .filter((n) => !n.IsFolder && !n.DeletionMark)
    .filter((n) => !n.ТипНоменклатуры || n.ТипНоменклатуры === "Товар") // goods only, no services
    .map((n) => ({
      sku: n.Артикул || "",
      name: n.Description,
      group: groupName.get(n.Parent_Key) || "",
      unit: unitName.get(n.ЕдиницаИзмерения_Key) || "",
      qty: qty.get(n.Ref_Key) || 0,
      price: price.get(n.Ref_Key) || null,
    }));
  return toPublic({ updated: new Date().toISOString(), items });
}

// ---------- helpers ----------

function safeEqual(a, b) {
  const x = Buffer.from(String(a)), y = Buffer.from(String(b));
  return x.length === y.length && require("crypto").timingSafeEqual(x, y);
}

module.exports = { toPublic, statusFor, categoryFor, fromReportRows, kv, kvGet, kvSet, publish, rollback, odataConfigured, pullFromOData, safeEqual };
