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

function toPublic(raw) {
  const items = [];
  for (const r of raw.items || []) {
    if (!r || !r.name) continue;
    const price = Number(r.price);
    // Visibility rule the shop owner controls from 1C: an item is on the site only if it has a retail price.
    if (config.requirePrice && !(price > 0)) continue;
    items.push({
      sku: String(r.sku || "").slice(0, 40),
      name: String(r.name).slice(0, 200),
      cat: categoryFor(r.name, r.group),
      group: String(r.group || "").slice(0, 80),
      unit: String(r.unit || "").slice(0, 12),
      status: statusFor(r.qty, r.unit),
      price: Number.isFinite(price) && price > 0 ? Math.round(price * 100) / 100 : null,
    });
  }
  items.sort((a, b) => a.name.localeCompare(b.name, "ru"));
  return { configured: true, updated: raw.updated || new Date().toISOString(), items };
}

// ---------- PUSH storage: Upstash Redis REST (Vercel Marketplace) ----------

const KEY = "max:stock:v1";

function kv() {
  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
  return url && token ? { url: url.replace(/\/$/, ""), token } : null;
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

module.exports = { toPublic, statusFor, categoryFor, kv, kvGet, kvSet, odataConfigured, pullFromOData, safeEqual };
