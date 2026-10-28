// POST /api/stock-upload — called by a 1C scheduled job (PUSH mode).
// Auth: header "Authorization: Bearer <STOCK_PUSH_TOKEN>".
// Body: { updated, items: [{ sku, name, group, unit, qty, price }] }
// Only the sanitised public view (statuses, retail prices) is stored.
const { toPublic, kvSet, safeEqual } = require("./_lib/stock");

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "method not allowed" });
  }
  const expected = process.env.STOCK_PUSH_TOKEN;
  if (!expected || expected.length < 24) {
    return res.status(503).json({ error: "upload is not configured" });
  }
  const got = String(req.headers.authorization || "").replace(/^Bearer\s+/i, "");
  if (!safeEqual(got, expected)) {
    return res.status(401).json({ error: "unauthorized" });
  }
  let body = req.body;
  if (typeof body === "string") {
    try { body = JSON.parse(body); } catch { return res.status(400).json({ error: "invalid JSON" }); }
  }
  if (!body || !Array.isArray(body.items)) {
    return res.status(400).json({ error: "expected { items: [...] }" });
  }
  if (body.items.length > 20000) {
    return res.status(413).json({ error: "too many items" });
  }
  try {
    const pub = toPublic(body);
    await kvSet(pub);
    return res.status(200).json({ ok: true, received: body.items.length, shown: pub.items.filter((i) => i.cat).length });
  } catch (e) {
    console.error("stock-upload:", e.message);
    return res.status(500).json({ error: "could not save" });
  }
};
