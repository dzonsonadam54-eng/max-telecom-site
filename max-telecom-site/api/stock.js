// GET /api/stock — public stock statuses and retail prices for the site.
// Never returns exact quantities. Returns {configured:false} until 1C is connected,
// in which case the site simply hides the stock block.
//
// Modes:
//   PULL (ODATA_URL set, cloud 1C): read OData; if storage is configured, keep the last good
//        result there and serve it when 1C is unreachable ("stale": true).
//   PUSH (no ODATA_URL, storage set, office 1C): serve what the 1C job uploaded.
const { kv, kvGet, kvSet, odataConfigured, pullFromOData } = require("./_lib/stock");

module.exports = async (req, res) => {
  if (req.method !== "GET" && req.method !== "HEAD") {
    res.setHeader("Allow", "GET, HEAD");
    return res.status(405).json({ error: "method not allowed" });
  }

  if (odataConfigured()) {
    try {
      const data = await pullFromOData();
      if (kv()) await kvSet(data).catch((e) => console.error("stock cache write:", e.message));
      res.setHeader("Cache-Control", "public, s-maxage=900, stale-while-revalidate=3600");
      return res.status(200).json(data);
    } catch (e) {
      console.error("stock pull:", e.message);
      const last = kv() ? await kvGet().catch(() => null) : null;
      if (last) {
        // 1C is down: serve the last good copy, retry soon.
        res.setHeader("Cache-Control", "public, s-maxage=120");
        return res.status(200).json({ ...last, stale: true });
      }
      res.setHeader("Cache-Control", "no-store");
      return res.status(502).json({ configured: true, error: "stock temporarily unavailable", items: [] });
    }
  }

  try {
    if (kv()) {
      const data = await kvGet();
      if (data) {
        res.setHeader("Cache-Control", "public, s-maxage=60, stale-while-revalidate=600");
        return res.status(200).json(data);
      }
    }
    res.setHeader("Cache-Control", "public, s-maxage=300");
    return res.status(200).json({ configured: false, items: [] });
  } catch (e) {
    console.error("stock:", e.message);
    res.setHeader("Cache-Control", "no-store");
    return res.status(502).json({ configured: true, error: "stock temporarily unavailable", items: [] });
  }
};
