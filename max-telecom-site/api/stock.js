// GET /api/stock — public stock statuses and retail prices for the site.
// Never returns exact quantities. Returns {configured:false} until 1C is connected,
// in which case the site simply hides the stock block.
const { kv, kvGet, odataConfigured, pullFromOData } = require("./_lib/stock");

module.exports = async (req, res) => {
  if (req.method !== "GET" && req.method !== "HEAD") {
    res.setHeader("Allow", "GET, HEAD");
    return res.status(405).json({ error: "method not allowed" });
  }
  try {
    if (kv()) {
      const data = await kvGet();
      if (data) {
        res.setHeader("Cache-Control", "public, s-maxage=60, stale-while-revalidate=600");
        return res.status(200).json(data);
      }
    }
    if (odataConfigured()) {
      const data = await pullFromOData();
      res.setHeader("Cache-Control", "public, s-maxage=900, stale-while-revalidate=3600");
      return res.status(200).json(data);
    }
    res.setHeader("Cache-Control", "public, s-maxage=300");
    return res.status(200).json({ configured: false, items: [] });
  } catch (e) {
    console.error("stock:", e.message);
    res.setHeader("Cache-Control", "no-store");
    return res.status(502).json({ configured: true, error: "stock temporarily unavailable", items: [] });
  }
};
