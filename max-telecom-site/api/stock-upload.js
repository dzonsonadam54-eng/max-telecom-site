// POST /api/stock-upload — updates the site catalog. Nobody connects to 1C: data arrives here.
// Auth: header "Authorization: Bearer <STOCK_PUSH_TOKEN>".
//
// Bodies (JSON):
//   { fileBase64, filename, dryRun? }  — 1C report saved as .xlsx (manual upload page or the mail relay)
//   { items: [{ sku, name, group, unit, qty, price }], updated? }  — structured push
//   { action: "rollback" }  — restore the previous catalog
// dryRun: parse and return a summary without publishing (used for the preview on the upload page).
// Only the public view (name, group, status, price) is stored; exact quantities are discarded.
const { toPublic, fromReportRows, publish, rollback, safeEqual } = require("./_lib/stock");
const { readXlsx } = require("./_lib/xlsx");

function summary(pub, received) {
  const groups = {};
  const statuses = { in: 0, low: 0, order: 0 };
  for (const i of pub.items) {
    const g = i.group || (pub.groups && pub.groups[i.g] && pub.groups[i.g].ru) || "—";
    groups[g] = (groups[g] || 0) + 1;
    statuses[i.status] = (statuses[i.status] || 0) + 1;
  }
  return { received, shown: pub.items.length, groups, statuses, sample: pub.items.slice(0, 8).map((i) => i.name) };
}

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
    return res.status(401).json({ error: "Неверный код доступа" });
  }

  let body = req.body;
  if (typeof body === "string") {
    try { body = JSON.parse(body); } catch { return res.status(400).json({ error: "invalid JSON" }); }
  }
  body = body || {};

  try {
    if (body.action === "rollback") {
      const prev = await rollback();
      if (!prev) return res.status(404).json({ error: "Предыдущей версии нет" });
      return res.status(200).json({ ok: true, restored: true, updated: prev.updated, shown: prev.items.length });
    }

    let raw;
    if (typeof body.fileBase64 === "string") {
      const buf = Buffer.from(body.fileBase64, "base64");
      if (buf.length > 4 * 1024 * 1024) return res.status(413).json({ error: "Файл слишком большой" });
      let rows;
      try { rows = readXlsx(buf); } catch (e) {
        return res.status(400).json({ error: "Не удалось прочитать файл. Сохраните отчёт из 1С как «Лист Excel 2007 (.xlsx)»" });
      }
      try { raw = { items: fromReportRows(rows), updated: body.updated }; } catch (e) {
        return res.status(422).json({ error: e.message });
      }
    } else if (Array.isArray(body.items)) {
      raw = body;
    } else {
      return res.status(400).json({ error: "expected { fileBase64 } or { items }" });
    }

    if (raw.items.length > 20000) return res.status(413).json({ error: "Слишком много позиций" });
    if (!raw.items.length) return res.status(422).json({ error: "В файле нет ни одной позиции" });

    const pub = toPublic(raw);
    const info = summary(pub, raw.items.length);
    if (body.dryRun) return res.status(200).json({ ok: true, dryRun: true, ...info });

    await publish(pub);
    return res.status(200).json({ ok: true, published: true, updated: pub.updated, ...info });
  } catch (e) {
    console.error("stock-upload:", e.message);
    return res.status(500).json({ error: "Не удалось сохранить. Попробуйте ещё раз" });
  }
};
