const { TableClient } = require("@azure/data-tables");

function getClient() {
  const cs = process.env.AZURE_STORAGE_CONNECTION_STRING;
  const table = process.env.SCHEDULES_TABLE_NAME || "Schedules";
  if (!cs) throw new Error("Missing AZURE_STORAGE_CONNECTION_STRING");
  return TableClient.fromConnectionString(cs, table);
}

function json(res, status, body) {
  res.status = status;
  res.headers = { "Content-Type": "application/json" };
  res.body = JSON.stringify(body);
  return res;
}

module.exports = async function (context, req) {
  try {
    const client = getClient();

    // GET default providers
    if ((req.method || "").toUpperCase() === "GET") {
      try {
        const entity = await client.getEntity("providers", "default");
        const providers = entity?.data ? JSON.parse(entity.data) : null;
        return (context.res = json(context.res, 200, { ok: true, providers }));
      } catch {
        return (context.res = json(context.res, 200, { ok: true, providers: null }));
      }
    }

    // POST save default providers
    const providers = req.body?.providers;
    if (!Array.isArray(providers)) {
      return (context.res = json(context.res, 400, { ok: false, error: "Body must include providers: []" }));
    }

    // Normalize: strings only, trimmed, unique, non-empty
    const clean = Array.from(new Set(
      providers
        .map(x => (x ?? "").toString().trim())
        .filter(Boolean)
    ));

    const entity = {
      partitionKey: "providers",
      rowKey: "default",
      updatedAt: new Date().toISOString(),
      data: JSON.stringify(clean)
    };

    await client.upsertEntity(entity, "Replace");
    return (context.res = json(context.res, 200, { ok: true, count: clean.length }));
  } catch (err) {
    context.log.error(err);
    return (context.res = json(context.res, 500, { ok: false, error: err.message || "Server error" }));
  }
};
