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

    // GET: return default roster
    if ((req.method || "").toUpperCase() === "GET") {
      try {
        const entity = await client.getEntity("roster", "default");
        const roster = entity?.data ? JSON.parse(entity.data) : null;
        return (context.res = json(context.res, 200, { ok: true, roster }));
      } catch (e) {
        // Not found -> return null roster
        return (context.res = json(context.res, 200, { ok: true, roster: null }));
      }
    }

    // POST: save default roster
    const roster = req.body?.roster;
    if (!Array.isArray(roster)) {
      return (context.res = json(context.res, 400, { ok: false, error: "Body must include roster: []" }));
    }

    // Minimal validation: required fields
    for (const r of roster) {
      if (!r || !r.id || !r.name || !r.role) {
        return (context.res = json(context.res, 400, { ok: false, error: "Each roster item must have id, name, role" }));
      }
    }

    const entity = {
      partitionKey: "roster",
      rowKey: "default",
      updatedAt: new Date().toISOString(),
      data: JSON.stringify(roster)
    };

    await client.upsertEntity(entity, "Replace");
    return (context.res = json(context.res, 200, { ok: true }));
  } catch (err) {
    context.log.error(err);
    return (context.res = json(context.res, 500, { ok: false, error: err.message || "Server error" }));
  }
};
