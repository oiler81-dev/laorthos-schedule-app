const { ensureTable } = require("../shared/table");

const NON_SCHEDULED_STATUSES = new Set(["OFF", "PTO", "Holiday", "Admin", "Call Only"]);

function isFloatId(id = "") {
  return id === "ma-float" || id === "xr-float";
}

module.exports = async function (context, req) {
  try {
    if (req.method !== "POST") {
      context.res = { status: 405, body: { error: "Method not allowed" } };
      return;
    }

    const body = req.body || {};
    const weekOf = body.weekOf;
    const items = Array.isArray(body.items) ? body.items : [];

    if (!weekOf) {
      context.res = { status: 400, body: { error: "Missing weekOf" } };
      return;
    }

    const maConflicts = new Map();

    for (const item of items) {
      const status = item.status || "Scheduled";

      if (status !== "Scheduled") continue;

      if (item.maId && !isFloatId(item.maId)) {
        const key = `${item.day}__${item.maId}`;
        if (maConflicts.has(key) && maConflicts.get(key) !== item.providerId) {
          context.res = {
            status: 400,
            body: {
              error: `${item.maName || item.maId} is assigned to multiple providers on ${item.day}.`
            }
          };
          return;
        }
        maConflicts.set(key, item.providerId);
      }
    }

    const client = await ensureTable("schedule");

    for (const item of items) {
      const status = item.status || "Scheduled";
      const rowKey = `${item.providerId}__${item.day}`;

      await client.upsertEntity({
        partitionKey: `WEEK#${weekOf}`,
        rowKey,
        weekOf,
        providerId: item.providerId || "",
        providerName: item.providerName || "",
        day: item.day || "",
        status,
        maId: status === "Scheduled" ? (item.maId || "") : "",
        maName: status === "Scheduled" ? (item.maName || "") : "",
        xrtId: status === "Scheduled" ? (item.xrtId || "") : "",
        xrtName: status === "Scheduled" ? (item.xrtName || "") : "",
        location: status === "Scheduled" ? (item.location || "") : "",
        secondaryLocation: status === "Scheduled" ? (item.secondaryLocation || "") : "",
        time: status === "Scheduled" ? (item.time || "") : "",
        xrRoom: status === "Scheduled" ? (item.xrRoom || "") : "",
        notes: item.notes || ""
      }, "Merge");
    }

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: { ok: true, count: items.length }
    };
  } catch (err) {
    context.res = {
      status: 500,
      body: { error: err.message || "Failed to save schedule." }
    };
  }
};
