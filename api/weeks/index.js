const { ensureTable } = require("../shared/table");

module.exports = async function (context, req) {
  try {
    const weekOf = req.query.weekOf;
    if (!weekOf) {
      context.res = { status: 400, body: { error: "Missing weekOf" } };
      return;
    }

    const client = await ensureTable("schedule");
    const items = [];

    const entities = client.listEntities({
      queryOptions: {
        filter: `PartitionKey eq 'WEEK#${weekOf}'`
      }
    });

    for await (const entity of entities) {
      items.push({
        weekOf,
        providerId: entity.providerId || "",
        providerName: entity.providerName || "",
        day: entity.day || "",
        status: entity.status || "Scheduled",
        maId: entity.maId || "",
        maName: entity.maName || "",
        xrtId: entity.xrtId || "",
        xrtName: entity.xrtName || "",
        location: entity.location || "",
        secondaryLocation: entity.secondaryLocation || "",
        time: entity.time || "",
        xrRoom: entity.xrRoom || "",
        notes: entity.notes || ""
      });
    }

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: { items }
    };
  } catch (err) {
    context.res = {
      status: 500,
      body: { error: err.message || "Failed to load week schedule." }
    };
  }
};
