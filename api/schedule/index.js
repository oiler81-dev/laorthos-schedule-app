const { TableClient } = require("@azure/data-tables");

module.exports = async function (context, req) {
  try {
    const connectionString = process.env.AZURE_STORAGE_CONNECTION_STRING;
    const tableName = process.env.SCHEDULES_TABLE_NAME || "Schedules";

    if (!connectionString) {
      context.res = { status: 500, body: "Missing storage connection string" };
      return;
    }

    const tableClient = TableClient.fromConnectionString(connectionString, tableName);

    const method = req.method.toUpperCase();

    if (method === "GET") {
      const weekOf = req.query.weekOf;
      const mode = req.query.mode || "published";

      if (!weekOf) {
        context.res = { status: 400, body: "Missing weekOf parameter" };
        return;
      }

      try {
        const entity = await tableClient.getEntity(weekOf, mode);

        context.res = {
          status: 200,
          headers: { "Content-Type": "application/json" },
          body: {
            weekOf,
            mode,
            data: JSON.parse(entity.data),
            updatedAt: entity.updatedAt,
            updatedBy: entity.updatedBy
          }
        };
      } catch {
        context.res = { status: 404, body: "Schedule not found" };
      }

      return;
    }

    if (method === "POST") {
      const headers = req.headers || {};
      const email =
        headers["x-ms-client-principal-name"] ||
        headers["X-MS-CLIENT-PRINCIPAL-NAME"] ||
        null;

      const editors = new Set([
        "nperez@unitymsk.com",
        "aledezma@laorthos.com"
      ]);

      if (!email || !editors.has(email.toLowerCase())) {
        context.res = { status: 403, body: "Not authorized" };
        return;
      }

      const { weekOf, mode, data } = req.body;

      if (!weekOf || !mode || !data) {
        context.res = { status: 400, body: "Missing weekOf, mode, or data" };
        return;
      }

      const entity = {
        partitionKey: weekOf,
        rowKey: mode,
        data: JSON.stringify(data),
        updatedAt: new Date().toISOString(),
        updatedBy: email
      };

      await tableClient.upsertEntity(entity, "Replace");

      context.res = {
        status: 200,
        body: { success: true }
      };

      return;
    }

    context.res = { status: 405, body: "Method not allowed" };

  } catch (err) {
    context.res = {
      status: 500,
      body: err.message
    };
  }
};
