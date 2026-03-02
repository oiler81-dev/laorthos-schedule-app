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

    const weeks = [];

    const entities = tableClient.listEntities({
      queryOptions: { filter: "RowKey eq 'published'" }
    });

    for await (const entity of entities) {
      weeks.push(entity.partitionKey);
    }

    weeks.sort((a, b) => (a < b ? 1 : -1));

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: { weeks }
    };

  } catch (err) {
    context.res = {
      status: 500,
      body: err.message
    };
  }
};
