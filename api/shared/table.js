const { TableClient, AzureNamedKeyCredential } = require("@azure/data-tables");

function getClient(tableName) {
  const accountName = process.env.STORAGE_ACCOUNT_NAME;
  const accountKey = process.env.STORAGE_ACCOUNT_KEY;

  if (!accountName || !accountKey) {
    throw new Error("Missing STORAGE_ACCOUNT_NAME or STORAGE_ACCOUNT_KEY app settings.");
  }

  const credential = new AzureNamedKeyCredential(accountName, accountKey);
  const serviceUrl = `https://${accountName}.table.core.windows.net`;
  return new TableClient(serviceUrl, tableName, credential);
}

async function ensureTable(tableName) {
  const client = getClient(tableName);
  try {
    await client.createTable();
  } catch (err) {
    if (err.statusCode !== 409) throw err;
  }
  return client;
}

module.exports = {
  getClient,
  ensureTable
};
