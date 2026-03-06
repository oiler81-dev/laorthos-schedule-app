const fs = require("fs");
const path = require("path");

function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter(Boolean);
  if (!lines.length) return [];
  const headers = lines[0].split(",").map(h => h.trim());
  return lines.slice(1).map(line => {
    const cols = line.split(",").map(v => v.trim());
    const row = {};
    headers.forEach((h, i) => row[h] = cols[i] || "");
    return row;
  });
}

function toId(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "");
}

module.exports = async function (context) {
  try {
    const csvPath = path.join(process.cwd(), "data", "providers.csv");
    const raw = fs.readFileSync(csvPath, "utf8");
    const rows = parseCsv(raw);

    const providers = rows
      .map(r => {
        const name = r.provider || r.name || r.Provider || r.Name || "";
        const location = r.location || r.Location || "";
        const specialty = r.specialty || r.Specialty || "";
        const locations = r.locations || r.Locations || "";
        return {
          id: toId(name),
          name,
          location,
          specialty,
          locations
        };
      })
      .filter(p => p.name);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: { providers }
    };
  } catch (err) {
    context.res = {
      status: 500,
      body: { error: err.message || "Failed to load providers." }
    };
  }
};
