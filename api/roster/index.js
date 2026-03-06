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
    const csvPath = path.join(process.cwd(), "data", "staff.csv");
    const raw = fs.readFileSync(csvPath, "utf8");
    const rows = parseCsv(raw);

    const staff = rows
      .map(r => {
        const name = r.name || r.staff || r.employee || r.Name || r.Staff || "";
        const roleRaw = (r.role || r.Role || "").toUpperCase();
        const role = roleRaw.includes("XR") ? "XRT" : "MA";
        const isFloat = /float/i.test(r.type || r.Type || r.notes || r.Notes || name);
        const email = r.email || r.Email || "";
        const location = r.location || r.Location || "";
        return {
          id: toId(name),
          name,
          role,
          isFloat,
          email,
          location
        };
      })
      .filter(s => s.name);

    if (!staff.some(s => s.id === "ma-float")) {
      staff.push({
        id: "ma-float",
        name: "MA Float",
        role: "MA",
        isFloat: true,
        email: "",
        location: ""
      });
    }

    if (!staff.some(s => s.id === "xr-float")) {
      staff.push({
        id: "xr-float",
        name: "XR Float",
        role: "XRT",
        isFloat: true,
        email: "",
        location: ""
      });
    }

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: { staff }
    };
  } catch (err) {
    context.res = {
      status: 500,
      body: { error: err.message || "Failed to load roster." }
    };
  }
};
