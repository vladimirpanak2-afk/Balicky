const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const XLSX = require("xlsx");

const app = express();
const PORT = Number(process.env.PORT || 8787);
const ROOT = path.resolve(__dirname, "..");
const DATA_DIR = path.join(ROOT, "data");
const SESSION_FILE = path.join(DATA_DIR, "sessions.json");

app.use(cors());
app.use(express.json({ limit: "5mb" }));
app.use(express.static(ROOT));

function readJsonSafe(filePath, fallback) {
  try {
    if (!fs.existsSync(filePath)) return fallback;
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch {
    return fallback;
  }
}

function writeJsonSafe(filePath, value) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, JSON.stringify(value, null, 2), "utf8");
}

function loadGlobalArrayFromScript(filePath, variableName) {
  if (!fs.existsSync(filePath)) return null;
  const code = fs.readFileSync(filePath, "utf8");
  const context = { window: {} };
  vm.createContext(context);
  vm.runInContext(code, context, { timeout: 3000, filename: filePath });
  const value = context.window[variableName];
  return Array.isArray(value) ? value : null;
}

function datasetFiles(countryRaw) {
  const country = String(countryRaw || "").toLowerCase();
  if (country === "cz") {
    return {
      catalog: [
        path.join(ROOT, "katalog_cz.js"),
        path.join(ROOT, "katalog.js"),
      ],
      packages: [
        path.join(ROOT, "balicky_cz.js"),
        path.join(ROOT, "balicky.js"),
      ],
    };
  }
  return {
    catalog: [path.join(ROOT, "katalog_sk.js"), path.join(ROOT, "katalog.js")],
    packages: [path.join(ROOT, "balicky_sk.js"), path.join(ROOT, "balicky.js")],
  };
}

function firstExisting(paths) {
  return paths.find((filePath) => fs.existsSync(filePath)) || null;
}

function loadPackagesFromXlsx(filePath) {
  if (!filePath || !fs.existsSync(filePath)) return null;
  const wb = XLSX.readFile(filePath, { cellDates: false });
  const sheetName = wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });
  if (!rows.length) return [];
  return rows.map((r) => [
    r.ID_BALICKU || "",
    r.ID_POLOZKY || "",
    r.NAZEV_POLOZKY_KRATKY || "",
    r.NEAKTIVNI || 0,
    r.PM || "",
  ]);
}

function loadDataset(country) {
  const files = datasetFiles(country);
  const catalogFile = firstExisting(files.catalog);
  const packagesFile = firstExisting(files.packages);
  const xlsxFallback = firstExisting([
    path.join(ROOT, `balicky_${String(country || "sk").toLowerCase()}.xlsx`),
    path.join(ROOT, "balicky_sk.xlsx"),
    path.join(ROOT, "balicky.xlsx"),
  ]);
  const catalog = catalogFile
    ? loadGlobalArrayFromScript(catalogFile, "KATALOG")
    : null;
  let packages = packagesFile
    ? loadGlobalArrayFromScript(packagesFile, "BALICKY")
    : null;
  if (!packages || !packages.length) {
    packages = loadPackagesFromXlsx(xlsxFallback) || [];
  }

  return {
    country: String(country || "sk").toLowerCase(),
    catalog: catalog || [],
    packages: packages || [],
    files: {
      catalog: catalogFile ? path.basename(catalogFile) : null,
      packages: packagesFile
        ? path.basename(packagesFile)
        : xlsxFallback
        ? path.basename(xlsxFallback)
        : null,
    },
  };
}

function sessionKey(country, user, pm) {
  return `${country}::${user}::${pm}`.toLowerCase();
}

app.get("/api/health", (_req, res) => {
  res.json({ ok: true });
});

app.get("/api/countries", (_req, res) => {
  res.json({
    countries: [
      { code: "sk", label: "Slovensko" },
      { code: "cz", label: "Cesko" },
    ],
  });
});

app.get("/api/data/:country", (req, res) => {
  const payload = loadDataset(req.params.country);
  res.json(payload);
});

app.get("/api/session", (req, res) => {
  const country = String(req.query.country || "sk").toLowerCase();
  const user = String(req.query.user || "").trim();
  const pm = String(req.query.pm || "").trim();
  if (!user || !pm) {
    return res.status(400).json({ error: "Missing user or pm query parameter." });
  }
  const sessions = readJsonSafe(SESSION_FILE, {});
  const key = sessionKey(country, user, pm);
  return res.json({ session: sessions[key] || null });
});

app.put("/api/session", (req, res) => {
  const country = String(req.body.country || "sk").toLowerCase();
  const user = String(req.body.user || "").trim();
  const pm = String(req.body.pm || "").trim();
  const session = req.body.session;
  if (!user || !pm || !session) {
    return res.status(400).json({ error: "Missing country, user, pm or session." });
  }

  const sessions = readJsonSafe(SESSION_FILE, {});
  const key = sessionKey(country, user, pm);
  sessions[key] = {
    ...session,
    country,
    savedBy: user,
    savedAt: new Date().toISOString(),
  };
  writeJsonSafe(SESSION_FILE, sessions);
  return res.json({ ok: true });
});

app.delete("/api/session", (req, res) => {
  const country = String(req.query.country || "sk").toLowerCase();
  const user = String(req.query.user || "").trim();
  const pm = String(req.query.pm || "").trim();
  if (!user || !pm) {
    return res.status(400).json({ error: "Missing user or pm query parameter." });
  }
  const sessions = readJsonSafe(SESSION_FILE, {});
  const key = sessionKey(country, user, pm);
  delete sessions[key];
  writeJsonSafe(SESSION_FILE, sessions);
  return res.json({ ok: true });
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
