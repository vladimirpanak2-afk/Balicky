const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { Firestore } = require("@google-cloud/firestore");
const XLSX = require("xlsx");

const app = express();
const PORT = Number(process.env.PORT || 8787);
const ROOT = path.resolve(__dirname, "..");
const DATA_DIR = process.env.DATA_DIR
  ? path.resolve(process.env.DATA_DIR)
  : path.join(ROOT, "data");
const SESSION_FILE = path.join(DATA_DIR, "sessions.json");
const USERS_FILE = path.join(DATA_DIR, "users.json");
const SESSION_BACKEND = String(process.env.SESSION_BACKEND || "").toLowerCase();
const USE_FIRESTORE =
  SESSION_BACKEND === "firestore" ||
  (!SESSION_BACKEND && !!process.env.K_SERVICE);
let firestore = null;
const DATASET_CACHE = new Map();

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

if (USE_FIRESTORE) {
  try {
    firestore = new Firestore();
    console.log("Session backend: Firestore");
  } catch (err) {
    console.warn("Firestore init failed, using file sessions backend.", err);
    firestore = null;
  }
} else {
  console.log("Session backend: file");
}

async function getSessionByKey(key) {
  if (firestore) {
    const doc = await firestore.collection("sessions").doc(key).get();
    return doc.exists ? doc.data() : null;
  }
  const sessions = readJsonSafe(SESSION_FILE, {});
  return sessions[key] || null;
}

async function setSessionByKey(key, value) {
  if (firestore) {
    await firestore.collection("sessions").doc(key).set(value);
    return;
  }
  const sessions = readJsonSafe(SESSION_FILE, {});
  sessions[key] = value;
  writeJsonSafe(SESSION_FILE, sessions);
}

async function deleteSessionByKey(key) {
  if (firestore) {
    await firestore.collection("sessions").doc(key).delete();
    return;
  }
  const sessions = readJsonSafe(SESSION_FILE, {});
  delete sessions[key];
  writeJsonSafe(SESSION_FILE, sessions);
}

function normalizeLogin(login) {
  return String(login || "").trim().toLowerCase();
}

function sanitizeUser(user) {
  if (!user) return null;
  return {
    login: normalizeLogin(user.login),
    firstName: String(user.firstName || "").trim(),
    lastName: String(user.lastName || "").trim(),
    active: user.active !== false,
  };
}

async function getUserByLogin(loginRaw) {
  const login = normalizeLogin(loginRaw);
  if (!login) return null;
  if (firestore) {
    const doc = await firestore.collection("users").doc(login).get();
    return doc.exists ? doc.data() : null;
  }
  const users = readJsonSafe(USERS_FILE, []);
  return users.find((u) => normalizeLogin(u.login) === login) || null;
}

async function verifyUser(loginRaw, pinRaw) {
  const login = normalizeLogin(loginRaw);
  const pin = String(pinRaw || "").trim();
  if (!login || !/^\d{6}$/.test(pin)) return null;
  const user = await getUserByLogin(login);
  if (!user || user.active === false) return null;
  if (String(user.pin || "") !== pin) return null;
  return sanitizeUser(user);
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

function fileStamp(filePath) {
  if (!filePath || !fs.existsSync(filePath)) return "missing";
  const s = fs.statSync(filePath);
  return `${path.basename(filePath)}:${s.size}:${Math.floor(s.mtimeMs)}`;
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
  const keyCountry = String(country || "sk").toLowerCase();
  const files = datasetFiles(keyCountry);
  const catalogFile = firstExisting(files.catalog);
  const packagesFile = firstExisting(files.packages);
  const xlsxFallback = firstExisting([
    path.join(ROOT, `balicky_${keyCountry}.xlsx`),
    path.join(ROOT, "balicky_sk.xlsx"),
    path.join(ROOT, "balicky.xlsx"),
  ]);
  const version = [
    fileStamp(catalogFile),
    fileStamp(packagesFile),
    fileStamp(xlsxFallback),
  ].join("|");

  const cached = DATASET_CACHE.get(keyCountry);
  if (cached && cached.version === version) return cached.payload;

  const catalog = catalogFile
    ? loadGlobalArrayFromScript(catalogFile, "KATALOG")
    : null;
  let packages = packagesFile
    ? loadGlobalArrayFromScript(packagesFile, "BALICKY")
    : null;
  if (!packages || !packages.length) {
    packages = loadPackagesFromXlsx(xlsxFallback) || [];
  }

  const payload = {
    country: keyCountry,
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
  DATASET_CACHE.set(keyCountry, { version, payload });
  return payload;
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

app.post("/api/auth/login", async (req, res) => {
  const login = req.body?.login;
  const pin = req.body?.pin;
  if (!login || !pin) {
    return res.status(400).json({ error: "Missing login or pin." });
  }
  try {
    const user = await verifyUser(login, pin);
    if (!user) return res.status(401).json({ error: "Invalid login or pin." });
    return res.json({ ok: true, user });
  } catch (err) {
    console.error("POST /api/auth/login failed", err);
    return res.status(500).json({ error: "Login failed." });
  }
});

app.get("/api/data/:country", (req, res) => {
  const payload = loadDataset(req.params.country);
  res.json(payload);
});

app.get("/api/session", async (req, res) => {
  const country = String(req.query.country || "sk").toLowerCase();
  const user = String(req.query.user || "").trim();
  const pm = String(req.query.pm || "").trim();
  if (!user || !pm) {
    return res.status(400).json({ error: "Missing user or pm query parameter." });
  }
  try {
    const key = sessionKey(country, user, pm);
    const session = await getSessionByKey(key);
    return res.json({ session: session || null });
  } catch (err) {
    console.error("GET /api/session failed", err);
    return res.status(500).json({ error: "Failed to load session." });
  }
});

app.put("/api/session", async (req, res) => {
  const country = String(req.body.country || "sk").toLowerCase();
  const user = String(req.body.user || "").trim();
  const pm = String(req.body.pm || "").trim();
  const session = req.body.session;
  if (!user || !pm || !session) {
    return res.status(400).json({ error: "Missing country, user, pm or session." });
  }

  try {
    const key = sessionKey(country, user, pm);
    await setSessionByKey(key, {
      ...session,
      country,
      savedBy: user,
      savedAt: new Date().toISOString(),
    });
    return res.json({ ok: true });
  } catch (err) {
    console.error("PUT /api/session failed", err);
    return res.status(500).json({ error: "Failed to save session." });
  }
});

app.delete("/api/session", async (req, res) => {
  const country = String(req.query.country || "sk").toLowerCase();
  const user = String(req.query.user || "").trim();
  const pm = String(req.query.pm || "").trim();
  if (!user || !pm) {
    return res.status(400).json({ error: "Missing user or pm query parameter." });
  }
  try {
    const key = sessionKey(country, user, pm);
    await deleteSessionByKey(key);
    return res.json({ ok: true });
  } catch (err) {
    console.error("DELETE /api/session failed", err);
    return res.status(500).json({ error: "Failed to delete session." });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
