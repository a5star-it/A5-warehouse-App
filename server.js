import express from "express";
import path from "path";
import fs from "fs";
import crypto from "crypto";
import bcrypt from "bcryptjs";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
app.use(express.json({ limit: "30mb" }));

/* -----------------------------------------------------------------
   Storage: a single JSON file holding { key: value } pairs. App data
   (inventory, inbound-records, ...) lives under plain keys; accounts,
   sessions, and the audit log live under reserved "__"-prefixed keys
   in the same file, and are never reachable through the generic
   /api/kv endpoints below.

   IMPORTANT: on most hosting platforms the filesystem is WIPED on
   every redeploy unless you attach a persistent volume — see
   README.md.
----------------------------------------------------------------- */
const DATA_FILE =
  process.env.DATA_FILE ||
  (process.env.RAILWAY_VOLUME_MOUNT_PATH ? path.join(process.env.RAILWAY_VOLUME_MOUNT_PATH, "data.json") : path.join(__dirname, "data.json"));

function loadStore() {
  try {
    return JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
  } catch (e) {
    return {};
  }
}
function saveStore(store) {
  fs.mkdirSync(path.dirname(DATA_FILE), { recursive: true });
  fs.writeFileSync(DATA_FILE, JSON.stringify(store), "utf8");
}

function getUsers() {
  return loadStore().__users || [];
}
function setUsers(users) {
  const store = loadStore();
  store.__users = users;
  saveStore(store);
}
function getSessions() {
  return loadStore().__sessions || {};
}
function setSessions(sessions) {
  const store = loadStore();
  store.__sessions = sessions;
  saveStore(store);
}
function appendAudit(entry) {
  const store = loadStore();
  const log = store.__audit || [];
  log.unshift({ id: crypto.randomUUID(), timestamp: new Date().toISOString(), ...entry });
  store.__audit = log.slice(0, 2000); // keep the log from growing without bound
  saveStore(store);
}

/* -----------------------------------------------------------------
   Bootstrap the first admin account from environment variables. If
   they're not set and no accounts exist yet, nobody can log in —
   the server logs a clear warning so this is easy to spot.
----------------------------------------------------------------- */
function ensureAdminSeed() {
  const users = getUsers();
  if (users.length > 0) return;
  const username = process.env.ADMIN_USERNAME;
  const password = process.env.ADMIN_PASSWORD;
  if (!username || !password) {
    console.warn(
      "[warehouse-app] No accounts exist yet and ADMIN_USERNAME / ADMIN_PASSWORD are not set. " +
        "Set both in your environment and restart the server to create the first admin account."
    );
    return;
  }
  setUsers([{ username, passwordHash: bcrypt.hashSync(password, 10), role: "admin", createdAt: new Date().toISOString() }]);
  console.log(`[warehouse-app] Seeded initial admin account: ${username}`);
}
try {
  ensureAdminSeed();
} catch (e) {
  console.error("[warehouse-app] Failed to seed admin account on startup:", e.message);
  console.error("[warehouse-app] The server will still start, but nobody will be able to log in until this is fixed.");
}

function generateToken() {
  return crypto.randomBytes(32).toString("hex");
}
const SESSION_DAYS = 30;

function requireAuth(req, res, next) {
  const authHeader = req.headers.authorization || "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;
  if (!token) return res.status(401).json({ error: "not authenticated" });
  const sessions = getSessions();
  const session = sessions[token];
  if (!session || new Date(session.expiresAt) < new Date()) {
    return res.status(401).json({ error: "session expired" });
  }
  req.user = { username: session.username, role: session.role };
  req.token = token;
  next();
}
function requireAdmin(req, res, next) {
  if (!req.user || req.user.role !== "admin") return res.status(403).json({ error: "admin only" });
  next();
}

/* -----------------------------------------------------------------
   Auth routes
----------------------------------------------------------------- */
app.post("/api/auth/login", (req, res) => {
  const { username, password } = req.body || {};
  const user = getUsers().find((u) => u.username === username);
  if (!user || !bcrypt.compareSync(password || "", user.passwordHash)) {
    return res.status(401).json({ error: "Invalid username or password" });
  }
  const token = generateToken();
  const sessions = getSessions();
  sessions[token] = {
    username: user.username,
    role: user.role,
    expiresAt: new Date(Date.now() + SESSION_DAYS * 24 * 60 * 60 * 1000).toISOString(),
  };
  setSessions(sessions);
  res.json({ token, username: user.username, role: user.role });
});

app.post("/api/auth/logout", requireAuth, (req, res) => {
  const sessions = getSessions();
  delete sessions[req.token];
  setSessions(sessions);
  res.json({ ok: true });
});

app.get("/api/auth/me", requireAuth, (req, res) => {
  res.json({ username: req.user.username, role: req.user.role });
});

/* -----------------------------------------------------------------
   User management (admin only)
----------------------------------------------------------------- */
app.get("/api/users", requireAuth, requireAdmin, (req, res) => {
  res.json({ users: getUsers().map((u) => ({ username: u.username, role: u.role, createdAt: u.createdAt })) });
});

app.post("/api/users", requireAuth, requireAdmin, (req, res) => {
  const { username, password, role } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: "username and password are required" });
  if (!["admin", "guest"].includes(role)) return res.status(400).json({ error: "role must be admin or guest" });
  const users = getUsers();
  if (users.some((u) => u.username.toLowerCase() === String(username).toLowerCase())) {
    return res.status(409).json({ error: "that username already exists" });
  }
  users.push({ username, passwordHash: bcrypt.hashSync(password, 10), role, createdAt: new Date().toISOString() });
  setUsers(users);
  appendAudit({ username: req.user.username, action: "user-create", summary: `Created ${role} account "${username}"` });
  res.json({ ok: true });
});

app.delete("/api/users/:username", requireAuth, requireAdmin, (req, res) => {
  const users = getUsers();
  const target = users.find((u) => u.username === req.params.username);
  if (!target) return res.status(404).json({ error: "not found" });
  if (target.role === "admin" && users.filter((u) => u.role === "admin").length <= 1) {
    return res.status(400).json({ error: "cannot delete the last admin account" });
  }
  setUsers(users.filter((u) => u.username !== req.params.username));
  const sessions = getSessions();
  Object.keys(sessions).forEach((t) => {
    if (sessions[t].username === req.params.username) delete sessions[t];
  });
  setSessions(sessions);
  appendAudit({ username: req.user.username, action: "user-delete", summary: `Deleted account "${req.params.username}"` });
  res.json({ ok: true });
});

/* -----------------------------------------------------------------
   Audit log — any signed-in user can add an entry (the server fills
   in the username from their session, never trusting the client);
   only admins can read the log back.
----------------------------------------------------------------- */
app.get("/api/audit-log", requireAuth, requireAdmin, (req, res) => {
  res.json({ entries: loadStore().__audit || [] });
});

app.post("/api/audit-log", requireAuth, (req, res) => {
  const { action, summary } = req.body || {};
  appendAudit({ username: req.user.username, action: action || "unknown", summary: summary || "" });
  res.json({ ok: true });
});

/* -----------------------------------------------------------------
   App data key/value store (auth required; reserved "__" keys are
   never reachable here)
----------------------------------------------------------------- */
app.get("/api/kv/:key", requireAuth, (req, res) => {
  if (req.params.key.startsWith("__")) return res.status(403).json({ error: "reserved key" });
  const store = loadStore();
  const value = Object.prototype.hasOwnProperty.call(store, req.params.key) ? store[req.params.key] : null;
  res.json({ value });
});

app.put("/api/kv/:key", requireAuth, (req, res) => {
  if (req.params.key.startsWith("__")) return res.status(403).json({ error: "reserved key" });
  const store = loadStore();
  const value = typeof req.body.value === "string" ? req.body.value : JSON.stringify(req.body.value);
  store[req.params.key] = value;
  saveStore(store);
  res.json({ ok: true });
});

/* -----------------------------------------------------------------
   Atomic inventory adjustments. Every inventory-mutating action in the
   app (inbound/outbound confirm, manual adjust, record deletion, stock
   count) goes through this endpoint instead of overwriting the whole
   inventory array from the client's in-memory copy. That copy can be
   stale — someone else may have confirmed a shipment since this
   browser last loaded — and blindly writing a full stale snapshot back
   would silently erase their change. Here, each operation only touches
   the specific SKU/field it names, computed fresh against the current
   on-disk state inside one synchronous request, so concurrent edits
   from different people never clobber each other.

   Body: { operations: [{ sku, name?, field: "qtyNew"|"qtyReturn",
           mode: "delta"|"set"|"delete", value? }] }
----------------------------------------------------------------- */
app.post("/api/inventory/adjust", requireAuth, (req, res) => {
  const { operations } = req.body || {};
  if (!Array.isArray(operations)) return res.status(400).json({ error: "operations array required" });

  const store = loadStore();
  let inventory = [];
  try {
    inventory = JSON.parse(store.inventory || "[]");
  } catch (e) {
    inventory = [];
  }
  const map = new Map(inventory.map((x) => [x.sku, x]));

  operations.forEach((op) => {
    const sku = op && op.sku;
    if (!sku) return;
    if (op.mode === "delete") {
      map.delete(sku);
      return;
    }
    if (!op.field) return;
    const cur = map.get(sku) || { sku, name: op.name || sku, qtyNew: 0, qtyReturn: 0 };
    if (op.mode === "delta") {
      cur[op.field] = (Number(cur[op.field]) || 0) + (Number(op.value) || 0);
    } else {
      // "set" stores the value as-is — quantities arrive as numbers, text
      // fields like brand arrive as strings; forcing everything through
      // Number() here would silently zero out any non-numeric field.
      cur[op.field] = op.value;
    }
    if (op.name && !cur.name) cur.name = op.name;
    map.set(sku, cur);
  });

  const next = Array.from(map.values());
  store.inventory = JSON.stringify(next);
  saveStore(store);
  res.json({ ok: true, inventory: next });
});

/* -----------------------------------------------------------------
   Same "never overwrite from a possibly-stale client snapshot"
   principle, applied to the other shared collections: inbound/outbound
   record lists, SKU mapping rules, and ignored codes. Each endpoint
   reads the CURRENT on-disk state and applies one precise change,
   instead of the client sending back a full array/object it built
   from data it may have loaded a while ago.
----------------------------------------------------------------- */
function loadArrayKey(key) {
  const store = loadStore();
  try {
    return JSON.parse(store[key] || "[]");
  } catch (e) {
    return [];
  }
}
function saveArrayKey(key, arr) {
  const store = loadStore();
  store[key] = JSON.stringify(arr);
  saveStore(store);
}
function loadObjectKey(key) {
  const store = loadStore();
  try {
    return JSON.parse(store[key] || "{}");
  } catch (e) {
    return {};
  }
}
function saveObjectKey(key, obj) {
  const store = loadStore();
  store[key] = JSON.stringify(obj);
  saveStore(store);
}
function normalizeSkuServer(s) {
  return String(s || "").trim().toUpperCase().replace(/[\s\-_./]/g, "");
}

const RECORD_KEYS = ["inbound-records", "outbound-records"];

app.post("/api/records/:key/add", requireAuth, (req, res) => {
  const key = req.params.key;
  if (!RECORD_KEYS.includes(key)) return res.status(400).json({ error: "invalid key" });
  const { record } = req.body || {};
  if (!record || !record.id) return res.status(400).json({ error: "record with an id is required" });
  const arr = loadArrayKey(key);
  arr.unshift(record);
  saveArrayKey(key, arr);
  res.json({ ok: true, records: arr });
});

app.post("/api/records/:key/remove", requireAuth, (req, res) => {
  const key = req.params.key;
  if (!RECORD_KEYS.includes(key)) return res.status(400).json({ error: "invalid key" });
  const { id } = req.body || {};
  if (!id) return res.status(400).json({ error: "id is required" });
  const arr = loadArrayKey(key).filter((r) => r.id !== id);
  saveArrayKey(key, arr);
  res.json({ ok: true, records: arr });
});

app.post("/api/aliases/set", requireAuth, (req, res) => {
  const { key, value } = req.body || {};
  if (!key || !value) return res.status(400).json({ error: "key and value are required" });
  const obj = loadObjectKey("sku-aliases");
  obj[key] = value;
  saveObjectKey("sku-aliases", obj);
  res.json({ ok: true, aliases: obj });
});

app.post("/api/aliases/delete", requireAuth, (req, res) => {
  const { key } = req.body || {};
  if (!key) return res.status(400).json({ error: "key is required" });
  const obj = loadObjectKey("sku-aliases");
  delete obj[key];
  saveObjectKey("sku-aliases", obj);
  res.json({ ok: true, aliases: obj });
});

app.post("/api/ignored-skus/add", requireAuth, (req, res) => {
  const { value } = req.body || {};
  if (!value) return res.status(400).json({ error: "value is required" });
  const arr = loadArrayKey("ignored-skus");
  if (!arr.some((s) => normalizeSkuServer(s) === normalizeSkuServer(value))) arr.push(value);
  saveArrayKey("ignored-skus", arr);
  res.json({ ok: true, items: arr });
});

app.post("/api/ignored-skus/remove", requireAuth, (req, res) => {
  const { value } = req.body || {};
  const arr = loadArrayKey("ignored-skus").filter((s) => s !== value);
  saveArrayKey("ignored-skus", arr);
  res.json({ ok: true, items: arr });
});

/* -----------------------------------------------------------------
   AI extraction proxy. The API key lives only on the server — the
   browser never sees it.
----------------------------------------------------------------- */
app.post("/api/extract", requireAuth, async (req, res) => {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "ANTHROPIC_API_KEY is not set on the server" });
  }
  try {
    const { model = "claude-sonnet-4-6", max_tokens = 1500, messages } = req.body;
    const upstream = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({ model, max_tokens, messages }),
    });
    const data = await upstream.json();
    res.status(upstream.status).json(data);
  } catch (e) {
    res.status(500).json({ error: e.message || "extraction request failed" });
  }
});

/* -----------------------------------------------------------------
   Serve the built frontend (npm run build -> dist/)

   index.html is served with no-cache so browsers always check for a new
   version on load — otherwise a browser can keep showing a stale cached
   copy of the app indefinitely after a redeploy, even on a fresh visit
   to the same URL. The hashed JS/CSS files in dist/assets/ are safe to
   cache aggressively since their filename changes on every build.
----------------------------------------------------------------- */
const distDir = path.join(__dirname, "dist");
app.use(
  express.static(distDir, {
    index: false, // don't let this serve index.html with default caching — handled explicitly below
    setHeaders: (res, filePath) => {
      if (filePath.endsWith(".html")) {
        res.setHeader("Cache-Control", "no-cache");
      }
    },
  })
);
app.get("*", (req, res) => {
  res.setHeader("Cache-Control", "no-cache");
  res.sendFile(path.join(distDir, "index.html"));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Warehouse app listening on port ${PORT}`);
  console.log(`Data file: ${DATA_FILE}`);
  console.log(`AI extraction: ${process.env.ANTHROPIC_API_KEY ? "configured" : "NOT configured (set ANTHROPIC_API_KEY)"}`);
  console.log(`Accounts: ${getUsers().length} configured`);
});
