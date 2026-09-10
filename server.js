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
ensureAdminSeed();

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
----------------------------------------------------------------- */
const distDir = path.join(__dirname, "dist");
app.use(express.static(distDir));
app.get("*", (req, res) => {
  res.sendFile(path.join(distDir, "index.html"));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Warehouse app listening on port ${PORT}`);
  console.log(`Data file: ${DATA_FILE}`);
  console.log(`AI extraction: ${process.env.ANTHROPIC_API_KEY ? "configured" : "NOT configured (set ANTHROPIC_API_KEY)"}`);
  console.log(`Accounts: ${getUsers().length} configured`);
});
