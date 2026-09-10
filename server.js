import express from "express";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
app.use(express.json({ limit: "30mb" }));

/* -----------------------------------------------------------------
   Storage: a single JSON file holding { key: value } pairs — mirrors
   the simple key/value model the app already used. Good enough for a
   small team's data volume; no database server to install or manage.

   IMPORTANT: on most hosting platforms (Railway, Render, etc.) the
   filesystem is WIPED on every redeploy unless you attach a persistent
   volume. Set DATA_FILE to a path inside that volume in production —
   see README.md for exact steps. Without a volume, your data will
   disappear the next time you deploy a change.
----------------------------------------------------------------- */
// Prefer a Railway-attached volume automatically if one is mounted, so the
// common case (Railway + Volume) needs zero manual path configuration.
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

// NOTE: there is no login/access control here — anyone with the URL can use
// and edit this app's data, same as the shared-storage behavior it had
// inside Claude. See README.md for simple ways to lock it down later
// (platform-level password wall, IP allowlist, or a real login system).

app.get("/api/kv/:key", (req, res) => {
  const store = loadStore();
  const value = Object.prototype.hasOwnProperty.call(store, req.params.key) ? store[req.params.key] : null;
  res.json({ value });
});

app.put("/api/kv/:key", (req, res) => {
  const store = loadStore();
  const value = typeof req.body.value === "string" ? req.body.value : JSON.stringify(req.body.value);
  store[req.params.key] = value;
  saveStore(store);
  res.json({ ok: true });
});

/* -----------------------------------------------------------------
   AI extraction proxy. The API key lives only on the server — the
   browser never sees it. Forwards whatever {model, max_tokens,
   messages} the frontend sends straight to Anthropic's Messages API.
----------------------------------------------------------------- */
app.post("/api/extract", async (req, res) => {
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
});
