import React, { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

/* ---------------------------------------------------------------
   Auth — a simple bearer-token session stored in localStorage.
   authToken lives at module scope so the plain async helpers below
   (kvGet/kvSet/extractFromFile/api*) can attach it without needing
   to be React components themselves.
--------------------------------------------------------------- */
let authToken = null;
function setAuthToken(t) {
  authToken = t;
}
function authHeaders() {
  return authToken ? { Authorization: `Bearer ${authToken}` } : {};
}
function notifyAuthExpired() {
  window.dispatchEvent(new Event("wh-auth-expired"));
}

async function apiGet(path) {
  const res = await fetch(path, { headers: authHeaders() });
  if (res.status === 401) {
    notifyAuthExpired();
    throw new Error("session expired");
  }
  return res.json();
}
async function apiPost(path, body) {
  const res = await fetch(path, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify(body || {}),
  });
  if (res.status === 401) {
    notifyAuthExpired();
    throw new Error("session expired");
  }
  return res.json();
}
async function apiDelete(path) {
  const res = await fetch(path, { method: "DELETE", headers: authHeaders() });
  if (res.status === 401) {
    notifyAuthExpired();
    throw new Error("session expired");
  }
  return res.json();
}

// Fire-and-forget audit trail — failures here should never block the
// action the person actually cares about.
function logAudit(action, summary) {
  apiPost("/api/audit-log", { action, summary }).catch(() => {});
}

/* ---------------------------------------------------------------
   Storage adapter — talks to our own Express + SQLite backend
   instead of the Claude-artifact window.storage API. Same call
   shape (throws on a missing key, mirroring window.storage.get's
   behavior) so the rest of the app needed no further changes.
--------------------------------------------------------------- */
async function kvGet(key) {
  const res = await fetch(`/api/kv/${encodeURIComponent(key)}`, { headers: authHeaders() });
  if (res.status === 401) {
    notifyAuthExpired();
    throw new Error("session expired");
  }
  if (!res.ok) throw new Error(`kv get failed: ${res.status}`);
  const data = await res.json();
  if (data.value == null) throw new Error("not found");
  return data; // { value: "<json string>" }
}

async function kvSet(key, value) {
  const res = await fetch(`/api/kv/${encodeURIComponent(key)}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({ value }),
  });
  if (res.status === 401) {
    notifyAuthExpired();
    return null;
  }
  if (!res.ok) return null;
  return { ok: true };
}
import {
  Package,
  PackagePlus,
  PackageMinus,
  ArrowLeft,
  Upload,
  Loader2,
  CheckCircle2,
  AlertTriangle,
  X,
  Plus,
  Pencil,
  Download,
  FileText,
  ChevronLeft,
  ChevronRight,
  Trash2,
  Link2,
  RotateCcw,
  Users,
  ShoppingCart,
  Clock,
  ShieldCheck,
  History,
  Search,
} from "lucide-react";

/* ---------------------------------------------------------------
   Design tokens — steel-navy warehouse ledger, colour-coded flows
--------------------------------------------------------------- */
const C = {
  bg: "#EDEEE9",
  bgHeader: "#1F2A3D",
  headerInk: "#EDEEE9",
  surface: "#FFFFFF",
  surfaceSoft: "#F5F5F0",
  border: "#DAD9D0",
  borderStrong: "#C7C6BB",
  ink: "#1E2430",
  inkSoft: "#6B7080",
  inventory: "#2C3A56",
  inboundNew: "#2F7A52",
  inboundNewSoft: "#E4F0E8",
  inboundReturn: "#3A6EA5",
  inboundReturnSoft: "#E2EAF3",
  outbound: "#B14A2E",
  outboundSoft: "#F3E4DE",
  outboundOrder: "#2E7D7B",
  outboundOrderSoft: "#DCEEEC",
  amber: "#C98A2C",
  amberSoft: "#F5E9D6",
  danger: "#C0392B",
};
const FONT_UI = '-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif';
const FONT_MONO = 'ui-monospace,SFMono-Regular,"SF Mono",Menlo,Consolas,monospace';

/* ---------------------------------------------------------------
   Helpers
--------------------------------------------------------------- */
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result.split(",")[1]);
    r.onerror = () => reject(new Error("Failed to read file"));
    r.readAsDataURL(file);
  });
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function monthKey(dateStr) {
  if (!dateStr) return null;
  return String(dateStr).slice(0, 7);
}

function fmtMoney(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function downloadCSV(rows, filename) {
  const csv = rows
    .map((r) => r.map((c) => `"${String(c ?? "").replace(/"/g, '""')}"`).join(","))
    .join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function normalizeSku(s) {
  return String(s || "")
    .trim()
    .toUpperCase()
    .replace(/[\s\-_./]/g, "");
}

function isSpreadsheetFile(file) {
  const name = (file.name || "").toLowerCase();
  return (
    name.endsWith(".xlsx") ||
    name.endsWith(".xls") ||
    name.endsWith(".csv") ||
    (file.type && file.type.includes("spreadsheet")) ||
    file.type === "text/csv"
  );
}

async function parseSpreadsheetFile(file) {
  const name = (file.name || "").toLowerCase();
  let wb;
  if (name.endsWith(".csv") || file.type === "text/csv") {
    const text = await file.text();
    wb = XLSX.read(text, { type: "string" });
  } else {
    const buf = await file.arrayBuffer();
    wb = XLSX.read(buf, { type: "array" });
  }
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }); // raw array-of-arrays
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" }); // object rows, first row as header
  const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
  return { grid, rows, headers };
}

// Scans a raw grid for one or more "SKU ... quantity" tables, wherever they sit
// in the sheet — handles files like Amazon's FBA shipment-plan export, which
// stack a metadata block plus several sub-tables (e.g. "Case packed" /
// "Individual units") with different column layouts in a single sheet.
function findSkuSections(grid) {
  const sections = [];
  let i = 0;
  while (i < grid.length) {
    const row = grid[i] || [];
    const lower = row.map((c) => String(c).trim().toLowerCase());
    const skuColIdx = lower.findIndex((c) => c === "sku");
    if (skuColIdx >= 0) {
      let qtyColIdx = lower.findIndex((c) => c === "total units");
      if (qtyColIdx < 0) qtyColIdx = lower.findIndex((c) => c.includes("total") && c.includes("unit"));
      if (qtyColIdx < 0) qtyColIdx = lower.findIndex((c) => c === "units" || c === "qty" || c === "quantity");
      if (qtyColIdx < 0) qtyColIdx = lower.findIndex((c) => c.includes("qty") || c.includes("quantity") || c.includes("pcs"));
      const nameColIdx = lower.findIndex((c) => c.includes("title") || c.includes("name") || c.includes("descri"));
      if (qtyColIdx >= 0) {
        let j = i + 1;
        const dataRows = [];
        while (j < grid.length) {
          const r = grid[j] || [];
          const isBlank = r.every((c) => String(c).trim() === "");
          if (isBlank) break;
          const rLower = r.map((c) => String(c).trim().toLowerCase());
          if (rLower.includes("sku")) break; // next section's header starts here
          dataRows.push(r);
          j++;
        }
        sections.push({ skuColIdx, qtyColIdx, nameColIdx, dataRows });
        i = j;
        continue;
      }
    }
    i++;
  }
  return sections;
}

// Aggregates every detected section into one items[] list, summing quantities
// when the same SKU appears more than once (e.g. across sections).
function itemsFromSkuSections(sections) {
  const map = new Map();
  sections.forEach(({ skuColIdx, qtyColIdx, nameColIdx, dataRows }) => {
    dataRows.forEach((r) => {
      const sku = String(r[skuColIdx] ?? "").trim();
      const qty = Number(r[qtyColIdx]) || 0;
      if (!sku || qty <= 0) return;
      const name = nameColIdx >= 0 ? String(r[nameColIdx] ?? "").trim() : "";
      const cur = map.get(sku) || { sku, name: "", qty: 0, unitPrice: 0 };
      cur.qty += qty;
      if (!cur.name && name) cur.name = name;
      map.set(sku, cur);
    });
  });
  return Array.from(map.values());
}

// Reads the metadata block at the top of an Amazon FBA shipment-plan export
// (Shipment ID, Shipment name, Ship to, Boxes, SKUs, Units) — order and row
// position don't matter, it just scans column A for known labels.
function parseFbaMeta(grid) {
  const meta = {};
  grid.forEach((row) => {
    const label = String(row[0] || "").trim().toLowerCase();
    const value = row[1];
    if (label === "shipment id") meta.shipmentId = String(value ?? "").trim();
    else if (label === "shipment name") meta.shipmentName = String(value ?? "").trim();
    else if (label === "ship to") meta.shipTo = String(value ?? "").trim();
    else if (label === "boxes") meta.boxes = Number(value) || 0;
    else if (label === "skus") meta.skus = Number(value) || 0;
    else if (label === "units") meta.units = Number(value) || 0;
  });
  return meta;
}

// --- Filename conventions -------------------------------------------------
// Inbound invoices are named after the order/invoice number itself.
function orderIdFromFilename(fileName) {
  return String(fileName || "").replace(/\.[^.]+$/, "").trim();
}

// Amazon FBA/order-fulfillment source files are named with just the ship
// date, e.g. "2026-09-10.html".
function dateFromFilename(fileName) {
  const base = String(fileName || "").replace(/\.[^.]+$/, "").trim();
  const m = base.match(/^(\d{4})[-_.](\d{1,2})[-_.](\d{1,2})$/);
  if (!m) return null;
  const y = m[1], mo = m[2].padStart(2, "0"), d = m[3].padStart(2, "0");
  return `${y}-${mo}-${d}`;
}

// Other-platform screenshots are named "{Platform} {YYYY-MM-DD}", e.g.
// "eBay 2026-09-10.png" — returns {platform, date} or null if it doesn't match.
function parsePlatformDateFilename(fileName) {
  const base = String(fileName || "").replace(/\.[^.]+$/, "").trim();
  const m = base.match(/^(.+?)\s+(\d{4})[-_.](\d{1,2})[-_.](\d{1,2})$/);
  if (!m) return null;
  const y = m[2], mo = m[3].padStart(2, "0"), d = m[4].padStart(2, "0");
  return { platform: m[1].trim(), date: `${y}-${mo}-${d}` };
}

// --- Amazon Seller Central packing-slip HTML (saved page) ---------------
function isAmazonPackingSlipFile(file) {
  const name = (file.name || "").toLowerCase();
  return name.endsWith(".html") || name.endsWith(".htm") || file.type === "text/html";
}

async function parseAmazonPackingSlip(file) {
  const text = await file.text();
  const doc = new DOMParser().parseFromString(text, "text/html");
  const norm = (s) => (s || "").replace(/\s+/g, " ").trim();
  const orderIdPattern = /\d{3}-\d{7}-\d{7}/;
  const rows = [];
  let currentOrderId = null;
  // Walk order-ID markers and item rows together, in document order, so each
  // row picks up the order it actually belongs to (Amazon puts the order ID
  // in its own <div class="a-section myo-orderId"> just above that order's table).
  doc.querySelectorAll("div.myo-orderId, tr").forEach((el) => {
    if (el.tagName === "DIV") {
      const m = norm(el.textContent).match(orderIdPattern);
      if (m) currentOrderId = m[0];
      return;
    }
    const detailsTd = el.querySelector("td.myo-mena-product-details");
    if (!detailsTd) return;
    const fields = {};
    detailsTd.querySelectorAll("div.a-row").forEach((row) => {
      const spans = Array.from(row.querySelectorAll("span")).map((s) => norm(s.textContent));
      if (spans.length >= 2 && /:$/.test(spans[0])) {
        fields[spans[0].replace(/:$/, "")] = spans[1];
      }
    });
    if (!fields["SKU"]) return;
    const nameSpan = detailsTd.querySelector("span.a-text-bold");
    const name = norm(nameSpan ? nameSpan.textContent : "");
    const tds = el.querySelectorAll(":scope > td");
    const qty = Number(norm(tds[0] ? tds[0].textContent : "")) || 0;
    const priceText = tds.length > 2 ? norm(tds[2].textContent) : "";
    const unitPrice = parseFloat(priceText.replace(/[^\d,.-]/g, "").replace(",", ".")) || 0;
    if (!fields["SKU"] || qty <= 0) return;
    rows.push({ sku: fields["SKU"], name, qty, unitPrice, orderId: currentOrderId });
  });
  // Deliberately NOT aggregated by SKU — keeping one row per order line lets
  // a later SKU search show exactly which order shipped how much, even when
  // several orders in the same file ship the same product.
  return rows;
}

function guessColumn(headers, keywords) {
  const lower = headers.map((h) => String(h).toLowerCase());
  for (const kw of keywords) {
    const idx = lower.findIndex((h) => h.includes(kw));
    if (idx >= 0) return headers[idx];
  }
  return "";
}

// A learned alias can map a raw SKU to ONE component (the common "multi-pack of
// the same item" case) or to SEVERAL different components (a bundle/kit made of
// different SKUs). Always normalize to a components[] array so downstream code
// doesn't need to special-case the single-target shape.
function aliasComponents(alias) {
  if (alias.components) return alias.components;
  return [{ sku: alias.sku, qty: Number(alias.qtyMultiplier) || 1 }];
}

// Resolves raw SKU strings against known inventory + the learned alias table.
// A learned alias carries one or more components, each with its own quantity
// multiplier — e.g. "MR100-2" → MR100 ×2 (multi-pack), or "COMBO-AB" → A ×2 + B ×1
// (a bundle of different products). Rows matching an existing SKU pass through
// untouched (×1). Rows matching a learned alias expand into one resolved item
// per component. Unknown rows are flagged with needsMapping so the person can
// teach the system what they mean. Rows on the ignore list (shipping fees,
// service charges that aren't real inventory) are dropped entirely first.
function resolveSkus(items, inventory, aliasMap, ignoredSkus = []) {
  const ignoredSet = new Set(ignoredSkus.map(normalizeSku));
  const invSkus = new Set(inventory.map((x) => x.sku));
  const result = [];
  items
    .filter((it) => !ignoredSet.has(normalizeSku(it.sku)))
    .forEach((it) => {
      const raw = it.sku;
      const rawQty = it.qty;
      const norm = normalizeSku(raw);
      const alias = aliasMap[norm];
      // A taught rule always wins — even when the raw text happens to be
      // identical to an existing SKU (e.g. a supplier reuses the base SKU on
      // their invoice to mean "a 4-pack of it"; the rule is what disambiguates).
      if (alias) {
        aliasComponents(alias).forEach((comp) => {
          const mult = Number(comp.qty) || 1;
          result.push({
            ...it,
            sku: comp.sku,
            name: it.name,
            unitPrice: it.unitPrice,
            rawSku: raw,
            rawQty,
            qty: rawQty * mult,
            qtyMultiplier: mult,
            mappedFrom: raw,
            needsMapping: false,
            autoMapped: true,
          });
        });
        return;
      }
      if (invSkus.has(raw)) {
        result.push({ ...it, rawQty, qtyMultiplier: 1, needsMapping: false, autoMapped: false });
        return;
      }
      result.push({ ...it, rawSku: raw, rawQty, qtyMultiplier: 1, needsMapping: !!raw });
    });
  return result;
}

async function extractFromFile(file, mode) {
  const base64 = await fileToBase64(file);
  const isPdf = file.type === "application/pdf";
  const mediaType = file.type || (isPdf ? "application/pdf" : "image/jpeg");
  const contentBlock = isPdf
    ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } }
    : { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } };

  const promptInbound = `You are a warehouse assistant. Carefully read this purchase or stock-return document (image or PDF) and extract the following information. Return ONLY a JSON object — no other text, explanation, or markdown code fences. Format:
{
  "supplier": "supplier or source name, null if not found",
  "invoice_date": "YYYY-MM-DD, null if not found",
  "items": [
    {"sku_guess": "product SKU or model number; if not found, create a short code from the product name", "name": "product name", "qty": integer quantity, "unit_price": unit price as a plain number with no currency symbol, 0 if not found}
  ],
  "total_amount": total amount as a number, or the sum of item amounts if not stated
}
Cover every line item on the document as completely as possible. Do not include shipping charges, service fees, taxes, or other non-product lines — only physical products. Some invoices have a warehouse location / shelf / bin column (often labeled "Place", "Ubicazione", "Location", or similar) sitting right next to the actual product code column (often labeled "Codice", "SKU", "Cod.", "Item code", or similar) — these are two separate pieces of information, even when they appear visually adjacent or on the same line. Use ONLY the product code column for sku_guess; never prepend, append, or merge a location/shelf/bin code into the SKU.`;

  const promptOutbound = `You are a warehouse assistant. Carefully read this shipping / outbound document (courier label, platform packing list, shipment manifest, etc. — image or PDF) and extract the following information. Return ONLY a JSON object — no other text, explanation, or markdown code fences. Format:
{
  "platform": "shipping platform or customer name, null if not found",
  "ship_date": "YYYY-MM-DD, null if not found",
  "items": [
    {"sku_guess": "product SKU or model number; if not found, create a short code from the product name", "name": "product name", "qty": integer quantity}
  ]
}
Cover every line item on the document as completely as possible. Do not include shipping charges, service fees, taxes, or other non-product lines — only physical products.`;

  const res = await fetch("/api/extract", {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1500,
      messages: [
        {
          role: "user",
          content: [contentBlock, { type: "text", text: mode === "inbound" ? promptInbound : promptOutbound }],
        },
      ],
    }),
  });
  if (res.status === 401) {
    notifyAuthExpired();
    throw new Error("session expired");
  }
  if (!res.ok) throw new Error("AI recognition request failed");
  const data = await res.json();
  const textBlocks = (data.content || [])
    .filter((b) => b.type === "text")
    .map((b) => b.text)
    .join("\n");
  const cleaned = textBlocks.replace(/```json|```/g, "").trim();
  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    throw new Error("Couldn't parse the AI response — please enter items manually");
  }
  if (!parsed.items || !Array.isArray(parsed.items)) parsed.items = [];
  return parsed;
}

/* ---------------------------------------------------------------
   Small shared UI atoms
--------------------------------------------------------------- */
function Btn({ children, onClick, color = C.inventory, variant = "solid", disabled, style, icon: Icon }) {
  const solid = variant === "solid";
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
        padding: "9px 16px",
        fontSize: 14,
        fontWeight: 600,
        fontFamily: FONT_UI,
        borderRadius: 4,
        border: `1.5px solid ${color}`,
        background: solid ? color : "transparent",
        color: solid ? "#fff" : color,
        cursor: disabled ? "not-allowed" : "pointer",
        opacity: disabled ? 0.5 : 1,
        transition: "filter 0.15s ease",
        ...style,
      }}
      onMouseEnter={(e) => !disabled && (e.currentTarget.style.filter = "brightness(1.08)")}
      onMouseLeave={(e) => (e.currentTarget.style.filter = "none")}
    >
      {Icon && <Icon size={15} strokeWidth={2.2} />}
      {children}
    </button>
  );
}

function Stamp({ label, color, bg }) {
  return (
    <span
      style={{
        display: "inline-block",
        fontFamily: FONT_MONO,
        fontSize: 11,
        letterSpacing: "0.06em",
        fontWeight: 700,
        color,
        background: bg,
        border: `1.5px dashed ${color}`,
        borderRadius: 3,
        padding: "2px 8px",
        transform: "rotate(-2deg)",
      }}
    >
      {label}
    </span>
  );
}

function Toast({ toast }) {
  if (!toast) return null;
  const isErr = toast.type === "error";
  return (
    <div
      style={{
        position: "fixed",
        bottom: 20,
        left: "50%",
        transform: "translateX(-50%)",
        background: isErr ? C.danger : C.inboundNew,
        color: "#fff",
        padding: "10px 18px",
        borderRadius: 5,
        fontSize: 13.5,
        fontFamily: FONT_UI,
        fontWeight: 600,
        display: "flex",
        alignItems: "center",
        gap: 8,
        boxShadow: "0 6px 20px rgba(0,0,0,0.18)",
        zIndex: 999,
        maxWidth: "90vw",
      }}
    >
      {isErr ? <AlertTriangle size={16} /> : <CheckCircle2 size={16} />}
      {toast.msg}
    </div>
  );
}

function TopBar({ title, subtitle, onBack, accent }) {
  return (
    <div
      style={{
        background: C.bgHeader,
        color: C.headerInk,
        padding: "18px 24px",
        display: "flex",
        alignItems: "center",
        gap: 14,
        borderBottom: `4px solid ${accent}`,
      }}
    >
      {onBack && (
        <button
          onClick={onBack}
          style={{
            background: "rgba(255,255,255,0.08)",
            border: "1px solid rgba(255,255,255,0.25)",
            borderRadius: 4,
            color: C.headerInk,
            padding: "7px 9px",
            cursor: "pointer",
            display: "flex",
          }}
        >
          <ArrowLeft size={16} />
        </button>
      )}
      <div>
        <div style={{ fontSize: 17, fontWeight: 700, fontFamily: FONT_UI, letterSpacing: "0.01em" }}>{title}</div>
        {subtitle && <div style={{ fontSize: 12, color: "rgba(237,238,233,0.65)", marginTop: 2, fontFamily: FONT_UI }}>{subtitle}</div>}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Upload box (shared by inbound / outbound)
--------------------------------------------------------------- */
function UploadBox({ accent, accentSoft, label, hint, onFile, busy }) {
  const inputRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);

  const handleFiles = (files) => {
    if (files && files[0]) onFile(files[0]);
  };

  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        handleFiles(e.dataTransfer.files);
      }}
      onClick={() => !busy && inputRef.current?.click()}
      style={{
        border: `2px dashed ${dragOver ? accent : C.borderStrong}`,
        background: dragOver ? accentSoft : C.surface,
        borderRadius: 6,
        padding: "36px 20px",
        textAlign: "center",
        cursor: busy ? "default" : "pointer",
        transition: "all 0.15s ease",
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept="image/*,application/pdf,.xlsx,.xls,.csv,.html,.htm"
        style={{ display: "none" }}
        onChange={(e) => handleFiles(e.target.files)}
      />
      {busy ? (
        <>
          <Loader2 size={26} color={accent} style={{ animation: "spin 1s linear infinite" }} />
          <div style={{ marginTop: 10, fontSize: 13.5, color: C.inkSoft, fontFamily: FONT_UI }}>Reading the document…</div>
        </>
      ) : (
        <>
          <Upload size={26} color={accent} />
          <div style={{ marginTop: 10, fontSize: 14.5, fontWeight: 600, color: C.ink, fontFamily: FONT_UI }}>{label}</div>
          <div style={{ marginTop: 4, fontSize: 12.5, color: C.inkSoft, fontFamily: FONT_UI }}>{hint}</div>
        </>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   HOME
--------------------------------------------------------------- */
function LoginScreen({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    if (!username.trim() || !password) return;
    setBusy(true);
    setError("");
    try {
      const res = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username: username.trim(), password }),
      });
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Login failed");
        return;
      }
      onLogin(data);
    } catch (e) {
      setError("Network error — please try again");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: C.bg, fontFamily: FONT_UI }}>
      <form
        onSubmit={submit}
        style={{
          background: C.surface,
          padding: "32px 30px",
          borderRadius: 6,
          width: 320,
          maxWidth: "90vw",
          borderTop: `4px solid ${C.inventory}`,
          boxShadow: "0 10px 30px rgba(0,0,0,0.10)",
        }}
      >
        <div style={{ fontSize: 12.5, letterSpacing: "0.08em", color: C.amber, fontWeight: 700, fontFamily: FONT_MONO, marginBottom: 4 }}>
          A5 STAR
        </div>
        <div style={{ fontSize: 19, fontWeight: 700, color: C.ink, marginBottom: 22 }}>Sign in</div>
        <Field label="Username">
          <input value={username} onChange={(e) => setUsername(e.target.value)} style={inputStyle} autoFocus />
        </Field>
        <Field label="Password">
          <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} style={inputStyle} />
        </Field>
        {error && (
          <div style={{ color: C.danger, fontSize: 12.5, fontFamily: FONT_UI, marginBottom: 12, display: "flex", gap: 6, alignItems: "center" }}>
            <AlertTriangle size={13} />
            {error}
          </div>
        )}
        <button
          type="submit"
          disabled={busy}
          style={{
            width: "100%",
            marginTop: 4,
            padding: "10px 16px",
            borderRadius: 4,
            border: "none",
            background: C.inventory,
            color: "#fff",
            fontWeight: 700,
            fontSize: 14,
            fontFamily: FONT_UI,
            cursor: busy ? "not-allowed" : "pointer",
            opacity: busy ? 0.6 : 1,
          }}
        >
          {busy ? "Signing in…" : "Sign in"}
        </button>
      </form>
    </div>
  );
}

function Home({ setView, inventory, inboundRecords, outboundRecords, auth, onLogout }) {
  const totalNew = inventory.reduce((s, x) => s + (x.qtyNew || 0), 0);
  const totalReturn = inventory.reduce((s, x) => s + (x.qtyReturn || 0), 0);
  const thisMonth = new Date().toISOString().slice(0, 7);
  const inboundThisMonth = inboundRecords.filter((r) => monthKey(r.invoiceDate || r.date) === thisMonth).length;
  const outboundThisMonth = outboundRecords.filter((r) => monthKey(r.shipDate || r.date) === thisMonth).length;

  const cards = [
    {
      key: "inventory",
      title: "Inventory",
      desc: "View new-stock and return-stock balances",
      accent: C.inventory,
      icon: Package,
      stat: `${totalNew + totalReturn} units in stock`,
    },
    {
      key: "inbound",
      title: "Inbound",
      desc: "Log new purchases or returned stock coming in",
      accent: C.inboundNew,
      icon: PackagePlus,
      stat: `${inboundThisMonth} records this month`,
    },
    {
      key: "outbound",
      title: "Outbound",
      desc: "Upload shipping documents to log stock going out",
      accent: C.outbound,
      icon: PackageMinus,
      stat: `${outboundThisMonth} shipments this month`,
    },
  ];
  if (auth?.role === "admin") {
    cards.push({
      key: "admin",
      title: "Admin",
      desc: "Manage team accounts and view the audit log",
      accent: C.amber,
      icon: ShieldCheck,
      stat: "Admin only",
    });
  }

  return (
    <div>
      <div style={{ background: C.bgHeader, color: C.headerInk, padding: "26px 24px 22px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div style={{ fontSize: 13, letterSpacing: "0.08em", color: C.amber, fontWeight: 700, fontFamily: FONT_MONO }}>
            A5 STAR · WAREHOUSE LEDGER
          </div>
          <div style={{ fontSize: 12, color: "rgba(237,238,233,0.75)", fontFamily: FONT_UI, textAlign: "right" }}>
            <div>
              {auth?.username} <span style={{ color: "rgba(237,238,233,0.5)" }}>({auth?.role})</span>
            </div>
            <button
              onClick={onLogout}
              style={{ background: "none", border: "none", color: C.amber, cursor: "pointer", fontSize: 12, fontFamily: FONT_UI, padding: 0, marginTop: 2 }}
            >
              Log out
            </button>
          </div>
        </div>
        <div style={{ fontSize: 22, fontWeight: 700, marginTop: 6, fontFamily: FONT_UI }}>Inbound / Outbound Management</div>
        <div style={{ fontSize: 12.5, color: "rgba(237,238,233,0.6)", marginTop: 4, fontFamily: FONT_MONO, display: "flex", alignItems: "center", gap: 6 }}>
          {todayISO()}
          <span style={{ opacity: 0.5 }}>·</span>
          <Users size={12} />
          Shared with your team
        </div>
      </div>

      <div
        style={{
          padding: 20,
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
          gap: 14,
        }}
      >
        {cards.map((c) => (
          <div
            key={c.key}
            onClick={() => setView(c.key)}
            style={{
              background: C.surface,
              borderRadius: 5,
              border: `1px solid ${C.border}`,
              borderLeft: `5px solid ${c.accent}`,
              padding: "20px 18px",
              cursor: "pointer",
              transition: "transform 0.12s ease, border-color 0.12s ease",
            }}
            onMouseEnter={(e) => (e.currentTarget.style.transform = "translateY(-2px)")}
            onMouseLeave={(e) => (e.currentTarget.style.transform = "none")}
          >
            <c.icon size={22} color={c.accent} strokeWidth={2} />
            <div style={{ fontSize: 17, fontWeight: 700, marginTop: 12, color: C.ink, fontFamily: FONT_UI }}>{c.title}</div>
            <div style={{ fontSize: 12.5, color: C.inkSoft, marginTop: 4, fontFamily: FONT_UI, lineHeight: 1.5 }}>{c.desc}</div>
            <div style={{ marginTop: 14, fontSize: 12.5, fontFamily: FONT_MONO, fontWeight: 700, color: c.accent }}>{c.stat}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   INVENTORY VIEW
--------------------------------------------------------------- */
function InventoryView({ setView, inventory, saveInventory, showToast, aliasMap, saveAliasMap, ignoredSkus, saveIgnoredSkus, inboundRecords, outboundRecords, isAdmin }) {
  const [query, setQuery] = useState("");
  const [editing, setEditing] = useState(null); // sku or 'new'
  const [form, setForm] = useState({ sku: "", name: "", qtyNew: 0, qtyReturn: 0 });
  const [showAliases, setShowAliases] = useState(false);
  const [showIgnored, setShowIgnored] = useState(false);
  const [showInsights, setShowInsights] = useState(false);

  const filtered = inventory.filter(
    (x) => x.sku.toLowerCase().includes(query.toLowerCase()) || x.name.toLowerCase().includes(query.toLowerCase())
  );
  const totalNew = inventory.reduce((s, x) => s + (x.qtyNew || 0), 0);
  const totalReturn = inventory.reduce((s, x) => s + (x.qtyReturn || 0), 0);

  const openNew = () => {
    setForm({ sku: "", name: "", qtyNew: 0, qtyReturn: 0 });
    setEditing("new");
  };
  const openEdit = (item) => {
    setForm({ ...item });
    setEditing(item.sku);
  };

  const submitForm = () => {
    if (!form.sku.trim() || !form.name.trim()) {
      showToast("SKU and product name are required", "error");
      return;
    }
    const next = [...inventory];
    const idx = next.findIndex((x) => x.sku === form.sku.trim());
    const record = {
      sku: form.sku.trim(),
      name: form.name.trim(),
      qtyNew: Number(form.qtyNew) || 0,
      qtyReturn: Number(form.qtyReturn) || 0,
    };
    if (idx >= 0) next[idx] = record;
    else next.push(record);
    saveInventory(next);
    logAudit("inventory-adjust", `Set ${record.sku} to New=${record.qtyNew}, Return=${record.qtyReturn}`);
    setEditing(null);
    showToast("Inventory updated");
  };

  const removeItem = (sku) => {
    saveInventory(inventory.filter((x) => x.sku !== sku));
    logAudit("inventory-delete", `Deleted inventory item ${sku}`);
    showToast("Item deleted");
  };

  const Section = ({ label, color, field }) => (
    <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 5, overflow: "hidden" }}>
      <div
        style={{
          padding: "12px 16px",
          borderBottom: `1px solid ${C.border}`,
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          background: color === C.inboundNew ? C.inboundNewSoft : C.amberSoft,
        }}
      >
        <Stamp label={label} color={color} bg={C.surface} />
        <span style={{ fontFamily: FONT_MONO, fontWeight: 700, fontSize: 16, color: C.ink }}>
          {inventory.reduce((s, x) => s + (x[field] || 0), 0)}
        </span>
      </div>
      <div style={{ maxHeight: 560, overflowY: "auto" }}>
        {filtered.filter((x) => (x[field] || 0) !== 0).length === 0 && (
          <div style={{ padding: 18, fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI }}>No items with stock</div>
        )}
        {filtered
          .filter((x) => (x[field] || 0) !== 0)
          .map((x) => (
          <div
            key={x.sku}
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              padding: "9px 16px",
              borderBottom: `1px solid ${C.surfaceSoft}`,
              fontSize: 13,
            }}
          >
            <div style={{ minWidth: 0 }}>
              <div style={{ fontFamily: FONT_MONO, fontWeight: 700, color: C.ink }}>{x.sku}</div>
              <div style={{ color: C.inkSoft, fontFamily: FONT_UI, fontSize: 12, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                {x.name}
              </div>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
              <span style={{ fontFamily: FONT_MONO, fontWeight: 700, fontSize: 15, color: x[field] < 0 ? C.danger : C.ink }}>{x[field]}</span>
              <button onClick={() => openEdit(x)} style={{ border: "none", background: "none", cursor: "pointer", color: C.inkSoft, display: "flex" }}>
                <Pencil size={13} />
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );

  return (
    <div>
      <TopBar title="Inventory" subtitle="New Stock (A) · Return Stock (B)" onBack={() => setView("home")} accent={C.inventory} />
      <div style={{ padding: 18 }}>
        <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search SKU or product name…"
            style={{
              flex: 1,
              minWidth: 180,
              padding: "9px 12px",
              borderRadius: 4,
              border: `1px solid ${C.border}`,
              fontSize: 13.5,
              fontFamily: FONT_UI,
              outline: "none",
            }}
          />
          <Btn onClick={openNew} color={C.inventory} icon={Plus}>
            Add / Adjust
          </Btn>
          <Btn onClick={() => setShowAliases(true)} color={C.inkSoft} variant="outline" icon={Link2}>
            SKU Mapping Rules
          </Btn>
          <Btn onClick={() => setShowIgnored(true)} color={C.inkSoft} variant="outline" icon={X}>
            Ignored Codes
          </Btn>
          {isAdmin && (
            <Btn onClick={() => setShowInsights(true)} color={C.inventory} variant="outline" icon={ShoppingCart}>
              Restock Insights
            </Btn>
          )}
          <Btn
            onClick={() => {
              const rows = [
                ["SKU", "Product name", "New stock (A)", "Return stock (B)"],
                ...inventory.map((x) => [x.sku, x.name, x.qtyNew, x.qtyReturn]),
              ];
              downloadCSV(rows, `inventory-${todayISO()}.csv`);
            }}
            color={C.inkSoft}
            variant="outline"
            icon={Download}
          >
            Export CSV
          </Btn>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 14 }}>
          <Section label="New Stock · A" color={C.inboundNew} field="qtyNew" />
          <Section label="Return Stock · B" color={C.amber} field="qtyReturn" />
        </div>

        <div style={{ marginTop: 14, fontSize: 12, color: C.inkSoft, fontFamily: FONT_MONO }}>
          Total: {totalNew} new · {totalReturn} return · {inventory.length} SKUs
        </div>
      </div>

      {editing && (
        <Modal onClose={() => setEditing(null)} title={editing === "new" ? "Add inventory item" : "Adjust inventory"} accent={C.inventory}>
          <Field label="SKU">
            <input
              value={form.sku}
              disabled={editing !== "new"}
              onChange={(e) => setForm({ ...form, sku: e.target.value })}
              style={inputStyle}
            />
          </Field>
          <Field label="Product name">
            <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} style={inputStyle} />
          </Field>
          <div style={{ display: "flex", gap: 10 }}>
            <Field label="New qty" style={{ flex: 1 }}>
              <input type="number" value={form.qtyNew} onChange={(e) => setForm({ ...form, qtyNew: e.target.value })} style={inputStyle} />
            </Field>
            <Field label="Return qty" style={{ flex: 1 }}>
              <input type="number" value={form.qtyReturn} onChange={(e) => setForm({ ...form, qtyReturn: e.target.value })} style={inputStyle} />
            </Field>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 16 }}>
            {editing !== "new" ? (
              <Btn color={C.danger} variant="outline" icon={Trash2} onClick={() => { removeItem(editing); setEditing(null); }}>
                Delete item
              </Btn>
            ) : <span />}
            <Btn color={C.inventory} onClick={submitForm}>Save</Btn>
          </div>
        </Modal>
      )}

      {showAliases && (
        <AliasManager aliasMap={aliasMap} saveAliasMap={saveAliasMap} inventory={inventory} isAdmin={isAdmin} onClose={() => setShowAliases(false)} />
      )}
      {showIgnored && <IgnoredSkuManager ignoredSkus={ignoredSkus} saveIgnoredSkus={saveIgnoredSkus} onClose={() => setShowIgnored(false)} />}
      {showInsights && (
        <InsightsModal inventory={inventory} inboundRecords={inboundRecords} outboundRecords={outboundRecords} onClose={() => setShowInsights(false)} />
      )}
    </div>
  );
}

function AliasManager({ aliasMap, saveAliasMap, inventory, isAdmin, onClose }) {
  const entries = Object.entries(aliasMap || {});
  const [showAdd, setShowAdd] = useState(false);
  const [rawInput, setRawInput] = useState("");
  const [components, setComponents] = useState([{ sku: "", qty: 1 }]);

  const remove = (key) => {
    const raw = aliasMap[key]?.raw || key;
    const next = { ...aliasMap };
    delete next[key];
    saveAliasMap(next);
    logAudit("alias-delete", `Removed SKU mapping rule for "${raw}"`);
  };
  const exportCSV = () => {
    const rows = [["Raw SKU", "Maps to SKU", "Qty"]];
    entries.forEach(([, v]) => {
      aliasComponents(v).forEach((c) => rows.push([v.raw, c.sku, c.qty]));
    });
    downloadCSV(rows, `sku-mapping-rules-${todayISO()}.csv`);
  };

  const updateComponent = (ci, field, val) => setComponents(components.map((c, i) => (i === ci ? { ...c, [field]: val } : c)));
  const addComponent = () => setComponents([...components, { sku: "", qty: 1 }]);
  const removeComponent = (ci) => setComponents(components.filter((_, i) => i !== ci));

  const saveRule = () => {
    const raw = rawInput.trim();
    if (!raw) return;
    const comps = components.map((c) => ({ sku: String(c.sku || "").trim(), qty: Number(c.qty) || 1 })).filter((c) => c.sku);
    if (comps.length === 0) return;
    saveAliasMap({ ...aliasMap, [normalizeSku(raw)]: { raw, components: comps } });
    logAudit("alias-add", `Added SKU mapping rule: "${raw}" → ${comps.map((c) => `${c.sku} ×${c.qty}`).join(", ")}`);
    setRawInput("");
    setComponents([{ sku: "", qty: 1 }]);
    setShowAdd(false);
  };

  return (
    <Modal title="SKU Mapping Rules" accent={C.inventory} onClose={onClose} wide>
      <datalist id="wh-sku-datalist-alias">
        {(inventory || []).map((inv) => (
          <option key={inv.sku} value={inv.sku}>
            {inv.name}
          </option>
        ))}
      </datalist>
      <div style={{ fontSize: 12.5, color: C.inkSoft, fontFamily: FONT_UI, marginBottom: 12 }}>
        Rules learned from your inbound/outbound documents. The same raw SKU is applied automatically next time —
        even if the raw text happens to match an existing SKU verbatim (e.g. a supplier reuses the base code to mean
        "a 4-pack of it" — the rule is what disambiguates). A rule can map to one product at a quantity, or several
        different products at once for a bundle.
      </div>
      {entries.length === 0 ? (
        <div style={{ fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI, padding: "8px 0" }}>
          No mapping rules yet. Rules are saved automatically when you resolve an unrecognized SKU during inbound or
          outbound — or add one manually below for cases the system can't flag on its own (like a raw SKU that's
          identical to an existing one).
        </div>
      ) : (
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 5, overflow: "hidden", marginBottom: 14 }}>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1.1fr 24px 1.6fr 26px",
              padding: "8px 12px",
              background: C.surfaceSoft,
              fontSize: 11.5,
              fontWeight: 700,
              color: C.inkSoft,
              fontFamily: FONT_UI,
            }}
          >
            <div>Raw SKU on document</div>
            <div></div>
            <div>Maps to</div>
            <div></div>
          </div>
          {entries.map(([key, v]) => (
            <div
              key={key}
              style={{
                display: "grid",
                gridTemplateColumns: "1.1fr 24px 1.6fr 26px",
                padding: "7px 12px",
                borderTop: `1px solid ${C.surfaceSoft}`,
                fontSize: 12.5,
                fontFamily: FONT_MONO,
                alignItems: "center",
              }}
            >
              <div>{v.raw}</div>
              <div style={{ color: C.inkSoft }}>→</div>
              <div>{aliasComponents(v).map((c) => `${c.sku} ×${c.qty}`).join(", ")}</div>
              <button onClick={() => remove(key)} style={{ border: "none", background: "none", cursor: "pointer", color: C.inkSoft }}>
                <Trash2 size={13} />
              </button>
            </div>
          ))}
        </div>
      )}

      {showAdd ? (
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 5, padding: 12, marginBottom: 14, background: C.surfaceSoft }}>
          <Field label="Raw SKU as it appears on the document">
            <input value={rawInput} onChange={(e) => setRawInput(e.target.value)} style={{ ...inputStyle, fontFamily: FONT_MONO }} placeholder="e.g. SH-MINI1G3" />
          </Field>
          <div style={{ fontSize: 12, color: C.inkSoft, marginBottom: 6, fontFamily: FONT_UI, fontWeight: 600 }}>Maps to</div>
          {components.map((c, ci) => (
            <div key={ci} style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 6 }}>
              <input
                list="wh-sku-datalist-alias"
                value={c.sku}
                onChange={(e) => updateComponent(ci, "sku", e.target.value)}
                placeholder="existing SKU"
                style={{ ...inputStyle, padding: "6px 8px", flex: 2, fontFamily: FONT_MONO }}
              />
              <span style={{ color: C.inkSoft }}>×</span>
              <input
                type="number"
                min="1"
                value={c.qty}
                onChange={(e) => updateComponent(ci, "qty", e.target.value)}
                style={{ ...inputStyle, padding: "6px 8px", width: 56 }}
              />
              {components.length > 1 && (
                <button onClick={() => removeComponent(ci)} style={{ border: "none", background: "none", cursor: "pointer", color: C.inkSoft }}>
                  <X size={14} />
                </button>
              )}
            </div>
          ))}
          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 8 }}>
            <Btn variant="outline" color={C.inkSoft} icon={Plus} onClick={addComponent} style={{ fontSize: 12.5, padding: "6px 10px" }}>
              Add component
            </Btn>
            <div style={{ display: "flex", gap: 8 }}>
              <Btn variant="outline" color={C.inkSoft} onClick={() => setShowAdd(false)}>
                Cancel
              </Btn>
              <Btn color={C.inventory} onClick={saveRule}>
                Save rule
              </Btn>
            </div>
          </div>
        </div>
      ) : (
        <Btn variant="outline" color={C.inventory} icon={Plus} onClick={() => setShowAdd(true)} style={{ marginBottom: 14 }}>
          Add rule manually
        </Btn>
      )}

      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        {isAdmin && (
          <Btn color={C.inventory} icon={Download} onClick={exportCSV} disabled={entries.length === 0}>
            Export CSV
          </Btn>
        )}
      </div>
    </Modal>
  );
}

function InsightsModal({ inventory, inboundRecords, outboundRecords, onClose }) {
  const [targetMonths, setTargetMonths] = useState(2);
  const [staleDays, setStaleDays] = useState(60);
  const { suggestions, staleItems } = computeInsights(
    inventory,
    inboundRecords,
    outboundRecords,
    3,
    Number(targetMonths) || 2,
    Number(staleDays) || 60
  );

  const exportSuggestions = () => {
    const rows = [
      ["SKU", "Product name", "Current new stock", "Avg monthly outbound (3mo)", "Months of coverage", "Suggested reorder qty"],
      ...suggestions.map((s) => [s.sku, s.name, s.currentQty, s.avgMonthly, s.monthsCoverage, s.suggestedQty]),
    ];
    downloadCSV(rows, `purchase-suggestions-${todayISO()}.csv`);
  };
  const exportStale = () => {
    const rows = [
      ["SKU", "Product name", "Current new stock", "Days since last movement", "Last movement date"],
      ...staleItems.map((s) => [s.sku, s.name, s.qty, s.daysSince, s.lastMovement]),
    ];
    downloadCSV(rows, `stale-stock-${todayISO()}.csv`);
  };

  return (
    <Modal title="Restock Insights" accent={C.inventory} onClose={onClose} wide>
      <div style={{ fontSize: 12.5, color: C.inkSoft, fontFamily: FONT_UI, marginBottom: 16 }}>
        Generated from your outbound and inbound history compared against current stock, as of right now — this is a
        suggestion to review, not an automatic order. Check back here whenever you want an updated read (e.g. once a
        month).
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <ShoppingCart size={16} color={C.inventory} />
        <div style={{ fontSize: 14, fontWeight: 700, fontFamily: FONT_UI, color: C.ink }}>Purchase suggestions</div>
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: C.inkSoft, fontFamily: FONT_UI }}>
          Target coverage
          <input
            type="number"
            min="1"
            value={targetMonths}
            onChange={(e) => setTargetMonths(e.target.value)}
            style={{ ...inputStyle, padding: "4px 6px", width: 48 }}
          />
          months
        </div>
      </div>
      <div style={{ fontSize: 11.5, color: C.inkSoft, fontFamily: FONT_UI, marginBottom: 8 }}>
        Based on average monthly outbound over the last 3 months. Only SKUs with recent outbound activity and below
        target coverage are listed.
      </div>
      {suggestions.length === 0 ? (
        <div style={{ fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI, padding: "6px 0 16px" }}>
          Nothing below target coverage right now.
        </div>
      ) : (
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 5, overflow: "hidden", marginBottom: 10, maxHeight: 220, overflowY: "auto" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1.4fr 0.7fr 0.9fr 0.8fr 0.9fr", padding: "7px 12px", background: C.surfaceSoft, fontSize: 11, fontWeight: 700, color: C.inkSoft, fontFamily: FONT_UI, position: "sticky", top: 0 }}>
            <div>SKU</div>
            <div>Product</div>
            <div>In stock</div>
            <div>Avg/mo</div>
            <div>Coverage</div>
            <div>Suggest order</div>
          </div>
          {suggestions.map((s) => (
            <div key={s.sku} style={{ display: "grid", gridTemplateColumns: "1fr 1.4fr 0.7fr 0.9fr 0.8fr 0.9fr", padding: "6px 12px", borderTop: `1px solid ${C.surfaceSoft}`, fontSize: 12, fontFamily: FONT_UI }}>
              <div style={{ fontFamily: FONT_MONO }}>{s.sku}</div>
              <div>{s.name}</div>
              <div style={{ fontFamily: FONT_MONO }}>{s.currentQty}</div>
              <div style={{ fontFamily: FONT_MONO }}>{s.avgMonthly}</div>
              <div style={{ fontFamily: FONT_MONO, color: s.monthsCoverage < 1 ? C.danger : C.ink }}>{s.monthsCoverage}mo</div>
              <div style={{ fontFamily: FONT_MONO, fontWeight: 700, color: C.inventory }}>+{s.suggestedQty}</div>
            </div>
          ))}
        </div>
      )}
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 20 }}>
        <Btn variant="outline" color={C.inventory} icon={Download} onClick={exportSuggestions} disabled={suggestions.length === 0} style={{ fontSize: 12.5, padding: "6px 10px" }}>
          Export CSV
        </Btn>
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <Clock size={16} color={C.amber} />
        <div style={{ fontSize: 14, fontWeight: 700, fontFamily: FONT_UI, color: C.ink }}>Stale stock</div>
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: C.inkSoft, fontFamily: FONT_UI }}>
          Flag after
          <input
            type="number"
            min="1"
            value={staleDays}
            onChange={(e) => setStaleDays(e.target.value)}
            style={{ ...inputStyle, padding: "4px 6px", width: 48 }}
          />
          days
        </div>
      </div>
      <div style={{ fontSize: 11.5, color: C.inkSoft, fontFamily: FONT_UI, marginBottom: 8 }}>
        New stock with no inbound or outbound movement for longer than the threshold above.
      </div>
      {staleItems.length === 0 ? (
        <div style={{ fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI, padding: "6px 0 8px" }}>Nothing flagged as stale right now.</div>
      ) : (
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 5, overflow: "hidden", marginBottom: 10, maxHeight: 220, overflowY: "auto" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1.6fr 0.7fr 1fr 1fr", padding: "7px 12px", background: C.surfaceSoft, fontSize: 11, fontWeight: 700, color: C.inkSoft, fontFamily: FONT_UI, position: "sticky", top: 0 }}>
            <div>SKU</div>
            <div>Product</div>
            <div>In stock</div>
            <div>Days idle</div>
            <div>Last movement</div>
          </div>
          {staleItems.map((s) => (
            <div key={s.sku} style={{ display: "grid", gridTemplateColumns: "1fr 1.6fr 0.7fr 1fr 1fr", padding: "6px 12px", borderTop: `1px solid ${C.surfaceSoft}`, fontSize: 12, fontFamily: FONT_UI }}>
              <div style={{ fontFamily: FONT_MONO }}>{s.sku}</div>
              <div>{s.name}</div>
              <div style={{ fontFamily: FONT_MONO }}>{s.qty}</div>
              <div style={{ fontFamily: FONT_MONO, color: C.danger, fontWeight: 700 }}>{s.daysSince}d</div>
              <div style={{ fontFamily: FONT_MONO }}>{s.lastMovement}</div>
            </div>
          ))}
        </div>
      )}
      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <Btn variant="outline" color={C.amber} icon={Download} onClick={exportStale} disabled={staleItems.length === 0} style={{ fontSize: 12.5, padding: "6px 10px" }}>
          Export CSV
        </Btn>
      </div>
    </Modal>
  );
}

function IgnoredSkuManager({ ignoredSkus, saveIgnoredSkus, onClose }) {
  const [input, setInput] = useState("");
  const add = () => {
    const val = input.trim();
    if (!val) return;
    if (ignoredSkus.some((s) => normalizeSku(s) === normalizeSku(val))) {
      setInput("");
      return;
    }
    saveIgnoredSkus([...ignoredSkus, val]);
    logAudit("ignored-add", `Added ignored code "${val}"`);
    setInput("");
  };
  const remove = (val) => {
    saveIgnoredSkus(ignoredSkus.filter((s) => s !== val));
    logAudit("ignored-delete", `Removed ignored code "${val}"`);
  };
  return (
    <Modal title="Ignored Codes" accent={C.inventory} onClose={onClose}>
      <div style={{ fontSize: 12.5, color: C.inkSoft, fontFamily: FONT_UI, marginBottom: 12 }}>
        Codes on this list (e.g. shipping fees, service charges) are dropped automatically from any inbound or outbound
        document — they're never treated as inventory SKUs.
      </div>
      <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
        <input
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && add()}
          placeholder="e.g. SPESE-TRASPORTO"
          style={{ ...inputStyle, fontFamily: FONT_MONO }}
        />
        <Btn color={C.inventory} icon={Plus} onClick={add} style={{ flexShrink: 0 }}>
          Add
        </Btn>
      </div>
      {ignoredSkus.length === 0 ? (
        <div style={{ fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI }}>No ignored codes yet.</div>
      ) : (
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 5, overflow: "hidden" }}>
          {ignoredSkus.map((s) => (
            <div
              key={s}
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                padding: "8px 12px",
                borderTop: `1px solid ${C.surfaceSoft}`,
                fontSize: 13,
                fontFamily: FONT_MONO,
              }}
            >
              {s}
              <button onClick={() => remove(s)} style={{ border: "none", background: "none", cursor: "pointer", color: C.inkSoft }}>
                <Trash2 size={13} />
              </button>
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

const inputStyle = {
  width: "100%",
  padding: "8px 10px",
  borderRadius: 4,
  border: `1px solid ${C.border}`,
  fontSize: 13.5,
  fontFamily: FONT_UI,
  outline: "none",
  boxSizing: "border-box",
};

function Field({ label, children, style }) {
  return (
    <div style={{ marginBottom: 12, ...style }}>
      <div style={{ fontSize: 12, color: C.inkSoft, marginBottom: 5, fontFamily: FONT_UI, fontWeight: 600 }}>{label}</div>
      {children}
    </div>
  );
}

function Modal({ title, accent, onClose, children, wide }) {
  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(20,24,32,0.45)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 500,
        padding: 16,
      }}
      onClick={onClose}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: C.surface,
          borderRadius: 6,
          width: wide ? 660 : 380,
          maxWidth: "100%",
          maxHeight: "88vh",
          overflowY: "auto",
          borderTop: `4px solid ${accent}`,
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 18px 0" }}>
          <div style={{ fontSize: 15.5, fontWeight: 700, fontFamily: FONT_UI, color: C.ink }}>{title}</div>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer", color: C.inkSoft }}>
            <X size={18} />
          </button>
        </div>
        <div style={{ padding: 18 }}>{children}</div>
      </div>
    </div>
  );
}

// SHA-256 content hash — used to spot a file that's been uploaded before,
// regardless of what it was renamed to.
async function hashFile(file) {
  const buf = await file.arrayBuffer();
  const digest = await crypto.subtle.digest("SHA-256", buf);
  return Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

// Looks for a previously-saved record (inbound or outbound, any type) whose
// source file had this exact content hash.
function findDuplicateByHash(hash, inboundRecords, outboundRecords) {
  const ib = inboundRecords.find((r) => r.fileHash === hash);
  if (ib) return { kind: "inbound", record: ib };
  const ob = outboundRecords.find((r) => r.fileHash === hash);
  if (ob) return { kind: "outbound", record: ob };
  return null;
}

function safeDate(v) {
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

// Compares recent outbound velocity per SKU (avg monthly qty shipped over the
// trailing `lookbackMonths`) against current New Stock to flag what's running
// low, and separately flags stock that's had no inbound/outbound movement in
// over `staleThresholdDays` — a signal of slow-moving or dead inventory.
// Reverses (or re-applies) the inventory effect of a record's item rows —
// sign = -1 to undo an inbound (subtract what it added), +1 to undo an
// outbound (add back what it removed). Used when an admin deletes a
// mistaken record so inventory stays consistent with what's actually on
// the books.
function applySignedQty(inventory, items, field, sign) {
  const map = new Map(inventory.map((x) => [x.sku, { ...x }]));
  items.forEach((it) => {
    const cur = map.get(it.sku) || { sku: it.sku, name: it.name, qtyNew: 0, qtyReturn: 0 };
    cur[field] = (Number(cur[field]) || 0) + sign * (Number(it.qty) || 0);
    map.set(it.sku, cur);
  });
  return Array.from(map.values());
}

function DeleteRecordConfirm({ record, onCancel, onConfirm }) {
  const count = record.items.length;
  return (
    <Modal title="Delete this record?" accent={C.danger} onClose={onCancel}>
      <div style={{ fontSize: 13.5, fontFamily: FONT_UI, color: C.ink, lineHeight: 1.6, marginBottom: 16 }}>
        This removes the record and automatically reverses the {count} item row{count === 1 ? "" : "s"} it applied to
        inventory, so stock stays accurate. This can't be undone from here — you'd need to re-upload the document to
        restore it.
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10 }}>
        <Btn variant="outline" color={C.inkSoft} onClick={onCancel}>
          Cancel
        </Btn>
        <Btn color={C.danger} icon={Trash2} onClick={onConfirm}>
          Delete &amp; reverse inventory
        </Btn>
      </div>
    </Modal>
  );
}

function computeInsights(inventory, inboundRecords, outboundRecords, lookbackMonths, targetMonths, staleThresholdDays) {
  const now = new Date();
  const cutoff = new Date(now);
  cutoff.setMonth(cutoff.getMonth() - lookbackMonths);

  const outboundInWindow = new Map();
  const lastOutbound = new Map();
  outboundRecords.forEach((r) => {
    const d = safeDate(r.shipDate || r.date);
    if (!d) return;
    r.items.forEach((it) => {
      const cur = lastOutbound.get(it.sku);
      if (!cur || d > cur) lastOutbound.set(it.sku, d);
      if (d >= cutoff) {
        outboundInWindow.set(it.sku, (outboundInWindow.get(it.sku) || 0) + (Number(it.qty) || 0));
      }
    });
  });

  const lastInbound = new Map();
  inboundRecords.forEach((r) => {
    const d = safeDate(r.invoiceDate || r.date);
    if (!d) return;
    r.items.forEach((it) => {
      const cur = lastInbound.get(it.sku);
      if (!cur || d > cur) lastInbound.set(it.sku, d);
    });
  });

  const suggestions = [];
  const staleItems = [];

  inventory.forEach((inv) => {
    const totalOut = outboundInWindow.get(inv.sku) || 0;
    const avgMonthly = totalOut / lookbackMonths;
    if (avgMonthly > 0) {
      const monthsCoverage = (Number(inv.qtyNew) || 0) / avgMonthly;
      if (monthsCoverage < targetMonths) {
        const suggestedQty = Math.ceil(targetMonths * avgMonthly - (Number(inv.qtyNew) || 0));
        if (suggestedQty > 0) {
          suggestions.push({
            sku: inv.sku,
            name: inv.name,
            currentQty: Number(inv.qtyNew) || 0,
            avgMonthly: Math.round(avgMonthly * 10) / 10,
            monthsCoverage: Math.round(monthsCoverage * 10) / 10,
            suggestedQty,
          });
        }
      }
    }

    if ((Number(inv.qtyNew) || 0) > 0) {
      const lo = lastOutbound.get(inv.sku);
      const li = lastInbound.get(inv.sku);
      const lastMovement = lo && li ? (lo > li ? lo : li) : lo || li;
      if (lastMovement) {
        const daysSince = Math.floor((now - lastMovement) / (1000 * 60 * 60 * 24));
        if (daysSince > staleThresholdDays) {
          staleItems.push({
            sku: inv.sku,
            name: inv.name,
            qty: Number(inv.qtyNew) || 0,
            daysSince,
            lastMovement: lastMovement.toISOString().slice(0, 10),
          });
        }
      }
    }
  });

  suggestions.sort((a, b) => a.monthsCoverage - b.monthsCoverage);
  staleItems.sort((a, b) => b.daysSince - a.daysSince);
  return { suggestions, staleItems };
}

function DuplicateFileWarning({ fileName, match, onCancel, onProceed }) {
  const label =
    match.kind === "inbound"
      ? `an inbound record (${match.record.type === "return" ? "Return Stock" : "New Stock Purchase"})`
      : `an outbound record (${match.record.type === "fba" ? "Amazon FBA Shipment" : "Order Fulfillment"})`;
  const dateStr = (match.record.date || "").slice(0, 10);
  return (
    <Modal title="This file looks familiar" accent={C.danger} onClose={onCancel}>
      <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
        <AlertTriangle size={20} color={C.danger} style={{ flexShrink: 0, marginTop: 2 }} />
        <div style={{ fontSize: 13.5, fontFamily: FONT_UI, color: C.ink, lineHeight: 1.6 }}>
          <strong>{fileName}</strong> has the exact same content as a file already processed on{" "}
          <strong>{dateStr || "an earlier date"}</strong>, saved as {label}. Uploading it again will add its items a
          second time.
        </div>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10 }}>
        <Btn variant="outline" color={C.inkSoft} onClick={onCancel}>
          Cancel
        </Btn>
        <Btn color={C.danger} onClick={onProceed}>
          Upload anyway
        </Btn>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------
   Editable extracted-items table with SKU-mapping controls
--------------------------------------------------------------- */
function ItemsEditor({ items, setItems, showPrice, inventory = [] }) {
  const update = (i, field, val) => {
    const next = [...items];
    const cur = { ...next[i], [field]: val };
    // Typing a SKU directly on a flagged row is a quick 1:1 correction —
    // still leaves the multiplier editable afterward (see the "mapped" box below).
    if (field === "sku" && cur.rawSku && cur.needsMapping) {
      const mult = Number(cur.components?.[0]?.qty) || 1;
      cur.mappedFrom = cur.rawSku;
      cur.needsMapping = false;
      cur.qtyMultiplier = mult;
      cur.qty = (cur.rawQty != null ? cur.rawQty : cur.qty) * mult;
    }
    next[i] = cur;
    setItems(next);
  };

  const remove = (i) => setItems(items.filter((_, idx) => idx !== i));
  const addRow = () => setItems([...items, { sku: "", name: "", qty: 1, unitPrice: 0 }]);

  // Re-derive qty from rawQty × the (possibly just-edited) multiplier — used
  // both by the initial mapping step and by the always-available correction
  // control on any already-mapped row, so the multiplier is never a dead end.
  const setRowMultiplier = (i, mult) => {
    const next = [...items];
    const cur = { ...next[i] };
    const m = Number(mult) || 1;
    cur.qtyMultiplier = m;
    cur.qty = (cur.rawQty != null ? cur.rawQty : cur.qty) * m;
    next[i] = cur;
    setItems(next);
  };

  // --- bundle-aware mapping controls for a still-unresolved row -----------
  const componentsFor = (it) => it.components || [{ sku: "", qty: it.qtyMultiplier || 1 }];

  const updateComponent = (i, ci, field, val) => {
    const next = [...items];
    const comps = componentsFor(next[i]).map((c, idx) => (idx === ci ? { ...c, [field]: val } : c));
    next[i] = { ...next[i], components: comps };
    setItems(next);
  };
  const addComponent = (i) => {
    const next = [...items];
    next[i] = { ...next[i], components: [...componentsFor(next[i]), { sku: "", qty: 1 }] };
    setItems(next);
  };
  const removeComponent = (i, ci) => {
    const next = [...items];
    next[i] = { ...next[i], components: componentsFor(next[i]).filter((_, idx) => idx !== ci) };
    setItems(next);
  };

  const applyMapping = (i) => {
    const it = items[i];
    const comps = componentsFor(it)
      .map((c) => ({ sku: String(c.sku || "").trim(), qty: Number(c.qty) || 1 }))
      .filter((c) => c.sku);
    const rawQty = it.rawQty != null ? it.rawQty : it.qty;
    if (comps.length === 0) {
      // nothing entered — treat the raw text itself as a new SKU, 1:1
      const next = [...items];
      next[i] = { ...it, sku: it.rawSku || it.sku, qty: rawQty, qtyMultiplier: 1, mappedFrom: undefined, needsMapping: false };
      setItems(next);
      return;
    }
    if (comps.length === 1) {
      const next = [...items];
      next[i] = {
        ...it,
        sku: comps[0].sku,
        qty: rawQty * comps[0].qty,
        qtyMultiplier: comps[0].qty,
        mappedFrom: it.rawSku,
        needsMapping: false,
      };
      setItems(next);
      return;
    }
    // Bundle: split this one row into one resolved row per component.
    const newRows = comps.map((c) => ({
      sku: c.sku,
      name: it.name,
      unitPrice: it.unitPrice,
      qty: rawQty * c.qty,
      qtyMultiplier: c.qty,
      rawSku: it.rawSku,
      rawQty,
      mappedFrom: it.rawSku,
      needsMapping: false,
      autoMapped: false,
    }));
    const next = [...items];
    next.splice(i, 1, ...newRows);
    setItems(next);
  };

  return (
    <div style={{ border: `1px solid ${C.border}`, borderRadius: 5, overflow: "hidden" }}>
      <datalist id="wh-sku-datalist">
        {inventory.map((inv) => (
          <option key={inv.sku} value={inv.sku}>
            {inv.name}
          </option>
        ))}
      </datalist>
      <div
        style={{
          display: "grid",
          gridTemplateColumns: showPrice ? "1fr 1.6fr 0.7fr 0.8fr 26px" : "1fr 1.8fr 0.7fr 26px",
          gap: 8,
          padding: "8px 12px",
          background: C.surfaceSoft,
          fontSize: 11.5,
          fontFamily: FONT_UI,
          fontWeight: 700,
          color: C.inkSoft,
        }}
      >
        <div>SKU</div>
        <div>Product name</div>
        <div>Qty</div>
        {showPrice && <div>Unit price</div>}
        <div></div>
      </div>
      {items.map((it, i) => (
        <div key={i} style={{ borderTop: `1px solid ${C.surfaceSoft}` }}>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: showPrice ? "1fr 1.6fr 0.7fr 0.8fr 26px" : "1fr 1.8fr 0.7fr 26px",
              gap: 8,
              padding: "6px 12px",
              alignItems: "center",
            }}
          >
            <input value={it.sku} onChange={(e) => update(i, "sku", e.target.value)} style={{ ...inputStyle, padding: "6px 8px", fontFamily: FONT_MONO }} />
            <input value={it.name} onChange={(e) => update(i, "name", e.target.value)} style={{ ...inputStyle, padding: "6px 8px" }} />
            <input type="number" value={it.qty} onChange={(e) => update(i, "qty", e.target.value)} style={{ ...inputStyle, padding: "6px 8px" }} />
            {showPrice && (
              <input type="number" value={it.unitPrice} onChange={(e) => update(i, "unitPrice", e.target.value)} style={{ ...inputStyle, padding: "6px 8px" }} />
            )}
            <button onClick={() => remove(i)} style={{ border: "none", background: "none", cursor: "pointer", color: C.inkSoft }}>
              <X size={14} />
            </button>
          </div>

          {it.needsMapping && (
            <div style={{ padding: "6px 12px 10px", background: C.amberSoft, fontSize: 12, fontFamily: FONT_UI }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, color: C.inkSoft, marginBottom: 6 }}>
                <AlertTriangle size={13} color={C.amber} style={{ flexShrink: 0 }} />
                <span>Unrecognized SKU "{it.rawSku}" — map it to one or more products (add more for a bundle):</span>
              </div>
              {componentsFor(it).map((c, ci) => (
                <div key={ci} style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 6 }}>
                  <input
                    list="wh-sku-datalist"
                    value={c.sku}
                    onChange={(e) => updateComponent(i, ci, "sku", e.target.value)}
                    placeholder="existing or new SKU"
                    style={{ ...inputStyle, padding: "4px 8px", flex: 2, minWidth: 120, fontFamily: FONT_MONO }}
                  />
                  <span style={{ color: C.inkSoft, whiteSpace: "nowrap" }}>×</span>
                  <input
                    type="number"
                    min="1"
                    value={c.qty}
                    onChange={(e) => updateComponent(i, ci, "qty", e.target.value)}
                    style={{ ...inputStyle, padding: "4px 8px", width: 56 }}
                  />
                  {componentsFor(it).length > 1 && (
                    <button onClick={() => removeComponent(i, ci)} style={{ border: "none", background: "none", cursor: "pointer", color: C.inkSoft }}>
                      <X size={13} />
                    </button>
                  )}
                </div>
              ))}
              <div style={{ display: "flex", gap: 8 }}>
                <Btn variant="outline" color={C.inkSoft} icon={Plus} onClick={() => addComponent(i)} style={{ fontSize: 12, padding: "5px 9px" }}>
                  Add component
                </Btn>
                <Btn color={C.amber} onClick={() => applyMapping(i)} style={{ fontSize: 12, padding: "5px 10px" }}>
                  Apply mapping
                </Btn>
              </div>
            </div>
          )}

          {!it.needsMapping && it.mappedFrom && (
            <div style={{ padding: "5px 12px 9px", background: C.surfaceSoft, fontSize: 11.5, fontFamily: FONT_UI }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, color: it.autoMapped ? C.inboundNew : C.inkSoft }}>
                <Link2 size={12} style={{ flexShrink: 0 }} />
                <span>
                  {it.autoMapped ? "Auto-matched via learned rule: " : "Mapped from: "}
                  {it.rawSku} → {it.sku}
                </span>
                <span style={{ marginLeft: "auto", whiteSpace: "nowrap" }}>units per row ×</span>
                <input
                  type="number"
                  min="1"
                  value={it.qtyMultiplier || 1}
                  onChange={(e) => setRowMultiplier(i, e.target.value)}
                  style={{ ...inputStyle, padding: "3px 7px", width: 50 }}
                />
              </div>
            </div>
          )}
        </div>
      ))}
      <div style={{ padding: 8 }}>
        <Btn variant="outline" color={C.inkSoft} icon={Plus} onClick={addRow} style={{ fontSize: 12.5, padding: "6px 10px" }}>
          Add row
        </Btn>
      </div>
    </div>
  );
}

function ColumnMapModal({ headers, accent, showPrice, onConfirm, onCancel }) {
  const [skuCol, setSkuCol] = useState(guessColumn(headers, ["sku", "codice", "model", "code"]));
  const [nameCol, setNameCol] = useState(guessColumn(headers, ["name", "product", "descriz", "articolo"]));
  const [qtyCol, setQtyCol] = useState(guessColumn(headers, ["qty", "quantit", "pcs", "quantity"]));
  const [priceCol, setPriceCol] = useState(guessColumn(headers, ["price", "prezzo", "unit"]));

  return (
    <Modal title="Match spreadsheet columns" accent={accent} onClose={onCancel}>
      <div style={{ fontSize: 12.5, color: C.inkSoft, fontFamily: FONT_UI, marginBottom: 14 }}>
        Columns are guessed automatically from the headers — please confirm or adjust.
      </div>
      <Field label="SKU column *">
        <select value={skuCol} onChange={(e) => setSkuCol(e.target.value)} style={inputStyle}>
          <option value="">Select…</option>
          {headers.map((h) => (
            <option key={h} value={h}>{h}</option>
          ))}
        </select>
      </Field>
      <Field label="Product name column">
        <select value={nameCol} onChange={(e) => setNameCol(e.target.value)} style={inputStyle}>
          <option value="">(none)</option>
          {headers.map((h) => (
            <option key={h} value={h}>{h}</option>
          ))}
        </select>
      </Field>
      <Field label="Quantity column *">
        <select value={qtyCol} onChange={(e) => setQtyCol(e.target.value)} style={inputStyle}>
          <option value="">Select…</option>
          {headers.map((h) => (
            <option key={h} value={h}>{h}</option>
          ))}
        </select>
      </Field>
      {showPrice && (
        <Field label="Unit price column">
          <select value={priceCol} onChange={(e) => setPriceCol(e.target.value)} style={inputStyle}>
            <option value="">(none)</option>
            {headers.map((h) => (
              <option key={h} value={h}>{h}</option>
            ))}
          </select>
        </Field>
      )}
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 14 }}>
        <Btn variant="outline" color={C.inkSoft} onClick={onCancel}>
          Cancel
        </Btn>
        <Btn color={accent} disabled={!skuCol || !qtyCol} onClick={() => onConfirm({ skuCol, nameCol, qtyCol, priceCol })}>
          Generate item list
        </Btn>
      </div>
    </Modal>
  );
}

function SkuSearchModal({ kind, records, onClose }) {
  const [sku, setSku] = useState("");
  const [months, setMonths] = useState(2);
  const isInbound = kind === "inbound";
  const accent = isInbound ? C.inboundNew : C.outbound;

  const norm = normalizeSku(sku);
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - (Number(months) || 2));

  const matches = [];
  if (norm) {
    records.forEach((r) => {
      const d = safeDate(isInbound ? r.invoiceDate || r.date : r.shipDate || r.date);
      if (!d || d < cutoff) return;
      (r.items || []).forEach((it) => {
        if (normalizeSku(it.sku) === norm) matches.push({ record: r, item: it, date: d });
      });
    });
  }
  matches.sort((a, b) => b.date - a.date);

  const gridCols = isInbound ? "1.1fr 0.9fr 0.6fr 0.7fr" : "0.6fr 1.3fr 0.9fr 0.6fr";

  return (
    <Modal title={isInbound ? "Search Purchase History" : "Search Shipment History"} accent={accent} onClose={onClose} wide>
      <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
        <Field label="SKU" style={{ flex: 1 }}>
          <input
            value={sku}
            onChange={(e) => setSku(e.target.value)}
            placeholder="Enter a SKU…"
            style={{ ...inputStyle, fontFamily: FONT_MONO }}
            autoFocus
          />
        </Field>
        <Field label="Look back" style={{ width: 130 }}>
          <select value={months} onChange={(e) => setMonths(e.target.value)} style={inputStyle}>
            <option value={1}>1 month</option>
            <option value={2}>2 months</option>
            <option value={3}>3 months</option>
            <option value={6}>6 months</option>
            <option value={12}>12 months</option>
          </select>
        </Field>
      </div>

      {!norm ? (
        <div style={{ fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI }}>Enter a SKU to search.</div>
      ) : matches.length === 0 ? (
        <div style={{ fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI }}>
          No {isInbound ? "purchase" : "shipment"} records found for this SKU in the selected window.
        </div>
      ) : (
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 5, maxHeight: 420, overflowY: "auto" }}>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: gridCols,
              padding: "8px 12px",
              background: C.surfaceSoft,
              fontSize: 11,
              fontWeight: 700,
              color: C.inkSoft,
              fontFamily: FONT_UI,
              position: "sticky",
              top: 0,
            }}
          >
            {isInbound ? (
              <>
                <div>Order</div>
                <div>Date</div>
                <div>Qty</div>
                <div>Unit price</div>
              </>
            ) : (
              <>
                <div>Type</div>
                <div>Reference</div>
                <div>Date</div>
                <div>Qty</div>
              </>
            )}
          </div>
          {matches.map((m, i) => (
            <div
              key={i}
              style={{
                display: "grid",
                gridTemplateColumns: gridCols,
                padding: "7px 12px",
                borderTop: `1px solid ${C.surfaceSoft}`,
                fontSize: 12.5,
                fontFamily: FONT_UI,
                alignItems: "center",
              }}
            >
              {isInbound ? (
                <>
                  <div style={{ fontFamily: FONT_MONO }}>{m.record.orderId || m.record.supplier || "—"}</div>
                  <div style={{ fontFamily: FONT_MONO }}>{m.record.invoiceDate || m.record.date?.slice(0, 10)}</div>
                  <div style={{ fontFamily: FONT_MONO }}>{m.item.qty}</div>
                  <div style={{ fontFamily: FONT_MONO }}>€{fmtMoney(m.item.unitPrice)}</div>
                </>
              ) : (
                <>
                  <div>
                    <Stamp
                      label={m.record.type === "fba" ? "FBA" : "Order"}
                      color={m.record.type === "fba" ? C.outbound : C.outboundOrder}
                      bg={C.surface}
                    />
                  </div>
                  <div style={{ fontFamily: FONT_MONO }}>
                    {m.record.type === "fba"
                      ? m.record.shipmentId || "—"
                      : [m.record.platform, m.item.orderId].filter(Boolean).join(" · ") || "—"}
                  </div>
                  <div style={{ fontFamily: FONT_MONO }}>{m.record.shipDate || m.record.date?.slice(0, 10)}</div>
                  <div style={{ fontFamily: FONT_MONO }}>{m.item.qty}</div>
                </>
              )}
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

function RecordsList({ records, dateField, showAmount, sourceLabel, isAdmin, onDeleteClick }) {
  if (records.length === 0) {
    return <div style={{ padding: 16, fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI }}>No records yet</div>;
  }
  return (
    <div style={{ marginTop: 10, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 5, maxHeight: 260, overflowY: "auto" }}>
      {records.slice(0, 30).map((r) => (
        <div key={r.id} style={{ padding: "10px 14px", borderBottom: `1px solid ${C.surfaceSoft}`, fontSize: 12.5, fontFamily: FONT_UI }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
            <span style={{ fontWeight: 700, color: C.ink }}>{r.supplier || r.platform || r.shipmentId || "—"}</span>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <span style={{ fontFamily: FONT_MONO, color: C.inkSoft }}>{r[dateField] || r.date?.slice(0, 10)}</span>
              {isAdmin && onDeleteClick && (
                <button
                  onClick={() => onDeleteClick(r)}
                  title="Delete this record"
                  style={{ border: "none", background: "none", cursor: "pointer", color: C.inkSoft, display: "flex" }}
                >
                  <Trash2 size={13} />
                </button>
              )}
            </div>
          </div>
          <div style={{ color: C.inkSoft, marginTop: 3 }}>
            {r.items.length} items{sourceLabel && sourceLabel(r) ? ` · ${sourceLabel(r)}` : ""}
            {showAmount && ` · Total €${fmtMoney(r.totalAmount)}`}
          </div>
        </div>
      ))}
    </div>
  );
}

function MonthlyReportModal({ records, onClose, title }) {
  const [month, setMonth] = useState(new Date().toISOString().slice(0, 7));
  const shift = (delta) => {
    const [y, m] = month.split("-").map(Number);
    const d = new Date(y, m - 1 + delta, 1);
    setMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
  };
  const monthRecords = records.filter((r) => monthKey(r.invoiceDate || r.date) === month);
  const totalAmount = monthRecords.reduce((s, r) => s + (Number(r.totalAmount) || 0), 0);
  const bySku = new Map();
  monthRecords.forEach((r) =>
    r.items.forEach((it) => {
      const cur = bySku.get(it.sku) || { sku: it.sku, name: it.name, qty: 0, amount: 0 };
      cur.qty += Number(it.qty) || 0;
      cur.amount += (Number(it.qty) || 0) * (Number(it.unitPrice) || 0);
      bySku.set(it.sku, cur);
    })
  );
  const skuRows = Array.from(bySku.values()).sort((a, b) => b.amount - a.amount);

  const exportCSV = () => {
    const rows = [["SKU", "Product name", "Qty", "Amount"], ...skuRows.map((r) => [r.sku, r.name, r.qty, r.amount.toFixed(2)])];
    rows.push([]);
    rows.push(["Records", monthRecords.length]);
    rows.push(["Total amount", totalAmount.toFixed(2)]);
    downloadCSV(rows, `${title.toLowerCase().replace(/\s+/g, "-")}-${month}.csv`);
  };

  return (
    <Modal title={title} accent={C.amber} onClose={onClose} wide>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 14, marginBottom: 16 }}>
        <button onClick={() => shift(-1)} style={{ border: "none", background: "none", cursor: "pointer", color: C.ink }}>
          <ChevronLeft size={18} />
        </button>
        <span style={{ fontFamily: FONT_MONO, fontWeight: 700, fontSize: 16 }}>{month}</span>
        <button onClick={() => shift(1)} style={{ border: "none", background: "none", cursor: "pointer", color: C.ink }}>
          <ChevronRight size={18} />
        </button>
      </div>

      <div style={{ display: "flex", gap: 12, marginBottom: 16 }}>
        <StatBox label="Records" value={monthRecords.length} />
        <StatBox label="Total value" value={`€${fmtMoney(totalAmount)}`} />
        <StatBox label="SKUs" value={skuRows.length} />
      </div>

      {skuRows.length === 0 ? (
        <div style={{ fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI, padding: 12 }}>No records this month</div>
      ) : (
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 5, overflow: "hidden", marginBottom: 14 }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1.6fr 0.6fr 0.8fr", padding: "8px 12px", background: C.surfaceSoft, fontSize: 11.5, fontWeight: 700, color: C.inkSoft, fontFamily: FONT_UI }}>
            <div>SKU</div>
            <div>Product name</div>
            <div>Qty</div>
            <div>Amount</div>
          </div>
          <div style={{ maxHeight: 260, overflowY: "auto" }}>
            {skuRows.map((r) => (
              <div key={r.sku} style={{ display: "grid", gridTemplateColumns: "1fr 1.6fr 0.6fr 0.8fr", padding: "7px 12px", borderTop: `1px solid ${C.surfaceSoft}`, fontSize: 12.5, fontFamily: FONT_UI }}>
                <div style={{ fontFamily: FONT_MONO }}>{r.sku}</div>
                <div>{r.name}</div>
                <div style={{ fontFamily: FONT_MONO }}>{r.qty}</div>
                <div style={{ fontFamily: FONT_MONO }}>€{fmtMoney(r.amount)}</div>
              </div>
            ))}
          </div>
        </div>
      )}
      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <Btn color={C.amber} icon={Download} onClick={exportCSV} disabled={skuRows.length === 0}>
          Export CSV
        </Btn>
      </div>
    </Modal>
  );
}

function StatBox({ label, value }) {
  return (
    <div style={{ flex: 1, background: C.surfaceSoft, borderRadius: 5, padding: "12px 14px", borderLeft: `3px solid ${C.amber}` }}>
      <div style={{ fontSize: 11.5, color: C.inkSoft, fontFamily: FONT_UI, fontWeight: 600 }}>{label}</div>
      <div style={{ fontSize: 19, fontWeight: 700, fontFamily: FONT_MONO, color: C.ink, marginTop: 3 }}>{value}</div>
    </div>
  );
}

/* ---------------------------------------------------------------
   INBOUND HUB — choose new-stock purchase vs. return-stock inbound
--------------------------------------------------------------- */
function InboundHub({ setView, inboundRecords }) {
  const [showSearch, setShowSearch] = useState(false);
  const thisMonth = new Date().toISOString().slice(0, 7);
  const countFor = (type) =>
    inboundRecords.filter((r) => (r.type || "new") === type && monthKey(r.invoiceDate || r.date) === thisMonth).length;

  const cards = [
    {
      key: "inbound-new",
      title: "New Stock Purchase",
      desc: "Upload purchase invoices for brand-new inventory",
      accent: C.inboundNew,
      icon: PackagePlus,
      stat: `${countFor("new")} records this month`,
    },
    {
      key: "inbound-return",
      title: "Return Stock Inbound",
      desc: "Upload documents for returned inventory coming back into stock",
      accent: C.inboundReturn,
      icon: RotateCcw,
      stat: `${countFor("return")} records this month`,
    },
  ];

  return (
    <div>
      <TopBar title="Inbound" subtitle="Choose the type of inbound stock" onBack={() => setView("home")} accent={C.inboundNew} />
      <div style={{ padding: "20px 20px 0" }}>
        <Btn color={C.inboundNew} variant="outline" icon={Search} onClick={() => setShowSearch(true)}>
          Search purchase history by SKU
        </Btn>
      </div>
      <div style={{ padding: 20, display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 14 }}>
        {cards.map((c) => (
          <div
            key={c.key}
            onClick={() => setView(c.key)}
            style={{
              background: C.surface,
              borderRadius: 5,
              border: `1px solid ${C.border}`,
              borderLeft: `5px solid ${c.accent}`,
              padding: "20px 18px",
              cursor: "pointer",
              transition: "transform 0.12s ease",
            }}
            onMouseEnter={(e) => (e.currentTarget.style.transform = "translateY(-2px)")}
            onMouseLeave={(e) => (e.currentTarget.style.transform = "none")}
          >
            <c.icon size={22} color={c.accent} strokeWidth={2} />
            <div style={{ fontSize: 17, fontWeight: 700, marginTop: 12, color: C.ink, fontFamily: FONT_UI }}>{c.title}</div>
            <div style={{ fontSize: 12.5, color: C.inkSoft, marginTop: 4, fontFamily: FONT_UI, lineHeight: 1.5 }}>{c.desc}</div>
            <div style={{ marginTop: 14, fontSize: 12.5, fontFamily: FONT_MONO, fontWeight: 700, color: c.accent }}>{c.stat}</div>
          </div>
        ))}
      </div>
      {showSearch && <SkuSearchModal kind="inbound" records={inboundRecords} onClose={() => setShowSearch(false)} />}
    </div>
  );
}

/* ---------------------------------------------------------------
   INBOUND FLOW — shared by "new" and "return" types
--------------------------------------------------------------- */
function InboundFlow({ type, inventory, saveInventory, inboundRecords, saveInboundRecords, outboundRecords, aliasMap, saveAliasMap, ignoredSkus, showToast, setView, isAdmin }) {
  const isNew = type === "new";
  const accent = isNew ? C.inboundNew : C.inboundReturn;
  const accentSoft = isNew ? C.inboundNewSoft : C.inboundReturnSoft;
  const targetField = isNew ? "qtyNew" : "qtyReturn";
  const partyLabel = isNew ? "Supplier" : "Source / returned by";
  const dateLabel = isNew ? "Invoice date" : "Return date";
  const title = isNew ? "New Stock Purchase" : "Return Stock Inbound";
  const docLabel = isNew ? "purchase invoice" : "return document";

  const [busy, setBusy] = useState(false);
  const [draft, setDraft] = useState(null);
  const [excelPending, setExcelPending] = useState(null);
  const [showReport, setShowReport] = useState(false);
  const [pendingDuplicate, setPendingDuplicate] = useState(null); // { file, hash, match }
  const [deleteTarget, setDeleteTarget] = useState(null);

  const typeRecords = inboundRecords.filter((r) => (r.type || "new") === type);

  const performDelete = (record) => {
    saveInventory(applySignedQty(inventory, record.items, targetField, -1));
    saveInboundRecords(inboundRecords.filter((r) => r.id !== record.id));
    logAudit(
      isNew ? "inbound-new-delete" : "inbound-return-delete",
      `Deleted ${isNew ? "new stock purchase" : "return stock inbound"} record from "${record.fileName}" (${record.supplier}) — reversed ${record.items.length} item rows`
    );
    setDeleteTarget(null);
    showToast("Record deleted, inventory reversed");
  };

  const handleFile = async (file) => {
    let hash = null;
    try {
      hash = await hashFile(file);
    } catch (e) {}
    if (hash) {
      const match = findDuplicateByHash(hash, inboundRecords, outboundRecords);
      if (match) {
        setPendingDuplicate({ file, hash, match });
        return;
      }
    }
    await processFile(file, hash);
  };

  const processFile = async (file, hash) => {
    if (isSpreadsheetFile(file)) {
      try {
        const { grid, rows, headers } = await parseSpreadsheetFile(file);
        const sections = findSkuSections(grid);
        const sectionItems = itemsFromSkuSections(sections);
        if (sectionItems.length > 0) {
          const resolved = resolveSkus(sectionItems, inventory, aliasMap, ignoredSkus);
          setDraft({
            supplier: "",
            invoiceDate: todayISO(),
            items: resolved,
            totalAmount: resolved.reduce((s, i) => s + i.qty * i.unitPrice, 0),
            fileName: file.name,
            fileHash: hash,
            orderId: orderIdFromFilename(file.name),
          });
          showToast(`Detected ${sectionItems.length} SKUs automatically from ${sections.length > 1 ? sections.length + " sections" : "the sheet"}`);
          return;
        }
        if (rows.length === 0) {
          showToast("The spreadsheet has no data rows", "error");
          return;
        }
        setExcelPending({ rows, headers, fileName: file.name, fileHash: hash });
      } catch (e) {
        showToast("Couldn't parse the spreadsheet — check the file format", "error");
      }
      return;
    }
    setBusy(true);
    try {
      const extracted = await extractFromFile(file, "inbound");
      const rawItems = (extracted.items || []).map((it) => ({
        sku: it.sku_guess || "",
        name: it.name || "",
        qty: Number(it.qty) || 0,
        unitPrice: Number(it.unit_price) || 0,
      }));
      setDraft({
        supplier: extracted.supplier || "",
        invoiceDate: extracted.invoice_date || todayISO(),
        items: resolveSkus(rawItems, inventory, aliasMap, ignoredSkus),
        totalAmount: Number(extracted.total_amount) || 0,
        fileName: file.name,
        fileHash: hash,
        orderId: orderIdFromFilename(file.name),
      });
      showToast("Extraction complete — please review and confirm");
    } catch (e) {
      showToast(e.message || "Recognition failed — please retry or enter manually", "error");
      setDraft({
        supplier: "",
        invoiceDate: todayISO(),
        items: [{ sku: "", name: "", qty: 1, unitPrice: 0 }],
        totalAmount: 0,
        fileName: file.name,
        fileHash: hash,
        orderId: orderIdFromFilename(file.name),
      });
    } finally {
      setBusy(false);
    }
  };

  const handleColumnsConfirmed = ({ skuCol, nameCol, qtyCol, priceCol }) => {
    const rawItems = excelPending.rows
      .map((r) => ({
        sku: String(r[skuCol] ?? "").trim(),
        name: nameCol ? String(r[nameCol] ?? "").trim() : "",
        qty: Number(r[qtyCol]) || 0,
        unitPrice: priceCol ? Number(r[priceCol]) || 0 : 0,
      }))
      .filter((it) => it.sku && it.qty > 0);
    const resolved = resolveSkus(rawItems, inventory, aliasMap, ignoredSkus);
    setDraft({
      supplier: "",
      invoiceDate: todayISO(),
      items: resolved,
      totalAmount: resolved.reduce((s, i) => s + i.qty * i.unitPrice, 0),
      fileName: excelPending.fileName,
      fileHash: excelPending.fileHash,
      orderId: orderIdFromFilename(excelPending.fileName),
    });
    setExcelPending(null);
    showToast(`Generated ${resolved.length} item rows from the spreadsheet — please review`);
  };

  const confirmInbound = () => {
    const cleanItems = draft.items.filter((it) => it.sku.trim() && Number(it.qty) > 0);
    if (cleanItems.length === 0) {
      showToast("No valid item rows", "error");
      return;
    }
    const map = new Map(inventory.map((x) => [x.sku, { ...x }]));
    cleanItems.forEach((it) => {
      const sku = it.sku.trim();
      if (map.has(sku)) {
        const cur = map.get(sku);
        cur[targetField] = (Number(cur[targetField]) || 0) + Number(it.qty);
        if (it.name) cur.name = it.name;
        map.set(sku, cur);
      } else {
        map.set(sku, { sku, name: it.name || sku, qtyNew: 0, qtyReturn: 0, [targetField]: Number(it.qty) });
      }
    });
    saveInventory(Array.from(map.values()));

    // Teach the alias table from any rows the person mapped by hand — grouped
    // by mappedFrom so a bundle (one raw SKU → several different components)
    // saves as a single multi-component rule instead of overwriting itself.
    const nextAliases = { ...aliasMap };
    let aliasChanged = false;
    const mapGroups = new Map();
    cleanItems.forEach((it) => {
      if (!it.mappedFrom) return;
      const norm = normalizeSku(it.mappedFrom);
      if (!mapGroups.has(norm)) mapGroups.set(norm, { raw: it.mappedFrom, components: [] });
      mapGroups.get(norm).components.push({ sku: it.sku.trim(), qty: Number(it.qtyMultiplier) || 1 });
    });
    mapGroups.forEach((val, norm) => {
      const existing = nextAliases[norm] ? aliasComponents(nextAliases[norm]) : null;
      const changed =
        !existing ||
        existing.length !== val.components.length ||
        existing.some((c, idx) => c.sku !== val.components[idx].sku || Number(c.qty) !== Number(val.components[idx].qty));
      if (changed) {
        nextAliases[norm] = { raw: val.raw, components: val.components };
        aliasChanged = true;
      }
    });
    if (aliasChanged) saveAliasMap(nextAliases);

    const record = {
      id: Date.now().toString(),
      date: new Date().toISOString(),
      invoiceDate: draft.invoiceDate,
      supplier: draft.supplier || "Unspecified",
      orderId: draft.orderId || "",
      fileName: draft.fileName,
      fileHash: draft.fileHash || null,
      items: cleanItems,
      totalAmount: Number(draft.totalAmount) || cleanItems.reduce((s, i) => s + i.qty * i.unitPrice, 0),
      type,
    };
    saveInboundRecords([record, ...inboundRecords]);
    logAudit(
      isNew ? "inbound-new-confirm" : "inbound-return-confirm",
      `${isNew ? "New stock purchase" : "Return stock inbound"} confirmed: ${cleanItems.length} item rows from "${draft.fileName}" (${record.supplier})`
    );
    setDraft(null);
    showToast(isNew ? "Purchase confirmed, inventory updated" : "Return inbound confirmed, inventory updated");
  };

  return (
    <div>
      <TopBar
        title={title}
        subtitle={isNew ? "Upload purchase invoices to log new inventory" : "Upload return documents to log inventory coming back in"}
        onBack={() => setView("inbound")}
        accent={accent}
      />
      <div style={{ padding: 18 }}>
        {!draft && (
          <UploadBox
            accent={accent}
            accentSoft={accentSoft}
            label={`Click or drag to upload a ${docLabel}`}
            hint="Supports image, PDF, or Excel/CSV · auto-extracts items, quantities and amounts"
            onFile={handleFile}
            busy={busy}
          />
        )}

        {draft && (
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 16 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <div style={{ fontSize: 13, fontFamily: FONT_UI, color: C.inkSoft, display: "flex", alignItems: "center", gap: 6 }}>
                <FileText size={14} /> {draft.fileName}
              </div>
              <Stamp label="Pending review" color={C.amber} bg={C.amberSoft} />
            </div>
            <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
              <Field label={partyLabel} style={{ flex: 1, minWidth: 160 }}>
                <input value={draft.supplier} onChange={(e) => setDraft({ ...draft, supplier: e.target.value })} style={inputStyle} />
              </Field>
              <Field label="Order ID" style={{ width: 140 }}>
                <input value={draft.orderId || ""} onChange={(e) => setDraft({ ...draft, orderId: e.target.value })} style={{ ...inputStyle, fontFamily: FONT_MONO }} />
              </Field>
              <Field label={dateLabel} style={{ width: 160 }}>
                <input type="date" value={draft.invoiceDate} onChange={(e) => setDraft({ ...draft, invoiceDate: e.target.value })} style={inputStyle} />
              </Field>
              <Field label="Total value" style={{ width: 130 }}>
                <input type="number" value={draft.totalAmount} onChange={(e) => setDraft({ ...draft, totalAmount: e.target.value })} style={inputStyle} />
              </Field>
            </div>
            <ItemsEditor items={draft.items} setItems={(items) => setDraft({ ...draft, items })} showPrice inventory={inventory} />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 14 }}>
              <Btn variant="outline" color={C.inkSoft} onClick={() => setDraft(null)}>
                Cancel
              </Btn>
              <Btn color={accent} icon={CheckCircle2} onClick={confirmInbound}>
                {isNew ? "Confirm Purchase" : "Confirm Return"}
              </Btn>
            </div>
          </div>
        )}

        <div style={{ marginTop: 20, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ fontSize: 13.5, fontWeight: 700, fontFamily: FONT_UI, color: C.ink }}>Recent records</div>
          {isAdmin && (
            <Btn color={C.amber} icon={Download} onClick={() => setShowReport(true)}>
              Generate Monthly Report
            </Btn>
          )}
        </div>
        <RecordsList
          records={typeRecords}
          dateField="invoiceDate"
          showAmount
          sourceLabel={(r) => (r.orderId ? `Order ${r.orderId}` : "")}
          isAdmin={isAdmin}
          onDeleteClick={setDeleteTarget}
        />
      </div>

      {excelPending && (
        <ColumnMapModal
          headers={excelPending.headers}
          accent={accent}
          showPrice
          onConfirm={handleColumnsConfirmed}
          onCancel={() => setExcelPending(null)}
        />
      )}

      {pendingDuplicate && (
        <DuplicateFileWarning
          fileName={pendingDuplicate.file.name}
          match={pendingDuplicate.match}
          onCancel={() => setPendingDuplicate(null)}
          onProceed={() => {
            const { file, hash } = pendingDuplicate;
            setPendingDuplicate(null);
            processFile(file, hash);
          }}
        />
      )}

      {showReport && (
        <MonthlyReportModal
          records={typeRecords}
          title={isNew ? "Monthly Report — New Stock Purchases" : "Monthly Report — Return Stock Inbound"}
          onClose={() => setShowReport(false)}
        />
      )}

      {deleteTarget && (
        <DeleteRecordConfirm record={deleteTarget} onCancel={() => setDeleteTarget(null)} onConfirm={() => performDelete(deleteTarget)} />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   OUTBOUND HUB — choose Amazon FBA shipment vs. order fulfillment
--------------------------------------------------------------- */
function OutboundHub({ setView, outboundRecords }) {
  const [showSearch, setShowSearch] = useState(false);
  const thisMonth = new Date().toISOString().slice(0, 7);
  const countFor = (type) =>
    outboundRecords.filter((r) => (r.type || "order") === type && monthKey(r.shipDate || r.date) === thisMonth).length;

  const cards = [
    {
      key: "outbound-fba",
      title: "Amazon FBA Shipment",
      desc: "Upload FBA shipment plans — stock going to Amazon's warehouse",
      accent: C.outbound,
      icon: PackageMinus,
      stat: `${countFor("fba")} shipments this month`,
    },
    {
      key: "outbound-order",
      title: "Order Fulfillment",
      desc: "Upload packing slips or order exports — stock going to customers",
      accent: C.outboundOrder,
      icon: FileText,
      stat: `${countFor("order")} shipments this month`,
    },
  ];

  return (
    <div>
      <TopBar title="Outbound" subtitle="Choose the type of outbound shipment" onBack={() => setView("home")} accent={C.outbound} />
      <div style={{ padding: "20px 20px 0" }}>
        <Btn color={C.outbound} variant="outline" icon={Search} onClick={() => setShowSearch(true)}>
          Search shipment history by SKU
        </Btn>
      </div>
      <div style={{ padding: 20, display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 14 }}>
        {cards.map((c) => (
          <div
            key={c.key}
            onClick={() => setView(c.key)}
            style={{
              background: C.surface,
              borderRadius: 5,
              border: `1px solid ${C.border}`,
              borderLeft: `5px solid ${c.accent}`,
              padding: "20px 18px",
              cursor: "pointer",
              transition: "transform 0.12s ease",
            }}
            onMouseEnter={(e) => (e.currentTarget.style.transform = "translateY(-2px)")}
            onMouseLeave={(e) => (e.currentTarget.style.transform = "none")}
          >
            <c.icon size={22} color={c.accent} strokeWidth={2} />
            <div style={{ fontSize: 17, fontWeight: 700, marginTop: 12, color: C.ink, fontFamily: FONT_UI }}>{c.title}</div>
            <div style={{ fontSize: 12.5, color: C.inkSoft, marginTop: 4, fontFamily: FONT_UI, lineHeight: 1.5 }}>{c.desc}</div>
            <div style={{ marginTop: 14, fontSize: 12.5, fontFamily: FONT_MONO, fontWeight: 700, color: c.accent }}>{c.stat}</div>
          </div>
        ))}
      </div>
      {showSearch && <SkuSearchModal kind="outbound" records={outboundRecords} onClose={() => setShowSearch(false)} />}
    </div>
  );
}

/* ---------------------------------------------------------------
   OUTBOUND — Amazon FBA Shipment
--------------------------------------------------------------- */
function OutboundFbaFlow({ setView, inventory, saveInventory, outboundRecords, saveOutboundRecords, inboundRecords, aliasMap, saveAliasMap, ignoredSkus, showToast, isAdmin }) {
  const [busy, setBusy] = useState(false);
  const [draft, setDraft] = useState(null); // + shipmentId, shipmentName, boxes
  const [excelPending, setExcelPending] = useState(null);
  const [showReport, setShowReport] = useState(false);
  const [pendingDuplicate, setPendingDuplicate] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);

  const typeRecords = outboundRecords.filter((r) => (r.type || "order") === "fba");

  const performDelete = (record) => {
    const field = record.source === "New Stock" ? "qtyNew" : "qtyReturn";
    saveInventory(applySignedQty(inventory, record.items, field, 1));
    saveOutboundRecords(outboundRecords.filter((r) => r.id !== record.id));
    logAudit("outbound-fba-delete", `Deleted FBA shipment ${record.shipmentId} — reversed ${record.items.length} item rows`);
    setDeleteTarget(null);
    showToast("Record deleted, inventory reversed");
  };

  const handleFile = async (file) => {
    let hash = null;
    try {
      hash = await hashFile(file);
    } catch (e) {}
    if (hash) {
      const match = findDuplicateByHash(hash, inboundRecords, outboundRecords);
      if (match) {
        setPendingDuplicate({ file, hash, match });
        return;
      }
    }
    await processFile(file, hash);
  };

  const processFile = async (file, hash) => {
    if (isSpreadsheetFile(file)) {
      try {
        const { grid, rows, headers } = await parseSpreadsheetFile(file);
        const meta = parseFbaMeta(grid);
        const sections = findSkuSections(grid);
        const sectionItems = itemsFromSkuSections(sections);
        if (sectionItems.length > 0) {
          const resolved = resolveSkus(sectionItems, inventory, aliasMap, ignoredSkus);
          setDraft({
            shipmentId: meta.shipmentId || "",
            shipmentName: meta.shipmentName || "",
            shipTo: meta.shipTo || "",
            boxes: meta.boxes || 0,
            shipDate: todayISO(),
            source: "qtyNew",
            items: resolved,
            fileName: file.name,
            fileHash: hash,
          });
          showToast(
            meta.shipmentId
              ? `Detected shipment ${meta.shipmentId} — ${resolved.length} SKUs, ${meta.boxes || 0} boxes`
              : `Detected ${resolved.length} SKUs automatically`
          );
          return;
        }
        if (rows.length === 0) {
          showToast("The spreadsheet has no data rows", "error");
          return;
        }
        setExcelPending({ rows, headers, fileName: file.name, meta, fileHash: hash });
      } catch (e) {
        showToast("Couldn't parse the spreadsheet — check the file format", "error");
      }
      return;
    }
    setBusy(true);
    try {
      const extracted = await extractFromFile(file, "outbound");
      const rawItems = (extracted.items || []).map((it) => ({
        sku: it.sku_guess || "",
        name: it.name || "",
        qty: Number(it.qty) || 0,
      }));
      setDraft({
        shipmentId: "",
        shipmentName: "",
        shipTo: "",
        boxes: 0,
        shipDate: extracted.ship_date || todayISO(),
        source: "qtyNew",
        items: resolveSkus(rawItems, inventory, aliasMap, ignoredSkus),
        fileName: file.name,
        fileHash: hash,
      });
      showToast("Extraction complete — please review and confirm");
    } catch (e) {
      showToast(e.message || "Recognition failed — please retry or enter manually", "error");
      setDraft({ shipmentId: "", shipmentName: "", shipTo: "", boxes: 0, shipDate: todayISO(), source: "qtyNew", items: [{ sku: "", name: "", qty: 1 }], fileName: file.name, fileHash: hash });
    } finally {
      setBusy(false);
    }
  };

  const handleColumnsConfirmed = ({ skuCol, nameCol, qtyCol }) => {
    const rawItems = excelPending.rows
      .map((r) => ({
        sku: String(r[skuCol] ?? "").trim(),
        name: nameCol ? String(r[nameCol] ?? "").trim() : "",
        qty: Number(r[qtyCol]) || 0,
      }))
      .filter((it) => it.sku && it.qty > 0);
    const resolved = resolveSkus(rawItems, inventory, aliasMap, ignoredSkus);
    const meta = excelPending.meta || {};
    setDraft({
      shipmentId: meta.shipmentId || "",
      shipmentName: meta.shipmentName || "",
      shipTo: meta.shipTo || "",
      boxes: meta.boxes || 0,
      shipDate: todayISO(),
      source: "qtyNew",
      items: resolved,
      fileName: excelPending.fileName,
      fileHash: excelPending.fileHash,
    });
    setExcelPending(null);
    showToast(`Generated ${resolved.length} item rows from the spreadsheet — please review`);
  };

  const confirmOutbound = () => {
    const cleanItems = draft.items.filter((it) => it.sku.trim() && Number(it.qty) > 0);
    if (cleanItems.length === 0) {
      showToast("No valid item rows", "error");
      return;
    }
    const field = draft.source;
    const map = new Map(inventory.map((x) => [x.sku, { ...x }]));
    let shortage = false;
    cleanItems.forEach((it) => {
      const sku = it.sku.trim();
      const cur = map.get(sku) || { sku, name: it.name || sku, qtyNew: 0, qtyReturn: 0 };
      const remaining = (Number(cur[field]) || 0) - Number(it.qty);
      if (remaining < 0) shortage = true;
      cur[field] = remaining;
      map.set(sku, cur);
    });
    saveInventory(Array.from(map.values()));

    const nextAliases = { ...aliasMap };
    let aliasChanged = false;
    const mapGroups = new Map();
    cleanItems.forEach((it) => {
      if (!it.mappedFrom) return;
      const norm = normalizeSku(it.mappedFrom);
      if (!mapGroups.has(norm)) mapGroups.set(norm, { raw: it.mappedFrom, components: [] });
      mapGroups.get(norm).components.push({ sku: it.sku.trim(), qty: Number(it.qtyMultiplier) || 1 });
    });
    mapGroups.forEach((val, norm) => {
      const existing = nextAliases[norm] ? aliasComponents(nextAliases[norm]) : null;
      const changed =
        !existing ||
        existing.length !== val.components.length ||
        existing.some((c, idx) => c.sku !== val.components[idx].sku || Number(c.qty) !== Number(val.components[idx].qty));
      if (changed) {
        nextAliases[norm] = { raw: val.raw, components: val.components };
        aliasChanged = true;
      }
    });
    if (aliasChanged) saveAliasMap(nextAliases);

    const record = {
      id: Date.now().toString(),
      date: new Date().toISOString(),
      shipDate: draft.shipDate,
      shipmentId: draft.shipmentId || "Unspecified",
      shipmentName: draft.shipmentName || "",
      shipTo: draft.shipTo || "",
      boxes: Number(draft.boxes) || 0,
      fileName: draft.fileName,
      fileHash: draft.fileHash || null,
      source: field === "qtyNew" ? "New Stock" : "Return Stock",
      items: cleanItems,
      type: "fba",
    };
    saveOutboundRecords([record, ...outboundRecords]);
    logAudit("outbound-fba-confirm", `FBA shipment ${record.shipmentId} confirmed: ${cleanItems.length} item rows, ${record.boxes} boxes`);
    setDraft(null);
    showToast(shortage ? "Shipment confirmed, but some items are now negative — please check" : "Shipment confirmed, inventory updated", shortage ? "error" : "success");
  };

  return (
    <div>
      <TopBar title="Amazon FBA Shipment" subtitle="Upload shipment plans for stock going to Amazon's warehouse" onBack={() => setView("outbound")} accent={C.outbound} />
      <div style={{ padding: 18 }}>
        {!draft && (
          <UploadBox
            accent={C.outbound}
            accentSoft={C.outboundSoft}
            label="Click or drag to upload an FBA shipment plan"
            hint="Supports Excel/CSV (auto-detects Shipment ID, boxes, SKUs) or image/PDF"
            onFile={handleFile}
            busy={busy}
          />
        )}

        {draft && (
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 16 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <div style={{ fontSize: 13, fontFamily: FONT_UI, color: C.inkSoft, display: "flex", alignItems: "center", gap: 6 }}>
                <FileText size={14} /> {draft.fileName}
              </div>
              <Stamp label="Pending review" color={C.amber} bg={C.amberSoft} />
            </div>
            <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
              <Field label="Shipment ID" style={{ flex: 1, minWidth: 160 }}>
                <input value={draft.shipmentId} onChange={(e) => setDraft({ ...draft, shipmentId: e.target.value })} style={{ ...inputStyle, fontFamily: FONT_MONO }} />
              </Field>
              <Field label="Ship date" style={{ width: 150 }}>
                <input type="date" value={draft.shipDate} onChange={(e) => setDraft({ ...draft, shipDate: e.target.value })} style={inputStyle} />
              </Field>
              <Field label="Boxes" style={{ width: 90 }}>
                <input type="number" value={draft.boxes} onChange={(e) => setDraft({ ...draft, boxes: e.target.value })} style={inputStyle} />
              </Field>
              <Field label="Outbound source" style={{ width: 150 }}>
                <select value={draft.source} onChange={(e) => setDraft({ ...draft, source: e.target.value })} style={inputStyle}>
                  <option value="qtyNew">New Stock (A)</option>
                  <option value="qtyReturn">Return Stock (B)</option>
                </select>
              </Field>
            </div>
            <ItemsEditor items={draft.items} setItems={(items) => setDraft({ ...draft, items })} showPrice={false} inventory={inventory} />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 14 }}>
              <Btn variant="outline" color={C.inkSoft} onClick={() => setDraft(null)}>
                Cancel
              </Btn>
              <Btn color={C.outbound} icon={CheckCircle2} onClick={confirmOutbound}>
                Confirm Shipment
              </Btn>
            </div>
          </div>
        )}

        <div style={{ marginTop: 20, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ fontSize: 13.5, fontWeight: 700, fontFamily: FONT_UI, color: C.ink }}>Recent shipments</div>
          {isAdmin && (
            <Btn color={C.amber} icon={Download} onClick={() => setShowReport(true)}>
              Generate Monthly Report
            </Btn>
          )}
        </div>
        <RecordsList
          records={typeRecords}
          dateField="shipDate"
          sourceLabel={(r) => `${r.boxes || 0} boxes`}
          isAdmin={isAdmin}
          onDeleteClick={setDeleteTarget}
        />
      </div>

      {excelPending && (
        <ColumnMapModal
          headers={excelPending.headers}
          accent={C.outbound}
          showPrice={false}
          onConfirm={handleColumnsConfirmed}
          onCancel={() => setExcelPending(null)}
        />
      )}

      {pendingDuplicate && (
        <DuplicateFileWarning
          fileName={pendingDuplicate.file.name}
          match={pendingDuplicate.match}
          onCancel={() => setPendingDuplicate(null)}
          onProceed={() => {
            const { file, hash } = pendingDuplicate;
            setPendingDuplicate(null);
            processFile(file, hash);
          }}
        />
      )}

      {showReport && <FbaMonthlyReportModal records={typeRecords} onClose={() => setShowReport(false)} />}

      {deleteTarget && (
        <DeleteRecordConfirm record={deleteTarget} onCancel={() => setDeleteTarget(null)} onConfirm={() => performDelete(deleteTarget)} />
      )}
    </div>
  );
}

function FbaMonthlyReportModal({ records, onClose }) {
  const [month, setMonth] = useState(new Date().toISOString().slice(0, 7));
  const shift = (delta) => {
    const [y, m] = month.split("-").map(Number);
    const d = new Date(y, m - 1 + delta, 1);
    setMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
  };
  const monthRecords = records.filter((r) => monthKey(r.shipDate || r.date) === month);

  // Group by Shipment ID
  const shipmentsMap = new Map();
  monthRecords.forEach((r) => {
    const key = r.shipmentId || r.id;
    const cur = shipmentsMap.get(key) || {
      shipmentId: r.shipmentId || "Unspecified",
      shipDate: r.shipDate,
      boxes: 0,
      items: new Map(),
    };
    cur.boxes += Number(r.boxes) || 0;
    r.items.forEach((it) => {
      const c = cur.items.get(it.sku) || { sku: it.sku, name: it.name, qty: 0 };
      c.qty += Number(it.qty) || 0;
      cur.items.set(it.sku, c);
    });
    shipmentsMap.set(key, cur);
  });
  const shipments = Array.from(shipmentsMap.values())
    .map((s) => ({ ...s, items: Array.from(s.items.values()) }))
    .sort((a, b) => (a.shipDate || "").localeCompare(b.shipDate || ""));

  const totalBoxes = shipments.reduce((s, x) => s + x.boxes, 0);
  const totalUnits = shipments.reduce((s, x) => s + x.items.reduce((a, i) => a + i.qty, 0), 0);

  const exportCSV = () => {
    const rows = [["Shipment ID", "Ship date", "SKU", "Product name", "Qty", "Boxes (shipment total)"]];
    shipments.forEach((s) => {
      s.items.forEach((it) => {
        rows.push([s.shipmentId, s.shipDate || "", it.sku, it.name, it.qty, s.boxes]);
      });
    });
    rows.push([]);
    rows.push(["Shipments", shipments.length]);
    rows.push(["Total boxes", totalBoxes]);
    rows.push(["Total units", totalUnits]);
    downloadCSV(rows, `fba-shipment-report-${month}.csv`);
  };

  return (
    <Modal title="Monthly Report — Amazon FBA Shipments" accent={C.amber} onClose={onClose} wide>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 14, marginBottom: 16 }}>
        <button onClick={() => shift(-1)} style={{ border: "none", background: "none", cursor: "pointer", color: C.ink }}>
          <ChevronLeft size={18} />
        </button>
        <span style={{ fontFamily: FONT_MONO, fontWeight: 700, fontSize: 16 }}>{month}</span>
        <button onClick={() => shift(1)} style={{ border: "none", background: "none", cursor: "pointer", color: C.ink }}>
          <ChevronRight size={18} />
        </button>
      </div>

      <div style={{ display: "flex", gap: 12, marginBottom: 16 }}>
        <StatBox label="Shipments" value={shipments.length} />
        <StatBox label="Total boxes" value={totalBoxes} />
        <StatBox label="Total units" value={totalUnits} />
      </div>

      {shipments.length === 0 ? (
        <div style={{ fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI, padding: 12 }}>No shipments this month</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 12, marginBottom: 14, maxHeight: 420, overflowY: "auto" }}>
          {shipments.map((s) => (
            <div key={s.shipmentId} style={{ border: `1px solid ${C.border}`, borderRadius: 5, overflow: "hidden" }}>
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  padding: "8px 12px",
                  background: C.surfaceSoft,
                  fontSize: 12.5,
                  fontFamily: FONT_UI,
                }}
              >
                <span style={{ fontFamily: FONT_MONO, fontWeight: 700, color: C.ink }}>{s.shipmentId}</span>
                <span style={{ color: C.inkSoft }}>
                  {s.shipDate || "—"} · {s.boxes} boxes · {s.items.reduce((a, i) => a + i.qty, 0)} units
                </span>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1.6fr 0.6fr", padding: "6px 12px", fontSize: 11.5, fontWeight: 700, color: C.inkSoft, fontFamily: FONT_UI }}>
                <div>SKU</div>
                <div>Product name</div>
                <div>Qty</div>
              </div>
              {s.items.map((it) => (
                <div key={it.sku} style={{ display: "grid", gridTemplateColumns: "1fr 1.6fr 0.6fr", padding: "5px 12px", borderTop: `1px solid ${C.surfaceSoft}`, fontSize: 12.5, fontFamily: FONT_UI }}>
                  <div style={{ fontFamily: FONT_MONO }}>{it.sku}</div>
                  <div>{it.name}</div>
                  <div style={{ fontFamily: FONT_MONO }}>{it.qty}</div>
                </div>
              ))}
            </div>
          ))}
        </div>
      )}
      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <Btn color={C.amber} icon={Download} onClick={exportCSV} disabled={shipments.length === 0}>
          Export CSV
        </Btn>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------
   OUTBOUND — Order Fulfillment
--------------------------------------------------------------- */
function OutboundOrderFlow({ setView, inventory, saveInventory, outboundRecords, saveOutboundRecords, inboundRecords, aliasMap, saveAliasMap, ignoredSkus, showToast, isAdmin }) {
  const [busy, setBusy] = useState(false);
  const [draft, setDraft] = useState(null);
  const [excelPending, setExcelPending] = useState(null);
  const [pendingDuplicate, setPendingDuplicate] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);

  const typeRecords = outboundRecords.filter((r) => (r.type || "order") === "order");

  const performDelete = (record) => {
    const field = record.source === "New Stock" ? "qtyNew" : "qtyReturn";
    saveInventory(applySignedQty(inventory, record.items, field, 1));
    saveOutboundRecords(outboundRecords.filter((r) => r.id !== record.id));
    logAudit("outbound-order-delete", `Deleted order fulfillment record from "${record.fileName}" (${record.platform}) — reversed ${record.items.length} item rows`);
    setDeleteTarget(null);
    showToast("Record deleted, inventory reversed");
  };

  const handleFile = async (file) => {
    let hash = null;
    try {
      hash = await hashFile(file);
    } catch (e) {}
    if (hash) {
      const match = findDuplicateByHash(hash, inboundRecords, outboundRecords);
      if (match) {
        setPendingDuplicate({ file, hash, match });
        return;
      }
    }
    await processFile(file, hash);
  };

  const processFile = async (file, hash) => {
    if (isAmazonPackingSlipFile(file)) {
      try {
        const parsedItems = await parseAmazonPackingSlip(file);
        if (parsedItems.length === 0) {
          showToast("Couldn't find any SKU rows in this page — check the file", "error");
          return;
        }
        const resolved = resolveSkus(parsedItems, inventory, aliasMap, ignoredSkus);
        setDraft({
          platform: "Amazon",
          shipDate: dateFromFilename(file.name) || todayISO(),
          source: "qtyNew",
          items: resolved,
          fileName: file.name,
          fileHash: hash,
        });
        showToast(`Detected ${resolved.length} SKUs from the Amazon packing slip`);
      } catch (e) {
        showToast("Couldn't parse this page — check the file format", "error");
      }
      return;
    }
    if (isSpreadsheetFile(file)) {
      try {
        const { grid, rows, headers } = await parseSpreadsheetFile(file);
        const sections = findSkuSections(grid);
        const sectionItems = itemsFromSkuSections(sections);
        if (sectionItems.length > 0) {
          const resolved = resolveSkus(sectionItems, inventory, aliasMap, ignoredSkus);
          setDraft({
            platform: "",
            shipDate: todayISO(),
            source: "qtyNew",
            items: resolved,
            fileName: file.name,
            fileHash: hash,
          });
          showToast(`Detected ${sectionItems.length} SKUs automatically from ${sections.length > 1 ? sections.length + " sections" : "the sheet"}`);
          return;
        }
        if (rows.length === 0) {
          showToast("The spreadsheet has no data rows", "error");
          return;
        }
        setExcelPending({ rows, headers, fileName: file.name, fileHash: hash });
      } catch (e) {
        showToast("Couldn't parse the spreadsheet — check the file format", "error");
      }
      return;
    }
    setBusy(true);
    try {
      const extracted = await extractFromFile(file, "outbound");
      const rawItems = (extracted.items || []).map((it) => ({
        sku: it.sku_guess || "",
        name: it.name || "",
        qty: Number(it.qty) || 0,
      }));
      // Filename convention "{Platform} {YYYY-MM-DD}" (e.g. eBay screenshots)
      // takes priority over whatever the AI guessed from the image content.
      const fnMeta = parsePlatformDateFilename(file.name);
      setDraft({
        platform: fnMeta?.platform || extracted.platform || "",
        shipDate: fnMeta?.date || extracted.ship_date || todayISO(),
        source: "qtyNew",
        items: resolveSkus(rawItems, inventory, aliasMap, ignoredSkus),
        fileName: file.name,
        fileHash: hash,
      });
      showToast("Extraction complete — please review and confirm");
    } catch (e) {
      showToast(e.message || "Recognition failed — please retry or enter manually", "error");
      setDraft({ platform: "", shipDate: todayISO(), source: "qtyNew", items: [{ sku: "", name: "", qty: 1 }], fileName: file.name, fileHash: hash });
    } finally {
      setBusy(false);
    }
  };

  const handleColumnsConfirmed = ({ skuCol, nameCol, qtyCol }) => {
    const rawItems = excelPending.rows
      .map((r) => ({
        sku: String(r[skuCol] ?? "").trim(),
        name: nameCol ? String(r[nameCol] ?? "").trim() : "",
        qty: Number(r[qtyCol]) || 0,
      }))
      .filter((it) => it.sku && it.qty > 0);
    const resolved = resolveSkus(rawItems, inventory, aliasMap, ignoredSkus);
    setDraft({
      platform: "",
      shipDate: todayISO(),
      source: "qtyNew",
      items: resolved,
      fileName: excelPending.fileName,
      fileHash: excelPending.fileHash,
    });
    setExcelPending(null);
    showToast(`Generated ${resolved.length} item rows from the spreadsheet — please review`);
  };

  const confirmOutbound = () => {
    const cleanItems = draft.items.filter((it) => it.sku.trim() && Number(it.qty) > 0);
    if (cleanItems.length === 0) {
      showToast("No valid item rows", "error");
      return;
    }
    const field = draft.source;
    const map = new Map(inventory.map((x) => [x.sku, { ...x }]));
    let shortage = false;
    cleanItems.forEach((it) => {
      const sku = it.sku.trim();
      const cur = map.get(sku) || { sku, name: it.name || sku, qtyNew: 0, qtyReturn: 0 };
      const remaining = (Number(cur[field]) || 0) - Number(it.qty);
      if (remaining < 0) shortage = true;
      cur[field] = remaining;
      map.set(sku, cur);
    });
    saveInventory(Array.from(map.values()));

    const nextAliases = { ...aliasMap };
    let aliasChanged = false;
    const mapGroups = new Map();
    cleanItems.forEach((it) => {
      if (!it.mappedFrom) return;
      const norm = normalizeSku(it.mappedFrom);
      if (!mapGroups.has(norm)) mapGroups.set(norm, { raw: it.mappedFrom, components: [] });
      mapGroups.get(norm).components.push({ sku: it.sku.trim(), qty: Number(it.qtyMultiplier) || 1 });
    });
    mapGroups.forEach((val, norm) => {
      const existing = nextAliases[norm] ? aliasComponents(nextAliases[norm]) : null;
      const changed =
        !existing ||
        existing.length !== val.components.length ||
        existing.some((c, idx) => c.sku !== val.components[idx].sku || Number(c.qty) !== Number(val.components[idx].qty));
      if (changed) {
        nextAliases[norm] = { raw: val.raw, components: val.components };
        aliasChanged = true;
      }
    });
    if (aliasChanged) saveAliasMap(nextAliases);

    const record = {
      id: Date.now().toString(),
      date: new Date().toISOString(),
      shipDate: draft.shipDate,
      platform: draft.platform || "Unspecified",
      fileName: draft.fileName,
      fileHash: draft.fileHash || null,
      source: field === "qtyNew" ? "New Stock" : "Return Stock",
      items: cleanItems,
      type: "order",
    };
    saveOutboundRecords([record, ...outboundRecords]);
    logAudit("outbound-order-confirm", `Order fulfillment confirmed: ${cleanItems.length} item rows from "${draft.fileName}" (${record.platform})`);
    setDraft(null);
    showToast(shortage ? "Outbound confirmed, but some items are now negative — please check" : "Outbound confirmed, inventory updated", shortage ? "error" : "success");
  };

  return (
    <div>
      <TopBar title="Order Fulfillment" subtitle="Upload packing slips or order exports to log stock shipped to customers" onBack={() => setView("outbound")} accent={C.outboundOrder} />
      <div style={{ padding: 18 }}>
        {!draft && (
          <UploadBox
            accent={C.outboundOrder}
            accentSoft={C.outboundOrderSoft}
            label="Click or drag to upload a shipping document"
            hint="Supports image, PDF, Excel/CSV, or an Amazon packing-slip page (.html) · auto-extracts items and quantities"
            onFile={handleFile}
            busy={busy}
          />
        )}

        {draft && (
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 16 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <div style={{ fontSize: 13, fontFamily: FONT_UI, color: C.inkSoft, display: "flex", alignItems: "center", gap: 6 }}>
                <FileText size={14} /> {draft.fileName}
              </div>
              <Stamp label="Pending review" color={C.amber} bg={C.amberSoft} />
            </div>
            <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
              <Field label="Platform / Customer" style={{ flex: 1, minWidth: 160 }}>
                <input value={draft.platform} onChange={(e) => setDraft({ ...draft, platform: e.target.value })} style={inputStyle} />
              </Field>
              <Field label="Ship date" style={{ width: 160 }}>
                <input type="date" value={draft.shipDate} onChange={(e) => setDraft({ ...draft, shipDate: e.target.value })} style={inputStyle} />
              </Field>
              <Field label="Outbound source" style={{ width: 160 }}>
                <select value={draft.source} onChange={(e) => setDraft({ ...draft, source: e.target.value })} style={inputStyle}>
                  <option value="qtyNew">New Stock (A)</option>
                  <option value="qtyReturn">Return Stock (B)</option>
                </select>
              </Field>
            </div>
            <ItemsEditor items={draft.items} setItems={(items) => setDraft({ ...draft, items })} showPrice={false} inventory={inventory} />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 14 }}>
              <Btn variant="outline" color={C.inkSoft} onClick={() => setDraft(null)}>
                Cancel
              </Btn>
              <Btn color={C.outboundOrder} icon={CheckCircle2} onClick={confirmOutbound}>
                Confirm Outbound
              </Btn>
            </div>
          </div>
        )}

        <div style={{ marginTop: 20, fontSize: 13.5, fontWeight: 700, fontFamily: FONT_UI, color: C.ink }}>Recent outbound records</div>
        <RecordsList
          records={typeRecords}
          dateField="shipDate"
          sourceLabel={(r) => r.source}
          isAdmin={isAdmin}
          onDeleteClick={setDeleteTarget}
        />
      </div>

      {excelPending && (
        <ColumnMapModal
          headers={excelPending.headers}
          accent={C.outboundOrder}
          showPrice={false}
          onConfirm={handleColumnsConfirmed}
          onCancel={() => setExcelPending(null)}
        />
      )}

      {pendingDuplicate && (
        <DuplicateFileWarning
          fileName={pendingDuplicate.file.name}
          match={pendingDuplicate.match}
          onCancel={() => setPendingDuplicate(null)}
          onProceed={() => {
            const { file, hash } = pendingDuplicate;
            setPendingDuplicate(null);
            processFile(file, hash);
          }}
        />
      )}

      {deleteTarget && (
        <DeleteRecordConfirm record={deleteTarget} onCancel={() => setDeleteTarget(null)} onConfirm={() => performDelete(deleteTarget)} />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   ADMIN — team accounts + audit log
--------------------------------------------------------------- */
function AdminView({ setView, auth }) {
  const [users, setUsersList] = useState([]);
  const [auditLog, setAuditLog] = useState([]);
  const [loadingUsers, setLoadingUsers] = useState(true);
  const [loadingAudit, setLoadingAudit] = useState(true);
  const [showAddUser, setShowAddUser] = useState(false);
  const [newUsername, setNewUsername] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [newRole, setNewRole] = useState("guest");
  const [error, setError] = useState("");

  const loadUsers = async () => {
    setLoadingUsers(true);
    try {
      const data = await apiGet("/api/users");
      setUsersList(data.users || []);
    } catch (e) {}
    setLoadingUsers(false);
  };
  const loadAudit = async () => {
    setLoadingAudit(true);
    try {
      const data = await apiGet("/api/audit-log");
      setAuditLog(data.entries || []);
    } catch (e) {}
    setLoadingAudit(false);
  };

  useEffect(() => {
    loadUsers();
    loadAudit();
  }, []);

  const addUser = async () => {
    setError("");
    if (!newUsername.trim() || !newPassword) {
      setError("Username and password are required");
      return;
    }
    const data = await apiPost("/api/users", { username: newUsername.trim(), password: newPassword, role: newRole });
    if (data.error) {
      setError(data.error);
      return;
    }
    setNewUsername("");
    setNewPassword("");
    setNewRole("guest");
    setShowAddUser(false);
    loadUsers();
    loadAudit();
  };

  const removeUser = async (username) => {
    const data = await apiDelete(`/api/users/${encodeURIComponent(username)}`);
    if (data.error) {
      setError(data.error);
      return;
    }
    loadUsers();
    loadAudit();
  };

  return (
    <div>
      <TopBar title="Admin" subtitle="Team accounts and audit log" onBack={() => setView("home")} accent={C.amber} />
      <div style={{ padding: 18 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
          <ShieldCheck size={16} color={C.amber} />
          <div style={{ fontSize: 14, fontWeight: 700, fontFamily: FONT_UI, color: C.ink }}>Team accounts</div>
          <Btn color={C.amber} icon={Plus} onClick={() => setShowAddUser(true)} style={{ marginLeft: "auto", fontSize: 12.5, padding: "6px 10px" }}>
            Add account
          </Btn>
        </div>
        {loadingUsers ? (
          <div style={{ fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI }}>Loading…</div>
        ) : (
          <div style={{ border: `1px solid ${C.border}`, borderRadius: 5, overflow: "hidden", marginBottom: 24 }}>
            <div style={{ display: "grid", gridTemplateColumns: "1.3fr 0.8fr 1fr 30px", padding: "8px 12px", background: C.surfaceSoft, fontSize: 11.5, fontWeight: 700, color: C.inkSoft, fontFamily: FONT_UI }}>
              <div>Username</div>
              <div>Role</div>
              <div>Created</div>
              <div></div>
            </div>
            {users.map((u) => (
              <div
                key={u.username}
                style={{
                  display: "grid",
                  gridTemplateColumns: "1.3fr 0.8fr 1fr 30px",
                  padding: "7px 12px",
                  borderTop: `1px solid ${C.surfaceSoft}`,
                  fontSize: 12.5,
                  fontFamily: FONT_UI,
                  alignItems: "center",
                }}
              >
                <div style={{ fontFamily: FONT_MONO }}>{u.username}</div>
                <div>
                  <Stamp label={u.role} color={u.role === "admin" ? C.amber : C.inkSoft} bg={C.surface} />
                </div>
                <div style={{ color: C.inkSoft, fontFamily: FONT_MONO }}>{(u.createdAt || "").slice(0, 10)}</div>
                {u.username !== auth.username && (
                  <button onClick={() => removeUser(u.username)} style={{ border: "none", background: "none", cursor: "pointer", color: C.inkSoft }}>
                    <Trash2 size={13} />
                  </button>
                )}
              </div>
            ))}
          </div>
        )}
        {error && <div style={{ color: C.danger, fontSize: 12.5, fontFamily: FONT_UI, marginTop: -14, marginBottom: 14 }}>{error}</div>}

        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
          <History size={16} color={C.amber} />
          <div style={{ fontSize: 14, fontWeight: 700, fontFamily: FONT_UI, color: C.ink }}>Audit log</div>
          <Btn variant="outline" color={C.inkSoft} onClick={loadAudit} style={{ marginLeft: "auto", fontSize: 12, padding: "5px 9px" }}>
            Refresh
          </Btn>
        </div>
        {loadingAudit ? (
          <div style={{ fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI }}>Loading…</div>
        ) : auditLog.length === 0 ? (
          <div style={{ fontSize: 13, color: C.inkSoft, fontFamily: FONT_UI }}>No activity recorded yet.</div>
        ) : (
          <div style={{ border: `1px solid ${C.border}`, borderRadius: 5, maxHeight: 420, overflowY: "auto" }}>
            {auditLog.map((e) => (
              <div key={e.id} style={{ padding: "9px 12px", borderTop: `1px solid ${C.surfaceSoft}`, fontSize: 12.5, fontFamily: FONT_UI }}>
                <div style={{ display: "flex", justifyContent: "space-between", gap: 10 }}>
                  <span style={{ fontWeight: 700, color: C.ink }}>{e.username}</span>
                  <span style={{ fontFamily: FONT_MONO, color: C.inkSoft, fontSize: 11.5, whiteSpace: "nowrap" }}>
                    {new Date(e.timestamp).toLocaleString()}
                  </span>
                </div>
                <div style={{ color: C.inkSoft, marginTop: 2 }}>{e.summary}</div>
              </div>
            ))}
          </div>
        )}
      </div>

      {showAddUser && (
        <Modal title="Add team account" accent={C.amber} onClose={() => setShowAddUser(false)}>
          <Field label="Username">
            <input value={newUsername} onChange={(e) => setNewUsername(e.target.value)} style={inputStyle} />
          </Field>
          <Field label="Password">
            <input type="password" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} style={inputStyle} />
          </Field>
          <Field label="Role">
            <select value={newRole} onChange={(e) => setNewRole(e.target.value)} style={inputStyle}>
              <option value="guest">Guest — full app access, inventory-level export only</option>
              <option value="admin">Admin — full access including all reports</option>
            </select>
          </Field>
          {error && <div style={{ color: C.danger, fontSize: 12.5, marginBottom: 10 }}>{error}</div>}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 10 }}>
            <Btn variant="outline" color={C.inkSoft} onClick={() => setShowAddUser(false)}>
              Cancel
            </Btn>
            <Btn color={C.amber} onClick={addUser}>
              Create account
            </Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   ROOT APP
--------------------------------------------------------------- */
export default function WarehouseApp() {
  const [auth, setAuth] = useState(null);
  const [authChecked, setAuthChecked] = useState(false);
  const [view, setView] = useState("home");
  const [loaded, setLoaded] = useState(false);
  const [inventory, setInventory] = useState([]);
  const [inboundRecords, setInboundRecords] = useState([]);
  const [outboundRecords, setOutboundRecords] = useState([]);
  const [skuAliases, setSkuAliases] = useState({});
  const [ignoredSkus, setIgnoredSkus] = useState([]);
  const [toast, setToast] = useState(null);

  const showToast = useCallback((msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3000);
  }, []);

  const handleLogin = useCallback((data) => {
    setAuthToken(data.token);
    const authObj = { token: data.token, username: data.username, role: data.role };
    localStorage.setItem("wh-auth", JSON.stringify(authObj));
    setLoaded(false);
    setView("home");
    setAuth(authObj);
  }, []);

  const handleLogout = useCallback(() => {
    apiPost("/api/auth/logout", {}).catch(() => {});
    localStorage.removeItem("wh-auth");
    setAuthToken(null);
    setAuth(null);
  }, []);

  // Restore a saved session on load, and log out automatically if the
  // server ever rejects a request as unauthenticated (expired/deleted session).
  useEffect(() => {
    try {
      const stored = localStorage.getItem("wh-auth");
      if (stored) {
        const parsed = JSON.parse(stored);
        setAuthToken(parsed.token);
        setAuth(parsed);
      }
    } catch (e) {}
    setAuthChecked(true);

    const onExpired = () => {
      localStorage.removeItem("wh-auth");
      setAuthToken(null);
      setAuth(null);
    };
    window.addEventListener("wh-auth-expired", onExpired);
    return () => window.removeEventListener("wh-auth-expired", onExpired);
  }, []);

  // All data is shared across your team — anyone signed in reads and
  // writes the same inventory, records, and SKU mapping rules.
  useEffect(() => {
    if (!auth) return;
    (async () => {
      try {
        const inv = await kvGet("inventory").catch(() => null);
        if (inv) setInventory(JSON.parse(inv.value));
      } catch (e) {}
      try {
        const ib = await kvGet("inbound-records").catch(() => null);
        if (ib) setInboundRecords(JSON.parse(ib.value));
      } catch (e) {}
      try {
        const ob = await kvGet("outbound-records").catch(() => null);
        if (ob) setOutboundRecords(JSON.parse(ob.value));
      } catch (e) {}
      try {
        const sa = await kvGet("sku-aliases").catch(() => null);
        if (sa) setSkuAliases(JSON.parse(sa.value));
      } catch (e) {}
      try {
        const ig = await kvGet("ignored-skus").catch(() => null);
        if (ig) {
          setIgnoredSkus(JSON.parse(ig.value));
        } else {
          // First run: seed with known non-inventory line items (shipping fees, service charges).
          const defaults = ["SPESE-TRASPORTO", "SAVE-EL"];
          setIgnoredSkus(defaults);
          kvSet("ignored-skus", JSON.stringify(defaults)).catch(() => {});
        }
      } catch (e) {}
      setLoaded(true);
    })();
  }, [auth]);

  const saveInventory = useCallback(async (next) => {
    setInventory(next);
    try {
      const ok = await kvSet("inventory", JSON.stringify(next));
      if (!ok) throw new Error();
    } catch (e) {
      showToast("Failed to save inventory — please retry", "error");
    }
  }, [showToast]);

  const saveInboundRecords = useCallback(async (next) => {
    setInboundRecords(next);
    try {
      const ok = await kvSet("inbound-records", JSON.stringify(next));
      if (!ok) throw new Error();
    } catch (e) {
      showToast("Failed to save inbound records", "error");
    }
  }, [showToast]);

  const saveOutboundRecords = useCallback(async (next) => {
    setOutboundRecords(next);
    try {
      const ok = await kvSet("outbound-records", JSON.stringify(next));
      if (!ok) throw new Error();
    } catch (e) {
      showToast("Failed to save outbound records", "error");
    }
  }, [showToast]);

  const saveAliasMap = useCallback(async (next) => {
    setSkuAliases(next);
    try {
      const ok = await kvSet("sku-aliases", JSON.stringify(next));
      if (!ok) throw new Error();
    } catch (e) {
      showToast("Failed to save SKU mapping rules", "error");
    }
  }, [showToast]);

  const saveIgnoredSkus = useCallback(async (next) => {
    setIgnoredSkus(next);
    try {
      const ok = await kvSet("ignored-skus", JSON.stringify(next));
      if (!ok) throw new Error();
    } catch (e) {
      showToast("Failed to save ignored codes", "error");
    }
  }, [showToast]);

  const globalStyle = (
    <style>{`
      @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
      * { box-sizing: border-box; }
      input:focus, select:focus { border-color: ${C.inventory} !important; }
      ::-webkit-scrollbar { width: 8px; height: 8px; }
      ::-webkit-scrollbar-thumb { background: ${C.borderStrong}; border-radius: 4px; }
    `}</style>
  );

  if (!authChecked) {
    return (
      <div style={{ fontFamily: FONT_UI, background: C.bg, minHeight: "100vh" }}>
        {globalStyle}
      </div>
    );
  }
  if (!auth) {
    return (
      <div style={{ fontFamily: FONT_UI, background: C.bg, minHeight: "100vh" }}>
        {globalStyle}
        <LoginScreen onLogin={handleLogin} />
      </div>
    );
  }

  const isAdmin = auth.role === "admin";

  return (
    <div style={{ fontFamily: FONT_UI, background: C.bg, minHeight: "100%", color: C.ink }}>
      {globalStyle}

      {!loaded ? (
        <div style={{ padding: 60, textAlign: "center", color: C.inkSoft }}>
          <Loader2 size={22} style={{ animation: "spin 1s linear infinite" }} />
          <div style={{ marginTop: 10, fontSize: 13 }}>Loading…</div>
        </div>
      ) : view === "home" ? (
        <Home
          setView={setView}
          inventory={inventory}
          inboundRecords={inboundRecords}
          outboundRecords={outboundRecords}
          auth={auth}
          onLogout={handleLogout}
        />
      ) : view === "inventory" ? (
        <InventoryView
          setView={setView}
          inventory={inventory}
          saveInventory={saveInventory}
          showToast={showToast}
          aliasMap={skuAliases}
          saveAliasMap={saveAliasMap}
          ignoredSkus={ignoredSkus}
          saveIgnoredSkus={saveIgnoredSkus}
          inboundRecords={inboundRecords}
          outboundRecords={outboundRecords}
          isAdmin={isAdmin}
        />
      ) : view === "inbound" ? (
        <InboundHub setView={setView} inboundRecords={inboundRecords} />
      ) : view === "inbound-new" || view === "inbound-return" ? (
        <InboundFlow
          type={view === "inbound-new" ? "new" : "return"}
          inventory={inventory}
          saveInventory={saveInventory}
          inboundRecords={inboundRecords}
          saveInboundRecords={saveInboundRecords}
          outboundRecords={outboundRecords}
          aliasMap={skuAliases}
          saveAliasMap={saveAliasMap}
          ignoredSkus={ignoredSkus}
          showToast={showToast}
          setView={setView}
          isAdmin={isAdmin}
        />
      ) : view === "outbound" ? (
        <OutboundHub setView={setView} outboundRecords={outboundRecords} />
      ) : view === "outbound-fba" ? (
        <OutboundFbaFlow
          setView={setView}
          inventory={inventory}
          saveInventory={saveInventory}
          outboundRecords={outboundRecords}
          saveOutboundRecords={saveOutboundRecords}
          inboundRecords={inboundRecords}
          aliasMap={skuAliases}
          saveAliasMap={saveAliasMap}
          ignoredSkus={ignoredSkus}
          showToast={showToast}
          isAdmin={isAdmin}
        />
      ) : view === "outbound-order" ? (
        <OutboundOrderFlow
          setView={setView}
          inventory={inventory}
          saveInventory={saveInventory}
          outboundRecords={outboundRecords}
          saveOutboundRecords={saveOutboundRecords}
          inboundRecords={inboundRecords}
          aliasMap={skuAliases}
          saveAliasMap={saveAliasMap}
          ignoredSkus={ignoredSkus}
          showToast={showToast}
          isAdmin={isAdmin}
        />
      ) : view === "admin" && isAdmin ? (
        <AdminView setView={setView} auth={auth} />
      ) : (
        <Home
          setView={setView}
          inventory={inventory}
          inboundRecords={inboundRecords}
          outboundRecords={outboundRecords}
          auth={auth}
          onLogout={handleLogout}
        />
      )}

      <Toast toast={toast} />
    </div>
  );
}
