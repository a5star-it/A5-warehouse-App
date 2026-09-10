# A5 STAR · Warehouse Ledger — self-hosted version

This is a standalone version of the warehouse in/outbound app: a Node.js
(Express) backend + React frontend, packaged so it can be deployed to
Railway and shared with your whole team at one URL.

## What changed from the Claude-artifact version

- **Data storage**: now a JSON file on the server (`data.json`) instead of
  Claude's built-in storage. Everyone who opens the URL shares the same
  data — same behavior as before, just running on your own infrastructure.
- **AI invoice/document recognition**: now calls Anthropic's API using
  **your own API key**, kept only on the server (`ANTHROPIC_API_KEY`
  environment variable) — the browser never sees it.
- **No login screen** — anyone with the URL can use and edit the data.
  Good enough for an internal team tool on a private URL; see
  "Locking it down later" below if you want more.

Everything else — the SKU mapping rules, the FBA shipment parsing, the
Amazon packing-slip parsing, the monthly reports, Restock Insights — is the
exact same code, unchanged.

---

## 1. Get an Anthropic API key (5 minutes)

1. Go to **console.anthropic.com** (this is a different account system
   than claude.ai — sign up if you don't have a console account).
2. **Settings → Billing** — add a payment method and add some prepaid
   credit (e.g. $5 to start; API usage is billed separately from any
   claude.ai subscription).
3. **Settings → API Keys → Create Key** — name it something like
   `warehouse-app`, then **copy the key immediately** (it starts with
   `sk-ant-...` and is only shown once — if you lose it, delete it and
   make a new one).

Keep this key somewhere safe for step 4 below.

---

## 2. Put the code on GitHub

Railway deploys from a GitHub repository, so the code needs to live there
first.

1. Create a free GitHub account if you don't have one (github.com).
2. Create a new, empty repository (e.g. `a5star-warehouse`).
3. Upload this whole folder to it. Easiest way if you're not familiar with
   git: on the repo's GitHub page, click **Add file → Upload files**, and
   drag in everything from this folder (keep the folder structure — the
   `src` folder should stay a folder).

---

## 3. Deploy on Railway

1. Go to **railway.com**, sign up (GitHub sign-in is easiest).
2. **New Project → Deploy from GitHub repo** → pick the repository you
   just created. Railway will detect it's a Node.js project automatically
   (it reads `package.json`) and run `npm install`, then `npm run build`,
   then `npm start`.
3. **Add persistent storage (important — do this before real use):**
   - Open your new service → **Volumes** tab → **Add Volume**.
   - Set the mount path to `/data` (any path works, but the server already
     knows to look for `RAILWAY_VOLUME_MOUNT_PATH` automatically, so `/data`
     needs no extra configuration).
   - Without this step, your data is on the container's local disk and
     **will be wiped every time you redeploy**.
4. **Set environment variables:** open the service → **Variables** tab →
   add:
   - `ANTHROPIC_API_KEY` = the key from step 1
   - (leave `PORT` and `DATA_FILE` alone — Railway and the volume step
     above already handle those)
5. **Get a public URL:** service → **Settings → Networking → Generate
   Domain**. Railway gives you a `*.up.railway.app` address — that's the
   link you share with your team.
6. Open the link. You should see the same three cards (Inventory / Inbound
   / Outbound) as before.

Every future change: push updated code to the same GitHub repo, and
Railway redeploys automatically.

---

## 4. Testing locally first (optional, recommended)

If you want to try it on your own computer before deploying:

```bash
npm install
cp .env.example .env
# edit .env and paste your ANTHROPIC_API_KEY
npm run dev
```

This runs the frontend (Vite, hot-reload) and backend together — open
`http://localhost:5173`.

To test the exact production build:

```bash
npm run build
npm start
```

then open `http://localhost:3000`.

---

## Locking it down later

Right now, anyone with the URL can open and edit everything — same as the
shared Claude artifact. If later you want to restrict it to your team
only, a few options, roughly easiest → most robust:

- Don't publish the URL anywhere public; treat it like an unlisted link.
- Put it behind your office/company VPN or a Railway private network.
- Add simple shared-password protection (a small code change — ask me and
  I can add it).
- A real login system with per-person accounts (a bigger change).

## Costs to expect

- **Railway**: usage-based; a small internal tool like this typically
  costs a few dollars a month (check Railway's current pricing).
- **Anthropic API**: pay-as-you-go per document processed. Cost per
  invoice/packing-slip scan is small (well under $0.01–$0.05 typically for
  a single-page document with the model this app uses), but keep an eye
  on console.anthropic.com's usage page, especially in the first weeks.

## If something breaks

- **Railway → your service → Deployments** shows build logs — read the
  error at the bottom if a deploy fails.
- **Railway → your service → Logs (Observability)** shows runtime logs
  from the running server — useful if the app loads but a feature (like
  AI extraction) fails.
- Common first-deploy issues: forgot to add `ANTHROPIC_API_KEY` (AI
  extraction will fail with a clear server error), or forgot the volume
  (data disappears after a redeploy).
