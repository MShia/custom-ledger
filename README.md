# Construction Project Finance Ledger

A dashboard for tracking the finances of a construction project between an
**owner** (who funds the work) and a **contractor** (who manages and spends).
It records owner payments and contractor expenses, computes the running
balance (funds held by the contractor, or the amount owed to them), and
provides category and personnel analytics plus a saved report registry.

All data is entered in-app and saved automatically in your browser. No backend,
no external file imports.

## Requirements

- Node.js 18 or newer (Node 20+ recommended)

## Setting your project defaults (name, owner, contractor, currency)

Open `src/App.jsx` and edit the **`PROJECT_DEFAULTS`** block near the top of the
file (around line 25):

```js
const PROJECT_DEFAULTS = {
  projectName: "New Construction Project",
  ownerName: "Owner",
  contractorName: "Contractor",
  currency: "Rs",
};
```

Change the strings, then rebuild/redeploy. These values show for any new
visitor whose browser has no saved data yet, so they appear immediately on a
fresh deployment.

Note: if you've already used the app in your own browser, your saved details
override these defaults on that browser. To preview the new defaults, use
**Settings → Reset all data**, open the site in a private/incognito window, or
clear the site's data.

## Setup

From this folder:

```bash
npm install
npm run dev
```

Then open the URL Vite prints (usually http://localhost:5173).

To open the project in VS Code from here: `code .`

## Build for production

```bash
npm run build      # outputs to dist/
npm run preview    # serves the production build locally
```

The `dist/` folder is a static site — deploy it to Vercel, Netlify, GitHub
Pages, or any static host. No server is required.

## How your data is stored

Inside the Claude app this dashboard uses a built-in storage API. On your own
machine, `src/main.jsx` installs a small drop-in replacement backed by the
browser's `localStorage`, so your payments, expenses, categories, personnel,
and generated reports persist across refreshes on that browser.

Because the data lives in the browser, it is specific to that browser/profile.
Use **Settings → Download backup (JSON)** to keep a copy or move data between
machines.

## Sharing data with another person

Because storage is per-browser, deploying the site does **not** share your data
with people who open it — everyone gets their own empty copy. There are two
built-in ways to hand your data to a third person:

### Option A — send them a backup file (snapshot handoff)

1. You: **Settings → Download backup (JSON)**. Send that `.json` file (email,
   drive, chat — anything).
2. Them: open the deployed site → **Settings → Import backup (JSON)** → pick the
   file. The app loads your data and saves it into *their* browser, and they
   continue from there.

This is a one-time copy: after import, your two copies are independent. If you
keep editing and want them current again, send a fresh backup to re-import.

### Option B — seed the deployment (everyone starts with the same data)

Bake a starting dataset into the build so anyone opening the deployed site
begins with it already loaded (only when their browser has no saved data yet):

1. **Settings → Download backup (JSON)**.
2. Open `src/seed.json`, delete the `null`, and paste in the full contents of
   that backup file.
3. Redeploy. New visitors now start from that data; once they make changes,
   their edits are saved to their own browser and the seed no longer overrides.

> Note: neither option is live collaboration. Two people editing will not see
> each other's changes. For shared, always-in-sync data across people you need a
> hosted database (e.g. Supabase or Firebase) wired into the storage layer in
> `src/main.jsx` — ask and this can be added.

## Features

- Owner → contractor payment ledger with running balance
- Expense entry with construction categories and sub-categories
- Search and filter expenses by category, person, and text
- Analytics: spend by category, monthly cash flow, per-person personnel spend
- Report snapshots saved to an in-app registry, exportable as a formatted
  **Word (.docx)** document or CSV
- A polished Word report (summary, category and personnel analytics, monthly
  cash flow, and a full transaction ledger appendix) generated in the browser
- Full-ledger CSV export and JSON backup
- Editable project details, categories, personnel, and currency symbol

## Project structure

```
construction-ledger/
├── index.html
├── package.json
├── vite.config.js
├── src/
│   ├── App.jsx        the entire dashboard
│   ├── main.jsx       entry point + localStorage storage shim
│   ├── seed.json      optional baked-in starter dataset (null = none)
│   └── index.css      Tailwind import + base page styles
```

## Tech

React 19, Vite, Tailwind CSS v4, Recharts, lucide-react, docx.
