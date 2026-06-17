# Hypertrophy App (v5.1)

A personal, science-based hypertrophy training logger — single-page React app (Vite), no backend.

## Program — Upper-Focus 4-Day (Push / Pull / Lower / Upper)
Built for upper-body size with legs kept moderate, machines/cables only. Chest days are
isolation-led (pre-exhaust) so the chest does the work before pressing, and triceps come
last. Each exercise carries a locked Garmin Connect name and a starting weight.

Rotation: **Push → Pull → Lower → Upper**, cycle on off-days.

## Deploy to Netlify

### Option 1 — Drag & drop (easiest)
1. `npm install` then `npm run build`
2. Drag the resulting `dist/` folder into [app.netlify.com/drop](https://app.netlify.com/drop)
   (or onto an existing site's Deploys tab — same URL keeps your saved data, which auto-migrates)

### Option 2 — Git deploy (recommended for ongoing updates)
1. This repo is on GitHub — in Netlify: "Add new site" → "Import from Git" → pick it.
2. Settings auto-detected from `netlify.toml` (build `npm run build`, publish `dist`).
3. Every push to `main` auto-deploys.

## Local development
```bash
npm install
npm run dev      # http://localhost:5173
npm run build    # outputs to dist/
```

## How data is stored
- `src/storage-shim.js` provides `window.storage` backed by `localStorage` — persists across
  refreshes/restarts on the same browser. It does NOT sync across devices: use **Backup to
  Clipboard** (Settings) and paste on the other device.
- History is keyed by **session** (`date__day`) so switching days never corrupts logs.
- `migrateHistory()` upgrades older date-keyed backups on load/import/restore.
- Clearing browser data / incognito wipes it — always keep a backup (Notes, email, Drive).

## Structure
- `src/App.jsx` — the whole app (program data, exercise guides, logging, history, progress, settings)
- `src/main.jsx`, `index.html`, `vite.config.js`, `public/assets/` + `assets/` (logo)

## Tech stack
React 18 · Vite 5 · lucide-react (icons) · xlsx (Excel import/export) · fully client-side.
