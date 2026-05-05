# Royal Glass — Work Order Dashboard

Google Apps Script project managed with clasp and GitHub.

---

## Environments

| Environment | Spreadsheet | Branch |
|---|---|---|
| Staging | Staging WO Dashboard | `staging` |
| Live | Work Order Dashboard | `main` |

**Rule: never deploy straight to live. Always test on staging first.**

---

## First-Time Setup

### 1. Install clasp globally

```bash
npm install -g @google/clasp
```

### 2. Install project dependencies

```bash
npm install
```

### 3. Log in to Google

```bash
clasp login
```

This opens a browser. Sign in with the same Google account that owns the spreadsheets.
Your auth token saves to `~/.clasprc.json` — this is gitignored and stays on your machine only.

### 4. Enable Apps Script API

Go to: https://script.google.com/home/usersettings

Turn **Google Apps Script API** ON.

### 5. Push code to staging first

```bash
npm run deploy:staging
```

### 6. Test in the staging spreadsheet

Open Staging WO Dashboard → 🔧 Royal Glass menu → run through functions.

### 7. When ready, push to live

```bash
npm run deploy:live
```

---

## Daily Workflow

### Making a change

```bash
# 1. Switch to staging branch
git checkout staging

# 2. Make your changes to src/Code.js

# 3. Deploy to staging
npm run deploy:staging

# 4. Test in Staging WO Dashboard

# 5. Commit your changes
git add .
git commit -m "describe what you changed"
git push origin staging

# 6. Merge to main and deploy live
git checkout main
git merge staging
npm run deploy:live
git push origin main
```

---

## Commands Reference

| Command | What it does |
|---|---|
| `npm run deploy:staging` | Push current code to Staging WO Dashboard |
| `npm run deploy:live` | Push current code to Live Work Order Dashboard |
| `npm run pull:staging` | Pull current staging script down to src/ |
| `npm run pull:live` | Pull current live script down to src/ |
| `clasp login` | Authenticate with Google |
| `clasp open` | Open the active Apps Script project in browser |

---

## Branch Strategy

```
main ─────────────────────────────────► Live
        ▲
staging ──────────────────────────────► Staging
        ▲
feature/your-change ──► staging ──► test ──► main
```

- `main` — always matches what is live
- `staging` — work in progress, tested before going live
- Feature branches — optional, for larger changes

---

## Project Structure

```
royal-glass-dashboard/
├── src/
│   ├── Code.js              ← All Apps Script code lives here
│   └── appsscript.json      ← Apps Script manifest (timezone, runtime)
├── .clasp.staging.json      ← Staging script ID
├── .clasp.live.json         ← Live script ID
├── .clasp.json              ← Active config (gitignored, swapped by deploy.js)
├── deploy.js                ← Deploy script (cross-platform)
├── package.json
├── .gitignore
└── README.md
```

---

## GitHub Setup (First Time)

```bash
git init
git add .
git commit -m "Initial commit — v2.0"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/royal-glass-dashboard.git
git push -u origin main

# Create staging branch
git checkout -b staging
git push -u origin staging
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `clasp: command not found` | Run `npm install -g @google/clasp` |
| `Error: Could not read API credentials` | Run `clasp login` |
| `Script ID not found` | Check `.clasp.staging.json` or `.clasp.live.json` has the correct ID |
| `Google Apps Script API has not been used` | Enable it at https://script.google.com/home/usersettings |
| Push succeeds but changes not visible | Open the spreadsheet, refresh, reload the menu |

---

## Script IDs

| Environment | Script ID |
|---|---|
| Staging | `1Pw9H3bXCS_UHM8ZaLxZ1UWpfN802IEcA4vF8tSqEeovbP6wIci1Y6LAT` |
| Live | `10xFAWGlibXl9fOkMeHktmBaYWyyjN9xV284T7zzNwLCiwQCfPF9zEiLs` |
