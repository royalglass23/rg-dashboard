# Royal Glass — Work Order Dashboard

Google Apps Script project managed with clasp and GitHub. Pulls live job and material data from ServiceM8, generates AI-powered short labels via Claude Haiku, and provides a production dashboard for Royal Glass NZ.

---

## Environments

| Environment | Spreadsheet | Branch | Script ID |
|---|---|---|---|
| Staging | Staging WO Dashboard | `staging` | `1Pw9H3bXCS_UHM8ZaLxZ1UWpfN802IEcA4vF8tSqEeovbP6wIci1Y6LAT` |
| Live | Work Order Dashboard | `main` | `10xFAWGlibXl9fOkMeHktmBaYWyyjN9xV284T7zzNwLCiwQCfPF9zEiLs` |

**Rule: never deploy straight to live. Always test on staging first.**

---

## Stack

- **Google Apps Script** — runtime
- **Google Sheets** — data store (future: PostgreSQL)
- **ServiceM8 API** — job and material source
- **Anthropic Claude Haiku** — AI label generation
- **clasp** — Apps Script CLI
- **GitHub** — version control

---

## Quick Start

### Prerequisites

- Node.js v16+
- npm
- Google account with access to both spreadsheets
- Apps Script API enabled at https://script.google.com/home/usersettings

### Setup

```bash
git clone https://github.com/royalglass23/royal-glass-dashboard.git
cd royal-glass-dashboard
npm install
clasp login
```

### Set API Keys

In each spreadsheet: **Extensions → Apps Script → Project Settings → Script Properties**

| Key | Value |
|---|---|
| `SM8_API_KEY` | ServiceM8 API key |
| `ANTHROPIC_API_KEY` | Anthropic Claude API key |

> Set these in both staging and live spreadsheets separately.

### Deploy

```bash
npm run deploy:staging   # push to staging — test here first
npm run deploy:live      # push to live — only after staging passes
```

---

## Development Workflow

```bash
# 1. Work on staging branch
git checkout staging

# 2. Make changes to src/Code.js

# 3. Deploy and test
npm run deploy:staging

# 4. Commit
git add .
git commit -m "describe the change"
git push origin staging

# 5. Merge to main and go live
git checkout main
git merge staging
npm run deploy:live
git push origin main
```

---

## Commands

| Command | Description |
|---|---|
| `npm run deploy:staging` | Push code to Staging WO Dashboard |
| `npm run deploy:live` | Push code to Live Work Order Dashboard |
| `npm run pull:staging` | Pull staging script down to src/ |
| `npm run pull:live` | Pull live script down to src/ |
| `clasp login` | Authenticate with Google |

---

## Project Structure

```
royal-glass-dashboard/
├── src/
│   ├── Code.js              ← All Apps Script source code
│   └── appsscript.json      ← Manifest (timezone: Pacific/Auckland, runtime: V8)
├── docs/
│   ├── UserGuide.md         ← Non-technical guide for Roxy and staff
│   └── TechnicalDocs.md     ← Full technical reference A–Z
├── .clasp.staging.json      ← Staging script ID
├── .clasp.live.json         ← Live script ID
├── .clasp.json              ← Active config (gitignored)
├── deploy.js                ← Cross-platform deploy script
├── package.json
├── .gitignore
└── README.md
```

---

## Key Design Decisions

- **Config-driven** — all settings in `CONFIG` at top of `Code.js`
- **UUID as PK** — matches ServiceM8 UUIDs, ready for DB migration
- **SM8 columns read-only** — cols A–E (Work Orders) and A–F (Job Materials) only written by refresh
- **Manual columns protected** — never overwritten by any automated process
- **Rolling backups** — max 5 kept, oldest auto-deleted
- **Audit log** — every manual edit recorded with old/new value and editor email

---

## Documentation

| Document | Audience |
|---|---|
| [User Guide](docs/UserGuide.md) | Roxy and office staff — non-technical |
| [Technical Docs](docs/TechnicalDocs.md) | Developers — full A–Z reference |

---

## Changelog

**v2.0 — May 2026** — Full refactor, CONFIG block, 7/30/Full refresh, backup/restore, audit log, clasp + GitHub pipeline
**v1.6 — April 2026** — Auto-label generation on refresh, two-line label format
**v1.5 — April 2026** — Label automation, col G validation fix
**v1.4 — April 2026** — Work Orders simplified, test sheet removed
**v1.0 — April 2026** — Initial build, SM8 integration, basic label generation
