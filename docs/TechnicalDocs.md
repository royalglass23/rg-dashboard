# Royal Glass Work Order Dashboard
## Technical Documentation

**Version:** 2.0
**Last updated:** May 2026
**Author:** Mike Tenorio — AI Programmer, Royal Glass Limited
**Stack:** Google Apps Script, Google Sheets, ServiceM8 API, Anthropic Claude API, clasp, GitHub

---

## Table of Contents

1. System Overview
2. Architecture & Data Flow
3. Repository Structure
4. Environment Setup
5. Script Properties (API Keys)
6. Google Sheets Structure
7. Function Reference
8. CONFIG Block
9. Deployment Workflow
10. GitHub Branch Strategy
11. Staging vs Live
12. Backup & Recovery
13. Audit Log
14. Known Limitations
15. Changelog
16. Testing Checklist
17. Future Roadmap

---

## 1. System Overview

A Google Apps Script project bound to two Google Sheets (staging and live). It integrates ServiceM8 (job management) and Anthropic Claude (AI label generation) to produce a production dashboard for Royal Glass NZ.

No external servers. No cron jobs. Everything runs inside Google's infrastructure via Apps Script triggers and menu-driven actions.

**Key design principles:**
- SM8 columns are read-only — never manually edited, always overwritten on refresh
- Manual columns are sacred — never touched by any automated process
- UUID is the primary key everywhere — matches SM8 UUIDs directly
- Config-driven — all settings live in the CONFIG block, nothing hardcoded elsewhere
- DB-ready — column map and naming convention designed for future PostgreSQL migration

---

## 2. Architecture & Data Flow

```
ServiceM8 API
    │
    ▼
_fetchJobs()          → filters by status='Work Order' and active=1
_fetchCompanyMap()    → resolves company_uuid → client name
_fetchMaterials()     → all active job materials
    │
    ▼
_filterExcludedItems()  → removes invoice/deposit items
    │
    ├──► _syncWorkOrdersSheet()    → upsert Work Orders sheet
    │         - updates cols A–E (SM8)
    │         - never touches col F (Status, manual)
    │         - deletes rows no longer in SM8
    │
    └──► _syncJobMaterialsSheet()  → upsert Job Materials sheet
              - updates cols A–F (SM8)
              - never touches cols G–Q (label + manual)
              - deletes rows no longer active
              │
              ▼
         generateShortLabels()
              - only runs on rows where col G is blank
              - calls Anthropic Claude Haiku API
              - writes label + hover note to col G
```

---

## 3. Repository Structure

```
royal-glass-dashboard/
├── src/
│   ├── Code.js              ← All Apps Script source code
│   └── appsscript.json      ← Apps Script manifest (timezone, runtime)
├── docs/
│   ├── UserGuide.md         ← Non-technical user documentation
│   └── TechnicalDocs.md     ← This file
├── .clasp.staging.json      ← Staging script ID
├── .clasp.live.json         ← Live script ID
├── .clasp.json              ← Active config (gitignored — swapped by deploy.js)
├── .clasprc.json            ← Auth tokens (gitignored — never commit)
├── deploy.js                ← Cross-platform deploy script (Node.js)
├── package.json             ← npm scripts
├── .gitignore
└── README.md
```

---

## 4. Environment Setup

### Prerequisites

- Node.js (v16+) — confirm with `node --version`
- npm — confirm with `npm --version`
- Google account with access to both spreadsheets
- Apps Script API enabled

### First-time setup

```bash
# 1. Clone the repo
git clone https://github.com/royalglass23/royal-glass-dashboard.git
cd royal-glass-dashboard

# 2. Install dependencies
npm install

# 3. Enable Apps Script API
# Go to: https://script.google.com/home/usersettings
# Turn ON: Google Apps Script API

# 4. Log in to Google via clasp
clasp login

# 5. Deploy to staging first
npm run deploy:staging

# 6. Test in Staging WO Dashboard spreadsheet

# 7. Deploy to live when confirmed
npm run deploy:live
```

### Subsequent deploys

```bash
# Make changes to src/Code.js
npm run deploy:staging   # push to staging
# test
npm run deploy:live      # push to live
```

---

## 5. Script Properties (API Keys)

API keys are stored in Apps Script Script Properties — not in code, not in the repo.

**How to set them:**
1. Open the spreadsheet
2. Extensions → Apps Script
3. Gear icon (Project Settings)
4. Script Properties tab
5. Add the following:

| Key | Value |
|---|---|
| `SM8_API_KEY` | ServiceM8 API key |
| `ANTHROPIC_API_KEY` | Anthropic Claude API key |

> Both spreadsheets (staging and live) need their own Script Properties set independently.

**How they are accessed in code:**
```javascript
PropertiesService.getScriptProperties().getProperty("SM8_API_KEY")
```

---

## 6. Google Sheets Structure

### Spreadsheets

| Name | Purpose | Script ID |
|---|---|---|
| Staging WO Dashboard | Development and testing | `1Pw9H3bXCS_UHM8ZaLxZ1UWpfN802IEcA4vF8tSqEeovbP6wIci1Y6LAT` |
| Work Order Dashboard | Production / live | `10xFAWGlibXl9fOkMeHktmBaYWyyjN9xV284T7zzNwLCiwQCfPF9zEiLs` |

### Work Orders sheet

| Col | Field | Source | DB field (future) |
|---|---|---|---|
| A | Job No | SM8 auto | job_no |
| B | Client Name | SM8 auto | client_name |
| C | Site Address | SM8 auto | site_address |
| D | Project Description | SM8 auto | project_desc |
| E | Date Created | SM8 auto | date_created |
| F | Status | Manual dropdown | status |
| G | Item ID | Hidden — PK | uuid |

### Job Materials sheet

| Col | Field | Source | DB field (future) |
|---|---|---|---|
| A | Job No | SM8 auto | job_no |
| B | Client Name | SM8 auto | client_name |
| C | Site Address | SM8 auto | site_address |
| D | Date Created | SM8 auto | date_created |
| E | Project Description | SM8 auto | project_desc |
| F | Item Name | SM8 hidden | item_name |
| G | Item - Short Label | AI generated | short_label |
| H | Installer | Manual | installer |
| I | Product Type | Manual | product_type |
| J | Design Status | Manual | design_status |
| K | Expected Completion Date | Manual | expected_date |
| L | Confirmed Date | Manual | confirmed_date |
| M | Glass Status | Manual | glass_status |
| N | Hardware Status | Manual | hardware_status |
| O | Site Ready | Manual checkbox | site_ready |
| P | Overall Status | Manual | overall_status |
| Q | Remarks | Manual | remarks |
| R | Item UUID | Hidden — PK | uuid |

---

## 7. Function Reference

### Public functions (menu-accessible)

| Function | Description |
|---|---|
| `onOpen()` | Builds the Royal Glass menu on spreadsheet open |
| `setupSheets()` | Creates both sheets with formatting and dropdowns. Run once. |
| `refreshLast7Days()` | Triggers refresh with 7-day date filter |
| `refreshLast30Days()` | Triggers refresh with 30-day date filter |
| `fullRefresh()` | Triggers refresh with no date filter |
| `generateShortLabels()` | Generates AI labels for rows where col G is blank |
| `backupData()` | Snapshots both sheets as hidden timestamped sheets |
| `loadBackup()` | Restores a selected backup over current data |
| `listBackups()` | Lists all available backup sheets |
| `deleteAllBackups()` | Deletes all backup sheets after confirmation |
| `clearData()` | Wipes all data rows from both sheets |
| `backfillNotes()` | One-time utility: adds hover notes to existing labels |

### Private functions (internal use)

| Function | Description |
|---|---|
| `_runRefresh(days)` | Core refresh engine. `days=null` for full refresh |
| `_syncWorkOrdersSheet()` | Upserts Work Orders sheet from SM8 data |
| `_syncJobMaterialsSheet()` | Upserts Job Materials sheet from SM8 data |
| `_fetchCompanyMap(apiKey)` | Returns UUID→name map of all SM8 companies |
| `_fetchJobs(apiKey, days)` | Fetches active Work Orders from SM8 |
| `_fetchMaterials(apiKey)` | Fetches all active job materials from SM8 |
| `_filterExcludedItems(materials)` | Removes invoice/deposit items from array |
| `_buildLabelPrompt(rawItem)` | Returns the Claude prompt string for label generation |
| `_createWorkOrdersSheet(ss)` | Creates and formats Work Orders sheet |
| `_createJobMaterialsSheet(ss)` | Creates and formats Job Materials sheet |
| `_ensureAuditLogSheet(ss)` | Creates Audit Log sheet if it does not exist |
| `_deleteAllDataRows(sheet)` | Clears and deletes all data rows except header |
| `_getDateDaysAgo(n)` | Returns ISO date string n days ago |
| `_writeAuditLog(e, ...)` | Appends an edit record to the Audit Log |
| `_getColumnName(sheetName, col)` | Returns human-readable column name for audit log |
| `_archiveAuditLogIfNeeded()` | Moves rows older than 30 days to archive sheet |

### Event handlers

| Function | Trigger |
|---|---|
| `onEdit(e)` | Fires on any cell edit — logs manual changes to Audit Log |

---

## 8. CONFIG Block

All configuration is centralised at the top of `Code.js`. Nothing else in the codebase should need to change for routine updates.

```javascript
const CONFIG = {
  sheets: { ... },          // Sheet names
  workOrders: { ... },      // Column index map + SM8 col count
  jobMaterials: { ... },    // Column index map + SM8 col count
  excludeKeywords: [...],   // Items excluded from dashboard
  ai: { ... },              // Claude model, token limit, key name
  sm8: { ... },             // SM8 base URL, key name
  theme: { ... },           // Hex colours
  dropdowns: { ... },       // All dropdown values
}
```

**To add a new excluded keyword:**
```javascript
excludeKeywords: ["partial invoice", "invoice", "deposit", "your new keyword"],
```

**To add a new installer to the dropdown:**
```javascript
installer: ["Royal Glass", "Mike", "New Person"],
```

**To change the AI model:**
```javascript
ai: { model: "claude-haiku-4-5-20251001", ... }
```

---

## 9. Deployment Workflow

```bash
# Standard deploy cycle
git checkout staging
# make changes to src/Code.js
npm run deploy:staging
# open Staging WO Dashboard and test
git add .
git commit -m "describe the change"
git push origin staging

git checkout main
git merge staging
npm run deploy:live
git push origin main
```

### deploy.js internals

```javascript
// Copies .clasp.{env}.json → .clasp.json, then runs clasp push --force
fs.copyFileSync(`.clasp.${env}.json`, '.clasp.json');
execSync('clasp push --force', { stdio: 'inherit' });
```

### npm scripts

| Command | What it runs |
|---|---|
| `npm run deploy:staging` | `node deploy.js staging` |
| `npm run deploy:live` | `node deploy.js live` |
| `npm run pull:staging` | Pulls current staging script to src/ |
| `npm run pull:live` | Pulls current live script to src/ |

---

## 10. GitHub Branch Strategy

```
main ────────────────────────────────► Live spreadsheet
  ▲
staging ─────────────────────────────► Staging spreadsheet
  ▲
feature/your-change ──► staging ──► test ──► main
```

| Branch | Maps to | Rule |
|---|---|---|
| `main` | Live | Only merge from staging after testing |
| `staging` | Staging | Work in progress |
| `feature/*` | Local only | Optional for larger changes |

**Never deploy directly to live from a feature branch.**

---

## 11. Staging vs Live

| | Staging | Live |
|---|---|---|
| Spreadsheet | Staging WO Dashboard | Work Order Dashboard |
| Script ID | `1Pw9H3...` | `10xFAW...` |
| Script Properties | Set separately | Set separately |
| SM8 data | Same source — real SM8 | Same source — real SM8 |
| Who uses it | Mike only | Roxy and staff |
| Branch | `staging` | `main` |

> Both environments connect to the same ServiceM8 account. Staging pulls real data — it is not a sandbox. Be mindful when testing destructive functions (Clear Data, Delete Backups).

---

## 12. Backup & Recovery

### Backup strategy

- Max 5 backups kept at any time
- Oldest backup auto-deleted when a 6th is created
- Backups stored as hidden sheets named: `Backup — DD MMM YYYY HH:MM | Sheet Name`
- `copyTo()` copies values, formatting, and notes

### Recovery

```
Menu → 📂 Load Backup → select number → confirm
```

`loadBackup()` restores both `getValues()` and `getNotes()` — labels and hover tooltips are preserved.

### Backup limitations

- Does not back up dropdowns or banding (formatting)
- Does not back up the Audit Log
- Max 5 per environment — older ones are gone permanently

---

## 13. Audit Log

### What is tracked

Every edit to a manual column fires `onEdit(e)` which calls `_writeAuditLog()`.

Tracked columns:
- Work Orders: col F (Status)
- Job Materials: cols H–Q (all manual columns)

### Log schema

| Col | Field |
|---|---|
| A | Timestamp |
| B | Edited By (Google account email) |
| C | Sheet name |
| D | Job No |
| E | Client Name |
| F | Column letter + name |
| G | Old value |
| H | New value |

### Archiving

`_archiveAuditLogIfNeeded()` runs at the start of every refresh.
- Rows older than 30 days move to a hidden sheet: `Audit Archive — MMM YYYY`
- Active log retains last 30 days only
- Archive sheets are permanent

---

## 14. Known Limitations

| Limitation | Detail |
|---|---|
| 6-minute execution limit | Apps Script times out at 6 minutes. Full Refresh with many items and labels may hit this. If it times out, run refresh again — it will skip already-labelled rows. |
| No real-time sync | Data only updates when a refresh is manually triggered. |
| SM8 API rate limits | `Utilities.sleep(300)` is added between label generation calls to avoid hitting Anthropic rate limits. |
| Backup does not restore formatting | `loadBackup()` restores values and notes but not dropdowns or alternating row colours. |
| Staging uses real SM8 data | There is no separate SM8 sandbox. All refreshes pull real production data. |
| `onEdit` is not triggered by script writes | The Audit Log only records edits made by a human in the sheet, not by the refresh script. |
| Google Sheets requires min 1 data row | `_deleteAllDataRows()` keeps row 2 as an empty row to satisfy this requirement. |

---

## 15. Changelog

### v2.0 — May 2026
- Full refactor — all settings moved to CONFIG block
- `_runRefresh(days)` replaces separate refresh functions
- 7-day, 30-day, and Full Refresh options added to menu
- Stale rows now deleted (not hidden) on refresh
- `backupData()` with rolling 5-backup limit
- `loadBackup()` with notes restoration
- `deleteAllBackups()` added
- `clearData()` now prompts for backup first and uses `deleteRows()` instead of `clearContent()`
- Audit Log auto-created on setup with warning-only protection
- `_archiveAuditLogIfNeeded()` — monthly archiving of audit rows older than 30 days
- Alternating row banding via `applyRowBanding()` replaces column background colours
- `backfillNotes()` utility for migrating existing labels
- Deploy pipeline: clasp + deploy.js + GitHub branches

### v1.6 — April 2026
- Two-line label format introduced (later simplified back to one line per scope)
- `generateShortLabels` runs automatically at end of every refresh
- `clearDataValidations` fix on col G before applying dropdowns
- `max_tokens` bumped from 80 to 120 for multi-scope items
- Item Labels Test sheet removed

### v1.5 — April 2026
- Auto-generate labels on refresh
- `applyColumnFormattingMaterials` fixed — clears leftover validation on col G before applying

### v1.4 — April 2026
- Work Orders simplified to 7 columns (summary only)
- All production detail moved to Job Materials
- `setupTestSheet`, `loadTestData`, `runLabelTest` functions removed
- Refresh split into two modes: 7-day and Full

### v1.0–v1.3 — April 2026
- Initial build
- ServiceM8 API integration (jobs + materials + companies)
- Safe merge logic using UUIDs
- Status dropdown on Work Orders
- Installer dropdown on Job Materials
- Basic label generation with Claude Haiku

---

## 16. Testing Checklist

Run through this after every deploy to staging before pushing to live.

### Refresh
- [ ] 7-day refresh completes without error toast
- [ ] 30-day refresh completes without error toast
- [ ] Full Refresh shows confirmation prompt and completes
- [ ] Sync toast shows correct counts
- [ ] SM8 columns updated on existing rows
- [ ] Manual columns not overwritten
- [ ] Invoice/deposit items not present in Job Materials
- [ ] New items get labels generated
- [ ] Existing labels not regenerated

### Labels
- [ ] Hovering over a label shows full item name
- [ ] Clearing a label cell and refreshing regenerates it
- [ ] Multi-scope items produce multiple lines
- [ ] Label format matches: `Name | L/H: value | Finish`

### Backup & Restore
- [ ] Backup creates two hidden sheets with timestamp
- [ ] List Backups shows correct count
- [ ] Load Backup restores values and hover notes
- [ ] Max 5 backups — 6th backup deletes oldest

### Audit Log
- [ ] Editing Status in Work Orders creates audit row
- [ ] Editing a manual column in Job Materials creates audit row
- [ ] Old value and new value both recorded correctly

### Clear Data
- [ ] Prompts for backup first
- [ ] Asks for confirmation before deleting
- [ ] Only header row remains after clear

---

## 17. Future Roadmap

| Item | Notes |
|---|---|
| PostgreSQL database | Replace Google Sheets as data store. CONFIG column map is ready for this migration. |
| n8n automation | Webhook-triggered refresh when SM8 job status changes |
| ServiceM8 webhook sync | Real-time rather than manual refresh |
| Quote tracking system | DocSend-style engagement tracking on client-facing quotes |
| Lead generation pipeline | Python scraping + n8n + ServiceM8 |
| Admin UI | Simple web dashboard replacing Google Sheets for Roxy |
